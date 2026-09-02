# Kaggle 残量管理 (kaggle_quota)

## なぜ自前計算なのか

| 手段 | 可否 |
|---|---|
| `kernels/status` API | 403（2026年以降、全トークン種別） |
| 公式の quota API | **存在しない** |
| Kaggle Web UI の quota 表示 | 認証Cookie + XSRF が必要。GAS からは実用不可 |
| ノートブック内の環境変数 | quota 情報は入っていない |

→ **GAS を残量の権威にする。** ノートブックの打刻を台帳に積み上げ、
`残量 = limit_hours − Σ(稼働秒 ∩ 現在のクォータ窓)` を返す。

実測とのズレは `quota_calibrate` で基準点を打ち直して吸収する。

## 構成

```
Kaggle Notebook                GAS (Web App)              クライアント
  quota_client.py    ──打刻──▶  kaggle_quota.gs   ──JSON──▶ curl / ダッシュボード
  (60秒ごと heartbeat)          └ usage_log シート          / 起動前チェック
```

## 台帳スキーマ（`kaggle_quota_ledger` / `usage_log` シート）

| 列 | 内容 |
|---|---|
| session_id | Kaggle セッションID |
| accel | GPU / TPU / CPU |
| kernel | `matsuda2026/qwen3-coder-30b` など |
| started_at | ISO8601 |
| last_heartbeat_at | ISO8601 |
| ended_at | ISO8601（未終了は空） |
| seconds | 稼働秒 |
| status | RUNNING / DONE / STALE |
| note | 締めた理由 |

スプレッドシートは初回アクセス時に自動生成され、ID が Script Properties
（`QUOTA_SS_ID`）に保存される。Script Properties の 9,216 文字制限を避ける
ため、台帳本体は必ずシート側に置く。

## API

ベース: `https://script.google.com/macros/s/XXXX/exec`

### 読み取り（トークン不要）

| action | パラメータ | 内容 |
|---|---|---|
| `quota` | `accel=GPU` | 残量を返す（メイン） |
| `quota_all` | – | 全アクセラレータをまとめて |
| `quota_can_start` | `accel`, `need_hours` | 起動可否の判定 |
| `quota_log` | `limit=30` | 直近セッション履歴 |

### 打刻（ノートブックから）

| action | パラメータ |
|---|---|
| `quota_session_start` | `session_id`, `kernel`, `accel`, `started_at`(任意) |
| `quota_heartbeat` | `session_id`, `accel` |
| `quota_end` | `session_id`, `note` |

### 変更系（`token` 必須）

| action | パラメータ | 内容 |
|---|---|---|
| `quota_calibrate` | `accel`, `used_hours` | 実測値で基準を打ち直す |
| `quota_clear_baseline` | `accel` | 基準を消して純粋な台帳積算に戻す |
| `quota_reap` | – | STALE セッションを手動で締める |
| `quota_purge` | – | 保持期間外の行を削除 |
| `quota_ledger_url` | – | 台帳シートのURLを取得 |

## レスポンス例

```json
{
  "ok": true,
  "accel": "GPU",
  "limit_hours": 30,
  "used_hours": 11.42,
  "remaining_hours": 18.58,
  "remaining_pct": 61.9,
  "window": {
    "mode": "weekly",
    "start": "2026-08-29T00:00:00.000Z",
    "next_reset": "2026-09-05T00:00:00.000Z",
    "next_reset_in_sec": 216000
  },
  "baseline": { "at": "2026-08-30T12:00:00.000Z", "used_hours": 4.5 },
  "sessions_in_window": 6,
  "running": [
    {
      "session_id": "38472910",
      "kernel": "matsuda2026/qwen3-coder-30b",
      "elapsed_sec": 3480
    }
  ],
  "can_start": true,
  "source": "gas-ledger (Kaggle has no public quota API)"
}
```

## curl

```bash
BASE="https://script.google.com/macros/s/XXXX/exec"
TOKEN="..."

# 残量確認
curl -sL "$BASE?action=quota&accel=GPU" | jq

# 起動前チェック（3時間ぶん必要）
curl -sL "$BASE?action=quota_can_start&accel=GPU&need_hours=3" | jq -r .allowed

# Kaggle UI が「11h 30m used」と表示していたので補正
curl -sL "$BASE?action=quota_calibrate&accel=GPU&used_hours=11.5&token=$TOKEN" | jq

# 履歴
curl -sL "$BASE?action=quota_log&limit=20" | jq
```

`-L` は必須（GAS Web App は 302 でリダイレクトする）。

## 導入手順

1. `kaggle_quota.gs` をプロジェクトに追加
2. `QUOTA_CONFIG.API_TOKEN` を長いランダム文字列に変更
3. 既存 `doGet` の先頭に差し込む:
   ```js
   function doGet(e) {
     try {
       var q = quotaHandle(e);
       if (q) return q;
       // ... 既存のルーティング
     } catch (err) {
       return ContentService.createTextOutput(
         JSON.stringify({ ok:false, error:String(err) })
       ).setMimeType(ContentService.MimeType.JSON);
     }
   }
   ```
4. 既存ハンドラにフックを1行ずつ追加:
   ```js
   // handleStarted() 内
   try { quotaSessionStart(sessionId, kernelSlug, 'GPU'); } catch (e) {}
   // handleHeartbeat() 内
   try { quotaHeartbeat(sessionId, 'GPU', kernelSlug); } catch (e) {}
   // 停止処理内
   try { quotaSessionEnd(sessionId, 'stopped'); } catch (e) {}
   ```
5. `installQuotaTriggers()` を一度だけ手動実行（5分ごとの STALE 回収 +
   日次 purge を登録。台帳URLがログに出る）
6. **「新しいバージョン」として再デプロイ**（保存だけでは反映されない）
7. `quotaSelfTest()` で動作確認
8. Kaggle UI の quota 表示を見て `quota_calibrate` を1回打つ

## 運用上の注意

- **キャリブレーションは週1回、リセット直後に打つのが一番効く。**
  基準点が現在のクォータ窓より古い場合は無視され、純粋な台帳積算に戻る。
- ブラウザから手動で起動したセッションは打刻されないので台帳から漏れる。
  漏れた分は次のキャリブレーションで吸収される。
- `SESSION_MAX_HOURS`（既定12h）は暴走加算の安全弁。ノートブックがクラッシュ
  して heartbeat が止まっても、最後の heartbeat 時点で締められる。
- `reset_mode` は既定 `weekly`（土曜 00:00 UTC）。実際のリセット挙動が
  ローリング窓に見える場合は `'rolling'` に切り替える。
- `limit_hours: 0` は無制限扱い（CPU 用）。
