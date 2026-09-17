# PLAUD 文字起こしパイプライン

## ファイル構成

| ファイル | 役割 |
|---|---|
| `plaud_transcribe_notion.py` | メイン（PLAUD取得 → 文字起こし → Notion登録 → index.json） |
| `transcribe_engines.py` | labelによるエンジン切り替え（whisper / qwen3-asr） |
| `qwen3_asr.py` | Qwen3-ASR（mlx-qwen3-asr）ラッパー |
| `speaker_diarize.py` | 声紋登録・照合。クラスタリング版と名前付けのみの版 |
| `llm_refine.py` | フィラー除去・表記統一（ルールベース） |
| `bench_asr.py` | whisper と Qwen3-ASR の比較検証 |
| `settings.json` | エンジン定義 |
| `normalize_dict.json` | 表記揺れ辞書 |

## セットアップ

```bash
# 基本
pip install requests

# Qwen3-ASR（日本語のタイムスタンプ整合込み）
pip install "mlx-qwen3-asr[aligner]"

# 話者分離（pyannote）
pip install "mlx-qwen3-asr[diarize]"
# https://huggingface.co/pyannote/speaker-diarization-community-1 の利用条件に同意してから
export PYANNOTE_AUTH_TOKEN=hf_xxxxx

# 声紋の名前付け
pip install resemblyzer

# 環境診断
mlx-qwen3-asr --doctor
```

## PLAUDアカウント（複数対応）

`~/.plaud_notion_sync.env` にトークンを並べるだけで、全アカウントを横断して
未登録ファイルを拾い、同じNotion DBに登録します。

```
PLAUD_TOKEN=xxxxx
PLAUD_WS_ID=ws_clQPe6Vll0

PLAUD_TOKEN_2=yyyyy
# PLAUD_WS_ID_2=ws_xxxxxxxx   ← 未設定なら workspaceId なしで既定ワークスペースを見る
# PLAUD_REFRESH_TOKEN_2=zzzzz ← 設定するとworkspaceToken(約10日で失効)を毎回自動更新する
```

`_3` `_4` `_5` と最大5アカウントまで。`PLAUD_DOMAIN_n` はリージョンが違う場合のみ指定。

`PLAUD_REFRESH_TOKEN_n` はブラウザの DevTools → Application → Local Storage →
`pld_xxx:workspaceList` の値に含まれる `refreshToken` をコピーして設定します。
設定しておくと `PLAUD_TOKEN_n`(workspaceToken)が期限切れになっても自動で再取得し、
ローテーションされた新しい refresh_token は `.env` に自動で書き戻されます。
ログ上の名前は `account1` `account2` と自動で付きます。
Notion側の重複判定は `https://web.plaud.ai/file/{id}` で行うため、アカウントが増えても
二重登録にはなりません。

## まず検証

```bash
python3 bench_asr.py ~/sample.m4a --labels whisper qwen3-asr
```

所要時間・RTF・文字数・話者一覧が並び、`bench_out/*.txt` に各出力が残るので目視比較できます。
初回は Qwen3-ASR-1.7B（約3.4GB）と ForcedAligner-0.6B のダウンロードが走ります。

## 本番実行

```bash
python3 plaud_transcribe_notion.py --label qwen3-asr   # Qwen3-ASR + 声紋 + フィラー除去/表記統一
python3 plaud_transcribe_notion.py --label whisper     # 従来通りのプレーンテキスト
```

cron はコマンド末尾に `--label qwen3-asr` を足すだけです。

## エンジン設定

| type | ASR | 話者 | 整形 |
|---|---|---|---|
| `whisper` | mlx-whisper | なし | なし |
| `qwen3-asr` | Qwen3-ASR | pyannoteで分離 → 声紋で命名 | ルールベース（フィラー除去・表記統一） |

`qwen3-asr` の主なキー:

- `model`: `Qwen/Qwen3-ASR-0.6B`（速い）/ `Qwen/Qwen3-ASR-1.7B`（高精度）
- `diarize`: 話者分離の有無。`minSpeakers` / `maxSpeakers` / `numSpeakers`
- `contextTerms`: 辞書に加えてASRに渡す固有名詞（人名など）
- `diarization.namePrefix` / `matchThreshold` / `minSegmentSec`: 声紋照合の挙動

## 声紋DB

`~/.plaud_speakers.json` に自動生成されます。初回は `担当者1` `担当者2` … で登録されるので、
`name` を `松田` `政次` 等に書き換えれば次回から実名で出力されます。
照合のたびに埋め込みが移動平均で更新され、回を重ねるほど安定します。

## 段階的フォールバック

どこかが欠けても止まらず、できる範囲まで自動的に落ちます。

```
Qwen3-ASR + pyannote + 声紋 + ルールベース整形     ← フル
  ↓ pyannote未導入
Qwen3-ASR + 単語グルーピング + 声紋 + ルールベース整形
  ↓ resemblyzer未導入
Qwen3-ASR + 話者不明 + ルールベース整形
  ↓ mlx-qwen3-asr未導入
エラー表示（--label whisper で従来動作に切り替え）
```
