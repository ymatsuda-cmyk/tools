# Excel読み込み仕様書

## 概要
バグ管理アドインのExcel連携機能における、Excelシートからのデータ読み込みに関する仕様をまとめたドキュメントです。

## シート構造

### 基本構造
- **対象シート名**: `バグ`
- **項目行（ヘッダー行）**: 2行目
- **入力例・設定行**: 3行目
- **データ開始行**: 4行目以降
- **対象列数**: 31列（A列〜AE列）

### 行の役割
| 行番号 | 役割 | 内容 |
|--------|------|------|
| 1行目 | （任意） | タイトル等（読み込み対象外） |
| 2行目 | 項目名 | 各列のラベル（ヘッダー） |
| 3行目 | 設定・入力例 | 各種設定値とデータ入力例 |
| 4行目以降 | データ | 実際のバグ情報データ |

## 列定義

### 列マッピング（A〜AE列）
| 列 | キー | 日本語名 | データ型 | 必須 | 説明 |
|----|------|----------|----------|------|------|
| A | id | ID | readonly | ○ | バグID（自動採番） |
| B | title | タイトル | text | ○ | バグのタイトル |
| C | status | 状況 | select | ○ | 新規/解析/修正/確認/再発/完了 |
| D | updated | 更新日 | date | ○ | 最終更新日（自動更新） |
| E | assignee | 担当者 | select | - | バグの担当者 |
| F | occurredOn | 発生日 | date | - | バグ発生日 |
| G | reporter | 登録者 | select | - | バグ登録者 |
| H | origin | 発生起因 | select | - | バグの発生起因 |
| I | originNumber | 起因番号 | text | - | 起因の識別番号 |
| J | steps | 再現手順 | textarea | - | バグの再現手順 |
| K | expected | 期待する動作 | textarea | - | 期待される正常動作 |
| L | actual | 実際の動作 | textarea | - | 実際に発生した動作 |
| M | reproRate | 再現率 | select | - | 毎回/時々/1回のみ |
| N | cause | 原因 | textarea | - | バグの原因分析 |
| O | analyst | 解析者 | select | - | バグを解析した担当者 |
| P | analysisDate | 解析日 | date | - | 解析実施日 |
| Q | scope | 影響範囲 | select | - | バグの影響範囲 |
| R | fix | 対応内容 | textarea | - | 修正内容 |
| S | fixVer | 修正Ver | text | - | 修正バージョン |
| T | fixer | 対応者 | select | - | 修正実施者 |
| U | fixDate | 対応日 | date | - | 修正実施日 |
| V | verify | 確認内容 | textarea | - | 確認テスト内容 |
| W | reject | 差し戻し | text | - | 差し戻し理由 |
| X | verifier | 確認者 | select | - | 確認実施者 |
| Y | verifyDate | 確認日 | date | - | 確認実施日 |
| Z | tag | タグ | text | - | バグに付与するタグ |
| AA | priority | 優先度 | select | - | バグの優先度 |
| AB | severity | 影響度 | select | - | バグの影響度 |
| AC | starred | 本日分 | text | - | 本日対応マーク |
| AD | periodStart | 期間開始 | date | - | 表示期間開始日 |
| AE | periodEnd | 期間終了 | date | - | 表示期間終了日 |

## 設定セルの仕様（3行目）

### C3セル：状態別表記設定
- **形式**: `元の状態:表示名/元の状態:表示名`
- **例**: `新規:新規/解析:解析中/修正:修正済み`
- **デフォルト**: 新規/解析/修正/確認/再発/完了
- **用途**: ステータス表示名のカスタマイズ

### E3セル：メンバーリスト設定
- **形式**: `/`区切りの文字列
- **例**: `政次/高橋/伊藤/松田`
- **用途**: 担当者・登録者・解析者・対応者・確認者の選択肢
- **注意**: このセルが担当者と登録者両方の設定に使用される

### H3セル：発生起因リスト設定
- **形式**: `/`区切りの文字列
- **例**: `定義(通常)/定義(電源断)/定義(通信断)`
- **デフォルト**: 定義(通常)/定義(電源断)/定義(通信断)

### Z3セル：プリセットタグ設定
- **形式**: `/`区切りの文字列
- **例**: `UI/RPA/通信/電源/設定/認証/データ/パフォーマンス`
- **用途**: タグ入力時の候補表示

### AA3セル：優先度リスト設定
- **形式**: `/`区切りの文字列
- **例**: `高（最優先）/中/低（改善）`
- **デフォルト**: 高（最優先）/中/低（改善）

### AB3セル：影響度リスト設定
- **形式**: `/`区切りの文字列
- **例**: `致命的/重大/軽微`
- **デフォルト**: 致命的/重大/軽微

### AD3・AE3セル：期間設定
- **AD3**: 期間開始日（日付形式またはExcelシリアル日付）
- **AE3**: 期間終了日（日付形式またはExcelシリアル日付）
- **用途**: トレンド分析の期間設定

## 読み込み処理フロー

### 1. 環境判定
```javascript
if (!state.inOffice) {
  // ブラウザ単体モード：デモデータを使用
  state.bugs = demoData();
  return;
}
```

### 2. 設定読み込み
- 3行目のC3〜AB3セルから各種設定を読み込み
- 各設定値をパース（`/`区切りの分割処理）
- デフォルト値の適用

### 3. データ範囲取得
```javascript
const used = sheet.getUsedRange(true);
const dataRange = sheet.getRangeByIndexes(
  DATA_START - 1, 0, rowCount - (DATA_START - 1), COL_COUNT
);
```

### 4. データ変換処理
- **日付変換**: Excelシリアル日付を文字列形式に変換
- **空値処理**: null/undefinedを空文字に統一
- **型変換**: すべての値をString型に統一

### 5. オブジェクト生成
各行のデータを以下の形式でオブジェクト化：
```javascript
const obj = { 
  rowIndex: DATA_START + r,  // Excel行番号を保持
  id: row[0],
  title: row[1],
  status: row[2],
  // ... 他の列データ
};
```

## 日付処理仕様

### Excelシリアル日付の変換
```javascript
function excelSerialToDateStr(v) {
  if (typeof v === 'number' && v > 0) {
    const date = new Date((v - 25569) * 86400 * 1000);
    const year = date.getUTCFullYear();
    const month = String(date.getUTCMonth() + 1).padStart(2, '0');
    const day = String(date.getUTCDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }
  return v;
}
```

## エラーハンドリング

### 空行の処理
- すべてのセルが空（''またはnull）の行は読み込み対象外
- `row.every(v => v === '' || v === null)`で判定

### シート不在時の処理
- 対象シート「バグ」が存在しない場合はエラー
- `sheet.getItem(SHEET_NAME)`で例外発生

### データ不足時の処理
- DATA_START行未満の場合は空配列を返す
- `rowCount < DATA_START`で判定

## 動作モード

### Excelアドインモード
- `state.inOffice = true`
- 実際のExcelシートからデータを読み込み
- Office.jsライブラリを使用
- Excel.run()コンテキスト内で処理

### ブラウザ単体モード
- `state.inOffice = false`  
- プリセットのデモデータを使用
- Excel連携なしで動作確認可能
- 開発・デバッグ用途

## パフォーマンス考慮事項

### 一括読み込み
- 全データを一度にRangeで取得
- 列ごとの個別読み込みは行わない
- `sheet.getRangeByIndexes()`で範囲指定

### 同期タイミング
- 設定読み込み後に`await ctx.sync()`
- データ読み込み後に`await ctx.sync()`
- 必要最小限の同期回数

### メモリ使用量
- 使用範囲のみを取得（`getUsedRange(true)`）
- 不要な列は読み込まない（COL_COUNT制限）

## 制限事項

### 対応データ型
- 文字列、数値、日付のみ対応
- 数式、ハイパーリンク、画像などは未対応

### 最大データ量
- Excelの行数制限に依存
- パフォーマンスは約1万行程度まで確認済み

### 文字エンコーディング
- UTF-8での日本語文字に対応
- 特殊文字、絵文字の動作は未保証

## 関連機能

### 保存処理
- `saveBugToExcel(bug)`関数で個別行の更新
- 更新日（D列）の自動更新
- 文字列折り返し設定の無効化

### フィルタリング
- 読み込み後に`applyFilters(bugs)`でフィルタ適用
- クライアントサイドでの絞り込み処理

### キーワードネットワーク
- 読み込んだデータを元にキーワード関連図を生成
- 別ウィンドウ（bug-network-tool.html）で表示