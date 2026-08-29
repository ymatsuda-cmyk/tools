kanban-addin/
├─ manifest.xml                 ← アドインのマニフェスト（Excel にサイドロード）
├─ src/
│  ├─ commands/
│  │   └─ commands.js          ← リボンボタン（ExecuteFunction）側のコード
│  └─ dialog/
│      ├─ kanban.html          ← カンバン UI（Dialog）の HTML
│      └─ kanban.js            ← カンバン UI のロジック
└─ (必要なら) package.json など


