// src/config.js — デプロイごとに変える値（秘密ではない）
// GAS をウェブアプリとしてデプロイした URL（…/exec）
export const GAS_URL = 'https://script.google.com/macros/s/AKfycbzptPsP4bWRkRlGpU_u1t6GXj0tPhAruwhrzATfmYZ1V7AVBcXTGHhoTr9jTv-50IHeeg/exec'
// 一覧を取り直す間隔（ほかの端末の登録を拾う）
export const POLL_MS = 120000
// Claude.ai で要約待ちを変換するときに開くページ（設定で端末ごとに変えられる）
export const CLAUDE_DEFAULT_URL = 'https://claude.ai/new'
// Mac の確認間隔（表示用。launchd の StartInterval と合わせる）
export const WORKER_EVERY_SEC = 300
// マーカーの色（書籍アプリと同じ3色）。名前は、ノートで「タグ」として使われる
export const MARK_TAGS = [{ id: 1, name: '重要' }, { id: 2, name: '確認' }, { id: 3, name: '疑問' }]
