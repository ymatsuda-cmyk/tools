# qwen3.8:27b データセット作成手順

CPU（Accelerator: None）+ Internet On のKaggle Notebookで、以下を1セルで実行する。
GPU週間枠を消費しない。

```python
import subprocess, os

# zstd（ollamaインストーラの展開に必要）
subprocess.run(['bash', '-lc', 'apt-get -qq update && apt-get -qq install -y zstd'], check=True)

# ollamaインストール
subprocess.run(['bash', '-lc', 'curl -fsSL https://ollama.com/install.sh | sh'], check=True)

# serve起動
subprocess.Popen(['ollama', 'serve'])
import time; time.sleep(5)

# pull（27bはサイズが大きいため数分〜数十分かかる）
subprocess.run(['ollama', 'pull', 'qwen3.8:27b'], check=True)

# --- ここで一旦 vision 対応を確認する（思い込みで進めない） ---
r = subprocess.run(['ollama', 'show', 'qwen3.8:27b'], capture_output=True, text=True)
print(r.stdout)
print('VISION:', 'vision' in r.stdout.lower())
print('AUDIO:', 'audio' in r.stdout.lower())

# 保存先を確認（通常 ~/.ollama/models 配下に blobs/manifests ができる）
print(subprocess.run(['bash', '-lc', 'ls -la ~/.ollama/models'], capture_output=True, text=True).stdout)
```

実行後、Notebookの出力（`~/.ollama/models/blobs` と `~/.ollama/models/manifests`）を
「Save Version」→「New Dataset」でデータセット化する。

- データセットref例: `matsuda2026/ollama-qwen3-8-27b`
- 階層はどこでもよい（`find_models_dir()` が `blobs` と `manifests` を両方含む
  ディレクトリを再帰探索するため）

## 重要: この案件での画像・音声対応について

依頼シートには「画像対応: はい」「音声対応: はい」とありましたが、事前調査の結果:

- **画像（vision）**: `qwen3.8:27b` は公式Ollamaライブラリ上で `vision` タグ付きの
  マルチモーダルモデルとして提供されており、`ollama show` の `Capabilities` にも
  `vision` が出ることを確認済みです（256K context、テキスト+画像入力）。
  → **画像対応は問題なく実現できます。**
- **音声（audio）**: 2026年8月時点のOllama公式ライブラリでは、`qwen3.8:27b` の
  Capabilities タグは `vision / tools / thinking` のみで、`audio` は含まれていません。
  Ollama自体、音声入力APIを一般提供しておらず（gemma4など一部モデル専用の対応に限られる）、
  `qwen3.8:27b` を音声対応モデルとして使うことはできません。
  → **音声対応は今回のモデル選定では実現できません。**

音声入力がどうしても必要な場合は、以下のいずれかを検討してください。

1. 別エンドポイントとして `gemma4:12b`（audio対応タグあり）や、音声専用モデル
   （`qwen2-audio` 系など）を追加する
2. 音声はブラウザ側でWhisper等を使って先にテキスト化し、テキストとして
   qwen3.8:27bへ渡す（`proxy_py.html` は音声パート `input_audio` を受け取っても
   ログを出して無視するだけの実装にしてあります）

今回提供する `proxy_py.html` は上記を踏まえ、**画像のみ変換対応・音声は非対応として
明示的にスキップする**実装にしています。