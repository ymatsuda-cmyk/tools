# qwen3-coder:30b データセット作成手順

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

# pull（MoEモデルだがダウンロードサイズは約19GB。数分〜十数分かかる）
subprocess.run(['ollama', 'pull', 'qwen3-coder:30b'], check=True)

# --- capabilities確認（vision非対応・thinking非対応のはず） ---
r = subprocess.run(['ollama', 'show', 'qwen3-coder:30b'], capture_output=True, text=True)
print(r.stdout)
print('VISION:', 'vision' in r.stdout.lower())
print('THINKING:', 'thinking' in r.stdout.lower())

# ~/.ollama/models を /kaggle/working 配下へコピー（Save Versionで拾われるようにするため）
subprocess.run(['bash', '-lc',
    'mkdir -p /kaggle/working/ollama-models && '
    'cp -r ~/.ollama/models/blobs /kaggle/working/ollama-models/ && '
    'cp -r ~/.ollama/models/manifests /kaggle/working/ollama-models/'
], check=True)

# 確認（Outputパネルの上限19.5GiBに対して、qwen3-coder:30bは約19GBなのでギリギリ。
# 前回のqwen3.8:27bデータセットと合わせて複数同時に置かない）
print(subprocess.run(['bash', '-lc', 'du -sh /kaggle/working/ollama-models; ls -la /kaggle/working/ollama-models'],
                      capture_output=True, text=True).stdout)
```

## 容量に関する注意

qwen3-coder:30b（Q4_K_M量子化）は約19GBあり、Kaggle Notebookの**Output上限（19.5GiB）にかなり近い**です。

- このセッションでは**他のモデルを同時にpullしない**（前回のqwen3.8:27b関連ファイルなどが`/kaggle/working`に残っていないか確認）
- 万一Output上限を超える場合は、`q4_K_M`より小さいタグが公式に無いため、以下のいずれかで対応:
  - 一度データセット化した後、モデル本体データセットと`ngrok-binary`データセットを**本番Notebook側でInputとして分けて追加**する（データセット作成用Notebookの容量とは別枠なので問題なし）
  - どうしても収まらない場合は`q4_K_M`を手動で`llama.cpp`系の外部量子化に差し替える必要があるが、通常はこの手順のままで収まるはず

保存後、**Save Version → Save & Run All (Commit)** → 完了後に **New Dataset** で切り出してください。
推奨ref: `matsuda2026/ollama-qwen3-coder-30b`

## capabilities確認結果の見方

- `VISION: False` が正しい（qwen3-coder:30bはコード特化でvision非対応）
- `THINKING`欄の有無に関わらず、Hugging Face上の情報では**このモデルはnon-thinking mode専用**（`<think></think>`ブロックを生成しない）とされているので、`enable_thinking`や`think`パラメータを気にする必要は基本的にありません
- MoE構成（総パラメータ30.5B、活性化3.3B）なので、`ollama show`の`Parameters`欄に`architecture`として`qwen3moe`系の表記が出るはずです

データセットができたら、`CONFIG_qwen3_coder_30b.json`の`datasets`配列に実際のref名を反映してください（既に`matsuda2026/ollama-qwen3-coder-30b`を仮に入れてあります）。
