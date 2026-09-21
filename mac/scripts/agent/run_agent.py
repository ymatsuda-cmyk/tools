"""
Phase2: GitHub IssueをVS Code + Clineへ投入し、実装完了を検知して
承認待ちJSONを出力する。

安定化方針:
- Cline入力欄は固定座標 X=1500 / Y=900 を直接クリックする。
- Clineの新規チャットボタンは自動操作しない。
  新規チャット操作でClineペインが閉じる・消える問題を回避する。
- VS Codeのコマンドパレットは使用しない。
- 日本語IMEの影響を避けるため、文字入力はpyperclip + Ctrl+Vで行う。
"""

from __future__ import annotations

import argparse
import base64
import json
import os
import pathlib
import re
import shutil
import subprocess
import sys
import time
from datetime import datetime


REQUIRED_ENV_VARS = [
    "GITHUB_OWNER",
    "GITHUB_REPO",
    "AGENT_REPLY_DIR",
    "AGENT_WAITING_DIR",
    "AGENT_STATE_DIR",
    "TARGET_REPO_PATH",
]

GITHUB_OWNER = ""
GITHUB_REPO = ""
GITHUB_REPOSITORY = ""
REPLY_DIR = pathlib.Path(".")
WAITING_DIR = pathlib.Path(".")
STATE_DIR = pathlib.Path(".")
TARGET_REPO = pathlib.Path(".")

VSCODE_BOOT_WAIT_SECONDS = 25
CLINE_FOCUS_WAIT_SECONDS = 2
PASTE_WAIT_SECONDS = 3
IMPLEMENTATION_TIMEOUT_SECONDS = 900
POLL_INTERVAL_SECONDS = 10
STABLE_DURATION_SECONDS = 60

# 実測したCline入力欄の固定座標
CLINE_INPUT_X = 1500
CLINE_INPUT_Y = 900


def log(message: str) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}", flush=True)


def fail(message: str, code: int = 1) -> None:
    log(f"[ERROR] {message}")
    raise SystemExit(code)


def validate_environment_variables() -> None:
    global GITHUB_OWNER
    global GITHUB_REPO
    global GITHUB_REPOSITORY
    global REPLY_DIR
    global WAITING_DIR
    global STATE_DIR
    global TARGET_REPO

    missing = []
    print()
    print("=" * 60)
    print("環境変数チェック")
    print("=" * 60)

    for name in REQUIRED_ENV_VARS:
        value = os.getenv(name)
        if value and value.strip():
            print(f"[OK] {name}: {value}")
        else:
            missing.append(name)
            print(f"[NG] {name}: 未設定")

    if missing:
        print()
        print("不足している環境変数:")
        for name in missing:
            print(f"- {name}")
        raise SystemExit(1)

    GITHUB_OWNER = os.environ["GITHUB_OWNER"]
    GITHUB_REPO = os.environ["GITHUB_REPO"]
    GITHUB_REPOSITORY = f"{GITHUB_OWNER}/{GITHUB_REPO}"
    REPLY_DIR = pathlib.Path(os.environ["AGENT_REPLY_DIR"])
    WAITING_DIR = pathlib.Path(os.environ["AGENT_WAITING_DIR"])
    STATE_DIR = pathlib.Path(os.environ["AGENT_STATE_DIR"])
    TARGET_REPO = pathlib.Path(os.environ["TARGET_REPO_PATH"])

    print("=" * 60)
    print("必要な環境変数はすべて設定されています。")
    print("=" * 60)
    print()


def run_command(
    command: list[str],
    cwd: pathlib.Path | None = None,
    check: bool = True,
) -> subprocess.CompletedProcess[str]:
    log(f"> {subprocess.list2cmdline(command)}")
    result = subprocess.run(
        command,
        cwd=str(cwd) if cwd else None,
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
        check=False,
    )


    if result.stdout and result.stdout.strip():
        print(result.stdout.rstrip())
    if result.stderr and result.stderr.strip():
        print(result.stderr.rstrip(), file=sys.stderr)

    if check and result.returncode != 0:
        fail(
            f"コマンド失敗 ({result.returncode}): "
            f"{subprocess.list2cmdline(command)}"
        )
    return result


def git(*args: str, check: bool = True) -> subprocess.CompletedProcess[str]:
    return run_command(["git", *args], cwd=TARGET_REPO, check=check)


def require_command(name: str) -> None:
    if not shutil.which(name):
        fail(f"'{name}' がPATHにありません。")


def validate_directories() -> None:
    if not (TARGET_REPO / ".git").exists():
        fail(f"Gitリポジトリではありません: {TARGET_REPO}")

    for directory in (REPLY_DIR, WAITING_DIR, STATE_DIR):
        try:
            directory.mkdir(parents=True, exist_ok=True)
        except OSError as error:
            fail(f"フォルダを作成できません: {directory} / {error}")


def state_path(issue_number: int) -> pathlib.Path:
    return STATE_DIR / f"issue-{issue_number}.json"


def load_state(issue_number: int) -> dict:
    path = state_path(issue_number)
    if not path.exists():
        log(f"状態ファイルが見つかりません: {path.name}")
        return {}

    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as error:
        log(f"状態ファイル読込失敗: {error}")
        return {}


def load_message_id(issue_number: int) -> str:
    return str(load_state(issue_number).get("messageId", "")).strip()


def update_state(issue_number: int, updates: dict) -> None:
    state = load_state(issue_number) or {"issueNumber": issue_number}
    state.update(updates)
    state["updatedAt"] = datetime.now().isoformat()

    path = state_path(issue_number)
    path.write_text(
        json.dumps(state, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    log(f"状態ファイル更新: {path.name}")


def load_issue(issue_number: int) -> dict:
    result = run_command(
        [
            "gh",
            "issue",
            "view",
            str(issue_number),
            "--repo",
            GITHUB_REPOSITORY,
            "--json",
            "number,title,body,url,labels,state",
        ],
        cwd=TARGET_REPO,
    )

    try:
        return json.loads(result.stdout)
    except json.JSONDecodeError as error:
        fail(f"Issue JSONを解析できません: {error}")
    return {}


def slugify(text: str, max_length: int = 40) -> str:
    value = re.sub(r"[^a-z0-9]+", "-", text.lower().strip()).strip("-")
    return value[:max_length].rstrip("-") or "task"


def prepare_branch(
    issue_number: int,
    issue_title: str,
    base_branch: str,
) -> str:
    branch = f"issue-{issue_number}-{slugify(issue_title)}"
    dirty = git("status", "--porcelain").stdout.strip()
    current = git("branch", "--show-current").stdout.strip()

    if dirty and current == branch:
        log(f"同一Issueブランチの変更を継続します: {branch}")
        return branch

    if dirty:
        fail(
            f"ブランチ {current} に未コミット変更があります。"
            "commitまたはstashしてから実行してください。"
        )

    if current == branch:
        log(f"既に対象ブランチです: {branch}")
        return branch

    git("fetch", "origin")
    git("checkout", base_branch)
    git("pull", "--ff-only", "origin", base_branch)

    exists = git(
        "show-ref",
        "--verify",
        "--quiet",
        f"refs/heads/{branch}",
        check=False,
    ).returncode == 0

    if exists:
        log(f"既存ブランチへ切り替えます: {branch}")
        git("checkout", branch)
    else:
        git("checkout", "-b", branch)

    return branch


def build_prompt(issue: dict, branch: str) -> str:
    labels = ", ".join(
        item.get("name", "") for item in issue.get("labels", [])
    ) or "なし"
    body = issue.get("body") or "(本文なし)"

    return f"""GitHub Issue #{issue['number']} を実装してください。

タイトル: {issue['title']}
URL: {issue['url']}
ラベル: {labels}

本文:
{body}

前提:
- 作業ブランチ {branch} は作成済みで、既にチェックアウトされています。
- 本リポジトリはビルド不要の静的サイトです。
- ターミナルはPowerShellです。

作業ルール:
1. 素のHTML / CSS / JavaScriptで実装してください。
2. TypeScriptやフレームワークを追加しないでください。
3. package.jsonやビルド設定を追加しないでください。
4. 新規ページはリポジトリ直下のディレクトリにindex.htmlを配置してください。
5. コメントは日本語で記述してください。
6. ファイルはUTF-8 BOMなしで保存してください。
7. old/配下とIssueに関係しないファイルは変更しないでください。
8. git操作、commit、push、PR作成は行わないでください。
9. 判断に迷う場合は勝手に決めず、質問してください。
10. 最後に変更ファイルと実装内容を簡潔に報告してください。


質問・確認ルール（最優先）:
- 要件が曖昧、矛盾、又は人の判断が必要な場合、過去の質問やサンプルを再利用しないでください。
- Cline標準の質問UIだけで停止しないでください。
- 質問UIを表示する前に、必ずリポジトリ直下へ .agent-question.json を新規作成してください。
- .agent-question.json はUTF-8 BOMなしで保存してください。
- issueNumberには必ず現在のIssue番号 {issue['number']} を設定してください。
- questionIdは現在の質問内容に固有の値にしてください。
- questionには、現在Clineが確認したい質問全文をそのまま記載してください。
- choicesには、現在Clineが提示する選択肢を最大3件記載してください。
- 各choices要素には id、title、description を含めてください。
- カスタム回答が必要な場合も、選択肢の1つとして明記してください。
- .agent-question.json を作成したら、Teamsから回答が返るまで実装ファイルを変更しないでください。

必須JSON形式:
{{
  "type": "cline_question",
  "issueNumber": {issue['number']},
  "questionId": "current-question-id",
  "question": "現在のIssueについて確認したい質問全文",
  "choices": [
    {{
      "id": "choice_1",
      "title": "選択肢1",
      "description": "選択肢1の詳細"
    }},
    {{
      "id": "choice_2",
      "title": "選択肢2",
      "description": "選択肢2の詳細"
    }}
  ]
}}


実装完了時の必須報告:
- 最終回答をチャットへ表示するだけでなく、リポジトリ直下へ .agent-summary.json をUTF-8 BOMなしで作成してください。
- .agent-summary.json は実装成果物ではなく、Teams承認画面へ結果を返すための一時ファイルです。
- issueNumberには必ず現在のIssue番号 {issue['number']} を設定してください。
- 実際に実施した内容だけを書き、推測や未確認事項を確認済みとして記載しないでください。
- ブラウザ確認やコンソール確認を実行していない場合は、notPerformedへ明記してください。
- git操作は実施しないでください。

必須JSON形式:
{{
  "issueNumber": {issue['number']},
  "implementation": [
    "実装した内容1",
    "実装した内容2"
  ],
  "verification": [
    "実際に確認した内容1",
    "実際に確認した内容2"
  ],
  "notPerformed": [
    "未実施の確認事項"
  ],
  "notes": [
    "承認者へ伝える注意事項"
  ]
}}
"""


def write_json(directory: pathlib.Path, file_name: str, payload: dict) -> bool:
    path = directory / file_name
    try:
        path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        log(f"JSON出力: {path}")
        return True
    except OSError as error:
        log(f"JSON出力失敗: {path} / {error}")
        return False


def notify_teams(
    message_type: str,
    issue_number: int,
    message: str,
    extra: dict | None = None,
) -> None:
    message_id = load_message_id(issue_number)
    if not message_id:
        log("messageIdが無いためTeams通知をスキップします。")
        return

    payload = {
        "type": message_type,
        "messageId": message_id,
        "issueNumber": issue_number,
        "status": message_type,
        "message": message,
    }
    if extra:
        payload.update(extra)

    write_json(
        REPLY_DIR,
        f"issue-{issue_number}-{message_type}-{datetime.now():%H%M%S}.json",
        payload,
    )


def copy_to_clipboard(text: str) -> bool:
    try:
        import pyperclip

        pyperclip.copy(text)
        log("クリップボードへコピーしました。")
        return True
    except Exception as error:
        log(f"クリップボードコピー失敗: {error}")
        return False


def launch_vscode() -> None:
    require_command("code")
    subprocess.Popen(["code", str(TARGET_REPO)], shell=True)
    log(f"VS Code起動後、{VSCODE_BOOT_WAIT_SECONDS}秒待機します。")
    time.sleep(VSCODE_BOOT_WAIT_SECONDS)


def focus_cline_input(pyautogui) -> None:
    """
    実測した固定座標でCline入力欄をクリックする。
    新規チャットボタンは操作しないため、Clineペインが消える問題を回避する。
    """
    log(f"Cline入力欄をクリック: X={CLINE_INPUT_X}, Y={CLINE_INPUT_Y}")
    pyautogui.click(CLINE_INPUT_X, CLINE_INPUT_Y)
    time.sleep(CLINE_FOCUS_WAIT_SECONDS)

    # フォーカスを安定させるため同じ位置をもう一度クリックする。
    pyautogui.click(CLINE_INPUT_X, CLINE_INPUT_Y)
    time.sleep(0.5)


def inject_prompt_gui(prompt: str) -> bool:
    try:
        import pyautogui
        import pyperclip
    except ImportError:
        log("pip install pyautogui pyperclip を実行してください。")
        return False

    pyautogui.FAILSAFE = True

    try:
        focus_cline_input(pyautogui)

        # 入力欄に以前の未送信文字が残っている場合だけ消去する。
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.3)
        pyautogui.press("backspace")
        time.sleep(0.3)

        # コマンド名などでクリップボードが上書きされないよう、直前に再コピーする。
        pyperclip.copy(prompt)
        pyautogui.hotkey("ctrl", "v")
        log("Cline入力欄へプロンプトを貼り付けました。")

        time.sleep(PASTE_WAIT_SECONDS)
        pyautogui.press("enter")
        log("Clineへプロンプトを送信しました。")
        return True

    except Exception as error:
        log(f"Cline自動投入に失敗しました: {error}")
        return False


def test_cline_click() -> bool:
    """Issue処理をせず、Cline入力欄のクリックだけ確認する。"""
    try:
        import pyautogui
    except ImportError:
        fail("pyautoguiが未インストールです。")

    pyautogui.FAILSAFE = True
    print()
    print("3秒後にCline入力欄をクリックします。")
    print(f"座標: X={CLINE_INPUT_X}, Y={CLINE_INPUT_Y}")
    print("VS Codeを前面に表示してください。")
    time.sleep(3)
    focus_cline_input(pyautogui)
    print("クリックしました。Cline入力欄にカーソルがあるか確認してください。")
    return True


def inject_prompt_manual() -> None:
    print()
    print("=" * 60)
    print("Cline入力欄を手動で選択し、Ctrl+V、Enterを実行してください。")
    print("=" * 60)
    input("投入したらEnterを押してください...")



def snapshot_changes() -> str:
    """エージェント制御ファイルを除外したGit変更状態を返す。"""
    output = git("status", "--porcelain", check=False).stdout
    excluded = {".agent-question.json", ".agent-summary.json"}
    lines: list[str] = []

    for line in output.splitlines():
        path = line[3:].strip().strip('"').replace("\\", "/")
        if path in excluded:
            continue
        lines.append(line)

    return "\n".join(lines).strip()





# ============================================================
# Cline質問 → Teams回答 → Cline再開
# ============================================================
DECISION_TIMEOUT_SECONDS = 3600
DECISION_POLL_INTERVAL_SECONDS = 3


def question_paths() -> tuple[pathlib.Path, pathlib.Path, pathlib.Path, pathlib.Path, pathlib.Path]:
    agent_root = STATE_DIR.parent
    question_dir = pathlib.Path(
        os.getenv("AGENT_QUESTION_DIR", str(agent_root / "question"))
    )
    decision_dir = pathlib.Path(
        os.getenv("AGENT_DECISION_DIR", str(agent_root / "decision"))
    )
    question_done_dir = question_dir / "done"
    decision_done_dir = decision_dir / "done"
    question_file = TARGET_REPO / ".agent-question.json"

    for directory in (
        question_dir,
        decision_dir,
        question_done_dir,
        decision_done_dir,
    ):
        directory.mkdir(parents=True, exist_ok=True)

    return (
        question_file,
        question_dir,
        decision_dir,
        question_done_dir,
        decision_done_dir,
    )


def decode_question_json(path: pathlib.Path) -> dict | None:
    if not path.exists():
        return None

    try:
        raw = path.read_bytes()
    except OSError as error:
        log(f"質問ファイル読込失敗: {error}")
        return None

    for encoding in ("utf-8-sig", "utf-8", "cp932", "shift_jis"):
        try:
            value = json.loads(raw.decode(encoding))
            if isinstance(value, dict):
                return value
        except (UnicodeDecodeError, json.JSONDecodeError):
            continue

    log("質問JSONを解析できません。")
    return None


def read_question_file(issue_number: int) -> dict | None:
    question_file, _, _, _, _ = question_paths()
    if not question_file.exists():
        return None

    try:
        question_mtime = question_file.stat().st_mtime
    except OSError as error:
        log(f"質問ファイルの更新時刻を取得できません: {error}")
        return None

    if QUESTION_SESSION_STARTED_AT and question_mtime < QUESTION_SESSION_STARTED_AT:
        log("前回Issueの質問ファイルを検出したため処理しません。")
        return None
    question = decode_question_json(question_file)
    if not question:
        return None

    choices = question.get("choices")
    if not isinstance(choices, list) or not choices:
        log("質問JSONにchoicesがありません。")
        return None

    valid_choices = []
    for index, choice in enumerate(choices[:3], start=1):
        if not isinstance(choice, dict):
            continue
        choice_id = str(choice.get("id", f"choice_{index}")).strip()
        title = str(choice.get("title", choice_id)).strip()
        description = str(choice.get("description", "")).strip()
        valid_choices.append(
            {"id": choice_id, "title": title, "description": description}
        )

    if not valid_choices:
        log("有効な選択肢がありません。")
        return None

    question["type"] = "cline_question"

    source_issue_number = question.get("issueNumber")
    if source_issue_number is None:
        log("質問JSONにissueNumberがありません。")
        return None

    try:
        source_issue_number = int(source_issue_number)
    except (TypeError, ValueError):
        log(f"質問JSONのissueNumberが不正です: {source_issue_number}")
        return None

    if source_issue_number != issue_number:
        log(
            "別Issueの質問ファイルを検出したため処理しません。 "
            f"現在Issue=#{issue_number}, 質問Issue=#{source_issue_number}"
        )
        return None

    question["issueNumber"] = issue_number
    question["questionId"] = str(
        question.get("questionId") or f"issue-{issue_number}-question"
    )
    question["question"] = str(
        question.get("question") or "Clineから確認があります。"
    )
    question["choices"] = valid_choices
    question["messageId"] = load_message_id(issue_number)
    question["status"] = "waiting_decision"
    question["createdAt"] = datetime.now().isoformat()

    # Power Automateで扱いやすい平坦な値を追加
    for index in range(1, 4):
        choice = valid_choices[index - 1] if index <= len(valid_choices) else {}
        question[f"choice{index}Id"] = str(choice.get("id", ""))
        question[f"choice{index}Title"] = str(choice.get("title", ""))
        question[f"choice{index}Description"] = str(
            choice.get("description", "")
        )

    return question


def publish_question(issue_number: int, question: dict) -> pathlib.Path:
    _, question_dir, _, _, _ = question_paths()
    output_path = question_dir / f"issue-{issue_number}.json"
    output_path.write_text(
        json.dumps(question, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    update_state(
        issue_number,
        {
            "status": "waiting_decision",
            "questionId": question["questionId"],
            "questionFile": str(output_path),
        },
    )
    log(f"Teams質問用JSON出力: {output_path}")
    return output_path


def wait_for_decision(issue_number: int, question: dict) -> dict:
    _, _, decision_dir, _, _ = question_paths()
    decision_path = decision_dir / f"issue-{issue_number}.json"
    started_at = time.time()
    valid_ids = {str(choice["id"]) for choice in question["choices"]}

    log(f"Teamsからの回答を待機します: {decision_path}")

    while True:
        if time.time() - started_at > DECISION_TIMEOUT_SECONDS:
            fail("Teamsからの回答待機がタイムアウトしました。")

        if decision_path.exists():
            try:
                decision = json.loads(
                    decision_path.read_text(
                        encoding="utf-8-sig"
                    )
                )

            except json.JSONDecodeError as error:
                log(
                    f"decision JSON解析失敗: "
                    f"{decision_path} / {error}"
                )
                time.sleep(
                    DECISION_POLL_INTERVAL_SECONDS
                )
                continue

            except UnicodeDecodeError as error:
                log(
                    f"decision文字コード解析失敗: "
                    f"{decision_path} / {error}"
                )
                time.sleep(
                    DECISION_POLL_INTERVAL_SECONDS
                )
                continue

            except OSError as error:
                log(
                    f"decisionファイル読込失敗: "
                    f"{decision_path} / {error}"
                )
                time.sleep(
                    DECISION_POLL_INTERVAL_SECONDS
                )
                continue
            
            selected = str(
                decision.get("selectedChoice")
                or decision.get("choice")
                or decision.get("result")
                or ""
            ).strip()

            if selected in valid_ids:
                decision["selectedChoice"] = selected

                custom_response = str(
                    decision.get(
                        "customResponse",
                        ""
                    )
                ).strip()

                custom_response_base64 = str(
                    decision.get(
                        "customResponseBase64",
                        ""
                    )
                ).strip()

                if custom_response_base64:
                    try:
                        custom_response = (
                            base64.b64decode(
                                custom_response_base64,
                                validate=True,
                            )
                            .decode("utf-8")
                            .strip()
                        )

                    except (
                        ValueError,
                        UnicodeDecodeError,
                    ) as error:
                        log(
                            "customResponseBase64解析失敗: "
                            f"{error}"
                        )

                        time.sleep(
                            DECISION_POLL_INTERVAL_SECONDS
                        )

                        continue

                decision["customResponse"] = (
                    custom_response
                )

                return decision

            log(f"回答JSONの選択値が不正です: {selected}")

        time.sleep(DECISION_POLL_INTERVAL_SECONDS)


def build_decision_prompt(question: dict, decision: dict) -> str:
    selected_id = decision["selectedChoice"]
    selected_choice = next(
        choice
        for choice in question["choices"]
        if str(choice.get("id", "")) == selected_id
    )

    custom_response = str(
        decision.get("customResponse", "")
    ).strip()

    custom_section = ""
    if custom_response:
        custom_section = (
            "\n追加指示（Teams自由入力）:\n"
            f"{custom_response}\n"
        )

    return f"""Teamsから確認事項への回答がありました。

質問:
{question['question']}

選択結果:
{selected_choice.get('title', selected_id)}

選択内容:
{selected_choice.get('description', '')}

{custom_section}
この回答を確定事項として、同じIssueの実装を再開してください。
- 同じIssueの作業を継続してください
- Clineの質問UIは使用しないでください
- .agent-question.jsonを作り直さないでください
- git操作は行わないでください
- UTF-8 BOMなしで保存してください
- 最後に変更ファイルと実装内容を報告してください
"""


def archive_question_files(issue_number: int) -> None:
    question_file, question_dir, decision_dir, question_done, decision_done = (
        question_paths()
    )
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    for source, destination_dir in (
        (question_dir / f"issue-{issue_number}.json", question_done),
        (decision_dir / f"issue-{issue_number}.json", decision_done),
    ):
        if source.exists():
            destination = destination_dir / f"{timestamp}_{source.name}"
            shutil.move(str(source), str(destination))
            log(f"処理済みへ移動: {destination}")

    question_file.unlink(missing_ok=True)


def handle_cline_question(issue_number: int) -> bool:
    question = read_question_file(issue_number)
    if not question:
        return False

    publish_question(issue_number, question)
    decision = wait_for_decision(issue_number, question)
    answer_prompt = build_decision_prompt(question, decision)

    # 新規チャットを開かず、現在のCline入力欄へ回答を送る
    if not inject_prompt_gui(answer_prompt):
        copy_to_clipboard(answer_prompt)
        inject_prompt_manual()

    update_state(
        issue_number,
        {
            "status": "implementation_resumed",
            "selectedChoice": decision["selectedChoice"],
        },
    )
    archive_question_files(issue_number)
    log("Teamsの回答をClineへ返し、実装監視を再開しました。")
    return True


def wait_for_implementation(issue_number: int) -> str:
    log("実装完了を待機します。")
    started_at = time.time()
    previous = snapshot_changes()
    stable_since: float | None = None

    while True:
        if time.time() - started_at > IMPLEMENTATION_TIMEOUT_SECONDS:
            log("実装待機がタイムアウトしました。")
            break

        question_file, _, _, _, _ = question_paths()
        if question_file.exists():
            if handle_cline_question(issue_number):
                previous = snapshot_changes()
                stable_since = None
                continue

        time.sleep(POLL_INTERVAL_SECONDS)
        current = snapshot_changes()

        if current != previous:
            log("変更を検知しました。")
            previous = current
            stable_since = None
            continue

        if not current:
            continue

        if stable_since is None:
            stable_since = time.time()
            log(f"変更停止。{STABLE_DURATION_SECONDS}秒後に完了判定します。")
            continue

        if time.time() - stable_since >= STABLE_DURATION_SECONDS:
            log("実装完了を検知しました。")
            break

    return previous



def list_changed_files() -> list[str]:
    files = []
    for line in snapshot_changes().splitlines():
        path = line[3:].strip().strip('"')
        if path == ".agent-question.json":
            continue
        if path:
            files.append(path)
    return files


def expand_changed_files(changed_files: list[str]) -> list[str]:
    expanded: set[str] = set()
    for relative_path in changed_files:
        target = TARGET_REPO / relative_path.rstrip("/")
        if target.is_dir():
            for child in target.rglob("*"):
                if child.is_file():
                    expanded.add(child.relative_to(TARGET_REPO).as_posix())
        else:
            expanded.add(relative_path.replace("\\", "/"))
    return sorted(expanded)


def write_waiting_json(
    issue: dict,
    branch: str,
    changed_files: list[str],
    preview_url: str = "",
    summary_override: str = "",
) -> None:
    issue_number = issue["number"]
    expanded_files = expand_changed_files(changed_files)
    changed_files_text = "\n".join(
        f"・{path}" for path in expanded_files
    ) or "・変更なし"

    summary_text = "\n".join(
        f"・ファイルを追加または更新: {path}"
        for path in expanded_files
    ) or "・変更内容を確認してください。"

    if summary_override.strip():
        summary_text = summary_override.strip()

    message = (
        f"⏳ Issue #{issue_number} の実装が完了しました\n\n"
        f"タイトル:\n{issue['title']}\n\n"
        # f"ブランチ:\n{branch}\n\n"
        f"変更ファイル:\n{changed_files_text}\n\n"
        f"実装内容:\n{summary_text}\n\n"
        f"プレビュー:\n{preview_url}\n\n"
        "内容を確認して承認してください。"
    )

    payload = {
        "type": "waiting_approval",
        "messageId": load_message_id(issue_number),
        "issueNumber": issue_number,
        "issueTitle": issue["title"],
        "issueUrl": issue["url"],
        "branch": branch,
        "changedFiles": expanded_files,
        "changedFilesText": changed_files_text,
        "summaryText": summary_text,
        "warningText": "",

        # 追加
        "previewUrl": preview_url,

        "diffStat": git("diff", "--stat", check=False).stdout.strip(),
        "status": "waiting_approval",
        "message": message,
    }

    write_json(WAITING_DIR, f"issue-{issue_number}.json", payload)
    update_state(
        issue_number,
        {
            "status": "waiting_approval",
            "branch": branch,
            "changedFiles": expanded_files,
        },
    )
    notify_teams(
        "waiting_approval",
        issue_number,
        message,
        {
            "branch": branch,
            "changedFiles": expanded_files,
            "summaryText": summary_text,
        },
    )




# ============================================================
# 質問ファイルの世代管理・整合性検証
# ============================================================
QUESTION_SESSION_STARTED_AT: float = 0.0


def archive_existing_file(
    source: pathlib.Path,
    done_dir: pathlib.Path,
    label: str,
) -> None:
    if not source.exists():
        return

    done_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    destination = done_dir / f"{timestamp}_{source.name}"

    try:
        shutil.move(str(source), str(destination))
        log(f"古い{label}を退避: {destination}")
    except OSError as error:
        fail(f"古い{label}を退避できません: {source} / {error}")


def prepare_question_session(issue_number: int) -> None:
    """Issue開始前に、前回の質問と同一Issueの残存入出力を退避する。"""
    global QUESTION_SESSION_STARTED_AT

    agent_root = STATE_DIR.parent
    question_dir = pathlib.Path(
        os.getenv("AGENT_QUESTION_DIR", str(agent_root / "question"))
    )
    decision_dir = pathlib.Path(
        os.getenv("AGENT_DECISION_DIR", str(agent_root / "decision"))
    )
    question_done = question_dir / "done"
    decision_done = decision_dir / "done"

    for directory in (question_dir, decision_dir, question_done, decision_done):
        directory.mkdir(parents=True, exist_ok=True)

    archive_existing_file(
        TARGET_REPO / ".agent-question.json",
        question_done,
        ".agent-question.json",
    )
    archive_existing_file(
        question_dir / f"issue-{issue_number}.json",
        question_done,
        "Teams質問JSON",
    )
    archive_existing_file(
        decision_dir / f"issue-{issue_number}.json",
        decision_done,
        "Teams回答JSON",
    )

    QUESTION_SESSION_STARTED_AT = time.time()
    log(f"質問セッション開始: Issue #{issue_number}")



# ============================================================
# Cline最終メッセージ連携
# ============================================================
CLINE_SUMMARY_FILE_NAME = ".agent-summary.json"


def cline_summary_path() -> pathlib.Path:
    return TARGET_REPO / CLINE_SUMMARY_FILE_NAME


def prepare_cline_summary(issue_number: int) -> None:
    """前回Issueのサマリーを退避し、今回のIssue用に初期化する。"""
    summary_path = cline_summary_path()
    if not summary_path.exists():
        return

    done_dir = STATE_DIR.parent / "summary" / "done"
    done_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    destination = done_dir / (
        f"{timestamp}_issue-{issue_number}_{CLINE_SUMMARY_FILE_NAME.lstrip('.')}"
    )

    try:
        shutil.move(str(summary_path), str(destination))
        log(f"前回のClineサマリーを退避: {destination}")
    except OSError as error:
        fail(f"前回のClineサマリーを退避できません: {error}")


def read_cline_summary(issue_number: int) -> str:
    """Clineが出力した構造化サマリーをTeams表示用テキストへ変換する。"""
    summary_path = cline_summary_path()
    if not summary_path.exists():
        log("Clineサマリーファイルがないため、変更ファイルから要約します。")
        return ""

    try:
        raw = summary_path.read_bytes()
    except OSError as error:
        log(f"Clineサマリー読込失敗: {error}")
        return ""

    data = None
    for encoding in ("utf-8-sig", "utf-8", "cp932", "shift_jis"):
        try:
            value = json.loads(raw.decode(encoding))
            if isinstance(value, dict):
                data = value
                break
        except (UnicodeDecodeError, json.JSONDecodeError):
            continue

    if not data:
        log("ClineサマリーJSONを解析できません。")
        return ""

    source_issue = data.get("issueNumber")
    try:
        source_issue = int(source_issue)
    except (TypeError, ValueError):
        log("ClineサマリーのissueNumberが不正です。")
        return ""

    if source_issue != issue_number:
        log(
            "別IssueのClineサマリーを検出したため使用しません。 "
            f"現在Issue=#{issue_number}, サマリーIssue=#{source_issue}"
        )
        return ""

    def normalize_items(value) -> list[str]:
        if isinstance(value, list):
            return [str(item).strip() for item in value if str(item).strip()]
        if isinstance(value, str) and value.strip():
            return [line.strip().lstrip("-・ ") for line in value.splitlines() if line.strip()]
        return []

    sections = [
        ("実装内容", normalize_items(data.get("implementation"))),
        ("確認結果", normalize_items(data.get("verification"))),
        ("未実施事項", normalize_items(data.get("notPerformed"))),
        ("注意事項", normalize_items(data.get("notes"))),
    ]

    output: list[str] = []
    for title, items in sections:
        if not items:
            continue
        output.append(f"【{title}】")
        output.extend(f"・{item}" for item in items)
        output.append("")

    text = "\n".join(output).strip()
    if text:
        log("Clineの最終サマリーをsummaryTextへ反映します。")
    return text


def archive_cline_summary(issue_number: int) -> None:
    summary_path = cline_summary_path()
    if not summary_path.exists():
        return

    done_dir = STATE_DIR.parent / "summary" / "done"
    done_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    destination = done_dir / f"{timestamp}_issue-{issue_number}.json"
    try:
        shutil.move(str(summary_path), str(destination))
        log(f"Clineサマリーを処理済みへ移動: {destination}")
    except OSError as error:
        log(f"Clineサマリー移動失敗: {error}")

def main() -> None:
    parser = argparse.ArgumentParser(
        description="GitHub IssueをClineへ投入し、承認待ちJSONを出力します。"
    )
    parser.add_argument("issue", type=int, nargs="?", help="GitHub Issue番号")
    parser.add_argument("--base", default=os.getenv("BASE_BRANCH", "main"))
    parser.add_argument("--manual", action="store_true")
    parser.add_argument("--no-vscode", action="store_true")
    parser.add_argument("--message-id", default="")
    parser.add_argument(
        "--test-click",
        action="store_true",
        help="Issue処理をせずCline入力欄のクリックだけテストする",
    )
    args = parser.parse_args()

    if args.test_click:
        test_cline_click()
        return

    if args.issue is None:
        parser.error("Issue番号を指定してください。")

    validate_environment_variables()
    validate_directories()

    for command in ("git", "gh"):
        require_command(command)

    run_command(["gh", "auth", "status"], cwd=TARGET_REPO)

    log("=" * 60)
    log(f"Cline 実装エージェント: Issue #{args.issue}")
    log(f"Cline入力座標: X={CLINE_INPUT_X}, Y={CLINE_INPUT_Y}")
    log("新規チャット操作: 無効")
    log("=" * 60)

    issue = load_issue(args.issue)
    if issue.get("state") != "OPEN":
        fail(f"Issue #{args.issue} はOPENではありません。")

    if args.message_id:
        update_state(
            args.issue,
            {
                "issueTitle": issue["title"],
                "issueUrl": issue["url"],
                "messageId": args.message_id.strip(),
            },
        )

    branch = prepare_branch(args.issue, issue["title"], args.base)

    update_state(
        args.issue,
        {"status": "implementation_started", "branch": branch},
    )

    prepare_question_session(args.issue)
    notify_teams(
        "implementation_started",
        args.issue,
        (
            f"🚀 Issue #{args.issue} の実装を開始しました\n\n"
            f"タイトル:\n{issue['title']}\n\n"
            f"ブランチ:\n{branch}"
        ),
        {"branch": branch},
    )

    prepare_cline_summary(args.issue)


    prompt = build_prompt(issue, branch)
    copy_to_clipboard(prompt)

    if not args.no_vscode:
        launch_vscode()

    injected = False
    if not args.manual:
        injected = inject_prompt_gui(prompt)

    if not injected:
        inject_prompt_manual()

    wait_for_implementation(args.issue)
    changed_files = list_changed_files()

    if not changed_files:
        log("変更が検出されませんでした。")
        update_state(args.issue, {"status": "implementation_failed"})
        notify_teams(
            "implementation_failed",
            args.issue,
            f"⚠️ Issue #{args.issue} の実装で変更が検出されませんでした。",
            {"branch": branch},
        )
        return

    preview_url = deploy_preview(
        issue_number=args.issue,
        changed_files=expand_changed_files(changed_files),
    )

    cline_summary = read_cline_summary(args.issue)


    write_waiting_json(
        issue=issue,
        branch=branch,
        changed_files=changed_files,
        preview_url=preview_url,
        summary_override=cline_summary,
)
    archive_cline_summary(args.issue)
    log("Phase2 完了。承認待ちへ移行します。")


def deploy_preview(
    issue_number: int,
    changed_files: list[str],
) -> str:
    """deploy_preview.pyを実行し、公開されたGitHub Pages URLを返す。"""
    deploy_script = pathlib.Path(__file__).resolve().parent / "deploy_preview.py"

    if not deploy_script.exists():
        log(f"[WARNING] deploy_preview.pyが見つかりません: {deploy_script}")
        return build_preview_url(changed_files)

    command = [
        sys.executable,
        str(deploy_script),
        "--issue",
        str(issue_number),
    ]

    for changed_file in changed_files:
        command.extend(["--changed-file", changed_file])

    log("tools-betaへ承認前プレビューを公開します。")
    result = run_command(
        command,
        cwd=deploy_script.parent,
        check=False,
    )

    if result.returncode != 0:
        log(
            "[WARNING] tools-betaへの公開に失敗しました。"
            "URL生成のみ継続します。"
        )
        return build_preview_url(changed_files)

    try:
        deploy_result = json.loads(result.stdout)
    except json.JSONDecodeError as error:
        log(f"[WARNING] deploy_preview.pyの出力JSONを解析できません: {error}")
        return build_preview_url(changed_files)

    preview_url = str(deploy_result.get("previewUrl", "")).strip()
    status = str(deploy_result.get("status", "")).strip()

    if preview_url:
        log(f"tools-beta公開結果: {status}")
        log(f"プレビューURL: {preview_url}")
        return preview_url

    log("[WARNING] previewUrlが空です。URL生成のみ継続します。")
    return build_preview_url(changed_files)

def build_preview_url(
    changed_files: list[str],
) -> str:

    base_url = (
        "https://ymatsuda-cmyk.github.io/"
        "tools-beta"
    )

    for path in changed_files:

        normalized = (
            path.replace("\\", "/")
        )

        parts = normalized.split("/")

        if (
            len(parts) >= 2
            and parts[-1].lower()
            == "index.html"
        ):

            directory = "/".join(
                parts[:-1]
            )

            return (
                f"{base_url}/"
                f"{directory}/"
            )

    return f"{base_url}/"


if __name__ == "__main__":
    main()
