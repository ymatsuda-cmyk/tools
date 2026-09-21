import base64
import json
import pathlib
import subprocess
import time

APPROVAL_DIR = pathlib.Path(
    r"C:\Users\matsuda\OneDrive - 株式会社日本ビジネスアシスト\work\agent\approval"
)

DONE_DIR = APPROVAL_DIR / "done"

DONE_DIR.mkdir(
    parents=True,
    exist_ok=True
)


def decode_rework_comment(
    encoded: str,
) -> str:
    if not encoded:
        return ""

    return (
        base64.b64decode(
            encoded,
            validate=True,
        )
        .decode("utf-8")
        .strip()
    )


while True:

    for file_path in APPROVAL_DIR.glob(
        "issue-*.json"
    ):

        try:

            data = json.loads(
                file_path.read_text(
                    encoding="utf-8"
                )
            )

            issue_number = int(
                data["issueNumber"]
            )

            action = str(
                data.get(
                    "action",
                    ""
                )
            ).lower()

            print(
                f"[APPROVAL] "
                f"Issue #{issue_number} "
                f"Action={action}"
            )

            #
            # 承認
            #
            if action == "approve":

                subprocess.Popen(
                    [
                        "python",
                        "finish_agent.py",
                        str(issue_number),
                        "--approve"
                    ]
                )

            #
            # 却下
            #
            elif action == "reject":

                print(
                    f"[REJECT] "
                    f"Issue #{issue_number}"
                )

            #
            # 再実装
            #
            elif action == "rework":

                comment = decode_rework_comment(
                    data.get(
                        "reworkCommentBase64",
                        ""
                    )
                )

                print(
                    f"[REWORK] "
                    f"Issue #{issue_number}"
                )

                print(comment)

                rework_file = pathlib.Path(
                    r"C:\repo\tools"
                ) / ".agent-rework.txt"

                rework_file.write_text(
                    comment,
                    encoding="utf-8"
                )

                subprocess.Popen(
                    [
                        "python",
                        "run_agent.py",
                        str(issue_number),
                        "--no-vscode"
                    ]
                )

            else:

                print(
                    f"[WARNING] "
                    f"Unknown action: "
                    f"{action}"
                )

            target = (
                DONE_DIR
                / file_path.name
            )

            file_path.rename(target)

        except Exception as error:

            print(error)

    time.sleep(3)