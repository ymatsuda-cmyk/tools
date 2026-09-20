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

            result = data["result"]

            if result == "approve":

                subprocess.Popen(
                    [
                        "python",
                        "finish_agent.py",
                        str(issue_number),
                        "--approve"
                    ]
                )

            target = (
                DONE_DIR
                / file_path.name
            )

            file_path.rename(target)

        except Exception as error:

            print(error)

    time.sleep(3)