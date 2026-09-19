import os
import json
import pathlib
import requests
from bs4 import BeautifulSoup

TOKEN = os.environ["GITHUB_TOKEN"]
OWNER = os.environ["GITHUB_OWNER"]
REPO = os.environ["GITHUB_REPO"]
WATCH_DIR = pathlib.Path(
    os.environ["AGENT_REQUEST_DIR"]
)

for file in WATCH_DIR.glob("*.json"):

    data = json.loads(
        file.read_text(
            encoding="utf-8"
        )
    )

    text = BeautifulSoup(
        data["message"],
        "html.parser"
    ).get_text("\n")

    title = text.split("\n")[0].strip()

    body = f"""
送信者:
{data['sender']}

Teams ID:
{data['id']}

内容:

{text}
"""

    response = requests.post(
        f"https://api.github.com/repos/{OWNER}/{REPO}/issues",
        headers={
            "Authorization": f"Bearer {TOKEN}",
            "Accept": "application/vnd.github+json"
        },
        json={
            "title": title,
            "body": body
        }
    )

    print(response.status_code)

    if response.status_code == 201:
        print("Issue作成成功")
        file.unlink()
    else:
        print(response.text)