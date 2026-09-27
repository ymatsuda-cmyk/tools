#!/usr/bin/env python3

import argparse
import json
import urllib.error
import urllib.request
from pathlib import Path

API_BASE = "https://api.jsonbin.io/v3/b"


def call(method, url, headers, body=None):
    """JSONBinを叩く。失敗してもここでは投げず、ステータスとボディをそのまま返す。"""

    data = (
        json.dumps(body).encode("utf-8")
        if body is not None
        else None
    )

    req = urllib.request.Request(
        url,
        data=data,
        headers=headers,
        method=method
    )

    try:
        with urllib.request.urlopen(req, timeout=30) as res:
            raw = res.read().decode("utf-8")
            return res.status, (json.loads(raw) if raw else None)

    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", errors="replace")
        try:
            parsed = json.loads(raw) if raw else None
        except json.JSONDecodeError:
            parsed = raw
        return e.code, parsed


def fetch_current(bin_id, api_key):
    """いま登録されている内容を取る。Binがまだ空/存在しなければ None。"""

    status, body = call(
        "GET",
        f"{API_BASE}/{bin_id}/latest",
        {
            "X-Master-Key": api_key,
            "X-Bin-Meta": "false"
        }
    )

    if status == 404:
        return None

    if status >= 300:
        raise RuntimeError(
            f"JSONBinの取得に失敗しました（HTTP {status}）: {body}"
        )

    return body


def push(bin_id, api_key, payload):
    """Binの中身を丸ごと置き換える。追記ではなく上書きなので、
    payload には常に「いま登録したい内容の全体」を渡すこと。"""

    status, body = call(
        "PUT",
        f"{API_BASE}/{bin_id}",
        {
            "Content-Type": "application/json",
            "X-Master-Key": api_key
        },
        payload
    )

    if status >= 300:
        raise RuntimeError(
            f"JSONBinへの登録に失敗しました（HTTP {status}）: {body}"
        )


def sync_config(config_path):

    config = json.loads(
        Path(config_path).read_text(encoding="utf-8")
    )

    if not config.get("enabled", True):
        print("無効設定のためスキップ")
        return

    source_file = Path(config["source_file"])

    bin_id = config["bin_id"]
    api_key = config["api_key"]

    print(f"同期開始 : {config['id']}")

    if not source_file.exists():
        raise FileNotFoundError(
            f"元ファイルが見つかりません: {source_file}"
        )

    #
    # 登録したい内容(ファイルの中身がそのままBinの中身になる)
    #
    payload = json.loads(
        source_file.read_text(encoding="utf-8")
    )

    #
    # 差分確認
    #
    # github_sync.py の「差分なし」と同じく、変わっていなければAPIは叩かない
    #
    current = fetch_current(bin_id, api_key)

    if current == payload:
        print("差分なし")
        return

    #
    # 登録
    #
    push(bin_id, api_key, payload)

    print(f"登録完了 : {source_file.name} -> bin {bin_id}")


def main():

    parser = argparse.ArgumentParser()

    parser.add_argument(
        "--config",
        required=True
    )

    args = parser.parse_args()

    sync_config(args.config)


if __name__ == "__main__":
    main()
