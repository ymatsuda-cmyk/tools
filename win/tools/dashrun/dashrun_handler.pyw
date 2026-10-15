# -*- coding: utf-8 -*-
"""dashrun:// URL スキームのハンドラ

launch と openmail の2種類の操作を受け付ける。
"""
import os
import sys
from urllib.parse import parse_qs, urlsplit

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import launch_core as core  # noqa: E402


def parse_dashrun_url(url):
    parts = urlsplit(url.strip())
    if parts.scheme.lower() != "dashrun":
        raise ValueError("dashrun:// 以外の URL です")
    if parts.fragment:
        raise ValueError("URL の形式が正しくありません")
    if parts.path not in ("", "/"):
        raise ValueError("未対応の操作です")
    action = parts.netloc.lower()
    query = parse_qs(parts.query, keep_blank_values=True)

    if action == "launch":
        if set(query) != {"path"} or len(query["path"]) != 1:
            raise ValueError("path 以外のパラメータは受け付けません")
        return ("launch", query["path"][0])
    if action == "openmail":
        if set(query) != {"account", "url"} or len(query["account"]) != 1 or len(query["url"]) != 1:
            raise ValueError("account と url 以外のパラメータは受け付けません")
        return ("openmail", query["account"][0], query["url"][0])
    raise ValueError("未対応の操作です")


def main(argv):
    if len(argv) < 2:
        core.show_error("dashrun はブラウザのリンクから呼び出して使います。")
        return 2
    url = argv[1]
    try:
        parsed = parse_dashrun_url(url)
    except ValueError as e:
        core.log(f"BADURL {url!r}: {e}")
        core.show_error(f"起動できません: {e}")
        return 2

    if parsed[0] == "launch":
        result = core.handle_launch_request(parsed[1])
    else:
        result = core.handle_openmail_request(parsed[1], parsed[2])

    if result["status"] in ("invalid", "error"):
        core.show_error(result["message"])
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))