# -*- coding: utf-8 -*-
"""dashrun:// URL スキームのハンドラ

ブラウザから  dashrun://launch?path=<URLエンコードした exe パス>  を受け取り、
launch_core で検証・許可確認をしてから exe を起動する。
受け付けるのは launch と path だけ。それ以外のパラメータが付いていたら起動しない。
"""
import os
import sys
from urllib.parse import parse_qs, urlsplit

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import launch_core as core  # noqa: E402


def parse_dashrun_url(url):
    """exe パスを返す。形式が違えば ValueError"""
    parts = urlsplit(url.strip())
    if parts.scheme.lower() != "dashrun":
        raise ValueError("dashrun:// 以外の URL です")
    if parts.netloc.lower() != "launch" or parts.path not in ("", "/"):
        raise ValueError("未対応の操作です")
    if parts.fragment:
        raise ValueError("URL の形式が正しくありません")
    query = parse_qs(parts.query, keep_blank_values=True)
    if set(query) != {"path"} or len(query["path"]) != 1:
        raise ValueError("path 以外のパラメータは受け付けません")
    return query["path"][0]


def main(argv):
    # ブラウザからは "dashrun://..." が 1 つだけ渡される。2つ目以降は使わない
    if len(argv) < 2:
        core.show_error("dashrun はブラウザのリンクから呼び出して使います。")
        return 2
    url = argv[1]
    try:
        path = parse_dashrun_url(url)
    except ValueError as e:
        core.log(f"BADURL {url!r}: {e}")
        core.show_error(f"起動できません: {e}")
        return 2
    result = core.handle_launch_request(path)
    if result["status"] in ("invalid", "error"):
        core.show_error(f"{result['message']}\n\n{result['path']}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
