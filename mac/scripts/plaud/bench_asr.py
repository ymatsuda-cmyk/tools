#!/usr/bin/env python3
"""whisper と Qwen3-ASR を同じ音声で比較する検証スクリプト

使い方:
    python3 bench_asr.py 音声ファイル
    python3 bench_asr.py 音声ファイル --labels whisper-large qwen3
    python3 bench_asr.py 音声ファイル --out ./bench_out

各labelについて 所要時間 / RTF / 文字数 を表示し、
出力テキストを --out ディレクトリに保存して目視比較できるようにする。
"""
import argparse
import json
import shutil
import subprocess
import sys
import time
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import transcribe_engines as TE  # noqa: E402


def audio_duration(path):
    ffprobe = shutil.which("ffprobe") or "/opt/homebrew/bin/ffprobe"
    try:
        res = subprocess.run(
            [ffprobe, "-v", "error", "-show_entries", "format=duration",
             "-of", "default=noprint_wrappers=1:nokey=1", str(path)],
            capture_output=True, text=True, timeout=60)
        return float(res.stdout.strip())
    except (OSError, ValueError, subprocess.SubprocessError):
        return 0.0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("audio")
    ap.add_argument("--labels", nargs="+", help="比較するlabel（既定: settings.json の全エンジン）")
    ap.add_argument("--settings")
    ap.add_argument("--out", default="./bench_out")
    args = ap.parse_args()

    audio = Path(args.audio).expanduser()
    if not audio.exists():
        print(f"❌ 音声ファイルが見つかりません: {audio}")
        return 1

    settings = TE.load_settings(args.settings)
    labels = args.labels or [e.get("label") for e in settings["engines"]]
    out_dir = Path(args.out).expanduser()
    out_dir.mkdir(parents=True, exist_ok=True)

    dur = audio_duration(audio)
    print(f"\n対象: {audio.name}  ({dur:.1f}秒 / {audio.stat().st_size/1024/1024:.1f} MB)")
    print("=" * 72)

    results = []
    for label in labels:
        print(f"\n▶ {label}")
        # 各labelを独立のtmpディレクトリで実行し、中間ファイルの混線を防ぐ
        work = out_dir / f"work_{label}"
        work.mkdir(parents=True, exist_ok=True)
        target = work / audio.name
        if not target.exists():
            shutil.copy2(audio, target)

        t0 = time.time()
        try:
            text = TE.transcribe(str(target), label=label, settings=settings)
        except Exception as e:  # noqa: BLE001
            print(f"  ❌ 失敗: {type(e).__name__}: {e}")
            results.append({"label": label, "ok": False, "error": f"{type(e).__name__}: {e}"})
            continue
        elapsed = time.time() - t0

        if not text:
            print("  ❌ 出力なし")
            results.append({"label": label, "ok": False, "error": "empty"})
            continue

        rtf = (elapsed / dur) if dur else 0.0
        speakers = sorted({ln.split("] ", 1)[1].split(":", 1)[0]
                           for ln in text.splitlines()
                           if "] " in ln and ":" in ln.split("] ", 1)[1]})
        print(f"  ⏱ {elapsed:.1f}秒  RTF={rtf:.3f}  {len(text)}文字"
              + (f"  話者={', '.join(speakers)}" if speakers else ""))

        txt_path = out_dir / f"{label}.txt"
        txt_path.write_text(text, encoding="utf-8")
        print(f"  📄 {txt_path}")
        print("  --- 冒頭300文字 ---")
        for line in text[:300].splitlines():
            print(f"  | {line}")

        results.append({"label": label, "ok": True, "elapsed": round(elapsed, 1),
                        "rtf": round(rtf, 3), "chars": len(text), "speakers": speakers})

    print("\n" + "=" * 72)
    print(f"{'label':<16}{'時間(s)':>10}{'RTF':>8}{'文字数':>10}  話者")
    for r in results:
        if r.get("ok"):
            print(f"{r['label']:<16}{r['elapsed']:>10.1f}{r['rtf']:>8.3f}{r['chars']:>10}"
                  f"  {', '.join(r.get('speakers') or [])}")
        else:
            print(f"{r['label']:<16}{'-':>10}{'-':>8}{'-':>10}  ❌ {r.get('error', '')}")

    summary = out_dir / "summary.json"
    summary.write_text(json.dumps({"audio": str(audio), "duration": dur, "results": results},
                                  ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n📄 {summary}\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
