#!/usr/bin/env python3
"""声紋登録 + 話者自動認識

- Whisperのセグメント区間ごとに声紋ベクトル（埋め込み）を作る
- 同一ファイル内でクラスタリングして話者をまとめる
- 声紋DB（JSON）と照合し、既知話者なら登録名、未知なら「担当者N」で新規登録

依存: resemblyzer（内部でtorch/numpyを使用）
    pip install resemblyzer
未インストールの場合は available() が False を返すので、
呼び出し側はLLMによるテキストベース話者推定にフォールバックする。
"""
import importlib.util
import json
import os
import shutil
import subprocess
import tempfile
import wave
from datetime import datetime, timezone
from pathlib import Path

_encoder = None


def available() -> bool:
    return importlib.util.find_spec("resemblyzer") is not None and \
           importlib.util.find_spec("numpy") is not None


def _get_encoder():
    global _encoder
    if _encoder is None:
        from resemblyzer import VoiceEncoder
        _encoder = VoiceEncoder()
    return _encoder


def _resolve_ffmpeg():
    ffmpeg = shutil.which("ffmpeg")
    if ffmpeg:
        return ffmpeg
    for c in ("/opt/homebrew/bin/ffmpeg", "/usr/local/bin/ffmpeg"):
        if Path(c).exists() and os.access(c, os.X_OK):
            return c
    return None


def _load_wav16k(audio_path):
    """音声を16kHzモノラルに変換してfloat32のnumpy配列で返す"""
    import numpy as np

    ffmpeg = _resolve_ffmpeg()
    if not ffmpeg:
        raise RuntimeError("ffmpeg が見つかりません")

    with tempfile.TemporaryDirectory() as td:
        wav_path = Path(td) / "audio16k.wav"
        cmd = [ffmpeg, "-y", "-i", str(audio_path), "-ac", "1", "-ar", "16000",
               "-c:a", "pcm_s16le", str(wav_path)]
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=1800)
        if res.returncode != 0 or not wav_path.exists():
            raise RuntimeError(f"ffmpeg変換に失敗: {(res.stderr or '')[:200]}")
        with wave.open(str(wav_path), "rb") as wf:
            frames = wf.readframes(wf.getnframes())
        pcm = np.frombuffer(frames, dtype=np.int16).astype(np.float32) / 32768.0
    return pcm


def _cosine(a, b):
    import numpy as np
    na, nb = np.linalg.norm(a), np.linalg.norm(b)
    if na == 0 or nb == 0:
        return -1.0
    return float(np.dot(a, b) / (na * nb))


def _embed_segments(pcm, segments, min_sec, pad_sec=0.25):
    """各セグメントの声紋ベクトルを作る。短すぎる区間は None。"""
    import numpy as np
    encoder = _get_encoder()
    sr = 16000
    embeddings = []
    for seg in segments:
        start = max(0.0, float(seg["start"]) - pad_sec)
        end = min(len(pcm) / sr, float(seg["end"]) + pad_sec)
        if end - start < min_sec:
            embeddings.append(None)
            continue
        chunk = pcm[int(start * sr): int(end * sr)]
        if chunk.size < int(min_sec * sr):
            embeddings.append(None)
            continue
        try:
            embeddings.append(np.asarray(encoder.embed_utterance(chunk), dtype=np.float32))
        except Exception:
            embeddings.append(None)
    return embeddings


def _cluster(embeddings, threshold):
    """貪欲法でクラスタリング（sklearn不要）。戻り値: セグメント→クラスタID"""
    import numpy as np
    centroids, counts, labels = [], [], []
    for emb in embeddings:
        if emb is None:
            labels.append(-1)
            continue
        best_i, best_sim = -1, -1.0
        for i, c in enumerate(centroids):
            sim = _cosine(emb, c)
            if sim > best_sim:
                best_i, best_sim = i, sim
        if best_i >= 0 and best_sim >= threshold:
            n = counts[best_i]
            centroids[best_i] = (centroids[best_i] * n + emb) / (n + 1)
            counts[best_i] = n + 1
            labels.append(best_i)
        else:
            centroids.append(emb.copy())
            counts.append(1)
            labels.append(len(centroids) - 1)
    # 未判定セグメントは前後のクラスタで補完
    for i, lb in enumerate(labels):
        if lb != -1:
            continue
        prev = next((labels[j] for j in range(i - 1, -1, -1) if labels[j] != -1), -1)
        nxt = next((labels[j] for j in range(i + 1, len(labels)) if labels[j] != -1), -1)
        labels[i] = prev if prev != -1 else (nxt if nxt != -1 else 0)
    if not centroids:
        return [0] * len(embeddings), [], []
    return labels, centroids, counts


def _load_db(db_path):
    p = Path(db_path).expanduser()
    if not p.exists():
        return {"speakers": []}
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        if isinstance(data, dict) and isinstance(data.get("speakers"), list):
            return data
    except (json.JSONDecodeError, OSError):
        pass
    return {"speakers": []}


def _save_db(db_path, db):
    p = Path(db_path).expanduser()
    p.parent.mkdir(parents=True, exist_ok=True)
    tmp = p.with_name(p.name + ".tmp")
    tmp.write_text(json.dumps(db, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    tmp.replace(p)


def _next_auto_name(db, prefix):
    used = set()
    for sp in db["speakers"]:
        name = sp.get("name", "")
        if name.startswith(prefix) and name[len(prefix):].isdigit():
            used.add(int(name[len(prefix):]))
    n = 1
    while n in used:
        n += 1
    return f"{prefix}{n}"


def name_labeled_speakers(audio_path, segments, db_path,
                          match_threshold=0.72, name_prefix="担当者",
                          max_sample_sec=30.0, min_segment_sec=0.7):
    """既に話者分離済み（SPEAKER_00 等）のセグメントに、声紋DBの名前を割り当てる。

    Qwen3-ASR + pyannote のように分離自体は外部で終わっている場合に使う。
    クラスタリングは行わず、ラベルごとに1つ声紋を作って照合するだけなので速い。
    """
    import numpy as np

    labels = [s.get("speaker") for s in segments if s.get("speaker")]
    if not labels:
        return segments, {}

    pcm = _load_wav16k(audio_path)
    encoder = _get_encoder()
    sr = 16000

    # ラベルごとに音声を連結（長い会議でも上限秒数までで十分）
    buckets = {}
    for seg in segments:
        lb = seg.get("speaker")
        if not lb:
            continue
        start, end = float(seg["start"]), float(seg["end"])
        if end - start < min_segment_sec:
            continue
        cur = buckets.setdefault(lb, [])
        total = sum(c.size for c in cur) / sr
        if total >= max_sample_sec:
            continue
        cur.append(pcm[int(start * sr): int(end * sr)])

    centroids = {}
    for lb, chunks in buckets.items():
        if not chunks:
            continue
        wav = np.concatenate(chunks)
        if wav.size < int(min_segment_sec * sr):
            continue
        try:
            centroids[lb] = np.asarray(encoder.embed_utterance(wav), dtype=np.float32)
        except Exception:  # noqa: BLE001
            continue
    if not centroids:
        raise RuntimeError("有効な声紋を抽出できませんでした")

    db = _load_db(db_path)
    known = [(sp, np.asarray(sp["embedding"], dtype=np.float32))
             for sp in db["speakers"] if sp.get("embedding")]

    mapping, used = {}, set()
    for lb, centroid in centroids.items():
        best_sp, best_sim = None, -1.0
        for sp, emb in known:
            sim = _cosine(centroid, emb)
            if sim > best_sim:
                best_sp, best_sim = sp, sim
        if best_sp is not None and best_sim >= match_threshold and best_sp["name"] not in used:
            n = min(int(best_sp.get("count", 1)), 20)
            old = np.asarray(best_sp["embedding"], dtype=np.float32)
            best_sp["embedding"] = ((old * n + centroid) / (n + 1)).tolist()
            best_sp["count"] = int(best_sp.get("count", 1)) + 1
            best_sp["updatedAt"] = datetime.now(timezone.utc).isoformat()
            mapping[lb] = best_sp["name"]
        else:
            new_name = _next_auto_name(db, name_prefix)
            db["speakers"].append({
                "name": new_name,
                "embedding": centroid.tolist(),
                "count": 1,
                "auto": True,
                "createdAt": datetime.now(timezone.utc).isoformat(),
                "updatedAt": datetime.now(timezone.utc).isoformat(),
            })
            known.append((db["speakers"][-1], centroid))
            mapping[lb] = new_name
        used.add(mapping[lb])

    _save_db(db_path, db)

    out = []
    for seg in segments:
        s = dict(seg)
        lb = s.get("speaker")
        if lb and lb in mapping:
            s["speaker"] = mapping[lb]
        out.append(s)
    return out, mapping


def assign_speakers(audio_path, segments, db_path,
                    cluster_threshold=0.62, match_threshold=0.72,
                    min_segment_sec=0.7, name_prefix="担当者"):
    """segments（start/end/text）に speaker を付与して返す。

    声紋DBに一致する話者がいればその名前、いなければ「担当者N」で新規登録する。
    DBのnameを手で「松田」等に書き換えれば、以降はその名前が使われる。
    """
    import numpy as np

    if not segments:
        return segments, {}

    pcm = _load_wav16k(audio_path)
    embeddings = _embed_segments(pcm, segments, min_segment_sec)
    if all(e is None for e in embeddings):
        raise RuntimeError("有効な声紋を抽出できませんでした")

    labels, centroids, counts = _cluster(embeddings, cluster_threshold)

    db = _load_db(db_path)
    known = []
    for sp in db["speakers"]:
        emb = sp.get("embedding")
        if emb:
            known.append((sp, np.asarray(emb, dtype=np.float32)))

    cluster_names, used_names = {}, set()
    for ci, centroid in enumerate(centroids):
        best_sp, best_sim = None, -1.0
        for sp, emb in known:
            sim = _cosine(centroid, emb)
            if sim > best_sim:
                best_sp, best_sim = sp, sim
        if best_sp is not None and best_sim >= match_threshold and best_sp["name"] not in used_names:
            # 既知話者：声紋を移動平均で更新して精度を上げる
            n = min(int(best_sp.get("count", 1)), 20)
            old = np.asarray(best_sp["embedding"], dtype=np.float32)
            best_sp["embedding"] = ((old * n + centroid) / (n + 1)).tolist()
            best_sp["count"] = int(best_sp.get("count", 1)) + 1
            best_sp["updatedAt"] = datetime.now(timezone.utc).isoformat()
            cluster_names[ci] = best_sp["name"]
        else:
            new_name = _next_auto_name(db, name_prefix)
            db["speakers"].append({
                "name": new_name,
                "embedding": centroid.tolist(),
                "count": 1,
                "auto": True,
                "createdAt": datetime.now(timezone.utc).isoformat(),
                "updatedAt": datetime.now(timezone.utc).isoformat(),
            })
            known.append((db["speakers"][-1], centroid))
            cluster_names[ci] = new_name
        used_names.add(cluster_names[ci])

    _save_db(db_path, db)

    out = []
    for seg, lb in zip(segments, labels):
        s = dict(seg)
        s["speaker"] = cluster_names.get(lb, f"{name_prefix}1")
        out.append(s)
    return out, cluster_names
