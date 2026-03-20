"""
B 站视频 URL → 语音识别 → 保存 txt + json

用法:
  python bilibili_transcribe.py "https://www.bilibili.com/video/BVxxxxxx"
  python bilibili_transcribe.py "https://www.bilibili.com/video/BVxxxxxx" --output-dir ./out --keep-audio
  python bilibili_transcribe.py "https://b23.tv/xxxx" --model base

依赖: pip install yt-dlp torch whisper
      系统需安装 ffmpeg（yt-dlp 转音频用）
"""

import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path


def check_dependencies():
    missing = []
    for pkg in ["torch", "whisper", "yt_dlp"]:
        try:
            __import__(pkg.replace("-", "_"))
        except ImportError:
            missing.append(pkg)
    if missing:
        print(f"[错误] 缺少依赖，请运行: pip install {' '.join(missing)}")
        sys.exit(1)
    # ffmpeg 用于 yt-dlp -x 转码
    try:
        subprocess.run(["ffmpeg", "-version"], capture_output=True, check=True)
    except (FileNotFoundError, subprocess.CalledProcessError):
        print("[错误] 未找到 ffmpeg，请先安装并加入 PATH")
        sys.exit(1)


def parse_bilibili_url(url: str) -> str | None:
    """从 B 站 URL 中解析出视频 id（BV 号或 av 等），用于输出文件名。"""
    url = url.strip()
    # https://www.bilibili.com/video/BV1xx...
    m = re.search(r"bilibili\.com/video/(BV[\w]+)", url, re.I)
    if m:
        return m.group(1)
    # https://b23.tv/xxx 会重定向，yt-dlp 可解析，这里用占位名
    if "b23.tv" in url or "bilibili.com" in url:
        return "bilibili_video"
    return None


def download_audio(url: str, output_path: Path) -> Path:
    """使用 yt-dlp 下载 B 站视频的音频，转为 wav 保存到 output_path。"""
    print("[1/3] 使用 yt-dlp 下载音频...")
    sys.stdout.flush()
    cmd = [
        sys.executable, "-m", "yt_dlp",
        "-x",
        "-f", "bestaudio",
        "--audio-format", "wav",
        "-o", str(output_path.with_suffix(".%(ext)s")),
        "--no-playlist",
        url,
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if result.returncode != 0:
        print(f"[错误] yt-dlp 下载失败:\n{result.stderr[-800:]}")
        sys.exit(1)
    # yt-dlp -o path.with_suffix(".%(ext)s") 会生成 path.wav（因为 --audio-format wav）
    wav = output_path.with_suffix(".wav")
    if not wav.exists():
        # 可能输出的是 m4a 等，用 ffmpeg 转 wav
        for f in output_path.parent.iterdir():
            if f.suffix.lower() in (".m4a", ".webm", ".opus", ".mp3") and f.stem == output_path.stem:
                subprocess.run([
                    "ffmpeg", "-y", "-i", str(f),
                    "-acodec", "pcm_s16le", "-ar", "16000", "-ac", "1",
                    str(wav)
                ], capture_output=True, check=True)
                f.unlink(missing_ok=True)
                break
        else:
            print("[错误] 未找到下载的音频文件")
            sys.exit(1)
    print(f"       已保存 -> {wav.name}")
    sys.stdout.flush()
    return wav


def transcribe_with_whisper(
    audio_path: Path,
    model_name: str,
    language: str | None,
) -> list[dict]:
    """Whisper 转录，返回 [{"start", "end", "text"}, ...]。

    始终使用 task="transcribe"，不做翻译。
    - language 为 None 时：自动识别语种（适合中英混合）。
    - language 为 "en"/"zh" 时：强制按英文/中文识别，避免自动识别出错。
    """
    import whisper

    lang_label = {
        None: "自动识别语言",
        "en": "英文",
        "zh": "中文",
    }.get(language, f"language={language}")

    print(f"[2/3] 转录中（Whisper {model_name}，{lang_label}，仅转写不翻译）...")
    sys.stdout.flush()
    model = whisper.load_model(model_name)
    result = model.transcribe(
        str(audio_path),
        language=language,
        task="transcribe",
        verbose=False,
        fp16=False,
    )
    segments = [
        {"start": seg["start"], "end": seg["end"], "text": seg["text"].strip()}
        for seg in result["segments"]
        if seg["text"].strip()
    ]
    total_chars = sum(len(s["text"]) for s in segments)
    print(f"       共 {len(segments)} 句，{total_chars} 字")
    sys.stdout.flush()
    return segments


def save_outputs(segments: list[dict], output_dir: Path, stem: str) -> tuple[Path, Path]:
    """写入 xxx_transcript.txt 和 xxx_segments.json，返回 (txt_path, json_path)。"""
    txt_path = output_dir / f"{stem}_transcript.txt"
    json_path = output_dir / f"{stem}_segments.json"
    lines = []
    for s in segments:
        lines.append(f"[{s['start']:.2f}-{s['end']:.2f}]")
        lines.append(s["text"])
        lines.append("")  # blank line between segments
    full_text = "\n".join(lines).rstrip("\n") + "\n"
    txt_path.write_text(full_text, encoding="utf-8")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(segments, f, ensure_ascii=False, indent=2)
    print(f"[3/3] 已保存 -> {txt_path.name}, {json_path.name}")
    sys.stdout.flush()
    return txt_path, json_path


def main():
    parser = argparse.ArgumentParser(
        description="输入 B 站视频网址，识别语音并保存为 txt 与 json。"
    )
    parser.add_argument("url", help="B 站视频 URL（如 https://www.bilibili.com/video/BVxxxxxx）")
    parser.add_argument("--output-dir", "-o", type=Path, default=Path("."),
                        help="输出目录，默认当前目录")
    parser.add_argument("--model", "-m", default="medium",
                        help="Whisper 模型: tiny/base/small/medium/large，默认 medium")
    parser.add_argument(
        "--lang",
        choices=["auto", "zh", "en"],
        default="auto",
        help="识别语言: auto=自动识别(默认), zh=按中文识别, en=按英文识别",
    )
    parser.add_argument("--keep-audio", action="store_true",
                        help="保留下载的音频文件，不删除")
    args = parser.parse_args()

    check_dependencies()
    video_id = parse_bilibili_url(args.url)
    if not video_id:
        print("[错误] 无法从 URL 解析出视频 ID，请使用标准 B 站视频链接")
        sys.exit(1)

    args.output_dir.mkdir(parents=True, exist_ok=True)
    work_dir = args.output_dir
    date_prefix = datetime.now().strftime("%Y_%m_%d")
    stem = f"{date_prefix}_{video_id}"

    # 始终把音频保存在输出目录下，便于你后续复用/检查
    audio_path = download_audio(args.url, work_dir / f"{stem}_audio")

    lang_param: str | None
    if args.lang == "auto":
        lang_param = None
    else:
        lang_param = args.lang

    segments = transcribe_with_whisper(audio_path, args.model, lang_param)
    save_outputs(segments, work_dir, stem)


if __name__ == "__main__":
    main()
