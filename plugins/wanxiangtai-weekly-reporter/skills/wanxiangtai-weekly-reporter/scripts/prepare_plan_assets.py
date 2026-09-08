from __future__ import annotations

import argparse
import io
import json
import subprocess
import urllib.request
from pathlib import Path

from PIL import Image, ImageOps


HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/140.0.0.0 Safari/537.36",
    "Referer": "https://one.alimama.com/",
}


def download(url: str) -> bytes:
    request = urllib.request.Request(url, headers=HEADERS)
    with urllib.request.urlopen(request, timeout=45) as response:
        return response.read()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("plan_json", type=Path)
    parser.add_argument("output_root", type=Path)
    args = parser.parse_args()

    data = json.loads(args.plan_json.read_text(encoding="utf-8"))
    target = args.output_root / data["campaignId"]
    target.mkdir(parents=True, exist_ok=True)

    image_paths: list[Path] = []
    video_frame_paths: list[Path] = []
    failures: list[dict[str, str]] = []

    for material in data["materials"]:
        try:
            content = download(material["url"])
            if material["type"] == "图片":
                path = target / f'{material["id"]}.png'
                with Image.open(io.BytesIO(content)) as source:
                    ImageOps.exif_transpose(source).convert("RGB").save(path, "PNG")
                image_paths.append(path)
            elif material["type"] == "视频":
                video_path = target / f'{material["id"]}.mp4'
                video_path.write_bytes(content)
                for label, second in (("frame", 0.5), ("t2", 2), ("t4", 4)):
                    frame_path = target / f'{material["id"]}-{label}.png'
                    subprocess.run(
                        ["ffmpeg", "-loglevel", "error", "-y", "-ss", str(second), "-i", str(video_path), "-frames:v", "1", str(frame_path)],
                        check=True,
                    )
                    if label == "frame":
                        video_frame_paths.append(frame_path)
        except Exception as exc:  # keep batch moving and report exact material
            failures.append({"id": material["id"], "error": str(exc)})

    print(
        json.dumps(
            {
                "campaignId": data["campaignId"],
                "directory": str(target),
                "images": len(image_paths),
                "video_frames": len(video_frame_paths),
                "failures": failures,
            },
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    main()
