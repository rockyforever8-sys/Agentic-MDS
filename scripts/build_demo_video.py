#!/usr/bin/env python3
"""Build a smooth 16:9 executive demo video from sequenced MDS screenshots."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[1]
STEPS = ROOT / "docs" / "demo-steps"
FRAMES = ROOT / "docs" / "demo-frames"
OUT = ROOT / "presentations" / "IMDS_Agentic_Workflow_Demo.mp4"

W, H = 1920, 1080
NAVY = (11, 31, 58)
GOLD = (196, 163, 90)
WHITE = (255, 255, 255)
FPS = 30
# Hold + dissolve: ~2.0s readable, ~0.35s crossfade — exec-demo pace (~45s for 18 steps)
HOLD_S = 1.85
XFADE_S = 0.35
TITLE_S = 3.2
END_S = 2.6


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    candidates = [
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for p in candidates:
        if Path(p).exists():
            return ImageFont.truetype(p, size)
    return ImageFont.load_default()


def _step_key(path: Path) -> tuple:
    stem = path.stem
    num = 0
    for i, ch in enumerate(stem):
        if ch.isdigit():
            j = i
            while j < len(stem) and stem[j].isdigit():
                j += 1
            num = int(stem[i:j])
            break
    return (num, stem.lower())


def step_label(path: Path) -> str:
    stem = path.stem
    # strip leading "12_" etc.
    i = 0
    while i < len(stem) and (stem[i].isdigit() or stem[i] == "_"):
        i += 1
    rest = stem[i:].replace("_", " ").strip()
    return rest[:48] if rest else stem


def list_steps() -> list[Path]:
    files = []
    for ext in ("*.png", "*.jpg", "*.jpeg", "*.PNG", "*.JPG"):
        files.extend(STEPS.glob(ext))
    files = [p for p in files if p.name.lower() != "readme.md"]
    files.sort(key=_step_key)
    return files


def letterbox(src: Image.Image) -> Image.Image:
    img = src.convert("RGB")
    canvas = Image.new("RGB", (W, H), NAVY)
    # leave 72px for lower third
    box_w, box_h = W, H - 72
    scale = min(box_w / img.width, box_h / img.height)
    nw, nh = max(1, int(img.width * scale)), max(1, int(img.height * scale))
    img = img.resize((nw, nh), Image.Resampling.LANCZOS)
    x = (box_w - nw) // 2
    y = (box_h - nh) // 2
    canvas.paste(img, (x, y))
    draw = ImageDraw.Draw(canvas)
    draw.rectangle((0, H - 72, W, H), fill=NAVY)
    draw.rectangle((0, H - 74, W, H - 72), fill=GOLD)
    return canvas


def caption(canvas: Image.Image, text: str, right: str) -> Image.Image:
    out = canvas.copy()
    draw = ImageDraw.Draw(out)
    draw.text((36, H - 50), text, font=font(26, True), fill=GOLD)
    bbox = draw.textbbox((0, 0), right, font=font(22, False))
    draw.text((W - 36 - (bbox[2] - bbox[0]), H - 48), right, font=font(22, False), fill=WHITE)
    return out


def title_card() -> Image.Image:
    im = Image.new("RGB", (W, H), NAVY)
    d = ImageDraw.Draw(im)
    d.rectangle((0, 0, 18, H), fill=GOLD)
    d.text((80, 280), "INTERNAL CONFIDENTIAL  ·  EXECUTIVE DEMONSTRATION", font=font(22, True), fill=GOLD)
    d.text((80, 360), "32 manual IMDS clicks.", font=font(56, True), fill=WHITE)
    d.text((80, 440), "Now one agentic workflow.", font=font(56, True), fill=WHITE)
    d.text((80, 560), "Automated review, accept / reject, and Propose — Johnson Electric IMDS desk.", font=font(28, False), fill=(197, 208, 220))
    d.text((80, 900), "Supplier Quality  ·  Agentic MDS  ·  August 2026", font=font(22, False), fill=GOLD)
    return im


def end_card() -> Image.Image:
    im = Image.new("RGB", (W, H), NAVY)
    d = ImageDraw.Draw(im)
    d.rectangle((0, 0, 18, H), fill=GOLD)
    d.text((80, 380), "Same-day inbox. Zero PPAP delay.", font=font(48, True), fill=WHITE)
    d.text((80, 480), "Green auto-accept  ·  Red auto-reject  ·  Human on novel cases", font=font(26, False), fill=(197, 208, 220))
    d.text((80, 900), "Invest in Agentic AI Today. Approve the 20-MDS pilot.", font=font(24, True), fill=GOLD)
    return im


def save_frame(im: Image.Image, idx: int) -> Path:
    FRAMES.mkdir(parents=True, exist_ok=True)
    p = FRAMES / f"f{idx:05d}.png"
    im.save(p, "PNG")
    return p


def blend(a: Image.Image, b: Image.Image, t: float) -> Image.Image:
    return Image.blend(a.convert("RGB"), b.convert("RGB"), t)


def main() -> int:
    steps = list_steps()
    if not steps:
        print(f"No screenshots in {STEPS}. Drop 01.png … N.png there.", file=sys.stderr)
        return 2
    FRAMES.mkdir(parents=True, exist_ok=True)
    for old in FRAMES.glob("*.png"):
        old.unlink()

    n = len(steps)
    hold = int(HOLD_S * FPS)
    xf = int(XFADE_S * FPS)
    title_n = int(TITLE_S * FPS)
    end_n = int(END_S * FPS)

    prepared = []
    for i, path in enumerate(steps, 1):
        raw = Image.open(path)
        framed = letterbox(raw)
        framed = caption(framed, f"STEP  {i:02d}  /  {n:02d}   ·   {step_label(path)}", "Agentic MDS  ·  IMDS UI")
        prepared.append(framed)

    seq = []
    title = title_card()
    end = end_card()
    seq.extend([title] * title_n)
    # fade title into first screenshot
    for k in range(xf):
        seq.append(blend(title, prepared[0], (k + 1) / xf))
    for i, frame in enumerate(prepared):
        seq.extend([frame] * hold)
        if i + 1 < len(prepared):
            nxt = prepared[i + 1]
            for k in range(xf):
                seq.append(blend(frame, nxt, (k + 1) / xf))
    for k in range(xf):
        seq.append(blend(prepared[-1], end, (k + 1) / xf))
    seq.extend([end] * end_n)

    for i, im in enumerate(seq):
        save_frame(im, i)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    cmd = [
        "ffmpeg", "-y",
        "-framerate", str(FPS),
        "-i", str(FRAMES / "f%05d.png"),
        "-c:v", "libx264",
        "-pix_fmt", "yuv420p",
        "-crf", "18",
        "-movflags", "+faststart",
        str(OUT),
    ]
    subprocess.run(cmd, check=True)
    print(f"Wrote {OUT}  ({n} screenshots, {len(seq)/FPS:.1f}s)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
