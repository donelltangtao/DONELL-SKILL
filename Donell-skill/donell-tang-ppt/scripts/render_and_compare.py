#!/usr/bin/env python3
"""Render a PPTX and create original-vs-rendered comparison images.

Usage:
  python scripts/render_and_compare.py deck.pptx --original-dir originals --out-dir qa

Original images should be named in natural order. The script renders the PPTX,
then makes one side-by-side comparison PNG per slide and an overview sheet.
"""

from __future__ import annotations

import argparse
import re
import subprocess
from pathlib import Path
from PIL import Image, ImageDraw

RENDERER = Path('/home/oai/skills/slides/container_tools/render_slides.py')


def natural_key(path: Path):
    return [int(x) if x.isdigit() else x.lower() for x in re.split(r'(\d+)', path.name)]


def fit(im: Image.Image, max_w: int, max_h: int) -> Image.Image:
    out = im.copy().convert('RGB')
    out.thumbnail((max_w, max_h))
    return out


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument('pptx', type=Path)
    ap.add_argument('--original-dir', type=Path, required=True)
    ap.add_argument('--out-dir', type=Path, required=True)
    args = ap.parse_args()

    args.out_dir.mkdir(parents=True, exist_ok=True)
    rendered_dir = args.out_dir / 'rendered'
    rendered_dir.mkdir(exist_ok=True)

    subprocess.run([
        'python', str(RENDERER), str(args.pptx), '--output_dir', str(rendered_dir)
    ], check=True)

    originals = sorted(
        [p for p in args.original_dir.iterdir() if p.suffix.lower() in {'.png', '.jpg', '.jpeg', '.webp'}],
        key=natural_key,
    )
    rendered = sorted(rendered_dir.glob('slide-*.png'), key=natural_key)

    if len(originals) != len(rendered):
        raise SystemExit(f'Page count mismatch: originals={len(originals)}, rendered={len(rendered)}')

    compare_paths = []
    for idx, (orig_path, render_path) in enumerate(zip(originals, rendered), start=1):
        orig = fit(Image.open(orig_path), 900, 506)
        rend = fit(Image.open(render_path), 900, 506)
        pad = 16
        height = max(orig.height, rend.height) + pad * 2
        canvas = Image.new('RGB', (orig.width + rend.width + pad * 3, height), (245, 245, 245))
        canvas.paste(orig, (pad, pad))
        canvas.paste(rend, (orig.width + pad * 2, pad))
        draw = ImageDraw.Draw(canvas)
        draw.text((pad + 6, pad + 6), f'{idx} 原图', fill='white')
        draw.text((orig.width + pad * 2 + 6, pad + 6), f'{idx} PPTX预览', fill='white')
        out = args.out_dir / f'compare-{idx:02d}.png'
        canvas.save(out)
        compare_paths.append(out)

    thumb_w, thumb_h, pad = 420, 236, 12
    rows = len(compare_paths)
    overview = Image.new('RGB', (thumb_w * 2 + pad * 3, rows * thumb_h + pad * (rows + 1)), (248, 248, 248))
    for row, (orig_path, render_path) in enumerate(zip(originals, rendered)):
        orig = fit(Image.open(orig_path), thumb_w, thumb_h)
        rend = fit(Image.open(render_path), thumb_w, thumb_h)
        y = pad + row * (thumb_h + pad)
        overview.paste(orig, (pad, y))
        overview.paste(rend, (thumb_w + pad * 2, y))
    overview.save(args.out_dir / 'overview.png')
    print(args.out_dir)


if __name__ == '__main__':
    main()
