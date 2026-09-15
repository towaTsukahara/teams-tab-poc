"""Teamsアプリ用のcolor.png/outline.pngを、外部ライブラリなし(標準ライブラリのみ)で生成する。"""
import struct
import zlib
import pathlib

OUT_DIR = pathlib.Path(__file__).resolve().parent.parent / "teams-package"


def write_png(path, width, height, pixels_rgba):
    def chunk(tag, data):
        return (
            struct.pack(">I", len(data))
            + tag
            + data
            + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
        )

    raw = bytearray()
    for y in range(height):
        raw.append(0)  # フィルタなし
        for x in range(width):
            raw.extend(pixels_rgba[y * width + x])

    ihdr = struct.pack(">IIBBBBB", width, height, 8, 6, 0, 0, 0)
    idat = zlib.compress(bytes(raw), 9)

    png = b"\x89PNG\r\n\x1a\n"
    png += chunk(b"IHDR", ihdr)
    png += chunk(b"IDAT", idat)
    png += chunk(b"IEND", b"")

    path.write_bytes(png)


def is_in_t(x, y, size):
    """「T」の形（横棒＋縦棒）に入っているかを判定する。"""
    bar_top, bar_bottom = size * 0.16, size * 0.34
    bar_left, bar_right = size * 0.16, size * 0.84
    stem_left, stem_right = size * 0.42, size * 0.58
    stem_bottom = size * 0.84

    in_bar = (bar_left <= x <= bar_right) and (bar_top <= y <= bar_bottom)
    in_stem = (stem_left <= x <= stem_right) and (bar_top <= y <= stem_bottom)
    return in_bar or in_stem


def make_color_icon(size=192):
    bg = (91, 33, 182, 255)  # 紫系の適当な背景色
    fg = (255, 255, 255, 255)
    pixels = []
    for y in range(size):
        for x in range(size):
            pixels.append(fg if is_in_t(x, y, size) else bg)
    return pixels


def make_outline_icon(size=32):
    transparent = (0, 0, 0, 0)
    white = (255, 255, 255, 255)
    pixels = []
    for y in range(size):
        for x in range(size):
            pixels.append(white if is_in_t(x, y, size) else transparent)
    return pixels


if __name__ == "__main__":
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    write_png(OUT_DIR / "color.png", 192, 192, make_color_icon(192))
    write_png(OUT_DIR / "outline.png", 32, 32, make_outline_icon(32))
    print(f"wrote {OUT_DIR / 'color.png'}")
    print(f"wrote {OUT_DIR / 'outline.png'}")
