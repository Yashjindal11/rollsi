"""Generates the ROLLSI icons (green tile, faint bars, white moving-average line) without dependencies."""
import math
import struct
import zlib
from pathlib import Path

GREEN = (16, 124, 65)
BAR = (70, 170, 110)
WHITE = (255, 255, 255)
BARS = [0.45, 0.62, 0.40, 0.70, 0.58, 0.82, 0.68]
LINE = [(0.12, 0.66), (0.30, 0.56), (0.50, 0.50), (0.70, 0.38), (0.88, 0.26)]
SS = 4  # supersampling factor for anti-aliasing


def seg_dist(px, py, ax, ay, bx, by):
    dx, dy = bx - ax, by - ay
    t = max(0.0, min(1.0, ((px - ax) * dx + (py - ay) * dy) / (dx * dx + dy * dy)))
    return math.hypot(px - (ax + t * dx), py - (ay + t * dy))


def color_at(u, v, size):
    r = 0.18
    cx, cy = min(max(u, r), 1 - r), min(max(v, r), 1 - r)
    if math.hypot(u - cx, v - cy) > r:
        return None
    width = max(0.075, 2.2 / size)
    if any(seg_dist(u, v, *LINE[i], *LINE[i + 1]) < width / 2 for i in range(len(LINE) - 1)):
        return WHITE
    if size >= 32:
        n = len(BARS)
        x0, x1 = 0.14, 0.86
        slot = (x1 - x0) / n
        i = int((u - x0) / slot) if x0 <= u < x1 else -1
        if 0 <= i < n and (u - x0 - i * slot) < slot * 0.7 and v > 1 - 0.14 - BARS[i] * 0.62 and v < 0.86:
            return BAR
    return GREEN


def render(size):
    rows = []
    for y in range(size):
        row = bytearray([0])
        for x in range(size):
            acc = [0, 0, 0, 0]
            for sy in range(SS):
                for sx in range(SS):
                    c = color_at((x + (sx + 0.5) / SS) / size, (y + (sy + 0.5) / SS) / size, size)
                    if c:
                        acc[0] += c[0]; acc[1] += c[1]; acc[2] += c[2]; acc[3] += 255
            n = SS * SS
            a = acc[3] // n
            cov = acc[3] / 255 or 1
            row += bytes([int(acc[0] / cov), int(acc[1] / cov), int(acc[2] / cov), a])
        rows.append(bytes(row))
    raw = b"".join(rows)

    def chunk(tag, data):
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

    ihdr = struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0)
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr) + chunk(b"IDAT", zlib.compress(raw, 9)) + chunk(b"IEND", b"")


out = Path(__file__).resolve().parent.parent / "assets"
out.mkdir(exist_ok=True)
for s in (16, 32, 64, 80, 128):
    (out / f"icon-{s}.png").write_bytes(render(s))
    print("wrote", out / f"icon-{s}.png")
