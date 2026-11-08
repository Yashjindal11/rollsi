"""Generates ROLLSI ribbon icons (16/32/80 px) as dependency-free PNGs: assets/rb-<name>-<size>.png."""
import math
import struct
import zlib
from pathlib import Path

G, L, D, R, Y, B, GR = "#107C41", "#7FBF9B", "#404040", "#D64541", "#F2A900", "#2F5597", "#A6A6A6"


def rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def seg_dist(px, py, ax, ay, bx, by):
    dx, dy = bx - ax, by - ay
    if dx == dy == 0:
        return math.hypot(px - ax, py - ay)
    t = max(0.0, min(1.0, ((px - ax) * dx + (py - ay) * dy) / (dx * dx + dy * dy)))
    return math.hypot(px - (ax + t * dx), py - (ay + t * dy))


def in_poly(x, y, pts):
    inside = False
    j = len(pts) - 1
    for i in range(len(pts)):
        xi, yi = pts[i]
        xj, yj = pts[j]
        if (yi > y) != (yj > y) and x < (xj - xi) * (y - yi) / (yj - yi + 1e-12) + xi:
            inside = not inside
        j = i
    return inside


# Primitives in unit coordinates (0..1, y down). Later primitives paint over earlier ones.
def rect(x0, y0, x1, y1, c):
    return ("rect", x0, y0, x1, y1, c)


def line(pts, w, c, dash=False):
    return ("line", pts, w, c, dash)


def dot(cx, cy, r, c):
    return ("dot", cx, cy, r, c)


def poly(pts, c):
    return ("poly", pts, c)


def arrow_head(x, y, dx, dy, s, c):
    n = math.hypot(dx, dy)
    ux, uy = dx / n, dy / n
    px, py = -uy, ux
    return poly([(x, y), (x - ux * s + px * s * 0.6, y - uy * s + py * s * 0.6), (x - ux * s - px * s * 0.6, y - uy * s - py * s * 0.6)], c)


def arc(cx, cy, r, a0, a1, n=24):
    return [(cx + r * math.cos(a0 + (a1 - a0) * k / n), cy + r * math.sin(a0 + (a1 - a0) * k / n)) for k in range(n + 1)]


ICONS = {
    "trend": [
        line([(0.08, 0.7), (0.22, 0.45), (0.36, 0.62), (0.5, 0.35), (0.64, 0.55), (0.78, 0.28), (0.92, 0.42)], 0.06, GR),
        line([(0.08, 0.66), (0.3, 0.55), (0.55, 0.46), (0.92, 0.33)], 0.11, G),
    ],
    "control": [
        line([(0.06, 0.25), (0.94, 0.25)], 0.05, GR, True),
        line([(0.06, 0.75), (0.94, 0.75)], 0.05, GR, True),
        line([(0.06, 0.5), (0.94, 0.5)], 0.04, GR),
        line([(0.08, 0.55), (0.25, 0.42), (0.4, 0.6), (0.55, 0.12), (0.7, 0.52), (0.92, 0.45)], 0.08, G),
        dot(0.55, 0.12, 0.09, R),
    ],
    "waterfall": [
        rect(0.06, 0.35, 0.24, 0.92, G),
        rect(0.3, 0.18, 0.48, 0.35, G),
        rect(0.54, 0.18, 0.72, 0.45, R),
        rect(0.78, 0.45, 0.96, 0.92, GR),
    ],
    "target": [
        rect(0.1, 0.45, 0.3, 0.92, G),
        rect(0.4, 0.3, 0.6, 0.92, G),
        rect(0.7, 0.55, 0.9, 0.92, G),
        line([(0.05, 0.35), (0.35, 0.35)], 0.09, D),
        line([(0.35, 0.22), (0.65, 0.22)], 0.09, D),
        line([(0.65, 0.4), (0.95, 0.4)], 0.09, D),
    ],
    "pareto": [
        rect(0.06, 0.3, 0.26, 0.92, G),
        rect(0.3, 0.5, 0.5, 0.92, G),
        rect(0.54, 0.66, 0.74, 0.92, L),
        rect(0.78, 0.78, 0.96, 0.92, L),
        line([(0.16, 0.42), (0.4, 0.2), (0.64, 0.12), (0.88, 0.08)], 0.07, D),
    ],
    "gantt": [
        rect(0.06, 0.12, 0.5, 0.3, G),
        rect(0.3, 0.41, 0.78, 0.59, G),
        rect(0.6, 0.7, 0.96, 0.88, L),
    ],
    "mekko": [
        rect(0.04, 0.08, 0.44, 0.5, G),
        rect(0.04, 0.5, 0.44, 0.92, L),
        rect(0.46, 0.08, 0.72, 0.62, B),
        rect(0.46, 0.62, 0.72, 0.92, Y),
        rect(0.74, 0.08, 0.96, 0.36, G),
        rect(0.74, 0.36, 0.96, 0.92, L),
    ],
    "diff": [
        rect(0.06, 0.55, 0.32, 0.92, GR),
        rect(0.4, 0.25, 0.66, 0.92, G),
        line([(0.32, 0.55), (0.84, 0.55)], 0.04, GR, True),
        line([(0.66, 0.25), (0.84, 0.25)], 0.04, GR, True),
        line([(0.84, 0.55), (0.84, 0.33)], 0.07, D),
        arrow_head(0.84, 0.22, 0, -1, 0.14, D),
    ],
    "cagr": [
        rect(0.08, 0.66, 0.26, 0.92, L),
        rect(0.41, 0.52, 0.59, 0.92, L),
        rect(0.74, 0.36, 0.92, 0.92, L),
        line([(0.14, 0.52), (0.72, 0.2)], 0.07, D),
        arrow_head(0.84, 0.13, 1, -0.55, 0.15, D),
        dot(0.45, 0.32, 0.12, Y),
    ],
    "totals": [
        rect(0.08, 0.5, 0.36, 0.92, G),
        rect(0.08, 0.32, 0.36, 0.5, L),
        rect(0.6, 0.62, 0.88, 0.92, G),
        rect(0.6, 0.46, 0.88, 0.62, L),
        line([(0.12, 0.17), (0.32, 0.17)], 0.08, D),
        line([(0.64, 0.31), (0.84, 0.31)], 0.08, D),
    ],
    "refresh": [
        line(arc(0.5, 0.52, 0.33, math.radians(-60), math.radians(240)), 0.1, G),
        arrow_head(0.68, 0.17, 0.7, 0.3, 0.2, G),
    ],
    "remove": [
        rect(0.12, 0.45, 0.4, 0.92, GR),
        rect(0.5, 0.3, 0.78, 0.92, GR),
        line([(0.2, 0.15), (0.9, 0.85)], 0.11, R),
        line([(0.9, 0.15), (0.2, 0.85)], 0.11, R),
    ],
    "style": [
        rect(0.08, 0.5, 0.3, 0.92, G),
        rect(0.39, 0.3, 0.61, 0.92, Y),
        rect(0.7, 0.62, 0.92, 0.92, B),
        line([(0.06, 0.94), (0.94, 0.94)], 0.05, D),
        poly([(0.78, 0.08), (0.82, 0.2), (0.94, 0.24), (0.82, 0.28), (0.78, 0.4), (0.74, 0.28), (0.62, 0.24), (0.74, 0.2)], Y),
    ],
    "palette": [
        dot(0.5, 0.5, 0.44, "#EDEDED"),
        dot(0.32, 0.32, 0.11, G),
        dot(0.62, 0.25, 0.11, Y),
        dot(0.75, 0.52, 0.11, B),
        dot(0.32, 0.65, 0.11, R),
        dot(0.58, 0.74, 0.09, "#FFFFFF"),
    ],
    "export": [
        rect(0.06, 0.24, 0.78, 0.94, D),
        rect(0.11, 0.29, 0.73, 0.89, "#FFFFFF"),
        rect(0.2, 0.62, 0.32, 0.82, G),
        rect(0.38, 0.5, 0.5, 0.82, G),
        rect(0.56, 0.4, 0.66, 0.82, G),
        line([(0.6, 0.4), (0.86, 0.14)], 0.08, B),
        arrow_head(0.94, 0.06, 1, -1, 0.2, B),
    ],
    "demo": [
        dot(0.5, 0.5, 0.44, G),
        poly([(0.38, 0.28), (0.38, 0.72), (0.76, 0.5)], "#FFFFFF"),
    ],
    "panel": [
        rect(0.06, 0.1, 0.94, 0.9, D),
        rect(0.11, 0.15, 0.6, 0.85, "#FFFFFF"),
        rect(0.65, 0.15, 0.89, 0.85, G),
        line([(0.2, 0.32), (0.5, 0.32)], 0.06, GR),
        line([(0.2, 0.5), (0.5, 0.5)], 0.06, GR),
        line([(0.2, 0.68), (0.42, 0.68)], 0.06, GR),
    ],
}


def color_at(prims, u, v, size):
    out = None
    min_w = 1.4 / size
    for p in prims:
        kind = p[0]
        if kind == "rect" and p[1] <= u <= p[3] and p[2] <= v <= p[4]:
            out = p[5]
        elif kind == "dot" and math.hypot(u - p[1], v - p[2]) <= p[3]:
            out = p[4]
        elif kind == "poly" and in_poly(u, v, p[1]):
            out = p[2]
        elif kind == "line":
            pts, w, c, dash = p[1], max(p[2], min_w), p[3], p[4]
            for i in range(len(pts) - 1):
                if seg_dist(u, v, *pts[i], *pts[i + 1]) <= w / 2:
                    if dash and size >= 32:
                        ax, ay = pts[i]
                        if int(math.hypot(u - ax, v - ay) / 0.09) % 2:
                            continue
                    out = c
                    break
    return out


def png(size, prims):
    ss = 4
    raw = bytearray()
    for y in range(size):
        raw.append(0)
        for x in range(size):
            acc = [0, 0, 0, 0]
            for sy in range(ss):
                for sx in range(ss):
                    c = color_at(prims, (x + (sx + 0.5) / ss) / size, (y + (sy + 0.5) / ss) / size, size)
                    if c:
                        r, g, b = rgb(c)
                        acc[0] += r
                        acc[1] += g
                        acc[2] += b
                        acc[3] += 1
            n = acc[3]
            raw += bytes([acc[0] // n, acc[1] // n, acc[2] // n, 255 * n // (ss * ss)]) if n else bytes(4)

    def chunk(tag, data):
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(bytes(raw), 9))
        + chunk(b"IEND", b"")
    )


out = Path(__file__).resolve().parent.parent / "assets"
for name, prims in ICONS.items():
    for size in (16, 32, 80):
        (out / f"rb-{name}-{size}.png").write_bytes(png(size, prims))
print(f"wrote {len(ICONS) * 3} ribbon icons to {out}")
