from pathlib import Path
import math
import struct
import zlib


ROOT_DIR = Path(__file__).resolve().parents[1]
OUTPUT_PATH = ROOT_DIR / "assets" / "app_mascot.ico"


def rgba(hex_color: str, alpha: int = 255):
    hex_color = hex_color.lstrip("#")
    return tuple(int(hex_color[index:index + 2], 16) for index in (0, 2, 4)) + (alpha,)


def make_canvas(size: int):
    return [[(0, 0, 0, 0) for _ in range(size)] for _ in range(size)]


def set_px(img, x: int, y: int, color):
    size = len(img)
    if 0 <= x < size and 0 <= y < size:
        img[y][x] = color


def fill_rect(img, x0: int, y0: int, x1: int, y1: int, color):
    for y in range(max(0, y0), min(len(img), y1)):
        row = img[y]
        for x in range(max(0, x0), min(len(img), x1)):
            row[x] = color


def fill_circle(img, cx: int, cy: int, radius: int, color):
    size = len(img)
    radius_sq = radius * radius
    for y in range(max(0, int(cy - radius)), min(size, int(cy + radius) + 1)):
        dy = y - cy
        row = img[y]
        for x in range(max(0, int(cx - radius)), min(size, int(cx + radius) + 1)):
            dx = x - cx
            if dx * dx + dy * dy <= radius_sq:
                row[x] = color


def fill_rounded_rect(img, x0: int, y0: int, x1: int, y1: int, radius: int, color):
    fill_rect(img, x0 + radius, y0, x1 - radius, y1, color)
    fill_rect(img, x0, y0 + radius, x1, y1 - radius, color)
    fill_circle(img, x0 + radius, y0 + radius, radius, color)
    fill_circle(img, x1 - radius - 1, y0 + radius, radius, color)
    fill_circle(img, x0 + radius, y1 - radius - 1, radius, color)
    fill_circle(img, x1 - radius - 1, y1 - radius - 1, radius, color)


def fill_triangle(img, p1, p2, p3, color):
    xs = [p1[0], p2[0], p3[0]]
    ys = [p1[1], p2[1], p3[1]]
    min_x = max(0, int(min(xs)))
    max_x = min(len(img) - 1, int(max(xs)))
    min_y = max(0, int(min(ys)))
    max_y = min(len(img) - 1, int(max(ys)))

    def area(a, b, c):
        return (b[0] - a[0]) * (c[1] - a[1]) - (b[1] - a[1]) * (c[0] - a[0])

    triangle_area = area(p1, p2, p3)
    if triangle_area == 0:
        return

    for y in range(min_y, max_y + 1):
        for x in range(min_x, max_x + 1):
            p = (x + 0.5, y + 0.5)
            w1 = area(p2, p3, p)
            w2 = area(p3, p1, p)
            w3 = area(p1, p2, p)
            if triangle_area > 0:
                inside = w1 >= 0 and w2 >= 0 and w3 >= 0
            else:
                inside = w1 <= 0 and w2 <= 0 and w3 <= 0
            if inside:
                img[y][x] = color


def draw_arc(img, cx: int, cy: int, radius: int, thickness: int, start_deg: int, end_deg: int, color):
    for degree in range(start_deg, end_deg + 1):
        radians = math.radians(degree)
        for offset in range(thickness):
            x = int(round(cx + math.cos(radians) * (radius - offset)))
            y = int(round(cy + math.sin(radians) * (radius - offset)))
            set_px(img, x, y, color)
            set_px(img, x, y + 1, color)


def draw_icon(size: int):
    img = make_canvas(size)
    teal = rgba("#0F766E")
    teal_dark = rgba("#115E59")
    cream = rgba("#FFF7E8")
    gold = rgba("#F0B429")
    rose = rgba("#F29C9C")
    navy = rgba("#102542")
    white = rgba("#FFFFFF")

    head_x0 = int(size * 0.18)
    head_y0 = int(size * 0.22)
    head_x1 = int(size * 0.82)
    head_y1 = int(size * 0.82)
    radius = max(2, int(size * 0.12))

    fill_triangle(img, (int(size * 0.28), int(size * 0.18)), (int(size * 0.38), int(size * 0.06)), (int(size * 0.42), int(size * 0.24)), teal)
    fill_triangle(img, (int(size * 0.72), int(size * 0.18)), (int(size * 0.62), int(size * 0.06)), (int(size * 0.58), int(size * 0.24)), teal)
    fill_circle(img, int(size * 0.50), int(size * 0.09), max(1, int(size * 0.035)), gold)
    fill_rect(img, int(size * 0.485), int(size * 0.10), int(size * 0.515), int(size * 0.22), gold)

    fill_rounded_rect(img, head_x0, head_y0, head_x1, head_y1, radius, teal)
    fill_rounded_rect(img, int(size * 0.26), int(size * 0.34), int(size * 0.74), int(size * 0.72), max(2, int(size * 0.08)), cream)
    fill_rect(img, int(size * 0.16), int(size * 0.27), int(size * 0.21), int(size * 0.56), teal_dark)
    fill_rect(img, int(size * 0.79), int(size * 0.27), int(size * 0.84), int(size * 0.56), teal_dark)

    eye_radius = max(1, int(size * 0.045))
    fill_circle(img, int(size * 0.40), int(size * 0.47), eye_radius, navy)
    fill_circle(img, int(size * 0.60), int(size * 0.47), eye_radius, navy)
    sparkle = max(1, int(size * 0.015))
    fill_circle(img, int(size * 0.39), int(size * 0.455), sparkle, white)
    fill_circle(img, int(size * 0.59), int(size * 0.455), sparkle, white)

    cheek_radius = max(1, int(size * 0.04))
    fill_circle(img, int(size * 0.31), int(size * 0.58), cheek_radius, rose)
    fill_circle(img, int(size * 0.69), int(size * 0.58), cheek_radius, rose)
    draw_arc(img, int(size * 0.50), int(size * 0.57), int(size * 0.12), max(1, int(size * 0.02)), 25, 155, navy)

    foot_radius = max(1, int(size * 0.035))
    fill_circle(img, int(size * 0.42), int(size * 0.86), foot_radius, gold)
    fill_circle(img, int(size * 0.58), int(size * 0.86), foot_radius, gold)
    return img


def build_png(img):
    height = len(img)
    width = len(img[0])
    raw = bytearray()
    for row in img:
        raw.append(0)
        for red, green, blue, alpha in row:
            raw.extend((red, green, blue, alpha))

    def chunk(tag: bytes, data: bytes):
        return (
            struct.pack(">I", len(data))
            + tag
            + data
            + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
        )

    ihdr = struct.pack(">IIBBBBB", width, height, 8, 6, 0, 0, 0)
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr) + chunk(b"IDAT", zlib.compress(bytes(raw), 9)) + chunk(b"IEND", b"")


def build_icon():
    sizes = [16, 32, 48, 64, 128, 256]
    png_images = [build_png(draw_icon(size)) for size in sizes]

    header = struct.pack("<HHH", 0, 1, len(png_images))
    entries = bytearray()
    payload = bytearray()
    offset = 6 + 16 * len(png_images)

    for size, png_data in zip(sizes, png_images):
        width_byte = 0 if size >= 256 else size
        height_byte = 0 if size >= 256 else size
        entries.extend(struct.pack("<BBBBHHII", width_byte, height_byte, 0, 0, 1, 32, len(png_data), offset))
        payload.extend(png_data)
        offset += len(png_data)

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_PATH.write_bytes(header + entries + payload)
    print(f"Icon generated at {OUTPUT_PATH}")


if __name__ == "__main__":
    build_icon()