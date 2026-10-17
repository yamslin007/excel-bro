"""出一张「功能区图标 旧 vs 新」对照图，供重启宿主后一眼确认。

为什么要单独出这张：功能区图标在 16/32px 下非常小，肉眼很难判断"到底变了没有"。
把旧图（备份在 outputs/logo-swap/before/）和新图按**真实尺寸**并排、再放大到可辨认，
就能在宿主里看一眼就知道是哪一版。

用法：python scripts/render_ribbon_icon_versions.py
"""
from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

REPO = Path(__file__).resolve().parent.parent
BEFORE = REPO / "outputs" / "logo-swap" / "before"
AFTER = REPO / "apps" / "excel-addin" / "public" / "assets"
OUT = REPO / "outputs" / "logo-swap" / "ribbon-icon-versions.png"

SIZES = (16, 32, 80)
ZOOM = 128
BG = (245, 245, 245)
INK = (55, 53, 47)
MUTED = (150, 148, 142)
SEP = (222, 222, 222)

FONT_CANDIDATES = (
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\msyhbd.ttc",
    r"C:\Windows\Fonts\simhei.ttf",
)


def load_font(size: int) -> ImageFont.FreeTypeFont:
    for path in FONT_CANDIDATES:
        if Path(path).exists():
            try:
                return ImageFont.truetype(path, size)
            except OSError:
                continue
    return ImageFont.load_default(size)


def paste_center(canvas: Image.Image, image: Image.Image, x: int, y: int, box: int) -> None:
    """把图标**按原尺寸**居中贴到 (x, y) 为左上角、边长 box 的格子里。"""
    canvas.alpha_composite(
        image,
        (x + (box - image.width) // 2, y + (box - image.height) // 2),
    )


def paste_zoomed(canvas: Image.Image, image: Image.Image, x: int, y: int, box: int) -> None:
    """把图标等比放大到最长边不超过 box，再居中贴进同一格。"""
    scale = min(box / image.width, box / image.height)
    size = (max(1, round(image.width * scale)), max(1, round(image.height * scale)))
    scaled = image.resize(size, Image.LANCZOS)
    canvas.alpha_composite(
        scaled,
        (x + (box - scaled.width) // 2, y + (box - scaled.height) // 2),
    )


def main() -> int:
    font_title = load_font(22)
    font_body = load_font(17)
    font_small = load_font(14)

    label_band = 26
    cell_w = ZOOM + 34
    row_height = label_band + ZOOM + 22
    header = 74
    left = 26
    col_gap = 30
    col_w = cell_w * 2
    width = left * 2 + col_w * 2 + col_gap
    height = header + row_height * len(SIZES) + 76

    canvas = Image.new("RGBA", (width, height), BG + (255,))
    draw = ImageDraw.Draw(canvas)

    draw.text((left, 20), "功能区图标：旧（EB 字标）  vs  新（绿意表格）", font=font_title, fill=INK)

    columns = ((left, "旧：EB 字标"), (left + col_w + col_gap, "新：绿意表格"))
    for x, label in columns:
        draw.text((x, header - 30), label, font=font_body, fill=INK)

    for index, size in enumerate(SIZES):
        top = header + index * row_height
        if index:
            draw.line((left, top - 10, width - left, top - 10), fill=SEP)

        before = Image.open(BEFORE / f"excel-bro-icon-{size}.png").convert("RGBA")
        after = Image.open(AFTER / f"excel-bro-icon-{size}-v2.png").convert("RGBA")

        chip_top = top + label_band
        for x, _ in columns:
            # 表头写清楚哪一格是原尺寸、哪一格是放大，免得看错。
            draw.text((x, top + 4), f"{size}px 原尺寸", font=font_small, fill=MUTED)
            draw.text((x + cell_w, top + 4), f"放大到 {ZOOM}px", font=font_small, fill=MUTED)

        for x, source in ((columns[0][0], before), (columns[1][0], after)):
            paste_center(canvas, source, x, chip_top, cell_w)
            paste_zoomed(canvas, source, x + cell_w, chip_top, ZOOM)

    for line_index, line in enumerate(
        (
            "宿主（Excel / WPS）按 URL 缓存功能区图标：只覆盖图片文件、不改 URL，功能区会一直显示旧图。",
            "新版图标 URL 带 -v2 版本号（见 manifest.xml 的 Icon.* 资源与 wps-addin/src/ribbon.ts 的 GetImage）。",
        )
    ):
        draw.text((left, height - 56 + line_index * 22), line, font=font_small, fill=MUTED)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    canvas.convert("RGB").save(OUT)
    print(f"write {OUT.relative_to(REPO)}  {canvas.width}x{canvas.height}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
