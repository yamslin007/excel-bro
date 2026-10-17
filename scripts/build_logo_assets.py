"""把用户给的 1254x1254 白底 JPEG 处理成插件可用的 logo 资源。

要点（踩过的坑都写在注释里）：

1. **不能用全局颜色抠白底**。角色脸部、表格单元格内部本身就是白色/极浅绿，
   按颜色阈值抠会把这些区域一并打穿。必须从**图像边框**做连通域洪泛，只清掉
   与边缘相连的那片底色；被绿色表框围住的白色区域自然被保住。

2. 原图 (0,0)/(1253,1253) 等边缘点实测是 (253,253,253)~(255,255,255)，
   底色均匀，单一种子 + 适中容差即可覆盖全部背景（四角同属一片连通域）。

3. 高分辨率下用硬掩膜、再交给 LANCZOS 降采样 —— 抗锯齿由重采样产生，
   比在小图上羽化边缘干净得多，也不会在深色底上留下白边。

4. 出**两套**裁切：
   - icon：内容占画布约 92%，给功能区 16/32/80 用，细节尽量大。
   - mark：内容占画布约 76%，给界面里的**圆形**徽标用 —— 圆会切掉四角，
     不留白就会把表框和柱状图的边角削掉。

用法：python scripts/build_logo_assets.py
"""
from __future__ import annotations

import sys
from pathlib import Path

from PIL import Image, ImageDraw

REPO = Path(__file__).resolve().parent.parent
SRC = Path(r"C:\Users\jia\.workbuddy\clipboard-images\clipboard-2026-10-12T03-22-21-783Z-3b0a4315.jpg")
OUT_PREVIEW = REPO / "outputs" / "logo-swap"

EXCEL_PUBLIC = REPO / "apps" / "excel-addin" / "public" / "assets"
WPS_PUBLIC = REPO / "apps" / "wps-addin" / "public" / "assets"
EXCEL_SRC_ASSETS = REPO / "apps" / "excel-addin" / "src" / "assets"

SENTINEL = (255, 0, 255)
FLOOD_THRESH = 60

# 功能区图标的 **URL 版本号**。换图时**必须**把它 +1，否则宿主不会重新拉图。
#
# 为什么需要它（2026-10-12 实机踩到）：两个宿主都**按 URL 缓存功能区图标**，
# 而且缓存得非常久 —— `apps/excel-addin/request-log.txt` 里 Excel 宿主
# （UA `ms-office`）最后一次拉 `/assets/excel-bro-icon-{16,32,80}.png` 是
# **2026-09-05**，之后一个多月再没拉过；期间图片文件被覆盖成新图，功能区照样
# 显示旧「EB」。只覆盖文件、不改 URL，宿主永远拿缓存。
#
# 取文件名后缀而不是 `?v=2` 查询串：文件名变更对任何缓存层（宿主位图缓存、
# WebView HTTP 缓存、QtWebEngine 缓存）都是**确定的另一个资源**，不依赖
# 各宿主是否在缓存键里带上查询串。
ICON_URL_VERSION = "v2"

# 原名字继续产出，作为**别名**。宿主可能仍持有老 URL（旧的 ribbon / 旧 manifest
# 缓存），保留别名可保证那些路径也拿到新图，而不是 404 出一个破图标。
# 两份由同一次运行从同一张源图产出，不会漂移。
KEEP_LEGACY_ALIASES = True


def cut_background(image: Image.Image) -> Image.Image:
    """从四条边向内洪泛，把与边缘相连的近白底清成透明。"""
    rgb = image.convert("RGB")
    width, height = rgb.size

    # 用洋红哨兵色标记背景：真图里没有这种颜色，之后按它取反建 alpha。
    work = rgb.copy()
    for seed in ((0, 0), (width - 1, 0), (0, height - 1), (width - 1, height - 1)):
        ImageDraw.floodfill(work, seed, SENTINEL, thresh=FLOOD_THRESH)

    alpha = Image.new("L", (width, height), 255)
    alpha.putdata([0 if pixel == SENTINEL else 255 for pixel in work.getdata()])

    out = rgb.convert("RGBA")
    out.putalpha(alpha)
    return out


def content_box(image: Image.Image) -> tuple[int, int, int, int]:
    """取 alpha 的内容包围盒。"""
    box = image.getchannel("A").getbbox()
    if box is None:
        raise SystemExit("抠完后图像全透明，背景阈值可能过大")
    return box


def square_with_content(image: Image.Image, ratio: float, size: int) -> Image.Image:
    """裁到内容，再按 ratio 缩放并居中放进 size×size 的方画布。"""
    cropped = image.crop(content_box(image))
    inner = max(1, int(round(size * ratio)))
    scale = min(inner / cropped.width, inner / cropped.height)
    target = (
        max(1, int(round(cropped.width * scale))),
        max(1, int(round(cropped.height * scale))),
    )
    resized = cropped.resize(target, Image.LANCZOS)
    canvas = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    canvas.paste(
        resized,
        ((size - resized.width) // 2, (size - resized.height) // 2),
        resized,
    )
    return canvas


def write_embedded_svg(png_path: Path, svg_path: Path, size: int) -> None:
    """把 PNG 以 base64 内嵌进一个 SVG，写到 svg_path。

    为什么保留 .svg 这条路：WPS 功能区原先用 `assets/excel-bro-icon.svg`，
    且已在宿主上验证可用。把 svg 内容换成"内嵌新图"而不是删掉它，可以保证
    **任何仍在引用 .svg 的路径也拿到新 logo**，不会留下一个还在显示退役 "EB"
    字标的死角文件。
    """
    import base64

    payload = base64.b64encode(png_path.read_bytes()).decode("ascii")
    svg_path.write_text(
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{size}" height="{size}" '
        f'viewBox="0 0 {size} {size}">\n'
        f"  <!-- Excel Bro logo（栅格内嵌）。真正的源图见 "
        f"scripts/build_logo_assets.py 与 apps/excel-addin/src/assets/。 -->\n"
        f'  <image width="{size}" height="{size}" '
        f'href="data:image/png;base64,{payload}" />\n'
        f"</svg>\n",
        encoding="utf-8",
    )


def main() -> int:
    if not SRC.exists():
        print(f"找不到源图：{SRC}", file=sys.stderr)
        return 1

    source = Image.open(SRC)
    print(f"源图：{source.size[0]}x{source.size[1]} {source.mode}")

    keyed = cut_background(source)
    box = content_box(keyed)
    print(
        f"内容包围盒：{box}  →  {box[2] - box[0]}x{box[3] - box[1]}"
        f"（四周白边被裁掉）"
    )

    for directory in (EXCEL_PUBLIC, WPS_PUBLIC, EXCEL_SRC_ASSETS):
        directory.mkdir(parents=True, exist_ok=True)

    # 功能区图标：Excel manifest 用 16/32/80 三档。带版本号的名字是 manifest
    # 真正引用的那一份；无版本号的名字只作别名（见 ICON_URL_VERSION 注释）。
    for size in (16, 32, 80):
        icon = square_with_content(keyed, 0.92, size)
        names = [f"excel-bro-icon-{size}-{ICON_URL_VERSION}.png"]
        if KEEP_LEGACY_ALIASES:
            names.append(f"excel-bro-icon-{size}.png")
        for name in names:
            icon.save(EXCEL_PUBLIC / name, optimize=True)
            print(f"  write excel-addin/public/assets/{name}")

    # WPS 功能区按 80px 出，命名与 Excel 一致便于对照。
    icon80 = square_with_content(keyed, 0.92, 80)
    wps_names = [f"excel-bro-icon-{ICON_URL_VERSION}.png"]
    if KEEP_LEGACY_ALIASES:
        wps_names.append("excel-bro-icon.png")
    for name in wps_names:
        icon80.save(WPS_PUBLIC / name, optimize=True)
        print(f"  write wps-addin/public/assets/{name}")

    # 两端各保留一份 .svg（内嵌同一张 PNG），避免退役的 "EB" 字标留在死路径上。
    write_embedded_svg(
        WPS_PUBLIC / "excel-bro-icon.png", WPS_PUBLIC / "excel-bro-icon.svg", 80
    )
    print("  write wps-addin/public/assets/excel-bro-icon.svg（内嵌新图）")
    write_embedded_svg(
        EXCEL_PUBLIC / "excel-bro-icon-80.png", EXCEL_PUBLIC / "excel-bro-icon.svg", 80
    )
    print("  write excel-addin/public/assets/excel-bro-icon.svg（内嵌新图）")

    # 界面圆形徽标：留白版本，供两端 import。放在 excel-addin/src 下作为**唯一来源**，
    # WPS 用相对路径 import 同一份文件（与该仓库已有的共享 styles.css 做法一致）。
    for size in (32, 64, 128, 256):
        mark = square_with_content(keyed, 0.76, size)
        mark.save(EXCEL_SRC_ASSETS / f"excel-bro-logo-{size}.png", optimize=True)
        print(f"  write excel-addin/src/assets/excel-bro-logo-{size}.png")

    # 另存一张 512 大图供预览与后续取用。
    square_with_content(keyed, 0.76, 512).save(OUT_PREVIEW / "logo-master-512.png")
    print("  write outputs/logo-swap/logo-master-512.png")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
