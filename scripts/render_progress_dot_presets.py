#!/usr/bin/env python
"""生成「有新进度」浮标的**四预设对照小样**（改配色前先给小样）。

做法是**真渲染**，不是手绘示意：
  1. 先用 `scripts/build_progress_dot_preview.mjs` 生成预览页（样式原文抽自
     `apps/excel-addin/src/styles.css`，令牌取自 `useTheme.ts`，字形标记反解自
     `components/ScrollToLatestButton.tsx`）；
  2. 用无头 Chrome 按 `?preset=<name>&sheet=<state>` 各渲染一张 —— `sheet=` 让页面
     只留**一张**「真实尺寸」卡片并贴到 (0,0)，预览页自己的外壳全部隐去；所以小样里
     看到的渐变、图形色、面板底色都是**实现里那一份**，不存在"示意和实机不一致"，
     裁剪框也不再受标题/说明文案换行的影响；
  3. 裁出那张卡片、拼成 2×2，并标出每个预设的圆环端点与前景图形色值。

产物：
  `outputs/progress-dot-preview/preset-sheet.png`         运行中（三点）
  `outputs/progress-dot-preview/preset-sheet-anchor.png`  已停下（锚）

依赖 Pillow（**不在项目 .venv 里**，用托管环境）：
  C:/Users/jia/.workbuddy/binaries/python/envs/default/Scripts/python.exe \
      scripts/render_progress_dot_presets.py
"""
from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

REPO = Path(__file__).resolve().parent.parent
PREVIEW_JS = REPO / "scripts" / "build_progress_dot_preview.mjs"
PREVIEW_HTML = REPO / "outputs" / "progress-dot-preview" / "index.html"
OUT_DIR = REPO / "outputs" / "progress-dot-preview"

CHROME_CANDIDATES = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
]
NODE_CANDIDATES = [
    r"C:\Users\jia\.workbuddy\binaries\node\versions\22.22.2-6\node.exe",
]

# 四个预设的圆环端点与前景图形色值 —— 与 useTheme.ts / :root 一致，仅用于**在图上标注**，
# 不参与渲染（渲染用的是真样式表那一份）。若与源码不一致，下面会打印警告。
PRESETS = [
    ("default", "#edf3ec → #e3ebe2", "#265639"),
    ("warm-orange", "#ffe4c4 → #ffd9b0", "#ae5f2e"),
    ("calm-blue", "#d4e8ff → #c2deff", "#2f64ab"),
    ("vivid-green", "#c8f5d9 → #b5f0ca", "#247c4f"),
]

# 两种状态各出一张。`state` 就是预览页的 `?sheet=` 取值，同时也是那一行的类名。
STATES = [
    ("running", "运行中 · 三点（EB 在跑，42px 浮标）", "preset-sheet.png"),
    ("idle", "已停下 · 锚（EB 已结束，42px 浮标）", "preset-sheet-anchor.png"),
]

# 截图模式下卡片贴在 (0,0)、宽 520 CSS、高 190(面板)+2(边框) = 192 CSS。
# 设备像素比 2 ⇒ 卡片正好是 1040×384 设备像素，这就是裁剪框本身。
WINDOW = (560, 220)
CROP = (0, 0, 1040, 384)
SCALE = 2  # 拼图时缩到 1/2，成品约 1000px 宽
LABEL_FONT = r"C:\Windows\Fonts\msyh.ttc"


def first_existing(paths: list[str], what: str) -> str:
    for p in paths:
        if Path(p).exists():
            return p
    raise SystemExit(f"找不到{what}：{'、'.join(paths)}")


def assert_presets_match_sources() -> None:
    """小样上的标注必须与真实源一致，否则这份小样会骗人。"""
    styles = (REPO / "apps" / "excel-addin" / "src" / "styles.css").read_text("utf-8")
    hook = (REPO / "apps" / "excel-addin" / "src" / "hooks" / "useTheme.ts").read_text("utf-8")
    problems = []
    # 字段名也要在：改名后仍拿旧名去标注，就等于小样描述了一个不存在的令牌。
    if "progressGlyph:" not in hook:
        problems.append("useTheme.ts 里找不到 progressGlyph 字段")
    for name, _, glyph in PRESETS:
        if glyph not in hook:
            problems.append(f"{name} 的前景色 {glyph} 在 useTheme.ts 里找不到")
    root_block = styles[styles.index(":root {"):].split("\n}")[0]
    if "--color-progress-glyph:" not in root_block:
        problems.append(":root 里没有 --color-progress-glyph")
    elif PRESETS[0][2] not in root_block:
        problems.append(f":root 里没有 default 的前景色 {PRESETS[0][2]}")
    if problems:
        raise SystemExit("小样标注与源码不一致：\n  - " + "\n  - ".join(problems))


def render_sheet(
    chrome: str,
    tmp: Path,
    state: str,
    title: str,
    out: Path,
    font: ImageFont.FreeTypeFont | ImageFont.ImageFont,
) -> tuple[int, int]:
    shots: list[Path] = []
    for name, _, _ in PRESETS:
        shot = tmp / f"{state}-{name}.png"
        # 注意：`--screenshot` 必须是 **Windows 绝对路径**，相对路径会报
        # "系统找不到指定的路径"。
        subprocess.run(
            [
                chrome,
                "--headless=new",
                "--disable-gpu",
                "--hide-scrollbars",
                "--force-device-scale-factor=2",
                f"--window-size={WINDOW[0]},{WINDOW[1]}",
                "--virtual-time-budget=1500",
                f"--screenshot={shot}",
                f"{PREVIEW_HTML.as_uri()}?preset={name}&sheet={state}",
            ],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        shots.append(shot)

    tiles = []
    for shot in shots:
        im = Image.open(shot).convert("RGB").crop(CROP)
        tiles.append(im.resize((im.width // SCALE, im.height // SCALE), Image.LANCZOS))

    w, h = tiles[0].size
    pad, top = 16, 48
    sheet = Image.new("RGB", (w * 2 + pad * 3, (h + top) * 2 + pad * 3), (252, 252, 251))
    draw = ImageDraw.Draw(sheet)
    for i, tile in enumerate(tiles):
        name, ring, glyph = PRESETS[i]
        x = pad + (i % 2) * (w + pad)
        y = pad + (i // 2) * (h + top + pad)
        draw.text((x + 2, y + 3), f"{name} · {title}", fill=(35, 35, 35), font=font)
        draw.text((x + 2, y + 24), f"圆环 {ring}   前景 {glyph}", fill=(110, 110, 110), font=font)
        sheet.paste(tile, (x, y + top))
        draw.rectangle([x - 1, y + top - 1, x + w, y + top + h], outline=(205, 205, 203))

    out.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(out)
    return sheet.size


def main() -> None:
    assert_presets_match_sources()

    node = first_existing(NODE_CANDIDATES, "node")
    chrome = first_existing(CHROME_CANDIDATES, "无头浏览器")

    subprocess.run([node, str(PREVIEW_JS)], check=True, cwd=REPO)
    if not PREVIEW_HTML.exists():
        raise SystemExit(f"预览页没生成：{PREVIEW_HTML}")

    try:
        font = ImageFont.truetype(LABEL_FONT, 17)
    except OSError:
        font = ImageFont.load_default()

    tmp = Path(tempfile.mkdtemp(prefix="progress-dot-presets-"))
    try:
        for state, title, filename in STATES:
            size = render_sheet(chrome, tmp, state, title, OUT_DIR / filename, font)
            print(f"已生成 {OUT_DIR / filename} {size}")
    finally:
        # 中间截图是**本脚本自己建的**临时目录，用完即清，不在 TEMP 里留垃圾。
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
