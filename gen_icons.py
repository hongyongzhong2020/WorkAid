# -*- coding: utf-8 -*-
"""从 RemixIcon 官方 SVG 包（_icons_pkg/remix.tgz）生成 WorkAid 按钮用白色 PNG 图标。

渲染方式：svglib + reportlab(rlPyCairo) 白底渲染，
再用 Pillow 把亮度转成 Alpha（白线黑底 → 白色抗锯齿透明图标）。

用法（隔离 venv）：
  C:\\Users\\Administrator\\.workbuddy\\binaries\\python\\envs\\default\\Scripts\\python.exe gen_icons.py
"""
import os
import tarfile

from svglib.svglib import svg2rlg
from reportlab.graphics import renderPM
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
TARBALL = os.path.join(HERE, "_icons_pkg", "remix.tgz")
OUT_DIR = os.path.join(HERE, "icons")
SIZE = 192          # 渲染尺寸（8x 超采样）
FINAL = 16          # 按钮显示尺寸（逻辑像素），LANCZOS 降到成品大小
VIEW = 24           # RemixIcon viewBox

# 按钮文字 → RemixIcon SVG（package/icons/ 下的路径）
BUTTON_ICONS = {
    "add_file":    "Document/file-add-line.svg",        # 添加文件
    "add_pic":     "Media/image-add-line.svg",          # 添加图片
    "add_folder":  "Document/folder-add-line.svg",      # 添加文件夹
    "remove":      "System/indeterminate-circle-line.svg",  # 移除选中
    "clear":       "System/delete-bin-line.svg",        # 清空列表
    "browse":      "Document/folder-open-line.svg",     # 浏览…
    "save":        "Device/save-line.svg",              # 另存为…
    "up":          "Arrows/arrow-up-line.svg",          # 上移
    "down":        "Arrows/arrow-down-line.svg",        # 下移
    "play":        "Media/play-line.svg",               # 开始转换 / 开始合并
    "stop":        "Media/stop-line.svg",               # 停止
}


def render_one(svg_text, out_path):
    # 黑线白底渲染，再把亮度取反作为 Alpha（黑线→不透明白、抗锯齿边缘平滑）
    svg_text = svg_text.replace("currentColor", "#000000")
    tmp = out_path + ".tmp.svg"
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(svg_text)
    try:
        d = svg2rlg(tmp)
        s = SIZE / VIEW
        d.width, d.height = SIZE, SIZE
        d.scale(s, s)
        png_tmp = out_path + ".tmp.png"
        renderPM.drawToFile(d, png_tmp, fmt="PNG", bg=0xFFFFFF)
        lum = Image.open(png_tmp).convert("L")
        from PIL import ImageOps
        alpha = ImageOps.invert(lum)
        rgba = Image.new("RGBA", (SIZE, SIZE), (255, 255, 255, 0))
        rgba.putalpha(alpha)
        rgba = rgba.resize((FINAL, FINAL), Image.LANCZOS)
        rgba.save(out_path)
    finally:
        for p in (tmp, out_path + ".tmp.png"):
            if os.path.exists(p):
                os.remove(p)


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    tar = tarfile.open(TARBALL)
    for key, rel in BUTTON_ICONS.items():
        svg = tar.extractfile("package/icons/" + rel).read().decode("utf-8")
        out = os.path.join(OUT_DIR, key + ".png")
        render_one(svg, out)
        print("OK", key, "<-", rel)
    print("done:", OUT_DIR)


if __name__ == "__main__":
    main()
