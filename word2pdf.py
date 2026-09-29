# -*- coding: utf-8 -*-
"""
WorkAid · 本地文档工具箱
- Word 转 PDF：调用 COM 接口（优先 Microsoft Word，回退 WPS 文字），批量转为 PDF。
- PDF 转图片：基于 PyMuPDF，将 PDF 每一页导出为 PNG / JPG。
- 图片转 PDF：基于 PyMuPDF，把多张图片按顺序合并为一个 PDF。
- PDF 合并：基于 PyMuPDF，把多个 PDF 按顺序合并为一个文件。
- PDF 拆分：基于 PyMuPDF，把一个 PDF 每页拆分或按页码范围提取为多个 PDF。
- PDF 加水印：基于 PyMuPDF，批量给 PDF 加文字 / 图片水印（平铺或居中）。
- PPT 转 PDF：调用 COM 接口（优先 Microsoft PowerPoint，回退 WPS 演示），批量导出 PDF。
全部在本地完成，文件不上传。
运行前提：Word 转 PDF 需本机已安装 Microsoft Word 或 WPS；PPT 转 PDF 需本机
已安装 Microsoft PowerPoint 或 WPS。
"""

import math
import os
import re
import sys
import ctypes
import queue
import threading
import traceback
import datetime
from tkinter import filedialog, messagebox
from tkinter import font as tkfont

# 运行环境兼容：优先用 ttkbootstrap 做现代化界面，缺失时退回标准 ttk
try:
    import ttkbootstrap as tb
    HAS_TTB = True
except Exception:
    tb = None
    HAS_TTB = False

# 拖拽支持：缺失时退回普通 Tk
try:
    import tkinterdnd2 as tkdnd
    HAS_DND = True
except Exception:
    tkdnd = None
    HAS_DND = False

import tkinter as tk
from tkinter import ttk

# ---------------------------------------------------------------------------
# 高 DPI 适配（必须在创建任何窗口之前调用）
# 声明 Per-Monitor DPI Aware，让系统不再对界面做位图拉伸，
# 从根本上消除高分屏（125%/150% 等）下文字发虚、圆角发毛的问题。
# ---------------------------------------------------------------------------
def _enable_dpi_awareness():
    """开启进程 DPI 感知，返回真实缩放比例（96dpi = 1.0）。"""
    if sys.platform != "win32":
        return 1.0
    try:
        # 2 = PROCESS_PER_MONITOR_DPI_AWARE（Win8.1+）
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()   # Win7 回退
        except Exception:
            return 1.0
    try:
        hdc = ctypes.windll.user32.GetDC(0)
        LOGPIXELSX = 88
        dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, LOGPIXELSX)
        ctypes.windll.user32.ReleaseDC(0, hdc)
        if dpi:
            return dpi / 96.0
    except Exception:
        pass
    return 1.0


UI_SCALE = _enable_dpi_awareness()


def s(value):
    """把按 96dpi 设计的逻辑像素值换算为当前屏幕的真实像素值。

    界面中所有硬编码的尺寸（窗口大小、内边距、控件长宽、圆角半径等）
    都是按 100% 缩放设计的，统一经此函数换算后，
    在 125% / 150% 等缩放下视觉大小保持一致。
    """
    if UI_SCALE == 1.0:
        return value
    return int(round(value * UI_SCALE))


def _apply_titlebar_colors(root, bg="#111111", text="#f0f0f0", border="#111111"):
    """Windows 11（Build 22000+）原生标题栏染色，与应用深色主题融为一体。

    通过 DwmSetWindowAttribute 设置标题栏背景 / 文字 / 窗口边框颜色，
    仅作用于本应用窗口，不影响系统其它部分，退出后无残留。
    Windows 10 及更早系统不支持这些属性（返回错误码），静默跳过保持默认。
    """
    if sys.platform != "win32":
        return
    try:
        dwm = ctypes.windll.dwmapi
        # tkinter 顶层窗口真正的 HWND 是内部子窗口的父窗口
        hwnd = ctypes.windll.user32.GetParent(root.winfo_id())

        def _colorref(hx):
            hx = hx.lstrip("#")
            r, g, b = int(hx[0:2], 16), int(hx[2:4], 16), int(hx[4:6], 16)
            return ctypes.c_uint((b << 16) | (g << 8) | r)   # COLORREF = 0x00BBGGRR

        def _apply(attr, value):
            dwm.DwmSetWindowAttribute(hwnd, attr, ctypes.byref(value), ctypes.sizeof(value))

        _apply(20, ctypes.c_uint(1))       # 深色模式：最小化/关闭按钮用浅色图标
        _apply(35, _colorref(bg))          # DWMWA_CAPTION_COLOR 标题栏背景
        _apply(36, _colorref(text))        # DWMWA_TEXT_COLOR 标题文字
        _apply(34, _colorref(border))      # DWMWA_BORDER_COLOR 窗口边框
    except Exception:
        pass

# COM 接口（Windows + Word 必需）
try:
    import win32com.client as win32
    import pythoncom
    HAS_WIN32 = True
except Exception:
    HAS_WIN32 = False

APP_NAME = "WorkAid"
APP_VERSION = "1.11.7"

# 版权与反馈信息（起始年份固定，结束年份自动取当前系统年份）
COPYRIGHT_START_YEAR = 2026
COPYRIGHT_AUTHOR = "hongyongzhong"
FEEDBACK_EMAIL = "hongyongzhong2020@126.com"


def copyright_text():
    """底部版权文字；第二个年份自动取当天所在年份。"""
    return f"© {COPYRIGHT_START_YEAR}-{datetime.datetime.now().year} " \
           f"{COPYRIGHT_AUTHOR}   Email: {FEEDBACK_EMAIL}"

# 支持的文件后缀
SUPPORTED_EXT = (".doc", ".docx", ".docm", ".rtf", ".wps")

# wdFormatPDF
WD_FORMAT_PDF = 17

# 状态 key → 显示文字
STATUS_TEXT = {
    "pending": "待转换",
    "working": "转换中",
    "ok": "完成",
    "error": "失败",
    "skip": "跳过",
}

# 状态 key → 前景色（None 表示用默认色）
STATUS_COLOR = {
    "ok": "#2ecc71",
    "error": "#e74c3c",
    "working": "#f39c12",
    "skip": "#95a5a6",
    "pending": None,
}


# ---------------------------------------------------------------------------
# 组件工厂：统一 ttkbootstrap 与标准 ttk 的差异
# ---------------------------------------------------------------------------
def _widget_class(name):
    """返回 ttkbootstrap（若有）或标准 ttk 的组件类。"""
    if HAS_TTB:
        return getattr(tb, name)
    return getattr(ttk, name)


def _resource_path(rel):
    """兼容打包后的资源路径（PyInstaller _MEIPASS）。"""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


# ---------------------------------------------------------------------------
# 圆角外观基础设施
# 界面内的框（按钮 / 输入框 / 下拉框 / 勾选框 / 进度条 / 列表 / 日志 /
# 拖拽区 / 标签页）统一改为 Canvas 自绘的圆角样式。
# ---------------------------------------------------------------------------
CARD_RADIUS = 12        # 卡片类容器圆角半径
CTRL_RADIUS = 12        # 按钮 / 输入框 / 下拉框圆角半径（统一为 12）
UI_FONT = "Microsoft YaHei UI"
UI_FONT_SIZE = 10

PALETTE = {
    "bg": "#111111",        # 窗口背景（黑色）
    "card": "#1d1d1d",      # 卡片 / 框背景
    "border": "#3a3a3a",    # 边框
    "fg": "#ffffff",        # 主文字
    "muted": "#8b949e",     # 次要文字
    "input": "#202020",     # 输入类控件背景（无边框方案下靠底色差区分）
    "primary": "#007AFF",
    "secondary": "#2c2c2c",
    "success": "#00bc8c",
    "danger": "#e74c3c",
    "info": "#3498db",
    "warning": "#f39c12",
    "track": "#292929",     # 进度条底槽
}


def _hex_to_rgb(color):
    color = str(color).lstrip("#")
    return tuple(int(color[i:i + 2], 16) for i in (0, 2, 4))


def _rgb_to_hex(rgb):
    return "#%02x%02x%02x" % tuple(
        max(0, min(255, int(round(v)))) for v in rgb)


def _mix(color_a, color_b, ratio):
    """把 color_a 向 color_b 混合 ratio（0~1）比例。"""
    a, b = _hex_to_rgb(color_a), _hex_to_rgb(color_b)
    return _rgb_to_hex([a[i] + (b[i] - a[i]) * ratio for i in range(3)])


def _shade(color, amount):
    """amount > 0 提亮，< 0 压暗。"""
    return _mix(color, "#ffffff" if amount > 0 else "#000000", abs(amount))


_ARC_QUARTER = math.pi / 2.0


def round_rect(canvas, x1, y1, x2, y2, radius, steps=None, **kwargs):
    """在 Canvas 上绘制圆角矩形（真圆弧采样），返回 item id。

    每个角沿 90° 圆弧按数学公式采样若干点，形成真正的圆形轮廓。
    相比旧版“直角顶点 + 样条平滑”的近似做法（每角仅 3 点，
    splinesteps 实际不生效、圆弧会被样条拉成鼓包），此实现轮廓
    完全贴合理想圆角，视觉更丝滑。
    """
    radius = max(0.0, min(radius, (x2 - x1) / 2.0, (y2 - y1) / 2.0))
    if radius <= 0.5:
        points = (x1, y1, x2, y1, x2, y2, x1, y2)
        return canvas.create_polygon(points, smooth=False, **kwargs)

    # 每个角的采样段数：半径越大采样越密，保证弧长步进 ≈1px，
    # 并限定在 4~16 之间以免点数过多拖慢重绘。
    if steps is None:
        steps = int(max(4, min(16, round(radius / 1.5) + 4)))

    pts = []
    # 右上角：-90° → 0°
    for i in range(steps + 1):
        a = -_ARC_QUARTER + _ARC_QUARTER * (i / steps)
        pts.extend((x2 - radius + radius * math.cos(a),
                    y1 + radius + radius * math.sin(a)))
    # 右下角：0° → 90°
    for i in range(steps + 1):
        a = _ARC_QUARTER * (i / steps)
        pts.extend((x2 - radius + radius * math.cos(a),
                    y2 - radius + radius * math.sin(a)))
    # 左下角：90° → 180°
    for i in range(steps + 1):
        a = _ARC_QUARTER + _ARC_QUARTER * (i / steps)
        pts.extend((x1 + radius + radius * math.cos(a),
                    y2 - radius + radius * math.sin(a)))
    # 左上角：180° → 270°
    for i in range(steps + 1):
        a = math.pi + _ARC_QUARTER * (i / steps)
        pts.extend((x1 + radius + radius * math.cos(a),
                    y1 + radius + radius * math.sin(a)))
    kwargs.setdefault("smooth", False)
    return canvas.create_polygon(pts, **kwargs)


def _load_scaled_photo(rel_path):
    """加载 PNG 并在高 DPI 下平滑缩放到当前缩放尺寸，返回 tk.PhotoImage。

    优先用 PIL 的 LANCZOS 重采样（任意倍数都平滑），
    PIL 不可用时回退为不缩放（或整数倍 zoom）。
    """
    path = _resource_path(rel_path)
    img = tk.PhotoImage(file=path)
    if UI_SCALE <= 1.0:
        return img
    try:
        from PIL import Image as _PILImage, ImageTk as _PILTk
        with _PILImage.open(path) as pil:
            pil = pil.convert("RGBA")
            size = (max(1, round(pil.width * UI_SCALE)),
                    max(1, round(pil.height * UI_SCALE)))
            pil = pil.resize(size, _PILImage.LANCZOS)
            return _PILTk.PhotoImage(pil)
    except Exception:  # noqa: BLE001
        factor = max(1, int(UI_SCALE))
        return img.zoom(factor, factor) if factor > 1 else img


def render_bg_photo(master, w, h):
    """生成纯色背景图（深蓝夜色，无纹理）。"""
    photo = tk.PhotoImage(width=w, height=h, master=master)
    photo.put(PALETTE["bg"], to=(0, 0, w, h))
    return photo


class BgCanvas(tk.Canvas):
    """带渐变+方格背景的画布容器（子控件可直接 pack 进来）。"""

    def __init__(self, master, **kw):
        kw.setdefault("bd", 0)
        kw.setdefault("highlightthickness", 0)
        kw.setdefault("bg", PALETTE["bg"])
        super().__init__(master, **kw)
        self._photo = None
        self._render_after = None
        self.bind("<Configure>", self._on_configure, add="+")

    def _on_configure(self, event):
        if self._render_after is not None:
            self.after_cancel(self._render_after)
        self._render_after = self.after(
            60, lambda: self._render(event.width, event.height))

    def _render(self, w, h):
        self._render_after = None
        if w < 4 or h < 4:
            return
        try:
            self._photo = render_bg_photo(self, w, h)
        except Exception:
            self._photo = None
            return
        self.redraw_bg()

    def redraw_bg(self):
        """重铺背景图（在 delete("all") 之后调用可恢复背景）。"""
        self.delete("bgimg")
        if self._photo is not None:
            self.create_image(0, 0, image=self._photo, anchor="nw",
                              tags="bgimg")
            self.tag_lower("bgimg")


def refresh_palette(style):
    """保留自定义「深蓝夜色」配色（不再被 darkly 主题的灰色覆盖）。"""
    PALETTE["track"] = _mix(PALETTE["card"], "#ffffff", 0.10)
    return PALETTE


def register_round_styles(style):
    """把列表（Treeview）等 ttk 控件也调成卡片配色，与圆角容器协调。"""
    try:
        style.configure("Treeview",
                        background=PALETTE["card"],
                        fieldbackground=PALETTE["card"],
                        foreground=PALETTE["fg"],
                        bordercolor=PALETTE["card"],
                        lightcolor=PALETTE["card"],
                        darkcolor=PALETTE["card"],
                        borderwidth=0, relief="flat", rowheight=s(28))
        style.configure("Treeview.Heading",
                        background=PALETTE["card"],
                        foreground=PALETTE["muted"],
                        borderwidth=0, relief="flat", padding=(s(8), s(6)))
        style.map("Treeview",
                  background=[("selected", PALETTE["primary"])],
                  foreground=[("selected", PALETTE["fg"])])
        style.map("Treeview.Heading",
                  background=[("active", _shade(PALETTE["card"], 0.08))])
    except Exception:
        pass


class RoundButton(tk.Canvas):
    """圆角按钮：Canvas 自绘，兼容 ttkbootstrap 的 bootstyle 写法。

    支持在文字左侧绘制小图标：优先使用 icons/ 目录下的 RemixIcon 矢量风格
    PNG（由 gen_icons.py 从官方 SVG 生成，白色线条透明底），缺失时回退到
    Segoe MDL2 Assets 字体字形。按按钮文字自动匹配，也可用 icon= 显式指定。
    """

    VARIANTS = ("primary", "secondary", "success", "danger", "info",
                "warning", "light", "dark")

    # 文字 → icons/ 图标名（icons/<名>.png，RemixIcon line 风格）
    ICON_BY_TEXT = {
        "添加文件": "add_file",
        "添加图片": "add_pic",
        "添加文件夹": "add_folder",
        "移除选中": "remove",
        "清空列表": "clear",
        "浏览…": "browse",
        "上移": "up",
        "下移": "down",
        "开始转换": "play",
        "开始合并": "play",
        "开始拆分": "play",
        "开始加水印": "play",
        "停止": "stop",
    }
    ICON_SIZE = 16                       # 图标显示尺寸（逻辑像素）
    ICON_DIR = _resource_path("icons")   # 兼容 PyInstaller 打包后路径
    _PNG_CACHE = {}                      # 类级共享：key → tk.PhotoImage / None

    # Segoe MDL2 字形兜底（icons 缺失时使用）
    GLYPH_BY_TEXT = {
        "添加文件": "\uE710",   # Add
        "添加图片": "\uE710",
        "添加文件夹": "\uE8B7",   # Folder
        "移除选中": "\uE711",     # Remove
        "清空列表": "\uE74D",     # Delete
        "浏览…": "\uE8E5",        # OpenFile
        "上移": "\uE74A",       # ChevronUp
        "下移": "\uE74B",       # ChevronDown
        "开始转换": "\uE768",     # Play
        "开始合并": "\uE768",
        "开始拆分": "\uE768",
        "开始加水印": "\uE768",
        "停止": "\uE71A",         # Stop
    }

    @classmethod
    def _load_png(cls, key):
        """加载并缓存图标 PNG；文件缺失或无法加载时返回 None。"""
        if key in cls._PNG_CACHE:
            return cls._PNG_CACHE[key]
        path = os.path.join(cls.ICON_DIR, key + ".png")
        img = None
        if os.path.exists(path):
            try:
                img = _load_scaled_photo(os.path.join("icons", key + ".png"))
            except Exception:  # noqa: BLE001
                img = None
        cls._PNG_CACHE[key] = img
        return img

    def __init__(self, master, text="", command=None, bootstyle="secondary",
                 radius=CTRL_RADIUS, pad_x=18, pad_y=9, font=None, fg=None,
                 parent_bg=None, icon=None, **kw):
        # 高 DPI：像素类参数统一按缩放比例换算（字体已由 Tk 的 point 机制处理）
        radius, pad_x, pad_y = s(radius), s(pad_x), s(pad_y)
        self._icon_size = s(self.ICON_SIZE)   # 图标显示尺寸（当前缩放）
        self._icon_gap = s(6)                 # 图标与文字间距
        self._text = text
        key = icon if icon is not None else self.ICON_BY_TEXT.get(text, "")
        self._icon = self.GLYPH_BY_TEXT.get(text, "") if icon is None else ""
        self._photo = self._load_png(key) if key else None
        self._command = command
        self._radius = radius
        self._pad_x = pad_x
        self._pad_y = pad_y
        self._variant = self._resolve_variant(bootstyle)
        self._fg = fg or PALETTE["fg"]
        self._parent_bg = parent_bg or PALETTE["bg"]
        self._hover = False
        self._pressed = False
        self._font = font or tkfont.Font(family=UI_FONT, size=UI_FONT_SIZE)
        self._icon_font = tkfont.Font(family="Segoe MDL2 Assets",
                                      size=UI_FONT_SIZE + 1)
        self._bw, self._bh = self._measure()
        tk.Canvas.__init__(self, master, width=self._bw, height=self._bh, bd=0,
                           highlightthickness=0, bg=self._parent_bg,
                           cursor="hand2", **kw)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)
        self.bind("<Configure>", self._on_resize)
        self._draw()

    def _on_resize(self, event):
        """控件被拉伸/缩放时同步缓存尺寸并重绘，保证内容始终居中。"""
        if event.width > 1 and event.height > 1 and \
                (event.width != self._bw or event.height != self._bh):
            self._bw, self._bh = event.width, event.height
            self._draw()

    # -- 内部 ---------------------------------------------------------------
    @staticmethod
    def _resolve_variant(style):
        if not style:
            return "secondary"
        name = str(style).split()[0].lower()
        for suffix in ("-outline", "-link", "-ghost", "-inverse"):
            if name.endswith(suffix):
                name = name[: -len(suffix)]
                break
        return name if name in RoundButton.VARIANTS else "secondary"

    def _measure(self):
        text_w = self._font.measure(self._text) if self._text else 0
        if self._photo is not None:
            icon_w = self._icon_size + self._icon_gap
        elif self._icon:
            icon_w = self._icon_font.measure(self._icon) + self._icon_gap
        else:
            icon_w = 0
        line_h = max(self._font.metrics("linespace"),
                     self._icon_font.metrics("linespace") if self._icon else 0)
        return text_w + icon_w + 2 * self._pad_x, line_h + 2 * self._pad_y + 2

    def _draw(self):
        self.delete("all")
        base = PALETTE.get(self._variant, PALETTE["secondary"])
        fill = base
        if self._pressed:
            fill = _shade(base, -0.14)
        elif self._hover:
            fill = _shade(base, 0.16)
        round_rect(self, 1, 1, self._bw - 1, self._bh - 1, self._radius,
                   fill=fill, outline="", width=0,
                   tags="btnbg")
        cy = self._bh / 2.0
        text_w = self._font.measure(self._text) if self._text else 0
        if self._photo is not None:
            icon_w = self._icon_size
        elif self._icon:
            icon_w = self._icon_font.measure(self._icon)
        else:
            icon_w = 0
        block = text_w + (icon_w + self._icon_gap if icon_w else 0)
        x = (self._bw - block) / 2.0
        if self._photo is not None:
            self.create_image(x + self._icon_size / 2.0, cy,
                              image=self._photo, tags="content")
            x += self._icon_size + self._icon_gap
        elif icon_w:
            self.create_text(x + icon_w / 2.0, cy, text=self._icon,
                             fill=self._fg, font=self._icon_font,
                             tags="content")
            x += icon_w + self._icon_gap
        if text_w:
            self.create_text(x + text_w / 2.0, cy, text=self._text,
                             fill=self._fg, font=self._font,
                             tags="content")
        # 二次精调：按内容真实包围盒整体平移，消除字体墨迹偏差，
        # 保证图标+文字在按钮内严格水平/垂直居中（与侧边栏一致）。
        content = self.find_withtag("content")
        if content:
            boxes = [self.bbox(i) for i in content]
            x0 = min(b[0] for b in boxes)
            x1 = max(b[2] for b in boxes)
            y0 = min(b[1] for b in boxes)
            y1 = max(b[3] for b in boxes)
            dx = (self._bw - (x0 + x1)) / 2.0
            dy = (self._bh - (y0 + y1)) / 2.0
            if abs(dx) > 0.05 or abs(dy) > 0.05:
                self.move("content", dx, dy)

    # -- 事件 ---------------------------------------------------------------
    def _on_enter(self, _event=None):
        self._hover = True
        self._draw()

    def _on_leave(self, _event=None):
        self._hover = False
        self._pressed = False
        self._draw()

    def _on_press(self, _event=None):
        self._pressed = True
        self._draw()

    def _on_release(self, _event=None):
        fire = self._pressed
        self._pressed = False
        self._draw()
        if fire and self._command is not None:
            self._command()

    # -- 接口兼容 -----------------------------------------------------------
    def configure(self, cnf=None, **kw):
        if cnf:
            kw = dict(cnf, **kw)
        changed = False
        if "text" in kw:
            self._text = kw.pop("text")
            key = self.ICON_BY_TEXT.get(self._text, "")
            self._photo = self._load_png(key) if key else None
            self._icon = self.GLYPH_BY_TEXT.get(self._text, "")
            changed = True
        if "bootstyle" in kw:
            self._variant = self._resolve_variant(kw.pop("bootstyle"))
            changed = True
        if "font" in kw:
            self._font = kw.pop("font")
            changed = True
        if kw:
            tk.Canvas.configure(self, **kw)
        if changed:
            self._bw, self._bh = self._measure()
            tk.Canvas.configure(self, width=self._bw, height=self._bh)
            self._draw()

    config = configure

    def pack(self, cnf=None, **kw):
        kw.pop("ipadx", None)      # 自绘控件已含内边距，忽略 ttk 的 ipad
        kw.pop("ipady", None)
        return tk.Canvas.pack(self, cnf, **kw)

    def invoke(self):
        if self._command is not None:
            self._command()


class RoundEntry(tk.Frame):
    """圆角输入框：Canvas 圆角底 + 无边框 Entry。"""

    def __init__(self, master, textvariable=None, height=34, radius=CTRL_RADIUS,
                 pad_x=12, font=None, parent_bg=None, **kw):
        # 高 DPI：像素类参数按缩放换算
        height, radius, pad_x = s(height), s(radius), s(pad_x)
        self._parent_bg = parent_bg or PALETTE["bg"]
        self._radius = radius
        self._pad_x = pad_x
        self._eh = height
        tk.Frame.__init__(self, master, bg=self._parent_bg, height=height, **kw)
        self.pack_propagate(False)
        self.canvas = tk.Canvas(self, bd=0, highlightthickness=0,
                                bg=self._parent_bg, height=height)
        self.canvas.pack(fill="both", expand=True)
        self.entry = tk.Entry(
            self.canvas, textvariable=textvariable, bd=0, relief="flat",
            highlightthickness=0, bg=PALETTE["input"], fg=PALETTE["fg"],
            insertbackground=PALETTE["fg"], disabledbackground=PALETTE["input"],
            selectbackground=PALETTE["primary"], selectforeground=PALETTE["fg"],
            font=font or tkfont.Font(family=UI_FONT, size=UI_FONT_SIZE))
        self._win = self.canvas.create_window(
            pad_x, height / 2.0, window=self.entry, anchor="w",
            height=max(s(18), height - s(14)))
        self.canvas.bind("<Configure>", self._redraw)
        self._redraw()

    def _redraw(self, _event=None):
        w = self.canvas.winfo_width() or 200
        h = self.canvas.winfo_height() or self._eh
        self.canvas.delete("cardbg")
        round_rect(self.canvas, 1, 1, w - 1, h - 1, self._radius,
                   fill=PALETTE["input"], outline="", width=0,
                   tags="cardbg")
        self.canvas.tag_lower("cardbg")
        self.canvas.itemconfigure(self._win,
                                  width=max(s(20), w - 2 * self._pad_x),
                                  height=max(s(18), h - s(14)))

    # 让外部像用 Entry 一样拿到变量
    def get(self):
        return self.entry.get()

    def configure(self, cnf=None, **kw):
        """兼容 configure：state 转发给内部 Entry，其余交给 Frame。"""
        if cnf:
            kw = dict(cnf, **kw)
        entry_kw = {}
        for key in ("state",):
            if key in kw:
                entry_kw[key] = kw.pop(key)
        if kw:
            tk.Frame.configure(self, **kw)
        if entry_kw:
            self.entry.configure(**entry_kw)

    config = configure


class RoundCombo(tk.Canvas):
    """圆角下拉框：自绘外观 + 菜单弹出，兼容 Combobox 常用接口。"""

    def __init__(self, master, values=(), width=12, height=34,
                 radius=CTRL_RADIUS, pad_x=12, font=None, parent_bg=None,
                 command=None, **kw):
        # 高 DPI：像素类参数按缩放换算
        height, radius, pad_x = s(height), s(radius), s(pad_x)
        self._values = [str(v) for v in values]
        self._index = 0 if self._values else -1
        self._command = command
        self._radius = radius
        self._pad_x = pad_x
        self._parent_bg = parent_bg or PALETTE["bg"]
        self._font = font or tkfont.Font(family=UI_FONT, size=UI_FONT_SIZE)
        char_w = self._font.measure("0") or 8
        widest = max([self._font.measure(v) for v in self._values] + [0])
        w = max(widest + 2 * pad_x + s(26), int(char_w * int(width)) + s(24))
        tk.Canvas.__init__(self, master, width=w, height=height, bd=0,
                           highlightthickness=0, bg=self._parent_bg,
                           cursor="hand2", **kw)
        self._menu = tk.Menu(self, tearoff=False, bd=0,
                             bg=PALETTE["card"], fg=PALETTE["fg"],
                             activebackground=PALETTE["primary"],
                             activeforeground=PALETTE["fg"], font=self._font)
        for i, value in enumerate(self._values):
            self._menu.add_command(label=value,
                                   command=lambda i=i: self.select_index(i))
        self.bind("<Button-1>", self._popup)
        self._draw()

    def _draw(self):
        self.delete("all")
        w = float(self.cget("width"))
        h = float(self.cget("height"))
        round_rect(self, 1, 1, w - 1, h - 1, self._radius,
                   fill=PALETTE["input"], outline="", width=0)
        text = self.get()
        self.create_text(self._pad_x, h / 2.0, text=text, anchor="w",
                         fill=PALETTE["fg"], font=self._font)
        # 右侧下拉箭头
        ax, ay = w - self._pad_x - 8, h / 2.0
        self.create_polygon(ax, ay - 2, ax + 9, ay - 2, ax + 4.5, ay + 4,
                            fill=PALETTE["muted"], outline="")

    def _popup(self, _event=None):
        if not self._values:
            return
        try:
            self._menu.tk_popup(self.winfo_rootx(),
                                self.winfo_rooty() + self.winfo_height())
        finally:
            self._menu.grab_release()

    # -- 接口兼容 -----------------------------------------------------------
    def current(self, index=None):
        if index is None:
            return self._index
        self.select_index(int(index))
        return None

    def get(self):
        if 0 <= self._index < len(self._values):
            return self._values[self._index]
        return ""

    def set(self, value):
        if value in self._values:
            self.select_index(self._values.index(value))

    def select_index(self, index):
        if not (0 <= index < len(self._values)):
            return
        if index != self._index:
            self._index = index
            self._draw()
        if self._command is not None:
            self._command(self.get())


class RoundProgress(tk.Canvas):
    """圆角进度条（胶囊形），支持 progress["value"] / ["maximum"] 读写。"""

    def __init__(self, master, mode="determinate", height=14,
                 parent_bg=None, **kw):
        height = s(height)                      # 高 DPI：按缩放换算
        self._value = 0.0
        self._maximum = 100.0
        self._parent_bg = parent_bg or PALETTE["bg"]
        self._ph = height
        tk.Canvas.__init__(self, master, height=height, bd=0,
                           highlightthickness=0, bg=self._parent_bg, **kw)
        self.bind("<Configure>", lambda _e: self._draw())

    def __setitem__(self, key, value):
        if key == "value":
            self._value = float(value)
        elif key == "maximum":
            self._maximum = max(1.0, float(value))
        else:
            tk.Canvas.configure(self, **{key: value})
            return
        self._draw()

    def __getitem__(self, key):
        if key == "value":
            return self._value
        if key == "maximum":
            return self._maximum
        return tk.Canvas.cget(self, key)

    def configure(self, cnf=None, **kw):
        if cnf:
            kw = dict(cnf, **kw)
        redraw = False
        if "value" in kw:
            self._value = float(kw.pop("value"))
            redraw = True
        if "maximum" in kw:
            self._maximum = max(1.0, float(kw.pop("maximum")))
            redraw = True
        if kw:
            tk.Canvas.configure(self, **kw)
        if redraw:
            self._draw()

    config = configure

    def _draw(self):
        self.delete("all")
        w = self.winfo_width() or 200
        h = self.winfo_height() or self._ph
        if w <= 2 or h <= 2:
            return
        radius = h / 2.0
        round_rect(self, 0, 0, w, h, radius,
                   fill=PALETTE["track"], outline="")
        ratio = 0.0 if self._maximum <= 0 else self._value / self._maximum
        ratio = max(0.0, min(1.0, ratio))
        if ratio > 0:
            bar_w = max(h, w * ratio)
            round_rect(self, 0, 0, bar_w, h, radius,
                       fill=PALETTE["success"], outline="")


class RoundCheck(tk.Canvas):
    """圆角勾选框：自绘圆角方块 + 对勾，绑定 BooleanVar。"""

    BOX = 18

    def __init__(self, master, text="", variable=None, command=None,
                 parent_bg=None, font=None, **kw):
        # 高 DPI：方块尺寸与间距按缩放换算
        self._box = s(self.BOX)
        self._gap = s(9)
        self._text = text
        self._var = variable if variable is not None else tk.BooleanVar(False)
        self._command = command
        self._font = font or tkfont.Font(family=UI_FONT, size=9)
        self._parent_bg = parent_bg or PALETTE["bg"]
        line_h = self._font.metrics("linespace")
        self._bh = max(self._box, line_h) + s(8)
        self._bw = self._box + self._gap + self._font.measure(text) + s(4)
        tk.Canvas.__init__(self, master, width=self._bw, height=self._bh, bd=0,
                           highlightthickness=0, bg=self._parent_bg,
                           cursor="hand2", **kw)
        self.bind("<Button-1>", self._toggle)
        try:
            self._trace = self._var.trace_add("write", lambda *_: self._draw())
        except Exception:
            self._trace = None
        self._draw()

    def _draw(self):
        self.delete("all")
        h = float(self._bh)
        box = self._box
        top = (h - box) / 2.0
        checked = bool(self._var.get())
        fill = PALETTE["primary"] if checked else PALETTE["input"]
        round_rect(self, 0, top, box, top + box, s(CTRL_RADIUS),
                   fill=fill, outline="", width=0)
        if checked:
            # 对勾按方块尺寸等比换算（基准 18px 方块时的坐标）
            k = box / 18.0
            self.create_line(4.0 * k, top + 9.0 * k,
                             7.2 * k, top + 13.0 * k,
                             14.0 * k, top + 5.5 * k,
                             fill=PALETTE["fg"], width=max(1, s(2)),
                             capstyle="round", joinstyle="round")
        self.create_text(box + self._gap, h / 2.0, text=self._text, anchor="w",
                         fill=PALETTE["fg"], font=self._font)

    def _toggle(self, _event=None):
        self._var.set(not bool(self._var.get()))
        if self._command is not None:
            self._command()

    def cget(self, key):
        if key == "text":
            return self._text
        if key == "variable":
            return str(self._var)
        return tk.Canvas.cget(self, key)


class RoundCard(tk.Frame):
    """圆角卡片容器：需要放置的控件请加到 .body 上。"""

    def __init__(self, master, radius=CARD_RADIUS, padding=8, bg=None,
                 outline=None, parent_bg=None, height=None, **kw):
        # 高 DPI：圆角半径与内边距按缩放换算
        radius, padding = s(radius), s(padding)
        if height:
            height = s(height)
        self._parent_bg = parent_bg or PALETTE["bg"]
        self._bg = bg or PALETTE["card"]
        # outline 参数保留以兼容旧调用签名；外观已改为无描边。
        self._outline = outline if outline is not None else PALETTE["border"]
        self._radius = radius
        self._pad = padding
        tk.Frame.__init__(self, master, bg=self._parent_bg, height=height, **kw)
        if height:
            self.pack_propagate(False)
        self.canvas = tk.Canvas(self, bd=0, highlightthickness=0,
                                bg=self._parent_bg, height=height or 0)
        self.canvas.pack(fill="both", expand=True)
        self.body = tk.Frame(self.canvas, bg=self._bg)
        self._win = self.canvas.create_window(padding, padding,
                                              window=self.body, anchor="nw")
        self.canvas.bind("<Configure>", self._on_configure)

    def _on_configure(self, event):
        w, h = event.width, event.height
        self.canvas.delete("cardbg")
        round_rect(self.canvas, 1, 1, w - 1, h - 1, self._radius,
                   fill=self._bg, outline="", width=0,
                   tags="cardbg")
        self.canvas.tag_lower("cardbg")
        self.canvas.itemconfigure(self._win,
                                  width=max(1, w - 2 * self._pad),
                                  height=max(1, h - 2 * self._pad))


class RoundNotebook(tk.Frame):
    """侧边栏标签页容器：LOGO 在侧栏顶部 + 纵向标签按钮，右侧为内容区。

    接口兼容 ttk.Notebook 的常用子集（add / select / tabs / index / tab）。
    """

    def __init__(self, master, bar_height=46, parent_bg=None, **kw):
        self._parent_bg = parent_bg or PALETTE["bg"]
        tk.Frame.__init__(self, master, bg=self._parent_bg, **kw)
        self._tabs = []            # [[child, text], ...]
        self._index = -1
        self._boxes = []           # [(y1, y2), ...] 各标签按钮的纵向范围
        self._font = tkfont.Font(family=UI_FONT, size=11, weight="bold")
        self._brand_font = tkfont.Font(family=UI_FONT, size=13, weight="bold")
        self._ver_font = tkfont.Font(family=UI_FONT, size=8)
        self._logo_img = None
        try:
            self._logo_img = _load_scaled_photo("logo_small.png")
        except Exception:  # noqa: BLE001
            self._logo_img = None

        # 侧边栏：LOGO + 名称 + 纵向标签（底色与窗口统一，靠标签色块与右分隔线区分）
        self._sidebar = tk.Canvas(self, width=s(174), bd=0, highlightthickness=0,
                                  bg=PALETTE["bg"], cursor="hand2")
        self._sidebar.pack(side="left", fill="y")
        self._sidebar.bind("<Button-1>", self._on_click)
        self._sidebar.bind("<Configure>", lambda _e: self._redraw(), add="+")

        # 右侧内容区
        self._body = tk.Frame(self, bg=self._parent_bg)
        self._body.pack(side="left", fill="both", expand=True)
        self.body = self._body       # 供外部放置标签页容器

    # -- 绘制 ---------------------------------------------------------------
    def _redraw(self):
        c = self._sidebar
        w = c.winfo_width() or s(174)
        h = c.winfo_height() or s(600)
        c.delete("all")
        self._boxes = []
        # 右侧分隔线
        c.create_line(w - 1, 0, w - 1, h, fill=PALETTE["border"])

        # 顶部 LOGO + 产品名 + 版本
        y = s(22)
        if self._logo_img is not None:
            c.create_image(w / 2.0, y + s(28), image=self._logo_img)
            y += s(66)
        c.create_text(w / 2.0, y + s(10), text=APP_NAME,
                      fill=PALETTE["fg"], font=self._brand_font)
        c.create_text(w / 2.0, y + s(28), text=f"v{APP_VERSION}",
                      fill=PALETTE["muted"], font=self._ver_font)

        # 纵向标签按钮
        y = y + s(52)
        bw = w - s(24)
        x1 = s(12)
        n = max(len(self._tabs), 1)
        # 标签较多时自动压缩高度与间距，保证 8 个标签也能完整显示
        bh = s(38)
        gap = s(8)
        avail = h - y - s(12)         # 底部留 12px 余量
        if n * (bh + gap) - gap > avail:
            gap = max(s(4), int((avail - n * s(30)) / max(n - 1, 1)))
            bh = max(s(26), int((avail - gap * (n - 1)) / n))
        for i, (_child, text) in enumerate(self._tabs):
            label = text.strip() or ("标签 %d" % (i + 1))
            active = (i == self._index)
            fill = PALETTE["primary"] if active else PALETTE["secondary"]
            round_rect(c, x1, y, x1 + bw, y + bh, s(CTRL_RADIUS),
                       fill=fill, outline="", width=0)
            c.create_text(x1 + bw / 2.0, y + bh / 2.0, text=label,
                          fill=PALETTE["fg"] if active else PALETTE["muted"],
                          font=self._font)
            self._boxes.append((y, y + bh))
            y += bh + gap

    def _on_click(self, event):
        for i, (y1, y2) in enumerate(self._boxes):
            if y1 <= event.y <= y2:
                self.select(i)
                return

    # -- 接口兼容 -----------------------------------------------------------
    def add(self, child, text="", **_kw):
        self._tabs.append([child, text])
        if self._index < 0:
            self.select(0)
        else:
            self._redraw()
        return child

    def select(self, tab_id=0):
        if not isinstance(tab_id, int):
            for i, (child, _text) in enumerate(self._tabs):
                if child is tab_id or str(child) == str(tab_id):
                    tab_id = i
                    break
        if not (0 <= tab_id < len(self._tabs)):
            return
        self._index = tab_id
        for i, (child, _text) in enumerate(self._tabs):
            if i == tab_id:
                child.pack(fill="both", expand=True)
            else:
                child.pack_forget()
        self._redraw()

    def tabs(self):
        return [str(child) for child, _text in self._tabs]

    def index(self, tab_id):
        for i, (child, _text) in enumerate(self._tabs):
            if child is tab_id or str(child) == str(tab_id):
                return i
        return -1

    def tab(self, tab_id, option=None, **kw):
        idx = tab_id if isinstance(tab_id, int) else self.index(tab_id)
        if not (0 <= idx < len(self._tabs)):
            return ""
        if option == "text":
            if kw:
                self._tabs[idx][1] = kw.get("text", self._tabs[idx][1])
                self._redraw()
            return self._tabs[idx][1]
        return self._tabs[idx][1]


# ---------------------------------------------------------------------------
# 核心转换函数（在后台线程中调用）
# ---------------------------------------------------------------------------
# 双引擎兼容：优先 Microsoft Office，未安装时自动回退 WPS（个人版免费）
WORD_PROG_IDS = ("Word.Application", "KWPS.Application")       # Word / WPS 文字
PPT_PROG_IDS = ("PowerPoint.Application", "KWPP.Application")  # PowerPoint / WPS 演示
ENGINE_NAME = {
    "Word.Application": "Microsoft Word",
    "KWPS.Application": "WPS 文字",
    "PowerPoint.Application": "Microsoft PowerPoint",
    "KWPP.Application": "WPS 演示",
}


def _dispatch_first(prog_ids):
    """按顺序尝试创建 COM 应用实例，返回 (app, ProgID)；全部失败返回 (None, 最后错误)。"""
    last_err = None
    for pid in prog_ids:
        try:
            return win32.DispatchEx(pid), pid
        except Exception as err:  # noqa: BLE001
            last_err = err
    return None, last_err


def convert_doc_to_pdf(word_file, out_dir, progress_cb=None, stop_flag=None):
    """将单个 Word 文档转为 PDF，返回 (状态, 信息)。状态: ok / skip / error"""
    filename = os.path.basename(word_file)
    if stop_flag is not None and stop_flag.is_set():
        return "skip", f"已取消：{filename}"

    if not os.path.exists(word_file):
        return "error", f"文件不存在：{filename}"

    base = os.path.splitext(filename)[0]
    pdf_path = os.path.join(out_dir, base + ".pdf")

    word = None
    doc = None
    try:
        # 每次独立启动实例（DispatchEx 进程隔离），避免单文件异常影响后续；
        # 优先 Microsoft Word，未安装时自动回退 WPS 文字
        word, engine = _dispatch_first(WORD_PROG_IDS)
        if word is None:
            return "error", (f"转换失败 {filename}：未检测到 Microsoft Word 或 WPS，"
                             f"无法转换 Word 文档（{engine}）")
        word.Visible = False
        try:
            word.DisplayAlerts = 0  # 关闭弹窗
        except Exception:
            pass

        doc = word.Documents.Open(
            os.path.abspath(word_file),
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
        )
        # SaveAs 第二个参数为 FileFormat，17 即 wdFormatPDF
        doc.SaveAs(os.path.abspath(pdf_path), WD_FORMAT_PDF)

        if progress_cb:
            progress_cb(filename)

        # 使用 WPS 转换时在结果中注明引擎，方便用户确认走了回退通道
        note = "" if engine == "Word.Application" else f"（{ENGINE_NAME[engine]}）"
        return "ok", f"已转换：{filename}{note}"

    except Exception as err:  # noqa: BLE001
        return "error", f"转换失败 {filename}：{err}"

    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass


# PowerPoint 导出 PDF 的常量（经实测验证）
PP_FIXED_FORMAT_PDF = 2     # FixedFormatType：ppFixedFormatTypePDF
PP_INTENT_PRINT = 2         # Intent：打印质量
PP_RANGE_ALL = 1            # RangeType：全部幻灯片


def _unique_out_path(out_dir, base_name, ext=".pdf"):
    """输出重名时自动加 _1、_2 后缀。"""
    path = os.path.join(out_dir, base_name + ext)
    if not os.path.exists(path):
        return path
    i = 1
    while True:
        cand = os.path.join(out_dir, f"{base_name}_{i}{ext}")
        if not os.path.exists(cand):
            return cand
        i += 1


def convert_ppt_to_pdf(ppt_file, out_dir, stop_flag=None):
    """将单个 PPT 文件导出为 PDF，返回 (状态, 信息)。状态: ok / skip / error"""
    filename = os.path.basename(ppt_file)
    if stop_flag is not None and stop_flag.is_set():
        return "skip", f"已取消：{filename}"

    if not os.path.exists(ppt_file):
        return "error", f"文件不存在：{filename}"
    if not filename.lower().endswith(PPT_EXTS):
        return "error", f"不支持的文件类型：{filename}"

    base = os.path.splitext(filename)[0]
    pdf_path = _unique_out_path(out_dir, base)

    app = None
    pres = None
    try:
        # DispatchEx 进程隔离 + DisplayAlerts=1（PowerPoint 的"关闭弹窗"值是 1，
        # 与 Word 的 0 不同；实测 0 会卡住）；
        # 优先 Microsoft PowerPoint，未安装时自动回退 WPS 演示
        app, engine = _dispatch_first(PPT_PROG_IDS)
        if app is None:
            return "error", (f"转换失败 {filename}：未检测到 Microsoft PowerPoint 或 WPS，"
                             f"无法转换 PPT 文件（{engine}）")
        try:
            app.DisplayAlerts = 1
        except Exception:
            pass

        # COM 只接受绝对路径；WithWindow=False 后台打开不闪窗口。
        # WPS 演示对命名参数支持不完整，失败时改用位置参数重试
        abs_src = os.path.abspath(ppt_file)
        try:
            pres = app.Presentations.Open(
                FileName=abs_src,
                ReadOnly=True, Untitled=False, WithWindow=False)
        except Exception:  # noqa: BLE001
            pres = app.Presentations.Open(abs_src, True, False, False)
        slide_count = pres.Slides.Count   # 必须 Close 前取

        out_abs = os.path.abspath(pdf_path)
        try:
            # 首选：SaveAs 32（ppSaveAsPDF），实测兼容性最好、矢量输出
            pres.SaveAs(out_abs, 32)
        except Exception:  # noqa: BLE001
            # 回退：ExportAsFixedFormat（部分环境动态绑定下会报 COM 转换错误）
            pres.ExportAsFixedFormat(
                out_abs, PP_FIXED_FORMAT_PDF, PP_INTENT_PRINT, 0, PP_RANGE_ALL)

        note = f"（{slide_count} 页"
        if engine != "PowerPoint.Application":
            note += f"·{ENGINE_NAME[engine]}"
        return "ok", f"已转换：{filename}{note}）"

    except Exception as err:  # noqa: BLE001
        return "error", f"转换失败 {filename}：{err}"

    finally:
        try:
            if pres is not None:
                pres.Close()
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# 主界面
# ---------------------------------------------------------------------------
class Word2PDFApp:
    # 仅深色模式，不提供主题切换
    DARK_THEME = "darkly"

    def __init__(self):
        if HAS_DND:
            self.root = tkdnd.TkinterDnD.Tk()
        else:
            self.root = tk.Tk()

        self.root.title(f"{APP_NAME} v{APP_VERSION}")

        # ---- 首屏定位：必须在窗口被映射（显示）之前一次算好 ----
        # 之前的写法先用 geometry() 设尺寸、update_idletasks() 逼 Tk 立刻建好
        # 窗口（此时默认出现在左上角并已可见），再第二次 geometry("+x+y")
        # 挪到屏幕中央 —— 用户就会看到「左上角闪一下再跳到中间」的两次跳转。
        # 现在改为：先 withdraw 隐藏窗口 → 计算尺寸与坐标 → 一次性设置
        # "宽x高+x+y" → deiconify 显示，全程只呈现一个最终位置。
        self.root.withdraw()
        gw, gh = self._compute_startup_geometry()
        wx, wy = self._compute_startup_position(gw, gh)
        self.root.geometry("%dx%d+%d+%d" % (gw, gh, wx, wy))
        self.root.minsize(s(780), min(s(600), gh))

        try:
            self.root.iconbitmap(_resource_path("app.ico"))
        except Exception:
            pass

        # Win11：标题栏染成与界面一致的深色（Win10 及以下自动跳过）
        _apply_titlebar_colors(self.root, bg=PALETTE["card"], text="#f0f0f0", border=PALETTE["card"])

        if HAS_TTB:
            self.style = tb.Style(theme=self.DARK_THEME)
            self.BTN_PRIMARY = "primary"
            self.BTN_DANGER = "danger"
            self.BTN_SECONDARY = "secondary"
            self.BTN_SUCCESS = "success"
        else:
            self.style = ttk.Style()
            try:
                self.style.theme_use("clam")
            except Exception:
                pass
            self.BTN_PRIMARY = None
            self.BTN_DANGER = None
            self.BTN_SECONDARY = None
            self.BTN_SUCCESS = None

        # 圆角外观：同步主题配色 + 注册列表样式
        refresh_palette(self.style)
        register_round_styles(self.style)
        # ttk 默认 Label / Frame 的底色同步为深蓝夜色（避免标题等露出灰底）
        try:
            self.style.configure("TLabel", background=PALETTE["bg"],
                                 foreground=PALETTE["fg"])
            self.style.configure("TFrame", background=PALETTE["bg"])
        except Exception:
            pass

        # 运行状态
        self.files = []          # [[完整路径, 状态 key], ...]
        self.out_dir = ""
        self._stop_flag = threading.Event()
        self._worker = None
        self._queue = queue.Queue()
        self._drop_hover = False

        # ---- 顶部标签页：多工具整合（圆角胶囊标签） ----
        Frame = _widget_class("Frame")
        Label = _widget_class("Label")

        # ---- 底部版权 / 反馈栏（先 pack，固定在窗口最底部，渐变背景）----
        self.footer = BgCanvas(self.root, height=s(34))
        self.footer.pack(side="bottom", fill="x")
        self.footer.bind("<Configure>", self._draw_footer, add="+")
        self._footer_item = None

        self.notebook = RoundNotebook(self.root, parent_bg=PALETTE["bg"])
        self.notebook.pack(fill="both", expand=True)
        self.word_tab = BgCanvas(self.notebook.body)
        self.img_tab = BgCanvas(self.notebook.body)
        self.pdf_tab = BgCanvas(self.notebook.body)
        self.merge_tab = BgCanvas(self.notebook.body)
        self.split_tab = BgCanvas(self.notebook.body)
        self.ppt_tab = BgCanvas(self.notebook.body)
        self.wm_tab = BgCanvas(self.notebook.body)
        self.pw_tab = BgCanvas(self.notebook.body)
        self.notebook.add(self.word_tab, text="Word 转 PDF")
        self.notebook.add(self.pw_tab, text="PDF 转 Word")
        self.notebook.add(self.ppt_tab, text="PPT 转 PDF")
        self.notebook.add(self.img_tab, text="PDF 转图片")
        self.notebook.add(self.wm_tab, text="PDF 加水印")
        self.notebook.add(self.merge_tab, text="PDF 合并")
        self.notebook.add(self.split_tab, text="PDF 拆分")
        self.notebook.add(self.pdf_tab, text="图片转 PDF")

        self._build_ui()          # 构建 Word 转 PDF 页
        self.img_page = PdfToImagesTab(self.img_tab, self)   # PDF 转图片页
        self.pdf_page = ImagesToPdfTab(self.pdf_tab, self)   # 图片转 PDF 页
        self.merge_page = MergePdfsTab(self.merge_tab, self)  # PDF 合并页
        self.split_page = SplitPdfTab(self.split_tab, self)  # PDF 拆分页
        self.ppt_page = PptToPdfTab(self.ppt_tab, self)      # PPT 转 PDF 页
        self.wm_page = WatermarkTab(self.wm_tab, self)       # PDF 加水印页
        self.pw_page = PdfToWordTab(self.pw_tab, self)       # PDF 转 Word 页
        # 页面切换时把拖拽落点重定向到当前页面（tkdnd 全局只能有一个活跃目标）
        self._page_by_tab = {
            str(self.word_tab): self,
            str(self.img_tab): self.img_page,
            str(self.pdf_tab): self.pdf_page,
            str(self.merge_tab): self.merge_page,
            str(self.split_tab): self.split_page,
            str(self.ppt_tab): self.ppt_page,
            str(self.wm_tab): self.wm_page,
            str(self.pw_tab): self.pw_page,
        }
        # RoundNotebook 是自绘控件（不是 ttk.Notebook），不会发出
        # <<NotebookTabChanged>>，因此直接包装它的 select() 来感知切页。
        self._orig_notebook_select = self.notebook.select

        def _select_and_sync(tab_id=0, _orig=self._orig_notebook_select):
            result = _orig(tab_id)
            # 页面已完成 pack，同步拖拽落点
            self._sync_active_drop()
            return result

        self.notebook.select = _select_and_sync

        self._apply_theme()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(100, self._poll_queue)
        self.root.after(220, self._sync_active_drop)   # 界面就绪后同步拖拽目标
        self._show_centered()                          # 一切就绪后再显示窗口

    # -- 启动定位 ------------------------------------------------------------
    def _compute_startup_geometry(self):
        """返回启动窗口尺寸 (宽, 高)，按屏幕自适应避免高分屏溢出。

        960x720 是按 100% 缩放设计的基准尺寸，在开启 DPI 感知后
        （1:1 渲染，不再被系统拉伸）需要按缩放比例换算，
        才能在高分屏上保持与以往一致的视觉大小。
        """
        gw, gh = s(960), s(720)
        try:
            sh = self.root.winfo_screenheight()
            gh = min(gh, int(sh * 0.85))
            gw = min(gw, int(self.root.winfo_screenwidth() * 0.95))
        except Exception:  # noqa: BLE001
            pass
        return gw, gh

    def _compute_startup_position(self, gw, gh):
        """返回窗口外框左上角坐标 (x, y)：在工作区内水平垂直居中。

        纯几何计算，不依赖窗口当前的实际位置，因此可在窗口显示前调用。
        """
        sw, sh = 1920, 1080
        try:
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
        except Exception:  # noqa: BLE001
            pass
        # Windows：取真实工作区（去掉任务栏），避免窗口压到任务栏下面
        try:
            import ctypes as _ctypes

            class _RECT(_ctypes.Structure):
                _fields_ = [("l", _ctypes.c_long), ("t", _ctypes.c_long),
                            ("r", _ctypes.c_long), ("b", _ctypes.c_long)]

            _rc = _RECT()
            _ctypes.windll.user32.SystemParametersInfoW(
                0x0030, 0, _ctypes.byref(_rc), 0)       # SPI_GETWORKAREA
            _w, _h = _rc.r - _rc.l, _rc.b - _rc.t
            if _w > 0 and _h > 0:
                sw, sh = _w, _h
        except Exception:  # noqa: BLE001
            pass
        wx = max((sw - gw) // 2, 0)
        wy = max((sh - gh) // 2, 0)
        return wx, wy

    def _show_centered(self):
        """显示窗口（位置已在 geometry("WxH+X+Y") 中定好，此处仅取消隐藏）。"""
        try:
            self.root.deiconify()
            self.root.lift()
        except Exception:  # noqa: BLE001
            pass

    # -- 拖拽落点同步 --------------------------------------------------------
    def _current_page(self):
        """返回当前可见功能页对象（失败时退回 Word 页）。"""
        idx = getattr(self.notebook, "_index", -1)
        tabs = getattr(self.notebook, "_tabs", [])
        if 0 <= idx < len(tabs):
            return self._page_by_tab.get(str(tabs[idx][0]), self)
        return self

    def _sync_active_drop(self):
        """把 tkdnd 活跃落点更新为当前可见页面。"""
        page = self._current_page()
        try:
            page._register_dnd()
        except Exception:  # noqa: BLE001
            pass
        return page

    def _register_dnd(self):
        """Word 转 PDF 页的拖拽落点（拖拽区 + 文件列表 + 空状态提示）。"""
        _register_page_drop(self, self.drop_canvas)
        if HAS_DND:
            for w in (self.tree, getattr(self, "empty_label", None)):
                if w is None:
                    continue
                try:
                    w.drop_target_register(tkdnd.DND_FILES)
                    w.dnd_bind("<<Drop>>", self._on_drop)
                except Exception:  # noqa: BLE001
                    pass

    def _unregister_dnd(self):
        if not (HAS_DND and tkdnd is not None):
            return
        for w in (getattr(self, "drop_canvas", None),
                  getattr(self, "tree", None),
                  getattr(self, "empty_label", None)):
            if w is None:
                continue
            try:
                w.drop_target_unregister()
            except Exception:  # noqa: BLE001
                pass

    # -- 组件快捷构造 --------------------------------------------------------
    def _button(self, master, text, style=None, command=None, **kw):
        """统一返回圆角自绘按钮。"""
        return RoundButton(master, text=text, bootstyle=style or "secondary",
                           command=command, parent_bg=PALETTE["bg"], **kw)

    def _draw_footer(self, _event=None):
        """在渐变背景上绘制上侧分隔线、底部版权文字与右下角缩放角标。"""
        f = self.footer
        f.delete("footer")
        f.redraw_bg()
        w = f.winfo_width() or 880
        h = f.winfo_height() or 34
        # 顶部分隔线：页脚与内容区的分界
        f.create_line(0, 0, w, 0, fill=PALETTE["border"], tags="footer")
        self._footer_item = f.create_text(
            w / 2, h / 2, text=copyright_text(), fill=PALETTE["muted"],
            font=("Microsoft YaHei UI", 9), tags="footer")

    def _make_status_bar(self, master, var):
        """状态栏：渐变背景 + 动态文字（替代纯色 Label）。"""
        bar = BgCanvas(master, height=s(24))
        holder = {"id": None}

        def draw(_e=None):
            bar.delete("st")
            bar.redraw_bg()
            holder["id"] = bar.create_text(
                20, 12, text=var.get(), anchor="w", fill=PALETTE["muted"],
                font=("Microsoft YaHei UI", 9), tags="st")

        bar.bind("<Configure>", draw, add="+")
        var.trace_add("write", lambda *_a: (
            bar.itemconfigure(holder["id"], text=var.get())
            if holder.get("id") else None))
        bar.pack(fill="x", padx=20, pady=(4, 0))
        return bar

    def _set_btn_style(self, btn, text, style):
        btn.configure(text=text, bootstyle=style)

    # -- 品牌 LOGO ------------------------------------------------------------
    def _logo_photo(self):
        """加载小尺寸 LOGO（tk 原生支持 PNG 透明），失败时返回 None。"""
        if not hasattr(self, "_logo_img"):
            try:
                self._logo_img = tk.PhotoImage(
                    file=_resource_path("logo_small.png"))
            except Exception:  # noqa: BLE001
                self._logo_img = None
        return self._logo_img

    def brand_row(self, master, title, subtitle):
        """标题行：大标题 + 灰色副标题（LOGO 已移至侧边栏顶部）。"""
        Label = _widget_class("Label")
        row = tk.Frame(master, bg=PALETTE["bg"])
        Label(row, text=title,
              font=("Microsoft YaHei UI", 17, "bold")).pack(side="left")
        row.pack(anchor="w")
        Label(master, text=subtitle, foreground=PALETTE["muted"],
              font=("Microsoft YaHei UI", 9)).pack(anchor="w", pady=(2, 0))

    # -- 主题（固定深色，无切换） -------------------------------------------
    def _apply_theme(self):
        c = getattr(self.style, "colors", None)
        fg = PALETTE["fg"] or getattr(c, "fg", "#dddddd")
        # 日志文本框（非 ttk 组件，需手动配色）
        self.log_text.configure(bg=PALETTE["card"], fg=fg,
                                insertbackground=fg)
        # 底部版权栏（年份实时取当前年份）
        if getattr(self, "footer", None) is not None:
            self._draw_footer()
        # 重画拖拽区（Canvas 非 ttk 组件）
        self._draw_drop_zone(self._drop_hover)
        # 其余功能页同样应用配色
        for page in (getattr(self, "img_page", None),
                     getattr(self, "pdf_page", None),
                     getattr(self, "merge_page", None),
                     getattr(self, "split_page", None),
                     getattr(self, "ppt_page", None),
                     getattr(self, "wm_page", None)):
            if page is not None:
                page._apply_theme()

    # -- UI 构建 ------------------------------------------------------------
    def _build_ui(self):
        Label = _widget_class("Label")
        Frame = _widget_class("Frame")
        Scrollbar = _widget_class("Scrollbar")
        Treeview = _widget_class("Treeview")

        PX = 20  # 统一水平留白
        BG = PALETTE["bg"]

        # ---- 顶部标题栏（LOGO + 标题） ----
        header = BgCanvas(self.word_tab)
        header.pack(fill="x", padx=PX, pady=(12, 6))
        self.brand_row(header, "Word 转 PDF",
                       "批量把 Word 文档转换为 PDF · 需本机安装 Microsoft Word 或 WPS")

        # ---- 工具栏（圆角按钮） ----
        toolbar = BgCanvas(self.word_tab)
        toolbar.pack(fill="x", padx=PX, pady=(4, 0))
        self._button(toolbar, "添加文件", self.BTN_PRIMARY,
                     self._add_files, pad_x=14).pack(side="left", padx=(0, 6))
        self._button(toolbar, "添加文件夹", self.BTN_SECONDARY,
                     self._add_folder, pad_x=14).pack(side="left", padx=6)
        self._button(toolbar, "移除选中", self.BTN_SECONDARY,
                     self._remove_selected, pad_x=14).pack(side="left", padx=6)
        self._button(toolbar, "清空列表", self.BTN_SECONDARY,
                     self._clear_list, pad_x=14).pack(side="left", padx=6)
        self.btn_up = self._button(toolbar, "上移", self.BTN_SECONDARY,
                                   lambda: self._move(-1), pad_x=14)
        self.btn_up.pack(side="left", padx=(14, 6))
        self.btn_down = self._button(toolbar, "下移", self.BTN_SECONDARY,
                                     lambda: self._move(1), pad_x=14)
        self.btn_down.pack(side="left", padx=6)

        # ---- 拖拽区（圆角虚线框） ----
        self.drop_canvas = tk.Canvas(self.word_tab, height=s(96), highlightthickness=0,
                                     bd=0, bg=BG, cursor="hand2")
        self.drop_canvas.pack(fill="x", padx=PX, pady=(10, 0))
        self.drop_canvas.bind("<Button-1>", lambda e: self._add_files())
        self.drop_canvas.bind("<Configure>",
                              lambda e: self._draw_drop_zone(self._drop_hover))
        if HAS_DND:
            self.drop_canvas.drop_target_register(tkdnd.DND_FILES)
            self.drop_canvas.dnd_bind("<<Drop>>", self._on_drop)
            self.drop_canvas.dnd_bind("<<DropEnter>>", self._on_drop_enter)
            self.drop_canvas.dnd_bind("<<DropLeave>>", self._on_drop_leave)
            _DND_ACTIVE["page"] = self       # 启动默认停在 Word 页

        # ---- 文件列表（圆角卡片） ----
        list_card = RoundCard(self.word_tab, padding=6, parent_bg=BG)
        list_card.pack(fill="both", expand=True, padx=PX, pady=(10, 0))
        list_frame = list_card.body

        cols = ("name", "path", "status")
        self.tree = Treeview(
            list_frame, columns=cols, show="headings", selectmode="extended",
            height=3)
        self.tree.heading("name", text="文件名")
        self.tree.heading("path", text="路径")
        self.tree.heading("status", text="状态")
        self.tree.column("name", width=230, anchor="w")
        self.tree.column("path", width=440, anchor="w")
        self.tree.column("status", width=90, anchor="center", stretch=False)

        vsb = Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # 空状态提示（叠加在列表中央）
        self.empty_label = tk.Label(
            list_frame,
            text="还没有添加文件\n点击「添加文件」或把文件/文件夹拖进上方区域",
            justify="center", bg=PALETTE["card"], fg=PALETTE["muted"],
            font=("Microsoft YaHei UI", 11))
        self.empty_label.place(relx=0.5, rely=0.5, anchor="center")

        # 拖拽落点：拖拽区 + 文件列表 + 空状态提示（提示浮在列表中央，
        # 若不注册为落点会吞掉落在提示文字上的拖拽事件）
        if HAS_DND:
            for _w in (self.tree, self.empty_label):
                try:
                    _w.drop_target_register(tkdnd.DND_FILES)
                    _w.dnd_bind("<<Drop>>", self._on_drop)
                except Exception:  # noqa: BLE001
                    pass

        self._refresh_tree_tags()

        # ---- 输出目录（圆角输入框） ----
        out_frame = BgCanvas(self.word_tab)
        out_frame.pack(fill="x", padx=PX, pady=(10, 0))
        Label(out_frame, text="保存到").pack(side="left")
        self.out_var = tk.StringVar()
        self.out_entry = RoundEntry(out_frame, textvariable=self.out_var,
                                    height=36, parent_bg=BG)
        self.out_entry.pack(side="left", fill="x", expand=True, padx=(10, 8))
        self._button(out_frame, "浏览…", self.BTN_SECONDARY,
                     self._choose_out_dir).pack(side="left")

        # ---- 进度条 + 转换按钮 ----
        prog_frame = BgCanvas(self.word_tab)
        prog_frame.pack(fill="x", padx=PX, pady=(10, 0))
        self.progress = RoundProgress(prog_frame, mode="determinate",
                                      height=14, parent_bg=BG)
        self.progress.pack(side="left", fill="x", expand=True, pady=8)
        self.btn_convert = self._button(
            prog_frame, "开始转换", self.BTN_SUCCESS, self._start_convert)
        self.btn_convert.pack(side="right", padx=(14, 0))

        self.status_var = tk.StringVar(value="就绪")
        self._make_status_bar(self.word_tab, self.status_var)

        # ---- 日志区（圆角卡片） ----
        log_card = RoundCard(self.word_tab, padding=8, height=104,
                             parent_bg=BG)
        log_card.pack(fill="x", padx=PX, pady=(8, 12))
        log_holder = log_card.body
        tk.Label(log_holder, text="转换日志", bg=PALETTE["card"],
                 fg=PALETTE["muted"],
                 font=("Microsoft YaHei UI", 9)).pack(anchor="w", pady=(0, 4))
        log_body = tk.Frame(log_holder, bg=PALETTE["card"])
        log_body.pack(fill="both", expand=True)
        self.log_text = tk.Text(
            log_body, height=4, state="disabled", wrap="word",
            font=("Consolas", 9), bd=0, relief="flat",
            highlightthickness=0, bg=PALETTE["card"], fg=PALETTE["fg"])
        log_sb = Scrollbar(log_body, orient="vertical",
                           command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_sb.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        log_sb.pack(side="right", fill="y")

    # -- 拖拽区绘制 ----------------------------------------------------------
    def _draw_drop_zone(self, hover=False):
        c = self.drop_canvas
        colors = self.style.colors
        c.configure(bg=PALETTE["bg"])
        c.delete("all")
        w = c.winfo_width() or 840
        h = c.winfo_height() or 96
        fill = _shade(PALETTE["card"], 0.05) if hover else PALETTE["card"]
        outline = PALETTE["info"] if hover else PALETTE["border"]
        round_rect(c, 4, 4, w - 4, h - 4, CARD_RADIUS, fill=fill,
                   outline=outline, width=2, dash=(8, 6))
        c.create_text(w // 2, h // 2 - 15,
                      text="⬇  将 Word 文件或文件夹拖到这里",
                      fill=PALETTE["fg"], font=("Microsoft YaHei UI", 12, "bold"))
        c.create_text(w // 2, h // 2 + 13,
                      text="支持 .doc / .docx / .docm / .rtf / .wps，或点击此区域选择文件",
                      fill=PALETTE["muted"], font=("Microsoft YaHei UI", 9))

    def _on_drop_enter(self, _event):
        self._drop_hover = True
        self._draw_drop_zone(True)

    def _on_drop_leave(self, _event):
        self._drop_hover = False
        self._draw_drop_zone(False)

    # -- 列表状态与着色 ------------------------------------------------------
    def _refresh_tree_tags(self):
        for key, color in STATUS_COLOR.items():
            if color:
                self.tree.tag_configure(key, foreground=color)

    def _refresh_list(self):
        self.tree.delete(*self.tree.get_children())
        for path, state in self.files:
            tags = (state,) if STATUS_COLOR.get(state) else ()
            self.tree.insert("", "end", values=(
                os.path.basename(path), path, STATUS_TEXT.get(state, state)),
                tags=tags)
        if self.files:
            self.empty_label.place_forget()
        else:
            self.empty_label.place(relx=0.5, rely=0.5, anchor="center")
        if self._worker and self._worker.is_alive():
            self.status_var.set(f"正在转换… {len(self.files)} 个文件")
        else:
            self.status_var.set(f"共 {len(self.files)} 个文件")

    # -- 文件操作 ------------------------------------------------------------
    def _move(self, delta):
        move_selected(self, delta, "文件")

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="选择要转换的 Word 文件",
            filetypes=[("Word 文档", "*.doc;*.docx;*.docm;*.rtf;*.wps"),
                       ("所有文件", "*.*")])
        if paths:
            self._append_paths(paths)

    def _add_folder(self):
        folder = filedialog.askdirectory(title="选择包含 Word 文件的文件夹")
        if not folder:
            return
        found = []
        for dirpath, _dirnames, filenames in os.walk(folder):
            for fn in filenames:
                if fn.lower().endswith(SUPPORTED_EXT):
                    found.append(os.path.join(dirpath, fn))
        if not found:
            messagebox.showinfo("提示", "该文件夹下未找到 Word 文档。")
            return
        self._append_paths(found)

    def _append_paths(self, paths):
        existing = {p for p, _ in self.files}
        added = 0
        for p in paths:
            p = os.path.normpath(p)
            if p in existing:
                continue
            self.files.append([p, "pending"])
            existing.add(p)
            added += 1
        self._refresh_list()
        self._log(f"已添加 {added} 个文件")

    def _remove_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        sel_paths = {self.tree.item(i, "values")[1] for i in sel}
        self.files = [f for f in self.files if f[0] not in sel_paths]
        self._refresh_list()

    def _clear_list(self):
        if self._worker and self._worker.is_alive():
            messagebox.showwarning("提示", "正在转换中，无法清空列表。")
            return
        self.files.clear()
        self._refresh_list()
        self.progress["value"] = 0

    def _on_drop(self, event):
        items = self.root.tk.splitlist(event.data)
        paths = []
        for it in items:
            if os.path.isdir(it):
                for dp, _dn, fns in os.walk(it):
                    for fn in fns:
                        if fn.lower().endswith(SUPPORTED_EXT):
                            paths.append(os.path.join(dp, fn))
            elif it.lower().endswith(SUPPORTED_EXT):
                paths.append(it)
        self._drop_hover = False
        self._draw_drop_zone(False)
        if paths:
            self._append_paths(paths)
        else:
            messagebox.showinfo("提示", "未识别到可转换的 Word 文档。")

    def _choose_out_dir(self):
        d = filedialog.askdirectory(title="选择 PDF 保存位置")
        if d:
            self.out_var.set(d)

    # -- 转换流程 ------------------------------------------------------------
    def _start_convert(self):
        if self._worker and self._worker.is_alive():
            self._stop_flag.set()
            self.status_var.set("正在停止…")
            return

        if not HAS_WIN32:
            messagebox.showerror(
                "错误", "缺少 pywin32 组件，无法调用 Word 转换。\n请安装：pip install pywin32")
            return

        if not self.files:
            messagebox.showinfo("提示", "请先添加要转换的 Word 文件。")
            return

        out_dir = self.out_var.get().strip()
        if not out_dir:
            messagebox.showinfo("提示", "请选择 PDF 保存位置。")
            return
        if not os.path.isdir(out_dir):
            messagebox.showwarning("提示", "保存目录不存在，请重新选择。")
            return

        self._stop_flag.clear()
        self.progress["maximum"] = len(self.files)
        self.progress["value"] = 0
        for i in range(len(self.files)):
            self.files[i][1] = "pending"
        self._refresh_list()

        self._log("开始批量转换…")
        self._set_btn_style(self.btn_convert, "停止", self.BTN_DANGER)

        self._worker = threading.Thread(
            target=self._convert_worker, args=(list(self.files), out_dir),
            daemon=True)
        self._worker.start()

    def _convert_worker(self, files_snapshot, out_dir):
        pythoncom.CoInitialize()
        ok_count = 0
        err_count = 0
        skip_count = 0
        try:
            for idx, (path, _status) in enumerate(files_snapshot):
                if self._stop_flag.is_set():
                    self._queue.put(("stopped", None))
                    break
                self._queue.put(("status", (idx, "working")))
                state, info = convert_doc_to_pdf(
                    path, out_dir, stop_flag=self._stop_flag)
                if state == "ok":
                    ok_count += 1
                elif state == "error":
                    err_count += 1
                else:
                    skip_count += 1
                self._queue.put(("item", (idx, state, info, path)))
            else:
                self._queue.put(("done", (ok_count, err_count, skip_count)))
        except Exception:  # noqa: BLE001
            self._queue.put(("fatal", traceback.format_exc()))
        finally:
            pythoncom.CoUninitialize()

    # -- 结果轮询 ------------------------------------------------------------
    def _poll_queue(self):
        try:
            while True:
                msg = self._queue.get_nowait()
                self._handle_msg(msg)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _handle_msg(self, msg):
        kind, payload = msg[0], msg[1]
        children = self.tree.get_children()
        if kind == "status":
            idx, state = payload
            if idx < len(children):
                self.files[idx][1] = state
                self.tree.item(children[idx], values=(
                    os.path.basename(self.files[idx][0]),
                    self.files[idx][0], STATUS_TEXT.get(state, state)),
                    tags=(state,) if STATUS_COLOR.get(state) else ())
        elif kind == "item":
            idx, state, info, path = payload
            if idx < len(children):
                self.files[idx][1] = state
                self.tree.item(children[idx], values=(
                    os.path.basename(path), path, STATUS_TEXT.get(state, state)),
                    tags=(state,) if STATUS_COLOR.get(state) else ())
            self.progress["value"] = idx + 1
            self._log(info)
        elif kind == "done":
            ok, err, skip = payload
            self.progress["value"] = self.progress["maximum"]
            self.status_var.set(f"完成：成功 {ok} · 失败 {err} · 跳过 {skip}")
            self._log(f"转换结束：成功 {ok} 个，失败 {err} 个，跳过 {skip} 个")
            self._set_btn_style(self.btn_convert, "开始转换", self.BTN_SUCCESS)
            self._worker = None
        elif kind == "stopped":
            self.status_var.set("已停止")
            self._log("转换已手动停止")
            self._set_btn_style(self.btn_convert, "开始转换", self.BTN_SUCCESS)
            self._worker = None
        elif kind == "fatal":
            self._log("发生严重错误：\n" + payload)
            self._set_btn_style(self.btn_convert, "开始转换", self.BTN_SUCCESS)
            self._worker = None

    # -- 日志 / 关闭 ---------------------------------------------------------
    def _log(self, text):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _on_close(self):
        if self._worker and self._worker.is_alive():
            if not messagebox.askyesno("退出", "正在转换中，确定要退出吗？"):
                return
            self._stop_flag.set()
        self.root.destroy()

    def run(self):
        self.root.mainloop()


def convert_pdf_to_images(pdf_file, out_dir, zoom, fmt, stop_flag=None):
    """将单个 PDF 的每一页渲染成图片。返回 (状态, 信息)。"""
    import fitz  # PyMuPDF（打包时已随依赖一起打入）

    filename = os.path.basename(pdf_file)
    if stop_flag is not None and stop_flag.is_set():
        return "skip", f"已取消：{filename}"
    if not os.path.exists(pdf_file):
        return "error", f"文件不存在：{filename}"

    try:
        doc = fitz.open(os.path.abspath(pdf_file))
    except Exception as err:  # noqa: BLE001
        return "error", f"无法打开 {filename}：{err}"

    base = os.path.splitext(filename)[0]
    ext = ".jpg" if fmt == "jpg" else ".png"
    count = 0
    try:
        page_total = doc.page_count
        for pg in range(page_total):
            if stop_flag is not None and stop_flag.is_set():
                return "skip", f"已取消：{filename}"
            page = doc[pg]
            # 放大矩阵：zoom = 目标dpi / 72。无任何旋转，按原始方向导出
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img_name = f"{base}_第{pg + 1:03d}页{ext}"
            pix.save(os.path.join(out_dir, img_name))
            count += 1
    finally:
        doc.close()

    return "ok", f"已转换：{filename}（共 {count} 页）"


def convert_pdf_to_word(pdf_file, out_dir, stop_flag=None):
    """将单个 PDF 转为 Word(.docx)，保留文字、图片与基本段落/表格结构。

    依赖 pdf2docx（纯 Python，基于 PyMuPDF），不要求本机安装 Office/WPS。
    返回 (状态, 信息)。状态: ok / skip / error
    """
    filename = os.path.basename(pdf_file)
    if stop_flag is not None and stop_flag.is_set():
        return "skip", f"已取消：{filename}"
    if not os.path.exists(pdf_file):
        return "error", f"文件不存在：{filename}"

    base = os.path.splitext(filename)[0]
    docx_path = _unique_out_path(out_dir, base, ".docx")

    src_doc = None
    try:
        import fitz  # 先用 PyMuPDF 做一次体检，能给出更友好的错误提示
        src_doc = fitz.open(os.path.abspath(pdf_file))
        page_count = src_doc.page_count
        if page_count == 0:
            return "error", f"转换失败 {filename}：PDF 不含任何页面"
    except Exception as err:  # noqa: BLE001
        return "error", f"无法打开 {filename}：{err}"
    finally:
        try:
            if src_doc is not None:
                src_doc.close()
        except Exception:  # noqa: BLE001
            pass

    try:
        from pdf2docx import Converter
    except Exception as err:  # noqa: BLE001
        return "error", f"转换失败 {filename}：缺少 PDF 转 Word 组件（{err}）"

    cv = None
    try:
        cv = Converter(os.path.abspath(pdf_file))
        cv.convert(os.path.abspath(docx_path))
        if stop_flag is not None and stop_flag.is_set():
            return "skip", f"已取消：{filename}"
    except Exception as err:  # noqa: BLE001
        return "error", f"转换失败 {filename}：{err}"
    finally:
        try:
            if cv is not None:
                cv.close()
        except Exception:  # noqa: BLE001
            pass

    return "ok", f"已转换：{filename}（共 {page_count} 页）"


# A4 纸张尺寸（单位：pt，1pt = 1/72 英寸）
A4_W, A4_H = 595.276, 841.89
MM_PT = 72.0 / 25.4          # 1 毫米 = 多少 pt


def images_to_pdf(img_paths, out_pdf, page_mode="auto", margin_mm=10,
                  stop_flag=None, progress_cb=None):
    """
    把多张图片按列表顺序合并成一个 PDF。

    page_mode:
      - "auto"  : 页面尺寸跟随图片本身（等同一张图一页，不做缩放）
      - "a4"    : A4 纵向，图片等比缩放后居中
      - "a4-l"  : A4 横向，图片等比缩放后居中
    margin_mm: 仅 A4 模式生效的页边距（毫米）

    返回 (状态, 信息)。状态: ok / skip / error
    """
    import fitz  # PyMuPDF（打包时已随依赖一起打入）

    if stop_flag is not None and stop_flag.is_set():
        return "skip", "已取消"

    out_pdf = os.path.abspath(out_pdf)
    parent = os.path.dirname(out_pdf)
    if parent and not os.path.isdir(parent):
        os.makedirs(parent, exist_ok=True)

    doc = fitz.open()
    failed = []
    total = len(img_paths)
    try:
        for i, path in enumerate(img_paths):
            if stop_flag is not None and stop_flag.is_set():
                doc.close()
                return "skip", "已取消"

            name = os.path.basename(path)
            try:
                if page_mode == "auto":
                    # 页面尺寸跟随图片（PyMuPDF 会自动处理 EXIF 方向）
                    src = fitz.open(path)
                    try:
                        buf = src.convert_to_pdf()
                    finally:
                        src.close()
                    tmp = fitz.open("pdf", buf)
                    try:
                        doc.insert_pdf(tmp)
                    finally:
                        tmp.close()
                else:
                    land = (page_mode == "a4-l")
                    pw, ph = (A4_H, A4_W) if land else (A4_W, A4_H)
                    src = fitz.open(path)
                    try:
                        iw, ih = src[0].rect.width, src[0].rect.height
                    finally:
                        src.close()
                    if iw <= 0 or ih <= 0:
                        raise ValueError("图片尺寸无效")
                    page = doc.new_page(width=pw, height=ph)
                    m = max(0.0, float(margin_mm)) * MM_PT
                    aw, ah = pw - 2 * m, ph - 2 * m
                    k = min(aw / iw, ah / ih)
                    w, h = iw * k, ih * k
                    x0, y0 = (pw - w) / 2, (ph - h) / 2
                    rect = fitz.Rect(x0, y0, x0 + w, y0 + h)
                    page.insert_image(rect, filename=path, keep_proportion=True)
            except Exception as err:  # noqa: BLE001
                failed.append(f"{name}（{err}）")

            if progress_cb is not None:
                try:
                    progress_cb(i + 1, total)
                except Exception:
                    pass

        pages = doc.page_count
        if pages == 0:
            doc.close()
            detail = "；".join(failed[:3]) if failed else "没有可用的图片"
            return "error", f"未能生成 PDF：{detail}"

        try:
            doc.save(out_pdf, garbage=3, deflate=True)
        except Exception as err:  # noqa: BLE001
            doc.close()
            return "error", f"保存失败（目标文件可能正被打开）：{err}"
        doc.close()
    except Exception as err:  # noqa: BLE001
        try:
            doc.close()
        except Exception:
            pass
        return "error", f"合并失败：{err}"

    size_mb = os.path.getsize(out_pdf) / 1024 / 1024
    info = f"已生成 {os.path.basename(out_pdf)}（{pages} 页，{size_mb:.1f} MB）"
    if failed:
        info += f"｜{len(failed)} 张图片读取失败：" + "，".join(failed[:3])
    return "ok", info


def merge_pdfs(pdf_paths, out_pdf, stop_flag=None, progress_cb=None):
    """
    按列表顺序把多个 PDF 文件合并成一个 PDF。

    逐个文件容错：加密/损坏的文件跳过，其余照常合并。
    返回 (状态, 信息)。状态: ok / skip / error
    """
    import fitz  # PyMuPDF（打包时已随依赖一起打入）

    if stop_flag is not None and stop_flag.is_set():
        return "skip", "已取消"

    out_pdf = os.path.abspath(out_pdf)
    parent = os.path.dirname(out_pdf)
    if parent and not os.path.isdir(parent):
        os.makedirs(parent, exist_ok=True)

    doc = fitz.open()
    failed = []
    total = len(pdf_paths)
    total_pages = 0
    try:
        for i, path in enumerate(pdf_paths):
            if stop_flag is not None and stop_flag.is_set():
                doc.close()
                return "skip", "已取消"

            name = os.path.basename(path)
            try:
                if not os.path.isfile(path):
                    raise ValueError("文件不存在")
                src = fitz.open(path)
                try:
                    if src.needs_pass:
                        raise ValueError("文件已加密，无法读取")
                    cnt = src.page_count
                    if cnt <= 0:
                        raise ValueError("空文件")
                    doc.insert_pdf(src)   # 保真合并：矢量/文字原样并入
                    total_pages += cnt
                finally:
                    src.close()
            except Exception as err:  # noqa: BLE001
                failed.append(f"{name}（{err}）")

            if progress_cb is not None:
                try:
                    progress_cb(i + 1, total)
                except Exception:
                    pass

        pages = doc.page_count
        if pages == 0:
            doc.close()
            detail = "；".join(failed[:3]) if failed else "没有可合并的 PDF"
            return "error", f"未能生成 PDF：{detail}"

        try:
            doc.save(out_pdf, garbage=3, deflate=True)
        except Exception as err:  # noqa: BLE001
            doc.close()
            return "error", f"保存失败（目标文件可能正被打开）：{err}"
        doc.close()
    except Exception as err:  # noqa: BLE001
        try:
            doc.close()
        except Exception:
            pass
        return "error", f"合并失败：{err}"

    size_mb = os.path.getsize(out_pdf) / 1024 / 1024
    info = (f"已生成 {os.path.basename(out_pdf)}"
            f"（{len(pdf_paths) - len(failed)} 个文件，共 {total_pages} 页，"
            f"{size_mb:.1f} MB）")
    if failed:
        info += f"｜{len(failed)} 个文件读取失败：" + "，".join(failed[:3])
    return "ok", info


def parse_page_ranges(text):
    """解析 "1-3, 5, 8-10" → [(1,3),(5,5),(8,10)]（1 起、闭区间、自动排序）。

    空文本返回 None（表示未填写）；含非法片段返回 []。
    """
    if text is None or not str(text).strip():
        return None
    groups = []
    for part in str(text).replace("，", ",").split(","):
        part = part.strip()
        if not part:
            continue
        m = re.match(r"^(\d+)\s*[-—~]\s*(\d+)$", part)
        if m:
            s, e = int(m.group(1)), int(m.group(2))
            if s > e:
                s, e = e, s
            groups.append((s, e))
        elif part.isdigit():
            n = int(part)
            groups.append((n, n))
        else:
            return []
    groups.sort()
    return groups


def split_pdf(pdf_file, out_dir, ranges=None, stop_flag=None, progress_cb=None):
    """拆分 PDF，返回 (状态, 信息)。状态: ok / skip / error

    ranges=None：每页一个 PDF；
    ranges=[(s,e), ...]：每段范围各生成一个 PDF（页码 1 起，闭区间）。
    """
    import fitz

    filename = os.path.basename(pdf_file)
    if stop_flag is not None and stop_flag.is_set():
        return "skip", "已取消"
    if not os.path.exists(pdf_file):
        return "error", f"文件不存在：{filename}"
    base = os.path.splitext(filename)[0]
    if out_dir and not os.path.isdir(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    try:
        src = fitz.open(pdf_file)
    except Exception as err:  # noqa: BLE001
        return "error", f"无法打开 {filename}：{err}"

    try:
        if src.needs_pass:
            return "error", f"文件已加密，无法拆分：{filename}"
        total = src.page_count
        if total <= 0:
            return "error", f"空文件：{filename}"
        if ranges is None:
            groups = [(i + 1, i + 1) for i in range(total)]
        else:
            groups = [(s, e) for s, e in ranges if 1 <= s <= e <= total]
            if not groups:
                return "error", f"页码范围无效（文档共 {total} 页）"

        def out_path(tag):
            path = os.path.join(out_dir, f"{base}_{tag}.pdf")
            i = 1
            while os.path.exists(path):
                path = os.path.join(out_dir, f"{base}_{tag}_{i}.pdf")
                i += 1
            return path

        done = 0
        try:
            for s, e in groups:
                if stop_flag is not None and stop_flag.is_set():
                    return "skip", f"已停止：已完成 {done}/{len(groups)} 个文件"
                tag = f"第{s}页" if s == e else f"{s}-{e}页"
                dst = fitz.open()
                try:
                    dst.insert_pdf(src, from_page=s - 1, to_page=e - 1)
                    dst.save(out_path(tag), garbage=3, deflate=True)
                finally:
                    dst.close()
                done += 1
                if progress_cb is not None:
                    try:
                        progress_cb(done, len(groups))
                    except Exception:
                        pass
        except Exception as err:  # noqa: BLE001
            return "error", f"拆分失败：{err}"
        return "ok", f"已拆分：{filename} → {len(groups)} 个 PDF（输出到 {out_dir}）"
    finally:
        try:
            src.close()
        except Exception:
            pass


def _pixmap_with_opacity(image_path, opacity):
    """加载图片为带 alpha 的 Pixmap 并整体乘以 opacity（0~1）。失败返回 None。"""
    import fitz

    try:
        pix = fitz.Pixmap(image_path)
        if pix.colorspace is not None and pix.colorspace.name != fitz.csRGB.name:
            pix = fitz.Pixmap(fitz.csRGB, pix)      # 统一转 RGB
        if pix.alpha == 0:
            pix = fitz.Pixmap(pix, 1)               # 加 alpha，RGB 自动拷贝
        samples = bytearray(pix.samples)
        n = pix.n
        for row in range(pix.height):
            base_i = row * pix.stride
            for col in range(pix.width):
                idx = base_i + col * n + (n - 1)
                samples[idx] = int(samples[idx] * opacity)
        return fitz.Pixmap(pix.colorspace, pix.width, pix.height,
                           bytes(samples), 1)
    except Exception:  # noqa: BLE001
        return None


# 页面切换时把「当前可见页的拖拽注册」提升到前台。
# tkdnd 在 Windows 上会记住最后一次注册窗口，较早注册的页面（如 Word 转 PDF）
# 会抢占拖拽事件，导致后建的页面收不到 Drop。因此每次切换到某页时重新注册一次，
# 让 tkdnd 的「最后注册者」始终等于当前页面，同时取消其它页面的注册以避免串页。
_DND_ACTIVE = {"page": None}


def _register_page_drop(page, widget):
    """把某功能页（及其列表控件）注册为拖拽落点，并使其成为 tkdnd 的当前目标。

    仅在 tkinterdnd2 可用且窗口已创建时生效；任何异常都静默忽略，不影响主流程。
    """
    if not (HAS_DND and tkdnd is not None) or widget is None:
        return
    try:
        if not widget.winfo_exists():
            return
    except Exception:  # noqa: BLE001
        return
    # 取消上一页的注册，避免拖拽落在隐藏页面上
    prev = _DND_ACTIVE.get("page")
    if prev is not None and prev is not page:
        try:
            prev._unregister_dnd()
        except Exception:  # noqa: BLE001
            pass
    try:
        widget.drop_target_register(tkdnd.DND_FILES)
    except Exception:  # noqa: BLE001
        pass
    _DND_ACTIVE["page"] = page


def move_selected(page, delta, noun="文件"):
    """批量列表通用「上移 / 下移」。

    各批量页的数据结构一致（self.files = [[路径, 状态], ...] + self.tree +
    self._worker + self._refresh_list()），故统一在此实现：
    delta=-1 上移，delta=+1 下移；支持多选，按顺序整体平移一格。
    """
    if not getattr(page, "files", None):
        return
    if getattr(page, "_worker", None) and page._worker.is_alive():
        return
    tree = page.tree
    sel = tree.selection()
    if not sel:
        messagebox.showinfo("提示", f"请先在列表里选中要调整顺序的{noun}。")
        return
    idxs = sorted((tree.index(i) for i in sel), reverse=(delta > 0))
    if delta < 0 and idxs[0] == 0:
        return
    if delta > 0 and idxs[-1] >= len(page.files) - 1:
        return
    for i in idxs:
        j = i + delta
        page.files[i], page.files[j] = page.files[j], page.files[i]
    page._refresh_list()
    children = tree.get_children()
    for i in idxs:
        tree.selection_add(children[i + delta])


def add_watermark(pdf_file, out_pdf, wm_type="text", text="内部资料",
                  image=None, tile=True, opacity=0.15, fontsize=48,
                  color=(1, 0, 0), angle=45, stop_flag=None, progress_cb=None):
    """给 PDF 每一页加水印，输出到 out_pdf。返回 (状态, 信息)。

    wm_type: "text"（文字水印）/ "image"（图片水印）
    tile: True 平铺整页 / False 页面居中单个
    """
    import fitz

    filename = os.path.basename(pdf_file)
    if stop_flag is not None and stop_flag.is_set():
        return "skip", "已取消"
    if not os.path.exists(pdf_file):
        return "error", f"文件不存在：{filename}"
    if wm_type == "image":
        if not image or not os.path.exists(image):
            return "error", "未选择水印图片或文件不存在"
    elif not (text or "").strip():
        return "error", "水印文字不能为空"

    try:
        doc = fitz.open(pdf_file)
    except Exception as err:  # noqa: BLE001
        return "error", f"无法打开 {filename}：{err}"

    try:
        if doc.needs_pass:
            return "error", f"文件已加密，无法加水印：{filename}"
        total = doc.page_count
        if total <= 0:
            return "error", f"空文件：{filename}"

        # 预生成图片水印（含透明度），只处理一次
        wm_pix = None
        if wm_type == "image":
            wm_pix = _pixmap_with_opacity(image, min(max(opacity, 0.05), 1.0))
            if wm_pix is None:
                return "error", f"水印图片无法读取：{os.path.basename(image)}"
            wm_png = wm_pix.tobytes("png")
            aspect = wm_pix.height / max(wm_pix.width, 1)

        rad = fitz.Matrix(angle % 360)

        for i, page in enumerate(doc):
            if stop_flag is not None and stop_flag.is_set():
                return "skip", f"已停止：已完成 {i}/{total} 页"
            rect = page.rect
            if wm_type == "text":
                if tile:
                    step_y = max(fontsize * 3.2, 90)
                    step_x = max(fontsize * 4.0, 110)
                    y = 40 + fontsize
                    while y < rect.height + fontsize:
                        x = 10
                        while x < rect.width:
                            page.insert_text(
                                (x, y), text, fontname="china-s",
                                fontsize=fontsize, color=color,
                                fill_opacity=opacity,
                                morph=(fitz.Point(x, y), rad))
                            x += step_x
                        y += step_y
                else:
                    size = max(rect.width, rect.height) * 0.14
                    page.insert_text(
                        (rect.width / 2, rect.height / 2), text,
                        fontname="china-s", fontsize=size, color=color,
                        fill_opacity=opacity,
                        morph=(fitz.Point(rect.width / 2, rect.height / 2), rad))
            else:
                # 图片水印：目标宽 = 页宽 30%（平铺）/ 45%（居中），保持宽高比
                target_w = rect.width * (0.30 if tile else 0.45)
                target_h = target_w * aspect
                if tile:
                    step_y = target_h * 2.6
                    step_x = target_w * 1.9
                    y = 20
                    while y < rect.height:
                        x = 10
                        while x < rect.width:
                            page.insert_image(
                                fitz.Rect(x, y, min(x + target_w, rect.width),
                                          min(y + target_h, rect.height)),
                                stream=wm_png, overlay=True)
                            x += step_x
                        y += step_y
                else:
                    cx, cy = rect.width / 2, rect.height / 2
                    page.insert_image(
                        fitz.Rect(cx - target_w / 2, cy - target_h / 2,
                                  cx + target_w / 2, cy + target_h / 2),
                        stream=wm_png, overlay=True)
            if progress_cb is not None:
                try:
                    progress_cb(i + 1, total)
                except Exception:
                    pass

        try:
            doc.save(out_pdf, garbage=3, deflate=True)
        except Exception as err:  # noqa: BLE001
            return "error", f"保存失败：{err}"
        label = (f"图片水印（{'平铺' if tile else '居中'}）" if wm_type == "image"
                 else f"「{text}」（{'平铺' if tile else '居中'}）")
        return "ok", f"已加水印：{filename} → {os.path.basename(out_pdf)}｜{label}"
    finally:
        try:
            doc.close()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# PDF 转图片页（标签页之二）
# ---------------------------------------------------------------------------
class PdfToImagesTab:
    """PDF 转图片功能页，界面与 Word 转 PDF 页保持一致。"""

    def __init__(self, parent, app):
        self.parent = parent
        self.app = app            # 复用主程序的样式/颜色/root
        self.root = app.root
        self.files = []           # [[完整路径, 状态 key], ...]
        self.out_var = tk.StringVar()
        self.fmt_var = tk.StringVar(value="png")
        self.zoom_var = tk.StringVar(value="2")
        self._stop_flag = threading.Event()
        self._worker = None
        self._queue = queue.Queue()
        self._build()
        self.root.after(100, self._poll_queue)

    # -- 组件快捷构造（与主程序同风格） -------------------------------------
    def _btn(self, master, text, style, command=None, **kw):
        return RoundButton(master, text=text, bootstyle=style or "secondary",
                           command=command,
                           parent_bg=kw.pop("parent_bg", PALETTE["bg"]), **kw)

    def _apply_theme(self):
        fg = PALETTE["fg"]
        self.log_text.configure(bg=PALETTE["card"], fg=fg, insertbackground=fg)

    # -- UI -----------------------------------------------------------------
    def _build(self):
        Label = _widget_class("Label")
        Frame = _widget_class("Frame")
        Treeview = _widget_class("Treeview")
        Scrollbar = _widget_class("Scrollbar")

        PX = 20
        BG = PALETTE["bg"]
        B = self.app  # 复用主程序按钮样式常量

        # 标题
        header = BgCanvas(self.parent)
        header.pack(fill="x", padx=PX, pady=(12, 6))
        self.app.brand_row(header, "PDF 转图片",
                           "把 PDF 每一页导出为 PNG / JPG 图片 · 本地处理，不上传")

        # 工具栏（圆角按钮）
        toolbar = BgCanvas(self.parent)
        toolbar.pack(fill="x", padx=PX, pady=(4, 0))
        self._btn(toolbar, "添加文件", B.BTN_PRIMARY,
                  self._add_files, pad_x=14).pack(side="left", padx=(0, 6))
        self._btn(toolbar, "添加文件夹", B.BTN_SECONDARY,
                  self._add_folder, pad_x=14).pack(side="left", padx=6)
        self._btn(toolbar, "移除选中", B.BTN_SECONDARY,
                  self._remove_selected, pad_x=14).pack(side="left", padx=6)
        self._btn(toolbar, "清空列表", B.BTN_SECONDARY,
                  self._clear_list, pad_x=14).pack(side="left", padx=6)
        self.btn_up = self._btn(toolbar, "上移", B.BTN_SECONDARY,
                                lambda: self._move(-1), pad_x=14)
        self.btn_up.pack(side="left", padx=(14, 6))
        self.btn_down = self._btn(toolbar, "下移", B.BTN_SECONDARY,
                                  lambda: self._move(1), pad_x=14)
        self.btn_down.pack(side="left", padx=6)

        # 文件列表（圆角卡片）
        list_card = RoundCard(self.parent, padding=6, parent_bg=BG)
        list_card.pack(fill="both", expand=True, padx=PX, pady=(10, 0))
        list_frame = list_card.body
        cols = ("name", "path", "status")
        self.tree = Treeview(list_frame, columns=cols, show="headings",
                             selectmode="extended", height=3)
        self.tree.heading("name", text="文件名")
        self.tree.heading("path", text="路径")
        self.tree.heading("status", text="状态")
        self.tree.column("name", width=230, anchor="w")
        self.tree.column("path", width=440, anchor="w")
        self.tree.column("status", width=90, anchor="center", stretch=False)
        vsb = Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self._refresh_tags()

        # 拖拽添加（直接拖到列表上）
        if HAS_DND:
            try:
                self.tree.drop_target_register(tkdnd.DND_FILES)
                self.tree.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass

        # 空状态提示（叠加在列表中央）
        self.empty_label = tk.Label(
            list_frame,
            text="还没有添加文件\n点击「添加文件」或把 PDF 文件拖到列表上",
            justify="center", bg=PALETTE["card"], fg=PALETTE["muted"],
            font=("Microsoft YaHei UI", 11))
        self.empty_label.place(relx=0.5, rely=0.5, anchor="center")
        if HAS_DND:
            # 空状态提示浮在列表中央，必须同样注册为落点，
            # 否则拖到提示文字上的文件会被这个标签吞掉、无法触发 Drop。
            try:
                self.empty_label.drop_target_register(tkdnd.DND_FILES)
                self.empty_label.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:  # noqa: BLE001
                pass

        # 输出目录（圆角输入框）
        out_frame = BgCanvas(self.parent)
        out_frame.pack(fill="x", padx=PX, pady=(10, 0))
        Label(out_frame, text="保存到").pack(side="left")
        self.out_entry = RoundEntry(out_frame, textvariable=self.out_var,
                                    height=36, parent_bg=BG)
        self.out_entry.pack(side="left", fill="x", expand=True, padx=(10, 8))
        self._btn(out_frame, "浏览…", B.BTN_SECONDARY,
                  self._choose_out_dir).pack(side="left")

        # 选项：清晰度 + 输出格式（圆角下拉框）
        opt_frame = BgCanvas(self.parent)
        opt_frame.pack(pady=(10, 0))  # 水平居中
        Label(opt_frame, text="清晰度").pack(side="left")
        self.dpi_combo = RoundCombo(
            opt_frame, width=14, parent_bg=BG,
            values=("标准 144 dpi", "高清 216 dpi（推荐）", "超清 288 dpi"))
        self.dpi_combo.current(1)
        self.dpi_combo.pack(side="left", padx=(10, 18))

        Label(opt_frame, text="输出格式").pack(side="left")
        self.fmt_combo = RoundCombo(
            opt_frame, width=10, parent_bg=BG,
            values=("PNG（无损，推荐）", "JPG（体积小）"))
        self.fmt_combo.current(0)
        self.fmt_combo.pack(side="left", padx=(10, 0))

        # 进度 + 转换按钮
        prog_frame = BgCanvas(self.parent)
        prog_frame.pack(fill="x", padx=PX, pady=(12, 0))
        self.progress = RoundProgress(prog_frame, mode="determinate",
                                      height=14, parent_bg=BG)
        self.progress.pack(side="left", fill="x", expand=True, pady=8)
        self.btn_convert = self._btn(
            prog_frame, "开始转换", B.BTN_SUCCESS, self._start_convert)
        self.btn_convert.pack(side="right", padx=(14, 0))

        self.status_var = tk.StringVar(value="就绪")
        self.app._make_status_bar(self.parent, self.status_var)

        # 日志（圆角卡片）
        log_card = RoundCard(self.parent, padding=8, height=104, parent_bg=BG)
        log_card.pack(fill="x", padx=PX, pady=(8, 12))
        log_holder = log_card.body
        tk.Label(log_holder, text="转换日志", bg=PALETTE["card"],
                 fg=PALETTE["muted"],
                 font=("Microsoft YaHei UI", 9)).pack(anchor="w", pady=(0, 4))
        log_body = tk.Frame(log_holder, bg=PALETTE["card"])
        log_body.pack(fill="both", expand=True)
        self.log_text = tk.Text(log_body, height=4, state="disabled",
                                wrap="word", font=("Consolas", 9),
                                bd=0, relief="flat", highlightthickness=0,
                                bg=PALETTE["card"], fg=PALETTE["fg"])
        log_sb = Scrollbar(log_body, orient="vertical",
                           command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_sb.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        log_sb.pack(side="right", fill="y")

    # -- 列表状态 ------------------------------------------------------------
    def _refresh_tags(self):
        for key, color in STATUS_COLOR.items():
            if color:
                self.tree.tag_configure(key, foreground=color)

    def _refresh_list(self):
        self.tree.delete(*self.tree.get_children())
        for path, state in self.files:
            tags = (state,) if STATUS_COLOR.get(state) else ()
            self.tree.insert("", "end", values=(
                os.path.basename(path), path, STATUS_TEXT.get(state, state)),
                tags=tags)
        if getattr(self, "empty_label", None) is not None:
            if self.files:
                self.empty_label.place_forget()
            else:
                self.empty_label.place(relx=0.5, rely=0.5, anchor="center")
        if self._worker and self._worker.is_alive():
            self.status_var.set(f"正在转换… {len(self.files)} 个文件")
        else:
            self.status_var.set(f"共 {len(self.files)} 个文件")

    # -- 文件操作 ------------------------------------------------------------
    def _append_paths(self, paths):
        existing = {p for p, _ in self.files}
        added = 0
        for p in paths:
            p = os.path.normpath(p)
            if p.lower().endswith(".pdf") and p not in existing:
                self.files.append([p, "pending"])
                existing.add(p)
                added += 1
        self._refresh_list()
        self._log(f"已添加 {added} 个文件")

    # -- 拖拽落点注册（由主界面在页面切换时调用） -----------------------------
    def _register_dnd(self):
        """把本页的文件列表注册为拖拽落点，并使其成为 tkdnd 的当前目标。"""
        _register_page_drop(self, self.tree)
        if HAS_DND:
            for w in (self.tree, getattr(self, "empty_label", None)):
                if w is None:
                    continue
                try:
                    w.drop_target_register(tkdnd.DND_FILES)
                    w.dnd_bind("<<Drop>>", self._on_drop)
                except Exception:  # noqa: BLE001
                    pass

    def _unregister_dnd(self):
        """取消本页的拖拽落点注册（页面切走后调用）。"""
        if not (HAS_DND and tkdnd is not None):
            return
        for w in (self.tree, getattr(self, "empty_label", None)):
            if w is None:
                continue
            try:
                w.drop_target_unregister()
            except Exception:  # noqa: BLE001
                pass

    def _on_drop(self, event):
        """拖拽添加：文件直接取用，文件夹递归扫描其中的 PDF。"""
        items = self.root.tk.splitlist(event.data)
        paths = []
        for it in items:
            if os.path.isdir(it):
                for dp, _dn, fns in os.walk(it):
                    for fn in fns:
                        if fn.lower().endswith(".pdf"):
                            paths.append(os.path.join(dp, fn))
            elif it.lower().endswith(".pdf"):
                paths.append(it)
        if paths:
            self._append_paths(sorted(paths, key=_natural_key))
        else:
            messagebox.showinfo("提示", "未识别到可转换的 PDF 文件。")

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="选择要转换的 PDF 文件",
            filetypes=[("PDF 文件", "*.pdf"), ("所有文件", "*.*")])
        if paths:
            self._append_paths(paths)

    def _add_folder(self):
        folder = filedialog.askdirectory(title="选择包含 PDF 的文件夹")
        if not folder:
            return
        found = []
        for dp, _dn, fns in os.walk(folder):
            for fn in fns:
                if fn.lower().endswith(".pdf"):
                    found.append(os.path.join(dp, fn))
        if not found:
            messagebox.showinfo("提示", "该文件夹下未找到 PDF 文件。")
            return
        self._append_paths(found)

    def _remove_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        sel_paths = {self.tree.item(i, "values")[1] for i in sel}
        self.files = [f for f in self.files if f[0] not in sel_paths]
        self._refresh_list()

    def _move(self, delta):
        move_selected(self, delta, "文件")

    def _clear_list(self):
        if self._worker and self._worker.is_alive():
            messagebox.showwarning("提示", "正在转换中，无法清空列表。")
            return
        self.files.clear()
        self._refresh_list()
        self.progress["value"] = 0

    def _choose_out_dir(self):
        d = filedialog.askdirectory(title="选择图片保存位置")
        if d:
            self.out_var.set(d)

    # -- 转换流程 ------------------------------------------------------------
    def _start_convert(self):
        if self._worker and self._worker.is_alive():
            self._stop_flag.set()
            self.status_var.set("正在停止…")
            return
        if not self.files:
            messagebox.showinfo("提示", "请先添加要转换的 PDF 文件。")
            return
        out_dir = self.out_var.get().strip()
        if not out_dir or not os.path.isdir(out_dir):
            messagebox.showwarning("提示", "请选择图片保存位置。")
            return

        dpi = int(self.dpi_combo.current() == 0 and 144 or
                  (216 if self.dpi_combo.current() == 1 else 288))
        zoom = dpi / 72.0
        fmt = "jpg" if "JPG" in self.fmt_combo.get() else "png"

        self._stop_flag.clear()
        self.progress["maximum"] = len(self.files)
        self.progress["value"] = 0
        for i in range(len(self.files)):
            self.files[i][1] = "pending"
        self._refresh_list()

        self._log(f"开始批量转换（{dpi} dpi，{fmt.upper()}）…")
        self._set_convert_btn("停止", self.app.BTN_DANGER)

        self._worker = threading.Thread(
            target=self._worker_run, args=(list(self.files), out_dir, zoom, fmt),
            daemon=True)
        self._worker.start()

    def _worker_run(self, files_snapshot, out_dir, zoom, fmt):
        ok = err = skip = 0
        try:
            for idx, (path, _s) in enumerate(files_snapshot):
                if self._stop_flag.is_set():
                    self._queue.put(("stopped", None))
                    break
                self._queue.put(("status", (idx, "working")))
                state, info = convert_pdf_to_images(
                    path, out_dir, zoom, fmt, stop_flag=self._stop_flag)
                if state == "ok":
                    ok += 1
                elif state == "error":
                    err += 1
                else:
                    skip += 1
                self._queue.put(("item", (idx, state, info, path)))
            else:
                self._queue.put(("done", (ok, err, skip)))
        except Exception:  # noqa: BLE001
            self._queue.put(("fatal", traceback.format_exc()))

    def _set_convert_btn(self, text, style):
        if HAS_TTB:
            self.btn_convert.configure(text=text, bootstyle=style)
        else:
            self.btn_convert.configure(text=text)

    def _poll_queue(self):
        try:
            while True:
                msg = self._queue.get_nowait()
                self._handle_msg(msg)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _handle_msg(self, msg):
        kind, payload = msg[0], msg[1]
        children = self.tree.get_children()
        if kind == "status":
            idx, state = payload
            if idx < len(children):
                self.files[idx][1] = state
                self.tree.item(children[idx], values=(
                    os.path.basename(self.files[idx][0]),
                    self.files[idx][0], STATUS_TEXT.get(state, state)),
                    tags=(state,) if STATUS_COLOR.get(state) else ())
        elif kind == "item":
            idx, state, info, path = payload
            if idx < len(children):
                self.files[idx][1] = state
                self.tree.item(children[idx], values=(
                    os.path.basename(path), path, STATUS_TEXT.get(state, state)),
                    tags=(state,) if STATUS_COLOR.get(state) else ())
            self.progress["value"] = idx + 1
            self._log(info)
        elif kind == "done":
            ok, err, skip = payload
            self.progress["value"] = self.progress["maximum"]
            self.status_var.set(f"完成：成功 {ok} · 失败 {err} · 跳过 {skip}")
            self._log(f"转换结束：成功 {ok} 个，失败 {err} 个，跳过 {skip} 个")
            self._set_convert_btn("开始转换", self.app.BTN_SUCCESS)
            self._worker = None
        elif kind == "stopped":
            self.status_var.set("已停止")
            self._log("转换已手动停止")
            self._set_convert_btn("开始转换", self.app.BTN_SUCCESS)
            self._worker = None
        elif kind == "fatal":
            self._log("发生严重错误：\n" + payload)
            self._set_convert_btn("开始转换", self.app.BTN_SUCCESS)
            self._worker = None

    def _log(self, text):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")


# ---------------------------------------------------------------------------
# 图片转 PDF 页（标签页之三）
# ---------------------------------------------------------------------------
IMG_EXTS = (".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp", ".gif")

PPT_EXTS = (".ppt", ".pptx", ".pptm", ".pps", ".ppsx")


def _natural_key(path):
    """自然排序键：让「图2」排在「图10」前面。"""
    import re
    return [int(s) if s.isdigit() else s.lower()
            for s in re.split(r"(\d+)", os.path.basename(path))]


def _human_size(num_bytes):
    if num_bytes >= 1024 * 1024:
        return f"{num_bytes / 1024 / 1024:.1f} MB"
    return f"{num_bytes / 1024:.0f} KB"


class ImagesToPdfTab:
    """图片转 PDF 功能页，界面与其他页保持一致。"""

    def __init__(self, parent, app):
        self.parent = parent
        self.app = app            # 复用主程序的样式/颜色/root
        self.root = app.root
        self.files = []           # [[完整路径, 状态 key], ...]
        self.out_var = tk.StringVar()
        self._stop_flag = threading.Event()
        self._worker = None
        self._queue = queue.Queue()
        self._out_touched = False  # 用户是否手动指定过输出文件
        self._build()
        self.root.after(100, self._poll_queue)

    # -- 组件快捷构造 --------------------------------------------------------
    def _btn(self, master, text, style, command=None, **kw):
        return RoundButton(master, text=text, bootstyle=style or "secondary",
                           command=command,
                           parent_bg=kw.pop("parent_bg", PALETTE["bg"]), **kw)

    def _apply_theme(self):
        fg = PALETTE["fg"]
        self.log_text.configure(bg=PALETTE["card"], fg=fg, insertbackground=fg)

    # -- UI -----------------------------------------------------------------
    def _build(self):
        Label = _widget_class("Label")
        Frame = _widget_class("Frame")
        Treeview = _widget_class("Treeview")
        Scrollbar = _widget_class("Scrollbar")

        PX = 20
        BG = PALETTE["bg"]
        B = self.app

        # 标题
        header = BgCanvas(self.parent)
        header.pack(fill="x", padx=PX, pady=(12, 6))
        self.app.brand_row(header, "图片转 PDF",
                           "多张图片按顺序合并为一个 PDF · 列表顺序即页面顺序")

        # 工具栏（圆角按钮）
        toolbar = BgCanvas(self.parent)
        toolbar.pack(fill="x", padx=PX, pady=(4, 0))
        self._btn(toolbar, "添加图片", B.BTN_PRIMARY,
                  self._add_files, pad_x=14).pack(side="left", padx=(0, 6))
        self._btn(toolbar, "添加文件夹", B.BTN_SECONDARY,
                  self._add_folder, pad_x=14).pack(side="left", padx=6)
        self._btn(toolbar, "移除选中", B.BTN_SECONDARY,
                  self._remove_selected, pad_x=14).pack(side="left", padx=6)
        self._btn(toolbar, "清空列表", B.BTN_SECONDARY,
                  self._clear_list, pad_x=14).pack(side="left", padx=6)
        self.btn_up = self._btn(toolbar, "上移", B.BTN_SECONDARY,
                                lambda: self._move(-1), pad_x=14)
        self.btn_up.pack(side="left", padx=(14, 6))
        self.btn_down = self._btn(toolbar, "下移", B.BTN_SECONDARY,
                                  lambda: self._move(1), pad_x=14)
        self.btn_down.pack(side="left", padx=6)

        # 图片列表（圆角卡片）
        list_card = RoundCard(self.parent, padding=6, parent_bg=BG)
        list_card.pack(fill="both", expand=True, padx=PX, pady=(10, 0))
        list_frame = list_card.body
        cols = ("idx", "name", "size", "status")
        self.tree = Treeview(list_frame, columns=cols, show="headings",
                             selectmode="extended", height=3)
        self.tree.heading("idx", text="顺序")
        self.tree.heading("name", text="图片")
        self.tree.heading("size", text="大小")
        self.tree.heading("status", text="状态")
        self.tree.column("idx", width=52, anchor="center", stretch=False)
        self.tree.column("name", width=360, anchor="w")
        self.tree.column("size", width=90, anchor="e", stretch=False)
        self.tree.column("status", width=90, anchor="center", stretch=False)
        vsb = Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self._refresh_tags()

        # 拖拽添加（直接拖到列表上）
        if HAS_DND:
            try:
                self.tree.drop_target_register(tkdnd.DND_FILES)
                self.tree.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass

        # 空状态提示（叠加在列表中央，提示可拖拽）
        self.empty_label = tk.Label(
            list_frame,
            text="还没有添加图片\n点击「添加图片」或把图片 / 文件夹拖到列表上",
            justify="center", bg=PALETTE["card"], fg=PALETTE["muted"],
            font=("Microsoft YaHei UI", 11))
        self.empty_label.place(relx=0.5, rely=0.5, anchor="center")
        if HAS_DND:
            # 空状态提示浮在列表中央，必须同样注册为落点，
            # 否则拖到提示文字上的文件会被这个标签吞掉、无法触发 Drop。
            try:
                self.empty_label.drop_target_register(tkdnd.DND_FILES)
                self.empty_label.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:  # noqa: BLE001
                pass

        # 输出文件（圆角输入框）
        out_frame = BgCanvas(self.parent)
        out_frame.pack(fill="x", padx=PX, pady=(10, 0))
        Label(out_frame, text="保存到").pack(side="left")
        self.out_entry = RoundEntry(out_frame, textvariable=self.out_var,
                                    height=36, parent_bg=BG)
        self.out_entry.pack(side="left", fill="x", expand=True, padx=(10, 8))
        self._btn(out_frame, "浏览…", B.BTN_SECONDARY,
                  self._choose_out_file).pack(side="left")

        # 选项：页面尺寸 + 页边距（圆角下拉框）
        opt_frame = BgCanvas(self.parent)
        opt_frame.pack(pady=(10, 0))  # 水平居中
        Label(opt_frame, text="页面尺寸").pack(side="left")
        self.page_combo = RoundCombo(
            opt_frame, width=20, parent_bg=BG,
            values=("跟随图片尺寸（推荐）", "A4 纵向（适合打印）", "A4 横向（适合打印）"))
        self.page_combo.current(0)
        self.page_combo.pack(side="left", padx=(10, 20))

        Label(opt_frame, text="页边距").pack(side="left")
        self.margin_combo = RoundCombo(
            opt_frame, width=10, parent_bg=BG,
            values=("0 mm", "10 mm（推荐）", "20 mm"))
        self.margin_combo.current(1)
        self.margin_combo.pack(side="left", padx=(10, 6))
        Label(opt_frame, text="（仅 A4 模式生效）", foreground=PALETTE["muted"],
              font=("Microsoft YaHei UI", 9)).pack(side="left")

        # 进度 + 转换按钮
        prog_frame = BgCanvas(self.parent)
        prog_frame.pack(fill="x", padx=PX, pady=(12, 0))
        self.progress = RoundProgress(prog_frame, mode="determinate",
                                      height=14, parent_bg=BG)
        self.progress.pack(side="left", fill="x", expand=True, pady=8)
        self.btn_convert = self._btn(
            prog_frame, "开始转换", B.BTN_SUCCESS, self._start_convert)
        self.btn_convert.pack(side="right", padx=(14, 0))

        self.status_var = tk.StringVar(value="就绪")
        self.app._make_status_bar(self.parent, self.status_var)

        # 日志（圆角卡片）
        log_card = RoundCard(self.parent, padding=8, height=104, parent_bg=BG)
        log_card.pack(fill="x", padx=PX, pady=(8, 12))
        log_holder = log_card.body
        tk.Label(log_holder, text="转换日志", bg=PALETTE["card"],
                 fg=PALETTE["muted"],
                 font=("Microsoft YaHei UI", 9)).pack(anchor="w", pady=(0, 4))
        log_body = tk.Frame(log_holder, bg=PALETTE["card"])
        log_body.pack(fill="both", expand=True)
        self.log_text = tk.Text(log_body, height=4, state="disabled",
                                wrap="word", font=("Consolas", 9),
                                bd=0, relief="flat", highlightthickness=0,
                                bg=PALETTE["card"], fg=PALETTE["fg"])
        log_sb = Scrollbar(log_body, orient="vertical",
                           command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_sb.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        log_sb.pack(side="right", fill="y")

    # -- 列表 ----------------------------------------------------------------
    def _refresh_tags(self):
        for key, color in STATUS_COLOR.items():
            if color:
                self.tree.tag_configure(key, foreground=color)

    def _refresh_list(self):
        self.tree.delete(*self.tree.get_children())
        for i, (path, state) in enumerate(self.files, 1):
            try:
                size = _human_size(os.path.getsize(path))
            except OSError:
                size = "—"
            tags = (state,) if STATUS_COLOR.get(state) else ()
            self.tree.insert("", "end", values=(
                i, os.path.basename(path), size,
                STATUS_TEXT.get(state, state)), tags=tags)
        if getattr(self, "empty_label", None) is not None:
            if self.files:
                self.empty_label.place_forget()
            else:
                self.empty_label.place(relx=0.5, rely=0.5, anchor="center")
        if self._worker and self._worker.is_alive():
            self.status_var.set(f"正在合并… 共 {len(self.files)} 张")
        else:
            self.status_var.set(f"共 {len(self.files)} 张图片")

    def _set_item_state(self, idx, state):
        children = self.tree.get_children()
        if idx >= len(children) or idx >= len(self.files):
            return
        self.files[idx][1] = state
        path = self.files[idx][0]
        try:
            size = _human_size(os.path.getsize(path))
        except OSError:
            size = "—"
        self.tree.item(children[idx], values=(
            idx + 1, os.path.basename(path), size,
            STATUS_TEXT.get(state, state)),
            tags=(state,) if STATUS_COLOR.get(state) else ())

    # -- 文件操作 ------------------------------------------------------------
    def _append_paths(self, paths, resort=False):
        existing = {p for p, _ in self.files}
        added = 0
        new_items = []
        for p in paths:
            p = os.path.normpath(p)
            if p.lower().endswith(IMG_EXTS) and p not in existing:
                new_items.append([p, "pending"])
                existing.add(p)
                added += 1
        if resort:
            new_items.sort(key=lambda it: _natural_key(it[0]))
        self.files.extend(new_items)
        self._refresh_list()
        if added and not self._out_touched and not self.out_var.get().strip():
            self._suggest_out_name()
        self._log(f"已添加 {added} 张图片")

    def _suggest_out_name(self):
        if not self.files:
            return
        folder = os.path.dirname(self.files[0][0])
        self.out_var.set(os.path.join(folder, "合并输出.pdf"))

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="选择要合并的图片（Ctrl / Shift 可多选）",
            filetypes=[("图片文件", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff *.webp *.gif"),
                       ("所有文件", "*.*")])
        if paths:
            # 按文件对话框返回的顺序（即用户选择顺序）加入
            self._append_paths(list(paths))

    def _add_folder(self):
        folder = filedialog.askdirectory(title="选择包含图片的文件夹")
        if not folder:
            return
        found = []
        for dp, _dn, fns in os.walk(folder):
            for fn in fns:
                if fn.lower().endswith(IMG_EXTS):
                    found.append(os.path.join(dp, fn))
        if not found:
            messagebox.showinfo("提示", "该文件夹下未找到图片文件。")
            return
        self._append_paths(found, resort=True)

    def _remove_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        idxs = sorted((self.tree.index(i) for i in sel), reverse=True)
        for i in idxs:
            if 0 <= i < len(self.files):
                self.files.pop(i)
        self._refresh_list()

    def _clear_list(self):
        if self._worker and self._worker.is_alive():
            messagebox.showwarning("提示", "正在合并中，无法清空列表。")
            return
        self.files.clear()
        self._out_touched = False
        self._refresh_list()
        self.progress["value"] = 0

    def _move(self, delta):
        if not self.files:
            return
        if self._worker and self._worker.is_alive():
            return
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先在列表里选中要调整顺序的图片。")
            return
        idxs = sorted((self.tree.index(i) for i in sel), reverse=(delta > 0))
        if delta < 0 and idxs[0] == 0:
            return
        if delta > 0 and idxs[0] >= len(self.files) - 1:
            return
        for i in idxs:
            j = i + delta
            self.files[i], self.files[j] = self.files[j], self.files[i]
        self._refresh_list()
        children = self.tree.get_children()
        for i in idxs:
            self.tree.selection_add(children[i + delta])

    def _choose_out_file(self):
        initial = self.out_var.get().strip()
        path = filedialog.asksaveasfilename(
            title="设置输出 PDF 保存名称与路径",
            defaultextension=".pdf",
            filetypes=[("PDF 文件", "*.pdf")],
            initialfile=os.path.basename(initial) if initial else "合并输出.pdf",
            initialdir=os.path.dirname(initial) if initial else None)
        if path:
            self.out_var.set(path)
            self._out_touched = True

    # -- 拖拽落点注册（由主界面在页面切换时调用） -----------------------------
    def _register_dnd(self):
        """把本页的文件列表注册为拖拽落点，并使其成为 tkdnd 的当前目标。"""
        _register_page_drop(self, self.tree)
        if HAS_DND:
            for w in (self.tree, getattr(self, "empty_label", None)):
                if w is None:
                    continue
                try:
                    w.drop_target_register(tkdnd.DND_FILES)
                    w.dnd_bind("<<Drop>>", self._on_drop)
                except Exception:  # noqa: BLE001
                    pass

    def _unregister_dnd(self):
        """取消本页的拖拽落点注册（页面切走后调用）。"""
        if not (HAS_DND and tkdnd is not None):
            return
        for w in (self.tree, getattr(self, "empty_label", None)):
            if w is None:
                continue
            try:
                w.drop_target_unregister()
            except Exception:  # noqa: BLE001
                pass

    def _on_drop(self, event):
        items = self.root.tk.splitlist(event.data)
        paths = []
        for it in items:
            if os.path.isdir(it):
                for dp, _dn, fns in os.walk(it):
                    for fn in fns:
                        if fn.lower().endswith(IMG_EXTS):
                            paths.append(os.path.join(dp, fn))
            elif os.path.isfile(it) and it.lower().endswith(IMG_EXTS):
                paths.append(it)
        if paths:
            self._append_paths(paths, resort=True)
        else:
            self._log("拖入的内容里没有可用的图片")

    # -- 合并流程 ------------------------------------------------------------
    def _start_convert(self):
        if self._worker and self._worker.is_alive():
            self._stop_flag.set()
            self.status_var.set("正在停止…")
            return
        if not self.files:
            messagebox.showinfo("提示", "请先添加要合并的图片。")
            return

        out_pdf = self.out_var.get().strip()
        if not out_pdf:
            messagebox.showwarning("提示", "请设置输出 PDF 的保存位置。")
            return
        if not out_pdf.lower().endswith(".pdf"):
            out_pdf += ".pdf"
            self.out_var.set(out_pdf)
        out_pdf = os.path.abspath(out_pdf)

        page_idx = self.page_combo.current()
        page_mode = ("auto", "a4", "a4-l")[page_idx]
        margin_mm = (0, 10, 20)[self.margin_combo.current()]

        paths = [p for p, _ in self.files]
        self._stop_flag.clear()
        self.progress["maximum"] = len(paths)
        self.progress["value"] = 0
        for i in range(len(self.files)):
            self.files[i][1] = "pending"
        self._refresh_list()

        label = {"auto": "跟随图片", "a4": "A4 纵向", "a4-l": "A4 横向"}[page_mode]
        self._log(f"开始转换 {len(paths)} 张图片（{label}）…")
        self._log(f"输出：{out_pdf}")
        self._set_convert_btn("停止", self.app.BTN_DANGER)

        self._worker = threading.Thread(
            target=self._worker_run,
            args=(paths, out_pdf, page_mode, margin_mm), daemon=True)
        self._worker.start()

    def _worker_run(self, paths, out_pdf, page_mode, margin_mm):
        def cb(done, total):
            self._queue.put(("progress", (done, total)))

        try:
            state, info = images_to_pdf(
                paths, out_pdf, page_mode, margin_mm,
                stop_flag=self._stop_flag, progress_cb=cb)
        except Exception:  # noqa: BLE001
            self._queue.put(("fatal", traceback.format_exc()))
            return
        self._queue.put(("result", (state, info, out_pdf)))

    def _set_convert_btn(self, text, style):
        if HAS_TTB:
            self.btn_convert.configure(text=text, bootstyle=style)
        else:
            self.btn_convert.configure(text=text)

    def _poll_queue(self):
        try:
            while True:
                self._handle_msg(self._queue.get_nowait())
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _handle_msg(self, msg):
        kind, payload = msg[0], msg[1]
        if kind == "progress":
            done, total = payload
            self.progress["maximum"] = total
            self.progress["value"] = done
            self.status_var.set(f"正在合并… {done}/{total}")
            self._set_item_state(done - 1, "ok")
        elif kind == "result":
            state, info, out_pdf = payload
            self.progress["value"] = self.progress["maximum"]
            self._log(info)
            if state == "ok":
                self.status_var.set("合并完成")
                self._set_convert_btn("开始转换", self.app.BTN_SUCCESS)
            elif state == "skip":
                self.status_var.set("已停止")
                self._log("合并已手动停止（未生成完整 PDF）")
                self._set_convert_btn("开始转换", self.app.BTN_SUCCESS)
            else:
                self.status_var.set("失败")
                messagebox.showerror("失败", info)
                self._set_convert_btn("开始转换", self.app.BTN_SUCCESS)
            self._worker = None
        elif kind == "fatal":
            self._log("发生严重错误：\n" + payload)
            self._set_convert_btn("开始转换", self.app.BTN_SUCCESS)
            self._worker = None

    def _log(self, text):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")


class MergePdfsTab:
    """PDF 合并功能页：多个 PDF 按顺序合并为一个文件。"""

    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.root = app.root
        self.files = []           # [[完整路径, 状态 key], ...]
        self.out_var = tk.StringVar()
        self._stop_flag = threading.Event()
        self._worker = None
        self._queue = queue.Queue()
        self._out_touched = False
        self._build()
        self.root.after(100, self._poll_queue)

    # -- 组件快捷构造 --------------------------------------------------------
    def _btn(self, master, text, style, command=None, **kw):
        return RoundButton(master, text=text, bootstyle=style or "secondary",
                           command=command,
                           parent_bg=kw.pop("parent_bg", PALETTE["bg"]), **kw)

    def _apply_theme(self):
        fg = PALETTE["fg"]
        self.log_text.configure(bg=PALETTE["card"], fg=fg, insertbackground=fg)

    # -- UI -----------------------------------------------------------------
    def _build(self):
        Label = _widget_class("Label")
        Frame = _widget_class("Frame")
        Treeview = _widget_class("Treeview")
        Scrollbar = _widget_class("Scrollbar")

        PX = 20
        BG = PALETTE["bg"]
        B = self.app

        # 标题
        header = BgCanvas(self.parent)
        header.pack(fill="x", padx=PX, pady=(12, 6))
        self.app.brand_row(header, "PDF 合并",
                           "多个 PDF 按顺序合并为一个文件 · 列表顺序即合并顺序")

        # 工具栏
        toolbar = BgCanvas(self.parent)
        toolbar.pack(fill="x", padx=PX, pady=(4, 0))
        self._btn(toolbar, "添加文件", B.BTN_PRIMARY,
                  self._add_files, pad_x=14).pack(side="left", padx=(0, 6))
        self._btn(toolbar, "添加文件夹", B.BTN_SECONDARY,
                  self._add_folder, pad_x=14).pack(side="left", padx=6)
        self._btn(toolbar, "移除选中", B.BTN_SECONDARY,
                  self._remove_selected, pad_x=14).pack(side="left", padx=6)
        self._btn(toolbar, "清空列表", B.BTN_SECONDARY,
                  self._clear_list, pad_x=14).pack(side="left", padx=6)
        self.btn_up = self._btn(toolbar, "上移", B.BTN_SECONDARY,
                                lambda: self._move(-1), pad_x=14)
        self.btn_up.pack(side="left", padx=(14, 6))
        self.btn_down = self._btn(toolbar, "下移", B.BTN_SECONDARY,
                                  lambda: self._move(1), pad_x=14)
        self.btn_down.pack(side="left", padx=6)

        # 文件列表（圆角卡片）
        list_card = RoundCard(self.parent, padding=6, parent_bg=BG)
        list_card.pack(fill="both", expand=True, padx=PX, pady=(10, 0))
        list_frame = list_card.body
        cols = ("idx", "name", "pages", "size", "status")
        self.tree = Treeview(list_frame, columns=cols, show="headings",
                             selectmode="extended", height=3)
        self.tree.heading("idx", text="顺序")
        self.tree.heading("name", text="PDF 文件")
        self.tree.heading("pages", text="页数")
        self.tree.heading("size", text="大小")
        self.tree.heading("status", text="状态")
        self.tree.column("idx", width=52, anchor="center", stretch=False)
        self.tree.column("name", width=330, anchor="w")
        self.tree.column("pages", width=60, anchor="center", stretch=False)
        self.tree.column("size", width=90, anchor="e", stretch=False)
        self.tree.column("status", width=90, anchor="center", stretch=False)
        vsb = Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self._refresh_tags()

        # 拖拽添加（直接拖到列表上）
        if HAS_DND:
            try:
                self.tree.drop_target_register(tkdnd.DND_FILES)
                self.tree.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass

        # 空状态提示（叠加在列表中央，提示可拖拽）
        self.empty_label = tk.Label(
            list_frame,
            text="还没有添加文件\n点击「添加文件」或把 PDF 文件 / 文件夹拖到列表上",
            justify="center", bg=PALETTE["card"], fg=PALETTE["muted"],
            font=("Microsoft YaHei UI", 11))
        self.empty_label.place(relx=0.5, rely=0.5, anchor="center")
        if HAS_DND:
            # 空状态提示浮在列表中央，必须同样注册为落点，
            # 否则拖到提示文字上的文件会被这个标签吞掉、无法触发 Drop。
            try:
                self.empty_label.drop_target_register(tkdnd.DND_FILES)
                self.empty_label.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:  # noqa: BLE001
                pass

        # 输出文件
        out_frame = BgCanvas(self.parent)
        out_frame.pack(fill="x", padx=PX, pady=(10, 0))
        Label(out_frame, text="保存到").pack(side="left")
        self.out_entry = RoundEntry(out_frame, textvariable=self.out_var,
                                    height=36, parent_bg=BG)
        self.out_entry.pack(side="left", fill="x", expand=True, padx=(10, 8))
        self._btn(out_frame, "浏览…", B.BTN_SECONDARY,
                  self._choose_out_file).pack(side="left")

        # 进度 + 合并按钮
        prog_frame = BgCanvas(self.parent)
        prog_frame.pack(fill="x", padx=PX, pady=(12, 0))
        self.progress = RoundProgress(prog_frame, mode="determinate",
                                      height=14, parent_bg=BG)
        self.progress.pack(side="left", fill="x", expand=True, pady=8)
        self.btn_convert = self._btn(
            prog_frame, "开始合并", B.BTN_SUCCESS, self._start_convert)
        self.btn_convert.pack(side="right", padx=(14, 0))

        self.status_var = tk.StringVar(value="就绪")
        self.app._make_status_bar(self.parent, self.status_var)

        # 日志（圆角卡片）
        log_card = RoundCard(self.parent, padding=8, height=104, parent_bg=BG)
        log_card.pack(fill="x", padx=PX, pady=(8, 12))
        log_holder = log_card.body
        tk.Label(log_holder, text="转换日志", bg=PALETTE["card"],
                 fg=PALETTE["muted"],
                 font=("Microsoft YaHei UI", 9)).pack(anchor="w", pady=(0, 4))
        log_body = tk.Frame(log_holder, bg=PALETTE["card"])
        log_body.pack(fill="both", expand=True)
        self.log_text = tk.Text(log_body, height=4, state="disabled",
                                wrap="word", font=("Consolas", 9),
                                bd=0, relief="flat", highlightthickness=0,
                                bg=PALETTE["card"], fg=PALETTE["fg"])
        log_sb = Scrollbar(log_body, orient="vertical",
                           command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_sb.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        log_sb.pack(side="right", fill="y")

    # -- 列表 ----------------------------------------------------------------
    def _refresh_tags(self):
        for key, color in STATUS_COLOR.items():
            if color:
                self.tree.tag_configure(key, foreground=color)

    def _pdf_meta(self, path):
        """返回 (页数文本, 大小文本)。读不到时给出占位。"""
        try:
            import fitz
            src = fitz.open(path)
            try:
                pages = str(src.page_count) if not src.needs_pass else "加密"
            finally:
                src.close()
        except Exception:  # noqa: BLE001
            pages = "—"
        try:
            size = _human_size(os.path.getsize(path))
        except OSError:
            size = "—"
        return pages, size

    def _refresh_list(self):
        self.tree.delete(*self.tree.get_children())
        for i, (path, state) in enumerate(self.files, 1):
            pages, size = self._pdf_meta(path)
            tags = (state,) if STATUS_COLOR.get(state) else ()
            self.tree.insert("", "end", values=(
                i, os.path.basename(path), pages, size,
                STATUS_TEXT.get(state, state)), tags=tags)
        if getattr(self, "empty_label", None) is not None:
            if self.files:
                self.empty_label.place_forget()
            else:
                self.empty_label.place(relx=0.5, rely=0.5, anchor="center")
        if self._worker and self._worker.is_alive():
            self.status_var.set(f"正在合并… 共 {len(self.files)} 个")
        else:
            self.status_var.set(f"共 {len(self.files)} 个 PDF 文件")

    def _set_item_state(self, idx, state):
        children = self.tree.get_children()
        if idx >= len(children) or idx >= len(self.files):
            return
        self.files[idx][1] = state
        path = self.files[idx][0]
        pages, size = self._pdf_meta(path)
        self.tree.item(children[idx], values=(
            idx + 1, os.path.basename(path), pages, size,
            STATUS_TEXT.get(state, state)),
            tags=(state,) if STATUS_COLOR.get(state) else ())

    # -- 文件操作 ------------------------------------------------------------
    def _append_paths(self, paths, resort=False):
        existing = {p for p, _ in self.files}
        added = 0
        new_items = []
        for p in paths:
            p = os.path.normpath(p)
            if p.lower().endswith(".pdf") and p not in existing:
                new_items.append([p, "pending"])
                existing.add(p)
                added += 1
        if resort:
            new_items.sort(key=lambda it: _natural_key(it[0]))
        self.files.extend(new_items)
        self._refresh_list()
        if added and not self._out_touched and not self.out_var.get().strip():
            self._suggest_out_name()
        self._log(f"已添加 {added} 个 PDF 文件")

    def _suggest_out_name(self):
        if not self.files:
            return
        folder = os.path.dirname(self.files[0][0])
        self.out_var.set(os.path.join(folder, "合并完成.pdf"))

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="选择要合并的 PDF（Ctrl / Shift 可多选）",
            filetypes=[("PDF 文件", "*.pdf"), ("所有文件", "*.*")])
        if paths:
            self._append_paths(list(paths))

    def _add_folder(self):
        folder = filedialog.askdirectory(
            title="选择包含 PDF 的文件夹（自动扫描所有子目录）")
        if not folder:
            return
        found = []
        for dp, _dn, fns in os.walk(folder):
            for fn in fns:
                if fn.lower().endswith(".pdf"):
                    found.append(os.path.join(dp, fn))
        if not found:
            messagebox.showinfo("提示", "该文件夹下未找到 PDF 文件。")
            return
        self._append_paths(found, resort=True)

    def _remove_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        idxs = sorted((self.tree.index(i) for i in sel), reverse=True)
        for i in idxs:
            if 0 <= i < len(self.files):
                self.files.pop(i)
        self._refresh_list()

    def _clear_list(self):
        if self._worker and self._worker.is_alive():
            messagebox.showwarning("提示", "正在合并中，无法清空列表。")
            return
        self.files.clear()
        self._out_touched = False
        self._refresh_list()
        self.progress["value"] = 0

    def _move(self, delta):
        if not self.files:
            return
        if self._worker and self._worker.is_alive():
            return
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先在列表里选中要调整顺序的文件。")
            return
        idxs = sorted((self.tree.index(i) for i in sel), reverse=(delta > 0))
        if delta < 0 and idxs[0] == 0:
            return
        if delta > 0 and idxs[0] >= len(self.files) - 1:
            return
        for i in idxs:
            j = i + delta
            self.files[i], self.files[j] = self.files[j], self.files[i]
        self._refresh_list()
        children = self.tree.get_children()
        for i in idxs:
            self.tree.selection_add(children[i + delta])

    def _choose_out_file(self):
        initial = self.out_var.get().strip()
        path = filedialog.asksaveasfilename(
            title="设置合并后 PDF 的保存名称与路径",
            defaultextension=".pdf",
            filetypes=[("PDF 文件", "*.pdf")],
            initialfile=os.path.basename(initial) if initial else "合并完成.pdf",
            initialdir=os.path.dirname(initial) if initial else None)
        if path:
            self.out_var.set(path)
            self._out_touched = True

    # -- 拖拽落点注册（由主界面在页面切换时调用） -----------------------------
    def _register_dnd(self):
        """把本页的文件列表注册为拖拽落点，并使其成为 tkdnd 的当前目标。"""
        _register_page_drop(self, self.tree)
        if HAS_DND:
            for w in (self.tree, getattr(self, "empty_label", None)):
                if w is None:
                    continue
                try:
                    w.drop_target_register(tkdnd.DND_FILES)
                    w.dnd_bind("<<Drop>>", self._on_drop)
                except Exception:  # noqa: BLE001
                    pass

    def _unregister_dnd(self):
        """取消本页的拖拽落点注册（页面切走后调用）。"""
        if not (HAS_DND and tkdnd is not None):
            return
        for w in (self.tree, getattr(self, "empty_label", None)):
            if w is None:
                continue
            try:
                w.drop_target_unregister()
            except Exception:  # noqa: BLE001
                pass

    def _on_drop(self, event):
        items = self.root.tk.splitlist(event.data)
        paths = []
        for it in items:
            if os.path.isdir(it):
                for dp, _dn, fns in os.walk(it):
                    for fn in fns:
                        if fn.lower().endswith(".pdf"):
                            paths.append(os.path.join(dp, fn))
            elif os.path.isfile(it) and it.lower().endswith(".pdf"):
                paths.append(it)
        if paths:
            self._append_paths(paths, resort=True)
        else:
            self._log("拖入的内容里没有可用的 PDF 文件")

    # -- 合并流程 ------------------------------------------------------------
    def _start_convert(self):
        if self._worker and self._worker.is_alive():
            self._stop_flag.set()
            self.status_var.set("正在停止…")
            return
        if not self.files:
            messagebox.showinfo("提示", "请先添加要合并的 PDF 文件。")
            return
        if len(self.files) < 2:
            if not messagebox.askyesno(
                    "提示", "列表里只有 1 个 PDF，合并后内容不变。确定继续吗？"):
                return

        out_pdf = self.out_var.get().strip()
        if not out_pdf:
            messagebox.showwarning("提示", "请设置合并后 PDF 的保存位置。")
            return
        if not out_pdf.lower().endswith(".pdf"):
            out_pdf += ".pdf"
            self.out_var.set(out_pdf)
        out_pdf = os.path.abspath(out_pdf)

        # 输出文件若也在合并列表里，会读到不完整内容，提前拦截
        if out_pdf in [os.path.abspath(p) for p, _ in self.files]:
            messagebox.showwarning(
                "提示", "输出文件与列表中的某个 PDF 相同，请换个保存位置。")
            return

        paths = [p for p, _ in self.files]
        self._stop_flag.clear()
        self.progress["maximum"] = len(paths)
        self.progress["value"] = 0
        for i in range(len(self.files)):
            self.files[i][1] = "pending"
        self._refresh_list()

        self._log(f"开始合并 {len(paths)} 个 PDF 文件…")
        self._log(f"输出：{out_pdf}")
        self._set_convert_btn("停止", self.app.BTN_DANGER)

        self._worker = threading.Thread(
            target=self._worker_run,
            args=(paths, out_pdf), daemon=True)
        self._worker.start()

    def _worker_run(self, paths, out_pdf):
        def cb(done, total):
            self._queue.put(("progress", (done, total)))

        try:
            state, info = merge_pdfs(
                paths, out_pdf,
                stop_flag=self._stop_flag, progress_cb=cb)
        except Exception:  # noqa: BLE001
            self._queue.put(("fatal", traceback.format_exc()))
            return
        self._queue.put(("result", (state, info, out_pdf)))

    def _set_convert_btn(self, text, style):
        if HAS_TTB:
            self.btn_convert.configure(text=text, bootstyle=style)
        else:
            self.btn_convert.configure(text=text)

    def _poll_queue(self):
        try:
            while True:
                self._handle_msg(self._queue.get_nowait())
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _handle_msg(self, msg):
        kind, payload = msg[0], msg[1]
        if kind == "progress":
            done, total = payload
            self.progress["maximum"] = total
            self.progress["value"] = done
            self.status_var.set(f"正在合并… {done}/{total}")
            self._set_item_state(done - 1, "ok")
        elif kind == "result":
            state, info, out_pdf = payload
            self.progress["value"] = self.progress["maximum"]
            self._log(info)
            if state == "ok":
                self.status_var.set("合并完成")
                self._set_convert_btn("开始合并", self.app.BTN_SUCCESS)
            elif state == "skip":
                self.status_var.set("已停止")
                self._log("合并已手动停止（未生成完整 PDF）")
                self._set_convert_btn("开始合并", self.app.BTN_SUCCESS)
            else:
                self.status_var.set("失败")
                messagebox.showerror("失败", info)
                self._set_convert_btn("开始合并", self.app.BTN_SUCCESS)
            self._worker = None
        elif kind == "fatal":
            self._log("发生严重错误：\n" + payload)
            self._set_convert_btn("开始合并", self.app.BTN_SUCCESS)
            self._worker = None

    def _log(self, text):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")


class SplitPdfTab:
    """PDF 拆分功能页：把一个 PDF 按页拆分或按页码范围提取为多个 PDF。"""

    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.root = app.root
        self.src_var = tk.StringVar()
        self.range_var = tk.StringVar()
        self.out_var = tk.StringVar()
        self._out_touched = False
        self._stop_flag = threading.Event()
        self._worker = None
        self._queue = queue.Queue()
        self._build()
        self.root.after(100, self._poll_queue)

    # -- 组件快捷构造 --------------------------------------------------------
    def _btn(self, master, text, style, command=None, **kw):
        return RoundButton(master, text=text, bootstyle=style or "secondary",
                           command=command,
                           parent_bg=kw.pop("parent_bg", PALETTE["bg"]), **kw)

    def _apply_theme(self):
        fg = PALETTE["fg"]
        self.log_text.configure(bg=PALETTE["card"], fg=fg, insertbackground=fg)

    # -- UI -----------------------------------------------------------------
    def _build(self):
        Label = _widget_class("Label")
        Scrollbar = _widget_class("Scrollbar")

        PX = 20
        BG = PALETTE["bg"]
        B = self.app

        # 标题
        header = BgCanvas(self.parent)
        header.pack(fill="x", padx=PX, pady=(12, 6))
        self.app.brand_row(header, "PDF 拆分",
                           "把一个 PDF 每页拆分或按页码范围提取 · 输出为多个 PDF")

        # 源文件行
        src_frame = BgCanvas(self.parent)
        src_frame.pack(fill="x", padx=PX, pady=(4, 0))
        Label(src_frame, text="源文件").pack(side="left")
        self.src_entry = RoundEntry(src_frame, textvariable=self.src_var,
                                    height=36, parent_bg=BG)
        self.src_entry.pack(side="left", fill="x", expand=True, padx=(10, 8))
        self.src_var.trace_add("write", lambda *_a: self._auto_out_dir())
        self._btn(src_frame, "浏览…", B.BTN_SECONDARY,
                  self._choose_src).pack(side="left")

        # ---- 拖拽区（圆角虚线框，把 PDF 拖进来即作为源文件） ----
        self.drop_canvas = tk.Canvas(self.parent, height=72, highlightthickness=0,
                                     bd=0, bg=BG, cursor="hand2")
        self.drop_canvas.pack(fill="x", padx=PX, pady=(8, 0))
        self.drop_canvas.bind("<Button-1>", lambda e: self._choose_src())
        self.drop_canvas.bind("<Configure>",
                              lambda e: self._draw_drop_zone(self._drop_hover))
        self._register_dnd()

        # 拆分方式 + 页码范围
        opt_frame = BgCanvas(self.parent)
        opt_frame.pack(fill="x", padx=PX, pady=(10, 0))
        Label(opt_frame, text="拆分方式").pack(side="left")
        self.mode_combo = RoundCombo(
            opt_frame, width=22, parent_bg=BG,
            values=("每页一个 PDF（全部页面）", "按页码范围拆分（每段一个文件）"))
        self.mode_combo.current(0)
        self.mode_combo.pack(side="left", padx=(10, 18))
        Label(opt_frame, text="页码范围").pack(side="left")
        self.range_entry = RoundEntry(opt_frame, textvariable=self.range_var,
                                      height=36, width=200, parent_bg=BG)
        self.range_entry.pack(side="left", padx=(10, 8))
        tip = Label(opt_frame, text="如 1-3,5,8-10", foreground=PALETTE["muted"])
        tip.configure(font=("Microsoft YaHei UI", 9))
        tip.pack(side="left")

        # 输出目录行（拆分页的文本框下移，给拖拽区留位置）
        out_frame = BgCanvas(self.parent)
        out_frame.pack(fill="x", padx=PX, pady=(10, 0))
        Label(out_frame, text="保存到").pack(side="left")
        self.out_entry = RoundEntry(out_frame, textvariable=self.out_var,
                                    height=36, parent_bg=BG)
        self.out_entry.pack(side="left", fill="x", expand=True, padx=(10, 8))
        self.out_var.trace_add("write", lambda *_a: setattr(self, "_out_touched", True))
        self._btn(out_frame, "浏览…", B.BTN_SECONDARY,
                  self._choose_out_dir).pack(side="left")

        # 进度 + 拆分按钮
        prog_frame = BgCanvas(self.parent)
        prog_frame.pack(fill="x", padx=PX, pady=(12, 0))
        self.progress = RoundProgress(prog_frame, mode="determinate",
                                      height=14, parent_bg=BG)
        self.progress.pack(side="left", fill="x", expand=True, pady=8)
        self.btn_convert = self._btn(
            prog_frame, "开始拆分", B.BTN_SUCCESS, self._start_split)
        self.btn_convert.pack(side="right", padx=(14, 0))

        self.status_var = tk.StringVar(value="就绪")
        self.app._make_status_bar(self.parent, self.status_var)

        # 日志（圆角卡片）
        log_card = RoundCard(self.parent, padding=8, height=104, parent_bg=BG)
        log_card.pack(fill="x", padx=PX, pady=(8, 12))
        log_holder = log_card.body
        tk.Label(log_holder, text="转换日志", bg=PALETTE["card"],
                 fg=PALETTE["muted"],
                 font=("Microsoft YaHei UI", 9)).pack(anchor="w", pady=(0, 4))
        log_body = tk.Frame(log_holder, bg=PALETTE["card"])
        log_body.pack(fill="both", expand=True)
        self.log_text = tk.Text(log_body, height=4, state="disabled",
                                wrap="word", font=("Consolas", 9),
                                bd=0, relief="flat", highlightthickness=0,
                                bg=PALETTE["card"], fg=PALETTE["fg"])
        log_sb = Scrollbar(log_body, orient="vertical",
                           command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_sb.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        log_sb.pack(side="right", fill="y")

        self._on_mode()
        # 构建完成后再挂联动回调，避免 current(0) 触发时 range_entry 还不存在
        self.mode_combo._command = lambda _i=None: self._on_mode()

    # -- 拖拽区 --------------------------------------------------------------
    _drop_hover = False

    def _register_dnd(self):
        """把拖拽区注册为放置目标：拖入 PDF 即设为源文件。

        同时通过 _register_page_drop 让本页成为 tkdnd 的当前活跃目标，
        避免被其它页面抢先接收拖拽事件。
        """
        if not HAS_DND:
            return
        try:
            self.drop_canvas.drop_target_register(tkdnd.DND_FILES)
            self.drop_canvas.dnd_bind("<<Drop>>", self._on_drop)
            self.drop_canvas.dnd_bind("<<DropEnter>>", self._on_drop_enter)
            self.drop_canvas.dnd_bind("<<DropLeave>>", self._on_drop_leave)
        except Exception:  # noqa: BLE001
            pass
        _register_page_drop(self, self.drop_canvas)

    def _unregister_dnd(self):
        """取消本页拖拽区的注册（页面切走后调用）。"""
        if not (HAS_DND and tkdnd is not None):
            return
        try:
            self.drop_canvas.drop_target_unregister()
        except Exception:  # noqa: BLE001
            pass

    def _draw_drop_zone(self, hover=False):
        c = self.drop_canvas
        c.configure(bg=PALETTE["bg"])
        c.delete("all")
        w = c.winfo_width() or 840
        h = c.winfo_height() or 72
        fill = _shade(PALETTE["card"], 0.05) if hover else PALETTE["card"]
        outline = PALETTE["info"] if hover else PALETTE["border"]
        round_rect(c, 4, 4, w - 4, h - 4, CARD_RADIUS, fill=fill,
                   outline=outline, width=2, dash=(8, 6))
        c.create_text(w // 2, h // 2 - 11,
                      text="⬇  将 PDF 文件拖到这里",
                      fill=PALETTE["fg"], font=("Microsoft YaHei UI", 11, "bold"))
        c.create_text(w // 2, h // 2 + 12,
                      text="支持 .pdf，或点击此区域选择文件",
                      fill=PALETTE["muted"], font=("Microsoft YaHei UI", 9))

    def _on_drop_enter(self, _event):
        self._drop_hover = True
        self._draw_drop_zone(True)

    def _on_drop_leave(self, _event):
        self._drop_hover = False
        self._draw_drop_zone(False)

    def _on_drop(self, event):
        items = self.root.tk.splitlist(event.data)
        pdfs = []
        for it in items:
            if os.path.isdir(it):
                # 文件夹：只取一级目录下的 PDF，自然排序后取第一个
                found = [os.path.join(it, fn) for fn in os.listdir(it)
                         if fn.lower().endswith(".pdf")
                         and os.path.isfile(os.path.join(it, fn))]
                pdfs.extend(sorted(found, key=_natural_key))
            elif it.lower().endswith(".pdf"):
                pdfs.append(it)
        self._drop_hover = False
        self._draw_drop_zone(False)
        if not pdfs:
            messagebox.showinfo("提示", "未识别到 PDF 文件。\n"
                                        "PDF 拆分一次只处理一个文件，请拖入一个 .pdf。")
            return
        if len(pdfs) > 1:
            self._log(f"一次只能拆分一个 PDF，已使用第一个：{os.path.basename(pdfs[0])}")
        self.src_var.set(os.path.normpath(pdfs[0]))
        self._log(f"已选择源文件：{os.path.basename(pdfs[0])}")

    # -- 交互 ----------------------------------------------------------------
    def _on_mode(self):
        """切换拆分方式：每页拆分时禁用范围输入提示。"""
        by_range = self.mode_combo.current() == 1
        self.range_entry.configure(
            state="normal" if by_range else "disabled")

    def _choose_src(self):
        path = filedialog.askopenfilename(
            title="选择要拆分的 PDF",
            filetypes=[("PDF 文件", "*.pdf"), ("所有文件", "*.*")])
        if path:
            self.src_var.set(os.path.normpath(path))

    def _auto_out_dir(self):
        if self._out_touched:
            return
        src = self.src_var.get().strip()
        if src and os.path.isfile(src):
            self.out_var.set(os.path.dirname(src))

    def _choose_out_dir(self):
        d = filedialog.askdirectory(title="选择输出文件夹")
        if d:
            self.out_var.set(os.path.normpath(d))
            self._out_touched = True

    def _log(self, text):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    # -- 拆分流程 --------------------------------------------------------------
    def _start_split(self):
        if self._worker and self._worker.is_alive():
            return
        src = self.src_var.get().strip()
        if not src or not os.path.isfile(src):
            messagebox.showwarning("提示", "请先选择要拆分的 PDF 文件。")
            return
        by_range = self.mode_combo.current() == 1
        ranges = None
        if by_range:
            ranges = parse_page_ranges(self.range_var.get())
            if ranges is None or ranges == []:
                messagebox.showwarning(
                    "提示", "请填写页码范围，如：1-3,5,8-10（多个范围用逗号分隔）。")
                return
        out_dir = self.out_var.get().strip() or os.path.dirname(src)
        if not os.path.isdir(out_dir):
            os.makedirs(out_dir, exist_ok=True)

        self._stop_flag.clear()
        self.progress["value"] = 0
        self._log(f"开始拆分 {os.path.basename(src)}（{'按范围' if by_range else '每页'}）…")
        self._log(f"输出：{out_dir}")
        self.btn_convert.configure(text="停止", bootstyle=self.app.BTN_DANGER)

        self._worker = threading.Thread(
            target=self._worker_run,
            args=(src, out_dir, ranges), daemon=True)
        self._worker.start()

    def _worker_run(self, src, out_dir, ranges):
        def cb(done, total):
            self._queue.put(("progress", (done, total)))

        try:
            state, info = split_pdf(src, out_dir, ranges,
                                    stop_flag=self._stop_flag, progress_cb=cb)
        except Exception:  # noqa: BLE001
            self._queue.put(("fatal", traceback.format_exc()))
            return
        self._queue.put(("result", (state, info)))

    def _poll_queue(self):
        try:
            while True:
                self._handle_msg(self._queue.get_nowait())
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _handle_msg(self, msg):
        kind, payload = msg[0], msg[1]
        if kind == "progress":
            done, total = payload
            self.progress["maximum"] = max(total, 1)
            self.progress["value"] = done
            self.status_var.set(f"正在拆分… {done}/{total}")
        elif kind == "result":
            state, info = payload
            self.progress["value"] = self.progress["maximum"]
            self._log(info)
            if state == "ok":
                self.status_var.set("拆分完成")
                self.btn_convert.configure(text="开始拆分",
                                           bootstyle=self.app.BTN_SUCCESS)
            elif state == "skip":
                self.status_var.set("已停止")
                self.btn_convert.configure(text="开始拆分",
                                           bootstyle=self.app.BTN_SUCCESS)
            else:
                self.status_var.set("失败")
                messagebox.showerror("失败", info)
                self.btn_convert.configure(text="开始拆分",
                                           bootstyle=self.app.BTN_SUCCESS)
            self._worker = None
        elif kind == "fatal":
            self._log("发生严重错误：\n" + payload)
            self.btn_convert.configure(text="开始拆分",
                                       bootstyle=self.app.BTN_SUCCESS)
            self._worker = None


# ---------------------------------------------------------------------------
# PDF 加水印页（标签页之七）
# ---------------------------------------------------------------------------
WM_COLORS = ("红色", "灰色", "黑色", "蓝色")
WM_COLOR_RGB = {"红色": (1.0, 0.0, 0.0), "灰色": (0.55, 0.55, 0.55),
                "黑色": (0.0, 0.0, 0.0), "蓝色": (0.0, 0.25, 0.8)}
WM_OPACITY = ("浅（12%）", "中（20%）", "深（35%）")
WM_OPACITY_VAL = {"浅（12%）": 0.12, "中（20%）": 0.20, "深（35%）": 0.35}
WM_FONTSIZE = ("小（24）", "中（40）", "大（64）")
WM_FONTSIZE_VAL = {"小（24）": 24, "中（40）": 40, "大（64）": 64}


class WatermarkTab:
    """PDF 加水印功能页：给批量 PDF 加文字或图片水印。"""

    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.root = app.root
        self.files = []           # [[完整路径, 状态 key], ...]
        self.out_var = tk.StringVar()
        self._stop_flag = threading.Event()
        self._worker = None
        self._queue = queue.Queue()
        self._build()
        self.root.after(100, self._poll_queue)

    # -- 组件快捷构造 --------------------------------------------------------
    def _btn(self, master, text, style, command=None, **kw):
        return RoundButton(master, text=text, bootstyle=style or "secondary",
                           command=command,
                           parent_bg=kw.pop("parent_bg", PALETTE["bg"]), **kw)

    def _apply_theme(self):
        fg = PALETTE["fg"]
        self.log_text.configure(bg=PALETTE["card"], fg=fg, insertbackground=fg)

    # -- UI -----------------------------------------------------------------
    def _build(self):
        Label = _widget_class("Label")
        Frame = _widget_class("Frame")
        Treeview = _widget_class("Treeview")
        Scrollbar = _widget_class("Scrollbar")

        PX = 20
        BG = PALETTE["bg"]
        B = self.app

        # 标题
        header = BgCanvas(self.parent)
        header.pack(fill="x", padx=PX, pady=(12, 6))
        self.app.brand_row(header, "PDF 加水印",
                           "批量给 PDF 加文字 / 图片水印 · 本地处理，不上传")

        # 工具栏
        toolbar = BgCanvas(self.parent)
        toolbar.pack(fill="x", padx=PX, pady=(4, 0))
        self._btn(toolbar, "添加文件", B.BTN_PRIMARY,
                  self._add_files, pad_x=14).pack(side="left", padx=(0, 6))
        self._btn(toolbar, "添加文件夹", B.BTN_SECONDARY,
                  self._add_folder, pad_x=14).pack(side="left", padx=6)
        self._btn(toolbar, "移除选中", B.BTN_SECONDARY,
                  self._remove_selected, pad_x=14).pack(side="left", padx=6)
        self._btn(toolbar, "清空列表", B.BTN_SECONDARY,
                  self._clear_list, pad_x=14).pack(side="left", padx=6)
        self.btn_up = self._btn(toolbar, "上移", B.BTN_SECONDARY,
                                lambda: self._move(-1), pad_x=14)
        self.btn_up.pack(side="left", padx=(14, 6))
        self.btn_down = self._btn(toolbar, "下移", B.BTN_SECONDARY,
                                  lambda: self._move(1), pad_x=14)
        self.btn_down.pack(side="left", padx=6)

        # 文件列表（圆角卡片）
        list_card = RoundCard(self.parent, padding=6, parent_bg=BG)
        list_card.pack(fill="both", expand=True, padx=PX, pady=(10, 0))
        list_frame = list_card.body
        cols = ("name", "path", "status")
        self.tree = Treeview(list_frame, columns=cols, show="headings",
                             selectmode="extended", height=3)
        self.tree.heading("name", text="文件名")
        self.tree.heading("path", text="路径")
        self.tree.heading("status", text="状态")
        self.tree.column("name", width=230, anchor="w")
        self.tree.column("path", width=440, anchor="w")
        self.tree.column("status", width=90, anchor="center", stretch=False)
        vsb = Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self._refresh_tags()

        # 拖拽添加（直接拖到列表上）
        if HAS_DND:
            try:
                self.tree.drop_target_register(tkdnd.DND_FILES)
                self.tree.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass

        # 空状态提示
        self.empty_label = tk.Label(
            list_frame,
            text="还没有添加文件\n点击「添加文件」或把 PDF 文件 / 文件夹拖到列表上",
            justify="center", bg=PALETTE["card"], fg=PALETTE["muted"],
            font=("Microsoft YaHei UI", 11))
        self.empty_label.place(relx=0.5, rely=0.5, anchor="center")
        if HAS_DND:
            # 空状态提示浮在列表中央，必须同样注册为落点，
            # 否则拖到提示文字上的文件会被这个标签吞掉、无法触发 Drop。
            try:
                self.empty_label.drop_target_register(tkdnd.DND_FILES)
                self.empty_label.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:  # noqa: BLE001
                pass

        # 水印设置（两行）
        opt1 = BgCanvas(self.parent)
        opt1.pack(fill="x", padx=PX, pady=(10, 0))
        Label(opt1, text="水印类型").pack(side="left")
        self.type_combo = RoundCombo(
            opt1, width=14, parent_bg=BG,
            values=("文字水印", "图片水印"))
        self.type_combo.current(0)
        self.type_combo.pack(side="left", padx=(10, 18))
        self.lbl_text = Label(opt1, text="水印文字")
        self.lbl_text.pack(side="left")
        self.text_var = tk.StringVar(value="内部资料")
        self.text_entry = RoundEntry(opt1, textvariable=self.text_var,
                                     height=36, width=170, parent_bg=BG)
        self.text_entry.pack(side="left", padx=(10, 8))
        self.img_var = tk.StringVar()
        self.btn_img = self._btn(opt1, "选择图片…", B.BTN_SECONDARY,
                                 self._choose_image)
        self.img_label = Label(opt1, text="", foreground=PALETTE["muted"])
        self.img_label.configure(font=("Microsoft YaHei UI", 9))
        self.img_label.pack(side="left", padx=(8, 0))

        opt2 = BgCanvas(self.parent)
        opt2.pack(fill="x", padx=PX, pady=(8, 0))
        Label(opt2, text="版式").pack(side="left")
        self.layout_combo = RoundCombo(
            opt2, width=13, parent_bg=BG, values=("平铺整页", "页面居中"))
        self.layout_combo.current(0)
        self.layout_combo.pack(side="left", padx=(10, 18))
        Label(opt2, text="颜色").pack(side="left")
        self.color_combo = RoundCombo(
            opt2, width=9, parent_bg=BG, values=WM_COLORS)
        self.color_combo.current(0)
        self.color_combo.pack(side="left", padx=(10, 18))
        Label(opt2, text="浓度").pack(side="left")
        self.opacity_combo = RoundCombo(
            opt2, width=11, parent_bg=BG, values=WM_OPACITY)
        self.opacity_combo.current(1)
        self.opacity_combo.pack(side="left", padx=(10, 18))
        Label(opt2, text="字号").pack(side="left")
        self.size_combo = RoundCombo(
            opt2, width=10, parent_bg=BG, values=WM_FONTSIZE)
        self.size_combo.current(1)
        self.size_combo.pack(side="left", padx=(10, 0))
        self._on_type()

        # 输出目录
        out_frame = BgCanvas(self.parent)
        out_frame.pack(fill="x", padx=PX, pady=(10, 0))
        Label(out_frame, text="保存到").pack(side="left")
        self.out_entry = RoundEntry(out_frame, textvariable=self.out_var,
                                    height=36, parent_bg=BG)
        self.out_entry.pack(side="left", fill="x", expand=True, padx=(10, 8))
        self._btn(out_frame, "浏览…", B.BTN_SECONDARY,
                  self._choose_out_dir).pack(side="left")

        # 进度 + 水印按钮
        prog_frame = BgCanvas(self.parent)
        prog_frame.pack(fill="x", padx=PX, pady=(12, 0))
        self.progress = RoundProgress(prog_frame, mode="determinate",
                                      height=14, parent_bg=BG)
        self.progress.pack(side="left", fill="x", expand=True, pady=8)
        self.btn_convert = self._btn(
            prog_frame, "开始加水印", B.BTN_SUCCESS, self._start_convert)
        self.btn_convert.pack(side="right", padx=(14, 0))

        self.status_var = tk.StringVar(value="就绪")
        self.app._make_status_bar(self.parent, self.status_var)

        # 日志（圆角卡片）
        log_card = RoundCard(self.parent, padding=8, height=104, parent_bg=BG)
        log_card.pack(fill="x", padx=PX, pady=(8, 12))
        log_holder = log_card.body
        tk.Label(log_holder, text="转换日志", bg=PALETTE["card"],
                 fg=PALETTE["muted"],
                 font=("Microsoft YaHei UI", 9)).pack(anchor="w", pady=(0, 4))
        log_body = tk.Frame(log_holder, bg=PALETTE["card"])
        log_body.pack(fill="both", expand=True)
        self.log_text = tk.Text(log_body, height=4, state="disabled",
                                wrap="word", font=("Consolas", 9),
                                bd=0, relief="flat", highlightthickness=0,
                                bg=PALETTE["card"], fg=PALETTE["fg"])
        log_sb = Scrollbar(log_body, orient="vertical",
                           command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_sb.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        log_sb.pack(side="right", fill="y")

    # -- 交互 ----------------------------------------------------------------
    def _on_type(self):
        """切换水印类型：文字模式显示文字输入，图片模式显示选择图片。"""
        by_image = self.type_combo.current() == 1
        if by_image:
            self.lbl_text.pack_forget()
            self.text_entry.pack_forget()
            self.btn_img.pack(side="left", padx=(0, 0),
                              before=self.img_label)
        else:
            self.btn_img.pack_forget()
            self.img_label.configure(text="")
            self.lbl_text.pack(side="left", before=self.img_label)
            self.text_entry.pack(side="left", padx=(10, 8),
                                 before=self.img_label)

    def _choose_image(self):
        path = filedialog.askopenfilename(
            title="选择水印图片（建议透明底 PNG）",
            filetypes=[("图片文件", "*.png;*.jpg;*.jpeg;*.bmp;*.gif"),
                       ("所有文件", "*.*")])
        if path:
            self.img_var.set(os.path.normpath(path))
            self.img_label.configure(
                text=os.path.basename(path)[:24])

    def _choose_out_dir(self):
        d = filedialog.askdirectory(title="选择 PDF 保存位置")
        if d:
            self.out_var.set(d)

    # -- 列表 ----------------------------------------------------------------
    def _refresh_tags(self):
        for key, color in STATUS_COLOR.items():
            if color:
                self.tree.tag_configure(key, foreground=color)

    def _refresh_list(self):
        self.tree.delete(*self.tree.get_children())
        for path, state in self.files:
            tags = (state,) if STATUS_COLOR.get(state) else ()
            self.tree.insert("", "end", values=(
                os.path.basename(path), path, STATUS_TEXT.get(state, state)),
                tags=tags)
        if getattr(self, "empty_label", None) is not None:
            if self.files:
                self.empty_label.place_forget()
            else:
                self.empty_label.place(relx=0.5, rely=0.5, anchor="center")
        if self._worker and self._worker.is_alive():
            self.status_var.set(f"正在处理… 共 {len(self.files)} 个文件")
        else:
            self.status_var.set(f"共 {len(self.files)} 个文件")

    def _set_item_state(self, idx, state):
        children = self.tree.get_children()
        if idx >= len(children) or idx >= len(self.files):
            return
        self.files[idx][1] = state
        path = self.files[idx][0]
        self.tree.item(children[idx], values=(
            os.path.basename(path), path, STATUS_TEXT.get(state, state)),
            tags=(state,) if STATUS_COLOR.get(state) else ())

    def _append_paths(self, paths):
        existing = {p for p, _ in self.files}
        added = 0
        for p in paths:
            p = os.path.normpath(p)
            if p.lower().endswith(".pdf") and p not in existing:
                self.files.append([p, "pending"])
                existing.add(p)
                added += 1
        self._refresh_list()
        self._log(f"已添加 {added} 个文件")

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="选择要加水印的 PDF 文件",
            filetypes=[("PDF 文件", "*.pdf"), ("所有文件", "*.*")])
        if paths:
            self._append_paths(paths)

    def _add_folder(self):
        folder = filedialog.askdirectory(title="选择包含 PDF 的文件夹")
        if not folder:
            return
        found = []
        for dp, _dn, fns in os.walk(folder):
            for fn in fns:
                if fn.lower().endswith(".pdf"):
                    found.append(os.path.join(dp, fn))
        if not found:
            messagebox.showinfo("提示", "该文件夹下未找到 PDF 文件。")
            return
        self._append_paths(sorted(found, key=_natural_key))

    def _remove_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        sel_paths = {self.tree.item(i, "values")[1] for i in sel}
        self.files = [f for f in self.files if f[0] not in sel_paths]
        self._refresh_list()

    def _move(self, delta):
        move_selected(self, delta, "文件")

    def _clear_list(self):
        if self._worker and self._worker.is_alive():
            messagebox.showwarning("提示", "正在处理中，无法清空列表。")
            return
        self.files.clear()
        self._refresh_list()

    # -- 拖拽落点注册（由主界面在页面切换时调用） -----------------------------
    def _register_dnd(self):
        """把本页的文件列表注册为拖拽落点，并使其成为 tkdnd 的当前目标。"""
        _register_page_drop(self, self.tree)
        if HAS_DND:
            for w in (self.tree, getattr(self, "empty_label", None)):
                if w is None:
                    continue
                try:
                    w.drop_target_register(tkdnd.DND_FILES)
                    w.dnd_bind("<<Drop>>", self._on_drop)
                except Exception:  # noqa: BLE001
                    pass

    def _unregister_dnd(self):
        """取消本页的拖拽落点注册（页面切走后调用）。"""
        if not (HAS_DND and tkdnd is not None):
            return
        for w in (self.tree, getattr(self, "empty_label", None)):
            if w is None:
                continue
            try:
                w.drop_target_unregister()
            except Exception:  # noqa: BLE001
                pass

    def _on_drop(self, event):
        items = self.root.tk.splitlist(event.data)
        paths = []
        for it in items:
            if os.path.isdir(it):
                for dp, _dn, fns in os.walk(it):
                    for fn in fns:
                        if fn.lower().endswith(".pdf"):
                            paths.append(os.path.join(dp, fn))
            elif it.lower().endswith(".pdf"):
                paths.append(it)
        if paths:
            self._append_paths(sorted(paths, key=_natural_key))
        else:
            messagebox.showinfo("提示", "未识别到 PDF 文件。")

    def _log(self, text):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    # -- 水印流程 --------------------------------------------------------------
    def _start_convert(self):
        if self._worker and self._worker.is_alive():
            self._stop_flag.set()
            self.status_var.set("正在停止…")
            self._log("正在停止…（当前文件处理完后停止）")
            return
        if not self.files:
            messagebox.showwarning("提示", "请先添加要加水印的 PDF 文件。")
            return
        wm_type = "image" if self.type_combo.current() == 1 else "text"
        image = self.img_var.get().strip()
        if wm_type == "image" and not image:
            messagebox.showwarning("提示", "图片水印模式请先「选择图片…」。")
            return
        text = self.text_var.get().strip()
        if wm_type == "text" and not text:
            messagebox.showwarning("提示", "请填写水印文字。")
            return
        tile = self.layout_combo.current() == 0
        color = WM_COLOR_RGB[self.color_combo.get()]
        opacity = WM_OPACITY_VAL[self.opacity_combo.get()]
        fontsize = WM_FONTSIZE_VAL[self.size_combo.get()]

        out_dir = self.out_var.get().strip()
        if not out_dir:
            out_dir = os.path.dirname(self.files[0][0])
        if not os.path.isdir(out_dir):
            os.makedirs(out_dir, exist_ok=True)

        self._stop_flag.clear()
        self.progress["maximum"] = len(self.files)
        self.progress["value"] = 0
        self._log(f"开始加水印（{len(self.files)} 个文件"
                  f"｜{'图片' if wm_type == 'image' else '文字'}"
                  f"｜{'平铺' if tile else '居中'}）…")
        self._log(f"输出：{out_dir}")
        self.btn_convert.configure(text="停止", bootstyle=self.app.BTN_DANGER)

        self._worker = threading.Thread(
            target=self._worker_run,
            args=(list(self.files), out_dir, wm_type, text, image,
                  tile, opacity, fontsize, color), daemon=True)
        self._worker.start()

    def _worker_run(self, files, out_dir, wm_type, text, image,
                     tile, opacity, fontsize, color):
        for i, (path, _state) in enumerate(files):
            if self._stop_flag.is_set():
                self._queue.put(("stopped", (i, len(files))))
                return
            base = os.path.splitext(os.path.basename(path))[0]
            out_pdf = os.path.join(out_dir, f"{base}_水印.pdf")
            k = 1
            while os.path.exists(out_pdf):
                out_pdf = os.path.join(out_dir, f"{base}_水印_{k}.pdf")
                k += 1

            def cb(done, total, _i=i):
                self._queue.put(("file_progress", (_i, done, total)))

            try:
                state, info = add_watermark(
                    path, out_pdf, wm_type=wm_type, text=text, image=image,
                    tile=tile, opacity=opacity, fontsize=fontsize,
                    color=color, stop_flag=self._stop_flag, progress_cb=cb)
            except Exception:  # noqa: BLE001
                self._queue.put(("done", (i, "error", traceback.format_exc())))
                return
            self._queue.put(("done", (i, state, info)))
        self._queue.put(("all_done", len(files)))

    def _poll_queue(self):
        try:
            while True:
                self._handle_msg(self._queue.get_nowait())
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _handle_msg(self, msg):
        kind, payload = msg[0], msg[1]
        if kind == "file_progress":
            fi, done, total = payload
            self.progress["maximum"] = max(len(self.files), 1)
            self.progress["value"] = fi + done / max(total, 1)
            self.status_var.set(f"正在加水印… 第 {fi + 1}/{len(self.files)} 个"
                                f"（{done}/{total} 页）")
        elif kind == "done":
            i, state, info = payload
            if state == "ok":
                self._set_item_state(i, "ok")
            elif state == "skip":
                self._set_item_state(i, "skip")
            else:
                self._set_item_state(i, "error")
                self._log(f"[{os.path.basename(self.files[i][0])}] {info}")
            self.status_var.set(f"已处理 {i + 1}/{len(self.files)}")
        elif kind == "stopped":
            done, total = payload
            self._log(f"已停止：完成 {done}/{total} 个文件")
            self.status_var.set("已停止")
            self._set_convert_btn()
        elif kind == "all_done":
            self.progress["value"] = self.progress["maximum"]
            ok = sum(1 for _p, s in self.files if s == "ok")
            self.status_var.set(f"完成：成功 {ok}/{len(self.files)}")
            self._log(f"全部完成：成功 {ok} 个，失败 "
                      f"{len(self.files) - ok} 个")
            self._set_convert_btn()

    def _set_convert_btn(self):
        self.btn_convert.configure(text="开始加水印",
                                   bootstyle=self.app.BTN_SUCCESS)
        self._worker = None


class PdfToWordTab:
    """PDF 转 Word 功能页：批量把 PDF 转成可编辑的 .docx。

    依赖 pdf2docx（纯 Python，无需本机安装 Office / WPS）。
    """

    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.root = app.root
        self.files = []           # [[完整路径, 状态 key], ...]
        self.out_var = tk.StringVar()
        self._stop_flag = threading.Event()
        self._worker = None
        self._queue = queue.Queue()
        self._build()
        self.root.after(100, self._poll_queue)

    # -- 组件快捷构造 --------------------------------------------------------
    def _btn(self, master, text, style, command=None, **kw):
        return RoundButton(master, text=text, bootstyle=style or "secondary",
                           command=command,
                           parent_bg=kw.pop("parent_bg", PALETTE["bg"]), **kw)

    def _apply_theme(self):
        fg = PALETTE["fg"]
        self.log_text.configure(bg=PALETTE["card"], fg=fg, insertbackground=fg)

    # -- UI -----------------------------------------------------------------
    def _build(self):
        Label = _widget_class("Label")
        Treeview = _widget_class("Treeview")
        Scrollbar = _widget_class("Scrollbar")

        PX = 20
        BG = PALETTE["bg"]
        B = self.app

        # 标题
        header = BgCanvas(self.parent)
        header.pack(fill="x", padx=PX, pady=(12, 6))
        self.app.brand_row(header, "PDF 转 Word",
                           "批量把 PDF 转成可编辑的 Word 文档 · 保留文字与图片")

        # 工具栏
        toolbar = BgCanvas(self.parent)
        toolbar.pack(fill="x", padx=PX, pady=(4, 0))
        self._btn(toolbar, "添加文件", B.BTN_PRIMARY,
                  self._add_files, pad_x=14).pack(side="left", padx=(0, 6))
        self._btn(toolbar, "添加文件夹", B.BTN_SECONDARY,
                  self._add_folder, pad_x=14).pack(side="left", padx=6)
        self._btn(toolbar, "移除选中", B.BTN_SECONDARY,
                  self._remove_selected, pad_x=14).pack(side="left", padx=6)
        self._btn(toolbar, "清空列表", B.BTN_SECONDARY,
                  self._clear_list, pad_x=14).pack(side="left", padx=6)
        self.btn_up = self._btn(toolbar, "上移", B.BTN_SECONDARY,
                                lambda: self._move(-1), pad_x=14)
        self.btn_up.pack(side="left", padx=(14, 6))
        self.btn_down = self._btn(toolbar, "下移", B.BTN_SECONDARY,
                                  lambda: self._move(1), pad_x=14)
        self.btn_down.pack(side="left", padx=6)

        # 文件列表（圆角卡片）
        list_card = RoundCard(self.parent, padding=6, parent_bg=BG)
        list_card.pack(fill="both", expand=True, padx=PX, pady=(10, 0))
        list_frame = list_card.body
        cols = ("name", "path", "status")
        self.tree = Treeview(list_frame, columns=cols, show="headings",
                             selectmode="extended", height=3)
        self.tree.heading("name", text="文件名")
        self.tree.heading("path", text="路径")
        self.tree.heading("status", text="状态")
        self.tree.column("name", width=230, anchor="w")
        self.tree.column("path", width=440, anchor="w")
        self.tree.column("status", width=90, anchor="center", stretch=False)
        vsb = Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self._refresh_tags()

        # 拖拽添加（直接拖到列表上）
        if HAS_DND:
            try:
                self.tree.drop_target_register(tkdnd.DND_FILES)
                self.tree.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass

        # 空状态提示
        self.empty_label = tk.Label(
            list_frame,
            text="还没有添加文件\n点击「添加文件」或把 PDF 文件 / 文件夹拖到列表上",
            justify="center", bg=PALETTE["card"], fg=PALETTE["muted"],
            font=("Microsoft YaHei UI", 11))
        self.empty_label.place(relx=0.5, rely=0.5, anchor="center")
        if HAS_DND:
            # 空状态提示浮在列表中央，必须同样注册为落点，
            # 否则拖到提示文字上的文件会被这个标签吞掉、无法触发 Drop。
            try:
                self.empty_label.drop_target_register(tkdnd.DND_FILES)
                self.empty_label.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:  # noqa: BLE001
                pass

        # 输出目录
        out_frame = BgCanvas(self.parent)
        out_frame.pack(fill="x", padx=PX, pady=(10, 0))
        Label(out_frame, text="保存到").pack(side="left")
        self.out_entry = RoundEntry(out_frame, textvariable=self.out_var,
                                    height=36, parent_bg=BG)
        self.out_entry.pack(side="left", fill="x", expand=True, padx=(10, 8))
        self._btn(out_frame, "浏览…", B.BTN_SECONDARY,
                  self._choose_out_dir).pack(side="left")

        # 进度 + 转换按钮
        prog_frame = BgCanvas(self.parent)
        prog_frame.pack(fill="x", padx=PX, pady=(12, 0))
        self.progress = RoundProgress(prog_frame, mode="determinate",
                                      height=14, parent_bg=BG)
        self.progress.pack(side="left", fill="x", expand=True, pady=8)
        self.btn_convert = self._btn(
            prog_frame, "开始转换", B.BTN_SUCCESS, self._start_convert)
        self.btn_convert.pack(side="right", padx=(14, 0))

        self.status_var = tk.StringVar(value="就绪")
        self.app._make_status_bar(self.parent, self.status_var)

        # 日志（圆角卡片）
        log_card = RoundCard(self.parent, padding=8, height=104, parent_bg=BG)
        log_card.pack(fill="x", padx=PX, pady=(8, 12))
        log_holder = log_card.body
        tk.Label(log_holder, text="转换日志", bg=PALETTE["card"],
                 fg=PALETTE["muted"],
                 font=("Microsoft YaHei UI", 9)).pack(anchor="w", pady=(0, 4))
        log_body = tk.Frame(log_holder, bg=PALETTE["card"])
        log_body.pack(fill="both", expand=True)
        self.log_text = tk.Text(log_body, height=4, state="disabled",
                                wrap="word", font=("Consolas", 9),
                                bd=0, relief="flat", highlightthickness=0,
                                bg=PALETTE["card"], fg=PALETTE["fg"])
        log_sb = Scrollbar(log_body, orient="vertical",
                           command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_sb.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        log_sb.pack(side="right", fill="y")

    # -- 目录 / 列表 ----------------------------------------------------------
    def _choose_out_dir(self):
        d = filedialog.askdirectory(title="选择 Word 保存位置")
        if d:
            self.out_var.set(d)

    def _refresh_tags(self):
        for key, color in STATUS_COLOR.items():
            if color:
                self.tree.tag_configure(key, foreground=color)

    def _refresh_list(self):
        self.tree.delete(*self.tree.get_children())
        for path, state in self.files:
            tags = (state,) if STATUS_COLOR.get(state) else ()
            self.tree.insert("", "end", values=(
                os.path.basename(path), path, STATUS_TEXT.get(state, state)),
                tags=tags)
        if getattr(self, "empty_label", None) is not None:
            if self.files:
                self.empty_label.place_forget()
            else:
                self.empty_label.place(relx=0.5, rely=0.5, anchor="center")
        if self._worker and self._worker.is_alive():
            self.status_var.set(f"正在转换… 共 {len(self.files)} 个文件")
        else:
            self.status_var.set(f"共 {len(self.files)} 个文件")

    def _set_item_state(self, idx, state):
        children = self.tree.get_children()
        if idx >= len(children) or idx >= len(self.files):
            return
        self.files[idx][1] = state
        path = self.files[idx][0]
        self.tree.item(children[idx], values=(
            os.path.basename(path), path, STATUS_TEXT.get(state, state)),
            tags=(state,) if STATUS_COLOR.get(state) else ())

    def _append_paths(self, paths):
        existing = {p for p, _ in self.files}
        added = 0
        for p in paths:
            p = os.path.normpath(p)
            if p.lower().endswith(".pdf") and p not in existing:
                self.files.append([p, "pending"])
                existing.add(p)
                added += 1
        self._refresh_list()
        self._log(f"已添加 {added} 个文件")

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="选择要转换的 PDF 文件",
            filetypes=[("PDF 文件", "*.pdf"), ("所有文件", "*.*")])
        if paths:
            self._append_paths(paths)

    def _add_folder(self):
        folder = filedialog.askdirectory(title="选择包含 PDF 的文件夹")
        if not folder:
            return
        found = []
        for dp, _dn, fns in os.walk(folder):
            for fn in fns:
                if fn.lower().endswith(".pdf"):
                    found.append(os.path.join(dp, fn))
        if not found:
            messagebox.showinfo("提示", "该文件夹下未找到 PDF 文件。")
            return
        self._append_paths(sorted(found, key=_natural_key))

    def _remove_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        sel_paths = {self.tree.item(i, "values")[1] for i in sel}
        self.files = [f for f in self.files if f[0] not in sel_paths]
        self._refresh_list()

    def _move(self, delta):
        move_selected(self, delta, "文件")

    def _clear_list(self):
        if self._worker and self._worker.is_alive():
            messagebox.showwarning("提示", "正在转换中，无法清空列表。")
            return
        self.files.clear()
        self._refresh_list()
        self.progress["value"] = 0

    # -- 拖拽落点注册（由主界面在页面切换时调用） -----------------------------
    def _register_dnd(self):
        """把本页的文件列表注册为拖拽落点，并使其成为 tkdnd 的当前目标。"""
        _register_page_drop(self, self.tree)
        if HAS_DND:
            for w in (self.tree, getattr(self, "empty_label", None)):
                if w is None:
                    continue
                try:
                    w.drop_target_register(tkdnd.DND_FILES)
                    w.dnd_bind("<<Drop>>", self._on_drop)
                except Exception:  # noqa: BLE001
                    pass

    def _unregister_dnd(self):
        """取消本页的拖拽落点注册（页面切走后调用）。"""
        if not (HAS_DND and tkdnd is not None):
            return
        for w in (self.tree, getattr(self, "empty_label", None)):
            if w is None:
                continue
            try:
                w.drop_target_unregister()
            except Exception:  # noqa: BLE001
                pass

    def _on_drop(self, event):
        items = self.root.tk.splitlist(event.data)
        paths = []
        for it in items:
            if os.path.isdir(it):
                for dp, _dn, fns in os.walk(it):
                    for fn in fns:
                        if fn.lower().endswith(".pdf"):
                            paths.append(os.path.join(dp, fn))
            elif it.lower().endswith(".pdf"):
                paths.append(it)
        if paths:
            self._append_paths(sorted(paths, key=_natural_key))
        else:
            messagebox.showinfo("提示", "未识别到 PDF 文件。")

    def _log(self, text):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    # -- 转换流程 --------------------------------------------------------------
    def _start_convert(self):
        if self._worker and self._worker.is_alive():
            self._stop_flag.set()
            self.status_var.set("正在停止…")
            self._log("正在停止…（当前文件处理完后停止）")
            return
        if not self.files:
            messagebox.showwarning("提示", "请先添加要转换的 PDF 文件。")
            return

        out_dir = self.out_var.get().strip()
        if not out_dir:
            out_dir = os.path.dirname(self.files[0][0])
        if not os.path.isdir(out_dir):
            os.makedirs(out_dir, exist_ok=True)

        self._stop_flag.clear()
        self.progress["maximum"] = len(self.files)
        self.progress["value"] = 0
        self._log(f"开始转换（{len(self.files)} 个文件）…")
        self._log(f"输出：{out_dir}")
        self.btn_convert.configure(text="停止", bootstyle=self.app.BTN_DANGER)

        self._worker = threading.Thread(
            target=self._worker_run,
            args=(list(self.files), out_dir), daemon=True)
        self._worker.start()

    def _worker_run(self, files, out_dir):
        for i, (path, _state) in enumerate(files):
            if self._stop_flag.is_set():
                self._queue.put(("stopped", (i, len(files))))
                return
            try:
                state, info = convert_pdf_to_word(
                    path, out_dir, stop_flag=self._stop_flag)
            except Exception:  # noqa: BLE001
                self._queue.put(("done", (i, "error", traceback.format_exc())))
                return
            self._queue.put(("done", (i, state, info)))
        self._queue.put(("all_done", len(files)))

    def _poll_queue(self):
        try:
            while True:
                self._handle_msg(self._queue.get_nowait())
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _handle_msg(self, msg):
        kind, payload = msg[0], msg[1]
        if kind == "done":
            i, state, info = payload
            if state == "ok":
                self._set_item_state(i, "ok")
                self._log(info)
            elif state == "skip":
                self._set_item_state(i, "skip")
                self._log(info)
            else:
                self._set_item_state(i, "error")
                self._log(f"[{os.path.basename(self.files[i][0])}] {info}")
            self.status_var.set(f"已处理 {i + 1}/{len(self.files)}")
            self.progress["value"] = i + 1
        elif kind == "stopped":
            done, total = payload
            self._log(f"已停止：完成 {done}/{total} 个文件")
            self.status_var.set("已停止")
            self._set_convert_btn()
        elif kind == "all_done":
            self.progress["value"] = self.progress["maximum"]
            ok = sum(1 for _p, s in self.files if s == "ok")
            self.status_var.set(f"完成：成功 {ok}/{len(self.files)}")
            self._log(f"全部完成：成功 {ok} 个，失败 "
                      f"{len(self.files) - ok} 个")
            self._set_convert_btn()

    def _set_convert_btn(self):
        self.btn_convert.configure(text="开始转换",
                                   bootstyle=self.app.BTN_SUCCESS)
        self._worker = None


class PptToPdfTab:
    """PPT 转 PDF 功能页：批量导出，依赖本机 Microsoft PowerPoint。"""

    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.root = app.root
        self.files = []           # [[完整路径, 状态 key], ...]
        self.out_var = tk.StringVar()
        self._stop_flag = threading.Event()
        self._worker = None
        self._queue = queue.Queue()
        self._build()
        self.root.after(100, self._poll_queue)

    # -- 组件快捷构造 --------------------------------------------------------
    def _btn(self, master, text, style, command=None, **kw):
        return RoundButton(master, text=text, bootstyle=style or "secondary",
                           command=command,
                           parent_bg=kw.pop("parent_bg", PALETTE["bg"]), **kw)

    def _apply_theme(self):
        fg = PALETTE["fg"]
        self.log_text.configure(bg=PALETTE["card"], fg=fg, insertbackground=fg)

    # -- UI -----------------------------------------------------------------
    def _build(self):
        Label = _widget_class("Label")
        Frame = _widget_class("Frame")
        Treeview = _widget_class("Treeview")
        Scrollbar = _widget_class("Scrollbar")

        PX = 20
        BG = PALETTE["bg"]
        B = self.app

        # 标题
        header = BgCanvas(self.parent)
        header.pack(fill="x", padx=PX, pady=(12, 6))
        self.app.brand_row(header, "PPT 转 PDF",
                           "批量把 PPT 演示文稿导出为 PDF · 需本机安装 Microsoft PowerPoint 或 WPS")

        # 工具栏
        toolbar = BgCanvas(self.parent)
        toolbar.pack(fill="x", padx=PX, pady=(4, 0))
        self._btn(toolbar, "添加文件", B.BTN_PRIMARY,
                  self._add_files, pad_x=14).pack(side="left", padx=(0, 6))
        self._btn(toolbar, "添加文件夹", B.BTN_SECONDARY,
                  self._add_folder, pad_x=14).pack(side="left", padx=6)
        self._btn(toolbar, "移除选中", B.BTN_SECONDARY,
                  self._remove_selected, pad_x=14).pack(side="left", padx=6)
        self._btn(toolbar, "清空列表", B.BTN_SECONDARY,
                  self._clear_list, pad_x=14).pack(side="left", padx=6)
        self.btn_up = self._btn(toolbar, "上移", B.BTN_SECONDARY,
                                lambda: self._move(-1), pad_x=14)
        self.btn_up.pack(side="left", padx=(14, 6))
        self.btn_down = self._btn(toolbar, "下移", B.BTN_SECONDARY,
                                  lambda: self._move(1), pad_x=14)
        self.btn_down.pack(side="left", padx=6)

        # 文件列表（圆角卡片）
        list_card = RoundCard(self.parent, padding=6, parent_bg=BG)
        list_card.pack(fill="both", expand=True, padx=PX, pady=(10, 0))
        list_frame = list_card.body
        cols = ("idx", "name", "size", "status")
        self.tree = Treeview(list_frame, columns=cols, show="headings",
                             selectmode="extended", height=3)
        self.tree.heading("idx", text="顺序")
        self.tree.heading("name", text="PPT 文件")
        self.tree.heading("size", text="大小")
        self.tree.heading("status", text="状态")
        self.tree.column("idx", width=52, anchor="center", stretch=False)
        self.tree.column("name", width=390, anchor="w")
        self.tree.column("size", width=90, anchor="e", stretch=False)
        self.tree.column("status", width=90, anchor="center", stretch=False)
        vsb = Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self._refresh_tags()

        # 拖拽添加（直接拖到列表上）
        if HAS_DND:
            try:
                self.tree.drop_target_register(tkdnd.DND_FILES)
                self.tree.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass

        # 空状态提示（叠加在列表中央，提示可拖拽）
        self.empty_label = tk.Label(
            list_frame,
            text="还没有添加文件\n点击「添加文件」或把 PPT 文件 / 文件夹拖到列表上",
            justify="center", bg=PALETTE["card"], fg=PALETTE["muted"],
            font=("Microsoft YaHei UI", 11))
        self.empty_label.place(relx=0.5, rely=0.5, anchor="center")
        if HAS_DND:
            # 空状态提示浮在列表中央，必须同样注册为落点，
            # 否则拖到提示文字上的文件会被这个标签吞掉、无法触发 Drop。
            try:
                self.empty_label.drop_target_register(tkdnd.DND_FILES)
                self.empty_label.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:  # noqa: BLE001
                pass

        # 输出目录
        out_frame = BgCanvas(self.parent)
        out_frame.pack(fill="x", padx=PX, pady=(10, 0))
        Label(out_frame, text="保存到").pack(side="left")
        self.out_entry = RoundEntry(out_frame, textvariable=self.out_var,
                                    height=36, parent_bg=BG)
        self.out_entry.pack(side="left", fill="x", expand=True, padx=(10, 8))
        self._btn(out_frame, "浏览…", B.BTN_SECONDARY,
                  self._choose_out_dir).pack(side="left")

        # 进度 + 转换按钮
        prog_frame = BgCanvas(self.parent)
        prog_frame.pack(fill="x", padx=PX, pady=(12, 0))
        self.progress = RoundProgress(prog_frame, mode="determinate",
                                      height=14, parent_bg=BG)
        self.progress.pack(side="left", fill="x", expand=True, pady=8)
        self.btn_convert = self._btn(
            prog_frame, "开始转换", B.BTN_SUCCESS, self._start_convert)
        self.btn_convert.pack(side="right", padx=(14, 0))

        self.status_var = tk.StringVar(value="就绪")
        self.app._make_status_bar(self.parent, self.status_var)

        # 日志（圆角卡片）
        log_card = RoundCard(self.parent, padding=8, height=104, parent_bg=BG)
        log_card.pack(fill="x", padx=PX, pady=(8, 12))
        log_holder = log_card.body
        tk.Label(log_holder, text="转换日志", bg=PALETTE["card"],
                 fg=PALETTE["muted"],
                 font=("Microsoft YaHei UI", 9)).pack(anchor="w", pady=(0, 4))
        log_body = tk.Frame(log_holder, bg=PALETTE["card"])
        log_body.pack(fill="both", expand=True)
        self.log_text = tk.Text(log_body, height=4, state="disabled",
                                wrap="word", font=("Consolas", 9),
                                bd=0, relief="flat", highlightthickness=0,
                                bg=PALETTE["card"], fg=PALETTE["fg"])
        log_sb = Scrollbar(log_body, orient="vertical",
                           command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_sb.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        log_sb.pack(side="right", fill="y")

    # -- 列表 ----------------------------------------------------------------
    def _refresh_tags(self):
        for key, color in STATUS_COLOR.items():
            if color:
                self.tree.tag_configure(key, foreground=color)

    def _refresh_list(self):
        self.tree.delete(*self.tree.get_children())
        for i, (path, state) in enumerate(self.files, 1):
            try:
                size = _human_size(os.path.getsize(path))
            except OSError:
                size = "—"
            tags = (state,) if STATUS_COLOR.get(state) else ()
            self.tree.insert("", "end", values=(
                i, os.path.basename(path), size,
                STATUS_TEXT.get(state, state)), tags=tags)
        if getattr(self, "empty_label", None) is not None:
            if self.files:
                self.empty_label.place_forget()
            else:
                self.empty_label.place(relx=0.5, rely=0.5, anchor="center")
        if self._worker and self._worker.is_alive():
            self.status_var.set(f"正在转换… 共 {len(self.files)} 个")
        else:
            self.status_var.set(f"共 {len(self.files)} 个 PPT 文件")

    def _set_item_state(self, idx, state):
        children = self.tree.get_children()
        if idx >= len(children) or idx >= len(self.files):
            return
        self.files[idx][1] = state
        path = self.files[idx][0]
        try:
            size = _human_size(os.path.getsize(path))
        except OSError:
            size = "—"
        self.tree.item(children[idx], values=(
            idx + 1, os.path.basename(path), size,
            STATUS_TEXT.get(state, state)),
            tags=(state,) if STATUS_COLOR.get(state) else ())

    # -- 文件操作 ------------------------------------------------------------
    def _append_paths(self, paths):
        existing = {p for p, _ in self.files}
        added = 0
        for p in paths:
            p = os.path.normpath(p)
            if p.lower().endswith(PPT_EXTS) and p not in existing:
                self.files.append([p, "pending"])
                existing.add(p)
                added += 1
        self._refresh_list()
        if added:
            self._log(f"已添加 {added} 个 PPT 文件")

    def _add_files(self):
        pats = " ".join("*" + e for e in PPT_EXTS)
        paths = filedialog.askopenfilenames(
            title="选择 PPT 文件（Ctrl / Shift 可多选）",
            filetypes=[("PowerPoint", pats), ("所有文件", "*.*")])
        if paths:
            self._append_paths(list(paths))

    def _add_folder(self):
        folder = filedialog.askdirectory(
            title="选择包含 PPT 的文件夹（自动扫描所有子目录）")
        if not folder:
            return
        found = []
        for dp, _dn, fns in os.walk(folder):
            for fn in fns:
                if fn.lower().endswith(PPT_EXTS):
                    found.append(os.path.join(dp, fn))
        if not found:
            messagebox.showinfo("提示", "该文件夹下未找到 PPT 文件。")
            return
        found.sort(key=_natural_key)
        self._append_paths(found)

    def _remove_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        idxs = sorted((self.tree.index(i) for i in sel), reverse=True)
        for i in idxs:
            if 0 <= i < len(self.files):
                self.files.pop(i)
        self._refresh_list()

    def _move(self, delta):
        move_selected(self, delta, "文件")

    def _clear_list(self):
        if self._worker and self._worker.is_alive():
            messagebox.showwarning("提示", "正在转换中，无法清空列表。")
            return
        self.files.clear()
        self._refresh_list()
        self.progress["value"] = 0

    def _choose_out_dir(self):
        folder = filedialog.askdirectory(title="选择 PDF 保存文件夹")
        if folder:
            self.out_var.set(folder)

    # -- 拖拽落点注册（由主界面在页面切换时调用） -----------------------------
    def _register_dnd(self):
        """把本页的文件列表注册为拖拽落点，并使其成为 tkdnd 的当前目标。"""
        _register_page_drop(self, self.tree)
        if HAS_DND:
            for w in (self.tree, getattr(self, "empty_label", None)):
                if w is None:
                    continue
                try:
                    w.drop_target_register(tkdnd.DND_FILES)
                    w.dnd_bind("<<Drop>>", self._on_drop)
                except Exception:  # noqa: BLE001
                    pass

    def _unregister_dnd(self):
        """取消本页的拖拽落点注册（页面切走后调用）。"""
        if not (HAS_DND and tkdnd is not None):
            return
        for w in (self.tree, getattr(self, "empty_label", None)):
            if w is None:
                continue
            try:
                w.drop_target_unregister()
            except Exception:  # noqa: BLE001
                pass

    def _on_drop(self, event):
        items = self.root.tk.splitlist(event.data)
        paths = []
        for it in items:
            if os.path.isdir(it):
                for dp, _dn, fns in os.walk(it):
                    for fn in fns:
                        if fn.lower().endswith(PPT_EXTS):
                            paths.append(os.path.join(dp, fn))
            elif os.path.isfile(it) and it.lower().endswith(PPT_EXTS):
                paths.append(it)
        if paths:
            self._append_paths(paths)
        else:
            self._log("拖入的内容里没有可用的 PPT 文件")

    # -- 转换流程 ------------------------------------------------------------
    def _start_convert(self):
        if self._worker and self._worker.is_alive():
            self._stop_flag.set()
            self.status_var.set("正在停止…")
            return
        if not self.files:
            messagebox.showinfo("提示", "请先添加要转换的 PPT 文件。")
            return

        out_dir = self.out_var.get().strip()
        if not out_dir:
            messagebox.showinfo("提示", "请选择 PDF 保存位置。")
            return
        if not os.path.isdir(out_dir):
            messagebox.showwarning("提示", "保存目录不存在，请重新选择。")
            return

        paths = [p for p, _ in self.files]
        self._stop_flag.clear()
        self.progress["maximum"] = len(paths)
        self.progress["value"] = 0
        for i in range(len(self.files)):
            self.files[i][1] = "pending"
        self._refresh_list()

        self._log(f"开始批量转换 {len(paths)} 个 PPT…")
        self._log(f"输出：{out_dir}")
        self._set_convert_btn("停止", self.app.BTN_DANGER)

        self._worker = threading.Thread(
            target=self._worker_run,
            args=(paths, out_dir), daemon=True)
        self._worker.start()

    def _worker_run(self, paths, out_dir):
        pythoncom.CoInitialize()
        ok_count = err_count = skip_count = 0
        try:
            for idx, path in enumerate(paths):
                if self._stop_flag.is_set():
                    self._queue.put(("stopped", None))
                    break
                self._queue.put(("status", (idx, "working")))
                state, info = convert_ppt_to_pdf(
                    path, out_dir, stop_flag=self._stop_flag)
                if state == "ok":
                    ok_count += 1
                elif state == "error":
                    err_count += 1
                else:
                    skip_count += 1
                self._queue.put(("item", (idx, state, info, path)))
            else:
                self._queue.put(("done", (ok_count, err_count, skip_count)))
        except Exception:  # noqa: BLE001
            self._queue.put(("fatal", traceback.format_exc()))
        finally:
            pythoncom.CoUninitialize()

    def _set_convert_btn(self, text, style):
        if HAS_TTB:
            self.btn_convert.configure(text=text, bootstyle=style)
        else:
            self.btn_convert.configure(text=text)

    def _poll_queue(self):
        try:
            while True:
                self._handle_msg(self._queue.get_nowait())
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _handle_msg(self, msg):
        kind, payload = msg[0], msg[1]
        if kind == "status":
            idx, state = payload
            self._set_item_state(idx, state)
        elif kind == "item":
            idx, state, info, _path = payload
            self.progress["value"] = idx + 1
            self.status_var.set(
                f"正在转换… {idx + 1}/{self.progress['maximum']}")
            self._set_item_state(idx, state)
            self._log(info)
        elif kind == "done":
            ok, err, skip = payload
            self.progress["value"] = self.progress["maximum"]
            self.status_var.set(f"完成：成功 {ok} · 失败 {err} · 跳过 {skip}")
            self._log(f"转换结束：成功 {ok} 个，失败 {err} 个，跳过 {skip} 个")
            self._set_convert_btn("开始转换", self.app.BTN_SUCCESS)
            self._worker = None
        elif kind == "stopped":
            self.status_var.set("已停止")
            self._log("转换已手动停止")
            self._set_convert_btn("开始转换", self.app.BTN_SUCCESS)
            self._worker = None
        elif kind == "fatal":
            self._log("发生严重错误：\n" + payload)
            self._set_convert_btn("开始转换", self.app.BTN_SUCCESS)
            self._worker = None

    def _log(self, text):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")


def main():
    # 隐藏自检模式：`--selftest <源文件> <输出目录>` 无界面执行一次转换，
    # 结果写入输出目录下的 _selftest_result.txt，用于打包后的功能验证。
    if "--selftest" in sys.argv:
        try:
            i = sys.argv.index("--selftest")
            src, out = sys.argv[i + 1], sys.argv[i + 2]
            state, info = convert_doc_to_pdf(src, out)
            with open(os.path.join(out, "_selftest_result.txt"),
                      "w", encoding="utf-8") as f:
                f.write(f"state={state}\ninfo={info}\n")
        except Exception:  # noqa: BLE001
            pass
        return

    # 隐藏自检模式（图片转 PDF）：`--selftest-images <输出.pdf> <图片...>`
    if "--selftest-images" in sys.argv:
        state, info, out_pdf = "error", "参数不足", ""
        try:
            i = sys.argv.index("--selftest-images")
            out_pdf = os.path.abspath(sys.argv[i + 1])
            state, info = images_to_pdf(sys.argv[i + 2:], out_pdf, "auto", 0)
        except Exception as err:  # noqa: BLE001
            info = str(err)
        try:
            with open(os.path.join(os.path.dirname(out_pdf),
                                   "_selftest_result.txt"),
                      "w", encoding="utf-8") as f:
                f.write(f"state={state}\ninfo={info}\n")
        except Exception:  # noqa: BLE001
            pass
        return

    # 隐藏自检模式（PDF 转图片）：`--selftest-pdf2img <pdf> <输出目录>`
    if "--selftest-pdf2img" in sys.argv:
        try:
            i = sys.argv.index("--selftest-pdf2img")
            src, out = sys.argv[i + 1], sys.argv[i + 2]
            state, info = convert_pdf_to_images(src, out, 216 / 72.0, "png")
            with open(os.path.join(out, "_selftest_result.txt"),
                      "w", encoding="utf-8") as f:
                f.write(f"state={state}\ninfo={info}\n")
        except Exception:  # noqa: BLE001
            pass
        return

    # 隐藏自检模式（PDF 转 Word）：`--selftest-pdf2word <pdf> <输出目录>`
    if "--selftest-pdf2word" in sys.argv:
        state, info = "error", "参数不足"
        try:
            i = sys.argv.index("--selftest-pdf2word")
            src, out = sys.argv[i + 1], sys.argv[i + 2]
            state, info = convert_pdf_to_word(src, out)
        except Exception as err:  # noqa: BLE001
            info = str(err)
        try:
            with open(os.path.join(out, "_selftest_result.txt"),
                      "w", encoding="utf-8") as f:
                f.write(f"state={state}\ninfo={info}\n")
        except Exception:  # noqa: BLE001
            pass
        return

    # 隐藏自检模式（PDF 合并）：`--selftest-merge <输出.pdf> <pdf...>`
    if "--selftest-merge" in sys.argv:
        state, info, out_pdf = "error", "参数不足", ""
        try:
            i = sys.argv.index("--selftest-merge")
            out_pdf = os.path.abspath(sys.argv[i + 1])
            state, info = merge_pdfs(sys.argv[i + 2:], out_pdf)
        except Exception as err:  # noqa: BLE001
            info = str(err)
        try:
            with open(os.path.join(os.path.dirname(out_pdf),
                                   "_selftest_result.txt"),
                      "w", encoding="utf-8") as f:
                f.write(f"state={state}\ninfo={info}\n")
        except Exception:  # noqa: BLE001
            pass
        return

    # 隐藏自检模式（PPT 转 PDF）：`--selftest-ppt <ppt文件> <输出目录>`
    if "--selftest-ppt" in sys.argv:
        state, info = "error", "参数不足"
        try:
            i = sys.argv.index("--selftest-ppt")
            src, out = sys.argv[i + 1], sys.argv[i + 2]
            state, info = convert_ppt_to_pdf(src, out)
        except Exception as err:  # noqa: BLE001
            info = str(err)
        try:
            with open(os.path.join(out, "_selftest_result.txt"),
                      "w", encoding="utf-8") as f:
                f.write(f"state={state}\ninfo={info}\n")
        except Exception:  # noqa: BLE001
            pass
        return

    # 隐藏自检模式（PDF 拆分）：`--selftest-split <pdf> <输出目录> [范围]`
    if "--selftest-split" in sys.argv:
        state, info = "error", "参数不足"
        try:
            i = sys.argv.index("--selftest-split")
            src, out = sys.argv[i + 1], sys.argv[i + 2]
            rng = sys.argv[i + 3] if len(sys.argv) > i + 3 else None
            state, info = split_pdf(src, out,
                                    parse_page_ranges(rng) if rng else None)
        except Exception as err:  # noqa: BLE001
            info = str(err)
        try:
            with open(os.path.join(out, "_selftest_result.txt"),
                      "w", encoding="utf-8") as f:
                f.write(f"state={state}\ninfo={info}\n")
        except Exception:  # noqa: BLE001
            pass
        return

    # 隐藏自检模式（PDF 加水印）：`--selftest-watermark <pdf> <输出pdf> [图片路径]`
    if "--selftest-watermark" in sys.argv:
        state, info = "error", "参数不足"
        try:
            i = sys.argv.index("--selftest-watermark")
            src, out = sys.argv[i + 1], sys.argv[i + 2]
            img = sys.argv[i + 3] if len(sys.argv) > i + 3 else None
            if img:
                state, info = add_watermark(src, out, wm_type="image",
                                            image=img, tile=True,
                                            opacity=0.3)
            else:
                state, info = add_watermark(src, out, wm_type="text",
                                            text="内部资料", tile=True,
                                            opacity=0.15, fontsize=40)
        except Exception as err:  # noqa: BLE001
            info = str(err)
        try:
            with open(os.path.join(os.path.dirname(out) or ".",
                                   "_selftest_result.txt"),
                      "w", encoding="utf-8") as f:
                f.write(f"state={state}\ninfo={info}\n")
        except Exception:  # noqa: BLE001
            pass
        return

    if not HAS_WIN32:
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(
                "缺少组件",
                "未检测到 pywin32 组件。\n\n"
                "请先安装：pip install pywin32\n然后重新运行本程序。")
            root.destroy()
        except Exception:
            pass
        sys.exit(1)

    app = Word2PDFApp()
    app.run()


if __name__ == "__main__":
    main()
