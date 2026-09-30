"""WorkAid 主题与基础设施：DPI 缩放、调色板、DWM 标题栏、绘图工具。"""

import math
import os
import re
import sys
import ctypes
import datetime
import tkinter as tk
from tkinter import ttk
from tkinter import font as tkfont

# 运行环境兼容：优先用 ttkbootstrap 做现代化界面，缺失时退回标准 ttk
try:
    import ttkbootstrap as tb
    HAS_TTB = True
except Exception:
    tb = None
    HAS_TTB = False

__all__ = [
    "APP_NAME",
    "APP_VERSION",
    "HAS_TTB",
    "tb",
    "CARD_RADIUS",
    "COPYRIGHT_AUTHOR",
    "COPYRIGHT_START_YEAR",
    "CTRL_RADIUS",
    "FEEDBACK_EMAIL",
    "PALETTE",
    "SUPPORTED_EXT",
    "SUPPORT_QR_FILE",
    "UI_FONT",
    "UI_FONT_SIZE",
    "UI_SCALE",
    "_ARC_QUARTER",
    "_DROP_ICON_CACHE",
    "_apply_titlebar_colors",
    "_clear_titlebar_icon",
    "_drop_icon_photo",
    "_enable_dpi_awareness",
    "_hex_to_rgb",
    "_load_scaled_photo",
    "_mix",
    "_resource_path",
    "_rgb_to_hex",
    "_shade",
    "_widget_class",
    "copyright_text",
    "render_bg_photo",
    "round_rect",
    "s",
]


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



def _clear_titlebar_icon(win):
    """隐藏窗口标题栏左侧的小图标（Windows；失败静默）。

    Win32 没有"不显示图标"的开关：把 WM_SETICON 置空只会回退到
    窗口类默认图标（Tk 羽毛）。这里改为设置 1×1 全透明图标，
    视觉上等于不显示。
    """
    if sys.platform != "win32":
        return
    try:
        user32 = ctypes.windll.user32
        # tkinter 顶层窗口真正的 HWND 是内部子窗口的父窗口
        hwnd = user32.GetParent(win.winfo_id()) or win.winfo_id()
        # 1×1 单色图标，AND 掩码全 1 → 完全透明
        hicon = user32.CreateIcon(None, 1, 1, 1, 1,
                                  b"\xff\xff", b"\xff\xff")
        WM_SETICON = 0x0080
        for wparam in (1, 0):              # ICON_BIG / ICON_SMALL
            user32.SendMessageW(hwnd, WM_SETICON, wparam, hicon)
        SWP = (0x0001 | 0x0002 | 0x0004 | 0x0020)   # NOSIZE|NOMOVE|NOZORDER|FRAMECHANGED
        user32.SetWindowPos(hwnd, None, 0, 0, 0, 0, SWP)
    except Exception:
        pass


APP_NAME = "WorkAid"
APP_VERSION = "1.13.0"


# 版权与反馈信息（起始年份固定，结束年份自动取当前系统年份）
COPYRIGHT_START_YEAR = 2026

COPYRIGHT_AUTHOR = "hongyongzhong"

FEEDBACK_EMAIL = "hongyongzhong2020@126.com"

# 微信赞赏码图片（放 app 目录 / 打进资源根目录即生效；缺失时弹窗降级为文字提示）
SUPPORT_QR_FILE = "donate_qr.png"



def copyright_text():
    """底部版权文字；第二个年份自动取当天所在年份。"""
    return f"© {COPYRIGHT_START_YEAR}-{datetime.datetime.now().year} " \
           f"{COPYRIGHT_AUTHOR}   Email: {FEEDBACK_EMAIL}"


# 支持的文件后缀
SUPPORTED_EXT = (".doc", ".docx", ".docm", ".rtf", ".wps")



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



# 拖拽区箭头图标缓存（PhotoImage 必须持引用，否则被 GC 后显示空白）
_DROP_ICON_CACHE = {}



def _drop_icon_photo():
    """加载拖拽区箭头图标（icons/drop.png，圆底白箭头），失败返回 None。"""
    if "icon" not in _DROP_ICON_CACHE:
        try:
            _DROP_ICON_CACHE["icon"] = _load_scaled_photo(
                os.path.join("icons", "drop.png"))
        except Exception:  # noqa: BLE001
            _DROP_ICON_CACHE["icon"] = None
    return _DROP_ICON_CACHE["icon"]



def render_bg_photo(master, w, h):
    """生成纯色背景图（深蓝夜色，无纹理）。"""
    photo = tk.PhotoImage(width=w, height=h, master=master)
    photo.put(PALETTE["bg"], to=(0, 0, w, h))
    return photo
