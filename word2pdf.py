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

# 拆分后的内部模块（theme/widgets/converters/pages）
from theme import *
from widgets import *
from converters import *
from pages import *

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
        self.rn_tab = BgCanvas(self.notebook.body)
        self.cv_tab = BgCanvas(self.notebook.body)
        self.notebook.add(self.word_tab, text="Word 转 PDF")
        self.notebook.add(self.pw_tab, text="PDF 转 Word")
        self.notebook.add(self.ppt_tab, text="PPT 转 PDF")
        self.notebook.add(self.img_tab, text="PDF 转图片")
        self.notebook.add(self.wm_tab, text="PDF 加水印")
        self.notebook.add(self.merge_tab, text="PDF 合并")
        self.notebook.add(self.split_tab, text="PDF 拆分")
        self.notebook.add(self.pdf_tab, text="图片转 PDF")
        self.notebook.add(self.rn_tab, text="图片重命名")
        self.notebook.add(self.cv_tab, text="图片格式转换")

        self._build_ui()          # 构建 Word 转 PDF 页
        self.img_page = PdfToImagesTab(self.img_tab, self)   # PDF 转图片页
        self.pdf_page = ImagesToPdfTab(self.pdf_tab, self)   # 图片转 PDF 页
        self.merge_page = MergePdfsTab(self.merge_tab, self)  # PDF 合并页
        self.split_page = SplitPdfTab(self.split_tab, self)  # PDF 拆分页
        self.ppt_page = PptToPdfTab(self.ppt_tab, self)      # PPT 转 PDF 页
        self.wm_page = WatermarkTab(self.wm_tab, self)       # PDF 加水印页
        self.pw_page = PdfToWordTab(self.pw_tab, self)       # PDF 转 Word 页
        self.rn_page = ImageRenameTab(self.rn_tab, self)     # 图片重命名页
        self.cv_page = ImageConvertTab(self.cv_tab, self)    # 图片格式转换页
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
            str(self.rn_tab): self.rn_page,
            str(self.cv_tab): self.cv_page,
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
        kw.setdefault("parent_bg", PALETTE["bg"])
        return RoundButton(master, text=text, bootstyle=style or "secondary",
                           command=command, **kw)

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
        # 右下角"支持作者"入口：点击打开打赏页，悬停高亮
        sup = f.create_text(w - 16, h / 2, text="☕ 支持作者", anchor="e",
                            fill=PALETTE["muted"],
                            font=("Microsoft YaHei UI", 9),
                            tags=("footer", "support"))
        f.tag_bind("support", "<Button-1>", self._open_support)
        f.tag_bind("support", "<Enter>",
                   lambda _e: f.itemconfigure(sup, fill=PALETTE["fg"]))
        f.tag_bind("support", "<Leave>",
                   lambda _e: f.itemconfigure(sup, fill=PALETTE["muted"]))

    def _load_qr_photo(self):
        """加载赞赏码并缩放到弹窗显示尺寸（逻辑 320px），失败返回 None。"""
        try:
            from PIL import Image as _PILImage, ImageTk as _PILTk
            path = _resource_path(SUPPORT_QR_FILE)
            target = s(160)
            with _PILImage.open(path) as pil:
                pil = pil.convert("RGBA")
                w, h = pil.size
                scale = target / float(max(w, h))
                pil = pil.resize((max(1, round(w * scale)),
                                  max(1, round(h * scale))), _PILImage.LANCZOS)
                return _PILTk.PhotoImage(pil)
        except Exception:  # noqa: BLE001
            return None

    def _open_support(self, _event=None):
        """弹出"支持作者"窗口：微信赞赏码扫码打赏（纯本地，无外网依赖）。"""
        qr = self._load_qr_photo()
        self._qr_photo = qr   # 持有引用，防止 ImageTk.PhotoImage 被 GC 后图像空白
        win = tk.Toplevel(self.root)
        win.title("支持作者")
        win.configure(bg=PALETTE["card"])
        win.resizable(False, False)
        win.transient(self.root)
        win.withdraw()   # 先隐藏，清完标题栏图标再显示，避免图标闪现
        # 客户区顶部自绘一行居中的"支持作者"；系统标题栏文字则染成与
        # 标题栏同色（隐形，须在窗口可见后设置）——两块同色区域无缝，
        # 视觉上标题文字居中
        title_row = tk.Frame(win, bg=PALETTE["card"], height=s(32))
        title_row.pack(fill="x")
        title_row.pack_propagate(False)
        tk.Label(title_row, text="支持作者", bg=PALETTE["card"],
                 fg=PALETTE["fg"],
                 font=("Microsoft YaHei UI", 9)).place(
                     relx=0.5, rely=0.5, anchor="center")
        body = tk.Frame(win, bg=PALETTE["card"], padx=28, pady=22)
        body.pack(fill="both", expand=True)
        tk.Label(body, text="如果 WorkAid 帮到了你\n欢迎请作者喝杯咖啡 ☕",
                 justify="center", bg=PALETTE["card"], fg=PALETTE["fg"],
                 font=("Microsoft YaHei UI", 12, "bold")).pack()
        if qr is not None:
            tk.Label(body, image=qr, bg=PALETTE["card"], bd=0).pack(pady=(12, 6))
            tk.Label(body, text="微信扫一扫 · 金额随意",
                     bg=PALETTE["card"], fg=PALETTE["muted"],
                     font=("Microsoft YaHei UI", 9)).pack()
        else:
            tk.Label(body, text="赞赏码即将上线\n也可以通过邮件联系作者：" + FEEDBACK_EMAIL,
                     justify="center", bg=PALETTE["card"], fg=PALETTE["muted"],
                     font=("Microsoft YaHei UI", 10)).pack(pady=(14, 10))
        # "关闭"按钮与弹窗背景同色（幽灵按钮）：悬停/按下仍有明暗反馈
        self._button(win, "关 闭", self.BTN_SECONDARY,
                     win.destroy, pad_x=22, fill=PALETTE["card"],
                     parent_bg=PALETTE["card"]).pack(pady=(4, 0))
        # 居中于主窗口；显示前清掉标题栏左侧图标（不显示 app.ico）
        win.update_idletasks()
        ww, wh = win.winfo_reqwidth(), win.winfo_reqheight()
        self.root.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - ww) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - wh) // 3
        win.geometry("+%d+%d" % (max(x, 0), max(y, 0)))
        _clear_titlebar_icon(win)
        win.deiconify()
        # DWM 标题栏属性（含"文字染成背景色隐形"）必须在窗口可见后设置，
        # 隐藏态设置会被系统重置（实测标题栏回退白色浅字）
        _apply_titlebar_colors(win, bg=PALETTE["card"], text=PALETTE["card"],
                               border=PALETTE["card"])
        win.grab_set()
        win.focus_set()

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
        # 图标 + 主提示整体居中（PNG 图标，避免字体缺字形显示为问号）
        text = "将 Word 文件或文件夹拖到这里"
        icon = _drop_icon_photo()
        tw = tkfont.Font(family="Microsoft YaHei UI", size=12,
                         weight="bold").measure(text)
        iw = icon.width() if icon is not None else 0
        gap = s(10) if icon is not None else 0
        x = (w - (iw + gap + tw)) / 2
        if icon is not None:
            c.create_image(x + iw / 2, h / 2 - 15, image=icon)
            x += iw + gap
        c.create_text(x + tw / 2, h / 2 - 15, text=text,
                      fill=PALETTE["fg"],
                      font=("Microsoft YaHei UI", 12, "bold"))
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
