"""WorkAid 自绘控件：BgCanvas 与 Round 系列（按钮/输入/下拉/进度/勾选/卡片/侧栏）。"""

import os
import tkinter as tk
from tkinter import messagebox
from tkinter import font as tkfont

# 拖拽支持：缺失时退回普通 Tk
try:
    import tkinterdnd2 as tkdnd
    HAS_DND = True
except Exception:
    tkdnd = None
    HAS_DND = False

from theme import *
__all__ = [
    "BgCanvas",
    "RoundButton",
    "RoundCard",
    "RoundCheck",
    "RoundCombo",
    "RoundEntry",
    "RoundNotebook",
    "RoundProgress",
    "_DND_ACTIVE",
    "_register_page_drop",
    "move_selected",
    "HAS_DND",
    "tkdnd",
    "refresh_palette",
    "register_round_styles",
]



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
                 parent_bg=None, icon=None, fill=None, **kw):
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
        self._fill = fill                     # 自定义底色（优先于 variant 配色）
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
        base = self._fill or PALETTE.get(self._variant, PALETTE["secondary"])
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

    def on_change(self, callback):
        """界面构建完成后再绑定选中项回调。

        current() / set() / select_index() 都会立即触发 command，若在构造函数
        里绑定，current(0) 会在后续控件尚未创建时就把回调跑一遍。所以需要联动的
        下拉框统一先建控件、再调用本方法挂回调。
        """
        self._command = callback



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
