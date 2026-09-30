"""WorkAid 功能页：10 个批量处理标签页类与共享工具。"""

import os
import re
import sys
import queue
import threading
import traceback
import datetime
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
from tkinter import font as tkfont

from theme import *
from widgets import *
from converters import *
__all__ = [
    "IMG_EXTS",
    "IMG_TARGET_EXTS",
    "ImageConvertTab",
    "ImageRenameTab",
    "ImagesToPdfTab",
    "MergePdfsTab",
    "PPT_EXTS",
    "PdfToImagesTab",
    "PdfToWordTab",
    "PptToPdfTab",
    "RN_MODES",
    "RN_STATUS_COLOR",
    "RN_STATUS_TEXT",
    "STATUS_COLOR",
    "STATUS_TEXT",
    "SplitPdfTab",
    "WM_COLORS",
    "WM_COLOR_RGB",
    "WM_FONTSIZE",
    "WM_FONTSIZE_VAL",
    "WM_OPACITY",
    "WM_OPACITY_VAL",
    "WatermarkTab",
    "_human_size",
    "_natural_key",
]


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
        # 图标 + 主提示整体居中（PNG 图标，避免字体缺字形显示为问号）
        text = "将 PDF 文件拖到这里"
        icon = _drop_icon_photo()
        tw = tkfont.Font(family="Microsoft YaHei UI", size=11,
                         weight="bold").measure(text)
        iw = icon.width() if icon is not None else 0
        gap = s(10) if icon is not None else 0
        x = (w - (iw + gap + tw)) / 2
        if icon is not None:
            c.create_image(x + iw / 2, h / 2 - 11, image=icon)
            x += iw + gap
        c.create_text(x + tw / 2, h / 2 - 11, text=text,
                      fill=PALETTE["fg"],
                      font=("Microsoft YaHei UI", 11, "bold"))
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



# ---------------------------------------------------------------------------
# 图片重命名（批量）
# ---------------------------------------------------------------------------
RN_STATUS_TEXT = {
    "pending": "待重命名",
    "ok": "已重命名",
    "skip": "冲突跳过",
    "same": "未更改",
    "error": "失败",
}

RN_STATUS_COLOR = {
    "ok": "#2ecc71",
    "skip": "#95a5a6",
    "same": "#95a5a6",
    "error": "#e74c3c",
    "pending": None,
}

RN_MODES = ("模板编号", "添加前缀", "添加后缀", "查找替换")



class ImageRenameTab:
    """图片批量重命名功能页：模板编号 / 前缀 / 后缀 / 查找替换，带实时预览。"""

    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.root = app.root
        self.files = []           # [[完整路径, 状态 key], ...]
        self._plans = {}          # idx -> 目标新文件名（预览用）
        self._build()
        self._bind_preview_traces()

    # -- 组件快捷构造 --------------------------------------------------------
    def _btn(self, master, text, style, command=None, **kw):
        return RoundButton(master, text=text, bootstyle=style or "secondary",
                           command=command,
                           parent_bg=kw.pop("parent_bg", PALETTE["bg"]), **kw)

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
        self.app.brand_row(header, "图片重命名",
                           "批量重命名图片文件 · 模板编号 / 前后缀 / 查找替换 · 实时预览")

        # 工具栏
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
        self._btn(toolbar, "上移", B.BTN_SECONDARY,
                  lambda: self._move(-1), pad_x=14).pack(side="left", padx=(14, 6))
        self._btn(toolbar, "下移", B.BTN_SECONDARY,
                  lambda: self._move(1), pad_x=14).pack(side="left", padx=6)

        # 文件列表（含新文件名预览列）
        list_card = RoundCard(self.parent, padding=6, parent_bg=BG)
        list_card.pack(fill="both", expand=True, padx=PX, pady=(10, 0))
        list_frame = list_card.body
        cols = ("idx", "old", "new", "size", "status")
        self.tree = Treeview(list_frame, columns=cols, show="headings",
                             selectmode="extended", height=3)
        self.tree.heading("idx", text="顺序")
        self.tree.heading("old", text="原文件名")
        self.tree.heading("new", text="新文件名（预览）")
        self.tree.heading("size", text="大小")
        self.tree.heading("status", text="状态")
        self.tree.column("idx", width=48, anchor="center", stretch=False)
        self.tree.column("old", width=250, anchor="w")
        self.tree.column("new", width=250, anchor="w")
        self.tree.column("size", width=80, anchor="e", stretch=False)
        self.tree.column("status", width=84, anchor="center", stretch=False)
        vsb = Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self._refresh_tags()

        if HAS_DND:
            try:
                self.tree.drop_target_register(tkdnd.DND_FILES)
                self.tree.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass

        self.empty_label = tk.Label(
            list_frame,
            text="还没有添加图片\n点击「添加图片」或把图片 / 文件夹拖到列表上",
            justify="center", bg=PALETTE["card"], fg=PALETTE["muted"],
            font=("Microsoft YaHei UI", 11))
        self.empty_label.place(relx=0.5, rely=0.5, anchor="center")
        if HAS_DND:
            try:
                self.empty_label.drop_target_register(tkdnd.DND_FILES)
                self.empty_label.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:  # noqa: BLE001
                pass

        # 命名模式行（先建 4 组参数 Frame，再建下拉框——current(0) 会触发回调）
        opt_frame = BgCanvas(self.parent)
        opt_frame.pack(fill="x", padx=PX, pady=(10, 0))
        Label(opt_frame, text="命名模式").pack(side="left")
        self.mode_combo = RoundCombo(opt_frame, width=12, parent_bg=BG,
                                     values=RN_MODES, command=self._on_mode_change)
        self.mode_combo.pack(side="left", padx=(10, 20))

        # 模式参数区（4 组，按模式显隐；RoundEntry 宽度为像素）
        # 0 模板编号
        self.f_tpl = tk.Frame(opt_frame, bg=BG)
        Label(self.f_tpl, text="前缀").pack(side="left")
        self.tpl_prefix_var = tk.StringVar(value="图片_")
        RoundEntry(self.f_tpl, textvariable=self.tpl_prefix_var,
                   width=s(130), height=30, parent_bg=BG
                   ).pack(side="left", padx=(6, 14))
        Label(self.f_tpl, text="起始序号").pack(side="left")
        self.tpl_start_var = tk.StringVar(value="1")
        RoundEntry(self.f_tpl, textvariable=self.tpl_start_var,
                   width=s(56), height=30, parent_bg=BG
                   ).pack(side="left", padx=(6, 14))
        Label(self.f_tpl, text="序号位数").pack(side="left")
        self.tpl_digits = RoundCombo(self.f_tpl, width=8, parent_bg=BG,
                                     values=("2 位", "3 位", "4 位"))
        self.tpl_digits.current(0)
        self.tpl_digits.pack(side="left", padx=(6, 0))
        # 1 添加前缀
        self.f_pre = tk.Frame(opt_frame, bg=BG)
        Label(self.f_pre, text="前缀文字").pack(side="left")
        self.pre_var = tk.StringVar()
        RoundEntry(self.f_pre, textvariable=self.pre_var,
                   width=s(150), height=30, parent_bg=BG
                   ).pack(side="left", padx=(6, 0))
        # 2 添加后缀
        self.f_suf = tk.Frame(opt_frame, bg=BG)
        Label(self.f_suf, text="后缀文字").pack(side="left")
        self.suf_var = tk.StringVar()
        RoundEntry(self.f_suf, textvariable=self.suf_var,
                   width=s(150), height=30, parent_bg=BG
                   ).pack(side="left", padx=(6, 0))
        # 3 查找替换
        self.f_rep = tk.Frame(opt_frame, bg=BG)
        Label(self.f_rep, text="查找").pack(side="left")
        self.find_var = tk.StringVar()
        RoundEntry(self.f_rep, textvariable=self.find_var,
                   width=s(110), height=30, parent_bg=BG
                   ).pack(side="left", padx=(6, 14))
        Label(self.f_rep, text="替换为").pack(side="left")
        self.repl_var = tk.StringVar()
        RoundEntry(self.f_rep, textvariable=self.repl_var,
                   width=s(110), height=30, parent_bg=BG
                   ).pack(side="left", padx=(6, 0))
        Label(self.f_rep, text="（仅改文件主名，不影响扩展名）", foreground=PALETTE["muted"],
              font=("Microsoft YaHei UI", 9)).pack(side="left", padx=(10, 0))

        # 进度 + 执行按钮
        prog_frame = BgCanvas(self.parent)
        prog_frame.pack(fill="x", padx=PX, pady=(12, 0))
        self.progress = RoundProgress(prog_frame, mode="determinate",
                                      height=14, parent_bg=BG)
        self.progress.pack(side="left", fill="x", expand=True, pady=8)
        self.btn_rename = self._btn(
            prog_frame, "开始重命名", B.BTN_SUCCESS, self._start_rename)
        self.btn_rename.pack(side="right", padx=(14, 0))

        self.status_var = tk.StringVar(value="就绪")
        self.app._make_status_bar(self.parent, self.status_var)

        # 日志（圆角卡片）
        log_card = RoundCard(self.parent, padding=8, height=104, parent_bg=BG)
        log_card.pack(fill="x", padx=PX, pady=(8, 12))
        log_holder = log_card.body
        tk.Label(log_holder, text="重命名日志", bg=PALETTE["card"],
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

        self._on_mode_change()

    def _bind_preview_traces(self):
        for var in (self.tpl_prefix_var, self.tpl_start_var,
                    self.pre_var, self.suf_var, self.find_var, self.repl_var):
            var.trace_add("write", lambda *_: self._refresh_list())

    def _on_mode_change(self, *_args):
        mode = self.mode_combo.current()
        for i, f in enumerate((self.f_tpl, self.f_pre, self.f_suf, self.f_rep)):
            if i == mode:
                f.pack(side="left")
            else:
                f.pack_forget()
        self._refresh_list()

    # -- 列表 ----------------------------------------------------------------
    def _refresh_tags(self):
        merged = dict(STATUS_COLOR)
        merged.update(RN_STATUS_COLOR)
        for key, color in merged.items():
            if color:
                self.tree.tag_configure(key, foreground=color)

    def _new_name_for(self, path, seq):
        """按当前规则计算 path 的新文件名（含扩展名）。"""
        folder = os.path.dirname(path)
        base = os.path.basename(path)
        stem, ext = os.path.splitext(base)
        mode = self.mode_combo.current()
        if mode == 0:      # 模板编号
            try:
                start = max(0, int(self.tpl_start_var.get().strip() or "1"))
            except ValueError:
                start = 1
            digits = (2, 3, 4)[self.tpl_digits.current()]
            new_stem = "%s%0*d" % (self.tpl_prefix_var.get(), digits, start + seq)
        elif mode == 1:    # 添加前缀
            new_stem = self.pre_var.get() + stem
        elif mode == 2:    # 添加后缀
            new_stem = stem + self.suf_var.get()
        else:              # 查找替换
            find = self.find_var.get()
            repl = self.repl_var.get()
            if not find:
                return base
            new_stem = stem.replace(find, repl)
        return new_stem + ext

    def _refresh_list(self):
        self.tree.delete(*self.tree.get_children())
        self._plans.clear()
        # 先整批计算目标名，检测批内重复（同目录 + 同目标名 → 冲突）
        seen = {}
        for seq, (path, _state) in enumerate(self.files):
            folder = os.path.dirname(path).lower()
            new_name = self._new_name_for(path, seq)
            key = (folder, new_name.lower())
            self._plans[seq] = new_name
            if new_name.lower() == os.path.basename(path).lower():
                self._plans[seq] = None          # 无需更改
            elif key in seen:
                self._plans[seq] = None          # 批内重名冲突
            else:
                seen[key] = seq
        for i, (path, state) in enumerate(self.files, 1):
            try:
                size = _human_size(os.path.getsize(path))
            except OSError:
                size = "—"
            new_name = self._plans.get(i - 1)
            tags = (state,) if RN_STATUS_COLOR.get(state) else ()
            self.tree.insert("", "end", values=(
                i, os.path.basename(path), new_name or "—", size,
                RN_STATUS_TEXT.get(state, state)), tags=tags)
        if getattr(self, "empty_label", None) is not None:
            if self.files:
                self.empty_label.place_forget()
            else:
                self.empty_label.place(relx=0.5, rely=0.5, anchor="center")
        self.status_var.set(f"共 {len(self.files)} 张图片")

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
        if added:
            self._log(f"已添加 {added} 张图片")

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="选择要重命名的图片（Ctrl / Shift 可多选）",
            filetypes=[("图片文件", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff *.webp *.gif"),
                       ("所有文件", "*.*")])
        if paths:
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
        self.files.clear()
        self._refresh_list()
        self.progress["value"] = 0

    def _move(self, delta):
        if not self.files:
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

    # -- 拖拽 ----------------------------------------------------------------
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

    def _register_dnd(self):
        """把本页的文件列表注册为拖拽落点（页面切换时由主界面调用）。"""
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
        if not (HAS_DND and tkdnd is not None):
            return
        try:
            self.tree.drop_target_unregister()
            self.empty_label.drop_target_unregister()
        except Exception:  # noqa: BLE001
            pass

    # -- 执行重命名 ----------------------------------------------------------
    def _start_rename(self):
        if not self.files:
            messagebox.showinfo("提示", "请先添加要重命名的图片。")
            return
        # 校验模板编号的起始序号
        if self.mode_combo.current() == 0:
            try:
                int(self.tpl_start_var.get().strip() or "1")
            except ValueError:
                messagebox.showwarning("提示", "起始序号必须是数字。")
                return

        done = skip = same = error = 0
        for seq, (path, _state) in enumerate(list(self.files)):
            new_name = self._plans.get(seq)
            folder = os.path.dirname(path)
            target = os.path.join(folder, new_name) if new_name else None
            if target is None:
                # 预览标记为 None：无需更改 或 批内重名冲突
                if new_name is None and self._new_name_for(path, seq).lower() \
                        == os.path.basename(path).lower():
                    same += 1
                    self.files[seq][1] = "same"
                else:
                    skip += 1
                    self.files[seq][1] = "skip"
                continue
            if os.path.exists(target) and os.path.abspath(target) != \
                    os.path.abspath(path):
                skip += 1
                self.files[seq][1] = "skip"
                self._log(f"冲突：{os.path.basename(path)} → {new_name}"
                          f"（目标已存在，跳过）")
                continue
            try:
                os.rename(path, target)
                done += 1
                self.files[seq][0] = target
                self.files[seq][1] = "ok"
            except OSError as exc:
                error += 1
                self.files[seq][1] = "error"
                self._log(f"失败：{os.path.basename(path)} → {exc}")

        self._refresh_list()
        total = len(self.files)
        self.progress["value"] = self.progress["maximum"] or 100
        self.status_var.set(
            f"完成：已重命名 {done} · 跳过 {skip} · 未更改 {same} · 失败 {error}"
            f"（共 {total}）")
        self._log(f"批量重命名结束：成功 {done} 个，冲突跳过 {skip} 个，"
                  f"未更改 {same} 个，失败 {error} 个")
        if done:
            messagebox.showinfo("完成", f"已重命名 {done} 个文件。")

    def _log(self, text):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")


# ---------------------------------------------------------------------------
# 图片格式转换（批量）
# ---------------------------------------------------------------------------
CV_QUALITY = ("高（95）", "中（85）", "低（70）")
CV_QUALITY_VAL = {"高（95）": 95, "中（85）": 85, "低（70）": 70}


def _norm_img_ext(ext):
    """归一化图片扩展名（.jpeg 视为 .jpg），用于“已是目标格式”的跳过判断。"""
    ext = (ext or "").lower()
    return ".jpg" if ext == ".jpeg" else ext


class ImageConvertTab:
    """图片格式转换功能页：PNG / JPG / WEBP / BMP / TIFF / ICO 批量互转。

    含透明通道的图片转 JPG 时自动铺白底；ICO 输出自动限制在 256×256 内；
    源文件与目标格式相同（.jpeg 与 .jpg 互视为同格式）的条目自动跳过。
    """

    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.root = app.root
        self.files = []           # [[完整路径, 状态 key], ...]
        self._queue = queue.Queue()
        self._worker = None
        self._stop_flag = threading.Event()
        self._build()
        self._on_format_change()
        self.root.after(100, self._poll_queue)

    # -- 组件快捷构造 --------------------------------------------------------
    def _btn(self, master, text, style, command=None, **kw):
        return RoundButton(master, text=text, bootstyle=style or "secondary",
                           command=command,
                           parent_bg=kw.pop("parent_bg", PALETTE["bg"]), **kw)

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
        self.app.brand_row(header, "图片格式转换",
                           "批量转换图片格式 · PNG / JPG / WEBP / BMP / TIFF / ICO"
                           " · 透明背景自动处理")

        # 工具栏
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

        # 文件列表（含输出文件名预览列）
        list_card = RoundCard(self.parent, padding=6, parent_bg=BG)
        list_card.pack(fill="both", expand=True, padx=PX, pady=(10, 0))
        list_frame = list_card.body
        cols = ("idx", "old", "new", "size", "status")
        self.tree = Treeview(list_frame, columns=cols, show="headings",
                             selectmode="extended", height=3)
        self.tree.heading("idx", text="顺序")
        self.tree.heading("old", text="原文件名")
        self.tree.heading("new", text="输出文件名")
        self.tree.heading("size", text="大小")
        self.tree.heading("status", text="状态")
        self.tree.column("idx", width=48, anchor="center", stretch=False)
        self.tree.column("old", width=240, anchor="w")
        self.tree.column("new", width=240, anchor="w")
        self.tree.column("size", width=80, anchor="e", stretch=False)
        self.tree.column("status", width=84, anchor="center", stretch=False)
        vsb = Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self._refresh_tags()

        if HAS_DND:
            try:
                self.tree.drop_target_register(tkdnd.DND_FILES)
                self.tree.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:  # noqa: BLE001
                pass

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

        # 保存到（留空 = 输出到原图所在文件夹）
        out_frame = BgCanvas(self.parent)
        out_frame.pack(fill="x", padx=PX, pady=(10, 0))
        Label(out_frame, text="保存到").pack(side="left")
        self.out_var = tk.StringVar()
        self.out_entry = RoundEntry(out_frame, textvariable=self.out_var,
                                    height=36, parent_bg=BG)
        self.out_entry.pack(side="left", fill="x", expand=True, padx=(10, 8))
        self._btn(out_frame, "浏览…", B.BTN_SECONDARY,
                  self._choose_out_dir).pack(side="left")

        # 选项行：目标格式 + 质量
        opt_frame = BgCanvas(self.parent)
        opt_frame.pack(fill="x", padx=PX, pady=(10, 0))
        Label(opt_frame, text="目标格式").pack(side="left")
        self.fmt_combo = RoundCombo(opt_frame, width=8, parent_bg=BG,
                                    values=IMG_TARGET_EXTS,
                                    command=self._on_format_change)
        self.fmt_combo.pack(side="left", padx=(10, 20))
        Label(opt_frame, text="图片质量").pack(side="left")
        self.qual_combo = RoundCombo(opt_frame, width=10, parent_bg=BG,
                                     values=CV_QUALITY)
        self.qual_combo.current(1)
        self.qual_combo.pack(side="left", padx=(10, 6))
        Label(opt_frame, text="（仅 JPG / WEBP 生效）",
              foreground=PALETTE["muted"],
              font=("Microsoft YaHei UI", 9)).pack(side="left")

        # 进度 + 转换按钮
        prog_frame = BgCanvas(self.parent)
        prog_frame.pack(fill="x", padx=PX, pady=(12, 0))
        self.progress = RoundProgress(prog_frame, mode="determinate",
                                      height=14, parent_bg=BG)
        self.progress.pack(side="left", fill="x", expand=True, pady=8)
        self.btn_convert = self._btn(prog_frame, "开始转换", B.BTN_SUCCESS,
                                     self._start_convert)
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

    def _row_values(self, idx):
        """列表第 idx 行的显示值（含按当前目标格式推算的输出文件名）。"""
        path, state = self.files[idx]
        stem, _ext = os.path.splitext(os.path.basename(path))
        new_name = stem + "." + (self.fmt_combo.get() or "PNG").lower()
        try:
            size = _human_size(os.path.getsize(path))
        except OSError:
            size = "—"
        return (idx + 1, os.path.basename(path), new_name, size,
                STATUS_TEXT.get(state, state))

    def _refresh_list(self):
        self.tree.delete(*self.tree.get_children())
        for i in range(len(self.files)):
            state = self.files[i][1]
            tags = (state,) if STATUS_COLOR.get(state) else ()
            self.tree.insert("", "end", values=self._row_values(i), tags=tags)
        if getattr(self, "empty_label", None) is not None:
            if self.files:
                self.empty_label.place_forget()
            else:
                self.empty_label.place(relx=0.5, rely=0.5, anchor="center")
        if hasattr(self, "status_var"):
            self.status_var.set(f"共 {len(self.files)} 张图片")

    def _on_format_change(self, *_args):
        if not hasattr(self, "status_var"):
            return                     # 构建期间触发，界面尚未就绪
        self._refresh_list()

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
        if added:
            self._log(f"已添加 {added} 张图片")

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="选择要转换的图片（Ctrl / Shift 可多选）",
            filetypes=[("图片文件", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff *.webp *.gif"),
                       ("所有文件", "*.*")])
        if paths:
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
            messagebox.showwarning("提示", "正在转换中，无法清空列表。")
            return
        self.files.clear()
        self._refresh_list()
        self.progress["value"] = 0

    def _choose_out_dir(self):
        d = filedialog.askdirectory(title="选择转换后图片的保存位置")
        if d:
            self.out_var.set(d)

    # -- 拖拽 ----------------------------------------------------------------
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

    def _register_dnd(self):
        """把本页的文件列表注册为拖拽落点（页面切换时由主界面调用）。"""
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
        if not (HAS_DND and tkdnd is not None):
            return
        try:
            self.tree.drop_target_unregister()
            self.empty_label.drop_target_unregister()
        except Exception:  # noqa: BLE001
            pass

    # -- 转换流程 ------------------------------------------------------------
    def _set_convert_btn(self, text, style):
        if HAS_TTB:
            self.btn_convert.configure(text=text, bootstyle=style)
        else:
            self.btn_convert.configure(text=text)

    def _start_convert(self):
        if self._worker and self._worker.is_alive():
            self._stop_flag.set()
            self.status_var.set("正在停止…")
            return
        if not self.files:
            messagebox.showinfo("提示", "请先添加要转换的图片。")
            return
        target_ext = "." + (self.fmt_combo.get() or "PNG").lower()
        quality = CV_QUALITY_VAL.get(self.qual_combo.get())
        out_dir_pref = self.out_var.get().strip()

        self._stop_flag.clear()
        self.progress["maximum"] = len(self.files)
        self.progress["value"] = 0
        for i in range(len(self.files)):
            self.files[i][1] = "pending"
        self._refresh_list()

        if out_dir_pref:
            self._log(f"开始转换 {len(self.files)} 张图片 → "
                      f"{target_ext.lstrip('.').upper()}（保存到 {out_dir_pref}）…")
        else:
            self._log(f"开始转换 {len(self.files)} 张图片 → "
                      f"{target_ext.lstrip('.').upper()}"
                      f"（未指定保存位置，输出到原图所在文件夹）…")
        self._set_convert_btn("停止", self.app.BTN_DANGER)

        self._worker = threading.Thread(
            target=self._worker_run,
            args=(list(self.files), target_ext, quality, out_dir_pref),
            daemon=True)
        self._worker.start()

    def _worker_run(self, files_snapshot, target_ext, quality, out_dir_pref):
        ok = err = skip = 0
        try:
            for idx, (path, _s) in enumerate(files_snapshot):
                if self._stop_flag.is_set():
                    self._queue.put(("stopped", None))
                    break
                src_ext = os.path.splitext(path)[1].lower()
                if _norm_img_ext(src_ext) == _norm_img_ext(target_ext):
                    skip += 1
                    self._queue.put((
                        "item", (idx, "skip",
                                 f"跳过：{os.path.basename(path)}"
                                 f" 已是 {target_ext.lstrip('.').upper()} 格式",
                                 path)))
                    continue
                self._queue.put(("status", (idx, "working")))
                out_dir = out_dir_pref or os.path.dirname(path)
                state, info = convert_image_format(
                    path, out_dir, target_ext, quality=quality,
                    stop_flag=self._stop_flag)
                if state == "ok":
                    ok += 1
                    info = (f"{os.path.basename(path)} → "
                            f"{os.path.basename(info)}")
                else:
                    err += 1
                self._queue.put(("item", (idx, state, info, path)))
            else:
                self._queue.put(("done", (ok, err, skip)))
        except Exception:  # noqa: BLE001
            self._queue.put(("fatal", traceback.format_exc()))

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
                self.tree.item(children[idx],
                               values=self._row_values(idx),
                               tags=(state,) if STATUS_COLOR.get(state) else ())
        elif kind == "item":
            idx, state, info, path = payload
            if idx < len(children):
                self.files[idx][1] = state
                self.tree.item(children[idx],
                               values=self._row_values(idx),
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
