"""WorkAid 转换核心：Word/PPT(COM)、图片转PDF、PDF转图片/Word、合并/拆分/水印。"""

import os
import re
import sys
import datetime

# COM 接口（Windows + Word 必需）
try:
    import win32com.client as win32
    import pythoncom
    HAS_WIN32 = True
except Exception:
    HAS_WIN32 = False

from theme import *
__all__ = [
    "A4_H",
    "A4_W",
    "ENGINE_NAME",
    "MM_PT",
    "PPT_PROG_IDS",
    "PP_FIXED_FORMAT_PDF",
    "PP_INTENT_PRINT",
    "PP_RANGE_ALL",
    "WD_FORMAT_PDF",
    "WORD_PROG_IDS",
    "_dispatch_first",
    "_ensure_dir",
    "_pixmap_with_opacity",
    "_unique_out_path",
    "add_watermark",
    "HAS_WIN32",
    "win32",
    "pythoncom",
    "convert_doc_to_pdf",
    "convert_image_format",
    "convert_pdf_to_images",
    "convert_pdf_to_word",
    "convert_ppt_to_pdf",
    "images_to_pdf",
    "IMG_TARGET_EXTS",
    "merge_pdfs",
    "parse_page_ranges",
    "split_pdf",
]


# COM 接口（Windows + Word 必需）
try:
    import win32com.client as win32
    import pythoncom
    HAS_WIN32 = True
except Exception:
    HAS_WIN32 = False


# wdFormatPDF
WD_FORMAT_PDF = 17



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
    if not _ensure_dir(out_dir):
        return "error", f"无法创建输出目录：{out_dir}"

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



def _ensure_dir(path):
    """确保输出目录存在（不存在则创建）。返回是否可用。"""
    try:
        if path and not os.path.isdir(path):
            os.makedirs(path, exist_ok=True)
        return True
    except Exception:  # noqa: BLE001
        return False



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
    if not _ensure_dir(out_dir):
        return "error", f"无法创建输出目录：{out_dir}"

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



def convert_pdf_to_images(pdf_file, out_dir, zoom, fmt, stop_flag=None):
    """将单个 PDF 的每一页渲染成图片。返回 (状态, 信息)。"""
    import fitz  # PyMuPDF（打包时已随依赖一起打入）

    filename = os.path.basename(pdf_file)
    if stop_flag is not None and stop_flag.is_set():
        return "skip", f"已取消：{filename}"
    if not os.path.exists(pdf_file):
        return "error", f"文件不存在：{filename}"
    if not _ensure_dir(out_dir):
        return "error", f"无法创建输出目录：{out_dir}"

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
    if not _ensure_dir(out_dir):
        return "error", f"无法创建输出目录：{out_dir}"

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
    if not _ensure_dir(os.path.dirname(os.path.abspath(out_pdf))):
        return "error", f"无法创建输出目录：{os.path.dirname(os.path.abspath(out_pdf))}"

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
# 图片格式转换
# ---------------------------------------------------------------------------

# 目标格式下拉框顺序（GUI 显示用）
IMG_TARGET_EXTS = ("PNG", "JPG", "WEBP", "BMP", "TIFF", "ICO")

# 各目标格式 Pillow 可直接保存的像素模式；不在表内的模式先转换再保存，
# 避免 "cannot write mode RGBA as JPEG" 一类报错（None 表示几乎全支持）。
_IMG_SAVE_MODES = {
    ".jpg": ("L", "RGB", "CMYK", "YCbCr"),
    ".jpeg": ("L", "RGB", "CMYK", "YCbCr"),
    ".png": ("1", "L", "LA", "P", "RGB", "RGBA", "I;16"),
    ".webp": ("L", "LA", "RGB", "RGBA"),
    ".bmp": ("1", "L", "P", "RGB", "RGBA"),
    ".tif": None,
    ".tiff": None,
    ".ico": ("L", "P", "RGB", "RGBA"),
}


def _has_alpha_channel(im):
    """判断图片是否携带可见透明通道。"""
    return im.mode in ("RGBA", "LA", "PA") or (
        im.mode == "P" and "transparency" in im.info)


def _flatten_alpha(im, bg_rgb=(255, 255, 255)):
    """把含透明的图片铺到纯色底上，返回不带透明通道的 RGB 图。"""
    from PIL import Image
    rgba = im.convert("RGBA")
    bg = Image.new("RGB", rgba.size, bg_rgb)
    bg.paste(rgba, mask=rgba.getchannel("A"))
    return bg


def _prepare_image_for_save(im, ext):
    """按目标扩展名调整图片模式/尺寸，返回可直接 save() 的图片对象。

    - JPG 不支持透明：含透明区域自动铺白底
    - 特殊模式（CMYK / I / F / YCbCr 等）转成常规 RGB / L
    - ICO 单帧最大 256×256，超出则等比缩小
    """
    from PIL import Image
    ext = ext.lower()
    if ext in (".jpg", ".jpeg"):
        if _has_alpha_channel(im):
            return _flatten_alpha(im)
        if im.mode == "1":
            return im.convert("L")
        if im.mode in ("P", "I", "I;16", "F"):
            return im.convert("RGB")
        return im
    ok_modes = _IMG_SAVE_MODES.get(ext)
    if ok_modes is not None and im.mode not in ok_modes:
        im = im.convert("RGBA" if _has_alpha_channel(im) else "RGB")
    if ext == ".ico" and max(im.size) > 256:
        im = im.copy()
        im.thumbnail((256, 256), Image.LANCZOS)
    return im


def convert_image_format(src_path, out_dir, target_ext,
                         quality=None, stop_flag=None):
    """把单张图片转换为目标格式，返回 (state, info)。

    state: "ok" / "error"
    info : 成功时为输出文件路径，失败时为错误说明
    target_ext: 目标扩展名（".png" 等，带点小写）
    quality: JPG / WEBP 质量（1-95 整数），None 表示用 Pillow 默认
    out_dir: 输出目录（不存在会自动创建）；重名自动加 _1、_2 后缀
    """
    from PIL import Image

    if stop_flag is not None and stop_flag.is_set():
        return "error", "已取消"
    _ensure_dir(out_dir)
    stem = os.path.splitext(os.path.basename(src_path))[0]
    out_path = _unique_out_path(out_dir, stem, target_ext)
    try:
        with Image.open(src_path) as im:
            im.load()
            frame = _prepare_image_for_save(im, target_ext)
            save_kw = {}
            if quality and target_ext.lower() in (".jpg", ".jpeg", ".webp"):
                save_kw["quality"] = int(quality)
            frame.save(out_path, **save_kw)
        return "ok", out_path
    except Exception as exc:  # noqa: BLE001
        try:
            if os.path.exists(out_path):
                os.remove(out_path)
        except OSError:
            pass
        return "error", str(exc)
