# WorkAid

> 一站式文档批处理工具箱 —— 8 个常用文档转换功能，批量处理，拖拽即用。

一款 Windows 桌面端文档处理工具，基于 Python + Tkinter 开发，界面全部自绘（圆角、深色主题、无描边纯色块风格）。

## 功能

| 功能 | 说明 |
| --- | --- |
| **Word 转 PDF** | 批量将 Word 文档转为 PDF |
| **PDF 转 Word** | 批量将 PDF 转为可编辑的 Word 文档（无需本机安装 Office） |
| **PPT 转 PDF** | 批量将 PPT 演示文稿转为 PDF |
| **PDF 转图片** | 将 PDF 每一页导出为图片 |
| **PDF 加水印** | 批量为 PDF 添加文字或图片水印，支持平铺/居中、浓度与角度调节 |
| **PDF 合并** | 将多个 PDF 合并为一个 |
| **PDF 拆分** | 按每页拆分或按页码范围拆分 |
| **图片转 PDF** | 将多张图片合并为一个 PDF |

所有批量功能均支持：**文件/文件夹拖拽添加**、**列表上移/下移排序**、**统一输出目录**、**实时进度**、**随时停止**、**完成后打开输出目录**。

## 下载使用

前往 [Releases](../../releases) 页面下载最新版 `WorkAid_vX.Y.Z.exe`。

- **免安装**：单文件绿色版，双击即可运行
- **无需配置**：所有依赖已打包进 exe，目标电脑无需安装 Python、Office 或任何运行库
- **首次启动稍慢**：单文件版需要解压资源，约 5~10 秒属正常现象

## 环境要求

- Windows 10 / 11（64 位）
- 若需使用 Word / PPT 转 PDF，目标电脑需安装 Microsoft Office 或 WPS

> 注：PDF 转 Word、PDF 加水印、PDF 合并、PDF 拆分、PDF 转图片、图片转 PDF 均为纯 Python 实现，**不依赖任何办公软件**。

## 从源码运行

```bash
# 建议使用 Python 3.12（3.15 beta 等版本存在第三方库 ABI 兼容问题）
python -m venv .venv312
.venv312\Scripts\activate

pip install pymupdf pdf2docx python-docx pillow tkinterdnd2 pywin32 pyinstaller

python word2pdf.py
```

## 打包

```bash
.venv312\Scripts\python.exe -m PyInstaller --noconfirm --clean WorkAid.spec
```

产物位于 `dist/WorkAid.exe`。

## 项目结构

```
WorkAid/
├── word2pdf.py          # 主程序（全部逻辑与自绘 UI）
├── WorkAid.spec         # PyInstaller 打包配置
├── app.ico              # 应用图标
├── gen_icons.py         # 图标生成脚本
└── icons/               # 界面图标资源
```

## 版本历史

| 版本 | 变更 |
| --- | --- |
| 1.10.3 | 全局去描边，控件改为纯色块风格 |
| 1.10.2 | 圆角改用真圆弧采样绘制，轮廓更平滑 |
| 1.10.1 | 全部控件圆角统一为 12px |
| 1.10.0 | 新增「PDF 转 Word」批量功能 |
| 1.9.0 | 所有批量页新增上移/下移排序 |
| 1.8.6 | 默认窗口高度调整为 720 |
| 1.8.5 | 修复按钮图标与文字未居中 |
| 1.8.0 | 新增「PDF 加水印」 |
| 1.7.0 | 新增「PDF 拆分」 |
| 1.6.0 | 界面改版为左侧边栏布局 |
| 1.5.0 | 支持 Office / WPS 双引擎 |
| 1.4.0 | 新增「PPT 转 PDF」 |
| 1.3.0 | 新增「PDF 合并」 |
| 1.2.0 | 新增「图片转 PDF」 |
| 1.1.0 | 新增「PDF 转图片」 |
| 1.0.0 | 首个版本：Word 转 PDF |

## License

MIT
