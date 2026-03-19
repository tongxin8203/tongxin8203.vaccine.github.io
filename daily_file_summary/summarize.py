#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
每日文件摘要工具
扫描指定文件夹中的 PDF/PPT/Word/Excel/TXT 文件，生成摘要 PDF。
"""

import json
import os
import sys
import re
import datetime
import traceback
import subprocess


def check_and_install_dependencies():
    """检查必要的依赖库，缺失则自动安装"""
    deps = {
        "PyPDF2": "PyPDF2",
        "pptx": "python-pptx",
        "docx": "python-docx",
        "openpyxl": "openpyxl",
        "xlrd": "xlrd",
        "olefile": "olefile",
        "reportlab": "reportlab",
        "chardet": "chardet",
    }
    missing = []
    for module_name, pip_name in deps.items():
        try:
            __import__(module_name)
        except ImportError:
            missing.append(pip_name)

    if not missing:
        return

    print(f"检测到缺少依赖库: {', '.join(missing)}")
    print("正在自动安装...")
    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install"] + missing,
            stdout=sys.stdout, stderr=sys.stderr
        )
        print("依赖安装完成!")
    except subprocess.CalledProcessError:
        print("=" * 50)
        print("自动安装失败，请手动运行:")
        print(f"  {sys.executable} -m pip install {' '.join(missing)}")
        print("=" * 50)
        sys.exit(1)


# --- 文件读取模块 ---

def read_pdf(filepath):
    """读取 PDF 文件文本内容"""
    from PyPDF2 import PdfReader
    try:
        reader = PdfReader(filepath)
        text_parts = []
        for page in reader.pages:
            t = page.extract_text()
            if t:
                text_parts.append(t)
        return "\n".join(text_parts)
    except Exception as e:
        return f"[读取失败: {e}]"


def read_pptx(filepath):
    """读取 PPTX 文件文本内容"""
    from pptx import Presentation
    try:
        prs = Presentation(filepath)
        text_parts = []
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        text = para.text.strip()
                        if text:
                            text_parts.append(text)
        return "\n".join(text_parts)
    except Exception as e:
        return f"[读取失败: {e}]"


def read_ppt_legacy(filepath):
    """读取旧版 .ppt 文件"""
    # 先尝试 python-pptx（有些 .ppt 其实是新格式改了扩展名）
    try:
        result = read_pptx(filepath)
        if not result.startswith("[读取失败"):
            return result
    except Exception:
        pass
    # 使用 olefile 从二进制 PPT 中提取文本
    return _extract_text_from_ole(filepath, "ppt")


def read_docx(filepath):
    """读取 DOCX 文件文本内容"""
    from docx import Document
    try:
        doc = Document(filepath)
        text_parts = [para.text for para in doc.paragraphs if para.text.strip()]
        return "\n".join(text_parts)
    except Exception as e:
        return f"[读取失败: {e}]"


def read_doc_legacy(filepath):
    """读取旧版 .doc 文件"""
    # 先尝试 python-docx（有些 .doc 其实是新格式改了扩展名）
    try:
        result = read_docx(filepath)
        if not result.startswith("[读取失败"):
            return result
    except Exception:
        pass
    # 使用 olefile 从二进制 DOC 中提取文本
    return _extract_text_from_ole(filepath, "doc")


def read_xlsx(filepath):
    """读取 Excel 文件文本内容"""
    from openpyxl import load_workbook
    try:
        wb = load_workbook(filepath, read_only=True, data_only=True)
        text_parts = []
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            text_parts.append(f"[工作表: {sheet_name}]")
            for row in ws.iter_rows(values_only=True):
                cells = [str(c) if c is not None else "" for c in row]
                line = " | ".join(cells).strip()
                if line.replace("|", "").strip():
                    text_parts.append(line)
        wb.close()
        return "\n".join(text_parts)
    except Exception as e:
        return f"[读取失败: {e}]"


def read_xls_legacy(filepath):
    """读取旧版 .xls 文件（使用 xlrd）"""
    import xlrd
    try:
        wb = xlrd.open_workbook(filepath)
        text_parts = []
        for sheet in wb.sheets():
            text_parts.append(f"[工作表: {sheet.name}]")
            for row_idx in range(sheet.nrows):
                cells = []
                for col_idx in range(sheet.ncols):
                    val = sheet.cell_value(row_idx, col_idx)
                    cells.append(str(val) if val else "")
                line = " | ".join(cells).strip()
                if line.replace("|", "").strip():
                    text_parts.append(line)
        return "\n".join(text_parts)
    except Exception as e:
        return f"[读取失败: {e}]"


def _extract_text_from_ole(filepath, file_type):
    """从旧版 Office 二进制文件 (OLE2) 中提取文本"""
    import olefile
    try:
        if not olefile.isOleFile(filepath):
            return "[读取失败: 文件不是有效的 OLE 格式]"

        ole = olefile.OleFileIO(filepath)
        text = ""

        if file_type == "doc":
            # Word .doc：文本存储在 "WordDocument" 流中
            # 尝试从 Word Document 流提取
            if ole.exists("WordDocument"):
                data = ole.openstream("WordDocument").read()
                # 尝试提取 UTF-16LE 编码的文本
                text = _extract_unicode_text(data)
            # 备用：遍历所有流提取文本
            if not text.strip():
                text = _extract_all_streams_text(ole)

        elif file_type == "ppt":
            # PowerPoint .ppt：文本通常在 "PowerPoint Document" 流中
            if ole.exists("PowerPoint Document"):
                data = ole.openstream("PowerPoint Document").read()
                text = _extract_ppt_text_records(data)
            if not text.strip():
                text = _extract_all_streams_text(ole)

        ole.close()
        return text if text.strip() else "[文件内容为空或无法提取文本]"

    except Exception as e:
        return f"[读取失败: {e}]"


def _extract_unicode_text(data):
    """从二进制数据中提取 Unicode (UTF-16LE) 文本片段"""
    # 查找连续的 UTF-16LE 字符序列（中文和ASCII混合）
    results = []
    try:
        decoded = data.decode("utf-16-le", errors="ignore")
        # 提取可读文本片段（至少4个字符）
        chunks = re.findall(r'[\u4e00-\u9fff\u3000-\u303f\uff00-\uffefa-zA-Z0-9\s，。！？、；：\u201c\u201d\u2018\u2019（）《》\-.,;:!?()\[\]+=/@#$%&*\\]{4,}', decoded)
        for chunk in chunks:
            cleaned = chunk.strip()
            if cleaned and len(cleaned) >= 4:
                results.append(cleaned)
    except Exception:
        pass
    return "\n".join(results)


def _extract_ppt_text_records(data):
    """从 PPT 二进制数据中提取文本记录"""
    # PPT 文本记录类型：
    # 0x0FA0 (4000) = TextCharsAtom (UTF-16LE)
    # 0x0FA8 (4008) = TextBytesAtom (ASCII/Latin-1)
    texts = []
    pos = 0
    while pos < len(data) - 8:
        # 每个记录: 2字节版本/实例, 2字节类型, 4字节长度
        rec_type = int.from_bytes(data[pos + 2:pos + 4], 'little')
        rec_len = int.from_bytes(data[pos + 4:pos + 8], 'little')

        if rec_len < 0 or rec_len > len(data) - pos - 8:
            pos += 1
            continue

        if rec_type == 0x0FA0:  # TextCharsAtom - UTF-16LE
            try:
                text = data[pos + 8:pos + 8 + rec_len].decode('utf-16-le', errors='ignore').strip()
                if text and len(text) >= 2:
                    texts.append(text)
            except Exception:
                pass
            pos += 8 + rec_len
        elif rec_type == 0x0FA8:  # TextBytesAtom - ASCII
            try:
                text = data[pos + 8:pos + 8 + rec_len].decode('latin-1', errors='ignore').strip()
                if text and len(text) >= 2:
                    texts.append(text)
            except Exception:
                pass
            pos += 8 + rec_len
        else:
            pos += 1

    return "\n".join(texts)


def _extract_all_streams_text(ole):
    """从 OLE 文件所有流中提取可读文本（备用方案）"""
    results = []
    for stream_path in ole.listdir():
        try:
            data = ole.openstream(stream_path).read()
            # 尝试 UTF-16LE
            text = _extract_unicode_text(data)
            if text.strip():
                results.append(text.strip())
        except Exception:
            continue
    return "\n".join(results)


def read_txt(filepath):
    """读取 TXT 文件文本内容，自动检测编码"""
    import chardet
    try:
        with open(filepath, "rb") as f:
            raw = f.read()
        detected = chardet.detect(raw)
        encoding = detected.get("encoding", "utf-8") or "utf-8"
        return raw.decode(encoding, errors="replace")
    except Exception as e:
        return f"[读取失败: {e}]"


# 文件类型 -> 读取函数映射
READERS = {
    ".pdf": read_pdf,
    ".pptx": read_pptx,
    ".ppt": read_ppt_legacy,
    ".docx": read_docx,
    ".doc": read_doc_legacy,
    ".xlsx": read_xlsx,
    ".xls": read_xls_legacy,
    ".txt": read_txt,
}


def make_summary(text, max_length=500):
    """简单摘要：取前 max_length 个字符"""
    text = text.strip()
    if not text:
        return "[文件内容为空]"
    if len(text) <= max_length:
        return text
    return text[:max_length] + "..."


def scan_folder(folder, file_types):
    """扫描文件夹，返回匹配的文件列表"""
    matched = []
    for root, _dirs, files in os.walk(folder):
        for fname in files:
            ext = os.path.splitext(fname)[1].lower()
            if ext in file_types:
                matched.append(os.path.join(root, fname))
    matched.sort()
    return matched


def generate_summary_pdf(summaries, output_path, date_str):
    """生成汇总 PDF 文件"""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    # 尝试注册中文字体
    font_name = "Helvetica"
    chinese_font_paths = [
        ("SimSun", "C:\\Windows\\Fonts\\simsun.ttc"),
        ("SimHei", "C:\\Windows\\Fonts\\simhei.ttf"),
        ("MicrosoftYaHei", "C:\\Windows\\Fonts\\msyh.ttc"),
        ("NotoSansSC", "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"),
        ("NotoSansSC", "/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc"),
    ]
    for fname, fpath in chinese_font_paths:
        if os.path.exists(fpath):
            try:
                pdfmetrics.registerFont(TTFont(fname, fpath))
                font_name = fname
                break
            except Exception:
                continue

    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=2 * cm,
        rightMargin=2 * cm,
        topMargin=2 * cm,
        bottomMargin=2 * cm,
    )

    title_style = ParagraphStyle(
        "TitleStyle",
        fontName=font_name,
        fontSize=18,
        leading=24,
        spaceAfter=20,
        alignment=1,  # 居中
    )
    heading_style = ParagraphStyle(
        "HeadingStyle",
        fontName=font_name,
        fontSize=12,
        leading=16,
        spaceAfter=6,
        textColor="darkblue",
    )
    body_style = ParagraphStyle(
        "BodyStyle",
        fontName=font_name,
        fontSize=10,
        leading=14,
        spaceAfter=10,
    )
    meta_style = ParagraphStyle(
        "MetaStyle",
        fontName=font_name,
        fontSize=9,
        leading=12,
        textColor="gray",
        spaceAfter=4,
    )

    elements = []

    # 标题
    title_text = f"每日文件摘要 - {date_str}"
    elements.append(Paragraph(title_text, title_style))
    elements.append(Spacer(1, 0.5 * cm))

    total_info = f"共扫描到 {len(summaries)} 个文件"
    elements.append(Paragraph(total_info, meta_style))
    elements.append(Spacer(1, 0.5 * cm))

    for i, item in enumerate(summaries, 1):
        # 文件标题
        safe_name = item["filename"].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        elements.append(Paragraph(f"{i}. {safe_name}", heading_style))

        # 文件路径和类型
        safe_path = item["filepath"].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        elements.append(Paragraph(f"路径: {safe_path}", meta_style))
        elements.append(Paragraph(f"类型: {item['filetype']}  大小: {item['filesize']}", meta_style))

        # 摘要内容
        safe_summary = item["summary"].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br/>")
        elements.append(Paragraph(safe_summary, body_style))
        elements.append(Spacer(1, 0.3 * cm))

    doc.build(elements)


def format_size(size_bytes):
    """格式化文件大小"""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.1f} MB"


def main():
    # 检查并自动安装依赖
    check_and_install_dependencies()

    # 加载配置
    script_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(script_dir, "config.json")

    if not os.path.exists(config_path):
        print(f"配置文件不存在: {config_path}")
        sys.exit(1)

    with open(config_path, "r", encoding="utf-8") as f:
        config = json.load(f)

    scan_folder_path = config.get("scan_folder", "")
    output_folder = config.get("output_folder", "D:\\wechat-summary")
    file_types = config.get("file_types", [".pdf", ".pptx", ".docx", ".xlsx", ".txt"])
    max_summary_length = config.get("max_summary_length", 500)

    # 支持命令行覆盖扫描路径
    if len(sys.argv) > 1:
        scan_folder_path = sys.argv[1]
    if len(sys.argv) > 2:
        output_folder = sys.argv[2]

    if not scan_folder_path or not os.path.isdir(scan_folder_path):
        print(f"扫描文件夹不存在: {scan_folder_path}")
        print("请修改 config.json 中的 scan_folder 或通过命令行参数指定:")
        print(f"  python {sys.argv[0]} <扫描路径> [输出路径]")
        sys.exit(1)

    # 确保输出目录存在
    os.makedirs(output_folder, exist_ok=True)

    # 日期
    today = datetime.date.today()
    date_str = today.strftime("%Y-%m-%d")

    print(f"扫描文件夹: {scan_folder_path}")
    print(f"输出文件夹: {output_folder}")
    print(f"日期: {date_str}")
    print()

    # 扫描文件
    file_types_lower = [t.lower() for t in file_types]
    files = scan_folder(scan_folder_path, file_types_lower)
    print(f"找到 {len(files)} 个文件")

    if not files:
        print("没有找到匹配的文件，跳过生成。")
        return

    # 逐个读取并生成摘要
    summaries = []
    success_count = 0
    fail_count = 0

    for filepath in files:
        filename = os.path.basename(filepath)
        ext = os.path.splitext(filename)[1].lower()
        reader = READERS.get(ext)
        if not reader:
            continue

        print(f"  处理: {filename} ...", end=" ")
        try:
            text = reader(filepath)
            summary = make_summary(text, max_summary_length)
        except Exception as e:
            summary = f"[处理失败: {e}]"
            traceback.print_exc()

        # 判断是否读取成功
        is_failed = summary.startswith("[读取失败") or summary.startswith("[处理失败") or summary.startswith("[文件内容为空")
        if is_failed:
            fail_count += 1
            print(f"失败")
            print(f"    原因: {summary}")
        else:
            success_count += 1
            print(f"成功")

        try:
            fsize = format_size(os.path.getsize(filepath))
        except OSError:
            fsize = "未知"

        summaries.append({
            "filename": filename,
            "filepath": filepath,
            "filetype": ext,
            "filesize": fsize,
            "summary": summary,
        })

    # 显示统计
    print()
    print(f"处理完成: 成功 {success_count} 个, 失败 {fail_count} 个, 共 {success_count + fail_count} 个")

    # 生成 PDF（文件被占用时自动换名）
    output_path = os.path.join(output_folder, f"{date_str}.pdf")
    for suffix in range(20):
        if suffix > 0:
            output_path = os.path.join(output_folder, f"{date_str}_{suffix}.pdf")
        try:
            # 测试文件是否可写
            with open(output_path, "ab") as _:
                pass
            break
        except PermissionError:
            print(f"  文件被占用: {output_path}，尝试换名...")
            continue

    print(f"生成摘要 PDF: {output_path}")
    generate_summary_pdf(summaries, output_path, date_str)
    print("完成!")


if __name__ == "__main__":
    main()
