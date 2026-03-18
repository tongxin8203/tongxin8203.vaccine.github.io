#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
每日文件摘要工具
扫描指定文件夹中的 PDF/PPT/Word/Excel/TXT 文件，生成摘要 PDF。
"""

import json
import os
import sys
import datetime
import traceback

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
    """读取 PPT/PPTX 文件文本内容"""
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


def read_docx(filepath):
    """读取 Word 文件文本内容"""
    from docx import Document
    try:
        doc = Document(filepath)
        text_parts = [para.text for para in doc.paragraphs if para.text.strip()]
        return "\n".join(text_parts)
    except Exception as e:
        return f"[读取失败: {e}]"


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
    ".ppt": read_pptx,
    ".docx": read_docx,
    ".doc": read_docx,
    ".xlsx": read_xlsx,
    ".xls": read_xlsx,
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

    # 扫描文件
    file_types_lower = [t.lower() for t in file_types]
    files = scan_folder(scan_folder_path, file_types_lower)
    print(f"找到 {len(files)} 个文件")

    if not files:
        print("没有找到匹配的文件，跳过生成。")
        return

    # 逐个读取并生成摘要
    summaries = []
    for filepath in files:
        filename = os.path.basename(filepath)
        ext = os.path.splitext(filename)[1].lower()
        reader = READERS.get(ext)
        if not reader:
            continue

        print(f"  处理: {filename}")
        try:
            text = reader(filepath)
            summary = make_summary(text, max_summary_length)
        except Exception as e:
            summary = f"[处理失败: {e}]"
            traceback.print_exc()

        summaries.append({
            "filename": filename,
            "filepath": filepath,
            "filetype": ext,
            "filesize": format_size(os.path.getsize(filepath)),
            "summary": summary,
        })

    # 生成 PDF
    output_path = os.path.join(output_folder, f"{date_str}.pdf")
    print(f"生成摘要 PDF: {output_path}")
    generate_summary_pdf(summaries, output_path, date_str)
    print("完成!")


if __name__ == "__main__":
    main()
