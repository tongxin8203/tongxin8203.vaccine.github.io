#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
疫苗日报生成器 - Vaccine Daily PDF Report Generator
每天自动收集疫苗相关信息，翻译成中文，并生成带来源链接的PDF日报
保存路径: C:\Users\WALVAX\Documents\VaccineDaily
"""

import feedparser
import requests
import os
import sys
import time
import logging
from datetime import datetime
from pathlib import Path
from html import unescape
import re

# 翻译
from deep_translator import GoogleTranslator

# ReportLab PDF generation
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether,
)
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY

# ──────────────────────────────────────────────
# 配置区 Configuration
# ──────────────────────────────────────────────

OUTPUT_DIR = r"C:\Users\WALVAX\Documents\VaccineDaily"

# 中文字体路径 (Windows系统自带，按优先级排列)
FONT_CANDIDATES = [
    (r"C:\Windows\Fonts\msyh.ttc",   r"C:\Windows\Fonts\msyhbd.ttc"),   # 微软雅黑
    (r"C:\Windows\Fonts\simhei.ttf", r"C:\Windows\Fonts\simhei.ttf"),   # 黑体
    (r"C:\Windows\Fonts\simsun.ttc", r"C:\Windows\Fonts\simsunb.ttf"),  # 宋体
]

# 新闻来源 RSS Feeds
RSS_SOURCES = [
    {
        "name": "WHO 免疫接种",
        "url": "https://www.who.int/feeds/entity/immunization/en/rss.xml",
        "category": "国际动态",
        "lang": "en",
    },
    {
        "name": "CDC 疫苗",
        "url": "https://tools.cdc.gov/api/v2/resources/media/316422.rss",
        "category": "国际动态",
        "lang": "en",
    },
    {
        "name": "Google新闻·疫苗（中文）",
        "url": "https://news.google.com/rss/search?q=%E7%96%AB%E8%8B%97&hl=zh-CN&gl=CN&ceid=CN:zh-Hans",
        "category": "国内动态",
        "lang": "zh",
    },
    {
        "name": "Google新闻·Vaccine Research",
        "url": "https://news.google.com/rss/search?q=vaccine+approval+clinical+trial&hl=en-US&gl=US&ceid=US:en",
        "category": "研究进展",
        "lang": "en",
    },
    {
        "name": "Reuters 健康",
        "url": "https://feeds.reuters.com/reuters/healthNews",
        "category": "国际动态",
        "lang": "en",
    },
]

# 颜色主题
COLOR_PRIMARY   = colors.HexColor("#005f73")
COLOR_SECONDARY = colors.HexColor("#0a9396")
COLOR_ACCENT    = colors.HexColor("#94d2bd")
COLOR_LIGHT     = colors.HexColor("#e9f5f7")
COLOR_DARK      = colors.HexColor("#001219")
COLOR_GRAY      = colors.HexColor("#6c757d")
COLOR_LINK      = colors.HexColor("#0066cc")
COLOR_WHITE     = colors.white
COLOR_ORIG_BG   = colors.HexColor("#f8f9fa")

MAX_ITEMS_PER_SOURCE = 5    # 每个来源最多抓取条数
SUMMARY_MAX_LEN      = 500  # 摘要最大字符数（翻译前）
REQUEST_TIMEOUT      = 15   # HTTP请求超时秒数
TRANSLATE_DELAY      = 0.4  # 翻译请求间隔（秒），避免触发限流

# ──────────────────────────────────────────────
# 日志
# ──────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)


# ──────────────────────────────────────────────
# 字体注册
# ──────────────────────────────────────────────

def register_fonts() -> bool:
    """注册中文字体，按优先级尝试，返回是否成功。"""
    for normal_path, bold_path in FONT_CANDIDATES:
        if not os.path.exists(normal_path):
            continue
        try:
            pdfmetrics.registerFont(TTFont("Chinese", normal_path))
            bold_src = bold_path if os.path.exists(bold_path) else normal_path
            pdfmetrics.registerFont(TTFont("ChineseBold", bold_src))
            logger.info("已加载字体: %s", normal_path)
            return True
        except Exception as exc:
            logger.warning("字体加载失败 %s: %s", normal_path, exc)
    logger.warning("未找到中文字体，中文可能显示为方块")
    return False


# ──────────────────────────────────────────────
# 文本工具
# ──────────────────────────────────────────────

def clean_html(raw: str) -> str:
    if not raw:
        return ""
    text = re.sub(r"<[^>]+>", " ", raw)
    text = unescape(text)
    return re.sub(r"\s+", " ", text).strip()


def is_chinese(text: str) -> bool:
    """判断文本是否已是中文（CJK字符占比>25%）。"""
    if not text:
        return False
    cjk = sum(1 for c in text if "\u4e00" <= c <= "\u9fff")
    return cjk / len(text) > 0.25


def safe_xml(text: str) -> str:
    """转义ReportLab Paragraph中的XML特殊字符（保留<a>标签外的内容安全）。"""
    return (text
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;"))


# ──────────────────────────────────────────────
# 翻译
# ──────────────────────────────────────────────

_translator = None

def get_translator() -> GoogleTranslator:
    global _translator
    if _translator is None:
        _translator = GoogleTranslator(source="auto", target="zh-CN")
    return _translator


def translate(text: str) -> str:
    """将文本翻译为中文；已是中文则直接返回。失败时返回原文。"""
    if not text or is_chinese(text):
        return text
    # GoogleTranslator 单次上限约5000字符
    chunk = text[:4800]
    try:
        result = get_translator().translate(chunk)
        time.sleep(TRANSLATE_DELAY)
        return result or text
    except Exception as exc:
        logger.warning("翻译失败: %s", exc)
        return text


def translate_item(item: dict) -> dict:
    """翻译单条新闻的标题和摘要，保留原文。"""
    item["title_zh"]   = translate(item["title"])
    item["summary_zh"] = translate(item["summary"]) if item.get("summary") else ""
    return item


# ──────────────────────────────────────────────
# 数据抓取
# ──────────────────────────────────────────────

def fetch_rss(source: dict) -> list:
    items = []
    try:
        headers = {"User-Agent": "VaccineDailyBot/2.0"}
        resp = requests.get(source["url"], headers=headers, timeout=REQUEST_TIMEOUT)
        resp.raise_for_status()
        feed = feedparser.parse(resp.content)

        for entry in feed.entries[:MAX_ITEMS_PER_SOURCE]:
            title   = clean_html(getattr(entry, "title", "（无标题）"))
            summary = clean_html(getattr(entry, "summary",
                                 getattr(entry, "description", "")))
            link    = getattr(entry, "link", "")
            pub     = getattr(entry, "published",
                              getattr(entry, "updated", ""))

            # 截断过长摘要
            if len(summary) > SUMMARY_MAX_LEN:
                summary = summary[:SUMMARY_MAX_LEN].rstrip() + "..."

            items.append({
                "title":      title,
                "title_zh":   "",          # 翻译后填充
                "summary":    summary,
                "summary_zh": "",          # 翻译后填充
                "link":       link,
                "pub_date":   pub,
                "source":     source["name"],
                "category":   source["category"],
                "lang":       source.get("lang", "en"),
            })

        logger.info("✓ %-30s %d 条", source["name"], len(items))

    except requests.exceptions.RequestException as exc:
        logger.warning("✗ %s 请求失败: %s", source["name"], exc)
    except Exception as exc:
        logger.warning("✗ %s 解析失败: %s", source["name"], exc)

    return items


def collect_all_news() -> dict:
    """抓取所有RSS源 → 翻译 → 按类别分组。"""
    all_items = []
    for source in RSS_SOURCES:
        all_items.extend(fetch_rss(source))

    total = len(all_items)
    logger.info("开始翻译 %d 条资讯...", total)

    translated = []
    for i, item in enumerate(all_items, 1):
        logger.info("  翻译 %d/%d: %s", i, total, item["title"][:50])
        translated.append(translate_item(item))

    # 按类别分组
    grouped: dict = {}
    for item in translated:
        grouped.setdefault(item["category"], []).append(item)

    return grouped


# ──────────────────────────────────────────────
# PDF 样式
# ──────────────────────────────────────────────

def build_styles(has_font: bool) -> dict:
    base = "Chinese"     if has_font else "Helvetica"
    bold = "ChineseBold" if has_font else "Helvetica-Bold"

    return {
        # 封面横幅
        "banner_title": ParagraphStyle(
            "BannerTitle", fontName=bold, fontSize=24,
            textColor=COLOR_WHITE, alignment=TA_CENTER,
            spaceAfter=4, leading=30,
        ),
        "banner_sub": ParagraphStyle(
            "BannerSub", fontName=base, fontSize=11,
            textColor=COLOR_ACCENT, alignment=TA_CENTER, spaceAfter=2,
        ),
        "banner_date": ParagraphStyle(
            "BannerDate", fontName=base, fontSize=10,
            textColor=COLOR_ACCENT, alignment=TA_CENTER,
        ),
        # 统计栏
        "stat_label": ParagraphStyle(
            "StatLabel", fontName=bold, fontSize=10,
            textColor=COLOR_PRIMARY, spaceAfter=2,
        ),
        "stat_value": ParagraphStyle(
            "StatValue", fontName=base, fontSize=10,
            textColor=COLOR_DARK, spaceAfter=2,
        ),
        # 分类标题栏
        "section": ParagraphStyle(
            "Section", fontName=bold, fontSize=13,
            textColor=COLOR_WHITE, backColor=COLOR_PRIMARY,
            alignment=TA_LEFT, leftIndent=6,
            spaceBefore=16, spaceAfter=8, leading=20,
        ),
        # 新闻条目 —— 中文翻译标题（主标题）
        "title_zh": ParagraphStyle(
            "TitleZH", fontName=bold, fontSize=11,
            textColor=COLOR_PRIMARY,
            spaceBefore=10, spaceAfter=2, leading=16,
        ),
        # 原文标题（小字灰色）
        "title_orig": ParagraphStyle(
            "TitleOrig", fontName=base, fontSize=8,
            textColor=COLOR_GRAY,
            spaceAfter=3, leading=11,
        ),
        # 来源 | 时间
        "meta": ParagraphStyle(
            "Meta", fontName=base, fontSize=8,
            textColor=COLOR_GRAY,
            spaceAfter=4, leading=11,
        ),
        # 中文摘要
        "summary_zh": ParagraphStyle(
            "SummaryZH", fontName=base, fontSize=9.5,
            textColor=COLOR_DARK,
            spaceAfter=4, leading=15, alignment=TA_JUSTIFY,
        ),
        # 原文摘要（折叠小字）
        "summary_orig": ParagraphStyle(
            "SummaryOrig", fontName=base, fontSize=7.5,
            textColor=COLOR_GRAY,
            spaceAfter=3, leading=11, alignment=TA_JUSTIFY,
        ),
        # 来源链接
        "link": ParagraphStyle(
            "Link", fontName=base, fontSize=8,
            textColor=COLOR_LINK,
            spaceAfter=2, leading=11,
        ),
        # 无资讯占位
        "no_news": ParagraphStyle(
            "NoNews", fontName=base, fontSize=9,
            textColor=COLOR_GRAY, alignment=TA_CENTER,
            spaceBefore=6, spaceAfter=6,
        ),
        # 页脚
        "footer": ParagraphStyle(
            "Footer", fontName=base, fontSize=7.5,
            textColor=COLOR_GRAY, alignment=TA_CENTER,
        ),
    }


# ──────────────────────────────────────────────
# PDF 组件
# ──────────────────────────────────────────────

def make_banner(styles: dict, today: datetime) -> Table:
    weekdays = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
    date_str = today.strftime(f"%Y年%m月%d日  {weekdays[today.weekday()]}")
    gen_str  = today.strftime("生成时间：%H:%M")

    t = Table([
        [Paragraph("疫苗日报", styles["banner_title"])],
        [Paragraph("Vaccine Daily Intelligence Report", styles["banner_sub"])],
        [Paragraph(f"{date_str}　　{gen_str}", styles["banner_date"])],
    ], colWidths=[17 * cm])

    t.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), COLOR_PRIMARY),
        ("TOPPADDING",    (0, 0), (-1, -1), 14),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("LEFTPADDING",   (0, 0), (-1, -1), 12),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 12),
    ]))
    return t


def make_stat_table(grouped: dict, styles: dict) -> Table:
    total = sum(len(v) for v in grouped.values())
    cats  = "　".join(f"{k}（{len(v)}条）" for k, v in grouped.items())

    data = [
        [Paragraph("今日统计", styles["stat_label"]),
         Paragraph(f"共 {total} 条资讯  ·  {len(RSS_SOURCES)} 个来源渠道  ·  已全部翻译为中文",
                   styles["stat_value"])],
        [Paragraph("内容分类", styles["stat_label"]),
         Paragraph(cats or "暂无数据", styles["stat_value"])],
    ]
    t = Table(data, colWidths=[3.2 * cm, 13.8 * cm])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), COLOR_LIGHT),
        ("GRID",          (0, 0), (-1, -1), 0.5, COLOR_ACCENT),
        ("TOPPADDING",    (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING",   (0, 0), (-1, -1), 8),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 8),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return t


def make_news_block(idx: int, item: dict, styles: dict, is_last: bool) -> list:
    """生成单条新闻的PDF段落块。"""
    block = []

    # ① 中文标题（主要，粗体蓝色）
    title_zh = item.get("title_zh") or item["title"]
    block.append(Paragraph(f"{idx}. {safe_xml(title_zh)}", styles["title_zh"]))

    # ② 原文标题（若与中文不同则显示）
    orig_title = item["title"]
    if orig_title and orig_title != title_zh:
        block.append(Paragraph(
            f"原文：{safe_xml(orig_title)}",
            styles["title_orig"]
        ))

    # ③ 元信息：来源 | 发布时间
    meta_parts = [f"来源：{safe_xml(item['source'])}"]
    if item.get("pub_date"):
        # 只取前25字符，去掉时区等冗余信息
        pub = item["pub_date"][:25].strip().rstrip(",")
        meta_parts.append(f"发布：{pub}")
    block.append(Paragraph("　|　".join(meta_parts), styles["meta"]))

    # ④ 中文摘要（核心内容）
    summary_zh = item.get("summary_zh", "").strip()
    if summary_zh:
        block.append(Paragraph(safe_xml(summary_zh), styles["summary_zh"]))

    # ⑤ 原文摘要（折叠小字，仅英文来源显示）
    orig_summary = item.get("summary", "").strip()
    if orig_summary and not is_chinese(orig_summary) and orig_summary != summary_zh:
        block.append(Paragraph(
            f"[原文摘要] {safe_xml(orig_summary)}",
            styles["summary_orig"]
        ))

    # ⑥ 来源链接（可点击）
    link = item.get("link", "").strip()
    if link:
        safe_link = link.replace("&", "&amp;")
        display   = (link[:90] + "…") if len(link) > 90 else link
        block.append(Paragraph(
            f'🔗 <a href="{safe_link}" color="#0066cc"><u>{safe_xml(display)}</u></a>',
            styles["link"]
        ))

    # ⑦ 分隔线
    if not is_last:
        block.append(HRFlowable(
            width="100%", thickness=0.5,
            color=COLOR_ACCENT, spaceBefore=4, spaceAfter=2,
        ))

    return block


# ──────────────────────────────────────────────
# PDF 主构建函数
# ──────────────────────────────────────────────

def build_pdf(grouped: dict, output_path: str, today: datetime):
    has_font = register_fonts()
    styles   = build_styles(has_font)

    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=2 * cm, rightMargin=2 * cm,
        topMargin=1.5 * cm, bottomMargin=1.8 * cm,
        title=f"疫苗日报 {today.strftime('%Y-%m-%d')}",
        author="VaccineDaily Auto-Generator",
        subject="每日疫苗资讯汇总（中文翻译版）",
    )

    story = []

    # 横幅
    story.append(make_banner(styles, today))
    story.append(Spacer(1, 8))

    # 统计栏
    story.append(make_stat_table(grouped, styles))
    story.append(Spacer(1, 10))

    if not grouped:
        story.append(Paragraph(
            "今日暂未获取到疫苗相关资讯，请检查网络连接后重试。",
            styles["no_news"]
        ))
    else:
        for category, items in grouped.items():
            story.append(Paragraph(f"  {category}", styles["section"]))

            if not items:
                story.append(Paragraph("本类别暂无资讯。", styles["no_news"]))
                continue

            for idx, item in enumerate(items, 1):
                is_last = (idx == len(items))
                block   = make_news_block(idx, item, styles, is_last)
                story.append(KeepTogether(block))

    # 页脚
    story.append(Spacer(1, 14))
    story.append(HRFlowable(width="100%", thickness=0.8, color=COLOR_SECONDARY))
    story.append(Spacer(1, 4))
    story.append(Paragraph(
        f"本报告由疫苗日报系统自动生成 · {today.strftime('%Y-%m-%d %H:%M:%S')} · "
        "数据来源：WHO / CDC / Google新闻 / Reuters · 内容经机器翻译，仅供参考",
        styles["footer"]
    ))

    doc.build(story)
    logger.info("PDF 已生成: %s", output_path)


# ──────────────────────────────────────────────
# 主入口
# ──────────────────────────────────────────────

def main():
    today    = datetime.now()
    date_str = today.strftime("%Y%m%d")
    out_dir  = Path(OUTPUT_DIR)

    try:
        out_dir.mkdir(parents=True, exist_ok=True)
    except Exception as exc:
        logger.error("无法创建输出目录 %s: %s", out_dir, exc)
        sys.exit(1)

    output_path = str(out_dir / f"VaccineDaily_{date_str}.pdf")

    logger.info("===== 疫苗日报生成开始 =====")
    logger.info("日期: %s", today.strftime("%Y年%m月%d日 %H:%M"))

    grouped = collect_all_news()

    try:
        build_pdf(grouped, output_path, today)
        logger.info("===== 生成完成 =====")
        logger.info("文件: %s", output_path)
    except Exception as exc:
        logger.error("PDF 生成失败: %s", exc, exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
