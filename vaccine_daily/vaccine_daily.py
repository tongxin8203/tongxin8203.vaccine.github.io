#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
疫苗日报生成器 - Vaccine Daily PDF Report Generator
每天自动收集疫苗相关信息并生成PDF日报
保存路径: C:\Users\WALVAX\Documents\VaccineDaily
"""

import feedparser
import requests
import os
import sys
import logging
import textwrap
from datetime import datetime
from pathlib import Path
from html import unescape
import re

# ReportLab PDF generation
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether, PageBreak
)
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY

# ──────────────────────────────────────────────
# 配置区 Configuration
# ──────────────────────────────────────────────

OUTPUT_DIR = r"C:\Users\WALVAX\Documents\VaccineDaily"

# 中文字体路径 (Windows系统自带)
FONT_PATHS = {
    "normal": r"C:\Windows\Fonts\msyh.ttc",        # 微软雅黑
    "bold":   r"C:\Windows\Fonts\msyhbd.ttc",       # 微软雅黑 Bold
    "fallback_normal": r"C:\Windows\Fonts\simhei.ttf",  # 黑体 (备用)
    "fallback_bold":   r"C:\Windows\Fonts\simhei.ttf",  # 黑体 (备用)
}

# 新闻来源 RSS Feeds
RSS_SOURCES = [
    {
        "name": "WHO 免疫接种",
        "url": "https://www.who.int/feeds/entity/immunization/en/rss.xml",
        "category": "国际动态",
        "lang": "en",
    },
    {
        "name": "CDC 疫苗新闻",
        "url": "https://tools.cdc.gov/api/v2/resources/media/316422.rss",
        "category": "国际动态",
        "lang": "en",
    },
    {
        "name": "Google新闻 - 疫苗",
        "url": "https://news.google.com/rss/search?q=%E7%96%AB%E8%8B%97&hl=zh-CN&gl=CN&ceid=CN:zh-Hans",
        "category": "国内动态",
        "lang": "zh",
    },
    {
        "name": "Google新闻 - vaccine",
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
COLOR_WHITE     = colors.white

MAX_ITEMS_PER_SOURCE = 5   # 每个来源最多抓取条数
REQUEST_TIMEOUT      = 15  # 请求超时秒数

# ──────────────────────────────────────────────
# 日志配置
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

def register_fonts():
    """注册中文字体，优先微软雅黑，备用黑体，最后使用内置字体。"""
    registered = False

    for key, path in FONT_PATHS.items():
        if not os.path.exists(path):
            continue
        try:
            if "bold" in key:
                pdfmetrics.registerFont(TTFont("ChineseBold", path))
            else:
                pdfmetrics.registerFont(TTFont("Chinese", path))
            registered = True
            logger.info("已注册字体: %s -> %s", key, path)
            break  # 成功就退出
        except Exception as exc:
            logger.warning("字体注册失败 %s: %s", path, exc)

    # Bold font
    bold_registered = False
    for key in ("bold", "fallback_bold"):
        path = FONT_PATHS.get(key, "")
        if not os.path.exists(path):
            continue
        try:
            pdfmetrics.registerFont(TTFont("ChineseBold", path))
            bold_registered = True
            break
        except Exception:
            pass

    if not registered:
        logger.warning("未找到中文字体，将使用内置字体（中文可能显示为方块）")
        return False

    if not bold_registered:
        # 复用普通字体作为粗体
        try:
            pdfmetrics.registerFont(TTFont("ChineseBold", FONT_PATHS.get("normal", "")))
        except Exception:
            pass

    return True


# ──────────────────────────────────────────────
# 文本清洗工具
# ──────────────────────────────────────────────

def clean_html(raw: str) -> str:
    """去除HTML标签，还原HTML实体。"""
    if not raw:
        return ""
    text = re.sub(r"<[^>]+>", " ", raw)
    text = unescape(text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def truncate(text: str, max_len: int = 200) -> str:
    if len(text) <= max_len:
        return text
    return text[:max_len].rstrip() + "…"


# ──────────────────────────────────────────────
# 数据抓取
# ──────────────────────────────────────────────

def fetch_rss(source: dict) -> list:
    """抓取单个RSS源，返回新闻条目列表。"""
    items = []
    url = source["url"]
    try:
        headers = {"User-Agent": "VaccineDailyBot/1.0 (+https://github.com/tongxin8203)"}
        resp = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
        resp.raise_for_status()
        feed = feedparser.parse(resp.content)

        for entry in feed.entries[:MAX_ITEMS_PER_SOURCE]:
            title   = clean_html(getattr(entry, "title", "（无标题）"))
            summary = clean_html(getattr(entry, "summary", getattr(entry, "description", "")))
            link    = getattr(entry, "link", "")
            pub     = getattr(entry, "published", getattr(entry, "updated", ""))

            items.append({
                "title":    title,
                "summary":  truncate(summary, 250),
                "link":     link,
                "pub_date": pub,
                "source":   source["name"],
                "category": source["category"],
            })
        logger.info("✓ %s: 获取 %d 条", source["name"], len(items))

    except requests.exceptions.RequestException as exc:
        logger.warning("✗ %s 请求失败: %s", source["name"], exc)
    except Exception as exc:
        logger.warning("✗ %s 解析失败: %s", source["name"], exc)

    return items


def collect_all_news() -> dict:
    """从所有RSS源收集新闻，按类别分组。"""
    grouped = {}
    for source in RSS_SOURCES:
        items = fetch_rss(source)
        cat = source["category"]
        grouped.setdefault(cat, []).extend(items)

    total = sum(len(v) for v in grouped.values())
    logger.info("共收集到 %d 条新闻，分 %d 个类别", total, len(grouped))
    return grouped


# ──────────────────────────────────────────────
# PDF 构建
# ──────────────────────────────────────────────

def build_styles(has_chinese_font: bool) -> dict:
    """构建PDF样式字典。"""
    base   = "Chinese"     if has_chinese_font else "Helvetica"
    bold   = "ChineseBold" if has_chinese_font else "Helvetica-Bold"

    styles = {
        "title": ParagraphStyle(
            "DocTitle",
            fontName=bold,
            fontSize=22,
            textColor=COLOR_WHITE,
            alignment=TA_CENTER,
            spaceAfter=4,
            leading=28,
        ),
        "subtitle": ParagraphStyle(
            "DocSubtitle",
            fontName=base,
            fontSize=11,
            textColor=COLOR_ACCENT,
            alignment=TA_CENTER,
            spaceAfter=2,
        ),
        "date_str": ParagraphStyle(
            "DateStr",
            fontName=base,
            fontSize=10,
            textColor=COLOR_ACCENT,
            alignment=TA_CENTER,
        ),
        "section_header": ParagraphStyle(
            "SectionHeader",
            fontName=bold,
            fontSize=13,
            textColor=COLOR_WHITE,
            backColor=COLOR_PRIMARY,
            alignment=TA_LEFT,
            leftIndent=6,
            rightIndent=6,
            spaceBefore=14,
            spaceAfter=6,
            leading=18,
        ),
        "news_title": ParagraphStyle(
            "NewsTitle",
            fontName=bold,
            fontSize=10,
            textColor=COLOR_PRIMARY,
            spaceBefore=8,
            spaceAfter=2,
            leading=14,
        ),
        "news_meta": ParagraphStyle(
            "NewsMeta",
            fontName=base,
            fontSize=8,
            textColor=COLOR_GRAY,
            spaceAfter=2,
            leading=10,
        ),
        "news_body": ParagraphStyle(
            "NewsBody",
            fontName=base,
            fontSize=9,
            textColor=COLOR_DARK,
            spaceAfter=4,
            leading=13,
            alignment=TA_JUSTIFY,
        ),
        "footer": ParagraphStyle(
            "Footer",
            fontName=base,
            fontSize=8,
            textColor=COLOR_GRAY,
            alignment=TA_CENTER,
        ),
        "no_news": ParagraphStyle(
            "NoNews",
            fontName=base,
            fontSize=9,
            textColor=COLOR_GRAY,
            alignment=TA_CENTER,
            spaceBefore=6,
            spaceAfter=6,
        ),
        "summary_label": ParagraphStyle(
            "SummaryLabel",
            fontName=bold,
            fontSize=10,
            textColor=COLOR_PRIMARY,
            spaceAfter=2,
        ),
        "summary_value": ParagraphStyle(
            "SummaryValue",
            fontName=base,
            fontSize=10,
            textColor=COLOR_DARK,
            spaceAfter=2,
        ),
    }
    return styles


def make_header_table(styles: dict, today: datetime) -> Table:
    """生成顶部标题横幅。"""
    date_cn   = today.strftime("%Y年%m月%d日")
    weekdays  = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
    weekday   = weekdays[today.weekday()]
    gen_time  = today.strftime("生成时间：%H:%M")

    title_para    = Paragraph("疫苗日报", styles["title"])
    subtitle_para = Paragraph("Vaccine Daily Intelligence Report", styles["subtitle"])
    date_para     = Paragraph(f"{date_cn}  {weekday}　　{gen_time}", styles["date_str"])

    header_table = Table(
        [[title_para], [subtitle_para], [date_para]],
        colWidths=[17 * cm],
    )
    header_table.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (-1, -1), COLOR_PRIMARY),
        ("TOPPADDING",   (0, 0), (-1, -1), 12),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 8),
        ("LEFTPADDING",  (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("ROUNDEDCORNERS", [6]),
    ]))
    return header_table


def make_summary_table(grouped: dict, styles: dict) -> Table:
    """生成今日摘要统计小表。"""
    total = sum(len(v) for v in grouped.values())
    cats  = "　".join(f"{k}({len(v)}条)" for k, v in grouped.items())

    data = [
        [Paragraph("今日摘要", styles["summary_label"]),
         Paragraph(f"共 {total} 条资讯，来源 {len(RSS_SOURCES)} 个渠道", styles["summary_value"])],
        [Paragraph("覆盖分类", styles["summary_label"]),
         Paragraph(cats or "暂无数据", styles["summary_value"])],
    ]
    t = Table(data, colWidths=[3.5 * cm, 13.5 * cm])
    t.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (-1, -1), COLOR_LIGHT),
        ("GRID",         (0, 0), (-1, -1), 0.5, COLOR_ACCENT),
        ("TOPPADDING",   (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 5),
        ("LEFTPADDING",  (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return t


def build_pdf(grouped: dict, output_path: str, today: datetime):
    """组装完整PDF文档。"""
    has_font = register_fonts()
    styles   = build_styles(has_font)

    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=2 * cm,
        rightMargin=2 * cm,
        topMargin=1.5 * cm,
        bottomMargin=1.8 * cm,
        title=f"疫苗日报 {today.strftime('%Y-%m-%d')}",
        author="VaccineDaily Auto-Generator",
        subject="每日疫苗资讯汇总",
    )

    story = []

    # ── 标题横幅
    story.append(make_header_table(styles, today))
    story.append(Spacer(1, 10))

    # ── 今日摘要
    story.append(make_summary_table(grouped, styles))
    story.append(Spacer(1, 12))

    # ── 各类别新闻
    if not grouped:
        story.append(Paragraph("今日暂未获取到疫苗相关资讯，请检查网络连接后重试。", styles["no_news"]))
    else:
        for category, items in grouped.items():
            # 分类标题
            story.append(Paragraph(f"  {category}", styles["section_header"]))

            if not items:
                story.append(Paragraph("本类别暂无资讯。", styles["no_news"]))
                continue

            for idx, item in enumerate(items, 1):
                block = []
                # 编号 + 标题
                title_text = f"{idx}. {item['title']}"
                block.append(Paragraph(title_text, styles["news_title"]))

                # 来源 + 时间
                meta_parts = [f"来源：{item['source']}"]
                if item.get("pub_date"):
                    meta_parts.append(f"发布：{item['pub_date'][:25]}")
                if item.get("link"):
                    meta_parts.append(f"链接：{item['link'][:80]}")
                block.append(Paragraph("　｜　".join(meta_parts), styles["news_meta"]))

                # 摘要
                if item.get("summary"):
                    block.append(Paragraph(item["summary"], styles["news_body"]))

                # 分隔线（非最后一条）
                if idx < len(items):
                    block.append(HRFlowable(
                        width="100%", thickness=0.4,
                        color=COLOR_ACCENT, spaceAfter=2, spaceBefore=2
                    ))

                story.append(KeepTogether(block))

    # ── 页脚
    story.append(Spacer(1, 16))
    story.append(HRFlowable(width="100%", thickness=0.8, color=COLOR_SECONDARY))
    story.append(Spacer(1, 4))
    footer_text = (
        f"本报告由疫苗日报系统自动生成 · 生成时间：{today.strftime('%Y-%m-%d %H:%M:%S')} · "
        "数据来源：WHO / CDC / Google新闻 / Reuters · 仅供参考"
    )
    story.append(Paragraph(footer_text, styles["footer"]))

    doc.build(story)
    logger.info("PDF 已生成: %s", output_path)


# ──────────────────────────────────────────────
# 主入口
# ──────────────────────────────────────────────

def main():
    today = datetime.now()
    date_str = today.strftime("%Y%m%d")

    # 确保输出目录存在
    out_dir = Path(OUTPUT_DIR)
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
        logger.info("输出目录: %s", out_dir)
    except Exception as exc:
        logger.error("无法创建输出目录 %s: %s", out_dir, exc)
        sys.exit(1)

    output_path = str(out_dir / f"VaccineDaily_{date_str}.pdf")

    logger.info("===== 疫苗日报生成开始 =====")
    logger.info("日期: %s", today.strftime("%Y年%m月%d日 %H:%M"))

    # 收集新闻
    grouped = collect_all_news()

    # 生成PDF
    try:
        build_pdf(grouped, output_path, today)
        logger.info("===== 生成完成 =====")
        logger.info("文件保存至: %s", output_path)
    except Exception as exc:
        logger.error("PDF 生成失败: %s", exc, exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
