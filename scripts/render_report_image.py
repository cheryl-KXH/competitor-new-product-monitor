from __future__ import annotations

import asyncio
import base64
import html
import re
import sys
import tempfile
import urllib.request
from collections import defaultdict
from datetime import date
from pathlib import Path
from typing import Any

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from generate_weekly_report import DataQualityReport, ReportRecord, clean_price_text, format_title_date, quality_record_label


ROOT = Path(__file__).resolve().parents[1]


def load_image_layout(path: Path | None = None) -> dict[str, Any]:
    import json

    layout_path = path or ROOT / "config" / "image_layout.json"
    with layout_path.open("r", encoding="utf-8") as f:
        return json.load(f)


def image_to_data_uri(path: Path) -> str:
    if not path.exists():
        return ""
    mime = "image/png"
    if path.suffix.lower() in {".jpg", ".jpeg"}:
        mime = "image/jpeg"
    elif path.suffix.lower() == ".webp":
        mime = "image/webp"
    data = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{data}"


def configured_font_path(configs: dict[str, dict[str, Any]], font_key: str) -> Path | None:
    font_cfg = configs.get("font_files", {}).get(font_key, {})
    candidates = [font_cfg.get("sourcePath"), ROOT / str(font_cfg.get("projectPath", ""))]
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return Path(candidate)
    return None


def font_face_css(configs: dict[str, dict[str, Any]], image_layout: dict[str, Any]) -> str:
    fonts = image_layout.get("fonts", {})
    chinese_key = fonts.get("chineseFontFileKey", "chineseFont")
    latin_key = fonts.get("latinFontFileKey", "latinFont")
    chinese_path = configured_font_path(configs, chinese_key)
    latin_path = configured_font_path(configs, latin_key)
    css: list[str] = []
    if chinese_path:
        css.append(
            "@font-face { font-family: 'ReportFont'; "
            f"src: url('{chinese_path.resolve().as_uri()}') format('truetype'); "
            "unicode-range: U+2E80-2EFF, U+3000-303F, U+3400-4DBF, U+4E00-9FFF, U+F900-FAFF, U+FF00-FFEF; }"
        )
    else:
        print(f"WARNING: HTML 中文字体文件不可访问：{chinese_key}", file=sys.stderr)
    if latin_path:
        css.append(
            "@font-face { font-family: 'ReportFont'; "
            f"src: url('{latin_path.resolve().as_uri()}') format('opentype'); "
            "unicode-range: U+0000-00FF, U+2000-206F, U+20A0-20CF; }"
        )
    else:
        print(f"WARNING: HTML 英文字体文件不可访问：{latin_key}", file=sys.stderr)
    return "\n".join(css)


def fetch_image_data_uri(url: str, record: ReportRecord, data_quality_report: DataQualityReport) -> str:
    if not url:
        return ""
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=20) as resp:
            content_type = resp.headers.get("Content-Type", "image/png").split(";")[0].strip() or "image/png"
            data = base64.b64encode(resp.read()).decode("ascii")
        return f"data:{content_type};base64,{data}"
    except Exception:
        data_quality_report.image_download_failures.append(f"{quality_record_label(record)}：产品外观图片下载失败")
        return ""


def escape_text(value: Any) -> str:
    return html.escape(str(value or ""), quote=True)


def html_text(value: Any) -> str:
    return escape_text(value).replace("\n", "<br>")


def render_price_html(value: str, wrap_after_slash: bool) -> str:
    text = clean_price_text(value)
    if wrap_after_slash and "/" in text:
        text = "/\n".join(part.strip() for part in text.split("/") if part.strip())

    escaped = html.escape(text, quote=True)
    pattern = re.compile(r"(?<=[(（])(\d+(?:\.\d+)?\s*元)(?=[)）])")
    escaped = pattern.sub(r'<span class="strike">\1</span>', escaped)
    return escaped.replace("\n", "<br>")


def visual_len(text: str) -> float:
    total = 0.0
    for char in str(text or ""):
        if char == "\n":
            continue
        if "\u4e00" <= char <= "\u9fff":
            total += 1.0
        elif char in "（）()[]【】":
            total += 0.62
        elif char in "，。；：！？、,.!?;:/":
            total += 0.42
        elif char.isspace():
            total += 0.34
        else:
            total += 0.56
    return total


def text_width_px(text: str, font_size_px: int, padding_px: int) -> int:
    return int(round(visual_len(text) * font_size_px + padding_px))


def clamp(value: int, minimum: int, maximum: int) -> int:
    return max(minimum, min(maximum, value))


def price_blocks(text: str) -> list[str]:
    value = clean_price_text(text)
    if not value:
        return [""]
    return [part.strip() for part in value.split("/") if part.strip()] or [value]


def estimate_wrapped_lines(text: str, width_px: int, font_size_px: int, explicit_breaks: bool = False) -> int:
    if not text:
        return 1
    parts = str(text).split("\n") if explicit_breaks else [str(text)]
    capacity = max(1.0, (width_px - 8) / max(1, font_size_px))
    lines = 0
    for part in parts:
        lines += max(1, int((visual_len(part) + capacity - 0.01) // capacity))
    return max(1, lines)


def compute_summary_column_widths(records: list[ReportRecord], layout: dict[str, Any]) -> dict[str, int]:
    summary = layout["summary"]
    fixed = dict(summary.get("fixedColumnWidthsPx", {}))
    table_width = int(summary["tableWidthPx"])
    min_max = summary.get("flexColumnMinMaxPx", {})
    padding = summary.get("textWidthPaddingPx", {})
    font_size = int(layout.get("fonts", {}).get("defaultSizePx", 12))

    fixed_width = sum(int(fixed[key]) for key in ("brand", "count", "category", "launchDate"))
    remaining = table_width - fixed_width

    max_name = max((record.product_name for record in records), key=visual_len, default="")
    name_min, name_max = min_max.get("productName", [118, 190])
    name_need = text_width_px(max_name, font_size, int(padding.get("productName", 18)))
    product_width = clamp(name_need, int(name_min), int(name_max))

    price_min, price_max = min_max.get("price", [100, 170])
    remark_min, remark_max = min_max.get("remark", [68, 150])
    product_width = min(product_width, max(int(name_min), remaining - int(price_min) - int(remark_min)))
    remaining_after_name = remaining - product_width

    longest_price = max((block for record in records for block in price_blocks(record.price)), key=visual_len, default="")
    price_need = text_width_px(longest_price, font_size, int(padding.get("price", 18)))
    price_width = clamp(price_need, int(price_min), min(int(price_max), remaining_after_name - int(remark_min)))

    remark_width = remaining_after_name - price_width
    longest_remark = max((record.remark for record in records), key=visual_len, default="")
    remark_need = text_width_px(longest_remark, font_size, int(padding.get("remark", 16)))
    if remark_width > int(remark_max):
        extra = remark_width - int(remark_max)
        price_room = int(price_max) - price_width
        give_to_price = min(extra, max(0, price_room))
        price_width += give_to_price
        remark_width -= give_to_price
    if remark_need < remark_width and price_width < int(price_max):
        spare = remark_width - max(int(remark_min), remark_need)
        give_to_price = min(spare, int(price_max) - price_width)
        price_width += give_to_price
        remark_width -= give_to_price

    return {
        "brand": int(fixed["brand"]),
        "count": int(fixed["count"]),
        "category": int(fixed["category"]),
        "productName": int(product_width),
        "launchDate": int(fixed["launchDate"]),
        "price": int(price_width),
        "remark": int(remark_width),
    }


def summary_line_count(record: ReportRecord, columns: dict[str, int], layout: dict[str, Any]) -> int:
    font_size = int(layout.get("fonts", {}).get("defaultSizePx", 12))
    price_lines = len(price_blocks(record.price)) if "/" in clean_price_text(record.price) else estimate_wrapped_lines(record.price, columns["price"], font_size)
    return max(
        1,
        estimate_wrapped_lines(record.category, columns["category"], font_size),
        estimate_wrapped_lines(record.product_name, columns["productName"], font_size),
        price_lines,
        estimate_wrapped_lines(record.remark, columns["remark"], font_size),
    )


def title_text(start: date, end: date, configs: dict[str, dict[str, Any]]) -> str:
    template = configs["excel_layout"].get("titleTemplate", "竞品新品周报{start_m}.{start_d}-{end_m}.{end_d}")
    return format_title_date(start, end, template)


def grouped_records(records: list[ReportRecord], brands: list[str]) -> dict[str, list[ReportRecord]]:
    grouped: dict[str, list[ReportRecord]] = defaultdict(list)
    for record in records:
        grouped[record.brand].append(record)
    return {brand: grouped[brand] for brand in brands if grouped.get(brand)}


def build_summary_html(records: list[ReportRecord], brands: list[str], layout: dict[str, Any]) -> str:
    grouped = grouped_records(records, brands)
    columns = compute_summary_column_widths(records, layout)
    colgroup = "".join(f'<col style="width:{columns[key]}px">' for key in ("brand", "count", "category", "productName", "launchDate", "price", "remark"))
    rows: list[str] = [
        '<table class="summary-table">',
        f"<colgroup>{colgroup}</colgroup>",
        '<thead><tr class="summary-header"><th>品牌</th><th>本周新品<br>数量</th><th>品类</th><th>新品名称</th><th>上市时间</th><th>价格</th><th>备注</th></tr></thead>',
        "<tbody>",
    ]
    for brand, brand_records in grouped.items():
        row_count = len(brand_records)
        for index, record in enumerate(brand_records):
            launch = f"{record.launch_date.month}月{record.launch_date.day}日" if record.launch_date else ""
            line_count = min(5, summary_line_count(record, columns, layout))
            row = [f'<tr class="line-count-{line_count}">']
            if index == 0:
                row.append(f'<td class="brand-cell" rowspan="{row_count}">{escape_text(brand)}</td>')
                row.append(f'<td class="count-cell" rowspan="{row_count}">{row_count}</td>')
            row.extend(
                [
                    f'<td class="category-cell">{html_text(record.category)}</td>',
                    f'<td class="product-name-cell">{html_text(record.product_name)}</td>',
                    f'<td class="launch-date-cell">{html_text(launch)}</td>',
                    f'<td class="price-cell">{render_price_html(record.price, True)}</td>',
                    f'<td class="remark-cell">{html_text(record.remark)}</td>',
                    "</tr>",
                ]
            )
            rows.append("".join(row))
    rows.extend(["</tbody>", "</table>"])
    return "\n".join(rows)


def build_tracked_html(brands: list[str], prefix: str) -> str:
    return escape_text(prefix + "、".join(brands))


def render_tracked_html(brands: list[str], prefix: str) -> str:
    return build_tracked_html(brands, prefix)


def build_detail_html(
    records: list[ReportRecord],
    brands: list[str],
    layout: dict[str, Any],
    data_quality_report: DataQualityReport,
) -> str:
    grouped = grouped_records(records, brands)
    summary_columns = compute_summary_column_widths(records, layout)
    label_width = summary_columns["brand"]
    value_width = int(layout["details"]["tableWidthPx"]) - label_width
    colgroup = f'<colgroup><col style="width:{label_width}px"><col style="width:{value_width}px"></colgroup>'
    rows: list[str] = []
    for brand, brand_records in grouped.items():
        rows.append('<table class="detail-table">')
        rows.append(colgroup)
        rows.append(
            f'<tr><th class="brand-header" colspan="2">{escape_text(brand)}本周新品数量： {len(brand_records)}个</th></tr>'
        )
        for record in brand_records:
            image_uri = fetch_image_data_uri(record.image_urls[0], record, data_quality_report) if record.image_urls else ""
            image_html = (
                f'<img class="product-image" src="{image_uri}" alt="{escape_text(record.product_name)}">'
                if image_uri
                else ""
            )
            rows.extend(
                [
                    f'<tr class="detail-row detail-row-compact"><th class="detail-label strong">新品名称</th><td class="detail-value name-row strong">{html_text(record.product_name)}</td></tr>',
                    f'<tr class="detail-row detail-row-compact"><th class="detail-label">产品系列归属</th><td class="detail-value series-row">{html_text(record.series)}</td></tr>',
                    f'<tr class="detail-row detail-row-text"><th class="detail-label">产品卖点介绍</th><td class="detail-value long-text">{html_text(record.selling_point)}</td></tr>',
                    f'<tr class="detail-row detail-row-compact"><th class="detail-label">产品价格</th><td class="detail-value">{render_price_html(record.price, False)}</td></tr>',
                    f'<tr class="detail-row detail-row-compact"><th class="detail-label">原料构成</th><td class="detail-value">{html_text(record.ingredients)}</td></tr>',
                    f'<tr class="detail-row detail-row-image"><th class="detail-label image-label">产品外观</th><td class="detail-value image-cell">{image_html}</td></tr>',
                ]
            )
        rows.append("</table>")
    return "\n".join(rows)


def build_html_document(
    records: list[ReportRecord],
    configs: dict[str, dict[str, Any]],
    image_layout: dict[str, Any],
    start: date,
    end: date,
    brands: list[str],
    data_quality_report: DataQualityReport,
) -> str:
    page = image_layout["page"]
    fonts = image_layout["fonts"]
    colors = image_layout["colors"]
    details = image_layout["details"]
    border_width = int(image_layout.get("border", {}).get("widthPx", 1))
    summary_columns = compute_summary_column_widths(records, image_layout)
    label_width = summary_columns["brand"]
    row_heights = image_layout.get("summary", {}).get("rowHeightByLineCountPx", {})
    row_height_css = "\n".join(
        f".summary-table tr.line-count-{count} > td {{ height: {height}px; }}"
        for count, height in row_heights.items()
    )
    font_css = font_face_css(configs, image_layout)
    logo_cfg = image_layout.get("logo", {})
    logo_uri = image_to_data_uri(ROOT / logo_cfg.get("path", ""))
    intro = configs["excel_layout"].get("introText", "")
    tracked_prefix = configs["excel_layout"].get("trackedBrandsPrefix", "*关注品牌包括：")
    tracked_html = render_tracked_html(brands, tracked_prefix)
    summary_html = build_summary_html(records, brands, image_layout)
    detail_html = build_detail_html(records, brands, image_layout, data_quality_report)
    css = f"""
{font_css}
* {{ box-sizing: border-box; }}
html, body {{ margin: 0; padding: 0; background: #{colors['white']}; }}
body {{
  width: {page['widthPx']}px;
  padding: {page['paddingTopPx']}px 0 {page['paddingBottomPx']}px;
  color: #{colors['black']};
  font-family: {fonts['family']};
  font-size: {fonts['defaultSizePx']}px;
  line-height: 1.35;
}}
.report {{ width: {page['widthPx']}px; margin: 0 auto; }}
.logo {{ display: block; height: {logo_cfg.get('heightPx', 46)}px; margin: 0 auto {logo_cfg.get('marginBottomPx', 8)}px; }}
.title {{ text-align: center; font-size: {fonts['titleSizePx']}px; font-weight: 700; margin: 0 0 18px; }}
.intro {{ width: {image_layout['summary']['tableWidthPx']}px; margin: 0 auto 2px; font-size: {fonts['introSizePx']}px; }}
table {{ border-collapse: collapse; table-layout: fixed; margin-left: auto; margin-right: auto; }}
th, td {{ border: {border_width}px solid #{colors['black']}; text-align: center; vertical-align: middle; padding: 2px 4px; word-break: break-word; }}
.summary-table {{ width: {image_layout['summary']['tableWidthPx']}px; margin-bottom: 2px; }}
.summary-table th, .summary-table td {{ padding: 1px 3px; line-height: 1.35; }}
.summary-table th {{ background: #{colors['headerFill']}; font-weight: 700; }}
.summary-header > th {{ height: {image_layout['summary'].get('headerHeightPx', 58)}px; }}
.summary-header > th:nth-child(2) {{ white-space: nowrap; word-break: keep-all; }}
.brand-cell, .product-name-cell {{ white-space: nowrap; word-break: keep-all; }}
.category-cell, .launch-date-cell {{ word-break: keep-all; }}
.price-cell {{ word-break: keep-all; overflow-wrap: normal; }}
.remark-cell {{ word-break: keep-all; overflow-wrap: break-word; }}
.brand-cell {{ background: #{colors['brandFill']}; font-weight: 700; }}
.count-cell {{ font-weight: 700; font-size: 14px; }}
.tracked {{ width: {image_layout['summary']['tableWidthPx']}px; margin: 0 auto {image_layout.get('trackedBrands', {}).get('marginBottomPx', 12)}px; font-size: {fonts['smallSizePx']}px; text-align: justify; text-align-last: left; white-space: normal; word-break: normal; overflow-wrap: normal; line-height: 1.35; }}
.detail-table {{ width: {details['tableWidthPx']}px; margin-top: 0; margin-bottom: 0; }}
.detail-table + .detail-table {{ margin-top: -{border_width}px; }}
.brand-header {{ height: {details.get('brandHeaderHeightPx', 20)}px; background: #{colors['detailHeaderFill']}; font-weight: 700; }}
.detail-label {{ width: {label_width}px; background: #{colors['detailLabelFill']}; font-weight: 400; white-space: nowrap; word-break: keep-all; }}
.detail-label.strong {{ font-weight: 700; }}
.detail-table th, .detail-table td {{ line-height: 1.35; padding: 1px 3px; }}
.detail-row-compact > th, .detail-row-compact > td {{ height: 18px; }}
.detail-row-text > th, .detail-row-text > td {{ min-height: 38px; }}
.detail-value {{ padding: {details['cellPaddingPx']}px 4px; }}
.name-row, .series-row {{ background: #{colors['detailLabelFill']}; }}
.strong {{ font-weight: 700; }}
.long-text {{ text-align: {image_layout.get('textAlignment', {}).get('sellingPoint', 'left')}; }}
.image-label {{ height: {details['imageHeightPx'] + 16}px; }}
.image-cell {{ height: {details['imageHeightPx'] + 16}px; padding: 6px; }}
.product-image {{ display: block; max-height: {details['imageHeightPx']}px; max-width: {details['imageMaxWidthPx']}px; margin: 0 auto; object-fit: contain; }}
.strike {{ text-decoration: line-through; }}
{row_height_css}
"""
    logo_html = f'<img class="logo" src="{logo_uri}" alt="logo">' if logo_uri else ""
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width={page['widthPx']}, initial-scale=1">
<title>{escape_text(title_text(start, end, configs))}</title>
<style>{css}</style>
</head>
<body>
<main class="report">
{logo_html}
<h1 class="title">{escape_text(title_text(start, end, configs))}</h1>
<div class="intro">{escape_text(intro)}</div>
{summary_html}
<div class="tracked">{tracked_html}</div>
{detail_html}
</main>
</body>
</html>
"""


async def screenshot_html_with_playwright(html_path: Path, png_path: Path, image_layout: dict[str, Any]) -> None:
    from playwright.async_api import async_playwright

    render_cfg = image_layout.get("render", {})
    width = int(image_layout.get("page", {}).get("widthPx", 640))
    scale = int(render_cfg.get("deviceScaleFactor", 2))
    executable = render_cfg.get("chromeExecutable")
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            executable_path=executable if executable and Path(executable).exists() else None,
            headless=True,
        )
        page = await browser.new_page(viewport={"width": width, "height": 900}, device_scale_factor=scale)
        await page.goto(html_path.resolve().as_uri(), wait_until="networkidle")
        await page.screenshot(path=str(png_path), full_page=True)
        await browser.close()


def render_report_outputs(
    records: list[ReportRecord],
    configs: dict[str, dict[str, Any]],
    image_layout: dict[str, Any],
    start: date,
    end: date,
    brands: list[str],
    html_path: Path | None,
    png_path: Path | None,
    data_quality_report: DataQualityReport,
) -> None:
    html_content = build_html_document(records, configs, image_layout, start, end, brands, data_quality_report)
    if html_path:
        html_path.parent.mkdir(parents=True, exist_ok=True)
        html_path.write_text(html_content, encoding="utf-8")
        screenshot_source = html_path
    else:
        temp = tempfile.NamedTemporaryFile("w", suffix=".html", encoding="utf-8", delete=False)
        try:
            temp.write(html_content)
            temp.close()
            screenshot_source = Path(temp.name)
        finally:
            pass

    try:
        if png_path:
            png_path.parent.mkdir(parents=True, exist_ok=True)
            asyncio.run(screenshot_html_with_playwright(screenshot_source, png_path, image_layout))
    finally:
        if not html_path:
            screenshot_source.unlink(missing_ok=True)
