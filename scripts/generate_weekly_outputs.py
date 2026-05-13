#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
from pathlib import Path

import generate_weekly_report as excel_report
from render_report_image import load_image_layout, render_report_outputs


ROOT = Path(__file__).resolve().parents[1]
VALID_OUTPUT_MODES = {"all", "excel", "html", "image"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="从钉钉 AI 表生成竞品新品周报 Excel、HTML 和 PNG 长图")
    parser.add_argument("--start-date", help="开始日期，格式 YYYY-MM-DD")
    parser.add_argument("--end-date", help="结束日期，格式 YYYY-MM-DD")
    parser.add_argument("--brands", help="品牌范围，使用英文逗号分隔，例如：霸王茶姬,古茗")
    parser.add_argument(
        "--output-mode",
        choices=sorted(VALID_OUTPUT_MODES),
        default="all",
        help="输出模式：all 默认输出 Excel+HTML+PNG；excel 只输出 Excel；html 只输出 HTML；image 只输出 PNG",
    )
    parser.add_argument("--output-dir", help="输出目录，默认读取 config/image_layout.json 的 outputDirectory")
    return parser.parse_args()


def report_stem(start, end) -> str:
    return f"竞品新品周报{start.isoformat()}_{end.isoformat()}"


def resolve_output_root(arg_output_dir: str | None, image_layout: dict) -> Path:
    output_dir = Path(arg_output_dir or image_layout.get("outputDirectory") or image_layout.get("outputDirectory", "outputs"))
    if not output_dir.is_absolute():
        output_dir = ROOT / output_dir
    return output_dir


def resolve_report_dir(output_root: Path, stem: str) -> Path:
    output_root.mkdir(parents=True, exist_ok=True)
    base = output_root / stem
    if not base.exists():
        return base
    escaped = re.escape(stem)
    pattern = re.compile(rf"^{escaped}_(\d+)$")
    max_suffix = 1
    for existing in output_root.iterdir():
        if not existing.is_dir():
            continue
        if existing.name == stem:
            max_suffix = max(max_suffix, 1)
            continue
        match = pattern.match(existing.name)
        if match:
            max_suffix = max(max_suffix, int(match.group(1)))
    return output_root / f"{stem}_{max_suffix + 1}"


def dedupe_quality_warnings(data_quality_report: excel_report.DataQualityReport) -> None:
    data_quality_report.missing_fields = list(dict.fromkeys(data_quality_report.missing_fields))
    data_quality_report.missing_images = list(dict.fromkeys(data_quality_report.missing_images))
    data_quality_report.image_download_failures = list(dict.fromkeys(data_quality_report.image_download_failures))


def main() -> int:
    args = parse_args()
    configs = excel_report.load_configs()
    image_layout = load_image_layout()

    warnings = excel_report.validate_config(configs)
    blocking = [w for w in warnings if "缺少" in w or "必须" in w or "需要" in w]
    if blocking:
        print("配置校验失败：")
        for warning in warnings:
            print(f"- {warning}")
        return 1
    for warning in warnings:
        print(f"WARNING: {warning}")

    start, end = excel_report.resolve_date_window(args)
    brands = excel_report.parse_brands(args.brands, configs["report_rules"])
    table_fields = excel_report.fetch_table_fields(configs)
    field_ids = excel_report.resolve_field_ids(configs["field_mapping"], table_fields)
    raw_records = excel_report.query_records(configs, field_ids, start, end)
    records = excel_report.normalize_records(raw_records, configs, field_ids, brands)
    data_quality_report = excel_report.collect_data_quality_report(records, configs["report_rules"])

    stem = report_stem(start, end)
    output_root = resolve_output_root(args.output_dir, image_layout)
    report_dir = resolve_report_dir(output_root, stem)
    report_dir.mkdir(parents=True, exist_ok=True)
    file_stem = report_dir.name

    output_paths: list[Path] = []
    if args.output_mode in {"all", "excel"}:
        xlsx_path = report_dir / f"{file_stem}.xlsx"
        excel_report.build_workbook(records, configs, start, end, brands, xlsx_path, data_quality_report)
        excel_report.cleanup_legacy_image_cache(report_dir)
        output_paths.append(xlsx_path)

    html_path = report_dir / f"{file_stem}.html" if args.output_mode in {"all", "html"} else None
    png_path = report_dir / f"{file_stem}.png" if args.output_mode in {"all", "image"} else None
    if html_path or png_path:
        render_report_outputs(
            records,
            configs,
            image_layout,
            start,
            end,
            brands,
            html_path,
            png_path,
            data_quality_report,
        )
        if html_path:
            output_paths.append(html_path)
        if png_path:
            output_paths.append(png_path)

    dedupe_quality_warnings(data_quality_report)
    excel_report.print_data_quality_warnings(data_quality_report)
    print(f"已生成目录：{report_dir}")
    print(f"记录数：{len(records)}")
    print(f"数据提醒：缺失字段 {data_quality_report.missing_field_count} 处，缺图 {data_quality_report.image_issue_count} 条。")
    for path in output_paths:
        print(f"本次生成文件：{path.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
