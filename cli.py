# cli.py
# -*- coding: utf-8 -*-
"""命令行模式生成报表（可选）。

用法示例：
    python cli.py -i "Order Status Report.xlsx" -o "Order Status Report_pivot.xlsx"

如果不指定 -o，会默认输出到：<输入文件名>_pivot.xlsx
"""

from __future__ import annotations

import argparse
import os
from datetime import datetime

from process import ReportError, generate_pivot_report_from_upload


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Sales Pivot Report Generator")
    p.add_argument("-i", "--input", required=True, help="输入 Excel 文件路径 (.xlsx/.xls)")
    p.add_argument("-o", "--output", default=None, help="输出 Excel 文件路径 (.xlsx)")
    p.add_argument("--report-date", default=None, help="报表日期（默认：今天，如 Jan-18）")
    return p.parse_args()


def main() -> None:
    args = parse_args()

    in_path = args.input
    if not os.path.isfile(in_path):
        raise SystemExit(f"输入文件不存在：{in_path}")

    base, _ext = os.path.splitext(in_path)
    out_path = args.output or f"{base}_pivot.xlsx"

    report_date = args.report_date or datetime.now().strftime("%b-%d")

    try:
        excel_bytes = open(in_path, "rb").read()
        pivot_bytes, stats = generate_pivot_report_from_upload(
            excel_bytes,
            filename=os.path.basename(in_path),
            report_date=report_date,
        )
        with open(out_path, "wb") as f:
            f.write(pivot_bytes)

        print("[OK] 报表已生成：", out_path)
        print("     报表日期：", stats.get("report_date"))
        print("     行数：", stats.get("rows_used"))
        print("     工厂：", ", ".join(stats.get("factories", [])))

    except ReportError as e:
        raise SystemExit(f"[ERROR] {e}")


if __name__ == "__main__":
    main()
