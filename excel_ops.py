#!/usr/bin/env python3
"""expense-tracker Excel operations."""

import argparse
import json
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DETAIL_SHEET = "明细"
SUMMARY_SHEET = "汇总"

HEADERS = ["序号", "录入时间", "发生日期", "金额（元）", "费用类型", "事由/摘要", "付款方式", "发票状态", "经办人", "备注"]
COL_WIDTHS = [6, 18, 12, 14, 10, 32, 10, 10, 10, 22]

# 报账系统
EXPENSE_TYPES_REIMBURSEMENT = [
    "办公用品", "邮电费", "版面费", "印刷费", "专利费",
    "设备燃动费", "维修费", "委托业务费", "水电费", "租赁费",
    "市内交通费", "活动费", "培训费", "会议费", "加班餐费",
    "公务招待费", "公务用车费", "因公出国费",
    "国内差旅费", "国际差旅费",
    "其他日常费用",
]
# 酬金系统
EXPENSE_TYPES_REMUNERATION = [
    "编审费", "答辩费", "奖励", "课酬", "评审费",
    "科研劳务费", "人才津贴", "咨询费", "稿酬",
    "科研绩效", "中学生奖助学金", "外籍专家劳务费",
    "高层次人才免税", "考务费", "其他劳务",
]

EXPENSE_TYPES = EXPENSE_TYPES_REIMBURSEMENT + EXPENSE_TYPES_REMUNERATION

# 酬金类天然无发票
NO_INVOICE_TYPES = {
    "编审费", "答辩费", "奖励", "课酬", "评审费",
    "科研劳务费", "人才津贴", "咨询费", "稿酬",
    "科研绩效", "中学生奖助学金", "外籍专家劳务费",
    "高层次人才免税", "考务费", "其他劳务",
}

C_DARK_BLUE, C_MID_BLUE, C_LIGHT_BLUE = "1F4E79", "2E75B6", "D6E4F0"
C_ALT, C_WHITE, C_GOLD = "EBF3FB", "FFFFFF", "FFF2CC"

THIN = Side(style="thin", color="CCCCCC")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
AMOUNT_FMT = '#,##0.00'

FONT_BODY   = Font(size=10, name="微软雅黑")
FONT_HEADER = Font(bold=True, color=C_WHITE, size=11, name="微软雅黑")
FONT_TITLE  = Font(bold=True, color=C_WHITE, size=15, name="微软雅黑")
FONT_TOTAL  = Font(bold=True, size=11, name="微软雅黑", color=C_DARK_BLUE)
FONT_LABEL  = Font(bold=True, size=11, name="微软雅黑")

FILL_HEADER = PatternFill("solid", start_color=C_DARK_BLUE)
FILL_TITLE  = PatternFill("solid", start_color=C_MID_BLUE)
FILL_TOTAL  = PatternFill("solid", start_color=C_LIGHT_BLUE)
FILL_ALT    = PatternFill("solid", start_color=C_ALT)
FILL_WHITE  = PatternFill("solid", start_color=C_WHITE)
FILL_GOLD   = PatternFill("solid", start_color=C_GOLD)

TITLE_ROW, HEADER_ROW, DATA_START = 1, 2, 3
NUM_COLS = len(HEADERS)


def _cell(ws, row, col, value=None):
    c = ws.cell(row=row, column=col, value=value)
    return c


def _setup_detail_sheet(ws, title: str):
    # Title row (merged)
    ws.merge_cells(start_row=TITLE_ROW, start_column=1, end_row=TITLE_ROW, end_column=NUM_COLS)
    tc = ws.cell(row=TITLE_ROW, column=1, value=title)
    tc.font = FONT_TITLE
    tc.fill = FILL_TITLE
    tc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[TITLE_ROW].height = 36

    # Header row
    for col, (h, w) in enumerate(zip(HEADERS, COL_WIDTHS), start=1):
        c = ws.cell(row=HEADER_ROW, column=col, value=h)
        c.font = FONT_HEADER
        c.fill = FILL_HEADER
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BORDER
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[HEADER_ROW].height = 22
    ws.freeze_panes = f"A{DATA_START}"


def _build_summary_sheet(ws, title: str):
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 18

    # Title
    ws.merge_cells("A1:B1")
    t = ws.cell(row=1, column=1, value=f"{title} — 汇总")
    t.font = FONT_TITLE
    t.fill = FILL_TITLE
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    d = f"'{DETAIL_SHEET}'"
    sections = [
        ("总览", [
            ("总金额（元）", f"=SUMIF({d}!D:D,\"<>待补\",{d}!D:D)"),
            ("总条数",       f"=COUNTA({d}!A:A)-1"),
        ]),
        ("按付款方式", [
            ("对公合计", f"=SUMIF({d}!G:G,\"对公\",{d}!D:D)"),
            ("对私合计", f"=SUMIF({d}!G:G,\"对私\",{d}!D:D)"),
        ]),
        ("按发票状态", [
            ("有发票合计", f"=SUMIF({d}!H:H,\"有发票\",{d}!D:D)"),
            ("无发票合计", f"=SUMIF({d}!H:H,\"无发票\",{d}!D:D)"),
            ("待补合计",   f"=SUMIF({d}!H:H,\"待补\",{d}!D:D)"),
        ]),
        ("按费用类型", [(et, f"=SUMIF({d}!E:E,\"{et}\",{d}!D:D)") for et in EXPENSE_TYPES]),
    ]

    row = 2
    for section_name, items in sections:
        # Section header
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        sh = ws.cell(row=row, column=1, value=f"  {section_name}")
        sh.font = Font(bold=True, color=C_WHITE, size=11, name="微软雅黑")
        sh.fill = PatternFill("solid", start_color=C_MID_BLUE)
        sh.alignment = Alignment(vertical="center")
        ws.row_dimensions[row].height = 20
        row += 1

        for label, formula in items:
            a = ws.cell(row=row, column=1, value=label)
            b = ws.cell(row=row, column=2, value=formula)
            fill = FILL_GOLD if label == "总金额（元）" else (FILL_ALT if row % 2 == 0 else FILL_WHITE)
            for c in (a, b):
                c.fill = fill
                c.border = BORDER
                c.alignment = Alignment(vertical="center", horizontal="right" if c.column == 2 else "left")
            a.font = FONT_TOTAL if label == "总金额（元）" else FONT_LABEL
            b.font = FONT_TOTAL if label == "总金额（元）" else FONT_BODY
            b.number_format = AMOUNT_FMT
            ws.row_dimensions[row].height = 18
            row += 1
        row += 1  # blank gap between sections


def _create_workbook(path: Path, title: str):
    wb = Workbook()
    ws_detail = wb.active
    ws_detail.title = DETAIL_SHEET
    _setup_detail_sheet(ws_detail, title)

    ws_sum = wb.create_sheet(SUMMARY_SHEET)
    _build_summary_sheet(ws_sum, title)

    wb.save(path)


def _next_row(ws) -> int:
    return ws.max_row + 1 if ws.max_row >= DATA_START else DATA_START


def append_rows(file_path: str, title: str, records: list) -> dict:
    path = Path(file_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    if not path.exists():
        _create_workbook(path, title)

    wb = load_workbook(path)
    ws = wb[DETAIL_SHEET] if DETAIL_SHEET in wb.sheetnames else wb.active

    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    start_row = _next_row(ws)

    for i, rec in enumerate(records):
        row_num = start_row + i
        seq = row_num - DATA_START + 1

        expense_type = rec.get("费用类型", "其他")
        # 劳务费等天然无发票类型自动填充
        invoice = "无发票" if expense_type in NO_INVOICE_TYPES else rec.get("发票状态", "待补")

        values = [
            seq,
            now,
            rec.get("发生日期", "待补"),
            rec.get("金额", "待补"),
            expense_type,
            rec.get("事由", "待补"),
            rec.get("付款方式", "待补"),
            invoice,
            rec.get("经办人", ""),
            rec.get("备注", ""),
        ]
        fill = FILL_ALT if row_num % 2 == 0 else FILL_WHITE
        for col, val in enumerate(values, start=1):
            c = ws.cell(row=row_num, column=col, value=val)
            c.border = BORDER
            c.fill = fill
            c.font = FONT_BODY
            c.alignment = Alignment(vertical="center", wrap_text=(col == 6))
            if col == 4 and isinstance(val, (int, float)):
                c.number_format = AMOUNT_FMT
                c.alignment = Alignment(vertical="center", horizontal="right")
        ws.row_dimensions[row_num].height = 18

    wb.save(path)
    return {"added": len(records), "total_rows": ws.max_row - DATA_START + 1, "file": str(path)}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--file", required=True)
    parser.add_argument("--title", default="报账明细")
    parser.add_argument("--data", required=True)
    args = parser.parse_args()

    records = json.loads(args.data)
    result = append_rows(args.file, args.title, records)
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
