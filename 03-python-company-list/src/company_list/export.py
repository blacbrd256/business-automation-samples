"""新設法人リストをCSV（Excelで文字化けしないBOM付きUTF-8）とExcelに書き出す。"""
from __future__ import annotations

import csv
from collections import Counter
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .houjin import Company

HEADERS = [
    ("corporate_number", "法人番号"), ("name", "商号"), ("kind", "法人種別"), ("post_code", "郵便番号"),
    ("address", "所在地"), ("prefecture", "都道府県"), ("city", "市区町村"),
    ("assignment_date", "法人番号指定日"), ("update_date", "更新日"),
]


def to_csv(companies: list[Company], path: Path) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow([label for _, label in HEADERS])
        for c in companies:
            row = c.to_row()
            w.writerow([row[k] for k, _ in HEADERS])


def to_excel(companies: list[Company], stats: dict, path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "新設法人リスト"
    ws.append([label for _, label in HEADERS])
    for c in companies:
        row = c.to_row()
        ws.append([row[k] for k, _ in HEADERS])

    head_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center")
    for col, width in zip("ABCDEFGHI", [16, 30, 12, 10, 44, 10, 14, 14, 12]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}{ws.max_row}"

    s = wb.create_sheet("集計")
    s.append(["項目", "件数"])
    for label, key in [("読み込んだ行", "total"), ("新規以外（変更・閉鎖など）", "not_new"),
                       ("法人番号の誤り", "invalid_number"), ("対象外の地域・種別", "filtered"),
                       ("重複をまとめた件数", "duplicates"), ("リストに残した件数", "kept")]:
        s.append([label, stats.get(key, 0)])
    s.append([])
    s.append(["都道府県別", "件数"])
    for pref, n in Counter(c.prefecture for c in companies).most_common():
        s.append([pref, n])
    s.append([])
    s.append(["法人種別", "件数"])
    for kind, n in Counter(c.kind for c in companies).most_common():
        s.append([kind, n])
    s.column_dimensions["A"].width = 28
    for row in s.iter_rows():
        if row[0].value in ("項目", "都道府県別", "法人種別"):
            for cell in row:
                cell.font = Font(bold=True)
    wb.save(path)
