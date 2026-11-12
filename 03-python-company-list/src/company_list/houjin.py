"""国税庁 法人番号公表サイトの「差分データ（CSV・Unicode）」を読み、新設法人のリストを作る。

差分データは1日ごとに公開され、1行が1件の変更を表す。
ここでは「新規」の行だけを取り出し、法人番号で重複をまとめ、表記を整える。
"""
from __future__ import annotations

import csv
import io
import re
import unicodedata
from dataclasses import dataclass, asdict
from datetime import date
from typing import Iterable, Iterator

# 差分データ（CSV形式・Unicode）の列順
COLUMNS = [
    "sequence_number", "corporate_number", "process", "correct", "update_date", "change_date",
    "name", "name_image_id", "kind", "prefecture_name", "city_name", "street_number",
    "address_image_id", "prefecture_code", "city_code", "post_code", "address_outside",
    "address_outside_image_id", "close_date", "close_cause", "successor_corporate_number",
    "change_cause", "assignment_date", "latest", "en_name", "en_prefecture_name",
    "en_city_name", "en_address_outside", "furigana", "hihyoji",
]

PROCESS_NEW = "01"  # 処理区分：新規

KIND_LABELS = {
    "101": "国の機関", "201": "地方公共団体", "301": "株式会社", "302": "有限会社",
    "303": "合名会社", "304": "合資会社", "305": "合同会社", "399": "その他の設立登記法人",
    "401": "外国会社等", "499": "その他",
}


@dataclass(frozen=True)
class Company:
    corporate_number: str
    name: str
    kind: str
    prefecture: str
    city: str
    street: str
    post_code: str
    assignment_date: date | None
    update_date: date | None

    @property
    def address(self) -> str:
        return f"{self.prefecture}{self.city}{self.street}"

    def to_row(self) -> dict:
        d = asdict(self)
        d["address"] = self.address
        d["assignment_date"] = self.assignment_date.isoformat() if self.assignment_date else ""
        d["update_date"] = self.update_date.isoformat() if self.update_date else ""
        return d


def is_valid_corporate_number(number: str) -> bool:
    """法人番号（13桁）のチェックデジットを検証する。

    先頭1桁がチェックデジット。残り12桁を右から n=1..12 として
    奇数桁は×1、偶数桁は×2 で合計し、9 −（合計 mod 9）が先頭と一致すれば正しい。
    """
    if not re.fullmatch(r"\d{13}", number or ""):
        return False
    body = number[1:]
    total = sum(int(d) * (1 if n % 2 else 2) for n, d in enumerate(reversed(body), start=1))
    return int(number[0]) == 9 - total % 9


def normalize_text(value: str) -> str:
    """全角英数・記号を半角に、連続する空白を1つにそろえる（NFKC）。"""
    v = unicodedata.normalize("NFKC", value or "")
    return re.sub(r"\s+", " ", v).strip()


def normalize_post_code(value: str) -> str:
    digits = re.sub(r"\D", "", normalize_text(value))
    return f"{digits[:3]}-{digits[3:]}" if len(digits) == 7 else ""


def _parse_date(value: str) -> date | None:
    try:
        return date.fromisoformat(value) if value else None
    except ValueError:
        return None


def read_rows(text: str) -> Iterator[dict]:
    """CSVテキストを列名付きの辞書として読み出す。列数が合わない行は飛ばす。"""
    for cells in csv.reader(io.StringIO(text.lstrip("﻿"))):
        if len(cells) < len(COLUMNS):
            continue
        yield dict(zip(COLUMNS, cells[: len(COLUMNS)]))


def extract_new_companies(
    rows: Iterable[dict],
    prefectures: set[str] | None = None,
    kinds: set[str] | None = None,
) -> tuple[list[Company], dict]:
    """新規設立の法人だけを取り出し、重複をまとめて返す。

    戻り値の2つ目は処理結果の内訳（件数）。
    """
    stats = {"total": 0, "not_new": 0, "invalid_number": 0, "filtered": 0, "duplicates": 0, "kept": 0}
    latest: dict[str, Company] = {}

    for r in rows:
        stats["total"] += 1
        if r["process"] != PROCESS_NEW:
            stats["not_new"] += 1
            continue
        number = normalize_text(r["corporate_number"])
        if not is_valid_corporate_number(number):
            stats["invalid_number"] += 1
            continue
        prefecture = normalize_text(r["prefecture_name"])
        if (prefectures and prefecture not in prefectures) or (kinds and r["kind"] not in kinds):
            stats["filtered"] += 1
            continue

        company = Company(
            corporate_number=number,
            name=normalize_text(r["name"]),
            kind=KIND_LABELS.get(r["kind"], "その他"),
            prefecture=prefecture,
            city=normalize_text(r["city_name"]),
            street=normalize_text(r["street_number"]),
            post_code=normalize_post_code(r["post_code"]),
            assignment_date=_parse_date(r["assignment_date"]),
            update_date=_parse_date(r["update_date"]),
        )
        prev = latest.get(number)
        if prev is not None:
            stats["duplicates"] += 1
            # 同じ法人番号が複数あれば、更新日の新しい方を残す
            if (company.update_date or date.min) < (prev.update_date or date.min):
                continue
        latest[number] = company

    result = sorted(latest.values(), key=lambda c: (c.assignment_date or date.min, c.corporate_number), reverse=True)
    stats["kept"] = len(result)
    return result, stats
