"""コマンドラインから実行する入口。

例：
  python -m company_list.cli data/diff_20261112.zip --pref 東京都 神奈川県 --kind 株式会社 合同会社 -o out
"""
from __future__ import annotations

import argparse
from pathlib import Path

from .export import to_csv, to_excel
from .fetch import load_text
from .houjin import KIND_LABELS, extract_new_companies, read_rows


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="法人番号の差分データから新設法人リストを作成します")
    p.add_argument("sources", nargs="+", help="差分データのURLまたはファイル（CSV/ZIP）。複数日分を指定可")
    p.add_argument("--pref", nargs="*", help="対象の都道府県（例：東京都 大阪府）")
    p.add_argument("--kind", nargs="*", help="対象の法人種別（例：株式会社 合同会社）")
    p.add_argument("-o", "--out", default="out", help="出力先フォルダ")
    a = p.parse_args(argv)

    kind_codes = None
    if a.kind:
        reverse = {v: k for k, v in KIND_LABELS.items()}
        unknown = [k for k in a.kind if k not in reverse]
        if unknown:
            p.error(f"不明な法人種別: {', '.join(unknown)}（指定できる値: {', '.join(reverse)}）")
        kind_codes = {reverse[k] for k in a.kind}

    rows = [row for src in a.sources for row in read_rows(load_text(src))]
    companies, stats = extract_new_companies(rows, set(a.pref) if a.pref else None, kind_codes)

    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    to_csv(companies, out / "new_companies.csv")
    to_excel(companies, stats, out / "new_companies.xlsx")

    print(f"読み込み {stats['total']}行 → 新設法人 {stats['kept']}社を出力しました（{out}）")
    for k, label in [("not_new", "新規以外"), ("invalid_number", "番号誤り"), ("filtered", "対象外"), ("duplicates", "重複")]:
        print(f"  {label}: {stats[k]}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
