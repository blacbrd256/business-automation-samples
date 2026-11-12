from pathlib import Path

import pytest
from openpyxl import load_workbook

from company_list import cli
from company_list.fetch import PoliteFetcher, load_text
from company_list.houjin import (
    extract_new_companies, is_valid_corporate_number, normalize_post_code, normalize_text, read_rows,
)

FIX = Path(__file__).parent / "fixtures"


def test_check_digit_with_real_number():
    # 国税庁自身の法人番号（公開情報）
    assert is_valid_corporate_number("7000012050002")
    assert not is_valid_corporate_number("8000012050002")
    assert not is_valid_corporate_number("700001205000")  # 12桁
    assert not is_valid_corporate_number("70000120500O2")  # 英字混入


def test_normalize():
    assert normalize_text("中央１丁目２番　サンプルビル　３Ｆ") == "中央1丁目2番 サンプルビル 3F"
    assert normalize_post_code("１５００００２") == "150-0002"
    assert normalize_post_code("15000") == ""


def test_extract_from_fixture():
    rows = list(read_rows(load_text(str(FIX / "diff_sample.zip"))))
    companies, stats = extract_new_companies(rows)
    assert stats["total"] == 42
    assert stats["not_new"] == 3
    assert stats["invalid_number"] == 1
    assert stats["duplicates"] == 2
    assert stats["kept"] == 36
    assert len({c.corporate_number for c in companies}) == len(companies)
    assert all("（旧）" not in c.name for c in companies), "古い更新日の重複は捨てる"
    assert all(not any(ch in c.street for ch in "０１２３４５６７８９　") for c in companies)
    dates = [c.assignment_date for c in companies]
    assert dates == sorted(dates, reverse=True), "新しい順に並ぶ"


def test_filter_by_prefecture_and_kind():
    rows = list(read_rows(load_text(str(FIX / "diff_sample.csv"))))
    companies, stats = extract_new_companies(rows, prefectures={"東京都"}, kinds={"305"})
    assert companies and all(c.prefecture == "東京都" and c.kind == "合同会社" for c in companies)
    assert stats["filtered"] > 0


def test_cli_outputs(tmp_path, capsys):
    assert cli.main([str(FIX / "diff_sample.zip"), "--pref", "東京都", "神奈川県", "-o", str(tmp_path)]) == 0
    csv_text = (tmp_path / "new_companies.csv").read_bytes()
    assert csv_text.startswith(b"\xef\xbb\xbf"), "Excelで文字化けしないBOM付き"
    wb = load_workbook(tmp_path / "new_companies.xlsx")
    assert wb.sheetnames == ["新設法人リスト", "集計"]
    assert wb["新設法人リスト"]["A1"].value == "法人番号"
    assert "新設法人" in capsys.readouterr().out


def test_cli_rejects_unknown_kind(tmp_path):
    with pytest.raises(SystemExit):
        cli.main([str(FIX / "diff_sample.zip"), "--kind", "株式会社X", "-o", str(tmp_path)])


def test_fetcher_waits_and_retries():
    calls, sleeps = [], []

    class Res:
        def __enter__(self): return self
        def __exit__(self, *a): return False
        def read(self): return b"ok"

    def opener(req, timeout):
        calls.append(req.full_url)
        if len(calls) < 3:
            raise OSError("一時的なエラー")
        return Res()

    f = PoliteFetcher(min_interval=3.0, retries=3, opener=opener, sleep=sleeps.append)
    assert f.get("https://example.com/data.zip") == b"ok"
    assert len(calls) == 3
    assert sleeps.count(2) == 1 and sleeps.count(4) == 1, "失敗後は2秒→4秒と待つ"


def test_fetcher_gives_up():
    def opener(req, timeout):
        raise OSError("down")

    f = PoliteFetcher(min_interval=0, retries=2, opener=opener, sleep=lambda s: None)
    with pytest.raises(RuntimeError, match="2回失敗"):
        f.get("https://example.com/x")
