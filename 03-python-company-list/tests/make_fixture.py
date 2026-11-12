"""テスト・動作確認用の差分データ（架空の法人）を作る。

実データと同じ列構成で、表記ゆれ・重複・番号誤り・新規以外の行などをわざと混ぜている。
実行：python tests/make_fixture.py
"""
import csv
import io
import random
import zipfile
from pathlib import Path

OUT = Path(__file__).parent / "fixtures"


def with_check_digit(body12: str) -> str:
    total = sum(int(d) * (1 if n % 2 else 2) for n, d in enumerate(reversed(body12), start=1))
    return str(9 - total % 9) + body12


def row(seq, number, process, name, kind, pref, city, street, post, assigned, updated):
    cells = [""] * 30
    cells[0], cells[1], cells[2], cells[3] = str(seq), number, process, "0"
    cells[4], cells[5], cells[6], cells[8] = updated, updated, name, kind
    cells[9], cells[10], cells[11], cells[15] = pref, city, street, post
    cells[22], cells[23] = assigned, "1"
    return cells


def main():
    rnd = random.Random(20261112)
    words = ["みらい", "あおぞら", "つばさ", "ひかり", "さくら", "かえで", "いろは", "そよかぜ", "あさひ", "ことのは"]
    trades = ["商事", "デザイン", "テクノロジー", "フーズ", "ホームサービス", "コンサルティング", "ラボ", "建設"]
    places = [("東京都", "渋谷区", "150"), ("東京都", "新宿区", "160"), ("東京都", "港区", "105"),
              ("神奈川県", "横浜市中区", "231"), ("神奈川県", "川崎市中原区", "211"), ("大阪府", "大阪市北区", "530"),
              ("埼玉県", "さいたま市大宮区", "330")]
    kinds = [("301", "株式会社{}"), ("305", "合同会社{}"), ("301", "株式会社{}")]

    rows, seq = [], 1
    for i in range(36):
        body = f"{rnd.randint(10**11, 10**12 - 1)}"
        num = with_check_digit(body)
        kind_code, fmt = rnd.choice(kinds)
        name = fmt.format(f"サンプル{rnd.choice(words)}{rnd.choice(trades)}")
        pref, city, zip3 = rnd.choice(places)
        n1, n2 = rnd.randint(1, 5), rnd.randint(1, 30)
        street = f"{rnd.choice(['本町', '中央', '栄町', '桜木町'])}{n1}丁目{n2}番{rnd.randint(1, 20)}号"
        if i % 5 == 0:  # 全角の数字・空白の表記ゆれ
            street = street.translate(str.maketrans("0123456789", "０１２３４５６７８９")) + "　サンプルビル　３Ｆ"
        post = f"{zip3}{rnd.randint(1000, 9999)}"
        day = f"2026-11-{rnd.randint(6, 12):02d}"
        rows.append(row(seq, num, "01", name, kind_code, pref, city, street, post, day, day)); seq += 1
        if i in (3, 17):  # 同じ法人番号の再送（古い更新日）→ 重複として除外
            rows.append(row(seq, num, "01", name + "（旧）", kind_code, pref, city, street, post, day, "2026-11-01")); seq += 1

    # 新規以外の行（商号変更・閉鎖）
    for process in ("12", "21", "71"):
        rows.append(row(seq, with_check_digit(f"{rnd.randint(10**11, 10**12 - 1)}"), process,
                        "株式会社サンプル既存商事", "301", "東京都", "千代田区", "丸の内1丁目1番1号", "1000005",
                        "2015-04-01", "2026-11-10")); seq += 1
    # チェックデジットが誤っている行
    bad = with_check_digit("123456789012")
    rows.append(row(seq, str((int(bad[0]) + 1) % 10) + bad[1:], "01", "株式会社サンプル入力ミス", "301",
                    "東京都", "渋谷区", "道玄坂1丁目1番1号", "1500043", "2026-11-11", "2026-11-11")); seq += 1
    rnd.shuffle(rows)

    buf = io.StringIO()
    csv.writer(buf, lineterminator="\r\n").writerows(rows)
    OUT.mkdir(exist_ok=True)
    (OUT / "diff_sample.csv").write_text(buf.getvalue(), encoding="utf-8-sig")
    with zipfile.ZipFile(OUT / "diff_sample.zip", "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("diff_sample.csv", "﻿" + buf.getvalue())
    print(f"{len(rows)}行を書き出しました")


if __name__ == "__main__":
    main()
