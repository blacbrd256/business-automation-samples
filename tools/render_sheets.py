"""テスト結果(result.json)のシート内容を、スプレッドシート風のPNG画像にする。
使い方: python render_sheets.py result.json out_dir シート名 [シート名...]
"""
import html
import os
import json
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

from playwright.sync_api import sync_playwright

JST = timezone(timedelta(hours=9))


MAXLEN = 60


def fmt(v):
    if v is None:
        return ""
    if isinstance(v, str) and len(v) == 24 and v.endswith("Z") and "T" in v:
        d = datetime.fromisoformat(v.replace("Z", "+00:00")).astimezone(JST)
        return d.strftime("%Y/%m/%d %H:%M")
    s = str(v)
    return s if len(s) <= MAXLEN else s[: MAXLEN - 1] + "…"


def page_html(title, rows):
    cols = max(len(r) for r in rows)
    letters = "ABCDEFGHIJKL"[:cols]
    head = "".join(f"<th class=col>{c}</th>" for c in letters)
    body = []
    for i, r in enumerate(rows, 1):
        cls = " class=hdr" if i == 1 else ""
        cells = "".join(f"<td>{html.escape(fmt(c))}</td>" for c in r + [""] * (cols - len(r)))
        body.append(f"<tr{cls}><th class=num>{i}</th>{cells}</tr>")
    return f"""<!doctype html><meta charset=utf-8>
<style>
body{{margin:0;font-family:'Noto Sans CJK JP','Noto Sans CJK SC',sans-serif;font-size:13px;background:#fff;color:#202124}}
.bar{{background:#f1f3f4;border-bottom:1px solid #dadce0;padding:8px 12px;font-size:12px;color:#5f6368}}
.tab{{display:inline-block;background:#e6f4ea;color:#137333;border-radius:4px 4px 0 0;padding:6px 14px;margin:8px 0 0 12px;font-weight:700}}
table{{border-collapse:collapse;margin:0}}
th,td{{border:1px solid #e2e3e3;padding:5px 9px;white-space:nowrap;text-align:left}}
th.col,th.num{{background:#f8f9fa;color:#5f6368;font-weight:400;text-align:center}}
th.num{{width:28px}}
tr.hdr td{{background:#e8f0fe;font-weight:700}}
</style>
<div class=bar>{html.escape(os.environ.get('CAPTION', 'テスト実行結果（Code.gs をモック環境で実行）'))}</div>
<div class=tab>{html.escape(title)}</div>
<table><tr><th class=num></th>{head}</tr>{''.join(body)}</table>"""


def main():
    global MAXLEN
    src, out_dir, *names = sys.argv[1:]
    if names and names[0].startswith("--max="):
        MAXLEN = int(names.pop(0)[6:])
    data = json.loads(Path(src).read_text(encoding="utf-8"))
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as p:
        b = p.chromium.launch(executable_path="/usr/bin/google-chrome")
        pg = b.new_page(viewport={"width": 1100, "height": 300}, device_scale_factor=2)
        for i, spec in enumerate(names, 1):
            # 「シート名:1,2,6」で表示する列を選べる（1始まり）
            n, _, cols = spec.partition(":")
            rows = data["sheets"][n]
            if cols:
                idx = [int(c) - 1 for c in cols.split(",")]
                rows = [[(r[j] if j < len(r) else "") for j in idx] for r in rows]
            pg.set_content(page_html(n, rows))
            pg.add_style_tag(content="body{display:inline-block;min-width:520px}")
            pg.locator("body").screenshot(path=str(out / f"{i:02d}.png"))
        b.close()


if __name__ == "__main__":
    main()
