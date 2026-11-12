"""差分データを取得する。URL でもローカルのファイル（CSV / ZIP）でも受け付ける。

公開データの取得でも相手のサーバーに負担をかけないよう、
リクエストの間隔を空け、失敗時は待ってから再試行する。
"""
from __future__ import annotations

import io
import time
import urllib.request
import zipfile
from pathlib import Path

USER_AGENT = "company-list-sample/1.0 (+https://github.com/)"


class PoliteFetcher:
    def __init__(self, min_interval: float = 3.0, retries: int = 3, opener=urllib.request.urlopen, sleep=time.sleep):
        self.min_interval = min_interval
        self.retries = retries
        self._opener = opener
        self._sleep = sleep
        self._last = 0.0

    def get(self, url: str) -> bytes:
        last_error: Exception | None = None
        for attempt in range(1, self.retries + 1):
            wait = self.min_interval - (time.monotonic() - self._last)
            if wait > 0:
                self._sleep(wait)
            self._last = time.monotonic()
            try:
                req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
                with self._opener(req, timeout=60) as res:
                    return res.read()
            except Exception as e:  # noqa: BLE001 - 通信エラーはまとめて再試行
                last_error = e
                self._sleep(2 ** attempt)
        raise RuntimeError(f"{url} の取得に{self.retries}回失敗しました: {last_error}")


def load_text(source: str, fetcher: PoliteFetcher | None = None) -> str:
    """URLまたはファイルパスから、差分データのCSVテキストを取り出す。"""
    if source.startswith(("http://", "https://")):
        data = (fetcher or PoliteFetcher()).get(source)
    else:
        data = Path(source).read_bytes()
    if data[:2] == b"PK":  # ZIPで配布されている場合は中のCSVを読む
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            name = next(n for n in z.namelist() if n.lower().endswith(".csv"))
            data = z.read(name)
    return data.decode("utf-8-sig")
