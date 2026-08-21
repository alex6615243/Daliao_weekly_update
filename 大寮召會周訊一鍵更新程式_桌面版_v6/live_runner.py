import time
import urllib.parse
import urllib.request

import ultimate_runner


app = ultimate_runner.app


def download_without_cache(url, cache_path=None):
    if cache_path:
        return cache_path.read_bytes()
    parsed = urllib.parse.urlsplit(url)
    query = urllib.parse.parse_qsl(parsed.query, keep_blank_values=True)
    query.append(("codex_refresh", str(int(time.time()))))
    refreshed_url = urllib.parse.urlunsplit(
        (parsed.scheme, parsed.netloc, parsed.path, urllib.parse.urlencode(query), parsed.fragment)
    )
    request = urllib.request.Request(
        refreshed_url,
        headers={
            "User-Agent": "Mozilla/5.0 WeeklyBulletinUpdater/1.1",
            "Cache-Control": "no-cache, no-store, max-age=0",
            "Pragma": "no-cache",
        },
    )
    with urllib.request.urlopen(request, timeout=30) as response:
        return response.read()


app.download = download_without_cache


if __name__ == "__main__":
    try:
        raise SystemExit(app.main())
    except Exception as exc:
        print(f"\n更新失敗：{exc}", file=app.sys.stderr)
        print("未覆寫任何原始檔。", file=app.sys.stderr)
        raise SystemExit(1)

