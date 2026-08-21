import re

import live_runner


app = live_runner.app
original_parse_page = app.parse_page


def chinese_number(text: str) -> int:
    digits = {"零": 0, "一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
    if text == "十":
        return 10
    if "十" in text:
        left, right = text.split("十", 1)
        return digits.get(left, 1) * 10 + digits.get(right, 0)
    if all(char in digits for char in text):
        return int("".join(str(digits[char]) for char in text))
    return 0


def cn_chapter_to_arabic(text: str):
    match = re.match(r"([\u4e00-\u9fff]+?)([零一二三四五六七八九十百]+)(\d+)$", text)
    if not match:
        raise ValueError(f"無法辨識讀經範圍：{text}")
    book = app.BOOK_NAMES.get(match.group(1), match.group(1))
    return book, chinese_number(match.group(2)), int(match.group(3))


def parse_page_with_wrapped_url_fix(raw):
    data = original_parse_page(raw)
    merged = []
    for item in data.traffic_items:
        if merged and re.search(r"https://[A-Za-z0-9-]+$", merged[-1]) and re.match(r"(?:github\.io|com|org|net)/", item):
            merged[-1] += "." + item
        else:
            merged.append(item)
    data.traffic_items = merged
    return data


app.cn_chapter_to_arabic = cn_chapter_to_arabic
app.parse_page = parse_page_with_wrapped_url_fix


if __name__ == "__main__":
    try:
        raise SystemExit(app.main())
    except Exception as exc:
        print(f"\n更新失敗：{exc}", file=app.sys.stderr)
        print("未覆寫任何原始檔。", file=app.sys.stderr)
        raise SystemExit(1)

