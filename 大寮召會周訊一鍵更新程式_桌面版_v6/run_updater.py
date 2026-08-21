import re

from lxml import html

import update_bulletin as app


def visible_lines(raw: bytes) -> list[str]:
    root = html.fromstring(raw)
    for node in root.xpath("//script|//style|//noscript"):
        node.drop_tree()
    nodes = root.xpath("//p[contains(@class,'zfr3Q')] | //h1 | //h2 | //h3")
    return [line for line in (app.clean("".join(node.itertext())) for node in nodes) if line]


def split_numbered_items(text: str) -> list[str]:
    parts = re.split(r"(?<!\d)(?=\d+[.]\s*)", text)
    return [re.sub(r"^\d+[.]\s*", "", part).strip() for part in parts if part.strip()]


def parse_traffic(lines: list[str]) -> list[str]:
    start = app.find_index(lines, "召會生活動態交通：") + 1
    end = app.find_index(lines, "每日靈糧", start)
    output: list[str] = []
    category = ""
    for original in lines[start:end]:
        line = original
        if line.startswith("★"):
            match = re.match(r"★([^：:]+)[：:]\s*(.*)", line)
            if not match:
                continue
            category, line = match.group(1).strip(), match.group(2).strip()
            if not line:
                continue
        if category == "照顧區":
            south = re.search(
                r"(?:^|\d+[.、]\s*)南B區[：:]\s*(.*?)(?=(?:\d+[.、]\s*)?(?:東區|中區|南A區|南C區|南D區|北區)[：:]|$)",
                line,
                re.I,
            )
            if south and south.group(1).strip():
                output.extend(split_numbered_items(south.group(1).strip()))
            continue
        for item in split_numbered_items(line):
            item = re.sub(r"^[①②③④⑤⑥⑦⑧⑨⑩]\s*", "", item).lstrip("★").strip()
            if len(item) > 3:
                output.append(item)
    return output


def parse_reading_progress(lines: list[str]) -> tuple[str, int, str, str]:
    daily = app.find_index(lines, "每日靈糧")
    prayer = app.find_index(lines, "禱", daily)
    header = next((x for x in lines[daily:prayer] if "晨興聖言" in x and "週" in x), "")
    match = re.search(r"[「『](.+?)[」』]晨興聖言第([一二三四五六七八九十\d]+)週", header)
    if not match:
        raise ValueError("無法辨識晨興聖言週次。")
    series = match.group(1).replace("（", "(").replace("）", ")")
    week_text = match.group(2)
    week = int(week_text) if week_text.isdigit() else app.CN_NUM.get(week_text, 0)
    progress = re.findall(r"讀經進度[：:]\s*([^週主\s]+)", " ".join(lines[prayer:]))
    progress = [x for x in progress if "複習" not in x]
    if len(progress) < 2:
        raise ValueError("讀經進度不足，無法計算起訖。")
    return series, week, progress[0].split("~", 1)[0], progress[-1].split("~", 1)[-1]


app.visible_lines = visible_lines
app.parse_traffic = parse_traffic
app.parse_reading_progress = parse_reading_progress


if __name__ == "__main__":
    try:
        raise SystemExit(app.main())
    except Exception as exc:
        print(f"\n更新失敗：{exc}", file=app.sys.stderr)
        print("未覆寫任何原始檔。", file=app.sys.stderr)
        raise SystemExit(1)

