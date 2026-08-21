import re

import run_updater


app = run_updater.app


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
    start_value = progress[0].split("~", 1)[0]
    last_value = progress[-1]
    if "~" in last_value:
        left, right = last_value.split("~", 1)
        prefix = re.match(r"(.+?)(\d+)$", left)
        end_value = (prefix.group(1) if prefix else "") + right
    else:
        end_value = last_value
    return series, week, start_value, end_value


app.parse_reading_progress = parse_reading_progress


if __name__ == "__main__":
    try:
        raise SystemExit(app.main())
    except Exception as exc:
        print(f"\n更新失敗：{exc}", file=app.sys.stderr)
        print("未覆寫任何原始檔。", file=app.sys.stderr)
        raise SystemExit(1)

