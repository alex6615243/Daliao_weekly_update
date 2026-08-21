import re
from datetime import date

import final_runner


app = final_runner.app


def latest_item_date(text: str, target: date) -> date | None:
    dates: list[date] = []
    for month, day in re.findall(r"(?<!\d)(\d{1,2})月(\d{1,2})日", text):
        try:
            dates.append(date(target.year, int(month), int(day)))
        except ValueError:
            pass
    for month, _start, end in re.findall(r"(?<!\d)(\d{1,2})月(\d{1,2})(?:~|～|至|-)(\d{1,2})日", text):
        try:
            dates.append(date(target.year, int(month), int(end)))
        except ValueError:
            pass
    return max(dates) if dates else None


app.latest_item_date = latest_item_date


if __name__ == "__main__":
    try:
        raise SystemExit(app.main())
    except Exception as exc:
        print(f"\n更新失敗：{exc}", file=app.sys.stderr)
        print("未覆寫任何原始檔。", file=app.sys.stderr)
        raise SystemExit(1)

