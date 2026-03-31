#!/usr/bin/env python3

import argparse
import json
from datetime import date, timedelta
from pathlib import Path

BASE_RULES = {
    "MONDAY": ["BURNABLE"],
    "THURSDAY": ["BURNABLE"],
    "FRIDAY": ["RECYCLE_PLASTIC"],
    "TUESDAY_1": ["OTHER_PLASTIC"],
    "TUESDAY_2": ["RECYCLABLE_HAZARDOUS"],
    "TUESDAY_3": ["OTHER_PLASTIC"],
    "TUESDAY_4": ["RECYCLABLE_HAZARDOUS"],
    "WEDNESDAY_1": ["LARGE_WASTE"],
    "WEDNESDAY_2": ["NON_BURNABLE"],
    "WEDNESDAY_3": ["LARGE_WASTE"],
    "WEDNESDAY_4": ["NON_BURNABLE"],
}


def expected_types(target_date):
    nth = (target_date.day - 1) // 7 + 1
    weekday = target_date.weekday()

    if weekday == 0:
        return BASE_RULES["MONDAY"]
    if weekday == 3:
        return BASE_RULES["THURSDAY"]
    if weekday == 4:
        return BASE_RULES["FRIDAY"]
    if weekday == 1 and nth in (1, 2, 3, 4):
        return BASE_RULES[f"TUESDAY_{nth}"]
    if weekday == 2 and nth in (1, 2, 3, 4):
        return BASE_RULES[f"WEDNESDAY_{nth}"]
    return []


def date_range(start, end):
    cursor = start
    while cursor <= end:
        yield cursor
        cursor += timedelta(days=1)


def main():
    parser = argparse.ArgumentParser(description="Validate imported trash schedule data.")
    parser.add_argument("schedule_json", help="Path to generated schedule JSON")
    args = parser.parse_args()

    payload = json.loads(Path(args.schedule_json).read_text(encoding="utf-8"))
    days = payload["days"]
    start = date.fromisoformat(payload["dateRange"]["start"])
    end = date.fromisoformat(payload["dateRange"]["end"])

    missing_dates = []
    anomalies = []

    for current in date_range(start, end):
        key = current.isoformat()
        actual = days.get(key)
        if actual is None:
            missing_dates.append(key)
            continue

        expected = expected_types(current)
        if actual != expected:
            anomalies.append({
                "date": key,
                "expected": expected,
                "actual": actual,
            })

    if missing_dates:
        print("Missing dates:")
        for key in missing_dates:
            print(f"  {key}")

    print(f"Checked {len(days)} dates from {start} to {end}")
    print(f"Detected {len(anomalies)} rule deviations")

    for item in anomalies:
        print(f"{item['date']}: expected={item['expected']} actual={item['actual']}")

    if missing_dates:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
