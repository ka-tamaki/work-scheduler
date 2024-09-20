# generate_weekends.py

import json
from datetime import datetime
import calendar
import os

def generate_weekends(start_year, end_year):
    weekends = {}
    for year in range(start_year, end_year + 1):
        weekends[str(year)] = {}
        for month in range(1, 13):
            weekends[str(year)][str(month)] = []
            cal = calendar.Calendar()
            for day, weekday in cal.itermonthdays2(year, month):
                if day == 0:
                    continue
                if weekday >= 5:  # 土曜日=5, 日曜日=6
                    weekends[str(year)][str(month)].append(day)
    return weekends

def merge_holidays(existing_holidays, weekends):
    for year, months in weekends.items():
        if year not in existing_holidays:
            existing_holidays[year] = {}
        for month, days in months.items():
            if month not in existing_holidays[year]:
                existing_holidays[year][month] = []
            for day in days:
                if day not in existing_holidays[year][month]:
                    existing_holidays[year][month].append(day)
    # ソートして重複を防ぐ
    for year, months in existing_holidays.items():
        for month, days in months.items():
            existing_holidays[year][month] = sorted(days)
    return existing_holidays

def main():
    current_year = datetime.now().year
    end_year = 2030
    weekends = generate_weekends(current_year, end_year)
    
    holidays_file = 'holidays.json'
    
    if os.path.exists(holidays_file):
        with open(holidays_file, 'r', encoding='utf-8') as f:
            try:
                existing_holidays = json.load(f)
            except json.JSONDecodeError:
                print("holidays.jsonが正しくフォーマットされていません。新規に作成します。")
                existing_holidays = {}
    else:
        existing_holidays = {}
    
    updated_holidays = merge_holidays(existing_holidays, weekends)
    
    with open(holidays_file, 'w', encoding='utf-8') as f:
        json.dump(updated_holidays, f, ensure_ascii=False, indent=4)
    
    print(f"holidays.jsonに{current_year}年から{end_year}年までの土日が追加されました。")

if __name__ == "__main__":
    main()
