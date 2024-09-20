# utils/holidays_initializer.py

from datetime import datetime
import calendar
import jpholiday
from data.holidays import HolidaysManager

def generate_weekends_and_holidays(start_year, end_year):
    """土日と祝日を生成"""
    holidays = {}
    for year in range(start_year, end_year + 1):
        holidays[str(year)] = {}
        for month in range(1, 13):
            holidays[str(year)][str(month)] = []
            cal = calendar.Calendar(firstweekday=0)  # 月曜日=0
            for day, weekday in cal.itermonthdays2(year, month):
                if day == 0:
                    continue
                date_obj = datetime(year, month, day)
                # 土日
                if weekday >= 5:
                    holidays[str(year)][str(month)].append(day)
                # 祝日
                if jpholiday.is_holiday(date_obj):
                    if day not in holidays[str(year)][str(month)]:
                        holidays[str(year)][str(month)].append(day)
    return holidays

def merge_holidays(existing_holidays, new_holidays):
    """既存の休日情報に新しい休日情報をマージ"""
    for year, months in new_holidays.items():
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

def initialize_holidays():
    """holidays.jsonを初期化または更新"""
    manager = HolidaysManager()
    existing_holidays = manager.holidays

    current_year = datetime.now().year
    end_year = 2030
    new_holidays = generate_weekends_and_holidays(current_year, end_year)

    updated_holidays = merge_holidays(existing_holidays, new_holidays)

    manager.holidays = updated_holidays
    manager.save_holidays()

    print(f"holidays.jsonに{current_year}年から{end_year}年までの土日と祝日が追加されました。")
