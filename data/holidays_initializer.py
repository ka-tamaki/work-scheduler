# data/holidays_initializer.py
import json
import calendar
from datetime import datetime, timedelta
import os
import jpholiday# 日本の祝日を取得

class HolidaysInitializer:
    def __init__(self, years=20):
        self.years = years
        self.today = datetime.today()
        self.start_year = self.today.year
        self.end_year = self.start_year + self.years

    def generate_holidays(self):
        for factory in ['yuki', 'kumagaya', 'shizuoka', 'kyoto', 'chiba']:
            holidays = []
            for year in range(self.start_year, self.end_year + 1):
                for month in range(1, 13):
                    for day in range(1, 32):
                        try:
                            date = datetime(year, month, day)
                            if date.weekday() >= 5 or jpholiday.is_holiday(date):
                                holidays.append({
                                    'date': date.strftime('%Y-%m-%d'),
                                    'weekday': calendar.day_name[date.weekday()]
                                })
                        except ValueError:
                            continue
            # 保存先パス
            filepath = os.path.join('holidays', f'{factory}.json')
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(holidays, f, ensure_ascii=False, indent=4)

if __name__ == "__main__":
    initializer = HolidaysInitializer()
    initializer.generate_holidays()
    print("休日情報の初期化が完了しました。")
