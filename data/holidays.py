# data/holidays.py

import json
import os
from config import HOLIDAYS_FILE
from tkinter import messagebox

class HolidaysManager:
    def __init__(self):
        self.holidays = {}
        self.load_holidays()

    def load_holidays(self):
        """holidays.jsonから休日情報をロード"""
        if os.path.exists(HOLIDAYS_FILE):
            with open(HOLIDAYS_FILE, 'r', encoding='utf-8') as f:
                try:
                    self.holidays = json.load(f)
                except json.JSONDecodeError:
                    messagebox.showerror("エラー", "holidays.jsonのフォーマットが正しくありません。")
                    self.holidays = {}
        else:
            self.holidays = {}

    def save_holidays(self):
        """holidays.jsonに休日情報を保存"""
        with open(HOLIDAYS_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.holidays, f, ensure_ascii=False, indent=4)

    def is_holiday(self, year, month, day):
        """指定された日が休日かどうかを判定"""
        year_str = str(year)
        month_str = str(month)
        return day in self.holidays.get(year_str, {}).get(month_str, [])

    def add_holiday(self, year, month, day):
        """休日を追加"""
        year_str = str(year)
        month_str = str(month)
        if year_str not in self.holidays:
            self.holidays[year_str] = {}
        if month_str not in self.holidays[year_str]:
            self.holidays[year_str][month_str] = []
        if day not in self.holidays[year_str][month_str]:
            self.holidays[year_str][month_str].append(day)
            self.save_holidays()

    def remove_holiday(self, year, month, day):
        """休日を削除"""
        year_str = str(year)
        month_str = str(month)
        if year_str in self.holidays and month_str in self.holidays[year_str]:
            if day in self.holidays[year_str][month_str]:
                self.holidays[year_str][month_str].remove(day)
                # 空になった月や年のエントリを削除
                if not self.holidays[year_str][month_str]:
                    del self.holidays[year_str][month_str]
                if not self.holidays[year_str]:
                    del self.holidays[year_str]
                self.save_holidays()
