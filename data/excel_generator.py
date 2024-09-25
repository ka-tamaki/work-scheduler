# data/excel_generator.py
from tkinter import messagebox
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import json
import os
import calendar
from datetime import datetime

class ExcelGenerator:
    def __init__(self, template_path):
        self.template_path = template_path
        self.wb = openpyxl.load_workbook(template_path)
        self.ws = self.wb.active

    def gregorian_to_reiwa(self, year, month):
        """
        西暦を令和に変換する関数。
        令和元年は2019年。
        """
        if year < 2019:
            raise ValueError("令和は2019年以降の年です。")
        reiwa_year = year - 2018
        return f"令和{reiwa_year}年{month}月"
    
    def is_holiday(self, year, month, day):
        # 年と月を文字列に変換
        year_str = str(year)
        month_str = str(month)
        
        # 工場ごとの休日データを取得
        if not hasattr(self, 'holidays_data'):
            raise AttributeError("holidays_dataがロードされていません。generate_excelを先に実行してください。")
        
        # 指定された年月の休日リストを取得
        if year_str in self.holidays_data and month_str in self.holidays_data[year_str]:
            return day in self.holidays_data[year_str][month_str]
        return False

    def get_top_left_cell(self, cell):
        for merged_range in self.ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                return self.ws.cell(row=merged_range.min_row, column=merged_range.min_col)
        return cell

    def apply_holidays(self, holidays, start_row, year, month):
        for day in holidays:
            try:
                # 日にちのセルを特定 (例: D8～AH8が日にち)
                col = 4 + day - 1  # D列は4
                col_letter = get_column_letter(col)
    
                # 曜日のセルはD9～AH9（開始行+1）
                day_cell_coord = f"{col_letter}{start_row + 1}"
                weekday_cell_coord = f"{col_letter}{start_row + 2}"
    
                day_cell = self.ws[day_cell_coord]
                weekday_cell = self.ws[weekday_cell_coord]
    
                # 結合セルの場合、左上のセルを取得
                top_left_day_cell = self.get_top_left_cell(day_cell)
                top_left_weekday_cell = self.get_top_left_cell(weekday_cell)
    
                # 日付の曜日を取得
                try:
                    date = datetime(year, month, day)
                    weekday_en = date.strftime('%a')  # 'Mon', 'Tue', etc.
                    # 英語の曜日を日本語にマッピング
                    weekday_jp = {
                        'Mon': '月',
                        'Tue': '火',
                        'Wed': '水',
                        'Thu': '木',
                        'Fri': '金',
                        'Sat': '土',
                        'Sun': '日'
                    }.get(weekday_en, '')
                except ValueError:
                    # 不正な日付（例: 2月30日など）の場合はスキップ
                    continue
    
                # 赤文字に設定
                red_font = Font(name='游ゴシック', size=14, bold=True, color='FF0000')
                if top_left_day_cell.value == day:
                    top_left_day_cell.font = red_font
                if top_left_weekday_cell.value.startswith(weekday_jp):
                    top_left_weekday_cell.font = red_font
            except Exception as e:
                print(f"Error applying holiday for {year}/{month}/{day}: {e}")

    def set_parameters(self, creation_date, title, month, start_row):
        """最初の月にのみ作成日とタイトルを設定"""
        # AE{start_row} に作成日を設定
        cell_coord = f'AE{start_row}'
        cell = self.ws[cell_coord]
        top_left_cell = self.get_top_left_cell(cell)
        top_left_cell.value = creation_date

        # A{start_row} にタイトルを設定
        cell_coord = f'A{start_row}'
        cell = self.ws[cell_coord]
        top_left_cell = self.get_top_left_cell(cell)
        top_left_cell.value = f"{title} 製造工程計画"

        # D{start_row} に月を設定
        cell_coord = f'D{start_row}'
        cell = self.ws[cell_coord]
        top_left_cell = self.get_top_left_cell(cell)
        top_left_cell.value = self.gregorian_to_reiwa(month.year, month.month)

    def set_month_parameters(self, month, start_row):
        """二つ目以降の月は月のみ設定"""
        # D{start_row} に月を設定
        cell_coord = f'D{start_row}'
        cell = self.ws[cell_coord]
        top_left_cell = self.get_top_left_cell(cell)
        top_left_cell.value = self.gregorian_to_reiwa(month.year, month.month)

    def save_and_open(self, output_path):
        self.wb.save(output_path)
        os.startfile(output_path)

    def generate_excel(self, title, start_date, end_date, factory, output_path):

        try:
            # 休日データのロード
            holidays_file = os.path.join(os.path.dirname(__file__), 'holidays', f'{factory}.json')
            if not os.path.exists(holidays_file):
                raise FileNotFoundError(f"工場 '{factory}' の休日JSONファイルが見つかりません。")
            
            with open(holidays_file, 'r', encoding='utf-8') as f:
                self.holidays_data = json.load(f)

            # 作成日を入力
            now = datetime.now()
            current_date_str = now.strftime('%Y/%m/%d')
            self.ws['AE1'].value = current_date_str

            # タイトルを入力
            self.ws['A3'].value = f"{title} 製造工程計画"

            # テーブルの開始位置
            first_table_start_row = 8
            rows_per_table = 13  # 各月のテーブルが13行を使用

            # 現在の行を初期化
            current_row = first_table_start_row

            # 開始年月から終了年月までループ
            year = start_date.year
            month = start_date.month
            end_year = end_date.year
            end_month = end_date.month

            while (year < end_year) or (year == end_year and month <= end_month):
                # 月のタイトルを入力
                reiwa_month = self.gregorian_to_reiwa(year, month)
                self.ws.cell(row=current_row, column=4, value=reiwa_month)

                # 月の日数を取得
                last_day = calendar.monthrange(year, month)[1]

                # 日にちを入力
                for day in range(1, last_day + 1):
                    col = 4 + day - 1  # D列は4
                    self.ws.cell(row=current_row + 1, column=col, value=day)
                
                # 曜日を入力
                for day in range(1, last_day + 1):
                    date = datetime(year, month, day)
                    weekday_en = date.strftime('%a')  # 'Mon', 'Tue', etc.
                    # 英語の曜日を日本語にマッピング
                    weekday_jp = {
                        'Mon': '月',
                        'Tue': '火',
                        'Wed': '水',
                        'Thu': '木',
                        'Fri': '金',
                        'Sat': '土',
                        'Sun': '日'
                    }.get(weekday_en, '')
                    col = 4 + day - 1
                    self.ws.cell(row=current_row + 2, column=col, value=weekday_jp)

                # 休日のセルに色を付ける（日にちと曜日のみ）
                for day in range(1, last_day + 1):
                    if self.is_holiday(year, month, day):
                        col = 4 + day - 1
                        # 日にちのセル（current_row +1）と曜日のセル（current_row +2）に色を付ける
                        for row_offset in [1, 2]:  # 1: 日にち, 2: 曜日
                            cell_row = current_row + row_offset
                            cell = self.ws.cell(row=cell_row, column=col)
                            cell.font = Font(name='游ゴシック', size=14, bold=True, color='FF0000')

                # 次の月のテーブル開始行を更新
                current_row += rows_per_table

                # 月をインクリメント
                if month == 12:
                    year += 1
                    month = 1
                else:
                    month += 1

        except Exception as e:
            messagebox.showerror("エラー", f"製造工程表の生成中にエラーが発生しました。\n{e}")

        # 保存して開く
        self.save_and_open(output_path)
