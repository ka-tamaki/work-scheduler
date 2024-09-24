# data/excel_generator.py
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import json
import os
from datetime import datetime
from utils.path_helper import get_template_path

class ExcelGenerator:
    def __init__(self, template_path):
        self.template_path = template_path
        self.wb = openpyxl.load_workbook(template_path)
        self.ws = self.wb.active

    def apply_holidays(self, holidays, start_row):
        for holiday in holidays:
            date_str = holiday['date']
            date = datetime.strptime(date_str, '%Y-%m-%d')
            day = date.day
            weekday = date.strftime('%a')  # 'Mon', 'Tue', etc.

            # 日にちのセルを特定 (例: D21～AH21が日にち)
            col = 3 + day  # D列は4列目なのでインデックス調整
            col_letter = get_column_letter(col + 1)  # openpyxlは1始まり

            # 曜日のセルはD22～AH22
            day_cell = f"{col_letter}{start_row}"
            weekday_cell = f"{col_letter}{start_row + 1}"

            # 赤文字に設定
            red_font = Font(color="FF0000")
            if self.ws[day_cell].value == day:
                self.ws[day_cell].font = red_font
            if self.ws[weekday_cell].value.startswith(weekday):
                self.ws[weekday_cell].font = red_font

    def set_parameters(self, creation_date, title, month, start_row):
        # 最初の月のみ作成日とタイトルを設定
        self.ws[f'AE{start_row}'] = creation_date  # 作成日
        self.ws[f'A{start_row}'] = f"{title}　製造工程計画"  # タイトル
        self.ws[f'D{start_row}'] = f"令和{month.year - 2018}年{month.month}月"  # 月

    def set_month_parameters(self, month, start_row):
        # 二つ目以降の月は月のみ設定
        self.ws[f'D{start_row}'] = f"令和{month.year - 2018}年{month.month}月"  # 月

    def save_and_open(self, output_path):
        self.wb.save(output_path)
        os.startfile(output_path)  # Windowsの場合

    def generate_excel(self, title, start_date, end_date, factory, output_path):
        current_date = start_date
        start_row = 21  # 初期の開始行（テンプレートに合わせて調整）
        is_first_month = True  # 最初の月かどうかのフラグ

        while current_date <= end_date:
            try:
                if is_first_month:
                    # 最初の月は作成日とタイトルを設定
                    creation_date = datetime.today().strftime('%Y/%m/%d')
                    self.set_parameters(creation_date, title, current_date, start_row)
                    is_first_month = False
                else:
                    # 二つ目以降の月は月のみ設定
                    self.set_month_parameters(current_date, start_row)

                # 休日情報の読み込み
                holidays_file = os.path.join('holidays', f'{factory}.json')
                with open(holidays_file, 'r', encoding='utf-8') as f:
                    holidays = json.load(f)

                # 期間の月に該当する休日をフィルタリング
                month_holidays = [
                    h for h in holidays 
                    if datetime.strptime(h['date'], '%Y-%m-%d').year == current_date.year 
                    and datetime.strptime(h['date'], '%Y-%m-%d').month == current_date.month
                ]

                # 休日を適用
                self.apply_holidays(month_holidays, start_row)

                # 次の月へ
                if current_date.month == 12:
                    year = current_date.year + 1
                    month = 1
                else:
                    year = current_date.year
                    month = current_date.month + 1
                current_date = datetime(year, month, 1)

                # 次の月の開始行を13行下げる
                start_row += 13
            except Exception as e:
                raise Exception(f"月 {current_date.strftime('%Y/%m')} の工程表生成中にエラーが発生しました: {e}")

        # 保存して開く
        self.save_and_open(output_path)
