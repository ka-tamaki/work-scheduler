# data/excel_generator.py
from tkinter import messagebox
import json
import os
import calendar
import xlwings as xw
from xlwings.constants import HAlign, VAlign, BorderWeight, LineStyle
from datetime import datetime

class ExcelGenerator:
    def __init__(self):
        self.wb = xw.Book()
        self.ws = self.wb.sheets.active
        self.ws.name = "製造工程表"

    def center_value_in_range(self, value, cell_range, name='游ゴシック', size=14, bold=True):
        """
        選択範囲の左上セルに値を入力し、選択範囲全体を縦横中央揃えにする関数。
        """
        # セル範囲を取得
        self.select_range = cell_range
        
        # 左上セルに値を入力
        self.select_range[0, 0].value = value

        # フォント設定
        self.select_range.api.Font.Name = name
        self.select_range.api.Font.Size = size
        self.select_range.api.Font.Bold = bold

        # 範囲全体の縦横中央揃え
        self.select_range.api.HorizontalAlignment = HAlign.xlHAlignCenterAcrossSelection
        self.select_range.api.VerticalAlignment = VAlign.xlVAlignCenter

    def format_range(self, cell_value, range_name, name='游ゴシック', size=14, bold=True):
        """
        セル範囲のフォーマットを設定する関数
        """
        # セル範囲を取得
        self.select_range = range_name

        # 水平方向と垂直方向を中央揃え
        self.select_range.api.HorizontalAlignment = HAlign.xlHAlignCenter
        self.select_range.api.VerticalAlignment = VAlign.xlVAlignCenter

        # フォント設定
        self.select_range.value = cell_value
        self.select_range.api.Font.Name = name
        self.select_range.api.Font.Size = size
        self.select_range.api.Font.Bold = bold

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
        """
        各工場の休日判定。
        休日セルの色付けに使用。
        """
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

    def generate_excel(self, title, start_date, end_date, factory, output_path):
        """
        excelの生成。
        """
        try:
            # 休日データのロード
            holidays_file = os.path.join(os.path.dirname(__file__), 'holidays', f'{factory}.json')
            if not os.path.exists(holidays_file):
                raise FileNotFoundError(f"工場 '{factory}' の休日JSONファイルが見つかりません。")
            
            with open(holidays_file, 'r', encoding='utf-8') as f:
                self.holidays_data = json.load(f)

            # スクリーンアップデートをオフにする
            xw.apps.active.screen_updating = False
            
            # セル幅、セル高の設定
            self.ws.range('A:A').column_width = 25
            self.ws.range('B:B').column_width = 30
            self.ws.range('C:C').column_width = 5.5
            self.ws.range('D:AH').column_width = 5
            self.ws.range('AI:AJ').column_width = 5.5
            self.ws.range('1:1').row_height = 32.5
            self.ws.range('3:3').row_height = 58.5
            self.ws.range('4:4').row_height = 32.5

            # 作成日を入力
            now = datetime.now()
            current_date_str = now.strftime('%Y/%m/%d')
            self.date_range = self.ws.range('AE1:AJ1')
            self.center_value_in_range(current_date_str, self.date_range, size=20)

            # タイトルを入力
            self.title_range = self.ws.range('A3:AJ3')
            self.center_value_in_range(f"{title} 製造工程計画", self.title_range, size=36)

            # 定型コメントを入力
            self.format_range("※以下の日程は生コン打設となります。", self.ws.range('A4'), size=20)
            self.ws.range('A4').api.HorizontalAlignment = HAlign.xlHAlignLeft
            self.format_range("ベルテクス株式会社", self.ws.range('AJ4'), size=20)
            self.ws.range('AJ4').api.HorizontalAlignment = HAlign.xlHAlignRight

            # テーブルの開始位置
            first_table_start_row = 7
            contents_rows = 8  # 項目の数
            rows_per_table = contents_rows + 5

            # 現在の行を初期化
            current_row = first_table_start_row

            # 開始年月から終了年月までループ
            year = start_date.year
            month = start_date.month
            end_year = end_date.year
            end_month = end_date.month
            
            # 各月の表ループ
            while (year < end_year) or (year == end_year and month <= end_month):
                # 各表のヘッダー部分作成
                self.items_range = self.ws.range(f'A{current_row}:A{current_row + 2}')
                self.center_value_in_range("項目", self.items_range)
                self.standards_range = self.ws.range(f'B{current_row}:B{current_row + 2}')
                self.center_value_in_range("規格", self.standards_range)
                self.orders_range = self.ws.range(f'C{current_row}:C{current_row + 2}')
                self.center_value_in_range("受注", self.orders_range)
                self.sums_range = self.ws.range(f'AI{current_row}:AJ{current_row + 1}')
                self.center_value_in_range("合計", self.sums_range)
                self.volumes_range = self.ws.range(f'AI{current_row + 2}:AI{current_row + 2}')
                self.center_value_in_range("数量", self.volumes_range)
                self.remains_range = self.ws.range(f'AJ{current_row + 2}:AJ{current_row + 2}')
                self.center_value_in_range("残り", self.remains_range)

                #セル高の設定
                self.ws.range(f'{current_row}:{current_row}').row_height = 32.5
                self.ws.range(f'{current_row + 1}:{current_row + 2}').row_height = 23
                self.ws.range(f'{current_row + 3}:{current_row + contents_rows + 3}').row_height = 36.8

                # 入力欄の設定
                self.items_col = self.ws.range(f'A{current_row + 3}:A{current_row + contents_rows + 3}') # 項目
                self.format_range("", self.items_col)
                self.standards_col = self.ws.range(f'B{current_row + 3}:B{current_row + contents_rows + 3}') # 規格
                self.format_range("", self.standards_col)
                self.orders_col = self.ws.range(f'C{current_row + 3}:C{current_row + contents_rows + 3}') # 受注数
                self.format_range("0", self.orders_col)
                self.num_col = self.ws.range(f'D{current_row + 3}:AH{current_row + contents_rows + 3}') # 製造数
                self.format_range("", self.num_col)
                for row_offset in range(3, contents_rows + 4): # 月合計数
                    cell_row = current_row + row_offset
                    self.format_range(f'=SUM(D{str(cell_row)}:AH{str(cell_row)})', self.ws.range(f'AI{cell_row}'))
                for row_offset in range(3, contents_rows + 4): # 残りの計算
                    cell_row = current_row + row_offset
                    if current_row == first_table_start_row:
                        self.format_range(f'=C{str(cell_row)}-AI{str(cell_row)}', self.ws.range(f'AJ{cell_row}'))
                    else:
                        self.format_range(f'=AJ{str(cell_row - rows_per_table)}-AI{str(cell_row)}', self.ws.range(f'AJ{cell_row}'))

                # 月のタイトルを入力
                self.reiwa_month = self.gregorian_to_reiwa(year, month)
                self.month_range = self.ws.range(f'D{current_row}:AH{current_row}')
                self.center_value_in_range(self.reiwa_month, self.month_range, size=20)

                # 月の日数を取得
                last_day = calendar.monthrange(year, month)[1]

                # 日にちを入力
                for day in range(1, last_day + 1):
                    col = 4 + day - 1  # D列は4
                    cell = self.ws.cells(current_row + 1, col)
                    self.format_range(day, cell)
                
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
                    cell = self.ws.cells(current_row + 2, col)
                    self.format_range(weekday_jp, cell)

                # 罫線の設定
                for i in range(11, 13): # 最後に細い点線を設定
                    border = self.ws.range(f'A{current_row}:AJ{current_row + contents_rows + 3}').api.Borders(i)
                    border.LineStyle = LineStyle.xlDash
                    border.Weight = BorderWeight.xlThin
                for i in range(7, 11):  # 外側の太線
                    border_left_top = self.ws.range(f'A{current_row}:AH{current_row + 2}').api.Borders(i)
                    border_left_top.LineStyle = LineStyle.xlContinuous
                    border_left_top.Weight = BorderWeight.xlMedium

                    border_right_top = self.ws.range(f'AI{current_row}:AJ{current_row + 2}').api.Borders(i)
                    border_right_top.LineStyle = LineStyle.xlContinuous
                    border_right_top.Weight = BorderWeight.xlMedium

                    border_left_buttom = self.ws.range(f'A{current_row + 3}:AH{current_row + contents_rows + 3}').api.Borders(i)
                    border_left_buttom.LineStyle = LineStyle.xlContinuous
                    border_left_buttom.Weight = BorderWeight.xlMedium

                    border_right_buttom = self.ws.range(f'AI{current_row + 3}:AJ{current_row + contents_rows + 3}').api.Borders(i)
                    border_right_buttom.LineStyle = LineStyle.xlContinuous
                    border_right_buttom.Weight = BorderWeight.xlMedium
                for col in ['A', 'B', 'C', 'AI']: # 列ごとの細い実線
                    border = self.ws.range(f'{col}{current_row}:{col}{current_row + contents_rows + 3}').api.Borders(10)
                    border.LineStyle = LineStyle.xlContinuous
                    border.Weight = BorderWeight.xlThin


                # 休日のセルに色を付ける（楯列すべて）
                for day in range(1, last_day + 1):
                    if self.is_holiday(year, month, day):
                        col = 4 + day - 1
                        for row_offset in range(1, rows_per_table - 1):
                            cell_row = current_row + row_offset
                            cell = self.ws.cells(cell_row, col)
                            cell.api.Font.Name = '游ゴシック'
                            cell.api.Font.Size = 14
                            cell.api.Font.Bold = True
                            cell.api.Font.Color = xw.utils.rgb_to_int((255, 0, 0))
                            cell.api.Interior.Color = xw.utils.rgb_to_int((255, 199, 206))

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
        finally:
            # スクリーンアップデートをオンに戻す
            xw.apps.active.screen_updating = True

