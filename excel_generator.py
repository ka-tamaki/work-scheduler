# excel_generator.py

import os
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
import calendar
from tkinter import messagebox
import config

def gregorian_to_reiwa(year, month):
    """
    西暦を令和に変換する関数。
    令和元年は2019年。
    """
    if year < 2019:
        raise ValueError("令和は2019年以降の年です。")
    reiwa_year = year - 2018
    return f"令和{reiwa_year}年{month}月"

def generate_schedule(title, start_year, start_month, end_year, end_month):
    try:
        template_path = config.TEMPLATE_PATH
        if not os.path.exists(template_path):
            messagebox.showerror("エラー", f"テンプレートファイル '{template_path}' が見つかりません。")
            return
        
        # 出力ディレクトリの作成
        if not os.path.exists(config.SAVE_DIR):
            os.makedirs(config.SAVE_DIR)
        
        wb = load_workbook(template_path)
        ws = wb.active  # アクティブなシートを選択
        
        # 作成日を入力
        current_date_str = datetime.now().strftime('%Y/%m/%d')
        ws['A1'].value = f"作成日: {current_date_str}"
        ws['A1'].alignment = Alignment(horizontal='left', vertical='center')
        ws['A1'].font = Font(size=12)
        
        # タイトルを入力
        ws['A2'].value = f"{title} 製造工程計画"
        ws['A2'].alignment = Alignment(horizontal='left', vertical='center')
        ws['A2'].font = Font(bold=True, size=14)
        
        # テーブルの開始位置
        first_table_start_row = 8
        rows_per_table = 13  # 各月のテーブルが13行を使用
        
        # 現在の行を初期化
        current_row = first_table_start_row
        
        # 開始年月から終了年月までループ
        year = start_year
        month = start_month
        
        while (year < end_year) or (year == end_year and month <= end_month):
            # 令和形式に変換
            reiwa_month = gregorian_to_reiwa(year, month)
            
            # 月をD行に入力
            ws.cell(row=current_row, column=4, value=reiwa_month)
            ws.cell(row=current_row, column=4).alignment = Alignment(horizontal='center', vertical='center')
            ws.cell(row=current_row, column=4).font = Font(bold=True, size=12)
            
            # 月の日数を取得
            last_day = calendar.monthrange(year, month)[1]
            
            # 日にちをD行+1に入力
            for day in range(1, last_day + 1):
                col = 4 + day - 1  # D列は4
                ws.cell(row=current_row + 1, column=col, value=day)
                ws.cell(row=current_row + 1, column=col).alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(row=current_row + 1, column=col).font = Font(bold=True)
            
            # 曜日をD行+2に入力
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
                ws.cell(row=current_row + 2, column=col, value=weekday_jp)
                ws.cell(row=current_row + 2, column=col).alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(row=current_row + 2, column=col).font = Font(bold=True)
            
            # フォントや色の設定（必要に応じて）
            # 例えば、ヘッダー行に背景色を設定する場合
            header_fill = PatternFill(start_color='FFD700', end_color='FFD700', fill_type='solid')  # ゴールド
            for col in range(4, 4 + last_day):
                ws.cell(row=current_row, column=col).fill = header_fill
                ws.cell(row=current_row + 1, column=col).fill = header_fill
                ws.cell(row=current_row + 2, column=col).fill = header_fill
            
            # 空白行19,20は自動で空白のまま
            
            # 次の月のテーブル開始行を更新
            current_row += rows_per_table
            
            # 月をインクリメント
            if month == 12:
                year += 1
                month = 1
            else:
                month += 1
        
        # 列幅の調整（必要に応じて）
        for col in range(4, 4 + 31):  # D列からAH列まで
            column_letter = get_column_letter(col)
            ws.column_dimensions[column_letter].width = 3  # 幅を3に設定（必要に応じて調整）
        
        # 保存ファイル名の作成
        save_filename = f"{title}_製造工程計画_{start_year}_{start_month:02}_{end_year}_{end_month:02}.xlsx"
        save_path = os.path.join(config.SAVE_DIR, save_filename)
        wb.save(save_path)
        
        messagebox.showinfo("成功", f"製造工程表を '{save_path}' として保存しました。")
    except Exception as e:
        messagebox.showerror("エラー", f"製造工程表の生成中にエラーが発生しました。\n{e}")
