# utils/path_helper.py
import os
import sys
from datetime import datetime

def resource_path(relative_path):
    """PyInstaller でパッケージされた場合と通常実行の場合でパスを取得"""
    try:
        # PyInstaller でパッケージされた場合
        base_path = sys._MEIPASS
    except Exception:
        # 通常の実行の場合
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_template_path():
    return resource_path(os.path.join('templates', 'template.xlsx'))

def get_output_path(title, start_date, end_date):
    # 出力パスの生成ロジック
    output_dir = resource_path('output')
    os.makedirs(output_dir, exist_ok=True)
    filename = f"{title}_{datetime.now().strftime('%Y/%m/%d_%H%M%S')}.xlsx"
    return os.path.join(output_dir, filename)
