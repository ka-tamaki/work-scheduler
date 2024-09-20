# config.py

import os

# 休日情報のファイルパス
HOLIDAYS_FILE = os.path.join(os.path.dirname(__file__), 'holidays.json')

# Excelテンプレートファイルのパス
EXCEL_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'templates', 'template.xlsx')

# Excel出力ディレクトリのパス
EXCEL_OUTPUT_DIR = os.path.join(os.path.dirname(__file__), 'output')