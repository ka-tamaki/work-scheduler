# utils/path_helper.py
import os
from datetime import datetime

def get_template_path():
    return os.path.join(os.path.dirname(__file__), '..', 'templates', 'template.xlsx')

def get_output_path(title, start_date, end_date):
    filename = f"{title}_{start_date.strftime('%Y%m')}_{end_date.strftime('%Y%m')}.xlsx"
    output_dir = os.path.join(os.path.dirname(__file__), '..', 'output')
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    return os.path.join(output_dir, filename)
