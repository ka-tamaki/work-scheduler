# utils/path_helper.py
import os
import sys

def resource_path(relative_path):
    """PyInstaller でパッケージされた場合と通常実行の場合でパスを取得"""
    try:
        # PyInstaller でパッケージされた場合
        base_path = sys._MEIPASS
    except Exception:
        # 通常の実行の場合
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)
