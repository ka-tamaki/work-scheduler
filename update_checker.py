# ui/update_checker.py

import requests
from packaging import version
import tkinter as tk
from tkinter import messagebox
import webbrowser
import sys

# アプリケーションの現在のバージョン
CURRENT_VERSION = "1.3.0"

# GitHubリポジトリ情報
GITHUB_USER = "ka-tamaki"
GITHUB_REPO = "work-scheduler"

def get_latest_release():
    url = f"https://api.github.com/repos/{GITHUB_USER}/{GITHUB_REPO}/releases/latest"
    try:
        response = requests.get(url)
        response.raise_for_status()
        return response.json()
    except requests.RequestException as e:
        print(f"アップデートチェック中にエラーが発生しました: {e}")
        return None

def check_for_update():
    latest_release = get_latest_release()
    if latest_release:
        latest_version = latest_release['tag_name'].lstrip('v')  # タグにvが含まれている場合
        if version.parse(latest_version) > version.parse(CURRENT_VERSION):
            return latest_release['html_url'], latest_version
    return None, None

def notify_user(download_url, latest_version):
    root = tk.Tk()
    root.withdraw()  # メインウィンドウを非表示にする
    response = messagebox.askyesno(
        "アップデートのお知らせ",
        f"新しいバージョン {latest_version} が利用可能です。\n\nアップデートしますか？"
    )
    if response:
        webbrowser.open(download_url)
    root.destroy()

def perform_update_check():
    download_url, latest_version = check_for_update()
    if download_url:
        notify_user(download_url, latest_version)
    
    print("perform_update_check実行")
