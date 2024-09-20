# main.py

import tkinter as tk
from ui.main_ui import MainUI
from utils.holidays_initializer import initialize_holidays

def main():
    # holidays.jsonを初期化または更新
    initialize_holidays()

    # Tkinterアプリケーションを開始
    root = tk.Tk()
    root.geometry("500x400")
    app = MainUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
