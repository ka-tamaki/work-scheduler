# main.py
import tkinter as tk
from ui.main_ui import MainUI
from update_checker import perform_update_check

def main():
    perform_update_check()

    root = tk.Tk()
    app = MainUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()

