# main.py
import tkinter as tk
from ui.main_ui import MainUI

def main():
    root = tk.Tk()
    app = MainUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()


'''
pyinstaller --onefile --windowed --add-data "data/holidays;data/holidays" --add-data "templates/template.xlsx;templates" ui/main_ui.py
'''