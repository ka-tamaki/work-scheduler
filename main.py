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
pyinstaller --onefile --add-data "templates/template.xlsx;templates" \
    --add-data "holidays/*.json;holidays" \
    main.py
'''