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
exe化

pyinstaller --onefile --windowed `
--exclude-module holidays_initializer `
--add-data "templates/template.xlsx;templates" `
--add-data "data/holidays;data/holidays" `
--hidden-import utils.path_helper `
main.py
'''