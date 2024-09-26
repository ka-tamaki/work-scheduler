pyinstaller --onefile --windowed `
--icon "assets\icon.ico" `
--exclude-module holidays_initializer `
--add-data "templates/template.xlsx;templates" `
--add-data "data/holidays;data/holidays" `
--hidden-import utils.path_helper `
main.py