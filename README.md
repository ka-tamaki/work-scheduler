pyinstaller --onefile --windowed `
--icon "assets\icon.ico" `
--exclude-module holidays_initializer `
--add-data "data/holidays;data/holidays" `
--hidden-import utils.path_helper `
main.py