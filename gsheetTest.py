from gsheets import Sheets

path="G:\\Mon Drive\\2020 equipment maintenance sw Budget xlsx.gsheet"
sheets = Sheets.from_files()
s=sheets.get(path)