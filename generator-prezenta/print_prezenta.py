import os
from datetime import date
from generate_libere import Libere
from generate_prezenta import Excel

if __name__ == '__main__':
    settings = Libere.loadJson("settings.json")
    path_input = settings["input_folder"]
    month = Excel.value_to_key(date.today().month)
    file = "Prezenta-{}.xlsx".format(month)

    os.startfile(path_input + file, 'print') # print only first sheet
