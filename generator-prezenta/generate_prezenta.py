from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, Alignment
from openpyxl.styles.borders import Border, Side
from generate_libere import Libere
from datetime import date
import calendar, math
_month = {"ianuarie": 1, "februarie": 2, "martie": 3, "aprilie": 4, "mai": 5, "iunie": 6, "iulie": 7, "august": 8, "septembrie": 9, "octombrie": 10, "noiembrie": 11, "decembrie": 12}

def get_work_days(month, year):
    max = calendar.monthrange(year, month)[1]
    lista = []
    for day in range(1, max+1):
        weekday = calendar.weekday(year, month, day)
        if weekday < calendar.SATURDAY:
            if day < 10:
                day = "0" + str(day)
            else:
                day = str(day)
            lista.append(day)
    return lista

class Excel:
    def __init__(self, output_folder=""):
        self.wb = Workbook()
        sh = self.wb.active
        sh.title = "Prezenta"
        self.sh = self.wb[sh.title]
        self.month = Excel.value_to_key(date.today().month)
        self.year = date.today().year
        self.default = Excel.default() #dictionary
        self.output_folder = output_folder

    @staticmethod
    def transformSelection(selection): # tuple (row, column) -> col+row
        if isinstance(selection, str):
            return selection # do nothing

        row = selection[0]
        column = selection[1]
        if column > 16384 or column < 1 or row > 1048576 or row < 1:
            raise Exception("Value exceed excel board.")
        letter = Excel.get_column_letter(column)

        return "".join(letter) + str(row)

    @staticmethod
    def get_column_letter(column):
        if isinstance(column, int): # 1
            col = [] # A, B, AA, AB, BA, BC
            alfa = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" # 26
            base = len(alfa)
            n = int(math.log(column, base)) # 2 1 0
            while column > 0:
                product  = 1 ; n1 = n
                while n1 > 0: # base ** n
                    product *= base
                    n1 -= 1
                value = int(column/product)
                col.append(alfa[value-1])
                column = column % product
                n -= 1
            return "".join(col)
        elif isinstance(column, str): # A12
            result = "".join([i for i in column if i is not i.isdigit()])
            return result #A

    def add_header_image(self, selection="B1"):
        if isinstance(selection, tuple) or isinstance(selection, list):
            select = Excel.transformSelection(selection)
        elif not(isinstance(selection, str)):
            raise Exception("Not String or tuple / list")
        else:
            select = selection
        img = Image("src/allora.PNG")
        self.sh.add_image(img, select)

    def add_value(self, selection, value=None):
        if isinstance(selection, str):
            self.sh[selection] = value
            self.sh[selection].font = Font(name=self.default["font_name"], size=self.default["font_size"], bold=self.default["font_bold"])
        elif isinstance(selection, tuple) or isinstance(selection, list):
            row = selection[0]
            column = selection[1]
            self.sh.cell(row=row, column=column, value=value)
            self.sh.cell(row=row, column=column).font = Font(name=self.default["font_name"], size=self.default["font_size"], bold=self.default["font_bold"])
        else:
            raise Exception("valueType")

    @staticmethod
    def default():
        return {"font_name": "Times New Roman",
                "font_size": "12",
                "font_name_table": "Calibri",
                "font_size_table": "11",
                "font_bold": False,
                "border": {
                    "left": ["thin", 'FF000000'],
                    "right": ["thin", 'FF000000'],
                    "top": ["thin", 'FF000000'],
                    "bottom": ["thin", 'FF000000']}
                }

    def font_style(self, selection, name=None, bold=False, size=None, wrap=False):####################se interfereaza font cu alignement
        if isinstance(selection, str):
            select = self.sh[selection]
        elif isinstance(selection, tuple) or isinstance(selection, list):
            select = self.sh.cell(row=selection[0], column=selection[1])
        else:
            select = None
            raise Exception("No selecting - font_style")

        select.font = Font(name=self.default["font_name"] if name is None else name, size=size if size is not None else self.default["font_size"], bold=True if bold else self.default["font_bold"])

    def font_align(self, selection, wrap = False):
        if isinstance(selection, str):
            select = self.sh[selection]
        elif isinstance(selection, tuple) or isinstance(selection, list):
            select = self.sh.cell(row=selection[0], column=selection[1])
        else:
            select = None
            raise Exception("No selecting - font_align")

        select.alignment = Alignment(wrap_text=True if wrap else False)

    def border_style(self, selection, listBorder=[]): # ma mai gandesc la border default
        if isinstance(selection, str): # [[thin, ffff], [thin, ffff], [], []] n-e-s-w
            select = self.sh[selection]
        elif isinstance(selection, tuple) or isinstance(selection, list):
            select = self.sh.cell(row=selection[0], column=selection[1])
        else:
            select = None
            raise Exception("No selecting - border_style")
        sideLeft = Side(style=self.default["border"]["left"][0], color=self.default["border"]["left"][1]) if len(listBorder) < 4 else Side(style=listBorder[3][0], color=listBorder[3][1])
        sideRight = Side(style=self.default["border"]["right"][0], color=self.default["border"]["right"][1]) if len(listBorder) < 4 else Side(style=listBorder[1][0], color=listBorder[1][1])
        sideTop = Side(style=self.default["border"]["top"][0], color=self.default["border"]["top"][1]) if len(listBorder) < 4 else Side(style=listBorder[0][0], color=listBorder[0][1])
        sideBottom = Side(style=self.default["border"]["bottom"][0], color=self.default["border"]["bottom"][1]) if len(listBorder) < 4 else Side(style=listBorder[2][0], color=listBorder[2][1])
        select.border = Border(left=sideLeft, right=sideRight, top=sideTop, bottom=sideBottom)

    def modify_column(self, column, value):
        if isinstance(column, str):
            self.sh.column_dimensions[column].width = value
        elif isinstance(column, int):
            self.sh.column_dimensions[get_column_letter(column)].width = value

    def modify_row(self, row, value):
        self.sh.row_dimensions[row].height = value

    @staticmethod
    def value_to_key(value): # Aprilie -> 4
        listKeys = list(_month.keys())
        listValues = list(_month.values())
        return listKeys[listValues.index(value)]

    def column_auto_size(self, selection):
        omisiuni = 100
        if isinstance(selection, int) or isinstance(selection, str):
            select = Excel.get_column_letter(selection)
        elif isinstance(selection, tuple) or isinstance(selection, list):
            select = Excel.get_column_letter(selection[1]) # A
        #self.sh.column_dimensions[select].width = 20

        blank = 0 ; row = 1 ; max_len = 0 ; is_wrap = False
        max_font_size = 0
        while blank <= omisiuni:
            thisCell = self.sh[select + str(row)]
            if  thisCell.value is None: #cell empty
                blank += 1
                row += 1
                continue
            else:
                blank = 0
                word_len = len(thisCell.value) # longest text
                if word_len > max_len:
                    max_len = word_len
                if thisCell.alignment.wrapText:
                    is_wrap = True
                if thisCell.font.sz > max_font_size: # biggest font size
                    max_font_size = thisCell.font.sz # 11.0
                ################################ verific daca e una merged
                row += 1
        if is_wrap:
            self.sh.column_dimensions[select].width = max_len * (max_font_size / 10) / 2 -2
            print(select, max_len * (max_font_size / 10) / 2) -2
        else:
            print(select, max_len * (max_font_size / 10)) -2
            self.sh.column_dimensions[select].width = max_len * (max_font_size / 10) -2

    def column_size(self, selection, width):
        pass

    def save(self):
        self.wb.save("{}Prezenta-{}.xlsx".format(self.output_folder, self.month))

if __name__ == '__main__':
    introductionData = ["SC ALLORA VISION TECH SRL",
    "Reg. Com. J40/11468/2011",
    "CUI:RO29146323",
    "Sediul: Calea Rahovei nr.266-268, corp 60, et.2, camera 30A",
    "incinta Electromagnetica\nBusiness\nPark.",
    "Contul: RO49INGB0000999903977417",
    "Persoana de contact: Nicoleta Baciu",
    "Telefon: 0723290110/0721153839"
    ]

    ORA_INCEPUT = "09:00"
    ORA_SFARSIT = "17:30"
    PAUZA = "12:30-13:00"
    year = date.today().year
    month = str(date.today().month) if date.today().month > 9 else "0" + str(date.today().month)
    days = get_work_days(date.today().month, date.today().year)

    settings = Libere.loadJson("settings.json")
    persons = settings["persons"]

    wb = Excel(settings["output_folder"])

    #adaugat imagine
    wb.add_header_image()

    # adaugat header text
    for i in range(4, 12):
        wb.add_value([i, 2], introductionData[i-4])
        if i == 8:
            wb.font_align([i, 2], wrap=True)
    wb.add_value("D12", "CONDICA PREZENTA")

    wb.font_style("B10", bold=True, size="10")
    wb.font_style("B11", bold=True)
    wb.font_style("D12", bold=True, size="14")

    #completat tabel
    curentRow = 15
    font_name_table = wb.default["font_name_table"]
    font_size_table = wb.default["font_size_table"]
    for section in range(len(days)):
        for row in range(1, 2+len(persons) +1): #6
            for col in range(1, 8+1):
                selection = [curentRow, col]
                wb.font_style(selection, name=font_name_table, size=font_size_table)
                if row == 1:
                    if col == 2:
                        wb.add_value(selection, "ZIUA:{}".format(days[section]))
                        wb.font_style(selection, bold=True)
                    elif col == 4:
                        wb.add_value(selection, "LUNA:{}".format(month))
                        wb.font_style(selection, bold=True)
                    elif col == 6:
                        wb.add_value(selection, "ANUL:{}".format(year))
                        wb.font_style(selection, bold=True)
                elif row == 2:
                    if col == 1:
                        wb.add_value(selection, "Nr.")
                    elif col == 2:
                        wb.add_value(selection, "Nume si prenume")
                    elif col == 3:
                        wb.add_value(selection, "Semnat. Venire")
                        wb.font_align(selection, wrap=True)
                    elif col == 4 or col == 6:
                        wb.add_value(selection, "Ora")
                    elif col == 5:
                        wb.add_value(selection, "Semnat plecare")
                        wb.font_align(selection, wrap=True)
                    elif col == 7:
                        wb.add_value(selection, "Pauza de masa")
                        wb.font_align(selection, wrap=True)
                    else: #8
                        wb.add_value(selection, "Observatii")
                else: # > 3..6
                    if col == 1:
                        wb.add_value(selection, str(row-2))
                    elif col == 2:
                        wb.add_value(selection, persons[row-3])
                    elif col == 4:
                        wb.add_value(selection, ORA_INCEPUT)
                    elif col == 6:
                        wb.add_value(selection, ORA_SFARSIT)
                    elif col == 7:
                        wb.add_value(selection, PAUZA)

                wb.border_style(selection)
            curentRow += 1

    # setat dimensiune coloane
    for col in range(1, 8+1): # -> h
        if col == 2: # skip B
            continue
        wb.column_auto_size(col)

    wb.save()
