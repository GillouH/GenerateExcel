from tkinter import Tk, Label, W, E, Frame, StringVar, Entry, Button, filedialog, OptionMenu, Spinbox, _setit
from tkinter.messagebox import showinfo, showwarning, showerror
from os import getcwd, chdir
from re import compile
from datetime import date
from locale import setlocale, LC_ALL
from subprocess import run
from xlsxwriter import Workbook
from xlsxwriter.utility import xl_cell_to_rowcol, xl_rowcol_to_cell
from json import loads, dumps


setlocale(LC_ALL, "fr_FR")
today = date.today()
monthList = ["{:%B}".format(date(today.year, month, 1)).capitalize() for month in range(1, 13)]


class Window(Tk):
    fileExtension = ".xlsx"
    devise = "€"
    floatingNumberRegex = compile("^[-+]?(0|[1-9]\d*)([,\.]\d{1,2})?$")
    memoryFile = "memory.txt"
    
    JSON_KEY = {
        "folder": "JSON_KEY_FOLDER",
        "file": "JSON_KEY_FILE",
        "solde": "JSON_KEY_SOLDE",
        "monthEnd": "JSON_KEY_MONTH_END",
        "yearEnd": "JSON_KEY_YEAR_END"
    }

    class FONT:
        LARGE = (None, 18)
        SMALL = (None, 15)

    def __init__(self, title=None):
        super().__init__()
        if title is not None:
            self.title(title)

        self.folderNameValue = StringVar()
        self.fileNameValue = StringVar()
        self.soldeValue = StringVar()
        self.monthEnd = StringVar()
        self.yearEnd = StringVar()
        self.restoreState()

        self.createFolderFrame(row=0, column=0)
        self.createFileFrame(row=1, column=0)
        self.createSoldeFrame(row=2, column=0)
        self.createDateFrame(row=3, column=0)
        Button(self, text="Générer", command=self.generate, font=Window.FONT.SMALL).grid(row=4, column=0, sticky=E)

    def createFrame(self, row, column, sticky=None, *paramList, **paramDict):
        frame = Frame(self, *paramList, **paramDict)
        frame.grid(row=row, column=column, sticky=sticky)
        return frame

    def createFolderFrame(self, row, column):
        frame = self.createFrame(row, column, sticky=W)

        Label(frame, text="Choisir l'emplacement du fichier :", font=Window.FONT.LARGE).grid(row=0, column=0, sticky=W)
        Entry(frame, textvariable=self.folderNameValue, width=50, font=Window.FONT.SMALL).grid(row=1, column=0)
        Button(frame, text="Parcourir", command=lambda: self.folderNameValue.set(filedialog.askdirectory()), font=Window.FONT.SMALL).grid(row=1, column=1)

    def createFileFrame(self, row, column):
        frame = self.createFrame(row, column, sticky=W)

        Label(frame, text="Nom du fichier : ", font=Window.FONT.LARGE).grid(row=0, column=0)
        Entry(frame, textvariable=self.fileNameValue, width=20, justify="right", font=Window.FONT.SMALL).grid(row=0, column=1)
        Label(frame, text=Window.fileExtension, font=Window.FONT.SMALL).grid(row=0, column=3)

    def createSoldeFrame(self, row, column):
        frame = self.createFrame(row, column, sticky=W)

        Label(frame, text="Solde actuel : ", font=Window.FONT.LARGE).grid(row=0, column=0)
        Entry(frame, textvariable=self.soldeValue, width=10, justify="right", font=Window.FONT.SMALL).grid(row=0, column=1)
        Label(frame, text=Window.devise, font=Window.FONT.SMALL).grid(row=0, column=3)

    def createDateFrame(self, row, column):
        frame = self.createFrame(row, column, sticky=W)

        Label(frame, text="Générer jusqu'au mois de : ", font=Window.FONT.LARGE).grid(row=0, column=0)
        self.optionMenu = OptionMenu(frame, self.monthEnd, *monthList[today.month-1:])
        self.optionMenu.configure(takefocus=True, font=Window.FONT.SMALL)
        self.optionMenu.grid(row=0, column=1)
        self.spinbox = Spinbox(frame, from_=today.year, to_=today.year+10, textvariable=self.yearEnd, font=Window.FONT.SMALL)
        self.spinbox.grid(row=0, column=2)
        def checkDate(evt=None):
            self.checkDate()
        self.spinbox["command"] = checkDate
        self.spinbox.bind("<Return>", checkDate)
        checkDate()

    def checkDate(self, generate=False):
        if not self.yearEnd.get().isdigit() or int(self.yearEnd.get()) < today.year:
            if generate:
                return False
            self.yearEnd.set(today.year)
        if int(self.yearEnd.get()) == today.year:
            self.refreshMonthMenu(monthList[today.month-1:])
            if self.monthEnd.get() in monthList[:today.month-1]:
                if generate:
                    return False
                self.monthEnd.set(monthList[today.month-1])
        else:
            self.refreshMonthMenu(monthList)
        self.spinbox["to"] = int(self.yearEnd.get())+1
        return True

    def refreshMonthMenu(self, values):
        self.optionMenu["menu"].delete(0, "end")
        for val in values:
            self.optionMenu["menu"].add_command(label=val, command=_setit(self.monthEnd, val))
    
    def generate(self):
        if not Window.floatingNumberRegex.match(self.soldeValue.get()):
            showwarning("Solde incorrect", "Le solde rengeigné est incorrecte : {}€".format(self.soldeValue.get()))
        elif not self.checkDate(generate=True):
            showwarning("Date incorrecte", "La date renseignée est incorrecte : {} {}".format(self.monthEnd.get(), self.yearEnd.get()))
        else:
            self.saveState()
            self.generateExcelFile()
            self.finish()

    def restoreState(self):
        chdir(getcwd())
        try:
            data = None
            with open(Window.memoryFile, "r") as file:
                data = loads(file.read())
            if data is not None:
                self.folderNameValue.set(data[Window.JSON_KEY["folder"]])
                self.fileNameValue.set(data[Window.JSON_KEY["file"]])
                self.soldeValue.set(data[Window.JSON_KEY["solde"]])
                self.monthEnd.set(data[Window.JSON_KEY["monthEnd"]])
                self.yearEnd.set(data[Window.JSON_KEY["yearEnd"]])
        except:
            self.folderNameValue.set(getcwd())
            self.fileNameValue.set("compte")
            self.soldeValue.set(0)
            self.monthEnd.set(monthList[today.month-1])
            self.yearEnd.set(today.year)

    def saveState(self):
        chdir(getcwd())
        data = {
            Window.JSON_KEY["folder"]: self.folderNameValue.get(),
            Window.JSON_KEY["file"]: self.fileNameValue.get(),
            Window.JSON_KEY["solde"]: self.soldeValue.get(),
            Window.JSON_KEY["monthEnd"]: self.monthEnd.get(),
            Window.JSON_KEY["yearEnd"]: self.yearEnd.get()
        }
        with open(Window.memoryFile, "w") as file:
            file.write(dumps(data, indent=4, ensure_ascii=False))

    def generateExcelFile(self):
        chdir(self.folderNameValue.get())
        excelFile = ExcelFile(self.fileNameValue.get()+Window.fileExtension)
        excelFile.generate(monthEndIndex=monthList.index(self.monthEnd.get()), yearEnd=int(self.yearEnd.get()), solde=float(self.soldeValue.get().replace(",", ".")))

    def finish(self):
        run(["explorer", self.folderNameValue.get().replace("/", "\\")])
        self.destroy()


class ExcelFile(Workbook):
    rowMax = 1048576
    columnMax = 16383

    FG = 0
    BG = 1

    def generate(self, monthEndIndex, yearEnd, solde):
        rowIndex = 2
        columnIndex = 2
        self.sheet = self.add_worksheet()
        self.generateFormatsDicts()
        self.setFormatColumn(columnIndex=columnIndex)
        rowIndex = self.setHeader(rowIndex=rowIndex, columnIndex=columnIndex)
        for yearToGenerate in range(yearEnd, today.year-1, -1):
            rowIndex = self.generateYear(
                rowIndex=rowIndex,
                columnIndex=columnIndex,
                yearToGenerate=yearToGenerate,
                monthEndIndex=monthEndIndex,
                yearEnd=yearEnd,
                solde=solde,
                upperYear=yearToGenerate==yearEnd,
                bottomYear=yearToGenerate==today.year
            )
        self.close()

    def generateFormatsDicts(self):
        ground = "ground"
        rotation = "rotation"
        num_format = "num_format"
        self.backgroundFormatDict = {ground: ExcelFile.BG}
        self.monthYearFormatDitct = {ground: ExcelFile.FG, rotation: 90}
        self.deviseFormatDict = {ground: ExcelFile.FG, num_format: '[White]_-#,##0.00" "€;[Black]-#,##0.00" "€'}
        self.dateFormatDict = {ground: ExcelFile.FG, num_format: "dd/mm/yyyy"}

    def createNewFormat(self, ground, **params):
        formatTmp = {"align": "center", "valign": "vcenter"}
        if ground is not None:
            bgColor = "bg_color"
            if ground == ExcelFile.BG:
                formatTmp[bgColor] = "#FFFF66"
            if ground == ExcelFile.FG:
                formatTmp[bgColor] = "#FF9999"
        formatTmp.update(params)
        return self.add_format(formatTmp)

    def setFormatColumn(self, columnIndex=0):
        backgroundFormat = self.createNewFormat(**self.backgroundFormatDict)
        if columnIndex > 0:
            self.sheet.set_column(0, columnIndex-1, None, backgroundFormat)
        numberKey = "numberKey"
        sizeKey = "sizeKey"
        columnSize = [
            {numberKey: 3, sizeKey: 10.71}, # Date
            {numberKey: 1, sizeKey: 16.29}, # Motif
            {numberKey: 1, sizeKey: 13.29}, # Commentaire
            {numberKey: 2, sizeKey: 10.71}, # Montant - Passé ?
            {numberKey: 1, sizeKey: 12.71}, # Date Passage
            {numberKey: 1, sizeKey: 18.14}, # Solde Prévisionnel
            {numberKey: 1, sizeKey: 10.71}, # Solde Réel
        ]
        for i in range(len(columnSize)):
            self.sheet.set_column(columnIndex, columnIndex+columnSize[i][numberKey], columnSize[i][sizeKey], backgroundFormat)
            columnIndex += columnSize[i][numberKey]
        self.sheet.set_column(columnIndex, ExcelFile.columnMax, None, backgroundFormat)

    def setHeader(self, rowIndex, columnIndex):
        firstRowIndex = rowIndex
        firstColumnIndex = columnIndex

        self.sheet.merge_range(rowIndex, columnIndex, rowIndex, columnIndex+2, "Date", self.createNewFormat(ground=ExcelFile.FG, top=5, left=5, right=2, bottom=6))

        rowIndexTmp = rowIndex + 1
        self.sheet.write(rowIndexTmp, columnIndex, "Année", self.createNewFormat(ground=ExcelFile.FG, left=5, right=2, bottom=5))
        columnIndex += 1
        for columnName in ["Mois", "Complète"]:
            self.sheet.write(rowIndexTmp, columnIndex, columnName, self.createNewFormat(ground=ExcelFile.FG, right=2, bottom=5))
            columnIndex += 1

        for columnName in ["Motif", "Commentaire", "Montant", "Passé ?", "Date Passage", "Solde Prévisionnel"]:
            self.sheet.merge_range(rowIndex, columnIndex, rowIndex+1, columnIndex, columnName, self.createNewFormat(ground=ExcelFile.FG, top=5, right=1, bottom=5))
            columnIndex+=1
        self.sheet.merge_range(rowIndex, columnIndex, rowIndex+1, columnIndex, "Solde Réel", self.createNewFormat(ground=ExcelFile.FG, top=5, right=5, bottom=5))
        rowIndex += 2

        # Vue rapide
            # Année
        columnIndex = firstColumnIndex
        for rowIndexTmp in range(rowIndex, rowIndex+2):
            self.sheet.write(rowIndexTmp, columnIndex, None, self.createNewFormat(ground=ExcelFile.FG, left=5, right=2, bottom=1))
        rowIndexTmp += 1
        self.sheet.write(rowIndexTmp, columnIndex, None, self.createNewFormat(ground=ExcelFile.FG, left=5, right=2, bottom=5))
        columnIndex += 1

            # Mois - Complète
        for columnIndexTmp in range(columnIndex, columnIndex+2):
            for rowIndexTmp in range(rowIndex, rowIndex+2):
                self.sheet.write(rowIndexTmp, columnIndexTmp, None, self.createNewFormat(ground=ExcelFile.FG, right=2, bottom=1))
            rowIndexTmp += 1
            self.sheet.write(rowIndexTmp, columnIndexTmp, None, self.createNewFormat(ground=ExcelFile.FG, right=2, bottom=5))
        columnIndex = columnIndexTmp + 1

            # Motif - Commentaire - Montant - Passé ? - Date Passage
        for columnIndexTmp in range(columnIndex, columnIndex+5):
            for rowIndexTmp in range(rowIndex, rowIndex+2):
                self.sheet.write(rowIndexTmp, columnIndexTmp, None, self.createNewFormat(ground=ExcelFile.FG, right=1, bottom=1))
            rowIndexTmp += 1
            self.sheet.write(rowIndexTmp, columnIndexTmp, None, self.createNewFormat(ground=ExcelFile.FG, right=1, bottom=5))
        columnIndex = columnIndexTmp + 1

            # Solde Prévisionnel
        rowIndexTmp = rowIndex
        self.sheet.write(rowIndexTmp, columnIndex, None, self.createNewFormat(ground=ExcelFile.FG, right=1, bottom=1))
        rowIndexTmp += 1
        self.sheet.write(rowIndexTmp, columnIndex, "={}".format(xl_rowcol_to_cell(rowIndexTmp+2, columnIndex)), self.createNewFormat(**self.deviseFormatDict, right=1, bottom=1))
        rowIndexTmp += 1
        self.sheet.write(rowIndexTmp, columnIndex, None, self.createNewFormat(ground=ExcelFile.FG, right=1, bottom=5))
        columnIndex += 1

            # Solde Réel
        self.sheet.write(rowIndex, columnIndex, None, self.createNewFormat(ground=ExcelFile.FG, right=5, bottom=1))
        rowIndex += 1
        self.sheet.write(rowIndex, columnIndex, "={}".format(xl_rowcol_to_cell(rowIndex+2, columnIndex)), self.createNewFormat(**self.deviseFormatDict, right=5, bottom=1))
        rowIndex += 1
        self.sheet.write(rowIndex, columnIndex, None, self.createNewFormat(ground=ExcelFile.FG, right=5, bottom=5))
        rowIndex += 1

        return rowIndex

    def generateYear(self, rowIndex, columnIndex, yearToGenerate, monthEndIndex, yearEnd, solde, upperYear=False, bottomYear=False):
        firstRowIndex = rowIndex
        firstColumnIndex = columnIndex

        if upperYear:
            columnIndex += 1

            # Mois - Complète
            for columnIndexTmp in range(columnIndex, columnIndex+2):
                self.sheet.write(rowIndex, columnIndexTmp, None, self.createNewFormat(ground=ExcelFile.FG, right=2, bottom=1))
                self.sheet.write(rowIndex+1, columnIndexTmp, None, self.createNewFormat(ground=ExcelFile.FG, right=2, bottom=2))
            columnIndex = columnIndexTmp + 1

            # Motif - Commentaire - Montant - Passé ? - Date Passage
            for columnIndexTmp in range(columnIndex, columnIndex+5):
                self.sheet.write(rowIndex, columnIndexTmp, None, self.createNewFormat(ground=ExcelFile.FG, right=1, bottom=1))
                self.sheet.write(rowIndex+1, columnIndexTmp, None, self.createNewFormat(ground=ExcelFile.FG, right=1, bottom=2))
            columnIndex = columnIndexTmp + 1

            # Solde Prévisionnel
            self.sheet.write(rowIndex, columnIndex, ExcelFile.formuleSP(rowIndex, columnIndex), self.createNewFormat(**self.deviseFormatDict, right=1, bottom=1))
            self.sheet.write(rowIndex+1, columnIndex, ExcelFile.formuleSP(rowIndex+1, columnIndex), self.createNewFormat(**self.deviseFormatDict, right=1, bottom=2))
            columnIndex += 1

            # Solde Réel
            self.sheet.write(rowIndex, columnIndex, ExcelFile.formuleSR(rowIndex, columnIndex), self.createNewFormat(**self.deviseFormatDict, right=5, bottom=1))
            self.sheet.write(rowIndex+1, columnIndex, ExcelFile.formuleSR(rowIndex+1, columnIndex), self.createNewFormat(**self.deviseFormatDict, right=5, bottom=2))
            columnIndex = firstColumnIndex
            rowIndex += 2

        lastMonthIndex = monthEndIndex if upperYear else len(monthList) - 1
        firstMonthIndex = today.month - 1 if bottomYear else 0

        for monthIndex in range(lastMonthIndex, firstMonthIndex-1, -1):
            monthBottomBorder = 5 if not bottomYear and monthIndex == firstMonthIndex else 2

            # Mois
            columnIndex += 1
            self.sheet.merge_range(rowIndex, columnIndex, rowIndex+4, columnIndex, monthList[monthIndex].upper(), self.createNewFormat(**self.monthYearFormatDitct, right=2, bottom=monthBottomBorder))
            # Complète
            columnIndex += 1
            for rowIndexTmp in range(rowIndex, rowIndex+4):
                self.sheet.write(rowIndexTmp, columnIndex, None, self.createNewFormat(**self.dateFormatDict, right=2, bottom=1))
            self.sheet.write(rowIndexTmp+1, columnIndex, None, self.createNewFormat(**self.dateFormatDict, right=2, bottom=monthBottomBorder))
            columnIndex += 1

            # Motif - Commentaire
            for columnIndexTmp in range(columnIndex, columnIndex+2):
                for rowIndexTmp in range(rowIndex, rowIndex+4):
                    self.sheet.write(rowIndexTmp, columnIndexTmp, None, self.createNewFormat(ground=ExcelFile.FG, right=1, bottom=1))
                self.sheet.write(rowIndexTmp+1, columnIndexTmp, None, self.createNewFormat(ground=ExcelFile.FG, right=1, bottom=monthBottomBorder))
            columnIndex = columnIndexTmp + 1

            # Montant
            for rowIndexTmp in range(rowIndex, rowIndex+4):
                self.sheet.write(rowIndexTmp, columnIndex, None, self.createNewFormat(**self.deviseFormatDict, right=1, bottom=1))
            self.sheet.write(rowIndexTmp+1, columnIndex, None, self.createNewFormat(**self.deviseFormatDict, right=1, bottom=monthBottomBorder))
            columnIndex += 1

            # Passé ?
            self.sheet.write(rowIndex, columnIndex, None, self.createNewFormat(ground=ExcelFile.FG, right=1, bottom=1))
            for rowIndexTmp in range(rowIndex+1, rowIndex+4):
                self.sheet.write(rowIndexTmp, columnIndex, "Non", self.createNewFormat(ground=ExcelFile.FG, right=1, bottom=1))
            self.sheet.write(rowIndexTmp+1, columnIndex, "Non", self.createNewFormat(ground=ExcelFile.FG, right=1, bottom=monthBottomBorder))
            columnIndex += 1

            # Date Passage
            for rowIndexTmp in range(rowIndex, rowIndex+4):
                self.sheet.write(rowIndexTmp, columnIndex, None, self.createNewFormat(**self.dateFormatDict, right=1, bottom=1))
            self.sheet.write(rowIndexTmp+1, columnIndex, None, self.createNewFormat(**self.dateFormatDict, right=1, bottom=monthBottomBorder))
            columnIndex += 1

            # Solde Prévisionnel
            for rowIndexTmp in range(rowIndex, rowIndex+4):
                self.sheet.write(rowIndexTmp, columnIndex, ExcelFile.formuleSP(rowIndexTmp, columnIndex), self.createNewFormat(**self.deviseFormatDict, right=1, bottom=1))
            self.sheet.write(rowIndexTmp+1, columnIndex, ExcelFile.formuleSP(rowIndexTmp+1, columnIndex), self.createNewFormat(**self.deviseFormatDict, right=1, bottom=monthBottomBorder))
            columnIndex += 1

            # Solde Réel
            for rowIndexTmp in range(rowIndex, rowIndex+4):
                self.sheet.write(rowIndexTmp, columnIndex, ExcelFile.formuleSR(rowIndexTmp, columnIndex), self.createNewFormat(**self.deviseFormatDict, right=5, bottom=1))
            self.sheet.write(rowIndexTmp+1, columnIndex, ExcelFile.formuleSR(rowIndexTmp+1, columnIndex), self.createNewFormat(**self.deviseFormatDict, right=5, bottom=monthBottomBorder))

            columnIndex = firstColumnIndex
            rowIndex = rowIndexTmp + 2

        if bottomYear:
            columnIndex += 1

            # Mois
            self.sheet.write(rowIndex, columnIndex, "BASE", self.createNewFormat(ground=ExcelFile.FG, right=2, bottom=5))
            columnIndex += 1

            # Complète
            self.sheet.write(rowIndex, columnIndex, "{:%x}".format(today), self.createNewFormat(**self.dateFormatDict, right=2, bottom=5))
            columnIndex += 1

            # Motif
            self.sheet.write(rowIndex, columnIndex, "SOLDE", self.createNewFormat(ground=ExcelFile.FG, right=1, bottom=5))
            columnIndex += 1

            # Commentaires
            self.sheet.write(rowIndex, columnIndex, None, self.createNewFormat(ground=ExcelFile.FG, right=1, bottom=5))
            columnIndex += 1

            # Montant
            self.sheet.write(rowIndex, columnIndex, solde, self.createNewFormat(**self.deviseFormatDict, right=1, bottom=5))
            columnIndex += 1

            # Passé ?
            self.sheet.write(rowIndex, columnIndex, "Oui", self.createNewFormat(ground=ExcelFile.FG, right=1, bottom=5))
            columnIndex += 1

            # Date Passage
            self.sheet.write(rowIndex, columnIndex, None, self.createNewFormat(**self.dateFormatDict, right=1, bottom=5))
            columnIndex += 1

            # Solde Provisoir
            self.sheet.write(rowIndex, columnIndex, ExcelFile.formuleSP(rowIndex, columnIndex, solde=True), self.createNewFormat(**self.deviseFormatDict, right=1, bottom=5))
            columnIndex += 1

            # Solde Réel
            self.sheet.write(rowIndex, columnIndex, ExcelFile.formuleSR(rowIndex, columnIndex, solde=True), self.createNewFormat(**self.deviseFormatDict, right=5, bottom=5))
            columnIndex = firstColumnIndex
            rowIndex += 1

        lastRowIndex = rowIndex - 1
        self.sheet.merge_range(firstRowIndex, columnIndex, lastRowIndex, columnIndex, yearToGenerate, self.createNewFormat(**self.monthYearFormatDitct, left=5, right=2, bottom=5))

        return rowIndex

    @classmethod
    def formuleSP(cls, rowIndex, columnIndex, solde=False):
        if solde:
            return "={}".format(xl_rowcol_to_cell(rowIndex, columnIndex-3))
        else:
            return "={}+{}".format(xl_rowcol_to_cell(rowIndex+1, columnIndex), xl_rowcol_to_cell(rowIndex, columnIndex-3))

    @classmethod
    def formuleSR(cls, rowIndex, columnIndex, solde=False):
        if solde:
            return "=IF({}=\"Oui\",{},0)".format(xl_rowcol_to_cell(rowIndex, columnIndex-3), xl_rowcol_to_cell(rowIndex, columnIndex-4))
        else:
            return "={}+IF({}=\"Oui\",{})".format(xl_rowcol_to_cell(rowIndex+1, columnIndex), xl_rowcol_to_cell(rowIndex, columnIndex-3), xl_rowcol_to_cell(rowIndex, columnIndex-4))


def main():
    window = Window("Génération de fichier Excel pour les comptes.")
    window.mainloop()


if __name__ == "__main__":
    main()