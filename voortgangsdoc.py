import openpyxl

class SourceXls:
    def __init__(self):
        self.file = "IO2A4-voortgang 2017-2018.xlsx"
        self.klas = self.file[0:5]
        self.pad = "C:\\Users\\Abu\\Documents\\Python\\"
        self.wb = openpyxl.load_workbook(self.pad + self.file, data_only=True)
        self.sheet = self.wb.get_sheet_by_name(self.klas)
        self.vakken = self.wb.get_sheet_names()
        self.actsheetindex = self.vakken.index(self.klas)
        self.actleerlingindex = -1
        self.leerlingen = []
        self.leerlingDocs = []
        self.namen = []
        self.llcoordinaten = []

    def initsheetindex(self):
        self.actsheetindex = self.vakken.index(self.klas)

    #Leerlingen staan in de tab met de naam van de klas en starten op de 4de rij
    def initleerlingen(self):
        j = 0
        for i in range(1, 50, 1):
            leerling = self.sheet.cell(row=i, column=1).value
            naam = self.sheet.cell(row=i, column=2).value

            if leerling is not None and naam is not None and str(leerling).isdigit():
                # print(str(leerling) + ' ' + str(naam))
                self.leerlingen.append(leerling)
                self.namen.append(naam)
                self.llcoordinaten.append("=" + self.klas + "!"  + self.sheet.cell(row=i, column=2).coordinate)
                j += 1
        print(self.leerlingen);

    def setnextsheet(self):
        if self.actsheetindex <= len(self.vakken) - 2:
            self.actsheetindex += 1
            self.sheet = self.wb.get_sheet_by_name(self.vakken[self.actsheetindex])
            return True
        else:
            return False

    def getsheetname(self):
        return self.vakken[self.actsheetindex]

    def setnextleerling(self):
        if self.actleerlingindex <= len(self.leerlingen) - 2:
            self.actleerlingindex += 1
            return True
        else:
            return False

    def zoekleerlingrij(self):
        j = 0
        for i in range(1, 50, 1):
            naam = self.sheet.cell(row=i, column=1)._value
            if naam is not None:
                if naam == self.namen[self.actleerlingindex] or naam == self.llcoordinaten[self.actleerlingindex]:
                    return i
        return -1

    # Wat is de rij waar de resultaten staan van de current leerling
    def zoekopleverrij(self):
        return 4 + self.actleerlingindex

    def kopieerrange(self, startCol, startRow, endCol, endRow):
        sheet = self.sheet
        rangeSelected = []
        # Loops through selected Rows
        for i in range(startRow, endRow + 1, 1):
            # Appends the row to a RowSelected list
            rowSelected = []
            for j in range(startCol, endCol + 1, 1):
                rowSelected.append(sheet.cell(row=i, column=j).value)
            # Adds the RowSelected List and nests inside the rangeSelected
            rangeSelected.append(rowSelected)

        return rangeSelected

