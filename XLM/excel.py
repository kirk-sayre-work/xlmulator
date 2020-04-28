"""@package excel.py
Partial implementation of xlrd.book object interface.
"""

# --- IMPORTS ------------------------------------------------------------------

class ExcelSheet(object):

    def __init__(self, cells, name="Sheet1"):
        self.cells = cells
        self.name = name

    def __repr__(self):
        r = ""
        r += "Sheet: " + self.name + "\n\n"
        for cell in self.cells.keys():
            r += str(cell) + "\t=\t'" + str(self.cells[cell]) + "'\n"
        return r
    
    def cell(self, row, col):
        if ((row, col) in self.cells):
            return self.cells[(row, col)]
        raise KeyError("Cell (" + str(row) + ", " + str(col) + ") not found.")

    def cell_value(self, row, col):
        return self.cell(row, col)
    
class ExcelBook(object):

    def __init__(self, cells=None, name="Sheet1"):

        # Create empty workbook to fill in later?
        self.sheets = []
        if (cells is None):
            return

        # Create single sheet workbook?
        self.sheets.append(ExcelSheet(cells, name))

    def __repr__(self):
        r = ""
        for sheet in self.sheets:
            r += str(sheet) + "\n"
        return r
        
    def sheet_names(self):
        r = []
        for sheet in self.sheets:
            r.append(sheet.name)
        return r

    def sheet_by_index(self, index):
        if (index < 0):
            raise ValueError("Sheet index " + str(index) + " is < 0")
        if (index >= len(self.sheets)):
            raise ValueError("Sheet index " + str(index) + " is > num sheets (" + str(len(self.sheets)) + ")")
        return self.sheets[index]

    def sheet_by_name(self, name):
        for sheet in self.sheets:
            if (sheet.name == name):
                return sheet
        raise ValueError("Sheet name '" + str(name) + "' not found.")

def make_book(cell_data):
    return ExcelBook(cell_data)
