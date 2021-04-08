from openpyxl import load_workbook
import shutil as sh

class Spreadsheets(object):

    def __init__(self):
        self.reimbursements_sheet = load_workbook(filename=sh.copyfile(r"eTransfer-Form-Blank.xlsx",r"eTransfer-Form-Fill.xlsx"))

mysheet = Spreadsheets()
