import sys
import openpyxl
import datetime
from dateutil.relativedelta import relativedelta

class ExcelHandler:

    def initialize_output_file(self, outputfile):
        print(sys._getframe().f_code.co_name, outputfile)

    def process_file(self, inputfile, start_date, end_date):
        print(sys._getframe().f_code.co_name, inputfile)
        print(f'start: {start_date.strftime("%d %m %Y")} - end: {end_date.strftime("%d %m %Y")}')
        wb = openpyxl.load_workbook(inputfile, read_only=True) # workbook
        results = list()
        next_date = start_date
        while next_date != end_date:
            wsname = next_date.strftime("%d%m%Y")
            if wsname in wb.sheetnames:
                ws = wb[wsname]
                row = self.__get_currency_row(ws)
                value = self.__get_currency(ws, row).value
                results.append(tuple([next_date.strftime("%d-%m-%Y"), value]))
            else:
                results.append(tuple([next_date.strftime("%d-%m-%Y"), None]))
            next_date = next_date + relativedelta(days=1)
        return results

    def append_output_to_file(self, output, outputfile):
        print(sys._getframe().f_code.co_name)
        print(output)

    def __get_currency_row(self, worksheet):
        col = 2
        for row in range(11, 33):
            cell = worksheet.cell(row, col)
            if cell.value == "USD":
                return row
        return None

    def __get_currency(self, worksheet, row):
        col = 7
        return worksheet.cell(row, col)