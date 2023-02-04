import openpyxl
import datetime
import sys

class ExcelHandler:

    def initialize_output_file(self, outputfile):
        print(sys._getframe().f_code.co_name, outputfile)

    def process_file(self, inputfile, start_date, end_date):
        print(sys._getframe().f_code.co_name, inputfile)
        print(f'start: {start_date.strftime("%d %m %Y")} - end: {end_date.strftime("%d %m %Y")}')

    def append_output_to_file(self, output, outputfile):
        print(sys._getframe().f_code.co_name)
        print(output)
