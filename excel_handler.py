import openpyxl
import sys

class ExcelHandler:

    def initialize_output_file(self, outputfile):
        print(sys._getframe().f_code.co_name, outputfile)

    def process_file(self, inputfile):
        print(sys._getframe().f_code.co_name, inputfile)

    def append_output_to_file(self, output, outputfile):
        print(sys._getframe().f_code.co_name)
        print(output)
