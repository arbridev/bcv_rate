import os
import datetime
from dateutil.relativedelta import relativedelta
from excel_handler import ExcelHandler

class Controller:
    inputdir = None
    outputdir = None
    year = None
    excel_handler = None

    def __init__(self, inputdir, outputdir, year):
        self.inputdir = inputdir
        self.outputdir = outputdir
        self.year = year
        self.excel_handler = ExcelHandler()

    def start(self):
        print("Process started...")
        outputfile = os.path.join(self.outputdir, f'bcv_daily_usd_rate_{self.year}.xlsx')
        self.excel_handler.initialize_output_file(outputfile)
        start_date = datetime.datetime(self.year, 1, 1)
        quarterIDs = ["a", "b", "c", "d"]
        for id in quarterIDs:
            next_start_date = start_date + relativedelta(months=3)
            end_date = next_start_date - relativedelta(days=1)
            output = self.excel_handler.process_file(os.path.join(self.inputdir, f'2_1_2{id}{self.year % 100}_smc.xlsx'), start_date, end_date)
            self.excel_handler.append_output_to_file(output, outputfile)
            start_date = next_start_date
        print("Process finished.")
