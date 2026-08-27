# controller.py
import os
import datetime
from dateutil.relativedelta import relativedelta
from excel_handler import ExcelHandler

class Controller:
    def __init__(self, inputdir, outputdir, year=None):
        """
        :param inputdir: Path to the folder containing input files.
        :param outputdir: Path where the final XLSX will be written.
        :param year: Optional. If None, defaults to the current calendar year.
        """
        self.inputdir = inputdir
        self.outputdir = outputdir

        # Resolve the year here – default to the current year if not supplied.
        self.year = year if year is not None else datetime.datetime.now().year

        self.excel_handler = ExcelHandler()

    def start(self):
        print("Process started...")
        outputfile = os.path.join(self.outputdir,
                                  f'bcv_daily_usd_rate_{self.year}.xlsx')
        self.excel_handler.initialize_output_file(outputfile)

        start_date = datetime.datetime(self.year, 1, 1)
        quarterIDs = ["a", "b", "c", "d"]

        for id in quarterIDs:
            next_start_date = start_date + relativedelta(months=3)
            end_date = next_start_date - relativedelta(days=1)

            input_file = os.path.join(
                self.inputdir, f'2_1_2{id}{self.year % 100}_smc.xlsx'
            )

            if not os.path.isfile(input_file):
                print(f"⚠️  Skipping quarter '{id}': file not found at {input_file}")
                start_date = next_start_date
                continue

            output = self.excel_handler.process_file(
                input_file, start_date, end_date
            )
            if output is not None:
                self.excel_handler.append_output_to_file(output, outputfile)

            start_date = next_start_date

        print("Process finished.")
