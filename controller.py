import os
from excel_handler import ExcelHandler

class Controller:
    inputdir = None
    outputdir = None
    excel_handler = ExcelHandler()

    def __init__(self, inputdir, outputdir):
        self.inputdir = inputdir
        self.outputdir = outputdir

    def start(self):
        print("Process started...")
        outputfile = os.path.join(self.outputdir, f'tasa_bcv_dolar_dia.xls')
        self.excel_handler.initialize_output_file(outputfile)
        quarterIDs = ["a", "b", "c", "d"]
        for id in quarterIDs:
            output = self.excel_handler.process_file(os.path.join(self.inputdir, f'2_1_2{id}22_smc.xls'))
            self.excel_handler.append_output_to_file(output, outputfile)
        print("Process finished.")
