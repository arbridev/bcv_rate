## external imports

import os
import sys
import getopt

## setup

appdir = os.path.dirname(os.path.realpath(__file__))
inputdir = os.path.join(appdir, "input")
outputdir = os.path.join(appdir, "output")

os.chdir(appdir)
sys.path.insert(1, os.getcwd())

## internal imports

from controller import Controller

## app

def main(argv):
    global inputdir
    year = 2023
    opts, args = getopt.getopt(argv,"hy:p:",["year=", "path="])
    for opt, arg in opts:
        if opt == '-h':
            print ('main.py -y <year> -p <path on the input directory>')
            sys.exit()
        elif opt in ("-y", "--year"):
            year = int(arg)
        elif opt in ("-p", "--path"):
            inputdir = os.path.join(inputdir, arg)
    print ('Year is', year)
    controller = Controller(inputdir, outputdir, year)
    controller.start()

if __name__ == "__main__":
   main(sys.argv[1:])