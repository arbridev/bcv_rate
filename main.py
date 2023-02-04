## external imports

import os
import sys

## setup

appdir = os.path.dirname(os.path.realpath(__file__))
inputdir = os.path.join(appdir, "input")
outputdir = os.path.join(appdir, "output")

os.chdir(appdir)
sys.path.insert(1, os.getcwd())

## internal imports

from controller import Controller

## app

controller = Controller(inputdir, outputdir)

controller.start()