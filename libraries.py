import time #built-in libraries-  written inside pyton interpreter in C langauge- import sys sys.builtin_module_names
import os #standard libraries- written in both C in python, reside in python libraries. import sys sys.prefix
import pandas #third party libraries- written by third party- to install pip3 install pandas

def datasheet(sheetPath):
    if os.path.exists(sheetPath):
        content= pandas.read_csv(sheetPath)
        name=list(content.TxnName)
        textcheck=list(content.TextCheck)
        return (name,textcheck)
    else:
        print("File doesn't exist")
