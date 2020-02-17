import os
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile

#defines texwriter function, writes data from extract block to text file
def textwriter():
    tw = open("example.txt","a")
    tw.write(data_extracted)
    tw.write(file_name)
    tw.write("\n")
    tw.close()

#defines extract block, iterates through list of excel file names and pulls data from specified row / column via pd.iloc method
file_name_list = [r"Y:\path\morepaths\example.xlsm", r"Y:\path\morepaths\example2.xlsm", r"Y:\path\anotherpath\example3.xlsm"] #enter all path names of excel sheets here
i = 0
while i < len(file_name_list):
    file_name = (file_name_list[i])
    df = pd.read_excel(file_name)
    data_extracted = str([df.iloc[0:0, 0:12]])
    textwriter()
    i += 1
print("done")
