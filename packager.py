# -*- coding: utf-8 -*-
# ---------------------------------------------------------------------------
# packager.py
# Created on: 2019-08-13 by Marlon Gonzalez and finalized on 2019-08-20
# Description: Packages data into shapefiles / KMZ's / Excel sheets based on desired specifications
# ---------------------------------------------------------------------------

print("Starting...")

# Imports modules
import os
import arcpy
from datatools import datafunctions
from datetime import datetime

start_time = datetime.now()

# Creates a feature layer from shapefile
def makelayer():
    linput = r"{}\SHP\{}_{}_{}.shp".format(homepath, charformatter(o), m, i)
    loutput = "{}_{}_{}_Layer".format(charformatter(o), m, i)
    arcpy.MakeFeatureLayer_management(linput, loutput)

# Checks if a feature layer already exists and creates one if not
def layerchecker():
    if not arcpy.Exists("{}_{}_{}_Layer".format(charformatter(o), m, i)):
        makelayer()

# Destroys feature layer at the end of the loop in order to prevent the computer's memory filling up
def destroylayer():
    if arcpy.Exists("{}_{}_{}_Layer".format(charformatter(o), m, i)):
        arcpy.Delete_management("{}_{}_{}_Layer".format(charformatter(o), m, i))
        print("destroying temporary layer")

# Defines first part of packaging process. Splits up the designated shapefile based off of a SQL query indicating desired fields
def packshapefile(*args):
    shppath = r"{}\SHP".format(homepath)
    pathchecker(shppath)
    outputfile = r"{}\{}_{}_{}.shp".format(shppath, charformatter(o), m, i)
    if filechecker(outputfile):
        sqlquery = """"SRC_RTE" = '{}' and "Main" = '{}'""".format(o, m)
        print("Creating {}_{}_{} shapefile...".format(o, m, i))
        arcpy.Select_analysis(inputfile, outputfile, sqlquery)
        print("Finished creating {}_{}_{} shapefile! Now starting KMZ conversion process".format(o, m, i))
    else:
        print("{}_{}_{} shapefile already exists, proceeding to KMZ conversion process".format(o, m, i))

# Defines second part of packaging process. Converts split shapefile into a KMZ file
def packkmz(*args):
    kmzpath = r"{}\KMZ".format(homepath)
    pathchecker(kmzpath)
    kmzoutputfile = r"{}\{}_{}_{}.kmz".format(kmzpath, charformatter(o), m, i)
    if filechecker(kmzoutputfile):
        layerchecker()
        print("Creating KMZ...")
        arcpy.LayerToKML_conversion("{}_{}_{}_Layer".format(charformatter(o), m, i), kmzoutputfile)
        print("Finished {}_{}_{} KMZ! Beginning Excel conversion process".format(o, m, i))
    else:
        print("{}_{}_{} KMZ already exists, proceeding to Excel conversion process".format(o, m, i))

# Defines last part of packaging process. Converts split shapefile into an excel sheet
def packexcel(*args):
    xlspath = r"{}\XLSX".format(homepath)
    pathchecker(xlspath)
    xlsoutputfile = r"{}\{}_{}_{}.xls".format(xlspath, charformatter(o), m, i)
    if filechecker(xlsoutputfile):
        layerchecker()
        print("Creating Excel sheet...")
        arcpy.TableToExcel_conversion("{}_{}_{}_Layer".format(charformatter(o), m, i), xlsoutputfile)
        print("Created Excel sheet and finished {}_{}_{} package!".format(o, m, i))
    else:
        print("{}_{}_{} Excel file already exists, finished {}_{}_{} package!".format(o, m, i, o, m, i))

# Inputting the parameters
print("Please input the following values (separate each value with a comma and a single space)")
sourceroute = (raw_input("Source Routes: ")).split(', ')
layer = (raw_input("Shapefile Name(s): ")).split(', ')
input_shpfile = (raw_input("Shapefile Path(s): ")).split(', ')
input_homepath = (raw_input("Destination Folder: ")).split(', ')
main = ["Main", "Lateral"]

# Iterative process
print("Initializing")
for i in layer:
    inputfile = input_shpfile[layer.index(i)]
    for o in sourceroute:
        homepath = r"{}\{}".format(input_homepath[0], o)
        for m in main:
            packshapefile()
            packkmz()
            packexcel()
            destroylayer()
            print("Now Looping...")
print("data packaging process successful!")
end_time = datetime.now()
print("Total execution time: {}".format(end_time - start_time))
