# -*- coding: utf-8 -*-
# ---------------------------------------------------------------------------
# pge_datafunctions.py
# Commonly used functions will be defined in this file and imported accordinglys
# ---------------------------------------------------------------------------
import os

# Defines character formatter function. which changes hyphens to underscores in order to prevent file name errors
def charformatter(field):
    newfield = field.replace("-","_")
    return newfield

# Defines path checking function. Checks to see if folders have yet to be created in shared drive and makes them if false
def pathchecker(path):
    checkpath = os.path.isdir(path)
    if checkpath != True:
        os.makedirs(path)

# Defines file checking function to prevent script from crashing if a file already exists with the same name
def filechecker(file):
    checkfile = os.path.exists(file)
    return not checkfile
