#!/usr/bin/env python
# coding: utf-8

"""
Authors: Ritvik Kapila and Gauri Gupta
"""

# You will need to install python xlsb library and update pandas, if not already done, on or after 29 January, using the following command for reading xlsb files

# pip install pandas --upgrade

# pip install pyxlsb

# Now continue with the code after updation

import pandas as pd

# Function for conversion of xlsb file to xlsx

def xlsb2xlsx(filename, output_filename):
    dfs = pd.read_excel(filename, engine = 'pyxlsb', sheet_name = None)
    writer = pd.ExcelWriter(output_filename + '.xlsx', engine = 'xlsxwriter')
    for sheet_name in dfs.keys():
        dfs[sheet_name].to_excel(writer, sheet_name = sheet_name, index = False)
    writer.save()    

# Function for conversion of xlsm file to xlsx

def xlsm2xlsx(filename, output_filename):
    dfs = pd.read_excel(filename, sheet_name = None)
    writer = pd.ExcelWriter(output_filename + '.xlsx', engine = 'xlsxwriter')
    for sheet_name in dfs.keys():
        dfs[sheet_name].to_excel(writer, sheet_name = sheet_name, index = False)
    writer.save()    

# xlsb2xlsx('test_xlsb.xlsb', 'out_xlsb')

# xlsm2xlsx('test_xlsm.xlsm', 'out_xlsm')
