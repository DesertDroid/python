#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
This script is for cleaning raw data from seb files
"""

import pandas as pd
import numpy as np

# enter periods to locate raw data file
print "Please enter the year of seb?"
year = raw_input('4 digits: --> ')
print "Please enter the month of seb?"
month = raw_input('2 digits: --> ')

# define path of raw data file
location = r'D:/repos/Python/excel/working_directory/rawdata/seb/seb_'
locationFile = location + year + month + r'.xlsx'

# define path of cleaned data file
destination = r'D:/repos/Python/excel/working_directory/cleaneddata/seb/seb_'
destinationFile = destination + year + month + r'.xlsx'

print "\nThe raw data file is: seb_%s%s.xlsx" % (year, month)
print "Is the file correct?\nFor 'Yes' hit Return, for 'No' hit CTRL-C (^C)."

raw_input("-->")

# Read data from file
df = pd.read_excel(locationFile, sheetname="sr.seb", skiprows=3)

# remove first column
del df['Unnamed: 0']
# remove last row
df = df[:-1]
# rename Total column
df.rename(columns={'kod':'pr_kod', 'Total':'ed_seb'}, inplace=True)
# add two columns based on periods
df = df.assign(year=year, month=month)

# convert columns to number
df['pr_kod'] = pd.to_numeric(df['pr_kod'], errors='coerce')
df['year'] = pd.to_numeric(df['year'], errors='coerce')
df['month'] = pd.to_numeric(df['month'], errors='coerce')

# Create a Pandas Excel writer using XlsxWriter as the engine
writer = pd.ExcelWriter(destinationFile, engine='xlsxwriter')
df.to_excel(writer, index=False)

# Save the results
writer.save()
