#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
This script is for cleaning raw data from sales files
"""

import pandas as pd
import numpy as np

# enter periods to locate raw data file
print "Please enter the year of sales?"
year = raw_input('4 digits: --> ')
print "Please enter the month of sales?"
month = raw_input('2 digits: --> ')

# define path of raw data file
location = r'D:/repos/Python/excel/working_directory/rawdata/sales/sales_'
locationFile = location + year + month + r'.xlsx'


# define path of cleaned data file
locationSeb = r'D:/repos/Python/excel/working_directory/cleaneddata/seb/seb_'
locationSebFile = locationSeb + year + month + r'.xlsx'

destination = r'D:/repos/Python/excel/working_directory/cleaneddata/sales/sales_'
destinationFile = destination + year + month + r'.xlsx'

print "\nThe raw data file is: sales_%s%s.xlsx" % (year, month)
print "Is the file correct?\nFor 'Yes' hit Return, for 'No' hit CTRL-C (^C)."

raw_input("-->")

# sead data from file
df_VP = pd.read_excel(locationFile, sheetname="VP")
df_Iznos = pd.read_excel(locationFile, sheetname="Iznos")
df_seb = pd.read_excel(locationSebFile)

# consolidate data from sheets VP and Iznos in a single DataFrame
df = df_VP.append(df_Iznos,ignore_index=True)


# rename all columns
df.columns = ['pz_kod', 'pazar', 'kl_kod', 'klient', 'pr_kod', 'produkt', 'm',
                'kol', 'ed_bzc', 'bzc', 'ed_nc', 'nc', 'strana', '1', '2',
                '3', '4']

# remove last columns
del df['1']
del df['2']
del df['3']
del df['4']

# merge two DataFrame by column pr_kod
df_sales = pd.merge(df, df_seb, on='pr_kod', how='left', sort=False)

# add new columns
df_sales = df_sales.assign(seb=(df_sales['kol'] * df_sales['ed_seb']))
df_sales = df_sales.assign(br_marj=(df_sales['nc'] - df_sales['seb']))

# create a Pandas Excel writer using XlsxWriter as the engine
writer = pd.ExcelWriter(destinationFile, engine='xlsxwriter')
df_sales.to_excel(writer, index=False)

# save the results
writer.save()
