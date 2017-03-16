#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
This script is for consolidating cleaned data from sales files
"""

import pandas as pd
import numpy as np
import glob


# define path of sales data files
location = r'D:/repos/Python/excel/working_directory/cleaneddata/sales/sales_*.xlsx'


# define path of consolidated data file
destination = r'D:/repos/Python/excel/working_directory/cleaneddata/sales/salesData.xlsx'


# consolidate all sales data in a single DataFrame
all_data = pd.DataFrame()
for f in glob.glob(location):
    df = pd.read_excel(f)
    all_data = all_data.append(df,ignore_index=True)

# create a Pandas Excel writer using XlsxWriter as the engine
writer = pd.ExcelWriter(destination, engine='xlsxwriter')
all_data.to_excel(writer, index=False)

# save the results
writer.save()
