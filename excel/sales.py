#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pandas as pd
from pandas import ExcelWriter
# from openpyxl.utils import get_column_letter
# import xlsxwriter
import numpy as np
# from openpyxl import Workbook


Location = r'../../WorkingDirectory/sales.xlsx'

df = pd.read_excel(Location, 0)

print df.index
print df.describe()
print df['Kod'].describe()
print df.dtypes
print df.head()

report = df.reset_index().groupby(['Kod']).sum()
del report['cena']
del report['index']

print report
print report.index
# report.reset_index()
# print report.size()

writer = ExcelWriter('../../WorkingDirectory/Lesson3.xlsx')
# writer = '../../WorkingDirectory/Lesson3.xlsx'
report.to_excel(writer, engine='xlsxwriter')
writer.save()
print "Done"

# data = report.values.as_matrix()

# wbe = Workbook()
# ws = wbe.active
# ws.title = "Report"

# for row in data:
#     ws.append(row)

# wbe.save('../../WorkingDirectory/Lesson3.xlsx')
