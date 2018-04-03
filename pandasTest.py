# Test code for evaluating pandas to replace xlrd/openpyxl

import pandas
fn = 'text.xls'

df = pandas.read_excel(open(fn))