# Download Baker Hughes rig count data

import pandas, datetime
from pyxlsb import open_workbook as open_xlsb
from datetime import datetime

import pdb

# Temp for development
fn = 'BHI test pivot.xlsb'
url = 'http://phx.corporate-ir.net/External.File?item=UGFyZW50SUQ9NjkxMDYzfENoaWxkSUQ9NDAxODE2fFR5cGU9MQ==&t=1'

def convertToDF(fn):
	"""Converts rig count *.xlsb file into Pandas 
	dataframe using the following arguments:
		fn - Name of the rig count data file"""

print('Starting conversion...')
df = []
i = 1
with open_xlsb(fn) as wb:
	with wb.get_sheet('Master Data') as sheet:
		for row in sheet.rows(sparse=True):
			pdb.set_trace()
			for item in row:
				print(item.v)
				df.append([item.v])

df = pandas.DataFrame(df[1:], columns=df[0])
