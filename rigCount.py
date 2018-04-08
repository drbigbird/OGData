# Download Baker Hughes rig count data

import pandas, datetime, requests, win32com.client, os
from datetime import datetime

import pdb, time
start = time.time()


# Temp for development
url = 'http://phx.corporate-ir.net/External.File?item=UGFyZW50SUQ9NjkxMDYzfENoaWxkSUQ9NDAxODE2fFR5cGU9MQ==&t=1'
tmp = os.getcwd() + '\\temp.xlsb'
fn = os.getcwd() + '\\BHGE rig data.xls'


def getRigFile():
	"""Download rig count *.xlsb file and convert into *.xls
	using Excel COM object"""
	
	res = requests.get(url)
	with open(tmp, 'wb') as output:
		output.write(res.content)
	output.close()

	# Save BHGE xlsb file as a xls file
	excel = win32com.client.Dispatch('Excel.Application')
	excel.DisplayAlerts = False
	excel.Visible = False
	doc = excel.Workbooks.Open(tmp)

	doc.SaveAs(fn, FileFormat=1)
	doc.Close()
	excel.Quit()


def cleanup():
	"""Remove Excel files"""

	os.remove(tmp)
	os.remove(fn)


getRigFile()
df = pandas.read_excel(fn, sheet_name='Master Data')
df.to_pickle('test_data.pickle')

cleanup()

print("Elapsed time: " + str(time.time()-start) + " sec")
