#! python3

import requests, win32com.client, os

tmp = os.getcwd() + '\\temp.xlsb'
fn = os.getcwd() + '\\BHGE rig data.xls'

# URL for North America rig pivot data
res = requests.get('http://phx.corporate-ir.net/External.File?item=UGFyZW50SUQ9NjkwNzUwfENoaWxkSUQ9NDAxNDIwfFR5cGU9MQ==&t=1')
with open('temp.xlsb', 'wb') as output:
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

os.remove(tmp)
