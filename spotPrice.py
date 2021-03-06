# Download O&G spot price data from EIA

import sys, datetime, pandas
from datetime import datetime

# Create lookup dictionary for converting commodity abbreviations
# (i.e., b - Brent, w - WTI, h - Henry Hub) to file name identifiers
# used on EIA website
abbrDict = {'b': 'pet/hist_xls/RBRTE', 'w': 'pet/hist_xls/RWTC', \
	'h': 'ng/hist_xls/RNGWHHD'}

# Create set containing valid time frequency options
freqOpt = {'d', 'w', 'm', 'a'}

# Define final output xlsx file name
exportfn = 'PriceData.xlsx'

def setOptions(args):
	"""Sets the options for what price to download
	with the arguments:
		comType - Commodity type
		freq - Frequency for pricing (mm/dd/yyyy)
		startDate - Starting date"""

	comType = ''
	freq = ''
	startDate = None
	dateInput = ''

	# Get commodity type
	if len(args) > 0 and args[0] in abbrDict:
		comType = args[0]
	else:
		while comType not in abbrDict:
			comType = input('Commodity (b - Brent/w - WTI/h - Henry Hub): ')

	# Get frequency
	if len(args) > 1 and args[1] in freqOpt:
		freq = args[1]
	else:
		while freq not in freqOpt:
			freq = input('Frequency (d - day/weekly - w/m - monthly/a - annual): ')

	# Get start date
	try:
		startDate = datetime.strptime(args[2],'%m/%d/%Y')
	except:
		pass
	
	while startDate == None:

		dateInput = input('Starting date (mm/dd/yyyy): ')
		
		try:
			startDate = datetime.strptime(dateInput,'%m/%d/%Y')
		except:
			if dateInput == '':
				startDate = datetime(1900,1,1)

	return comType, freq, startDate

# Get download options and construct URL string
comType, freq, startDate = setOptions(sys.argv[1:])
url = 'https://www.eia.gov/dnav/' + abbrDict[comType] + freq + '.xls'

# Download data and filter for start date
df = pandas.read_excel(url, sheet_name=1, header = 2)
df = df[df.Date > startDate]

# Export data to new .xlsx file
writer = pandas.ExcelWriter(exportfn)
df.to_excel(writer, 'Historical prices', startrow = 1, startcol = 1, \
	index = False)
writer.save()
