import csv, sys, json
from openpyxl import Workbook
import xlsxwriter

json_data = open("dictionary.json")
dictData = json.load(json_data)
wb = Workbook()
ws = wb.active

def parseWriteHeaders(reader):
	reader.next()
	ws.append(["Date", "Vendor", "Description", "Type", "Transaction Medium", "Debit"])

def parseWriteData(row):
	date, desc, origDesc, amount, transaction, category, account, labels, notes = row
	amount = amountParser(row)
	account = accountParser(row)
	category = categoryParser(row)

	if account and category:
		ws.append([date, desc, notes, category, account, amount])

def isVenmo(origDesc):
	return 'venmo' in origDesc.lower()

def accountParser(row):
	date, desc, origDesc, amount, transaction, category, account, labels, notes = row
	if (isVenmo(origDesc)):
		return 'Venmo'
	try:
		return dictData['description'][account]
	except KeyError:
		pass
	return account

def categoryParser(row):
	date, desc, origDesc, amount, transaction, category, account, labels, notes = row
	try:
		return dictData['category'][category]
	except KeyError:
		pass
	return "Misc"

def amountParser(row):
	date, desc, origDesc, amount, transaction, category, account, labels, notes = row
	if transaction == "credit":
		return float(amount)*-1
	return float(amount)

with open(sys.argv[1], 'rb') as mintData:
	reader = csv.reader(mintData)
	parseWriteHeaders(reader)
	for row in reader:
		parseWriteData(row)

wb.save("tester.xlsx")
workbook = xlsxwriter.Workbook('tester.xlsx')
worksheet = workbook.add_worksheet()

with open(sys.argv[1], 'rb') as mintData:
	reader = csv.reader(mintData)
	parseWriteHeaders(reader)
	rowNum = 2
	for row in reader:
		worksheet.write_datetime(rowNum, 0, row[0], "mm/dd/yyyy")
		rowNum += 1

workbook.close()