from time import time
import xlsxwriter
import json
import os
from datetime import date, datetime, timedelta

def get_change(current, previous):
    if current == previous:
        return 0
    try:
        result = (abs(current - previous) / previous) * 100.0
        if previous > current:
            return -result
        else:
            return result 
    except ZeroDivisionError:
        return float('inf')


def getCellFormat(percentage, good, bad):
    if percentage > 0:
        return good
    elif percentage < 0:
        return bad
    else:
        return {}

# change the date depeding from where you want the report to start
# minmum start date is 13-11-2021 as that is the earliest we have data
startDate = '13-11-2021'

file = open('files/data-'+startDate+'.json',)
 
data = json.load(file)
exportFileName = 'CustomReport-' + startDate + 'to'+ datetime.now().strftime('%d-%m-%Y')
workbook = xlsxwriter.Workbook('files/customReports/'+exportFileName+'.xlsx')

goodStyle = workbook.add_format()
goodStyle.set_bg_color('#C6EFCE')
goodStyle.set_font_color('#006100')

badStyle = workbook.add_format()
badStyle.set_bg_color('#FFC7CE')
badStyle.set_font_color('#9C0006')

worksheet = workbook.add_worksheet()
worksheet.set_column(0,0,25)
worksheet.set_column(1,1,10)
worksheet.set_column(2,50,25)
header_format = workbook.add_format({'bold': True})
worksheet.write('A1', 'Name', header_format)
worksheet.write('B1', 'Symbol', header_format)
worksheet.write('C1', 'Network', header_format)
worksheet.write('D1', 'Price', header_format)
worksheet.write('E1', 'Added', header_format)
entryNumber = 2
entryMap = {}
priceMap = {}

for entry in data['data']:
    dateAdded=datetime.strptime(entry["date_added"].split('T')[0], '%Y-%m-%d')
    worksheet.write('A'+str(entryNumber), entry["name"])
    worksheet.write('B'+str(entryNumber), entry["symbol"])
    entryMap[entry['symbol']] = entryNumber
    if entry['platform'] is not None:
        worksheet.write('C'+str(entryNumber), entry["platform"]['name'])
    numbers = entry['quote']['USD']
    worksheet.write('D'+str(entryNumber), str(numbers['price']))
    priceMap[entry['symbol']] = numbers['price']
    worksheet.write('E'+str(entryNumber), entry['date_added'])
    entryNumber += 1
     
file.close()
lastChar = 'E'
d = datetime.strptime(startDate, '%d-%m-%Y')

while d <= datetime.today():
    d = d + timedelta(days=1)
    if os.path.exists('files/data-'+d.strftime('%d-%m-%Y')+'.json'):
        file = open('files/data-'+d.strftime('%d-%m-%Y')+'.json')
        compareData = json.load(file)
        worksheet.write(chr(ord(lastChar) + 1) + '1', d.date().strftime('%d-%m-%Y'), header_format)
        worksheet.write(chr(ord(lastChar) + 2) + '1', 'Change from '+ startDate, header_format)
        for entry in compareData['data']:
            if entry['symbol'] in entryMap:
                numbers = entry['quote']['USD']
                worksheet.write(chr(ord(lastChar) + 1) + str(entryMap[entry['symbol']]), str(numbers['price']))
                change = get_change(numbers['price'], priceMap[entry['symbol']])
                worksheet.write(chr(ord(lastChar) + 2) + str(entryMap[entry['symbol']]), change, getCellFormat(change, goodStyle, badStyle))
        lastChar = chr(ord(lastChar) + 2)
        file.close()

workbook.close()  


