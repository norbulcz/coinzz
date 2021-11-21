from time import time
import xlsxwriter
import json
from datetime import date, datetime, timedelta

file = open('files/data-'+datetime.now().strftime('%d-%m-%Y')+'.json',)
 
data = json.load(file)
today = datetime.today()
daysAgo7 = today - timedelta(days=7)
print(daysAgo7)
exportFileName = 'Export-' + datetime.now().strftime('%d-%m-%Y')
workbook = xlsxwriter.Workbook('files/'+exportFileName+'.xlsx')
worksheet = workbook.add_worksheet()
worksheet.set_column(0,0,25)
worksheet.set_column(1,1,10)
worksheet.set_column(2,7,25)
header_format = workbook.add_format({'bold': True})
worksheet.write('A1', 'Name', header_format)
worksheet.write('B1', 'Symbol', header_format)
worksheet.write('C1', 'Network', header_format)
worksheet.write('D1', '% 1h', header_format)
worksheet.write('E1', '% 24h', header_format)
worksheet.write('F1', '% 7d', header_format)
worksheet.write('G1', 'Volume 24h', header_format)
worksheet.write('H1', 'Volume change 24h', header_format)
worksheet.write('I1', 'Added', header_format)
entryNumber = 2
for entry in data['data']:
    dateAdded=datetime.strptime(entry["date_added"].split('T')[0], '%Y-%m-%d')
    if daysAgo7 < dateAdded:
        print(entry['symbol'])
        print('A'+str(entryNumber))
        worksheet.write('A'+str(entryNumber), entry["name"])
        worksheet.write('B'+str(entryNumber), entry["symbol"])
        if entry['platform'] is not None:
            worksheet.write('C'+str(entryNumber), entry["platform"]['name'])
        numbers = entry['quote']['USD']
        worksheet.write('D'+str(entryNumber), str(numbers['percent_change_1h']))
        worksheet.write('E'+str(entryNumber), str(numbers['percent_change_24h']))
        worksheet.write('F'+str(entryNumber), str(numbers['percent_change_7d']))
        worksheet.write('G'+str(entryNumber), str(numbers['volume_24h']))
        worksheet.write('H'+str(entryNumber), str(numbers['volume_change_24h']))
        worksheet.write('I'+str(entryNumber), entry['date_added'])
        entryNumber += 1

workbook.close()       
file.close()
