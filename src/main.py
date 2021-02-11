# import xlsxwriter module 
import xlsxwriter 
from lxml.html import parse
    
workbook = xlsxwriter.Workbook('output.xlsx') 
worksheet = workbook.add_worksheet() 
worksheet.write('A1', 'Link Text')
worksheet.write('B1', 'Link url')
worksheet.write('C1', 'TC name')
doc = parse('http://www.google.com').getroot()

for i, div in enumerate(doc.cssselect('a')):
# 	 print (div.text_content(), div.get('href'))
	 worksheet.write('A'+str(i+2), div.text_content())
	 worksheet.write('B'+str(i+2), div.get('href'))
	 worksheet.write('C'+str(i+2), 'TC'+str(i+1)+'_'+div.text_content().lower()+'_click')
	 if i>15:
	 	break

workbook.close() 
