# import xlsxwriter module 
import xlsxwriter 
import urllib.request as urllib2
from bs4 import BeautifulSoup

workbook = xlsxwriter.Workbook('output.xlsx') 
worksheet = workbook.add_worksheet() 
worksheet.write('A1', 'Link Text')
worksheet.write('B1', 'Link url')
worksheet.write('C1', 'TC name')

quote_page = 'https://github.com/'
page = urllib2.urlopen(quote_page)
soup = BeautifulSoup(page, 'lxml')

for i, div in enumerate(soup.find_all('a')):
	link_text = " ".join(str(div.text).split())
	if link_text:
		print(link_text)
		worksheet.write('A'+str(i+2), link_text)
		worksheet.write('B'+str(i+2), div.get('href'))
		worksheet.write('C'+str(i+2), 'TC'+str(i+1)+'_'+link_text.lower()+'_click')
	if i>15:
		break

workbook.close() 
