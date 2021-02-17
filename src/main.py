import xlsxwriter 
import urllib.request as urllib2
from bs4 import BeautifulSoup

workbook = xlsxwriter.Workbook('output.xlsx') 
worksheet = workbook.add_worksheet() 
worksheet.write('A1', 'Link Text')
worksheet.write('B1', 'Link url')
worksheet.write('C1', 'TC name')

quote_page = 'https://flipkart.com/'
page = urllib2.urlopen(quote_page)
soup = BeautifulSoup(page, 'lxml')

for i, div in enumerate(soup.find_all('a', href=True)):
	link_text = " ".join(str(div.text).split())
	link_url = div.get('href')
	if link_url is not None:
		worksheet.write('A'+str(i+2), link_text)
		if link_url.startswith("/"):
			worksheet.write('B'+str(i+2), quote_page+link_url)
		else:
			worksheet.write('B'+str(i+2), link_url)
		worksheet.write('C'+str(i+2), 'TC'+str(i+1)+'_'+link_text.lower()+'_click')
	if i>100000:
		break

workbook.close() 
