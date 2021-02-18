import xlsxwriter 
import urllib.request as urllib2
from urllib.parse import urlparse
from bs4 import BeautifulSoup

workbook = xlsxwriter.Workbook('output.xlsx') 

cell_format = workbook.add_format()
cell_format.set_bg_color('green') # not working.. todo
worksheet = workbook.add_worksheet() 
worksheet.write('A1', 'Use Case Name')
worksheet.write('B1', 'Test Case Name')
worksheet.write('C1', 'Scenario')
worksheet.write('D1', 'Use Case')
worksheet.write('E1', 'Test Case Title')
worksheet.write('F1', 'Test Case Description')
worksheet.write('G1', 'Expected Results')
worksheet.write('H1', 'Type of Test Case')
worksheet.write('I1', 'Status')
worksheet.write('J1', 'Comments')

quote_page = 'https://flipkart.com/'
home = urlparse(quote_page).netloc
page = urllib2.urlopen(quote_page)
soup = BeautifulSoup(page, 'lxml')
anchors_list = soup.find_all('a', href=True)
for i, div in enumerate(anchors_list):
	link_text = " ".join(str(div.text).split())
	link_url = div.get('href')
	if link_url is not None:
		worksheet.write('A'+str(i+2), 'UC'+str(i+1)+'_'+link_text.lower()+'_click')
		worksheet.write('B'+str(i+2), 'TC'+str(i+1)+'_'+link_text.lower()+'_click')
		worksheet.write('C'+str(i+2), link_text)
		worksheet.write('D'+str(i+2), 'Validating '+link_text+' link')
		worksheet.write('E'+str(i+2), '['+home+']['+link_text+']')
		worksheet.write('F'+str(i+2), 'Objective: To Validate opening of '+link_text+' link. \nPre-requisite - User should have desired access to the '+home+' . \nTest steps: \n1. Go to '+home+' .\n2. Click on '+link_text+' link.')
		worksheet.write('G'+str(i+2), '1. '+link_text+' link should open.')
		worksheet.write('H'+str(i+2), 'Functional')
	if i>100000:
		break

'''if link_url.startswith("/"):
	worksheet.write('B'+str(i+2), quote_page+link_url)
else:
	worksheet.write('B'+str(i+2), link_url)'''
# for j in range(i):
# 	worksheet.set_column(i,i,30)
workbook.close() 
