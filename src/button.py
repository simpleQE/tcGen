import xlsxwriter 
import urllib.request as urllib2
from urllib.parse import urlparse
from bs4 import BeautifulSoup

workbook = xlsxwriter.Workbook('output.xlsx') 

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

quote_page = 'https://www.flipkart.com/'
home = urlparse(quote_page).netloc     # toDo  extract proper home name
page = urllib2.urlopen(quote_page)
soup = BeautifulSoup(page, 'lxml')
buttons_list = soup.find_all('button')
for i, div in enumerate(buttons_list):
# 	import ipdb;ipdb.set_trace()
	button_text = " ".join(str(div.text).split())
	worksheet.write('A'+str(i+2), 'UC'+str(i+1)+'_'+button_text.lower()+'_button_click')
	worksheet.write('B'+str(i+2), 'TC'+str(i+1)+'_'+button_text.lower()+'_button_click')
	worksheet.write('C'+str(i+2), button_text)
	worksheet.write('D'+str(i+2), 'Validating '+button_text+' button')
	worksheet.write('E'+str(i+2), '['+home+']['+button_text+']')
	worksheet.write('F'+str(i+2), 'Objective: To Validate clicking '+button_text+' button. \nPre-requisite - User should have desired access to the '+home+' . \nTest steps: \n1. Go to '+home+' .\n2. Click on '+button_text+' button.')
	button_type = div.get('type')
	button_onclick = div.get('onclick')
	if button_onclick is not None:
		worksheet.write('G'+str(i+2), '1. '+button_text+' button click should activate respective onClick function.')
	elif button_type is not None:
		if button_type.lower()=='submit':
			worksheet.write('G'+str(i+2), '1. '+button_text+' button click should activate submit action for respective input field.')
		elif button_type.lower()=='reset':
			worksheet.write('G'+str(i+2), '1. '+button_text+' button click should reset all input fields to default.')
		elif button_type.lower()=='button':
			worksheet.write('G'+str(i+2), '1. '+button_text+' button should get clicked.')
		worksheet.write('H'+str(i+2), 'Smoke')
	else:
		worksheet.write('G'+str(i+2), '1. '+button_text+' button click should do nothing.')
	if i>100000:
		break
workbook.close() 
