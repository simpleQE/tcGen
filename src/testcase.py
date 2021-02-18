import xlsxwriter 
import urllib.request as urllib2
from urllib.parse import urlparse
from bs4 import BeautifulSoup
import click


class ManualTestCases(object):
	"""Class to generate test cases and its structure"""
	def __init__(self):
		self.workbook = xlsxwriter.Workbook('/tmp/output.xlsx')

		self.cell_format = self.workbook.add_format()
		self.cell_format.set_bg_color('green') # not working.. todo
		self.worksheet = self.workbook.add_worksheet()
		self.worksheet.write('A1', 'Use Case Name')
		self.worksheet.write('B1', 'Test Case Name')
		self.worksheet.write('C1', 'Scenario')
		self.worksheet.write('D1', 'Use Case')
		self.worksheet.write('E1', 'Test Case Title')
		self.worksheet.write('F1', 'Test Case Description')
		self.worksheet.write('G1', 'Expected Results')
		self.worksheet.write('H1', 'Type of Test Case')
		self.worksheet.write('I1', 'Status')
		self.worksheet.write('J1', 'Comments')

	def test_case_generator(self, url):
		"""generates tests"""
		quote_page = url
		home = urlparse(quote_page).netloc
		page = urllib2.urlopen(quote_page)
		soup = BeautifulSoup(page, 'lxml')

		# This is for anchors tag
		anchors_list = soup.find_all('a')
		# import ipdb;ipdb.set_trace()
		for i, div in enumerate(anchors_list):
			link_text = " ".join(str(div.text).split())
			link_url = div.get('href')
			if link_url is None:
				if div.img:
					link_url = div.img['src']
			self.worksheet.write('A'+str(i+2), 'UC'+str(i+1)+'_'+link_text.lower()+'_click')
			self.worksheet.write('B'+str(i+2), 'TC'+str(i+1)+'_'+link_text.lower()+'_click')
			self.worksheet.write('C'+str(i+2), link_text)
			self.worksheet.write('D'+str(i+2), 'Validating '+link_text+' link')
			self.worksheet.write('E'+str(i+2), '['+home+']['+link_text+']')
			self.worksheet.write('F'+str(i+2), 'Objective: To Validate opening of '+link_text+' link. \nPre-requisite - User should have desired access to the '+home+' . \nTest steps: \n1. Go to '+home+' .\n2. Click on '+link_text+' link.')
			self.worksheet.write('G'+str(i+2), '1. '+link_text+' link should open.')
			self.worksheet.write('H'+str(i+2), 'Functional')
			if link_url.startswith("/"):
				self.worksheet.write('K'+str(i+2), quote_page+link_url)
			else:
				self.worksheet.write('K'+str(i+2), link_url)
			k=i
			if i > 100000:
				break
		# import ipdb;ipdb.set_trace()
		# This is for button tag
		buttons_list = soup.find_all('button')
		for i, div in enumerate(buttons_list):
			i=i+k
			button_text = " ".join(str(div.text).split())
			self.worksheet.write('A'+str(i+2), 'UC'+str(i+1)+'_'+button_text.lower()+'_button_click')
			self.worksheet.write('B'+str(i+2), 'TC'+str(i+1)+'_'+button_text.lower()+'_button_click')
			self.worksheet.write('C'+str(i+2), button_text)
			self.worksheet.write('D'+str(i+2), 'Validating '+button_text+' button')
			self.worksheet.write('E'+str(i+2), '['+home+']['+button_text+']')
			self.worksheet.write('F'+str(i+2), 'Objective: To Validate clicking '+button_text+' button. \nPre-requisite - User should have desired access to the '+home+' . \nTest steps: \n1. Go to '+home+' .\n2. Click on '+button_text+' button.')
			button_type = div.get('type')
			button_onclick = div.get('onclick')
			if button_onclick is not None:
				self.worksheet.write('G'+str(i+2), '1. '+button_text+' button click should activate respective onClick function.')
			elif button_type is not None:
				if button_type.lower() == 'submit':
					self.worksheet.write('G'+str(i+2), '1. '+button_text+' button click should activate submit action for respective input field.')
				elif button_type.lower() == 'reset':
					self.worksheet.write('G'+str(i+2), '1. '+button_text+' button click should reset all input fields to default.')
				elif button_type.lower() == 'button':
					self.worksheet.write('G'+str(i+2), '1. '+button_text+' button should get clicked.')
				self.worksheet.write('H'+str(i+2), 'Smoke')
			else:
				self.worksheet.write('G'+str(i+2), '1. '+button_text+' button click should do nothing.')
			if i > 100000:
				break

		self.workbook.close()


@click.command(help="Provide url")
@click.option("-u", "--url", default=None, help="URL to generate test cases")
def generate(url):
	"""tests"""
	tests = ManualTestCases()
	tests.test_case_generator(url)
