from bs4 import BeautifulSoup
from selenium import webdriver
from time import sleep
from xlsxwriter import Workbook
import xlrd
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
import os


class Run:
	def __init__(self,url='http://csrankings.org/#/index?',path='F:/Downloads/chromedriver.exe'):
		area = ['ai&vision&mlmining&nlp&ir', 'arch&comm&sec&mod&hpc&mobile&metrics&ops&plan&soft&da&bed', 'act&crypt&log', 'graph&chi&robotics&bio&visualization&ecom']
		area_name = ['AI', 'Systems', 'Theory', 'Interdisciplinary areas']
		"""country = ['&au', '&at', '&br', '&ca', '&cn', '&dk', '&fr', '&de', '&gr', '&hk', '&in', '&il',
					 '&it', '&jp', '&nl', '&nz' '&kr', '&es', '&ch', '&tr', '&uk', '&us']
		country_name = ['Australia', 'Austria', 'Brazil', 'Canada', 'China', 'Denmark', 'France', 'Germany', 'Greece', 'Hong Kong',
						 'India', 'Israel', 'Italy', 'Japan', 'Netherlands', 'New Zealand', 'South Korea', 'Spain', 'Switzerland', 'Turkey', 'United Kingdom',
						 'USA', 'the world']"""
		country = ['&world']
		country_name = ['the world']		
		for i in range(len(country)):
			for j in range(len(area)):
				self.url  = url+area[j]+country[i]
				self.path = path
				self.array=[]
				self.driver = webdriver.Chrome(self.path)
				self.driver.get(self.url)
				if not os.path.exists('faculty/'+country_name[i]):
					os.mkdir('faculty/'+country_name[i])
				self.workbook=Workbook('faculty/'+country_name[i]+'/'+area_name[j]+'.xlsx')

				try:
					drop=Select(self.driver.find_element_by_xpath('//*[@id="regions"]'))
					drop.select_by_visible_text(country_name[i])
					self.extract()
				except Exception as e:
					print(e)

				self.driver.close()
				self.workbook.close()


	def convert(self,data):
		item = ""
		data = data.strip()
		data = data.replace('&','%26')
		data = data.split(' ')
		for i in range(len(data)-1):
			val = str(data[i])

			item = item + val + "%20"
		item = item + data[len(data)-1]+"-widget"
		return item

	def database(self,data,row_id):
		try :
			print("Start database")
			tbody = data.find('tbody')

			prev = tbody.find('tr')
			res = prev.find_all('td')
			self.worksheet.write(self.count,0,res[1].find('a').string)
			self.worksheet.write(self.count,1,res[1].find('a')['href'])
			self.worksheet.write(self.count,2,res[2].find('a').string)
			self.worksheet.write(self.count,3,res[3].find('small').string)
			try:
				self.worksheet.write(self.count,4,res[1].find('a',attrs={"title": "Click for author's Google Scholar page."})['href'])
			except Exception as e:
				print("Not Found Google Scholar page")
			try:
				self.worksheet.write(self.count,5,res[1].find('a',attrs={"title": "Click for author's DBLP entry."})['href'])
			except Exception as e:
				print("Not Found DBLP page")
			self.count = self.count + 1
			
	
			for i in range(1,100):
				try:
					mid1 = prev.find_next_sibling('tr')
					nex = mid1.find_next_sibling('tr')
					res = nex.find_all('td')
					self.worksheet.write(self.count,0,res[1].find('a').string)
					self.worksheet.write(self.count,1,res[1].find('a')['href'])
					self.worksheet.write(self.count,2,res[2].find('a').string)
					self.worksheet.write(self.count,3,res[3].find('small').string)
					try:
						self.worksheet.write(self.count,4,res[1].find('a',attrs={"title": "Click for author's Google Scholar page."})['href'])
					except Exception as e:
						print("Not Found Google Scholar page")
					try:
						self.worksheet.write(self.count,5,res[1].find('a',attrs={"title": "Click for author's DBLP entry."})['href'])
					except Exception as e:
						print("Not Found DBLP page")
					self.count = self.count + 1

					prev=nex
				except Exception as e:
					print("database")
					print(i)
					print(e)
					break

		except Exception as e:
			print(e)




	def extract(self):
		try :
			print("Extract")
			sleep(10)
			self.count = 1
			soup = BeautifulSoup(self.driver.page_source,'html.parser')
			
			table = soup.find('table',{'id':'ranking'})
			tbody = table.find('tbody')

			prev = tbody.find('tr')
			res = prev.find_all('td')

			self.worksheet=self.workbook.add_worksheet(""+str(res[1].contents[2].string.replace(" ", ""))[0:28])
			self.worksheet.write(0,0,"Name")
			self.worksheet.write(0,1,"Link")
			self.worksheet.write(0,2,"Pubs")
			self.worksheet.write(0,3,"Adj")
			self.worksheet.write(0,4,"Google Scholar Link")
			self.worksheet.write(0,5,"DBLP Link")
			row_id = self.convert(str(res[1].contents[2].string))


			strnew='//*[@id="'+row_id+'"]'
			try:
				element=self.driver.find_element_by_xpath(strnew)
				element.click()
			except Exception as e:
				print("Link not found")

			cont_mid1 = prev.find_next_sibling('tr')
			cont_mid2 = cont_mid1.find_next_sibling('tr')

			self.database(cont_mid2,row_id)

			element.click()

			for i in range(1,200):
				try:
					self.count = 1
					#mid1 = prev.find_next_sibling('tr')
					#mid2 = mid1.find_next_sibling('tr')
					#nex = mid2.find_next_sibling('tr')
					nex = prev.find_next_sibling('tr')
					while len(nex) == 1:
						nex = nex.find_next_sibling('tr')
					#print(len(mid1), len(mid2), len(nex))
					res = nex.find_all('td')
					
					self.worksheet=self.workbook.add_worksheet(""+str(res[1].contents[2].string.replace(" ", ""))[0:28] )
					self.worksheet.write(0,0,"Name")
					self.worksheet.write(0,1,"Link")
					self.worksheet.write(0,2,"Pubs")
					self.worksheet.write(0,3,"Adj")
					self.worksheet.write(0,4,"Google Scholar Link")
					self.worksheet.write(0,5,"DBLP Link")
					row_id = self.convert(str(res[1].contents[2].string))

					try:
						element=self.driver.find_element_by_xpath('//*[@id="'+row_id+'"]')
						element.click()
					except Exception as e:
						print(e)

					cont_mid1 = nex.find_next_sibling('tr')
					cont_mid2 = cont_mid1.find_next_sibling('tr')
					self.database(cont_mid2,row_id)
					element.click()
					prev=nex

				except Exception as e:
					print('extract')
					print(i)
					print(e)
					break

		except Exception as e:
			print(e)



if __name__ =='__main__' :
	run=Run()
	print("Done")
