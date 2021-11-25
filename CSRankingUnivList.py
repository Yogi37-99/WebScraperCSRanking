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
					 '&it', '&jp', '&nl', '&nz' '&kr', '&es', '&ch', '&tr', '&uk', '&us', '&world']
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
				if not os.path.exists('ranking/'+country_name[i]):
					os.mkdir('ranking/'+country_name[i])
				sleep(2)
				self.workbook=Workbook('ranking/'+country_name[i]+'/'+area_name[j]+'.xlsx')
				self.worksheet=self.workbook.add_worksheet("Ranking")		
				try:
					drop=Select(self.driver.find_element_by_id('regions'))
					drop.select_by_visible_text(country_name[i])
					self.extract()
				except Exception as e:
					print(e)

				self.driver.close()
				self.workbook.close()
			

	def extract(self):
		try :
			print("Extract")

			
			print("Done")
			sleep(10)


			self.worksheet.write(0,0,"University")
			self.worksheet.write(0,1,"Count")
			self.worksheet.write(0,2,"Faculty")
			soup = BeautifulSoup(self.driver.page_source,'html.parser')
			
			table = soup.find('table',{'id':'ranking'})
			tbody = table.find('tbody')

			prev = tbody.find('tr')
			res = prev.find_all('td')

			self.worksheet.write(1,0,res[1].contents[2].string)
			self.worksheet.write(1,1,res[2].string)
			self.worksheet.write(1,2,res[3].string)
	
			for i in range(1,187):
				try:
					mid1 = prev.find_next_sibling('tr')
					mid2 = mid1.find_next_sibling('tr')
					nex = mid2.find_next_sibling('tr')
					res = nex.find_all('td')
					self.worksheet.write(i+1,0,res[1].contents[2].string)
					self.worksheet.write(i+1,1,res[2].string)
					self.worksheet.write(i+1,2,res[3].string)
					prev=nex
				except Exception as e:
					print(i)
					

		except Exception as e:
			print(e)



if __name__ =='__main__' :
	run=Run()
	print("Done")