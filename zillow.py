
import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QCoreApplication
from PyQt5 import QtCore
from PyQt5.QtGui import QIcon
import os
from PyQt5 import QtWidgets
from datetime import datetime
from lxml import html
from bs4 import BeautifulSoup
import xlsxwriter
import xlrd
import re
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import csv 
import requests

from PyQt5 import QtCore

from PyQt5.QtCore import QThread

try:
    _fromUtf8 = QtCore.QString.fromUtf8
except AttributeError:
    def _fromUtf8(s):
        return s
from PyQt5.QtCore import QObject, pyqtSignal


urls = []
data = []

from excel_writer import write_to_excel


class getPostsThread(QThread):
	add_post = pyqtSignal()
	finish_task = pyqtSignal()

	def __init__(self):
		QThread.__init__(self)

	def __del__(self):
		self.wait()

	def _get_top_post(self):
		driver = webdriver.Chrome()
		for url in urls:
			driver.get(url["full_url"])
			sup = BeautifulSoup(driver.page_source,"lxml")
			output = {}
			output["Str #"] = url["str_num"]
			output["Street Name"] = url["str_name"]
			output["City"] = url["city"]
			output["Zip"] =  url["zip"]
			try:
				try:
					link = sup.find("li",attrs={"id":"hdp-popout-menu"}).find("a")["href"]
					f_link = "http://www.zillow.com/"+link
					driver.get(f_link)
					# driver.get("http://www.zillow.com/homedetails/14806-Bridge-Creek-Dr-Midlothian-VA-23113/88874836_zpid/")
					driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
					time.sleep(10)
					su = BeautifulSoup(driver.page_source,"lxml")
					output["Link to Zillow"] = f_link
			
					bed = su.find("header", class_="zsg-content-header addr").find("h3")
					try:
						if bed is not None:
							b = str(bed).split("</span>")
							beds = b[1].replace('<span class="addr_bbs">',"")
							baths = b[3].replace('<span class="addr_bbs">',"")
							sqft = b[5].replace('<span class="addr_bbs">',"")
							output["BR"] = beds
							output["BA"] = baths
							sq = str(sqft).replace("sqft","").strip()
							output["Sq ft"] = sq
					except:
						pass

					try:
						rent_zest = su.find_all("div", class_="zest zsg-lg-1-3 zsg-md-1-1")[1].find("div",class_="zest-value").text
						if "mo" not in str(rent_zest):
							rent_zest = su.find_all("div", class_="zest zsg-lg-1-3 zsg-md-1-1")[0].find("div",class_="zest-value").text
						output["Rent Zestimate"] = str(rent_zest).replace("$","").strip()

						
					except:
						try:
							rent_zest = su.find("div", class_="zsg-g zest-double").find_all("div",class_="zest-value")[1].text
							if "mo" not in str(rent_zest):
								rent_zest = su.find("div", class_="zsg-g zest-double").find_all("div",class_="zest-value")[0].text
							output["Rent Zestimate"] = str(rent_zest).replace("$","").strip()
							
						except:
							pass

						
					try:
						zest = su.find("div", class_="zsg-g zest-double").find_all("div",class_="zest-value")[0].text
						if "mo" in str(zest):
							zest = su.find("div", class_="zsg-g zest-double").find_all("div",class_="zest-value")[1].text
						if "mo" not in str(zest):
							output["Zestimate"] = str(zest).replace("$","").replace("Zestimate\xae:","").strip()
					except:
						try:
							zest = su.find_all("div", class_="zest zsg-lg-1-3 zsg-md-1-1")[0].find("div",class_="zest-value").text
							if "mo" in str(zest):
								zest = su.find_all("div", class_="zest zsg-lg-1-3 zsg-md-1-1")[1].find("div",class_="zest-value").text
							output["Zestimate"] = str(zest).replace("$","").replace("Zestimate\xae:","").strip()
						except:
							pass


					try:
						price_type = su.find("div", class_="estimates").find_all("div", class_="home-summary-row")[0].text
						if "Rent" in price_type:
							price = su.find("div", class_="estimates").find_all("div", class_="home-summary-row")[1].text
							output["Rent"] = price.strip().replace("$","").strip()
						elif "Sale" in price_type:
							price = su.find("div", class_="estimates").find_all("div", class_="home-summary-row")[1].text
							print (price)
							if "Zestimate" in price:
								price = su.find("div", class_="estimates").find_all("div", class_="home-summary-row")[0].text
							output["Sales Price"] = price.strip().replace("$","").strip()
						elif "Sold" in price_type:
							output["Sold Amt"] = price_type.strip().replace("Sold:","").replace("$","")
						else:
							price = su.find("div", class_="estimates").find_all("div", class_="home-summary-row")[1].text
							print (price) 
							if "Zestimate" in price:
								price = su.find("div", class_="estimates").find_all("div", class_="home-summary-row")[0].text
							output["Sales Price"] = price.strip().replace("$","").strip()


						if "Sold" in price_type:
							typ = price_type.split(":")
							output["Listing Type"] = typ[0].strip()
						else:
							output["Listing Type"] = price_type.strip()
						if "LOT" in price_type:
							output["Home type"] =  price_type.strip()

					except Exception as e:
						print (e)
						pass
					try:
						nearby_school = su.find_all("div", class_="nearby-schools-name")
						for n in nearby_school:
							school = str(n.text).strip()
							if "Elementary" in school:
								output["Nearby Elementary"] = school.replace("Elementary","")
							elif "Middle" in school:
								output["Nearby Middle"] = school
							elif "High" in school:
								output["Nearby High"] = school
					except:
						pass
					try:
						fact = su.find("div", class_="hdp-facts").find_all("li")
						type_list = ["Cooperative", "Condo", "Townhome", "Lot", "Land", "Apartment", "Apt", "Multi Family", "Townhouse"]

						for f in fact:
							if "Single Family" in str(f.text) or "Condo" in str(f.text) or "Cooperative" in str(f.text) or "Townhouse" in str(f.text) or "Townhome" in str(f.text) or "Apt" in str(f.text) or "Apartment" in str(f.text) or "Multi Family" in str(f.text):
								home_type = str(f.text)
								output["Home type"] = home_type.replace("PropertyType:","").replace("Property Type:","").strip()
								print ("Home type === %s"%home_type)
							elif "acres" in str(f.text):
								Lot = str(f.text).replace("Lot:","").replace("Lot Description:","")
								output["Lot Size"] = Lot.strip()
								print ("Lot === %s"%Lot)
							elif "MLS" in str(f.text):
								mils = str(f.text).replace("MLS #:","")
								print ("Mils === %s"%mils)
								if "{" not in mils:
									output["MLS"] = mils.strip()
							elif "days on" in str(f.text):
								zillow_time = str(f.text)
								output["Days on Zillow"] = int(zillow_time.replace("days on Zillow","").strip())
								print ("zillow time === %s"%zillow_time)
							elif "shopper" in str(f.text):
								shoppers = str(f.text)
								print ("shoppers === %s"%shoppers)
								output["Shopper Saves"] = shoppers.strip().replace("shoppers saved this home","").replace("shopper saved this home","").strip()
							elif "view" in str(f.text):
								view = str(f.text)
								if "{" not in view:
									output["Views"] = view.replace("All time views:","").strip()
									print ("view === %s"%view)
							elif "HOA" in str(f.text):
								hoa = str(f.text)
								output["HOA"] = hoa.strip().replace("HOA Fee:","").replace("HOAFee:","").strip()
					except:
						pass
					try:
						price_history = su.find("div", attrs={"id":"hdp-price-history"}).find_all("tr")
						loop = 0
						for t in price_history:
							loop = loop +1
							
							if loop>1:
								if loop == 2:
									price_info = t.find_all("td")
									date = str(price_info[0].text).strip()
									event = str(price_info[1].text).strip()
									price = str(price_info[2].find("span").text).strip()
									source = str(price_info[4].text).strip()
									output["Date_1"] = date
									output["Event_1"] = event
									output["Price_1"] = price.replace("$","").strip()
									output["Source_1"] = source
					
								elif loop == 3:
									price_info = t.find_all("td")
									date = str(price_info[0].text).strip()
									event = str(price_info[1].text).strip()
									price = str(price_info[2].find("span").text).strip()
									source = str(price_info[4].text).strip()
									output["Date_2"] = date
									output["Event_2"] = event
									output["Price_2"] = price.replace("$","").strip()
									output["Source_2"] = source
								elif loop == 4:
									price_info = t.find_all("td")
									date = str(price_info[0].text).strip()
									event = str(price_info[1].text).strip()
									price = str(price_info[2].find("span").text).strip()
									source = str(price_info[4].text).strip()
									output["Date_3"] = date
									output["Event_3"] = event
									output["Price_3"] = price.replace("$","").strip()
									output["Source_3"] = source

					except:
						pass
					try:
						tax = su.find("div", attrs={"id":"hdp-tax-history"}).find_all("tr")
						loop = 0
						tax_hist = []
						succ = 0
						for t in tax:
							loop = loop +1
							if loop>1:
								try:
									if loop == 2 :
										tax_info = t.find_all("td")
										year = str(tax_info[0].text)
										property_taxes = str(tax_info[1]).strip()
										prop = property_taxes.split("<span")
										pro = prop[0].replace("""<td class="numeric">""","").replace("$","")
										property_assest = str(tax_info[3].text)
										output["Tax Year_1"] = year
										output["Property Taxes_1"] = pro
										output["Tax Assessment_1"] = property_assest.replace("$","").strip()
										succ = 2
								except Exception as e:
									print (e)
									pass
								if succ == 2:
									if loop == 3:
										tax_info = t.find_all("td")
										year = str(tax_info[0].text)
										property_taxes = str(tax_info[1]).strip()
										prop = property_taxes.split("<span")
										pro = prop[0].replace("""<td class="numeric">""","").replace("$","")
										property_assest = str(tax_info[3].text)
										output["Property Taxes_2"] = pro
										output["Tax Year_2"] = year
										output["Tax Assessment_2"] = property_assest.replace("$","").strip()
								else:
									if loop == 3 :
										tax_info = t.find_all("td")
										year = str(tax_info[0].text)
										property_taxes = str(tax_info[1]).strip()
										prop = property_taxes.split("<span")
										pro = prop[0].replace("""<td class="numeric">""","").replace("$","")
										property_assest = str(tax_info[3].text)


										property_assest = str(tax_info[3].text)
										output["Tax Year_1"] = year
										output["Property Taxes_1"] = pro
										output["Tax Assessment_1"] = property_assest.replace("$","").strip()
										
									elif loop == 4:
										tax_info = t.find_all("td")
										year = str(tax_info[0].text)
										
										property_assest = str(tax_info[3].text)
										property_taxes = str(tax_info[1]).strip()
										prop = property_taxes.split("<span")
										pro = prop[0].replace("""<td class="numeric">""","").replace("$","")
										property_assest = str(tax_info[3].text)
										output["Property Taxes_2"] = pro
										output["Tax Year_2"] = year
										output["Tax Assessment_2"] = property_assest.replace("$","").strip()
					except Exception as e:
						print (e)
						pass
					data.append(output)
					print (output)
				except Exception as e:
					print (e)
					pass
			except Exception as e:
				print (e)
			self.add_post.emit()


		uniqe_list = []
		new_list = [x["Link to Zillow"] for x in data if "Link to Zillow" in x]
		uniqe = list(set(new_list))
		for u in uniqe:
			z = 0
			for d in data:
					if "Link to Zillow" in d:
						if u == d["Link to Zillow"]:
							if z == 0:
								z = z+1
								uniqe_list.append(d)
							else:
								continue
		dat = time.strftime('%Y-%m-%d %H:%M')
		print ("uniqe_list", uniqe_list)
		workbook = xlsxwriter.Workbook('Zillow_Data-{}.xlsx'.format(dat))
		worksheet = workbook.add_worksheet('Zillow Data')
		write_to_excel(workbook, worksheet, uniqe_list)
		with open('Zillow_Data-{}.csv'.format(dat), 'w') as f:
			dict_writer = csv.DictWriter(f, uniqe_list[0].keys())
			dict_writer.writeheader()
			dict_writer.writerows(uniqe_list)
		workbook.close()
		driver.close()
		driver.quit()

		self.finish_task.emit()

	def run(self):
		self._get_top_post()
           



class Example(QWidget):

	def __init__(self):
		super().__init__()
		self.setGeometry(50, 50, 500, 500)
		self.setWindowTitle("Log In to your facebook account")
		self.setStyleSheet("background-color: #3b5998;")
		self.show()

		self.initUI()

	def initUI(self):
		LineEditStyle = "background-color: white;padding: 10px 10px 10px 20px; font-size: 14px; font-family: consolas;" \
                        "border: 2px solid #3BBCE3; border-radius: 4px; width:100px;"
		logo = QLabel("Zillow Scraper")
		logo.setStyleSheet("font-size: 50px; font-weight:bold;color:white;")
		logo.setAlignment(QtCore.Qt.AlignCenter | QtCore.Qt.AlignVCenter)

		email = QLineEdit()
		email.setPlaceholderText("Enter Your Search Param")
		email.setStyleSheet(LineEditStyle)

		btn = QtWidgets.QPushButton('Select a XLSX', self)
		btn.setStyleSheet(LineEditStyle)
		btn.resize(btn.sizeHint())
		btn.clicked.connect(self.browse_file)

		self.progress_bar = QtWidgets.QProgressBar(self)
		self.progress_bar.setProperty("value", 0)
		self.progress_bar.setObjectName(_fromUtf8("progress_bar"))
		self.progress_bar.setStyleSheet(LineEditStyle)
		

		self.get_thread = getPostsThread()
		self.get_thread.add_post.connect(self.add_post)
		self.get_thread.finish_task.connect(self.done)
	

		self.s_btn = QtWidgets.QPushButton('Start Crawler', self)
		self.s_btn.setStyleSheet(LineEditStyle)
		self.s_btn.resize(btn.sizeHint())
		self.s_btn.clicked.connect(self.start_crawling)
		self.s_btn.setEnabled(True)


		self.st_btn = QtWidgets.QPushButton('Stop', self)
		self.st_btn.setStyleSheet(LineEditStyle)
		self.st_btn.resize(btn.sizeHint())
		self.st_btn.setEnabled(False)
		self.st_btn.clicked.connect(self.get_thread.terminate)


		self.textEdit = QLineEdit()
		
		layout = QVBoxLayout()

		self.setLayout(layout)
		layout.addWidget(logo)
		# layout.addStretch()
		# layout.addWidget(email)
		layout.addWidget(btn)
		layout.addWidget(self.textEdit)
		layout.addWidget(self.progress_bar)
		layout.addStretch(1)
		layout.addWidget(self.s_btn)
		layout.addWidget(self.st_btn)
		layout.addStretch(1)



	def browse_file(self):
		fname = QFileDialog.getOpenFileName(self, 
		                  "Open file", os.path.expanduser('~'))
		if fname[0] and "xlsx" in fname[0]:
			self.textEdit.setText(fname[0])
			book = xlrd.open_workbook(fname[0])
			sheet = book.sheet_by_index(0)
			for i in range(1, sheet.nrows):
				url = {}
				o = sheet.row_values(i)
				str_num = o[0]
				try:
					str_num = int(o[0])
				except:
					str_num = o[0]
				full_add = str(str_num) +"-"+o[1] + "-" + o[2] + "-" + str(int(o[3]))
				full_url = "http://www.zillow.com/homes/" + "-".join(full_add.split(" ")) + "_rb"
				url["str_num"] = str_num
				url["str_name"] = o[1]
				url["city"] = o[2]
				url["zip"] = int(o[3])
				url["full_url"] = full_url
				urls.append(url)
			print(urls)

	def add_post(self):
		print(" Called Signal")
		self.progress_bar.setValue(self.progress_bar.value()+1)

	def start_crawling(self):
		self.progress_bar.setMaximum(len(urls))
		self.st_btn.setEnabled(True)
		self.get_thread.start()

	def done(self):
		self.st_btn.setEnabled(False)
		self.s_btn.setEnabled(True)
		self.progress_bar.setValue(0)
		QtWidgets.QMessageBox.information(self, "Done!", "Done fetching data!")
		

if __name__ == '__main__':

	app = QApplication(sys.argv)
	ex = Example()
	sys.exit(app.exec_())
