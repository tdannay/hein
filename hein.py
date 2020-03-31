### Tim Dannay - 4/29/15 ###
### Script for automating the downloads of articles from HeinOnline ###
### See heinlist.csv in the same directory for the list of title/author search terms to be used by this script ###

### Imports modules for browser control, operating system keypress recognition, sleep function, and csv reader ###

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import win32com.client
import time
import csv

### Creates Custom Firefox profile to reduce impact of popups and to set the download directory ###

fp = webdriver.FirefoxProfile()
fp.set_preference("browser.download.folderList", 2)
fp.set_preference("browser.download.manager.showWhenStarting", False)
fp.set_preference("browser.download.dir", "C:\Users\Tdannay\Dropbox\Saves\HeinPDFs")
fp.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/pdf")
fp.set_preference("browser.download.manager.showAlertOnComplete", False)
fp.set_preference("browser.startup.homepage_override.mstone", "ignore")
fp.set_preference("startup.homepage_welcome_url.additional","about:blank")

browser = webdriver.Firefox(firefox_profile=fp)

### Create python list of author/title combinations to be searched, pulling the data from a csv file in the same directory ###

f = open('heinlist.csv')
searchlist = list(csv.reader(f))

### Go to HeinOnline ###

browser.get('http://www.heinonline.org/HOL/Welcome')

### Click on 'Advanced Search' ###

adv_search = browser.find_element_by_link_text('Advanced Search')
adv_search.click()

### Assign variable for listing errors ###

errorList = []

### Begin loop through all CSV elements ###

### Enter author and title search terms ###

for elem in range(len(searchlist)):
	searchstring_title = browser.find_element_by_xpath("/html/body/div[5]/div/div[2]/form/table/tbody/tr[1]/td[2]/input")
	searchstring_title.clear()
	searchstring_title.send_keys(searchlist[elem][1])
	searchstring_author = browser.find_element_by_xpath("/html/body/div[5]/div/div[2]/form/table/tbody/tr[2]/td[2]/input")
	searchstring_author.clear()
	searchstring_author.send_keys(searchlist[elem][0] + Keys.RETURN)
	
### Flag window so the program knows how to return to it after other windows have been opened/closed ###

	first_window = browser.current_window_handle

### Attempts to find the "PDF/Download" button
### If it fails, it appends the faulty search terms to a list and restarts the loop beginning with the next CSV element###

	try:
		downloadPDF = browser.find_element_by_link_text('PDF/Download')
	except:
		errorList.append(searchlist[elem])
		continue

### Click 'PDF/Download' for the first search result entry ###

	downloadPDF.click()

### Switch focus of script to the newly opened download window ###

	browser.switch_to_window(browser.window_handles[-1])

### Click the first 'Print/Download' button ###

	savePDF = browser.find_element_by_name('submit')
	savePDF.click()

### Make the script wait 5 seconds to allow time for download prompt to load ###

	time.sleep(5)

### Press 'Enter' to select 'OK' ###

	shell = win32com.client.Dispatch("WScript.Shell")
	shell.SendKeys("{ENTER}")

### Close download window and switch control back to the search window ###

	browser.close()	
	browser.switch_to_window(first_window)

### Loop restarts until end of csv list ###

### Writes the list of faulty searches to a CSV file ###

with open('csvErrorList.csv', 'wb') as csvfile:
	wr = csv.writer(csvfile, quoting=csv.QUOTE_ALL)
	wr.writerows(errorList)