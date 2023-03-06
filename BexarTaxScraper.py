#############################################
#                modules
#############################################
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common import exceptions
import time
import pandas as pd
import openpyxl
import requests
import re
import numpy

#############################################
#              variables  
#############################################
options = Options()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-gpu')
bexarTaxURL = "https://bexar.acttax.com/act_webdev/bexar/index.jsp"
APURL = "https://bexar.acttax.com/act_webdev/bexar/showdetail2.jsp?can="
propertyTaxDataList = []
totalTimeElapsed = 0.0

inputFile = input("Enter filename containing list of addresses: ")
if inputFile[-5:] == ".xlsx":
    inputFile = inputFile
else:
    inputFile = inputFile + ".xlsx"
    
addressList = pd.read_excel(inputFile)
addressList.to_excel("backupAddressList.xlsx")
print("A backup of the original address list was saved as backupAddressList.xlsx")
lastRow = len(addressList.index)

#############################################
# function name: propertyTaxData
# purpose: pulls pertinent property tax data  
# from each individual property page and 
# populates a property data list
# input: address; the address of each property
# return: none
#############################################
def propertyTaxData(address):
    searchDriver = webdriver.Chrome(options=options)
    searchDriver.set_page_load_timeout(45)
    searchDriver.get(bexarTaxURL)
    time.sleep(2)
    counter = 0
    while True and counter < 10:
        try:
            searchType = searchDriver.find_element_by_xpath('//*[@id="site-content"]/table/tbody/tr/td/center/form/table/tbody/tr[3]/td[2]/div[3]/select/option[2]').click()
            redoST = False
        except exceptions.NoSuchElementException:
            
        try:    
            searchBox = searchDriver.find_element_by_xpath('//*[@id="criteria"]')
            redoSB = False
        except exceptions.NoSuchElementException:
            
        searchBox.send_keys(address)
        try:
            searchButton = searchDriver.find_element_by_name('submit').click()
            redoSButton = False
        except exceptions.NoSuchElementException:
        
        counter = counter + 1
        time.sleep(2)
        if counter == 10:
            print ("Search page not loading correctly for " + str(address))
            return
        if (redoST == False and redoSB == False and redoSButton == False):
            break
    
    ###### initial account page data ######
    try:
        accountPages = searchDriver.find_elements_by_xpath('//*[@id="account-container"]/td[2]/a')
        pages = []
        for accountPage in accountPages:
             pages.append(accountPage.text)
    except exceptions.NoSuchElementException: # populates the variables with blanks when there are no search results
        print("No search results for " + address) 
        accountNumber = "No result"
        ownerName = ""
        mailingStreet = ""
        mailingCity = ""
        currentTaxLevy = ""
        currentAmtDue = ""
        priorYearAmtDue = ""
        totalAmtDue = ""
        totalMktVal = ""
        exemptions = ""
        lastRec = ""
        rollYear = ""
        amount = ""
        description = ""
        payer = ""
        paymentHistLink = "" 
        
        ###### property data list ######
        listItem = {
            'Property Address': address,
            'Account Number': accountNumber,
            'Owner Name': ownerName,
            'Mailing Street': mailingStreet,
            'Mailing City/State': mailingCity,
            'Current Year Tax Levy': currentTaxLevy,
            'Current Year Amount Due': currentAmtDue,
            'Prior Year Amount Due': priorYearAmtDue,
            'Total Amount Due': totalAmtDue,
            'Total Market Value': totalMktVal,
            'Exemptions': exemptions,
            'Last Payment Receipt Date': lastRec,
            'Roll Year': rollYear,
            'Amount': amount,
            'Description': description,
            'Payer': payer,
            'Payment History': paymentHistLink
        }

        propertyTaxDataList.append(listItem)
        searchDriver.close()
        
        ###### write to master excel file ###### 
        outputList = pd.DataFrame(propertyTaxDataList)
        outputList.to_excel(inputFile)
        return
    
    searchDriver.close()

    for page in pages:
        APDriver = webdriver.Chrome(options=options)
        APDriver.get(APURL + page)
        APDriver.set_page_load_timeout(45)
        time.sleep(2)
        try:
            mapAvail = APDriver.find_element_by_xpath('//*[@id="site-content"]/font/div/table[2]/tbody/tr/td[1]/b/b/a')
            accountNumber = APDriver.find_element_by_xpath('//*[@id="site-content"]/font/div/table[2]/tbody/tr/td[1]/div[1]').text
            accountNumber = accountNumber.split(' ')
            accountNumber = accountNumber[2]
            
            mailAddress = APDriver.find_element_by_xpath('//*[@id="site-content"]/font/div/table[2]/tbody/tr/td[1]/div[2]').text
            mailAddress = mailAddress.split('\n')
            ownerName = mailAddress[1]
            mailingStreet = mailAddress[2]
            mailingCity = mailAddress[3]
            
            propAddress = APDriver.find_element_by_xpath('//*[@id="site-content"]/font/div/table[2]/tbody/tr/td[1]/div[3]').text
            propAddress = propAddress.split('\n')
            propAddress = propAddress[1]
            if address != propAddress:
                address = propAddress
            
            currentTaxLevy = APDriver.find_element_by_xpath('//*[@id="site-content"]/font/div/table[2]/tbody/tr/td[1]/div[5]').text
            currentTaxLevy = currentTaxLevy.split(' ')
            currentTaxLevy = currentTaxLevy[-1]
            currentTaxLevy = currentTaxLevy.replace('$', '')
            currentTaxLevy = currentTaxLevy.replace(',', '')
           
            currentAmtDue = APDriver.find_element_by_xpath('//*[@id="site-content"]/font/div/table[2]/tbody/tr/td[1]/div[6]').text
            currentAmtDue = currentAmtDue.split(' ')
            currentAmtDue = currentAmtDue[-1]
            currentAmtDue = currentAmtDue.replace('$', '')
            currentAmtDue = currentAmtDue.replace(',', '')
            
            dTest = APDriver.find_element_by_xpath('//*[@id="site-content"]/font/div/table[2]/tbody/tr/td[1]/div[7]').text
            dTest = dTest.split(' ')
            
            if dTest[0] == "Delinquent":
                priorYearAmtDue = APDriver.find_element_by_xpath('//*[@id="site-content"]/font/div/table[2]/tbody/tr/td[1]/div[9]').text
                totalAmtDue = APDriver.find_element_by_xpath('//*[@id="site-content"]/font/div/table[2]/tbody/tr/td[1]/div[10]').text
            else:
                priorYearAmtDue = APDriver.find_element_by_xpath('//*[@id="site-content"]/font/div/table[2]/tbody/tr/td[1]/div[8]').text
                totalAmtDue = APDriver.find_element_by_xpath('//*[@id="site-content"]/font/div/table[2]/tbody/tr/td[1]/div[9]').text
            
            priorYearAmtDue = priorYearAmtDue.split(' ')
            priorYearAmtDue = priorYearAmtDue[-1]
            priorYearAmtDue = priorYearAmtDue.replace('$', '')
            priorYearAmtDue = priorYearAmtDue.replace(',', '')
            
            totalAmtDue = totalAmtDue.split(' ')
            totalAmtDue = totalAmtDue[-1]
            totalAmtDue = totalAmtDue.replace('$', '')
            totalAmtDue = totalAmtDue.replace(',', '')
            
            if totalAmtDue == "0.00":
                totalMktVal = APDriver.find_element_by_xpath('//*[@id="site-content"]/font/div/table[2]/tbody/tr/td[2]/div[2]').text
                exemptions = APDriver.find_element_by_xpath('//*[@id="site-content"]/font/div/table[2]/tbody/tr/td[2]/div[7]').text
            else:
                totalMktVal = APDriver.find_element_by_xpath('//*[@id="site-content"]/font/div/table[2]/tbody/tr/td[2]/div[3]').text
                exemptions = APDriver.find_element_by_xpath('//*[@id="site-content"]/font/div/table[2]/tbody/tr/td[2]/div[8]').text
                
            totalMktVal = totalMktVal.split(' ')
            totalMktVal = totalMktVal[-1]
            totalMktVal = totalMktVal.replace('$', '')
            totalMktVal = totalMktVal.replace(',', '')
            exemptions = exemptions.replace('Exemptions (current year only):\n', '')
            exemptions = exemptions.replace('\n', ', ')

            ###### payment history page data ######
            paymentHistPage = "https://bexar.acttax.com/act_webdev/bexar/reports/paymentinfo.jsp?can=" + accountNumber + "&ownerno=0"
            paymentHistLink = '=HYPERLINK("' + paymentHistPage + '", "Payment History")'
            payHistButton = APDriver.find_element_by_link_text('Payment History').click()
            lastRec = APDriver.find_element_by_xpath('//*[@id="site-content"]/table/tbody[1]/tr[1]/td[1]').text
            rollYear = APDriver.find_element_by_xpath('//*[@id="site-content"]/table/tbody[1]/tr[1]/td[2]').text
            amount = APDriver.find_element_by_xpath('//*[@id="site-content"]/table/tbody[1]/tr[1]/td[3]').text
            amount = amount.replace('$', '')
            amount = amount.replace(',', '')
            description = APDriver.find_element_by_xpath('//*[@id="site-content"]/table/tbody[1]/tr[1]/td[4]').text
            payer = APDriver.find_element_by_xpath('//*[@id="site-content"]/table/tbody[1]/tr[1]/td[5]').text

            ###### property data list ######
            listItem = {
                'Property Address': address,
                'Account Number': accountNumber,
                'Owner Name': ownerName,
                'Mailing Street': mailingStreet,
                'Mailing City/State': mailingCity,
                'Current Year Tax Levy': currentTaxLevy,
                'Current Year Amount Due': currentAmtDue,
                'Prior Year Amount Due': priorYearAmtDue,
                'Total Amount Due': totalAmtDue,
                'Total Market Value': totalMktVal,
                'Exemptions': exemptions,
                'Last Payment Receipt Date': lastRec,
                'Roll Year': rollYear,
                'Amount': amount,
                'Description': description,
                'Payer': payer,
                'Payment History': paymentHistLink
            }

            propertyTaxDataList.append(listItem)
            APDriver.close()

            ###### write to master excel file ###### 
            outputList = pd.DataFrame(propertyTaxDataList)
            outputList.to_excel(inputFile)
            print("Data for ", address," saved to file")
            
        except exceptions.NoSuchElementException:
            continue
    
    
#############################################
#             search program  
#############################################       
for i in range(0, lastRow):
    nextAddress = addressList.loc[i, 'Address']
    propertyTaxData(nextAddress)
    
print("Output File Complete")