################################################################################
# modules
################################################################################
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common import exceptions
import time
import pandas as pd
import openpyxl
import requests
import re
import numpy

################################################################################
# variable definitions
################################################################################
options = Options()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-gpu')
searchPageURL = "https://bexar.acttax.com/act_webdev/bexar/index.jsp" 
acctPageURL = "https://bexar.acttax.com/act_webdev/bexar/showdetail2.jsp?can=" 
propertyTaxDataList = [] #multidimensional array containing all property data
dudAddresses = [] #list of addresses with no results
# totalTimeElapsed = 0.0

################################################################################
# function definitions  
################################################################################

################################################################################
# excelFileIn(): Asks user for the excel filename containing a list of property
# addresses and stores the filename in a variable
# file: string variable containing filename
# return: file
################################################################################
def excelFileIn():
    file = input("Enter filename containing list of addresses: ")
    if file[-5:] == ".xlsx":
        file = file
    else:
        file = file + ".xlsx"
    return file

################################################################################
# getFileLength(): finds the number of lines in an excel file
# file: string variable containing filename
# return: number of lines in the file
################################################################################
def getFileLength(file):
    pyList = pd.read_excel(file)
    length = len(pyList.index)
    return length

################################################################################
# getAccountList(): finds the list of accounts associated with an address
# URL: the search page URL
# address: the address being searched
# return: a list of account numbers associated with the address
# ***if return is [0, 0, 0] then the search page did not load
################################################################################
def getAccountList(URL, address):
    searchDriver = webdriver.Chrome(options=options)
    searchDriver.set_page_load_timeout(120)
    searchDriver.get(URL)
    time.sleep(2)
    counter = 0
    
    while counter < 10: #tries to load the search page at most 10x with 2 sec
        #sleep between tries
        try:
            searchType = searchDriver.find_element_by_xpath('//*[@id="site-content"]'\
                                                            +'/table/tbody/tr/td/'\
                                                            +'center/form/table/tbody'\
                                                            +'/tr[3]/td[2]/div[3]/'\
                                                            +'select/option[2]').click()
            redoST = False #redo finding the "search type" drop down
        except exceptions.NoSuchElementException:
            redoST = True
        try:
            inputBox = searchDriver.find_element_by_xpath('//*[@id="criteria"]')
            inputBox.send_keys(address)
            redoIB = False #redo finding the input box
        except exceptions.NoSuchElementException:
            redoIB = True
        try:
            searchButton = searchDriver.find_element_by_name('submit').click()
            redoSB = False #redo finding the search button
        except exceptions.NoSuchElementException:
            redoSB = True
        counter = counter + 1
        if counter == 10: 
            print ("Search page not loading correctly for " + str(address))
            accounts = [0, 0, 0]
            return accounts #10th try, returning [0, 0, 0]
        if (redoST == False and redoIB == False and redoSB == False):
            break
        time.sleep(2)
    
    #move from search page to results page
    try:
        accounts = searchDriver.find_elements_by_xpath('//*[@id="account-'\
                                                       'container"]/td[2]/a')
        tempAccounts = []
        for account in accounts:
            tempAccounts.append(account.text) #changes accounts from webelement
            #to string
        accounts = tempAccounts #moves the string version of accounts back into
        #accounts list
    except exceptions.NoSuchElementException:
        formatHolder = 0

    try:
        searchButton = searchDriver.find_element_by_name('submit')
        print("No search results for " + address)
        accounts = [0, 0, 0]
    except exceptions.NoSuchElementException:
        formatHolder = 0
    
    searchDriver.close() 
    return accounts

################################################################################
# getPropertyData(): scrapes property data from the account page
# URL: the account page URL
# account: the account number
# return: list of property data
################################################################################
def getPropertyData(URL, account):
    pageDriver = webdriver.Chrome(options=options)
    pageDriver.get(URL + account)
    pageDriver.set_page_load_timeout(120)
    time.sleep(2)
    
    try:
        mapAvailable = pageDriver.find_element_by_xpath('//*[@id="site-content"]'\
                                                        +'/font/div/table[2]/'\
                                                        +'tbody/tr/td[1]/b/'\
                                                        +'b/a').text
    except exceptions.NoSuchElementException:
        mapAvailable = "Map Unavailable"
        
    if mapAvailable != "Map Unavailable":
        mapAvailable = "Map Available"
            
    try:
        mailAddress = pageDriver.find_element_by_xpath('//*[@id="site-content"]/font/div'\
                                                 +'/table[2]/tbody/tr/td[1]/div[2]').text
    except exceptions.NoSuchElementException:
        time.sleep(3)
        return False
    mailAddress = mailAddress.split('\n') 
    ownerName = mailAddress[1]
    mailingStreet = mailAddress[-2]
    mailingCity = mailAddress[-1]
    
    
    propAddress = pageDriver.find_element_by_xpath('//*[@id="site-content"]/font/'\
                                                 +'div/table[2]/tbody/tr/td[1]/div[3]').text
    propAddress = propAddress.split('\n')
    propAddress = propAddress[0]
    
    currentTaxLevy = pageDriver.find_element_by_xpath('//*[@id="site-content"]/font'\
                                                    +'/div/table[2]/tbody/tr/td[1]/div[5]').text
    currentTaxLevy = currentTaxLevy.split(' ')
    currentTaxLevy = currentTaxLevy[-1]
    currentTaxLevy = currentTaxLevy.replace('$', '')
    currentTaxLevy = currentTaxLevy.replace(',', '')
    
    currentAmtDue = pageDriver.find_element_by_xpath('//*[@id="site-content"]/font'\
                                                   +'/div/table[2]/tbody/tr/td[1]/div[6]').text
    currentAmtDue = currentAmtDue.split(' ')
    currentAmtDue = currentAmtDue[-1]
    currentAmtDue = currentAmtDue.replace('$', '')
    currentAmtDue = currentAmtDue.replace(',', '')
    
    dTest = pageDriver.find_element_by_xpath('//*[@id="site-content"]/font/div/table'\
                                           +'[2]/tbody/tr/td[1]/div[7]').text
    dTest = dTest.split(' ')
    
    if dTest[0] == "Delinquent":
        priorYearAmtDue = pageDriver.find_element_by_xpath('//*[@id="site-content"]/'\
                                                         +'font/div/table[2]/tbody/tr/td[1]/div[9]').text
        totalAmtDue = pageDriver.find_element_by_xpath('//*[@id="site-content"]/font'\
                                                     +'/div/table[2]/tbody/tr/td[1]/div[10]').text
    else:
        priorYearAmtDue = pageDriver.find_element_by_xpath('//*[@id="site-content"]/font/div'\
                                                         +'/table[2]/tbody/tr/td[1]/div[8]').text
        totalAmtDue = pageDriver.find_element_by_xpath('//*[@id="site-content"]/font/div/'\
                                                     +'table[2]/tbody/tr/td[1]/div[9]').text

    priorYearAmtDue = priorYearAmtDue.split(' ')
    priorYearAmtDue = priorYearAmtDue[-1]
    priorYearAmtDue = priorYearAmtDue.replace('$', '')
    priorYearAmtDue = priorYearAmtDue.replace(',', '')
    
    totalAmtDue = totalAmtDue.split(' ')
    totalAmtDue = totalAmtDue[-1]
    totalAmtDue = totalAmtDue.replace('$', '')
    totalAmtDue = totalAmtDue.replace(',', '')
    
    if totalAmtDue == "0.00":
        totalMktVal = pageDriver.find_element_by_xpath('//*[@id="site-content"]/font/div'\
                                                     +'/table[2]/tbody/tr/td[2]/div[2]').text
        exemptions = pageDriver.find_element_by_xpath('//*[@id="site-content"]/font/div/'\
                                                    +'table[2]/tbody/tr/td[2]/div[7]').text
    else:
        totalMktVal = pageDriver.find_element_by_xpath('//*[@id="site-content"]/font/div/'\
                                                     +'table[2]/tbody/tr/td[2]/div[3]').text
        exemptions = pageDriver.find_element_by_xpath('//*[@id="site-content"]/font/div/'\
                                                    +'table[2]/tbody/tr/td[2]/div[8]').text

    totalMktVal = totalMktVal.split(' ')
    totalMktVal = totalMktVal[-1]
    totalMktVal = totalMktVal.replace('$', '')
    totalMktVal = totalMktVal.replace(',', '')
    exemptions = exemptions.replace('Exemptions (current year only):\n', '')
    exemptions = exemptions.replace('\n', ', ')
    
    paymentHistPage = "https://bexar.acttax.com/act_webdev/bexar/reports/"\
    +"paymentinfo.jsp?can=" + account + "&ownerno=0"
    paymentHistLink = '=HYPERLINK("' + paymentHistPage + '", "Payment History")'
    payHistButton = pageDriver.find_element_by_link_text('Payment History').click()
    try:
        lastRec = pageDriver.find_element_by_xpath('//*[@id="site-content"]/table'\
                                                 +'/tbody[1]/tr[1]/td[1]').text
    except exceptions.NoSuchElementException:
        lastRec = "N/A"
    try:
        rollYear = pageDriver.find_element_by_xpath('//*[@id="site-content"]/table'\
                                                  +'/tbody[1]/tr[1]/td[2]').text
    except exceptions.NoSuchElementException:
        rollYear = "N/A"
    try:
        amount = pageDriver.find_element_by_xpath('//*[@id="site-content"]/table/'\
                                                +'tbody[1]/tr[1]/td[3]').text
        amount = amount.replace('$', '')
        amount = amount.replace(',', '')
    except exceptions.NoSuchElementException:
        amount = "N/A"
    try:
        description = pageDriver.find_element_by_xpath('//*[@id="site-content"]/'\
                                                     +'table/tbody[1]/tr[1]/td[4]').text
    except exceptions.NoSuchElementException:
        description = "N/A"
    try:
        payer = pageDriver.find_element_by_xpath('//*[@id="site-content"]/table/'\
                                               +'tbody[1]/tr[1]/td[5]').text
    except exceptions.NoSuchElementException:
        payer = "N/A"

    listItem = {
            'Property Address': address,
            'Account Number': account,
            'Owner Name': ownerName,
            'Mailing Street': mailingStreet,
            'Mailing City/State': mailingCity,
            'Map': mapAvailable,
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
    
    pageDriver.close()
    return listItem

################################################################################
# startIndex(): asks the user which address to start with and returns the index
# of that address in the list
# addresses: pandas dataframe of addresses
# lastRow: the index of the last address in addresses
# return: index; the index of the address given by the user
################################################################################
def startIndex(addresses, lastRow):
    index = -1
    addyStart = str(input("Type which address to start with in the list: "))
    addyStart = addyStart.lower()
    for k in range(0, lastRow):
        address = addresses.loc[k, 'Address'].lower()
        if addyStart == address:
            index = k
            return index

################################################################################
# main program  
################################################################################
inputExcelFile = excelFileIn()
lastRow = getFileLength(inputExcelFile)
addresses = pd.read_excel(inputExcelFile)
dudAddresses.append("Address")
initIndex = 0
initIndex = startIndex(addresses, lastRow)
for i in range(initIndex, lastRow):
    address = addresses.loc[i, 'Address']
    accounts = getAccountList(searchPageURL, address)
    if accounts == [0, 0, 0]:
        dudAddresses.append(address)
        outputDudFile = pd.DataFrame(dudAddresses)
        outputDudFile.to_excel("BexarNoResults.xlsx", header = False, index = False)
    else:
        for account in accounts:
            tries = 0
            while tries < 3:
                outputItem = getPropertyData(acctPageURL, account)
                if outputItem == False:
                    tries = tries + 1
                else:
                    break
            if tries == 3:
                continue
            propertyTaxDataList.append(outputItem)
            outputFile = pd.DataFrame(propertyTaxDataList)
            outputFile.to_excel("BexarTaxOutputFile.xlsx", index = False)
        
print("Output File Complete")
    