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

while True:
    county = str(input("Enter county you wish to search"))
    county = county.lower()
    searchURL = str(input("Paste the search page url: "))
    if searchURL[-13:] == "&pageSize=100":
        searchURL = searchURL
    else:
        searchURL = searchURL + "&pageSize=100"
    pageStart = int(input("Enter the index of the page you'd like to start the search (1, 2, etc.): "))
    print("Warning!!!!!!!!!")
    print("If you are restarting a previous search, enter a different filename for the output excel file.")
    print("The program will overwrite the previous file if you use the same filename.")
    outputFile = input('Enter output file name: ')
    if outputFile[-5:] == ".xlsx":
        outputFile = outputFile
    else:
        outputFile = outputFile + ".xlsx"

    referenceIDList = []  #initial list of reference IDs from search page
    propertyList = []  #list of properties and their attributes which will be written to an excel file
    totalTimeElapsed = 0.0

    if county == "austin":
        propURLPre = "https://esearch.austincad.org/Property/View/" # Austin County 
    elif county == "bandera":
        propURLPre = "https://esearch.bancad.org/Property/View/" # Bandera County
    elif county == "bell":
        propURLPre = "https://esearch.bellcad.org/Property/View/" # Bell County
    elif county == "bosque":
        propURLPre = "https://esearch.bosquecad.com/Property/View/" # Bosque County
    elif county == "brazos":
        propURLPre = "https://esearch.brazoscad.org/Property/View/" # Brazos County
    elif county == "briscoe":
        propURLPre = "https://esearch.briscoecad.org/Property/View/" # Briscoe County
    elif county == "burleson":
        propURLPre = "https://esearch.burlesonappraisal.com/Property/View/" # Burleson County
    elif county == "callahan":
        propURLPre = "https://esearch.callahancad.org/Property/View/" # Callahan County
    elif county == "camp":
        propURLPre = "https://esearch.campcad.org/Property/View/" # Camp County
    elif county == "carson":
        propURLPre = "https://esearch.carsoncad.org/Property/View/" # Carson County
    elif county == "coke":
        propURLPre = "https://esearch.cokecad.org/Property/View/" # Coke County
    elif county == "comanche":
        propURLPre = "https://esearch.comanchecad.org/Property/View/" # Comanche County
    elif county == "concho":
        propURLPre = "https://esearch.conchocad.org/Property/View/" # Concho County
    elif county == "coryell":
        propURLPre = "https://esearch.coryellcad.org/Property/View/" # Coryell County
    elif county == "crane":
        propURLPre = "https://esearch.cranecad.org/Property/View/" # Crane County
    elif county == "dallam":
        propURLPre = "https://esearch.dallamcad.org/Property/View/" # Dallam County
    elif county == "dimmit":
        propURLPre = "https://esearch.dimmit-cad.org/Property/View/" # Dimmit County
    elif county == "duval":
        propURLPre = "https://esearch.duvalcad.org/Property/View/" # Duval County
    elif county == "ellis":
        propURLPre = "https://esearch.elliscad.com/Property/View/" # Ellis County
    elif county == "erath":
        propURLPre = "https://esearch.erath-cad.com/Property/View/" # Erath County
    elif county == "fannin":
        propURLPre = "https://esearch.fannincad.org/Property/View/" # Fannin County
    elif county == "fayette":
        propURLPre = "https://esearch.fayettecad.org/Property/View/" # Fayette County
    elif county == "fisher":
        propURLPre = "https://esearch.fishercad.org/Property/View/" # Fisher County
    elif county == "gaines":
        propURLPre = "https://esearch.gainescad.org/Property/View/" # Gaines County
    elif county == "garza":
        propURLPre = "https://esearch.garzacad.org/Property/View/" # Garza County
    elif county == "gray":
        propURLPre = "https://esearch.graycad.org/Property/View/" # Gray County
    elif county == "grayson":
        propURLPre = "https://esearch.graysonappraisal.org/Property/View/" # Grayson County
    elif county == "gregg":
        propURLPre = "https://esearch.gcad.org/Property/View/" # Gregg County
    elif county == "hansford":
        propURLPre = "https://esearch.hansfordcad.org/Property/View/" # Hansford County
    elif county == "hardin":
        propURLPre = "https://esearch.hardin-cad.org/Property/View/" # Hardin County
    elif county == "hays":
        propURLPre = "https://esearch.hayscad.com/Property/View/" # Hays County
    elif county == "henderson":
        propURLPre = "https://esearch.henderson-cad.org/Property/View/" # Henderson County
    elif county == "howard":
        propURLPre = "https://esearch.howardcad.org/Property/View/" # Howard County
    elif county == "hunt":
        propURLPre = "https://esearch.hunt-cad.org/Property/View/" # Hunt County
    elif county == "jim hogg":
        propURLPre = "https://esearch.jimhogg-cad.org/Property/View/" # Jim Hogg County
    elif county == "kaufman":
        propURLPre = "https://esearch.kaufman-cad.org/Property/View/" # Kaufman County
    elif county == "kendall":
        propURLPre = "https://esearch.kendallad.org/Property/View/" # Kendall County
    elif county == "la salle":
        propURLPre = "https://esearch.lasallecad.com/Property/View/" # La Salle County
    elif county == "lamar":
        propURLPre = "https://esearch.lamarcad.org/Property/View/" # Lamar County
    elif county == "lee":
        propURLPre = "https://esearch.lee-cad.org/Property/View/" # Lee County
    elif county == "liberty":
        propURLPre = "https://esearch.libertycad.com/Property/View/" # Liberty County
    elif county == "llano":
        propURLPre = "https://esearch.llanocad.net/Property/View/" # Llano County
    elif county == "madison":
        propURLPre = "https://esearch.madisoncad.org/Property/View/" # Madison County
    elif county == "mcmullen":
        propURLPre = "https://esearch.mcmullencad.org/Property/View/" # McMullen County
    elif county == "medina":
        propURLPre = "https://esearch.medinacad.org/Property/View/" # Medina County
    elif county == "mills":
        propURLPre = "https://esearch.millscad.org/Property/View/" # Mills County
    elif county == "moore":
        propURLPre = "https://esearch.moorecad.org/Property/View/" # Moore County
    elif county == "navarro":
        propURLPre = "https://esearch.navarrocad.com/Property/View/" # Navarro County
    elif county == "oldham":
        propURLPre = "https://esearch.oldhamcad.org/Property/View/" # Oldham County
    elif county == "orange":
        propURLPre = "https://esearch.orangecad.net/Property/View/" # Orange County
    elif county == "parmer":
        propURLPre = "https://esearch.parmercad.org/Property/View/" # Parmer County
    elif county == "pecos":
        propURLPre = "https://esearch.pecoscad.org/Property/View/" # Pecos County
    elif county == "rains":
        propURLPre = "https://esearch.rainscad.org/Property/View/" # Rains County
    elif county == "real":
        propURLPre = "https://esearch.realcad.org/Property/View/" # Real County
    elif county == "reeves":
        propURLPre = "https://esearch.reevescad.org/Property/View/" # Reeves County
    elif county == "roberts":
        propURLPre = "https://esearch.robertscad.org/Property/View/" # Roberts County
    elif county == "san patricio":
        propURLPre = "https://esearch.sanpatcad.org/Property/View/" # San Patrico County
    elif county == "san saba":
        propURLPre = "https://esearch.sansabacad.org/Property/View/" # San Saba County
    elif county == "schleicher":
        propURLPre = "https://esearch.schleichercad.org/Property/View/" # Schleicher County
    elif county == "shackelford":
        propURLPre = "https://esearch.shackelfordcad.com/Property/View/" # Shackelford County
    elif county == "somervell":
        propURLPre = "https://esearch.somervellcad.net/Property/View/" # Somervell County
    elif county == "throckmorton":
        propURLPre = "http://esearch.throckmortoncad.org/Property/View/" # Throckmorton County
    elif county == "titus":
        propURLPre = "https://esearch.titus-cad.org/Property/View/" # Titus County
    elif county == "van zandt":
        propURLPre = "https://esearch.vzcad.org/Property/View/" # Van Zandt County
    elif county == "walker":
        propURLPre = "https://esearch.walkercad.org/Property/View/" # Walker County
    elif county == "wharton":
        propURLPre = "https://esearch.whartoncad.net/Property/View/" # Wharton County
    elif county == "wood":
        propURLPre = "https://esearch.woodcad.net/Property/View/" # Wood County
    elif county == "zavala":
        propURLPre = "https://esearch.zavalacad.com/Property/View/" # Zavala County
    else
        print("You did not enter a supported county.")
        userChoice = str(input("Retry? y/n: "))
        if userChoice == 'n':
            return
        elif userChoice != 'y':
            print("Invalid choice. Program terminated.")
            return


#############################################
# function name: referenceIDSearch
# purpose: pulls the list of reference ID's from 
# each search page
# input: none
# return: none
#############################################
def referenceIDSearch():
    for i in range(1, 101):
        try:
            referenceID = driver.find_element_by_xpath('//*[@id="grid"]/div[2]/table/tbody/tr[' + str(i) + ']/td[2]').text
            referenceIDList.append(referenceID)
        except exceptions.NoSuchElementException:
            break 
            
#############################################
# function name: propertyDataFunction
# purpose: pulls pertinent property data from 
# each individual property page and populates
# a property data list
# input: none
# return: none
#############################################
def propertyDataFunction():
    for j in range(0, len(referenceIDList)):
        propertyURL = propURLPre + referenceIDList[j]
        propertyDriver = webdriver.Chrome(options=options)
        propertyDriver.set_page_load_timeout(45)
        try:
            propertyDriver.get(propertyURL)
            time.sleep(2)
        except exceptions.TimeoutException:
            print("There was an error with " + referenceIDList[j] + ". The page timed out.")
            continue

        # property details box on website
        quickRefID = referenceIDList[j]
        webpageHyperlink = '=HYPERLINK("' + propertyURL + '")'
        try:
            legalDescription = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]/div[1]/div/table/tbody/tr[3]/td').text
        except exceptions.NoSuchElementException:
            legalDescription = ""
        try:
            geoID = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]/div[1]/div/table/tbody/tr[4]/td').text
        except exceptions.NoSuchElementException:
            geoID = ""
        try:
            agent = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]/div[1]/div/table/tbody/tr[5]/td').text
        except exceptions.NoSuchElementException:
            agent = ""
        try:
            propertyType = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]/div[1]/div/table/tbody/tr[6]/td').text
        except exceptions.NoSuchElementException:
            propertyType = ""
        try:
            propertyAddress = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]/div[1]/div/table/tbody/tr[8]/td').text
        except exceptions.NoSuchElementException:
            propertyAddress = ""
        try:
            mapID = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]/div[1]/div/table/tbody/tr[9]/td').text
        except exceptions.NoSuchElementException:
            mapID = ""
        try:
            neighborhoodCode = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]/div[1]/div/table/tbody/tr[10]/td').text
        except exceptions.NoSuchElementException:
            neighborhoodCode = ""
        try:
            ownerID = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]/div[1]/div/table/tbody/tr[12]/td').text
        except exceptions.NoSuchElementException:
            ownerID = ""
        try:
            ownerName = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]/div[1]/div/table/tbody/tr[13]/td').text
        except exceptions.NoSuchElementException:
            ownerName = ""
        try:
            ownerAddress = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]/div[1]/div/table/tbody/tr[14]/td').text
        except exceptions.NoSuchElementException:
            ownerAddress = ""
        try:
            percentOwnership = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]/div[1]/div/table/tbody/tr[15]/td').text
        except exceptions.NoSuchElementException:
            percentOwnership = ""
        try:
            taxExemptions = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]/div[1]/div/table/tbody/tr[16]/td/span').text
        except exceptions.NoSuchElementException:
            taxExemptions = ""

        # property values box on website
        try:
            improvementHomesiteVal = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]/div[2]/div[1]/table/tbody/tr[2]/td').text
        except exceptions.NoSuchElementException:
            improvementHomesiteVal = ""
        try:
            improvementNonHomesiteVal = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]/div[2]/div[1]/table/tbody/tr[3]/td').text
        except exceptions.NoSuchElementException:
            improvementNonHomesiteVal = ""
        try:
            landHomesiteVal = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]/div[2]/div[1]/table/tbody/tr[4]/td').text
        except exceptions.NoSuchElementException:
            landHomesiteVal = ""
        try:
            landNonHomesiteVal = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]/div[2]/div[1]/table/tbody/tr[5]/td').text
        except exceptions.NoSuchElementException:
            landNonHomesiteVal = ""
        try:
            agMarketVal = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]/div[2]/div[1]/table/tbody/tr[6]/td').text
        except exceptions.NoSuchElementException:
            agMarketVal = ""
        try:
            valueMethod = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]/div[2]/div[1]/table/tbody/tr[8]/td').text
        except exceptions.NoSuchElementException:
            valueMethod = ""
        try:
            marketVal = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]/div[2]/div[1]/table/tbody/tr[9]/td').text
        except exceptions.NoSuchElementException:
            marketVal = ""
        try:
            agUseVal = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]/div[2]/div[1]/table/tbody/tr[10]/td').text
        except exceptions.NoSuchElementException:
            agUseVal = ""
        try:
            appraisedVal = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]/div[2]/div[1]/table/tbody/tr[12]/td').text
        except exceptions.NoSuchElementException:
            appraisedVal = ""
        try:
            homesteadCapLoss = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]/div[2]/div[1]/table/tbody/tr[13]/td').text
        except exceptions.NoSuchElementException:
            homesteadCapLoss = ""
        try:
            assessedVal = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]/div[2]/div[1]/table/tbody/tr[15]/td').text
        except exceptions.NoSuchElementException:
            assessedVal = ""

        # property land box on website
        try:
            propertyLandType = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[6]/div[2]/table/tbody/tr[2]/td[1]').text
        except exceptions.NoSuchElementException:
            propertyLandType = ""
        try:
            landTypeDescription = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[6]/div[2]/table/tbody/tr[2]/td[2]').text
        except exceptions.NoSuchElementException:
            landTypeDescription = ""
        try:
            acreage = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[6]/div[2]/table/tbody/tr[2]/td[3]').text
        except exceptions.NoSuchElementException:
            acreage = ""
        try:
            landSquareFeet = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[6]/div[2]/table/tbody/tr[2]/td[4]').text
        except exceptions.NoSuchElementException:
            landSquareFeet = ""
        try:
            effectiveFrontage = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[6]/div[2]/table/tbody/tr[2]/td[5]').text
        except exceptions.NoSuchElementException:
            effectiveFrontage = ""
        try:
            effectiveDepth = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[6]/div[2]/table/tbody/tr[2]/td[6]').text
        except exceptions.NoSuchElementException:
            effectiveDepth = ""
        try:
            landMarketVal = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[6]/div[2]/table/tbody/tr[2]/td[7]').text
        except exceptions.NoSuchElementException:
            landMarketVal = ""
        try:
            landProductionVal = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[6]/div[2]/table/tbody/tr[2]/td[8]').text
        except exceptions.NoSuchElementException:
            landProductionVal = ""

        # property deed history box on website
        deedDate = []
        deedType = []
        documentDesc = []
        deedGrantor = []
        deedGrantee = []
        deedVolume = []
        deedPage = []
        deedNumber = []
        moreDeed = '=HYPERLINK("' + propertyURL + '")'

        for k in range (2, 5):
            try:
                dDate = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[8]/div[2]/table/tbody/tr[' + str(k) + ']/td[1]').text
                deedDate.append(dDate)
            except exceptions.NoSuchElementException:
                dDate = ""
                deedDate.append(dDate)
            try:
                dType = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[8]/div[2]/table/tbody/tr[' + str(k) + ']/td[2]').text
                deedType.append(dType)
            except exceptions.NoSuchElementException:
                dType = ""
                deedType.append(dType)
            try:
                dDesc = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[8]/div[2]/table/tbody/tr[' + str(k) + ']/td[3]').text
                documentDesc.append(dDesc)
            except exceptions.NoSuchElementException:
                dDesc = ""
                documentDesc.append(dDesc)
            try:
                dGrantor = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[8]/div[2]/table/tbody/tr[' + str(k) + ']/td[4]').text
                deedGrantor.append(dGrantor)
            except exceptions.NoSuchElementException:
                dGrantor = ""
                deedGrantor.append(dGrantor)
            try:
                dGrantee = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[8]/div[2]/table/tbody/tr[' + str(k) + ']/td[5]').text
                deedGrantee.append(dGrantee)
            except exceptions.NoSuchElementException:
                dGrantee = ""
                deedGrantee.append(dGrantee)
            try:
                dVolume = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[8]/div[2]/table/tbody/tr[' + str(k) + ']/td[6]').text
                deedVolume.append(dVolume)
            except exceptions.NoSuchElementException:
                dVolume = ""
                deedVolume.append(dVolume)
            try:
                dPage = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[8]/div[2]/table/tbody/tr[' + str(k) + ']/td[7]').text
                deedPage.append(dPage)
            except exceptions.NoSuchElementException:
                dPage = ""
                deedPage.append(dPage)
            try:
                dNumber = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[8]/div[2]/table/tbody/tr[' + str(k) + ']/td[8]').text
                deedNumber.append(dNumber)
            except exceptions.NoSuchElementException:
                dNumber = ""
                deedNumber.append(dNumber)
        if deedDate[1] == "" and deedDate[2] == "":
            moreDeed = "False"
            
        ########## property data list ##############
        propertyData = {
            'Quick Reference ID': quickRefID,
            'Property Webpage Hyperlink': webpageHyperlink,
            'Property Address': propertyAddress,
            'Owner ID': ownerID,
            'Owner Name': ownerName,
            'Owner Address': ownerAddress,
            'Percent Ownership': percentOwnership,
            'Tax Exemptions': taxExemptions,
            'Agent': agent,
            'Property Type': propertyType,
            'Legal Description': legalDescription,
            'Geographic ID': geoID,
            'Map ID': mapID,
            'Neighborhood Code': neighborhoodCode,
            'Improvement Homesite Value': improvementHomesiteVal,
            'Improvement Non-Homesite Value': improvementNonHomesiteVal,
            'Land Homesite Value': landHomesiteVal,
            'Land Non-Homesite Value': landNonHomesiteVal,
            'Market Value': marketVal,
            'Ag Market Value': agMarketVal,
            'Ag Use Value': agUseVal,
            'Appraised Value': appraisedVal,
            'Value Method': valueMethod,
            'Assessed Value': assessedVal,
            'Homestead Cap Loss': homesteadCapLoss,
            'Property Land Type': propertyLandType,
            'Land Type Description': landTypeDescription,
            'Acreage': acreage,
            'Square Footage': landSquareFeet,
            'Frontage': effectiveFrontage,
            'Depth': effectiveDepth,
            'Land Market Value': landMarketVal,
            'Land Production Value': landProductionVal,
            'Deed Date': deedDate[0],
            'Deed Type': deedType[0],
            'Document Description': documentDesc[0],
            'Deed Grantor': deedGrantor[0],
            'Deed Grantee': deedGrantee[0],
            'Deed Volume': deedVolume[0],
            'Number of pages': deedPage[0],
            'Deed Number': deedNumber[0],
            'More Deed Info': moreDeed,
        }
        #############################################

        ############ deed history list ##############
#         deedHistory = {
#             'Quick Reference ID': quickRefID,
#             'Property Webpage Hyperlink': webpageHyperlink,
#             'Property Address': propertyAddress,
#             'Deed Date': deedDate,
#             'Deed Type': deedType,
#             'Document Description': documentDesc,
#             'Deed Grantor': deedGrantor,
#             'Deed Grantee': deedGrantee,
#             'Deed Volume': deedVolume,
#             'Number of pages': deedPage,
#             'Deed Number': deedNumber,
#         }
        #############################################
       
        propertyList.append(propertyData)
        propertyDriver.close()
        
        #############################################
        #       write to master excel file  
        #############################################
        df = pd.DataFrame(propertyList)
        df.to_excel(outputFile)

#############################################
#              search begin  
#############################################
driver = webdriver.Chrome(options=options)
driver.set_page_load_timeout(45)
tic = time.perf_counter()  # ********** timer start **********
try:
    driver.get(searchURL)
except exceptions.TimeoutException:
    print("There was a timeout error with page 1 of ", lastPage)
time.sleep(2)

# display number of pages and number of properties in the results
header = driver.find_element_by_xpath('//*[@id="page-header"]').text
header = header.split()
lastPage = int(header[3])
totalProperties = int(header[6])
print("There are " + str(lastPage) + " pages and " + str(totalProperties) + " properties in your search results")

if pageStart == 1:
    referenceIDSearch()
    propertyDataFunction()
else:
    print("Search will begin on page:", pageStart)
    
driver.close()
toc = time.perf_counter()  # ********** timer stop **********
totalTimeElapsed = totalTimeElapsed + (toc-tic)
print("Page 1 of ", lastPage, " -- Done -- ", (toc-tic)/60, " Minutes, ", totalTimeElapsed/60, " Minutes Total")

if pageStart > 1:
    for p in range(pageStart, lastPage+1):  # handling subsequent search pages
        driver = webdriver.Chrome(options=options)
        driver.set_page_load_timeout(45)
        tic = time.perf_counter()  # ********** timer start **********
        searchURL = searchURL +  '&page=' + str(p)
        try:
            driver.get(searchURL)
        except exceptions.TimeoutException:
            print("There was a timeout error with page ", p, " of ", lastPage)
            continue
        time.sleep(2)
        referenceIDSearch()
        propertyDataFunction()
        driver.close()
        toc = time.perf_counter()  # ********** timer stop **********
        totalTimeElapsed = totalTimeElapsed + (toc-tic)
        print("Page ", p, " of ", lastPage, " -- Done -- ", (toc-tic)/60, " Minutes, ", totalTimeElapsed/60, " Minutes Total")  # status update
else:   
    for p in range(2, lastPage):  # handling subsequent search pages
        driver = webdriver.Chrome(options=options)
        driver.set_page_load_timeout(45)
        tic = time.perf_counter()  # ********** timer start **********
        searchURL = searchURL +  '&page=' + str(p)
        try:
            driver.get(searchURL)
        except exceptions.TimeoutException:
            print("There was a timeout error with page ", p, " of ", lastPage)
            continue
        time.sleep(2)
        referenceIDSearch()
        propertyDataFunction()
        driver.close()
        toc = time.perf_counter()  # ********** timer stop **********
        totalTimeElapsed = totalTimeElapsed + (toc-tic)
        print("Page ", p, " of ", lastPage, " -- Done -- ", (toc-tic)/60, " Minutes, ", totalTimeElapsed/60, " Minutes Total")  # status update

print("Output File Complete") 