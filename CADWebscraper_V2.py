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
countyList = {
    "austin",
    "bandera",
    "bell",
    "bosque",
    "brazos",
    "briscoe",
    "burleson",
    "callahan",
    "camp",
    "carson",
    "coke",
    "comanche",
    "concho",
    "coryell",
    "crane",
    "dallam",
    "dimmit",
    "duval",
    "ellis",
    "erath",
    "fannin",
    "fayette",
    "fisher",
    "gaines",
    "garza",
    "gray",
    "grayson",
    "gregg",
    "hansford",
    "hardin",
    "hays",
    "henderson",
    "howard",
    "hunt",
    "jim Hogg",
    "kaufman",
    "kendall",
    "la Salle",
    "lamar",
    "lee",
    "liberty",
    "llano",
    "madison",
    "mcMullen",
    "medina",
    "mills",
    "moore",
    "navarro",
    "oldham",
    "orange",
    "parmer",
    "pecos",
    "rains",
    "real",
    "reeves",
    "roberts",
    "san patricio",
    "san saba",
    "schleicher",
    "shackelford",
    "somervell",
    "throckmorton",
    "titus",
    "van zandt",
    "walker",
    "wharton",
    "wood",
    "zavala",
}
county = str(input("Enter county you wish to search: "))
county = county.lower()

################################################################################
# function definitions  
################################################################################

################################################################################
# inputLoop(county): places the user in an input loop if the selected county
# is not supported by the program
# county: county name input by user
# return: no return
################################################################################
def inputLoop(county):
    if checkCounty(county) == False:
        print("You entered an unsupported county.")
        chooseList = str(input("Would you like to see a list of supported counties? y/n: "))
        chooseList = chooseList.lower()
        if chooseList == "y":
            for item in countyList:
                print(item)
        chooseAgain = str(input("Enter another county? y/n: "))
        chooseAgain = chooseAgain.lower()
        if chooseAgain == "y":
            county = str(input("Enter county you wish to search: "))
            county = county.lower()
            inputLoop(county)

################################################################################
# checkCounty(county): checks the county input by the user against the counties
# that are supported by the program
# county: county name input by user
# return: True if the county is valid; False if the county is not valid
################################################################################
def checkCounty(county):
    if county in countyList:
        return True
    else:
        return False

################################################################################
# getSearchPage(county): finds the URL of the search page for the county that
# is input by the user
# county: county name input by user
# return: the URL of the search page as a string
################################################################################
def getSearchPage(county):
    if county == "austin":
        page = "https://esearch.austincad.org/"
    elif county == "bandera":
        page = "https://esearch.bancad.org/"
    elif county == "bell":
        page = "https://esearch.bellcad.org/"
    elif county == "bosque":
        page = "https://esearch.bosquecad.com/"
    elif county == "brazos":
        page = "https://esearch.brazoscad.org/"
    elif county == "briscoe":
        page = "https://esearch.briscoecad.org/"
    elif county == "burleson":
        page = "https://esearch.burlesonappraisal.com/"
    elif county == "callahan":
        page = "https://esearch.callahancad.org/"
    elif county == "camp":
        page = "https://esearch.campcad.org/"
    elif county == "carson":
        page = "https://esearch.carsoncad.org/"
    elif county == "coke":
        page = "https://esearch.cokecad.org/"
    elif county == "comanche":
        page = "https://esearch.comanchecad.org/"
    elif county == "concho":
        page = "https://esearch.conchocad.org/"
    elif county == "coryell":
        page = "https://esearch.coryellcad.org/"
    elif county == "crane":
        page = "https://esearch.cranecad.org/"
    elif county == "dallam":
        page = "https://esearch.dallamcad.org/"
    elif county == "dimmit":
        page = "https://esearch.dimmit-cad.org/"
    elif county == "duval":
        page = "https://esearch.duvalcad.org/"
    elif county == "erath":
        page = "https://esearch.erath-cad.com/"
    elif county == "fannin":
        page = "https://esearch.fannincad.org/"
    elif county == "fayette":
        page = "https://esearch.fayettecad.org/"
    elif county == "fisher":
        page = "https://esearch.fishercad.org/"
    elif county == "gaines":
        page = "https://esearch.gainescad.org/"
    elif county == "garza":
        page = "https://esearch.garzacad.org/"
    elif county == "gray":
        page = "https://esearch.graycad.org/"
    elif county == "grayson":
        page = "https://esearch.graysonappraisal.org/"
    elif county == "gregg":
        page = "https://esearch.gcad.org/"
    elif county == "hansford":
        page = "https://esearch.hansfordcad.org/"
    elif county == "hardin":
        page = "https://esearch.hardin-cad.org/"
    elif county == "hays":
        page = "https://esearch.hayscad.com/"
    elif county == "henderson":
        page = "https://esearch.henderson-cad.org/"
    elif county == "howard":
        page = "https://esearch.howardcad.org/"
    elif county == "hunt":
        page = "https://esearch.hunt-cad.org/"
    elif county == "jim hogg":
        page = "https://esearch.jimhogg-cad.org/"
    elif county == "kaufman":
        page = "https://esearch.kaufman-cad.org/"
    elif county == "kendall":
        page = "https://esearch.kendallad.org/"
    elif county == "la salle":
        page = "https://esearch.lasallecad.com/"
    elif county == "lamar":
        page = "https://esearch.lamarcad.org/"
    elif county == "lee":
        page = "http://esearch.lee-cad.org/"
    elif county == "liberty":
        page = "https://esearch.libertycad.com/"
    elif county == "llano":
        page = "https://esearch.llanocad.net/"
    elif county == "madison":
        page = "https://esearch.madisoncad.org/"
    elif county == "mcmullen":
        page = "https://esearch.mcmullencad.org/"
    elif county == "medina":
        page = "https://esearch.medinacad.org/"
    elif county == "mills":
        page = "https://esearch.millscad.org/"
    elif county == "moore":
        page = "http://esearch.moorecad.org/"
    elif county == "navarro":
        page = "http://esearch.navarrocad.com/"
    elif county == "oldham":
        page = "https://esearch.oldhamcad.org/"
    elif county == "orange":
        page = "https://esearch.orangecad.net/"
    elif county == "parmer":
        page = "https://esearch.parmercad.org/"
    elif county == "pecos":
        page = "https://esearch.pecoscad.org/"
    elif county == "rains":
        page = "http://esearch.rainscad.org/"
    elif county == "real":
        page = "https://esearch.realcad.org/"
    elif county == "reeves":
        page = "https://esearch.reevescad.org/"
    elif county == "roberts":
        page = "https://esearch.robertscad.org/"
    elif county == "san patricio":
        page = "https://esearch.sanpatcad.org/"
    elif county == "san saba":
        page = "http://esearch.sansabacad.org/"
    elif county == "schleicher":
        page = "https://esearch.schleichercad.org/"
    elif county == "shackelford":
        page = "https://esearch.shackelfordcad.com/"
    elif county == "somervell":
        page = "https://esearch.somervellcad.net/"
    elif county == "throckmorton":
        page = "https://esearch.throckmortoncad.org/"
    elif county == "titus":
        page = "https://esearch.titus-cad.org/"
    elif county == "van zandt":
        page = "https://esearch.vzcad.org/"
    elif county == "walker":
        page = "https://walkercad.org/property-search/"
    elif county == "wharton":
        page = "http://esearch.whartoncad.net/"
    elif county == "wood":
        page = "https://esearch.woodcad.net/"
    else:
        page = "https://esearch.zavalacad.com/"
    return page

################################################################################
# getPropertyPagePre(county): finds the preamble of the property page URL for
# the county selected by the user
# county: county name input by user
# return: the preamble of the property page URL as a string
################################################################################
def getPropertyPagePre(county):
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
    else:
        propURLPre = "https://esearch.zavalacad.com/Property/View/" # Zavala County
    return propURLPre

################################################################################
# getSearchType(county): asks the user for which type of search he would like
# to perform (Regular, By owner, By address, By ID, or Advanced)
# county: county name input by user
# return: searchType
################################################################################
def getSearchType(county):
    print("Choose a search type (1, 2, 3, 4, or 5).")
    searchType = str(input("1.Regular search\n2.By owner\n3.By address\n4.By ID\n5.Advanced\n"))
    return searchType
    

################################################################################
# getSearchArgument(searchType): gets the string to add on to the search URL
# in order to do the search
# searchType: type of search to be conducted (Regular, By owner, By address,
# By ID, or Advanced)
# return: preamble variable which contains the string of the search criteria
################################################################################
def getSearchArgument(searchType):
    preamble = "Search/Result?keywords="
    if searchType == "1": # regular search
        query = str(input("Enter the search query: "))
        preamble = appendArgument(preamble, query)
        return preamble
    if searchType == "2": # by owner search: ownerName, ptype, DBA, taxYear
        preamble = preamble + "OwnerName:"
        ownerName = str(input("Enter owner name: "))
        preamble = appendArgument(preamble, ownerName)
        preamble = appendPropertyType(preamble)
        dbaQ = str(input("Enter a DBA name? y/n: "))
        dbaQ = dbaQ.lower()
        if dbaQ == "y":
            DBA = str(input("Enter a DBA name: "))
            preamble = preamble + "%20DoingBusinessAs:"
            preamble = appendArgument(preamble, DBA)
        tYearQ = str(input("Enter a tax year? (default is current year) y/n: "))
        tYearQ = tYearQ.lower()
        if tYearQ == "y":
            taxYear = str(input("Enter a tax year: "))
            preamble = preamble + "%20Year:"
            preamble = preamble + taxYear
    if searchType == "3":# by address search:streetNum, pType, streetName, taxYear
        preamble = preamble + "StreetNumber:"
        streetNumber = str(input("Enter street number: "))
        streetName = str(input("Enter the street name: "))
        preamble = preamble + streetNumber
        preamble = appendPropertyType(preamble)
        preamble = preamble + "%20StreetName:"
        preamble = appendArgument(preamble, streetName)
        tYearQ = str(input("Enter a tax year? (default is current year) y/n: "))
        tYearQ = tYearQ.lower()
        if tYearQ == "y":
            taxYear = str(input("Enter a tax year: "))
            preamble = preamble + "%20Year:"
            preamble = preamble + taxYear
    if searchType == "4":# by ID:quickRefID, pType, taxYear
        preamble = preamble + "PropertyId:"
        quickRefID = str(input("Enter a quick reference ID: "))
        preamble = preamble + quickRefID
        preamble = appendPropertyType(preamble)
        tYearQ = str(input("Enter a tax year? (default is current year) y/n: "))
        tYearQ = tYearQ.lower()
        if tYearQ == "y":
            taxYear = str(input("Enter a tax year: "))
            preamble = preamble + "%20Year:"
            preamble = preamble + taxYear
    if searchType == "5":
        ownerQ = str(input("Enter an owner name? y/n: ")) #add owner name
        ownerQ = ownerQ.lower()
        if ownerQ == "y":
            preamble = preamble + "OwnerName:"
            ownerName = str(input("Enter owner name: "))
            preamble = appendArgument(preamble, ownerName)
        addressQ = str(input("Enter address? y/n: ")) #add address
        addressQ = addressQ.lower()
        if addressQ == "y":
            streetNumber = str(input("Enter street number: "))
            streetName = str(input("Enter street name: "))
            if preamble[-1] != "=":
                preamble = preamble + "%20"
            preamble = preamble + "StreetNumber:"
            preamble = preamble + streetNumber
        idQ = str(input("Enter quick reference ID? y/n: ")) #add quick ref ID
        idQ = idQ.lower()
        if idQ == "y":
            if preamble[-1] != "=":
                preamble = preamble + "%20"
            preamble = preamble + "PropertyId:"
            quickRefID = str(input("Enter quick reference ID: "))
            preamble = preamble + quickRefID  
        pTypeQ = str(input("Choose a property type? y/n: ")) #add property type
        pTypeQ = pTypeQ.lower()
        if pTypeQ == "y":
            if preamble[-1] != "=":
                preamble = preamble + "%20"
            print("Choose a property type (1, 2, 3, 4, or 5)")
            pType = int(input("1.Real\n2.Personal\n3.Mineral\n4.Auto\n5.Mobile Home\n"))
            if pType == 1:
                preamble = preamble + 'PropertyType:Real'
            if pType == 2:
                preamble = preamble + 'PropertyType:Personal'
            if pType == 3:
                preamble = preamble + 'PropertyType:Mineral'
            if pType == 4:
                preamble = preamble + 'PropertyType:Auto'
            if pType == 5:
                preamble = preamble + 'PropertyType:"Mobile%20Home"'
        dbaQ = str(input("Enter a DBA name? y/n: ")) #add DBA
        dbaQ = dbaQ.lower()
        if dbaQ == "y":
            if preamble[-1] != "=":
                preamble = preamble + "%20"
            DBA = str(input("Enter a DBA name: "))
            preamble = preamble + "DoingBusinessAs:"
            preamble = appendArgument(preamble, DBA)
        if addressQ == "y":
            preamble = preamble + "%20StreetName:"
            preamble = appendArgument(preamble, streetName)
        tYearQ = str(input("Enter a tax year? (default is current year) y/n: ")) #add tax year
        tYearQ = tYearQ.lower()
        if tYearQ == "y":
            if preamble[-1] != "=":
                preamble = preamble + "%20"
            taxYear = str(input("Enter a tax year: "))
            preamble = preamble + "Year:"
            preamble = preamble + taxYear
            
    return preamble    
    
################################################################################
# doSearch(searchType, county): performs the property search by calling 
# referenceIDSearch() and propertySearch() functions
# searchType: type of search to be conducted (Regular, By owner, By address,
# By ID, or Advanced)
# county: county name input by user
# return: no return
################################################################################
def doSearch(searchType, county):
    referenceIDList = referenceIDSearch(searchType, county)
    if referenceIDList == [0, 0, 0]:
        print("Search program terminating...")
        return
    print("Reference ID compilation successful...")
    print("Property data search commencing...")
    propertySearch(referenceIDList, county)

    

################################################################################
# propertySearch: performs the property search and exports excel files
# containing property data and a list of reference IDs that did not load
# IDlist: list of quick reference IDs
# county: county name input by user
# return: no return
################################################################################    
def propertySearch(IDlist, county):
    URL = getPropertyPagePre(county)
    noLoadList = []
    noLoadList.append("Quick Reference ID")
    propertyList = []
    
    for j in range(0, len(IDlist)):
        propertyURL = URL + IDlist[j]
        propertyDriver = webdriver.Chrome(options=options)
        propertyDriver.set_page_load_timeout(120)
        try:
            propertyDriver.get(propertyURL)
            time.sleep(2)
            ##########property details box on website##########
            quickRefID = IDlist[j]
            webpageHyperlink = '=HYPERLINK("' + propertyURL + '")'
            try:
                legalDescription = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]'\
                                                                        +'/div[1]/div/table/tbody/tr[3]/td').text
            except exceptions.NoSuchElementException:
                legalDescription = ""
            try:
                geoID = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]/div[1]'\
                                                             +'/div/table/tbody/tr[4]/td').text
            except exceptions.NoSuchElementException:
                geoID = ""
            try:
                agent = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]/div[1]'\
                                                             +'/div/table/tbody/tr[5]/td').text
            except exceptions.NoSuchElementException:
                agent = ""
            try:
                propertyType = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]'\
                                                                    +'/div[1]/div/table/tbody/tr[6]/td').text
            except exceptions.NoSuchElementException:
                propertyType = ""
            try:
                propertyAddress = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]'\
                                                                       +'/div[1]/div/table/tbody/tr[8]/td').text
            except exceptions.NoSuchElementException:
                propertyAddress = ""
            try:
                mapID = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]/div[1]'\
                                                             +'/div/table/tbody/tr[9]/td').text
            except exceptions.NoSuchElementException:
                mapID = ""
            try:
                neighborhoodCode = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]'\
                                                                        +'/div[1]/div/table/tbody/tr[10]/td').text
            except exceptions.NoSuchElementException:
                neighborhoodCode = ""
            try:
                ownerID = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]/div[1]'\
                                                               +'/div/table/tbody/tr[12]/td').text
            except exceptions.NoSuchElementException:
                ownerID = ""
            try:
                ownerName = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]'\
                                                                 +'/div[1]/div/table/tbody/tr[13]/td').text
            except exceptions.NoSuchElementException:
                ownerName = ""
            try:
                ownerAddress = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]'\
                                                                    +'/div[1]/div/table/tbody/tr[14]/td').text
            except exceptions.NoSuchElementException:
                ownerAddress = ""
            try:
                percentOwnership = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]'\
                                                                        +'/div[1]/div/table/tbody/tr[15]/td').text
            except exceptions.NoSuchElementException:
                percentOwnership = ""
            try:
                taxExemptions = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]'\
                                                                     +'/div[1]/div/table/tbody/tr[16]/td/span').text
            except exceptions.NoSuchElementException:
                taxExemptions = ""
            ##########property values box on website##########
            try:
                improvementHomesiteVal = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]'\
                                                                              +'/div[3]/div[2]/div[1]'\
                                                                              +'/table/tbody/tr[2]/td').text
            except exceptions.NoSuchElementException:
                improvementHomesiteVal = ""
            try:
                improvementNonHomesiteVal = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]'\
                                                                                 +'/div[3]/div[2]/div[1]'\
                                                                                 +'/table/tbody/tr[3]/td').text
            except exceptions.NoSuchElementException:
                improvementNonHomesiteVal = ""
            try:
                landHomesiteVal = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]'\
                                                                       +'/div[2]/div[1]/table/tbody/tr[4]/td').text
            except exceptions.NoSuchElementException:
                landHomesiteVal = ""
            try:
                landNonHomesiteVal = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]'\
                                                                          +'/div[2]/div[1]/table/tbody/tr[5]/td').text
            except exceptions.NoSuchElementException:
                landNonHomesiteVal = ""
            try:
                agMarketVal = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]'\
                                                                   +'/div[2]/div[1]/table/tbody/tr[6]/td').text
            except exceptions.NoSuchElementException:
                agMarketVal = ""
            try:
                valueMethod = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]'\
                                                                   +'/div[2]/div[1]/table/tbody/tr[8]/td').text
            except exceptions.NoSuchElementException:
                valueMethod = ""
            try:
                marketVal = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]'\
                                                                 +'/div[2]/div[1]/table/tbody/tr[9]/td').text
            except exceptions.NoSuchElementException:
                marketVal = ""
            try:
                agUseVal = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]'\
                                                                +'/div[2]/div[1]/table/tbody/tr[10]/td').text
            except exceptions.NoSuchElementException:
                agUseVal = ""
            try:
                appraisedVal = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]'\
                                                                    +'/div[2]/div[1]/table/tbody/tr[12]/td').text
            except exceptions.NoSuchElementException:
                appraisedVal = ""
            try:
                homesteadCapLoss = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]'\
                                                                        +'/div[2]/div[1]/table/tbody/tr[13]/td').text
            except exceptions.NoSuchElementException:
                homesteadCapLoss = ""
            try:
                assessedVal = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[3]'\
                                                                   +'/div[2]/div[1]/table/tbody/tr[15]/td').text
            except exceptions.NoSuchElementException:
                assessedVal = ""
            ##########property land box on website##########
            try:
                propertyLandType = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[6]'\
                                                                        +'/div[2]/table/tbody/tr[2]/td[1]').text
            except exceptions.NoSuchElementException:
                propertyLandType = ""
            try:
                landTypeDescription = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]'\
                                                                           +'/div[6]/div[2]/table/tbody'\
                                                                           +'/tr[2]/td[2]').text
            except exceptions.NoSuchElementException:
                landTypeDescription = ""
            try:
                acreage = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[6]'\
                                                               +'/div[2]/table/tbody/tr[2]/td[3]').text
            except exceptions.NoSuchElementException:
                acreage = ""
            try:
                landSquareFeet = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]'\
                                                                      +'/div[6]/div[2]/table/tbody/tr[2]/td[4]').text
            except exceptions.NoSuchElementException:
                landSquareFeet = ""
            try:
                effectiveFrontage = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]'\
                                                                         +'/div[6]/div[2]/table/tbody/tr[2]/td[5]').text
            except exceptions.NoSuchElementException:
                effectiveFrontage = ""
            try:
                effectiveDepth = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]'\
                                                                      +'/div[6]/div[2]/table/tbody/tr[2]/td[6]').text
            except exceptions.NoSuchElementException:
                effectiveDepth = ""
            try:
                landMarketVal = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]'\
                                                                     +'/div[6]/div[2]/table/tbody/tr[2]/td[7]').text
            except exceptions.NoSuchElementException:
                landMarketVal = ""
            try:
                landProductionVal = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]'\
                                                                         +'/div[6]/div[2]/table/tbody/tr[2]/td[8]').text
            except exceptions.NoSuchElementException:
                landProductionVal = ""
            ##########property deed history box on website##########
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
                    dDate = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[8]'\
                                                                 +'/div[2]/table/tbody/tr[' + str(k) + ']/td[1]').text
                    deedDate.append(dDate)
                except exceptions.NoSuchElementException:
                    dDate = ""
                    deedDate.append(dDate)
                try:
                    dType = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[8]'\
                                                                 +'/div[2]/table/tbody/tr[' + str(k) + ']/td[2]').text
                    deedType.append(dType)
                except exceptions.NoSuchElementException:
                    dType = ""
                    deedType.append(dType)
                try:
                    dDesc = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[8]'\
                                                                 +'/div[2]/table/tbody/tr[' + str(k) + ']/td[3]').text
                    documentDesc.append(dDesc)
                except exceptions.NoSuchElementException:
                    dDesc = ""
                    documentDesc.append(dDesc)
                try:
                    dGrantor = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[8]'\
                                                                    +'/div[2]/table/tbody/tr[' + str(k) + ']/td[4]').text
                    deedGrantor.append(dGrantor)
                except exceptions.NoSuchElementException:
                    dGrantor = ""
                    deedGrantor.append(dGrantor)
                try:
                    dGrantee = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[8]'\
                                                                    +'/div[2]/table/tbody/tr[' + str(k) + ']/td[5]').text
                    deedGrantee.append(dGrantee)
                except exceptions.NoSuchElementException:
                    dGrantee = ""
                    deedGrantee.append(dGrantee)
                try:
                    dVolume = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[8]'\
                                                                   +'/div[2]/table/tbody/tr[' + str(k) + ']/td[6]').text
                    deedVolume.append(dVolume)
                except exceptions.NoSuchElementException:
                    dVolume = ""
                    deedVolume.append(dVolume)
                try:
                    dPage = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[8]'\
                                                                 +'/div[2]/table/tbody/tr[' + str(k) + ']/td[7]').text
                    deedPage.append(dPage)
                except exceptions.NoSuchElementException:
                    dPage = ""
                    deedPage.append(dPage)
                try:
                    dNumber = propertyDriver.find_element_by_xpath('//*[@id="detail-page"]/div[8]'\
                                                                   +'/div[2]/table/tbody/tr[' + str(k) + ']/td[8]').text
                    deedNumber.append(dNumber)
                except exceptions.NoSuchElementException:
                    dNumber = ""
                    deedNumber.append(dNumber)
            if deedDate[1] == "" and deedDate[2] == "":
                moreDeed = "False"

            ##########property data list##########
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


            propertyList.append(propertyData)
            propertyDriver.close()

            #######################################################
            #       write to master excel file  
            #######################################################
            df = pd.DataFrame(propertyList)
            outputFile = county + "CADSearch.xlsx"
            df.to_excel(outputFile)
            print("Property " + str(j+1) + " of " + str(len(IDlist)) + " completed.") 
        except exceptions.TimeoutException:
            print("There was an error with " + str(IDlist[j])\
                  + ". The page timed out.")
            print("Property " + str(j+1) + " of " + str(len(IDlist)) + " completed unsuccessfully.")
            propertyDriver.close()
            #######################################################
            #       write to non load excel file  
            #######################################################
            noLoadList.append(IDlist[j])
            df = pd.DataFrame(noLoadList)
            noLoadFile = county + "NonLoadList.xlsx"
            df.to_excel(noLoadFile)
            
################################################################################
# referenceIDSearch(county): creates a list of reference IDs that corresponds
# to the CAD search conducted by the user
# county: county name input by user
# return: quickRefIDs; list of quick reference IDs
################################################################################        
def referenceIDSearch(searchType, county):
    quickRefIDs = []
    URL = getSearchPage(county)
    searchArg = getSearchArgument(searchType)
#     print("Search Arg")
#     print(searchArg)
    URL = URL + searchArg
    URL = URL + "&pageSize=100"
    print("Initial search page URL")
    print(URL)
    statusDriver = webdriver.Chrome(options=options)
    statusDriver.set_page_load_timeout(120)
    try:
        statusDriver.get(URL)
        time.sleep(3)
        header = statusDriver.find_element_by_xpath('//*[@id="page-header"]').text   
        header = header.split()
        try:
            lastPage = int(header[3])
        except ValueError:
            print("There were no search results.")
            quickRefIDs = [0, 0, 0]
            return quickRefIDs
        totalProperties = int(header[6])
        if lastPage == 1:
            print("There is 1 page and " + str(totalProperties) + " properties in your search results")
        else:
            print("There are " + str(lastPage) + " pages and " + str(totalProperties) + " properties in your search results")
        statusDriver.close()
    except exceptions.TimeoutException:
        print("The search page would not load.")
        statusDriver.close()
        tryAgain = str(input("Try again? y/n: "))
        tryAgain = tryAgain.lower()
        if tryAgain == "y":
            doSearch(county)
        else:
            quickRefIDs = [0, 0, 0]
            return quickRefIDs
        
    pageStart = int(input("Enter the index of the page you'd like to start the search (1, 2, etc.): "))
    print("Starting search on page ", pageStart)
    for p in range(pageStart, lastPage + 1):
        refIDdriver = webdriver.Chrome(options=options)
        refIDdriver.set_page_load_timeout(120)
        if p > 1:
            searchURL = URL + "&page=" + str(p)
        else:
            searchURL = URL
        try:
            refIDdriver.get(searchURL)
            for i in range(1, 101):
                try:
                    referenceID = refIDdriver.find_element_by_xpath('//*[@id="grid"]/div[2]/table/tbody/tr['\
                                                               + str(i) + ']/td[2]').text
                    quickRefIDs.append(referenceID)
                except exceptions.NoSuchElementException:
                    break
            refIDdriver.close()
        except exceptions.TimeoutException:
            print("There was a timeout error with page ", p, " of ", lastPage)
            refIDdriver.close()
    
    return quickRefIDs

################################################################################
# appendArgument(pre, arg): appends a new search criteria to the search string
# pre: the existing search string
# arg: the argument to append to the search string
# return: pre; the search string
################################################################################   
def appendArgument(pre, arg):
    space = " "
    if space in arg:
        arg = arg.lower()
        arg.split(" ")
        pre = pre + '"'
        for piece in arg:
            pre = pre + piece
            pre = pre + "%20"
        pre.rstrip('0')
        pre.rstrip('2')
        pre.rstrip('%')
        pre = pre + '"'
    else:
        pre = pre + arg
            
    return pre

################################################################################
# appendPropertyType(pre): appends the property type criteria to the search string
# pre: the existing search string
# return: pre; the search string
################################################################################
def appendPropertyType(pre):
    pTypeQ = str(input("Choose a property type? y/n: "))
    pTypeQ = pTypeQ.lower()
    if pTypeQ == "y":
        print("Choose a property type (1, 2, 3, 4, or 5)")
        pType = str(input("1.Real\n2.Personal\n3.Mineral\n4.Auto\n5.Mobile Home\n"))
        if pType == "1":
            pre = pre + '%20PropertyType:Real'
        if pType == "2":
            pre = pre + '%20PropertyType:Personal'
        if pType == "3":
            pre = pre + '%20PropertyType:Mineral'
        if pType == "4":
            pre = pre + '%20PropertyType:Auto'
        if pType == "5":
            pre = pre + '%20PropertyType:"Mobile%20Home"'
                
    return pre

################################################################################
# main program  
################################################################################
inputLoop(county)
searchType = getSearchType(county)
doSearch(searchType, county)
print("Program Complete")