#------------- IMPORT STATEMENTS -------------------------------# 
from mmap import mmap, ACCESS_READ                                                       
from xlrd import open_workbook, cellname #to read the excel file
from tempfile import TemporaryFile
from xlwt import Workbook
import mechanize
import re
from bs4 import BeautifulSoup, SoupStrainer
from urllib2 import urlopen
from selenium import webdriver

#-------------LOAD EXCEL FILE/WEBSITE HERE -------------------------------#
terms_collection = Workbook(encoding='ascii') #created the workbook
terms_sheet = terms_collection.add_sheet('Terms') #add the sheet

komen = open_workbook('topiccodes2.xls') #loads into komen test
#print komen.nsheets 
sheet = komen.sheet_by_index(0)
#print sheet.nrows 
#print sheet.ncols
b = 0
wr_cnt = 0


#-------------READ FROM EXCEL FILE HERE -------------------------------# 

#for row_index in range(sheet.nrows) :
    #for col_index in range(sheet.ncols)
       
       # print sheet.cell(row_index, col_index).value


#--------------BEGIN LOOP\READS ALL Values in Description Colum---------------------
for row_index in range(sheet.nrows) :

    x = [] #array for term grabbing
    terms = [] #array for terms
    new_t = []
    #-----------reads the excel sheet and uploads info-------------#
    a = sheet.cell(row_index, 2).value + sheet.cell(row_index, 1).value
    code = sheet.cell(row_index, 0).value
    vacr = a #stores the value in 'a' because of variable change later
    a = a.encode('utf-8')
    print a
    print b
    b += 1
    #-------------------searches for info online----------#
    url = "http://www.nlm.nih.gov/mesh/MeSHonDemand.html"
    br = mechanize.Browser()
    br.set_handle_robots(False) # ignore robots
    br.open(url)
    br.select_form(name="MTIForm") #selects the area to input it
    br["InputText"] = a
    res = br.submit()
    content = res.read() #holds html 
    #print content
    
    #----------------------------scans the html and grabs terms---------------------------#
    soup = BeautifulSoup(content)
    
    
    for a in soup.find_all('a', href=True): #bad programming practice, changed variable purpose!!!
      #  print a['href']
        if 'http://www.nlm.nih.gov/cgi/mesh/2015/MB_cgi?term=' in a['href']:
            x.append(a['href'])

   # print x

    for o in range(1, len(x), 2):

       print x[o]
       terms.append(x[o][x[o].rfind("=") +1:x[o].rfind("")])
       
       

    print terms

    for i in terms:
        print i

    for y in terms:
        terms_sheet.write(wr_cnt, 2, y)
        terms_sheet.write(wr_cnt, 3, wr_cnt)
        terms_sheet.write(wr_cnt, 1, vacr)
        terms_sheet.write(wr_cnt, 0, code)
        wr_cnt += 1
        
    
    # terms_sheet.write(row_index, 2, terms)
    
   # for z in range(len(terms)):
     #        terms_sheet.write(z, 2, terms[z])
     #        terms_sheet.write(z, 1, z)#row_index) 
    #----------------------------puts terms into excel sheet---------------------------#


#print terms
terms_collection.save('terms.xls')

    

    
     


    # SEARCH FOR A ON MeSH WebSite #

    

        

#STEP 1: BEGIN LOOP/READ FROM EXCEL SHEET


# for number of rows in column from sheet whatever






#STEP 2: SEARCH MESH WEBPAGE AND GRAB TERMS



#url = "http://www.nlm.nih.gov/mesh/MeSHonDemand.html"
#br = mechanize.Browser()
#br.set_handle_robots(False) # ignore robots
#br.open(url)
#br.select_form(name="MTIForm") #selects the area to input it
#br["InputText"] = "Endostatin, a new protein recently discovered in our laboratory is the most powerful of the known inhibitors of blood vessel growth. We propose to determine (1) if it can inhibit human breast cancer in mice; (2) the mechanism of its ability to eradicate the tumors when administered with another angiogenesis inhibitor angistatin, and (3) the molecular mechanism of endostatin."  #inputs into text area (WOOH! FIGURED IT OUT!)
#res = br.submit()
#content = res.read()
#print content




#STEP 3: STORE TERMS & ITERATE/END LOOP
