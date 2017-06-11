import pandas as pd
import numpy as np
import csv
from xlutils.copy import copy
import xlsxwriter
import xlrd
import xlwt
import os
import scipy.stats as st
import plotly.figure_factory as FF
import csv 
import re
import datetime
import scipy.stats as sps
from datetime import datetime, timedelta
from calendar import monthrange





def doCountryDataCheck(threshold,s):

        dir = r"D:\\Style\\2017-05\\Book country Files Slimline\\Datachecks\\"  

        country = []

        if (s == "ICB"):
            country = ['AUSTRALIA ICB', 'BELGIUM ICB', 'BRAZIL ICB', 
                   'CANADA ICB', 'CHINA ICB', 'DENMARK ICB', 'FRANCE ICB',
                   'GERMANY ICB', 'HONG KONG ICB', 'INDIA ICB', 'IRELAND ICB', 'ITALY ICB', 
                   'JAPAN ICB', 'NETHERLANDS ICB', 'NORWAY ICB', 'RUSSIA ICB', 'SINGAPORE ICB', 
                   'SOUTH AFRICA ICB', 'SPAIN ICB', 'SWEDEN ICB', 'SWITZERLAND ICB', 
                   'UNITED KINGDOM ICB', 'UNITED STATES ICB']
        else:     
            country  = ['AUSTRALIA GICS', 'BELGIUM GICS', 'BRAZIL GICS', 
                   'CANADA GICS', 'CHINA GICS', 'DENMARK GICS', 'FRANCE GICS',
                   'GERMANY GICS', 'HONG KONG GICS', 'INDIA GICS', 'IRELAND GICS', 'ITALY GICS', 
                   'JAPAN GICS', 'NETHERLANDS GICS', 'NORWAY GICS', 'RUSSIA GICS', 'SINGAPORE GICS', 
                   'SOUTH AFRICA GICS', 'SPAIN GICS', 'SWEDEN GICS', 'SWITZERLAND GICS', 
                   'UNITED KINGDOM GICS', 'UNITED STATES GICS']

        for c in country:
  
            file = dir + "\\DataCheck " + c + ".xlsx"
        
            book=xlrd.open_workbook(file)                         
            sheet=book.sheet_by_name('Sheet1')

            print("************")
            print(c)
            print("************")
            for j in range(0,495,14):
            
                k=0
                if (j>238): 
                    k=j+4
                else:
                    k=j

      
                previous = []
                currentReb = []                              
                for i in range(19,132):                  

                    previous.append(sheet.cell_value(27+k,i))
                    currentReb.append(sheet.cell_value(29+k,i))
                          
                
                chi2 = sps.chisquare(previous,currentReb)              
                 
                if (chi2[0] > threshold):
                    print(sheet.cell_value(26+k,0)) #name of plot                                
                   # print(chi2)




def doRegionalDataCheck():

    dir = r"D:\\Style\\2017-05\\Country EconomicRegion Slimline2\\"

#    region = ['All Emerging ICB', 'AW Developed ICB', 'BRIC ICB', 'Euro Zone ICB']

    region = ['All Emerging ICB']

    for f in region:
  
        file = dir + "\\" + f +"\\DataCheck " + f + ".xlsx"
        
        previous = []
        currentReb = []

        print(f) 

        for j in range (0,504,14):
          
            for i in range(19,32):                  
                book=xlrd.open_workbook(file)                         
                sheet=book.sheet_by_name('Sheet1')
                previous.append(sheet.cell_value(27+j,i))
                currentReb.append(sheet.cell_value(29+j,i))
                

            print(sheet.cell_value(26+j,0)) #name of plot
            print(previous)
            print(currentReb)

            
            chi2 = sps.chisquare(previous,currentReb)              
           
            print(chi2)
