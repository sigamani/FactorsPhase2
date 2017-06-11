import plotly.plotly as py
import plotly.graph_objs as go
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
import module1 as m1
from datetime import datetime, timedelta
from calendar import monthrange



def getCorrelationOfReturns(country):

    dir = r"C:\\Users\\michael\\Desktop\\SRMA2\\"

    factors = ['CAEY', 'InversePEG', 'InversePEGY', 'Net Buyback Yld', 'Net Dbt Paydown Yld',
               'Net Payout Yld', 'Asset Growth 3yr', 'Asset Growth 5yr',
               'Capex Growth 3yr', 'Capex Growth 5yr', 'Historic Dividend Growth', 
               'IBES ROE','IBES 12M Fcast R 3M', 'IBES 12M Fcast R 1M', 
               'Stable Ergs Gr 5yr', 'Stable Sales Gr 5yr',
               'Daily 1yr Vol', 'Mmntm 12-1','Trading Turnover 3mth',
               'Interest Coverage Ratio (ex-Fin)', 'Current Ratio (ex-Fin)', 'Quick Ratio (ex-Fin)']


          
    nfactors = len(factors)        
    arr = []

    for j in range(nfactors):    

        factor = factors[j].replace(" (*)","")
        file = dir + "Michael Style 50-100 " + factor + " " + country + " MC M1 200705 to 201704 Dec.xlsx"


        book=xlrd.open_workbook(file)                         
        sheet=book.sheet_by_name('Return Data')

        list = []
        for i in range(5,123):
            value = sheet.cell_value(12,i)
            list.append(value)

        arr.append([])
        arr[j].append(list)
    

    results = []        
    for k in range(nfactors):

        results.append([])
        for l in range(nfactors):
                    
            corr = np.corrcoef(arr[k],arr[l])
            temp = corr[0,1]
            temp2 = round(temp,2)
            results[k].append(temp2)


    fig = FF.create_annotated_heatmap(x = factors, y = factors, z=results)
    layout = go.Layout(
    title=country,
    xaxis=dict(ticks='', ticksuffix='', side='bottom'),
        width=350,
        height=300,
        margin=go.Margin(
        l=180,
        r=80,
        b=100,
        t=100,
        pad=4
        ),
    autosize=False          
    )


    py.iplot(fig, filename=country+'-Correlation',layout=layout)


def getMarketsAnalyzerOutput(s):


    dir = r"C:\\Users\\michael\\Desktop\\SRMA2\\"

    factors = ['CAEY', 'InversePEG', 'InversePEGY', 'Net Buyback Yld', 'Net Dbt Paydown Yld',
               'Net Payout Yld', 'Asset Growth 3yr', 'Asset Growth 5yr',
               'Capex Growth 3yr', 'Capex Growth 5yr', 'Historic Dividend Growth', 
               'IBES ROE','IBES 12M Fcast R 3M', 'IBES 12M Fcast R 1M', 
               'Stable Ergs Gr 5yr', 'Stable Sales Gr 5yr',
               'Daily 1yr Vol', 'Mmntm 12-1','Trading Turnover 3mth',
               'Interest Coverage Ratio (ex-Fin)', 'Current Ratio (ex-Fin)', 'Quick Ratio (ex-Fin)']

    countries = ['UNITED STATES','CHINA', 'JAPAN', 'UNITED KINGDOM','SWITZERLAND', 'GERMANY', 'HONG KONG',
                 'FRANCE', 'CANADA', 'AUSTRALIA', 'NETHERLANDS', 'SOUTH AFRICA', 'SINGAPORE']

    #countries = ['UNITED STATES','CHINA']

    if (s == 'Return_10Year'):
        plottitle = '10 Year Return (%)'
        y=21
        x=4
        rounddp=0
        percent=100
    if (s == 'Return_1Year'):
        plottitle = '12 Month Return (%)'
        y=22
        x=4
        rounddp=0
        percent=100
    if (s == 'TrackingError'):
        plottitle = 'Tracking Error (%)'
        y=21
        x=6
        rounddp=0
        percent=100
    if (s == 'TrackingError_2Year'):
        plottitle = '2 Year Tracking Error (%)'
        y=22
        x=6
        rounddp=0
        percent=100
    if (s == 'StdDev'):
        plottitle = 'Standard Deviation'
        y=23
        x=6
        rounddp=3
        percent=1
    if (s == 'StdDev_2Year'):
        plottitle = '2 Year Standard Deviation'
        y=24
        x=6
        rounddp=3
        percent=1
    if (s == 'StyleBeta'):
        plottitle = 'Beta'
        y=25
        x=6
        rounddp=2
        percent=1
    if (s == 'Regularity_3Month'):
        plottitle = '3 Month Regularity'
        y=21
        x=8
        rounddp=3
        percent=1
    if (s == 'Regularity_6Month'):
        plottitle = '6 Month Regularity'
        y=22
        x=8
        rounddp=3
        percent=1
    if (s == 'Regularity_12Month'):
        plottitle = '12 Month Regularity'
        y=23
        x=8
        rounddp=3
        percent=1
    if (s == 'Identity'):
        plottitle = 'Identity (%)'
        y=25
        x=2
        rounddp=0
        percent=100
    if (s == 'Attrib'):
        plottitle = 'Attribution'
        y=23
        x=2
        rounddp=2
        percent=1


    corr = [] 
    for i in range(len(factors)):

            corr.append([])
            for coun in countries:

                file = dir + "Michael Style 50-100 " + factors[i] + " " + coun + " MC M1 200705 to 201704 Dec.xlsx"
                
                book=xlrd.open_workbook(file)                         
                sheet=book.sheet_by_name('Style Graph')
                value = sheet.cell_value(y,x)
                corr[i].append(value*percent)
                 
   

    #layout = go.Layout(
    #title=s,
    #xaxis=dict(ticks='', ticksuffix='', side='bottom'),
    #width=450,
    #height=300,
    #    margin=go.Margin(
    #    l=180,
    #    r=50,
    #    b=100,
    #    t=100,
    #    pad=4
    #    ),
    #autosize=False          
    #)

    fig = go.Heatmap(x = countries, y = factors, z=np.round(corr,rounddp))
    py.plot([fig], filename=s)


def heatmap():

    trace = go.Heatmap(z=[[1, 20, 30],
                      [20, 1, 60],
                      [30, 60, 1]])
    data=[trace]
    py.iplot(data, filename='basic-heatmap')


def changeName(dir):

    country = ['CHINA', 'JAPAN', 'SWITZERLAND', 'GERMANY', 'HONG KONG', 'UNITED STATES', 'UNITED KINGDOM',
                 'FRANCE', 'CANADA', 'AUSTRALIA', 'NETHERLANDS', 'SOUTH AFRICA', 'SINGAPORE']

    for c in country:

        for filename in os.listdir(dir):
        
            if (c in filename):

                temp = filename.split(c)[-1].split('MC')[0]
                filename2= filename.replace(temp,' ')
               # print("move \"" +dir+filename + "\" \"" +dir+ filename2+"\"")
                os.system("move \"" +dir+filename + "\" \"" +dir+ filename2+"\"")


def doCoveragePlot():

    file = r"C:\\Users\\michael\\Documents\\coverage.csv"   
    df = pd.read_csv(file)
    df['DataDate'] = pd.to_datetime(df['DataDate'])
    data = []


    factorIDs = [117,118,119,107,109,110,120,121,122,123,
                 124,125,126,127,128,129,130,131,132,133,
                 134,135,10103,116]

    for f in factorIDs: 
            
        criteria = (df['FactorID']==f)
        factorname = df[criteria].iat[0,1]       

        trace = go.Scatter(
            x = df[criteria].DataDate,
            y = df[criteria].WgtCoverage,
            name=factorname
        )
   
        data.append(trace)

    layout = go.Layout(
    title="Coverage",
    xaxis = dict( )            
    )

    fig = go.Figure(data = data, layout=layout)
    py.iplot(fig, filename="Coverage")


def makePercentiles():

    file = r"C:\\Users\\michael\\Documents\\percentiles2.csv"   
    df = pd.read_csv(file)
    df['DataDate'] = pd.to_datetime(df['DataDate'])

    factorIDs = [117,118,119,107,109,110,120,121,122,123,
                 124,125,126,127,128,129,130,131,132,133,
                 134,135]

    factorIDs = [132] 

    percentiles = [0.01, 0.02, 0.05, 0.1, 0.25, 
                   0.5, 0.75, 0.9, 0.95, 0.98, 0.99]
    data = []     
    factorname = ""
 
    for f in factorIDs:

        for p in percentiles:
       
            criteria = ((df['FactorID']==f) & (df['Percentile']==p))
            #criteria = ((df['FactorID']==f) & (df['Percentile']==p) & (df['DataDate'] > '1990-01-01'))

            trace = go.Scatter(
                x = df[criteria].DataDate,
                y = df[criteria].Value,
                name=p
            )
   
            data.append(trace)
            factorname = df[criteria].iat[0,1]       

        layout = go.Layout(
            title=factorname,
            xaxis = dict( )            
        )


        print("Made (%i): %s" %  (f,factorname))

        fig = go.Figure(data = data, layout=layout)
        py.iplot(fig, filename=factorname)
        data = []

def makeTrendingStat(s):


    dir = r"C:\\Users\\michael\\Desktop\\SRMA2\\"

    factors = ['CAEY', 'InversePEG', 'InversePEGY', 'Net Buyback Yld', 'Net Dbt Paydown Yld',
               'Net Payout Yld', 'Asset Growth 3yr', 'Asset Growth 5yr',
               'Capex Growth 3yr', 'Capex Growth 5yr', 'Historic Dividend Growth', 
               'IBES ROE','IBES 12M Fcast R 3M', 'IBES 12M Fcast R 1M', 
               'Stable Ergs Gr 5yr', 'Stable Sales Gr 5yr',
               'Daily 1yr Vol', 'Mmntm 12-1','Trading Turnover 3mth',
               'Interest Coverage Ratio (ex-Fin)', 'Current Ratio (ex-Fin)', 'Quick Ratio (ex-Fin)']


    countries = ['UNITED STATES','CHINA', 'JAPAN', 'UNITED KINGDOM','SWITZERLAND', 'GERMANY', 'HONG KONG',
                 'FRANCE', 'CANADA', 'AUSTRALIA', 'NETHERLANDS', 'SOUTH AFRICA', 'SINGAPORE']

    if (s == 'Regularity_3Month'):
        plottitle = '3 Month Regularity'
        y=21        
        x=8
        sd=0.183
        rounddp=2
        percent=1
    if (s == 'Regularity_6Month'):
        plottitle = '6 Month Regularity'
        y=22
        x=8
        sd=0.289
        rounddp=2
        percent=1
    if (s == 'Regularity_12Month'):
        plottitle = '12 Month Regularity'
        y=23
        x=8
        sd=0.428
        rounddp=2
        percent=1

    result = [] 
    for i in range(len(factors)):

            result.append([])
            for coun in countries:

                factor = factors[i].replace(" (*)","")

                file = dir + "Michael Style 50-100 " + factor + " " + coun + " MC M1 200705 to 201704 Dec.xlsx"
                book=xlrd.open_workbook(file)                         
                sheet=book.sheet_by_name('Style Graph')
                value = sheet.cell_value(y,x)
               
                if (s == 'Regularity_3Month' or s == 'Regularity_6Month' or s == 'Regularity_12Month'):
                    z_score = value/sd
                    p_value = sps.norm.sf(abs(z_score))

                    if value > 0:

                        result[i].append(1 - (2*p_value))    
                    if value < 0:

                        result[i].append(-1+ (2*p_value))
                else:
                    result[i].append(value*percent)


    data = FF.create_annotated_heatmap(x = countries, y = factors, z=np.round(result,rounddp))

    layout = go.Layout(
        autosize=False,
        width=350,
        height=300,
        margin=go.Margin(
            l=1800,
            r=60,
            b=100,
            t=100,
            pad=4
        ),
        paper_bgcolor='#7f7f7f',
        plot_bgcolor='#c7c7c7'
    )

    #fig = go.Figure(data=data, layout=layout)
    #py.iplot(fig, filename=s+'_Trending')
    #test 

    fig = FF.create_annotated_heatmap(x = countries, y = factors, z=np.round(result,rounddp))
    py.iplot(fig, filename=s+'_Trending', layout = layout)

   
def getMarketsAnalyzerOutputAll():


    dir = r"C:\\Users\\michael\\Desktop\\SRMA3\\"

    #stat = ['Return_10Year', 'Return_1Year', 'TrackingError','TrackingError_2Year','StdDev','StdDev_2Year','StyleBeta',
	 #      'Regularity_3Month','Regularity_6Month','Regularity_12Month','Identity']

    stat = ['Identity']
    countries = ['CHINA']
  #  countries = ['UNITED STATES', 'CHINA', 'JAPAN', 'UNITED KINGDOM', 'SWITZERLAND', 'GERMANY', 'HONG KONG',
  #              'FRANCE', 'CANADA', 'AUSTRALIA', 'NETHERLANDS', 'SOUTH AFRICA', 'SINGAPORE']
  

    factors = []
   
    country = "CHINA"
    for filename in os.listdir(dir):
        
        if (country in filename):

            try:
                temp = filename.split('50-100')[-1].split(country)[0]        
                factors.append(temp)
            except:                  
                print("File: %s does not exist" % filename)

    nfactors = len(factors)        
    arr = []
    list_dfs = []

    for s in stat:

        if (s == 'Return_10Year'):
            plottitle = '10 Year Return (%)'
            y=21
            x=4
            rounddp=0
            percent=100
        if (s == 'Return_1Year'):
            plottitle = '12 Month Return (%)'
            y=22
            x=4
            rounddp=0
            percent=100
        if (s == 'TrackingError'):
            plottitle = 'Tracking Error (%)'
            y=21
            x=6
            rounddp=0
            percent=100
        if (s == 'TrackingError_2Year'):
            plottitle = '2 Year Tracking Error (%)'
            y=22
            x=6
            rounddp=0
            percent=100
        if (s == 'StdDev'):
            plottitle = 'Standard Deviation'
            y=23
            x=6
            rounddp=3
            percent=1
        if (s == 'StdDev_2Year'):
            plottitle = '2 Year Standard Deviation'
            y=24
            x=6
            rounddp=3
            percent=1
        if (s == 'StyleBeta'):
            plottitle = 'Beta'
            y=25
            x=6
            rounddp=2
            percent=1
        if (s == 'Regularity_3Month'):
            plottitle = '3 Month Regularity'
            y=21
            x=8
            rounddp=3
            percent=1
        if (s == 'Regularity_6Month'):
            plottitle = '6 Month Regularity'
            y=22
            x=8
            rounddp=3
            percent=1
        if (s == 'Regularity_12Month'):
            plottitle = '12 Month Regularity'
            y=23
            x=8
            rounddp=3
            percent=1
        if (s == 'Identity'):
            plottitle = 'Identity (%)'
            y=25
            x=2
            rounddp=0
            percent=100
        if (s == 'Attrib'):
            plottitle = 'Attribution'
            y=23
            x=2



        val = [] 

        for i in range(len(factors)):

                val.append([])
                for c in countries:

                    file = dir + "Michael Style 50-100" + factors[i] + "" + c + " MC M1 199705 to 201704 Dec.xlsx"
                    print(file)
                    book=xlrd.open_workbook(file)                         
                    sheet=book.sheet_by_name('Style Graph')
                    value = sheet.cell_value(y,x)
                    val[i].append(value*percent)


        dframe = pd.DataFrame(data=val,   
        index=factors,    
        columns=countries)

        list_dfs.append(dframe)       
                        
    writer = pd.ExcelWriter("C:\\Users\\michael\\Google Drive\\FactorsProject\\SRMA10.xlsx")
  
   
    for n, df in enumerate(list_dfs):
        df.to_excel(writer,'%s' % stat[n])
    writer.save()


def getCorrelationOfReturnsAll():

    c = ['UNITED STATES', 'CHINA', 'JAPAN', 'UNITED KINGDOM', 'SWITZERLAND', 'GERMANY', 'HONG KONG',
                 'FRANCE', 'CANADA', 'AUSTRALIA', 'NETHERLANDS', 'SOUTH AFRICA', 'SINGAPORE']

    dir = r"C:\\Users\\michael\\Desktop\\SRMA3\\"
    list_dfs = []

    c = ['UNITED STATES', 'CHINA']
    factors = [" 1Yr Vol ", " AGR GMI "] 

    for country in c:

        factors = []
   
        for filename in os.listdir(dir):
        
            if (country in filename):

                nfactors = len(factors)        
                arr = []
 

        for j in range(nfactors):    



            file = dir + "Michael Style 50-100" + factors[j] + "" + country + " MC M1 199705 to 201704 Dec.xlsx"
            book=xlrd.open_workbook(file)
            first_sheet = book.sheet_by_index(0)

            count = book.nsheets 
            list = []

            if (count ==1):
                print("Missing Return Series for %s and %s." % (factors[j], country))
              
                for i in range(5,123):
                    value = 10E23
                    list.append(value)

                arr.append([])
                arr[j].append(list)

            else:                     
                sheet=book.sheet_by_name('Return Data')
               
                for i in range(5,123):
                    value = sheet.cell_value(12,i)
                    list.append(value)

                arr.append([])
                arr[j].append(list)
    

        results = []        
        for k in range(nfactors):

            results.append([])
            for l in range(nfactors):
                    
                corr = np.corrcoef(arr[k],arr[l])
                temp = corr[0,1]
               # temp2 = round(temp,2)
                results[k].append(temp)

                print("%s, %s, %s, %s, %s, %s, %f" % (factors[l], m1.getStyle(factors[l]), country, factors[k], m1.getStyle(factors[k]), country, temp))


def getStyle(ShortName):
    
    style = "Nothing"
    if (ShortName in " Book Value " ): style = "Value"
    if (ShortName in " Div Yld " ): style = "Value"
    if (ShortName in " Engs Yld " ): style = "Value"
    if (ShortName in " Cfl Yld " ): style = "Value"
    if (ShortName in " Sales to Pr " ): style = "Value"
    if (ShortName in " RoE " ): style = "Growth"
    if (ShortName in " Ergs Gr " ): style = "Growth"
    if (ShortName in " Inc to Sales " ): style = "Growth"
    if (ShortName in " Sales Gr " ): style = "Growth"
    if (ShortName in " IBES 12M Gr " ): style = "Growth"
    if (ShortName in " IBES FY1 R 3M " ): style = "Growth"
    if (ShortName in " Market Cap " ): style = "Risk"
    if (ShortName in " Beta " ): style = "Risk"
    if (ShortName in " Mmntm ST " ): style = "Momentum"
    if (ShortName in " Mmntm 12M " ): style = "Momentum"
    if (ShortName in " Debt to Eq " ): style = "Risk"
    if (ShortName in " Foreign Sales " ): style = "Risk"
    if (ShortName in " IBES FY1 R 1M " ): style = "Growth"
    if (ShortName in " IBES FY2 R 3M " ): style = "Growth"
    if (ShortName in " IBES FY2 R 1M " ): style = "Growth"
    if (ShortName in " EBITDA to Pr " ): style = "Value"
    if (ShortName in " Stable Ergs Gr " ): style = "Quality"
    if (ShortName in " Stable Sales Gr " ): style = "Quality"
    if (ShortName in " Stable IBES 12M Gr " ): style = "Quality"
    if (ShortName in " Stable IBES FY1 R " ): style = "Quality"
    if (ShortName in " Stable Returns " ): style = "Quality"
    if (ShortName in " Earnings " ): style = "Growth"
    if (ShortName in " Employees " ): style = "Other"
    if (ShortName in " Mmntm 6M " ): style = "Momentum"
    if (ShortName in " IBES Sales Yld " ): style = "Value"
    if (ShortName in " IBES Div Yld " ): style = "Value"
    if (ShortName in " IBES EPS Yld " ): style = "Value"
    if (ShortName in " IBES EPS LTG " ): style = "Growth"
    if (ShortName in " IBES Sales LTG " ): style = "Growth"
    if (ShortName in " IBES Sales 12M Gr " ): style = "Growth"
    if (ShortName in " Sales to EV " ): style = "Value"
    if (ShortName in " EBITDA to EV " ): style = "Value"
    if (ShortName in " Sustainable GR " ): style = "Growth"
    if (ShortName in " Low Accruals " ): style = "Quality"
    if (ShortName in " FCf Yld " ): style = "Value"
    if (ShortName in " Carbon Footprint " ): style = "ESG"
    if (ShortName in " Impact Ratio " ): style = "ESG"
    if (ShortName in " Environment MSCI " ): style = "ESG"
    if (ShortName in " Climate Ch MSCI " ): style = "ESG"
    if (ShortName in " Nat Res Use MSCI " ): style = "ESG"
    if (ShortName in " Waste Mgmt MSCI " ): style = "ESG"
    if (ShortName in " Environ Opp's MSCI " ): style = "ESG"
    if (ShortName in " Social MSCI " ): style = "ESG"
    if (ShortName in " Human Cap MSCI " ): style = "ESG"
    if (ShortName in " Prod Safety MSCI " ): style = "ESG"
    if (ShortName in " Social Opp MSCI " ): style = "ESG"
    if (ShortName in " Governance MSCI " ): style = "ESG"
    if (ShortName in " Corp Gov MSCI " ): style = "ESG"
    if (ShortName in " Bus Ethics MSCI " ): style = "ESG"
    if (ShortName in " Gov & Pub Plcy MSCI " ): style = "ESG"
    if (ShortName in " SSI RavenPack " ): style = "Momentum"
    if (ShortName in " CVI RavenPack " ): style = "Momentum"
    if (ShortName in " ESG MSCI " ): style = "ESG"
    if (ShortName in " Accounting GMI " ): style = "ESG"
    if (ShortName in " AGR GMI " ): style = "ESG"
    if (ShortName in " FAM GMI " ): style = "ESG"
    if (ShortName in " ESG oekom " ): style = "ESG"
    if (ShortName in " Environment oekom " ): style = "ESG"
    if (ShortName in " Social & Gov oekom " ): style = "ESG"   
    if (ShortName in " Environ Opp_s MSCI " ): style = "ESG"
    if (ShortName in " ROIC " ): style = "Growth"
    if (ShortName in " ROA " ): style = "Growth"
    if (ShortName in " Asset Turnover " ): style = "Other"
    if (ShortName in " EBIT to EV " ): style = "Value"
    if (ShortName in " Div Pay Ratio " ): style = "Value"
    if (ShortName in " Gross Prof Marg " ): style = "Growth"
    if (ShortName in " Gross Prof to Assts " ): style = "Growth"
    if (ShortName in " Op Prof Margin " ): style = "Growth"
    if (ShortName in " Net Buyback Yld " ): style = "Value"
    if (ShortName in " Net Dbt Paydown Yld " ): style = "Value"
    if (ShortName in " Net Payout Yld " ): style = "Value"
    if (ShortName in " Tot Share Yld " ): style = "Value"
    if (ShortName in " 3Yr Vol " ): style = "Risk"
    if (ShortName in " Assets to Eq " ): style = "Other"
    if (ShortName in " Ergs Gr 5yr " ): style = "Growth"
    if (ShortName in " Sales Gr 5yr " ): style = "Growth"
    if (ShortName in " Sales Gr 3yr " ): style = "Growth"
    if (ShortName in " 5Yr Vol " ): style = "Risk"
    if (ShortName in " 1Yr Vol " ): style = "Risk"
    if (ShortName in " Exp to Sh Rate " ): style = "Economic"
    if (ShortName in " Exp to Infl " ): style = "Economic"
    if (ShortName in " Exp to Gold " ): style = "Economic"
    if (ShortName in " Exp to Oil " ): style = "Economic"
    if (ShortName in " Exp to Ccy " ): style = "Economic"
    if (ShortName in " Exp to GDP Surp " ): style = "Economic"
    if (ShortName in " ESG GMI " ): style = "ESG"
    if (ShortName in " Environment GMI " ): style = "ESG"
    if (ShortName in " Social GMI " ): style = "ESG"
    if (ShortName in " Governance GMI " ): style = "ESG"
    if (ShortName in " Board GMI " ): style = "ESG"
    if (ShortName in " Owner GMI " ): style = "ESG"
    if (ShortName in " Pay GMI " ): style = "ESG"
    if (ShortName in " Market Cap FF " ): style = "Risk"
    if (ShortName in " Stable EarningsGR Wght " ): style = "Quality"
    if (ShortName in " Stable SalesGR Wght " ): style = "Quality"
    if (ShortName in " Stable Fcast 12M EarnGR Wght " ): style = "Quality"
    if (ShortName in " Stable IBES FY1 RevFcast Wght " ): style = "Quality"
    if (ShortName in " Stable Returns Wght " ): style = "Quality"
    if (ShortName in " IBES ROE " ): style = "Growth"
    if (ShortName in " Capex Growth 3yr " ): style = "Growth"
    if (ShortName in " Capex Growth 5yr " ): style = "Growth"
    if (ShortName in " Asset Growth 3yr " ): style = "Growth"
    if (ShortName in " Asset Growth 5yr " ): style = "Growth"
    if (ShortName in " IBES 12M Fcast R 1M " ): style = "Growth"
    if (ShortName in " IBES 12M Fcast R 3M " ): style = "Growth"
    if (ShortName in " Historic Dividend Growth " ): style = "Growth"
    if (ShortName in " Daily 1yr Vol " ): style = "Risk"
    if (ShortName in " Mmntm 12-1 " ): style = "Momentum"
    if (ShortName in " CAEY " ): style = "Value"
    if (ShortName in " InversePEG " ): style = "Value"
    if (ShortName in " InversePEGY " ): style = "Value"
    if (ShortName in " Current Ratio (ex-Fin) " ): style = "Other"
    if (ShortName in " Interest Coverage Ratio (ex-Fin) " ): style = "Other"
    if (ShortName in " Quick Ratio (ex-Fin) " ): style = "Other"   
    if (ShortName in " Trading Turnover 3mth " ): style = "Other"  
    if (ShortName in " Stable Ergs Gr 5yr " ): style = "Quality"
    if (ShortName in " Stable Sales Gr 5yr " ): style = "Quality"  

    return style