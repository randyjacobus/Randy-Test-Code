# -*- coding: utf-8 -*-
"""
Created on Sun Jun 14 18:46:07 2015

This code downloads the municipla closed end fund data for hte Nuveen funds.

@authors: Randy, Jenica, Ricky and Nikkita
"""

# Import Modules
import xlrd
import pandas as pd
import urllib
import requests
from bs4 import BeautifulSoup
import re
from collections import defaultdict
import os

#*****************************************************************************
# Create a nuveenunidata.xls file for the uni data
# Create a nuveenhtmldata.xls for the html data
# Create a nuveenportfolio.xls file for the portfolio data
# This code will determine the working directory and overwrite 
# any existing files
#*****************************************************************************
workingdirectory=os.getcwd()
nuveenunifile ='nuveenunidata.xls'
nuveenhtmlfile='nuveenhtmldata.xls'
nuveenportfile='nuveenportfolio.xls'
nuveenunidata = os.path.join(workingdirectory,nuveenunifile)
nuveenhtmldata = os.path.join(workingdirectory,nuveenhtmlfile)
nuveenportdata = os.path.join(workingdirectory,nuveenportfile)
open(nuveenunifile,'w')
open(nuveenhtmlfile,'w')
open(nuveenportfile,'w')
nuveenuniurl = 'http://www.nuveen.com/Home/Documents/Viewer.aspx?fileId=65923'
nuveenhtmlurl= 'http://www.nuveen.com/CEF/DailyPricingTaxExempt.aspx'
#******************************************************************************
# Excel Parser - Find uni data
#******************************************************************************
# Import Excel File from Web
file = requests.get(nuveenuniurl)
output = open(nuveenunidata, "wb")
output.write(file.content)
output.close()

# Load File
wb = xlrd.open_workbook(filename)

# Identify the sheet the data is in in
sheet_name = 'Municipal Funds'
sheet = wb.sheet_by_name(sheet_name)

# Use list comprhension to pull the data fields
starting_row = 0
uni=[[sheet.cell_value(i,0), sheet.cell_value(i,1), sheet.cell_value(i,3), 
      sheet.cell_value(i,5), sheet.cell_value(i,6)] 
      for i in range(starting_row, sheet.nrows) 
      if len(sheet.cell(i,0).value)==3]

# Create a list of tickers for later use
tickers = [uni[i][0] for i in range (0,len(uni))]
#*******************************************************************************
# HTML Parser - Find basic data
#*******************************************************************************
#little function that cleans data
def cleanText(value):
    return value.replace("$", "").replace("%", "").replace(",", "");

# beautifuksoup looking for tr and td tags
soup = BeautifulSoup(urllib.request.urlopen(nuveenhtmlurl).read())
soup('table')[0].prettify()
html = []
for tr in soup('table')[0].findAll('tr'):
    tds = tr.findAll('td')
    if len(tds) > 1:
        html.append([tds[0].text, cleanText(tds[3].text), cleanText(tds[5].text), 
                     cleanText(tds[7].text), cleanText(tds[11].text), 
                     cleanText(tds[12].text)])
#print(html[0])
#******************************************************************************
# Analyze portfolio data
#******************************************************************************
# Identify the file and the sheet the data is in in, including the starting row.
print(nuveenportdata)
wb = xlrd.open_workbook(nuveenportdata)
sheet_name = 'Tax-Exempt Municipal Debt'
sheet = wb.sheet_by_name(sheet_name)
starting_row = 5 
ending_row = sheet.nrows

data = defaultdict(list)

stateCol = 3
portfolioCol = 6
marketvalueCol = 9
sectionCol = 16
tobpositionCol = 1

# load all rows except FRS bonds
for i in range(starting_row, ending_row):
    if (sheet.cell_value(i, tobpositionCol) != 'FRS'):
        data[sheet.cell_value(i,0)].append([sheet.cell_value(i, stateCol), 
        sheet.cell_value(i, portfolioCol), sheet.cell_value(i, marketvalueCol), 
        sheet.cell_value(i, sectionCol)])

print(data)

portfoliodata =[[sheet.cell_value(i,0)].append([sheet.cell_value(i, stateCol), 
        sheet.cell_value(i, portfolioCol), sheet.cell_value(i, marketvalueCol), 
        sheet.cell_value(i, sectionCol) for i in range (starting_row,ending_row)
        if (sheet.cell_value(i, tobpositionCol) != 'FRS')]

print(portfoliodata)

marketvalueListPos = 2
portfolioListPos = 1
stateListPos = 0
sectionListPos = 3

Nuveen_Output_ALL = []
for fund in Nuveen_Output:
    
    ticker = fund[0]

    # market value
    market_value_fund = [data[ticker][i][marketvalueListPos] for i in range(0, len(data[ticker]))]
    market_value_fund_sum = sum(market_value_fund)

    # in PR
    portfolio_fund_in_PR = [data[ticker][i][portfolioListPos]/100 for i in range(0, len(data[ticker])) if data[ticker][i][stateListPos] == 'PR']
    portfolio_fund_in_PR_sum = sum(portfolio_fund_in_PR)

    # In IL
    portfolio_fund_in_IL = [data[ticker][i][portfolioListPos]/100 for i in range(0, len(data[ticker])) if data[ticker][i][stateListPos] == 'IL']
    portfolio_fund_in_IL_sum = sum(portfolio_fund_in_IL)

    # In Consumer Staples
    portfolio_fund_in_TOB = [data[ticker][i][portfolioListPos]/100 for i in range(0, len(data[ticker])) if data[ticker][i][sectionListPos] == 'Consumer Staples']
    portfolio_fund_in_TOB_sum = sum(portfolio_fund_in_TOB)

    # In Health Care
    portfolio_fund_in_Smoke = [data[ticker][i][portfolioListPos]/100 for i in range(0, len(data[ticker])) if data[ticker][i][sectionListPos] == 'Health Care']
    portfolio_fund_in_Smoke_sum = sum(portfolio_fund_in_Smoke)       

    
    Nuveen_Output_ALL.append(fund + [market_value_fund_sum, portfolio_fund_in_PR_sum, portfolio_fund_in_IL_sum, portfolio_fund_in_TOB_sum, portfolio_fund_in_Smoke_sum])

filenameAnalysis ='C:\\Users\\Jenica\\Desktop\\Projects\\Week2\\Week2Code\\Data\\NuveenAnalysis.xlsx'
# need code to save Nuveen_Output_ALL to filenameAnalysis

#*******************************************************************************
# Output the data to the excel file using data nitro.
#*******************************************************************************

