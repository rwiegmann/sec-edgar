import feedparser
import os.path
import sys, getopt
import time
import socket
from urllib.request import urlopen
from urllib.error import URLError, HTTPError
import xml.etree.ElementTree as ET
import zipfile
import zlib
import xlsxwriter

##############################################

#    SET DATE HERE

year = 2016
month = 12


##############################################



# edgarFilingsFeed = 'https://www.sec.gov/Archives/edgar/monthly/xbrlrss-2017-11.xml'
edgarFilingsFeed = 'http://www.sec.gov/Archives/edgar/monthly/xbrlrss-' + str(year) + '-' + str(month).zfill(2) + '.xml'

feedFile = urlopen(edgarFilingsFeed)
feedData = feedFile.read()
feedFile.close()

feed = feedparser.parse(feedData)



# Create a workbook and add a worksheet.
#workbook = xlsxwriter.Workbook('filings.xlsx')

workbook = xlsxwriter.Workbook('filings' + str(year) + '-' + str(month).zfill(2) + '.xlsx')
worksheet = workbook.add_worksheet()


# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

# Iterate over the data and write it out row by row.
for item in (feed.entries):
    worksheet.write(row, col,     item['summary'])
    worksheet.write(row, col + 1, item['title'])
    worksheet.write(row, col + 2, item['edgar_ciknumber'])
    worksheet.write(row, col + 3, item['published'])
    row += 1

workbook.close()









