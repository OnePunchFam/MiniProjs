import urllib
import os, sys
import urllib2
import webbrowser
import csv
import xlrd
from xlrd import open_workbook,XL_CELL_TEXT
import xlwt
import sys
import time
from xlutils.copy import copy
from urllib import urlencode, unquote
from urlparse import urlparse, parse_qsl, ParseResult
from json import dumps
from urllib2 import urlopen
from urlparse import urlparse
from urlparse import urlsplit


from urllib import urlencode
from urlparse import parse_qs, urlsplit, urlunsplit




period = raw_input('What frequency? 12 or 3?: ')
savedfile = raw_input('What file is you doing?: ')

mainFile = 'holla bruh.csv'
other_file = 'CSLT Income Statement.csv'

Metrics = ['Revenue' , 'Net income' , 'Earnings per share' , 'Operating expenses' , 'Sales, General and administrative' , 'Diluted' , 'EBITDA' , 'Net income available to common shareholders']
#------------------------------------------------------------------------------------------------------

#stock_url = 'http://financials.morningstar.com/ajax/ReportProcess4CSV.html?t=' + ticker + '&reportType=' + reportType + '&period=12&dataType=A&order=asc&columnYear=5&number=3'

def set_query_parameter(url, param_name, param_value):
    """Given a URL, set or replace a query parameter and return the
    modified URL.

    >>> set_query_parameter('http://financials.morningstar.com/ajax/ReportProcess4CSV.html?t=aapl&reportType=is&period=12&dataType=A&order=asc&columnYear=5&number=3', 't', 'td')
    'http://financials.morningstar.com/ajax/ReportProcess4CSV.html?t=td&reportType=is&period=12&dataType=A&order=asc&columnYear=5&number=3'

    """
    scheme, netloc, path, query_string, fragment = urlsplit(url)
    query_params = parse_qs(query_string)

    query_params[param_name] = [param_value]
    new_query_string = urlencode(query_params, doseq=True)

    return urlunsplit((scheme, netloc, path, new_query_string, fragment))

def download_file(tick):
    webbrowser.open('http://financials.morningstar.com/ajax/ReportProcess4CSV.html?t='+tick+'&reportType=is&period='+period+'&dataType=A&order=asc&columnYear=5&number=3')
	

#------------------------------------------------------------------------------------------------------

def loop_through_tickers():
    for i in range(0, 2):
        ticker = read_cell(1, i, mainFile)
        currentMetric = Metrics[i]
        read_entire_row(find_metric_row(currentMetric))


#------------------------------------------------------------------------------------------------------

def read_cell(x, y, csvFile):
    try:
        #open the file and read the cells
        with open(csvFile, 'r') as f: 
        #should csvfile be replaced with other file?
            reader = csv.reader(f)
            y_count = 0
            for n in reader:
                if y_count == y:
                    cell = n[x]
                    print(cell)
                    return cell
                y_count += 1
        #return "no value found"
    except IOError:
        return ''

#------------------------------------------------------------------------------------------------------

def read_entire_row(row):
  #read rows for symbol (0, 1 = stock symbol)
  #(0 , 2 = most recent year) (0, 3 = 2nd year etc)  
    for i in range(0, 2):
        print(read_cell(i,row, other_file))
    print "\n"

#------------------------------------------------------------------------------------------------------


def read_stockfile_cells(value='Revenue'):
    # Gives the row/column of cell with given string
    # value = 'Total Liabilities'
    book = open_workbook(savedfile)
    sheet = book.sheet_by_index(0)
    cell = sheet.cell(0,0)
    for row in range(sheet.nrows):
        for column in range(sheet.ncols):
            if sheet.cell(row,column).value == value:
                print row, column
                return column


#------------------------------------------------------------------------------------------------------

def write_into_file(row, column, value):
    rb = xlrd.open_workbook(savedfile)
    wb = copy(rb)
    #need to make a writable copy
    w_sheet = wb.get_sheet(0)
    #primary worksheet
    w_sheet.write(row, column, str(value))
    # (first) = row number
    # (second) = column number
    # value of metric
    
    wb.save(savedfile)

#write_into_file(2, 58, 'nigga')

#------------------------------------------------------------------------------------------------------

def download_stock_files():
 #take symobls from mainfile and return the list   
    year = input('What year you want?: ')
    dict = {'2016':5 , '2015':4 , '2014':3 , '2013':2 , '2012' :1}
    for i in range(1, 2):
        # (1,1) = downloads nothing
        # (1,2) = first stock, 
        # (1,3) = second stock, etc
        ticker = read_cell(1, i, mainFile)
        print read_cell(1, i, mainFile)
        download_file(ticker)
        time.sleep(2)
        other_file = ticker + " Income Statement.csv"
        print(other_file)
        for met in Metrics:
            
            row = find_metric_row(met)
            print(row , 'metric row 2016') #this is the test: print row --- saves metric row as row
            
            main_col = read_stockfile_cells(met)
            print(main_col , 'mainfile column')
            # this is the test: print main_col ---- saves column for mainfile with same metric name
            
            #col 5 is for 2016
            try:
                data_val = read_cell(dict[str(year)], row, other_file)
            except IndexError:
                dat_val = 'no val'
            print(data_val , 'data value')# this is the test: print data_val ---- saves data value for column (2016)
            
            write_into_file(i, main_col, str(data_val))
        os.remove(other_file)






#------------------------------------------------------------------------------------------------------


def find_metric_row(data_Name):
  #find the row of the metric in the file  
    row_metric = None
    for i in range(0, 60):
        category = read_cell(0, i, other_file)
        if str(category) == data_Name:
            row_metric = i
            break
        
    return row_metric  


'''
example:
print find_metric_row("Total liabilities")
print find_metric_row("Total assets")
print find_metric_row("Liabilities")
print find_metric_row("Total stockholders' equity")
print find_metric_row("Total current liabilities")
print find_metric_row("Total current assets")
print find_metric_row("Total current liabilities")
print find_metric_row("Total cash")
print find_metric_row("Retained earnings")
'''


#-------------------------------------------------------------------------------------------------------


def loop_through_metrics():
    #loop through the different metrics i want for each csv file, ex. TA, TL, TE... etc.
    for i in range(0 , len(Metrics)):
        currentMetric = Metrics[i]
        read_entire_row(find_metric_row(currentMetric))

#loop_through_metrics()


'''
example:
read_entire_row(find_metric_row('Total current assets'))
read_entire_row(find_metric_row('Total current liabilities'))
read_entire_row(find_metric_row('Total cash'))
read_entire_row(find_metric_row('Total assets'))
read_entire_row(find_metric_row('Retained earnings'))
read_entire_row(find_metric_row('Total cash'))
read_entire_row(find_metric_row("Total stockholders' equity"))
read_entire_row(find_metric_row("Total liabilities"))
'''

#------------------------------------------------------------------------------------------------------------

def read_stock_file_rows():
 #read full row of stock spread sheet xlsx file   
    from xlrd import open_workbook,XL_CELL_TEXT
    book = open_workbook(savedfile)
    sheet = book.sheet_by_index(0)
    cell = sheet.cell(0,0)
    for i in range(sheet.ncols):
        print sheet.cell_type(1,i),sheet.cell_value(1,i)

#read_stock_file_rows()

#------------------------------------------------------------------------------------------------------

#read_stockfile_cells()

#------------------------------------------------------------------------------------------------------
download_stock_files()



