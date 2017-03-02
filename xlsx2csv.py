#coding: utf-8
import csv
import os
import sys
import xlrd

def xlsx2csv(path):
  for subdir, dirs, files in os.walk(path):
      for file in files:
          filepath = subdir + os.sep + file
          if filepath.endswith(".xls") or filepath.endswith(".xlsx"): 
              print(filepath)
              Excel2CSV(filepath)



def Excel2CSV(ExcelFile):
     workbook = xlrd.open_workbook(ExcelFile)
     worksheet = workbook.sheet_by_index(0)
     csvfile = open(ExcelFile + '.csv', 'wb')
     wr = csv.writer(csvfile, quoting=csv.QUOTE_ALL)

     for rownum in xrange(worksheet.nrows):
         wr.writerow(
             list(x.encode('utf-8') if type(x) == type(u'') else x
                  for x in worksheet.row_values(rownum)))

     csvfile.close()

if __name__ == '__main__':
    if len(sys.argv) != 2:
        print('usage: python xlsx2csv.py <path>')
    else:
        xlsx2csv(sys.argv[1])
    sys.exit(0)
