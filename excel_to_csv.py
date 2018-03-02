import xlrd
import csv

def excel_to_csv(ExcelFile, SheetName, CSVFile):
     workbook = xlrd.open_workbook(ExcelFile)
     worksheet = workbook.sheet_by_name(SheetName)
     csvfile = open(CSVFile, 'wb')
     #wr = csv.writer(csvfile, delimiter=';', quoting=csv.QUOTE_NONE)
     wr = csv.writer(csvfile, delimiter=';',quotechar='|', quoting=csv.QUOTE_MINIMAL)
     for rownum in xrange(worksheet.nrows):
         wr.writerow(
             list(x.encode('utf-8') if type(x) == type(u'') else x
                  for x in worksheet.row_values(rownum)))
     csvfile.close()


ExcelFile = "GL2017-5_analyse_correctifs_etech_010318_04.xlsx"
CSVFile = "GL2017-5_0417.csv"
SheetName = "Feuil1"

excel_to_csv(ExcelFile, SheetName, CSVFile)



