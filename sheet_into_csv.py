import xlrd
import csv
xls = xlrd.open_workbook(r'ODOO_PRIX_VENTE_GROSSISTECENTRE&PROVINCE_DETAILLANTCENTRE&PROVINCE_080618.xlsx', on_demand=True)
sheets = xls.sheet_names()

for sheet_name in sheets:
    worksheet = xls.sheet_by_name(sheet_name)
    with open("csv/fixed_" + sheet_name + ".csv", 'wb') as csvfile:
        wr = csv.writer(csvfile, delimiter=';', quotechar='|', quoting=csv.QUOTE_MINIMAL)
        for rownum in xrange(worksheet.nrows):
            wr.writerow(
                list(x.encode('utf-8') if type(x) == type(u'') else x
                     for x in worksheet.row_values(rownum)))
