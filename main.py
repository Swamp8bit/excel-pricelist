
import xlrd
import xlwt
#open workbook
xl_workbook = xlrd.open_workbook('RoughtPriceData.xls',formatting_info=True, encoding_override="utf-8")
#get sheet_names
sheet_names = xl_workbook.sheet_names() #get sheet_names in workbook
product={} #create dictionary for product with fields article, product name, volume and price
product_catalogue=[] # create product catalogue for products
if "__main__"==__name__:


    # for i in sheet_names:
    #     print i
    for sheet_name in sheet_names:

        xl_sheet = xl_workbook.sheet_by_name(sheet_name)
        print ('Sheet name: %s' % xl_sheet.name)


        #first_row=xl_sheet.row(9)

        try:
            rows = range(10,63,1)
            for i in rows:
                row = xl_sheet.row(i)
                for j in row:
                    print j.value
        except IndexError:
            continue



