import xlwt
workbook = xlwt.Workbook(encoding='utf-8')

worksheet = workbook.add_sheet('1')

#worksheet.insert_bitmap('./Image/d8433551-55cf-455a-83ee-856012e1691f&0.bmp',1,1)
#worksheet.write(0,0,xlwt.Formula('HYPERLINK("%s";"Link")' % './Image/d8433551-55cf-455a-83ee-856012e1691f&0.bmp'))
worksheet.write(0, 0, xlwt.Formula('"test " & HYPERLINK("http://google.com")'))
workbook.save('simple.xls')
