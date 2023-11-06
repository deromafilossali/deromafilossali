import win32com.client as excel
Excel = excel.gencache.EnsureDispatch('Excel.Application')

Location = 'C:\\Users\\rodrigogr\\Desktop\\ACD\\'
Document = 'ACD'
Workbook = Excel.Workbooks.Open(Location + Document + '.xlsx')
#Excel.Visible = False
Sheet = Workbook.Worksheets(1)
first_row = 2
second_row = 3

a = Sheet.Cells(first_row,6).Value.date()
b = Sheet.Cells(first_row,3).Value.date()

print(a)
print(b)

Workbook.Save()
Workbook.Close()
Excel.Application.Quit()