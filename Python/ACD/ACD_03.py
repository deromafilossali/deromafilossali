from threading import main_thread
import win32com.client as excel
#print(win32com.__gen_path__)
from datetime import timedelta
Excel = excel.gencache.EnsureDispatch('Excel.Application')

Location = 'C:\\Users\\rodrigogr\\Desktop\\ACD\\'
Document = 'ACD'
Workbook = Excel.Workbooks.Open(Location + Document + '.xlsx')
Excel.Visible = False
ACD_Sheet = Workbook.Worksheets("ACD")
Work_Sheet = Workbook.Worksheets("Work")
main_row = 2
next_row = main_row + 1
write_work = False
work_row = 2
print(ACD_Sheet.Cells(main_row,1))

def compare_parent():
    global work_row
    parent_main_row = ACD_Sheet.Cells(main_row,1).Value
    parent_next_row = ACD_Sheet.Cells(next_row,1).Value    
    if (parent_main_row == parent_next_row) :
        write_work = True
        if ((ACD_Sheet.Cells(main_row,6).Value is not None) and (ACD_Sheet.Cells(next_row,5).Value is not None) and (ACD_Sheet.Cells(next_row,6).Value is not None)): 
            after_main_row = ACD_Sheet.Cells(main_row,6).Value.date()
            after_next_row = ACD_Sheet.Cells(next_row,6).Value.date()
            if ( after_main_row == (after_next_row + timedelta(days=1))):
                ACD_Sheet.Cells(next_row,6).EntireRow.Interior.ColorIndex = 6
                ACD_Sheet.Cells(main_row,6).EntireRow.Interior.ColorIndex = 6
                a = Work_Sheet.Cells(work_row,2).Value
                print(a)
                if (Work_Sheet.Cells(work_row,2).Value < after_main_row or Work_Sheet.Cells(work_row,2).Value is None):
                    Work_Sheet.Cells(work_row,1).Value = parent_next_row
                    Work_Sheet.Cells(work_row,2).Value = after_main_row
                else:
                    work_row += 1

# def compare_rows_rev01():
#     parent_main_row = ACD_Sheet.Cells(main_row,1).Value
#     parent_next_row = ACD_Sheet.Cells(next_row,1).Value
#     if (parent_main_row == parent_next_row) :
#         #print("parent " + str(parent_main_row) + " / " + str(parent_next_row) )
#         createdby_main_row = ACD_Sheet.Cells(main_row,2).Value
#         createdby_next_row = ACD_Sheet.Cells(next_row,2).Value
#         datecreated_fist_row = ACD_Sheet.Cells(main_row,3).Value.date()
#         datecreated_next_row = ACD_Sheet.Cells(next_row,3).Value.date()
#         if (createdby_main_row == createdby_next_row and datecreated_fist_row == datecreated_next_row ):
#             #print("createdby " + str(createdby_main_row) + " / " + str(createdby_next_row) )
#             #print("datecreated " + str(datecreated_fist_row) + " / " + str(datecreated_next_row) )
#             after_main_row = ACD_Sheet.Cells(main_row,6).Value.date()
#             before_next_row = ACD_Sheet.Cells(next_row,5).Value.date()
#             if (after_main_row == before_next_row ) :
#                 #print("after1 and before2 " + str(after_main_row) + " / " + str(before_next_row) )
#                 ACD_Sheet.Cells(next_row,6).Interior.ColorIndex = 6
#                 ACD_Sheet.Cells(main_row,5).Interior.ColorIndex = 6
#                 #print('Colored')
                
while ACD_Sheet.Cells(next_row,1).Value is not None:
    #print("rows " + str(main_row) + " / " + str(next_row) )
    compare_parent()
    main_row += 1
    next_row += 1
    print('row '+str(main_row))

Workbook.Save()
Workbook.Close()
Excel.Application.Quit()
