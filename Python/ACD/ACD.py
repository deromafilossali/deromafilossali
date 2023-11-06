from threading import main_thread
import win32com.client as excel
#print(win32com.__gen_path__)
from datetime import timedelta
from datetime import datetime
import requests
import json

Excel = excel.gencache.EnsureDispatch('Excel.Application')

SugarCRM = 'https://ruhrpumpendev.sugarondemand.com'
data = [('grant_type', 'password'), ('client_id', 'sugar'), ('client_secret', ''), ('username', 'rodrigogr'), ('password', 'Ruhr2021'), ('platform', 'mobile'),]
response = requests.post(SugarCRM + '/rest/v11_4/oauth2/token/', data=data)
access_token = response.json()['access_token']
head = {'Authorization' : 'Bearer ' + access_token}

Location = 'C:\\Users\\rodrigogr\\Desktop\\ACD\\'
Document = 'ACD'
Workbook = Excel.Workbooks.Open(Location + Document + '.xlsx')
Excel.Visible = True
ACD_Sheet = Workbook.Worksheets("ACD")
Work_Sheet = Workbook.Worksheets("Work")
main_row = 2
next_row = main_row + 1
work_row = 2
date_placeholder = "2000-00-00"
def compare_parent():
    if ( Work_Sheet.Cells(work_row,2).Value is None ):
        Work_Sheet.Cells(work_row,2).Value = ACD_Sheet.Cells(main_row,6).Value #Write date in "Work" if empty
        Work_Sheet.Cells(work_row,3).Value = ACD_Sheet.Cells(main_row,6).Value
    if ((ACD_Sheet.Cells(main_row,6).Value is not None) and (ACD_Sheet.Cells(next_row,5).Value is not None) and (ACD_Sheet.Cells(next_row,6).Value is not None)): 
        after_main_row = ACD_Sheet.Cells(main_row,6).Value.date()
        after_next_row = ACD_Sheet.Cells(next_row,6).Value.date()
        if ( (after_main_row + timedelta(days=1)) == after_next_row ):
            ACD_Sheet.Cells(next_row,6).EntireRow.Interior.ColorIndex = 6
            ACD_Sheet.Cells(main_row,6).EntireRow.Interior.ColorIndex = 6         
            work_date = Work_Sheet.Cells(work_row,2).Value.date()
            work_date_2 = Work_Sheet.Cells(work_row,3).Value.date()
            if ( work_date > after_main_row ):
                Work_Sheet.Cells(work_row,2).Value = str(after_main_row)
            if ( work_date_2 < after_next_row ):
                Work_Sheet.Cells(work_row,3).Value = str(after_next_row)
    return Work_Sheet.Cells(work_row,2).Value.date()

def update_sugar_record(UpdateSugarRecord_date,UpdateSugarRecord_id):
    Update_Record = {}
    print(UpdateSugarRecord_date)
    Update_Record["actual_close_date_c"] = str(UpdateSugarRecord_date)
    requests.put(SugarCRM + '/rest/v11_4/Quotes/'+ UpdateSugarRecord_id ,headers=head,data=json.dumps(Update_Record))


while ACD_Sheet.Cells(next_row,1).Value is not None:
    parent_main_row = ACD_Sheet.Cells(main_row,1).Value
    parent_next_row = ACD_Sheet.Cells(next_row,1).Value
    Work_Sheet.Cells(work_row,1).Value = parent_next_row
    if (parent_main_row == parent_next_row) :
        CompareParent = compare_parent()
        if ( date_placeholder != CompareParent):
            update_sugar_record(CompareParent,parent_main_row)
        date_placeholder = CompareParent
    else:
        work_row += 1
        date_placeholder = "2000-00-00"
    main_row += 1
    next_row += 1
    print('row '+str(main_row))

Workbook.Save()
Workbook.Close()
Excel.Application.Quit()