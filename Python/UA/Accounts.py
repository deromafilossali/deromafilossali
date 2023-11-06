import win32com.client as excel
#print(win32com.__gen_path__)
import requests
import json

def main():
    Excel = excel.gencache.EnsureDispatch('Excel.Application')

    SugarCRM = 'https://ruhrpumpen.sugarondemand.com'
    data = [('grant_type', 'password'), ('client_id', 'sugar'), ('client_secret', ''), ('username', 'rodrigogr'), ('password', 'Ruhr2021'), ('platform', 'mobile'),]
    response = requests.post(SugarCRM + '/rest/v11_4/oauth2/token/', data=data)
    access_token = response.json()['access_token']
    head = {'Authorization' : 'Bearer ' + access_token}

    Location = 'C:\\Users\\rodrigogr\\OneDrive - ruhrpumpen.com\\Backup\\Python\\Accounts\\'
    Document = 'Accounts'
    Workbook = Excel.Workbooks.Open(Location + Document + '.xlsx')
    Excel.Visible = False
    Sheet = Workbook.Worksheets("Accounts")
    main_row = 2

    while Sheet.Cells(main_row,1).Value is not None:
        Account_id = Sheet.Cells(main_row,2).Value
        Update_Record = {}
        Update_Record["agent_distributor_c"] = ""
        Update_Record["account_id1_c"] = ""
        if (Sheet.Cells(main_row,11).Value is None ):
            print(Sheet.Cells(main_row,1).Value)
            requests.put(SugarCRM + '/rest/v11_4/Accounts/'+ Account_id ,headers=head,data=json.dumps(Update_Record))
            Sheet.Cells(main_row,11).Value = "Distributor Removed"
        main_row += 1

    Workbook.Save()
    Workbook.Close()
    Excel.Application.Quit()

main()