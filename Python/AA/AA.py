import win32com.client as excel
import requests
import logging
import os
import json

def main():
    SugarCRM = 'https://ruhrpumpen.sugarondemand.com'
    data = [('grant_type', 'password'), ('client_id', 'sugar'), ('client_secret', ''), ('username', 'rpServices'), ('password', 'Pass2018'), ('platform', 'mobile'),]
    response = requests.post(SugarCRM + '/rest/v11_20/oauth2/token/', data=data)
    access_token = response.json()['access_token']
    head = {'Authorization' : 'Bearer ' + access_token}

    Excel = excel.gencache.EnsureDispatch('Excel.Application')
    Document = 'Accounts reassignation from Perron'
    Workbook = Excel.Workbooks.Open(os.getcwd() +'\\'+ Document + '.xlsx')
    Excel.Visible = True
    WorkSheet = Workbook.Worksheets(1)

    logging.basicConfig(filename= Document +".log", level=logging.INFO,format='[ %(asctime)s ] %(levelname)s : %(message)s', datefmt='%m/%d/%Y %I:%M %p')
    logging.info('Connected')

    row = 3
    while (WorkSheet.Cells(row,1).Value is not None):
        GPS = str(int(WorkSheet.Cells(row,1).Value))

        Acc = requests.get(SugarCRM + '/rest/v11_21/Accounts/?filter[0][gps_id_c][$contains]=' + GPS, headers=head )
        Account_json = {}
        if len(Acc.json()['records']) > 0:
            if(Acc.json()['records'][0]['name'] == WorkSheet.Cells(row,2).Value) :
                if ( WorkSheet.Cells(row,5).Value is not None ):
                    user = list(WorkSheet.Cells(row,5).Value.split(" "))
                    Usr = requests.get(SugarCRM + '/rest/v11_21/Users/?filter[0][first_name][$contains]=' + user[0] + '&filter[0][last_name][$contains]=' + user[-1], headers=head )
                    Account_json["assigned_user_id"] = Usr.json()['records'][0]['id']
                    Account_json["assigned_user_name"] = Usr.json()['records'][0]['name']
                    logging.info('Assignation [ ' + Acc.json()['records'][0]['name'] + ' ] | Before : ' + Acc.json()['records'][0]['assigned_user_name'] + ' [' + Acc.json()['records'][0]['assigned_user_id'] + ']')
                    logging.info('Assignation [ ' + Acc.json()['records'][0]['name'] + ' ] | After : ' + Usr.json()['records'][0]['name'] + ' [' + Usr.json()['records'][0]['id'] + ']')

                if (WorkSheet.Cells(row,6).Value is not None):
                    account_class = WorkSheet.Cells(row,6).Value
                    Account_json["account_class_c"] = account_class.replace("Type ","")
                    logging.info('Account Class [ ' + Acc.json()['records'][0]['name'] + ' ] | Before : ' + Acc.json()['records'][0]['account_class_c'])
                    logging.info('Account Class [ ' + Acc.json()['records'][0]['name'] + ' ] | After : ' + account_class.replace("Type ",""))
        if ( len(Account_json) > 0 ):
            response = requests.put(SugarCRM + '/rest/v11_21/Accounts/'+ Acc.json()['records'][0]['id'] ,headers=head,data=json.dumps(Account_json))
            if response.status_code == 200 :
                WorkSheet.Cells(row,7).Value = "Updated"
        print(row)
        row += 1

    logging.info('Successful')
    Workbook.Save()
    Workbook.Close()
    Excel.Application.Quit()

main()