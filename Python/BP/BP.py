import requests
import json
import pyodbc
import logging

def main():
    #[requests] Stablishing conexion to Sugar (API)
    SugarCRM = 'https://ruhrpumpendev.sugarondemand.com'
    #data = [('grant_type', 'password'), ('client_id', 'sugar'), ('client_secret', ''), ('username', 'rodrigogr'), ('password', 'Ruhr2021'), ('platform', 'mobile'),]
    data = [('grant_type', 'password'), ('client_id', 'sugar'), ('client_secret', ''), ('username', 'rpServices'), ('password', 'Pass2018'), ('platform', 'mobile'),]
    response = requests.post(SugarCRM + '/rest/v11_4/oauth2/token/', data=data)
    access_token = response.json()['access_token']
    head = {'Authorization' : 'Bearer ' + access_token}
    #[pyodbc] Stablishing Conexion to Baan (SQL)
    query = pyodbc.connect('Driver={SQL Server};''Server=tulbaansql;''Database=baanlivedb;''UID=xlsmty;''PWD=read;').cursor()
    select = "SELECT TOP(10) * FROM BaaN_Business_Partners_List;"
    query.execute(select)
    #for records in range(len(response.json()['records'])):
    logging.basicConfig(filename="BP.log", level=logging.INFO,format='[ %(asctime)s ] %(levelname)s : %(message)s', datefmt='%m/%d/%Y %I:%M %p')
    logging.info('Connected')

    for i in query:
        #[json] Building JSON (Sugar ID VS SQL ID)
        BP = {}
        BP["name"] = str(i.BP_Code)
        BP["description"] = i.BP_Name
        BP["bp_code"] = i.BP_Code
        BP["status"] = i.Status
        BP["taxid"] = dif_zero(i.TaxID)
        BP["country"] = i.Country
        BP["state"] = i.State
        BP["creationdate"] = date_filter(i.CreationDate)
        BP["creationuser"] = i.CreationUser
        BP["lastmodificationdate"] = date_filter(i.LastModificationDate)
        BP["lastmodificationuser"] = i.LastModificationUser
        #[requests] Get the Sugar record to update
        response = requests.get(SugarCRM + '/rest/v11_4/bp_business_partners/filter?filter[0][name][$equals]=' + str(i.BP_Code),headers=head)
        if( len(response.json()['records']) > 0 ):
            SugarID = response.json()['records'][0]['id']
            response = requests.put(SugarCRM + '/rest/v11_4/bp_business_partners/'+ SugarID ,headers=head,data=json.dumps(BP))
            logging.info('PUT | ' + SugarID + ' | ' + str(i.BP_Code) + ' / ' + str(i.BP_Name))
        else:
            response = requests.post(SugarCRM + '/rest/v11_4/bp_business_partners',headers=head,data=json.dumps(BP))
            logging.info('POST | ' + str(i.BP_Code) + ' / ' + str(i.BP_Name))

def date_filter(datefilter):
    if (datefilter == "1970-01-01"):
        datefilter = None
    return datefilter

def dif_zero(difzero):
    if (difzero == '0'):
       difzero = None
    return difzero

main()