from tkinter import *
from tkinter import ttk
import tkinter.messagebox
import win32com.client as excel
import requests
import os
from datetime import date, datetime,timedelta



root = Tk()
root.title("Approval Report Generator")
root.wm_attributes('-toolwindow', 'True')
root.call('wm', 'attributes', '.', '-topmost', '1')

frame = ttk.Frame(root, padding=10)
frame.grid()

def generate_report(GPS,Byt):
    SugarCRM = 'https://ruhrpumpen.sugarondemand.com'
    data = [('grant_type', 'password'), ('client_id', 'sugar'), ('client_secret', ''), ('username', 'rpServices'), ('password', 'Pass2018'), ('platform', 'mobile'),]
    response = requests.post(SugarCRM + '/rest/v11_20/oauth2/token/', data=data)
    access_token = response.json()['access_token']
    head = {'Authorization' : 'Bearer ' + access_token}

    GPS = GPS.strip()

    Byts = requests.get(SugarCRM + '/rest/v11_20/nth_Buyouts/?filter[0][gps_number_c][$contains]=' + GPS + '&filter[1][buyout_type_c][$equals]=' + Byt, headers=head )
    if len(GPS) > 5 :
        if len(Byts.json()['records']) > 0 :
            Excel = excel.gencache.EnsureDispatch('Excel.Application')
            Workbook = Excel.Workbooks.Add()
            Workbook.Worksheets(1).Name = GPS
            Sheet = Workbook.Worksheets(GPS)
            Excel.Visible = True
            Sheet.Cells.Interior.Color = rgbToInt((255,255,255))

            MainRow = 3

            Sheet.Range("E2").Value = Byts.json()['records'][0]['gps_number_c']
            Sheet.Range("E2").Style = 'Normal'
            Sheet.Range("E2").Font.Color = rgbToInt((6,121,200))
            Sheet.Range("E2").Font.Bold = True
            Sheet.Range("G2").Value = 'GPS COST'
            Sheet.Range("G2:J2").Interior.Color = rgbToInt((6,121,200))
            Sheet.Range("G2:J2").Font.Color = rgbToInt((255,255,255))
            Sheet.Range("I2").Value = 'SELL PRICE'
            Sheet.Range("G2:H2").MergeCells = True
            Sheet.Range("I2:J2").MergeCells = True

            Sheet.Cells(MainRow,2).Value = 'Lead Time'
            Sheet.Cells(MainRow,3).Value = 'TAG'
            Sheet.Cells(MainRow,4).Value = 'Pump Type'
            Sheet.Cells(MainRow,5).Value = 'Description'
            Sheet.Cells(MainRow,6).Value = 'Qty.'
            Sheet.Cells(MainRow,7).Value = 'Each'
            Sheet.Cells(MainRow,8).Value = 'Total'
            Sheet.Cells(MainRow,9).Value = 'Each'
            Sheet.Cells(MainRow,10).Value = 'Total'
            Sheet.Range("B3:J3").Interior.Color = rgbToInt((6,121,200))
            Sheet.Range("B3:J3").Font.Color = rgbToInt((255,255,255))

            for byt in Byts.json()['records']:      
                BS_Count = requests.get(SugarCRM + '/rest/v11_20/BS_Buyouts_Summary/count&filter[0][bs_buyouts_summary_nth_buyoutsnth_buyouts_ida]='+ byt['id'] , headers=head )

                if ( BS_Count.json()['record_count'] > 0 ) and MainRow == 3:
                    BSs = requests.get(SugarCRM + '/rest/v11_20/BS_Buyouts_Summary/&order_by=bs_buyouts_summary_accounts_name&filter[0][bs_buyouts_summary_nth_buyoutsnth_buyouts_ida]='+ byt['id'] , headers=head )
                    MainColumn = 12
                    for bs in BSs.json()['records']:
                        Sheet.Range( index(MainColumn,True) + "2").Value = bs['bs_buyouts_summary_accounts_name']
                        Sheet.Range( index(MainColumn,True) + "2:" + index(MainColumn,False) + "2" ).MergeCells = True
                        Sheet.Range( index(MainColumn,True) + "2:" + index(MainColumn,False) + "2" ).Interior.Color = rgbToInt((6,121,200))
                        Sheet.Range( index(MainColumn,True) + "2:" + index(MainColumn,False) + "2" ).Font.Color = rgbToInt((255,255,255))

                        Sheet.Cells(MainRow,MainColumn).Value = 'Each'
                        Sheet.Cells(MainRow,MainColumn + 1).Value = 'Total'
                        Sheet.Cells(MainRow,MainColumn + 2).Value = 'ETA'
                        Sheet.Cells(MainRow,MainColumn + 3).Value = 'IncoTerms'
                        Sheet.Cells(MainRow,MainColumn + 4).Value = 'Location'
                        Sheet.Cells(MainRow,MainColumn + 5).Value = 'Payment'
                        Sheet.Cells(MainRow,MainColumn + 6).Value = 'Warranty'
                        Sheet.Cells(MainRow,MainColumn + 7).Value = 'Cost'
                        Sheet.Cells(MainRow,MainColumn + 8).Value = 'Sales'

                        Sheet.Range( index(MainColumn,True) + "3:" + index(MainColumn,False) + "3" ).Interior.Color = rgbToInt((6,121,200))
                        Sheet.Range( index(MainColumn,True) + "3:" + index(MainColumn,False) + "3" ).Font.Color = rgbToInt((255,255,255))

                        MainColumn += 10

                MainRow += 1
                Product = requests.get(SugarCRM + '/rest/v11_20/Products/&filter[0][quote_id]=' + byt['quotes_nth_buyouts_1quotes_ida'] + '&filter[1][gps_item_number_c]=' + byt['tag_c'] , headers=head )
                try :
                    size = Product.json()['records'][0]['size_qli_c']
                except :
                    size = "ERROR"
                Quote = requests.get(SugarCRM + '/rest/v11_20/Quotes/' + byt['quotes_nth_buyouts_1quotes_ida'] , headers=head )
                try:
                    dates = datetime.strptime(Quote.json()['oe_order_date_c'], '%Y-%m-%d') + timedelta(days=(int(byt['cust_delivery_time_c'])*7)) - datetime.now()
                    dates = dates.days/7
                except :
                    dates = "ERROR"
                Sheet.Cells(MainRow,2).Value = dates
                #Sheet.Cells(MainRow,2).NumberFormat = "mm/dd/yyyy"
                Sheet.Cells(MainRow,3).Value = byt['tag_c']
                Sheet.Cells(MainRow,4).Value = size
                Sheet.Cells(MainRow,5).Value = description(byt)
                Sheet.Cells(MainRow,6).Value = byt['quantity_c']
                Sheet.Cells(MainRow,7).Value = byt['cost_usd_c']
                Sheet.Cells(MainRow,7).Style = 'Currency'
                Sheet.Cells(MainRow,8).Value = int(byt['cost_usd_c']) * int(byt['quantity_c'])
                Sheet.Cells(MainRow,8).Style = 'Currency'
                Sheet.Cells(MainRow,9).Value = byt['price_usd_c']
                Sheet.Cells(MainRow,9).Style = 'Currency'
                Sheet.Cells(MainRow,10).Value = byt['extended_price_c']
                Sheet.Cells(MainRow,10).Style = 'Currency'

                if ( BS_Count.json()['record_count'] > 0 ):
                    BSs = requests.get(SugarCRM + '/rest/v11_20/BS_Buyouts_Summary/&order_by=bs_buyouts_summary_accounts_name&filter[0][bs_buyouts_summary_nth_buyoutsnth_buyouts_ida]='+ byt['id'] , headers=head )
                    MainColumn = 12
                    for bs in BSs.json()['records']:
                        Sheet.Cells(MainRow,MainColumn).Value = bs['price_usd']
                        Sheet.Cells(MainRow,MainColumn).Style = 'Currency'
                        MainColumn += 1
                        Sheet.Cells(MainRow,MainColumn).Value = bs['total_usd']
                        Sheet.Cells(MainRow,MainColumn).Style = 'Currency'
                        MainColumn += 1
                        Sheet.Cells(MainRow,MainColumn).Value = bs['delivery_time']
                        MainColumn += 1
                        Sheet.Cells(MainRow,MainColumn).Value = incoterms(bs['inco_terms'])
                        MainColumn += 1
                        Sheet.Cells(MainRow,MainColumn).Value = bs['location_incoterms']
                        MainColumn += 1
                        Sheet.Cells(MainRow,MainColumn).Value = bs['payment_method']
                        MainColumn += 1
                        Sheet.Cells(MainRow,MainColumn).Value = warranties(bs['warranty'])
                        MainColumn += 1
                        Sheet.Cells(MainRow,MainColumn).Value = bs['margin_cost']/100
                        Sheet.Cells(MainRow,MainColumn).Style = 'Percent'
                        MainColumn += 1
                        Sheet.Cells(MainRow,MainColumn).Value = bs['margin_sales']/100
                        Sheet.Cells(MainRow,MainColumn).Style = 'Percent'
                        MainColumn += 2

            MainRow += 1
            Sheet.Cells(MainRow,6).Formula = '=SUM(F4:F' + str(MainRow - 1) + ')' #Qty
            #Sheet.Cells(MainRow,7).Formula = '=SUM(G4:G' + str(MainRow - 1) + ')' #Each GPS Cost
            Sheet.Cells(MainRow,8).Formula = '=SUM(H4:H' + str(MainRow - 1) + ')' #Total GPS Cost
            #Sheet.Cells(MainRow,9).Formula = '=SUM(I4:I' + str(MainRow - 1) + ')' #Each Sell Price
            Sheet.Cells(MainRow,10).Formula = '=SUM(J4:J' + str(MainRow - 1) + ')' #Total Sell Price
            Sheet.Range("F" + str(MainRow) + ":J" + str(MainRow)).Interior.Color = rgbToInt((6,121,200))
            Sheet.Range("F" + str(MainRow) + ":J" + str(MainRow)).Font.Color = rgbToInt((255,255,255))

            if ( BS_Count.json()['record_count'] > 0 ):
                BSs = requests.get(SugarCRM + '/rest/v11_20/BS_Buyouts_Summary/&order_by=bs_buyouts_summary_accounts_name&filter[0][bs_buyouts_summary_nth_buyoutsnth_buyouts_ida]='+ byt['id'] , headers=head )
                MainColumn = 12
                for bs in BSs.json()['records']:
                    Sheet.Range( index(MainColumn,True) + str(MainRow) + ":" + index(MainColumn,False) + str(MainRow) ).Interior.Color = rgbToInt((6,121,200))
                    Sheet.Range( index(MainColumn,True) + str(MainRow) + ":" + index(MainColumn,False) + str(MainRow) ).Font.Color = rgbToInt((255,255,255) )
                    #Sheet.Cells( MainRow , MainColumn ).Formula = '=SUM(' + index(MainColumn,True) + '4:' + index(MainColumn,True) + str(MainRow - 1) + ')' #Each Supplier
                    MainColumn += 1
                    Sheet.Cells( MainRow , MainColumn ).Formula = '=SUM(' + index(MainColumn,True) + '4:' + index(MainColumn,True) + str(MainRow - 1) + ')' #Total Supplier
                    MainColumn += 6
                    Sheet.Cells( MainRow , MainColumn ).Formula = 1 - Sheet.Cells( MainRow , MainColumn - 6 ).Value / Sheet.Cells( MainRow , 8 ).Value
                    Sheet.Cells( MainRow , MainColumn ).Style = 'Percent'
                    MainColumn += 1
                    Sheet.Cells( MainRow , MainColumn ).Formula = 1 - Sheet.Cells( MainRow , MainColumn - 7 ).Value / Sheet.Cells( MainRow , 10 ).Value
                    Sheet.Cells( MainRow , MainColumn ).Style = 'Percent'
                    MainColumn += 2

            Sheet.Range("2:2").Font.Bold  = True
            Sheet.Range("3:3").Font.Bold  = True
            Sheet.Range(str(MainRow) + ":" + str(MainRow) ).Font.Bold  = True
            Sheet.Columns.AutoFit()
            Sheet.Range("A:A").ColumnWidth = 2
            Sheet.Range("K:K").ColumnWidth = 2
            Sheet.Range("U:U").ColumnWidth = 2
            Sheet.Range("E:E").ColumnWidth = 30
            Sheet.Range("E:E").WrapText = True
            Sheet.Range("C:C").ColumnWidth = 15
            Sheet.Range("C:C").WrapText = True
            Sheet.Columns.HorizontalAlignment  = excel.constants.xlCenter
            Sheet.Columns.VerticalAlignment   = excel.constants.xlCenter
            Sheet.Range("K4").Select()
            Excel.ActiveWindow.FreezePanes = True
            Workbook.SaveAs(os.getcwd() + '\Approval Report ' + str(GPS) + ' - ' + Byt + '.xlsx')
            Workbook.Close()
            Excel.Application.Quit()

            tkinter.messagebox.showinfo("Saved", "Report Generated")
        else :
            tkinter.messagebox.showerror("Not Found", "No Records Found")
    else :
        tkinter.messagebox.showerror("GPS Length", "Longer GPS Number Required")
    
def rgbToInt(rgb):
    colorInt = rgb[0] + (rgb[1] * 256) + (rgb[2] * 256 * 256)
    return colorInt

def warranties(w):
    match w:
        case "M1218":
            return "12/18 Months"
        case "M1224":
            return "12/24 Months"
        case "Y3":
            return "3 Years"
        case "Y4":
            return "4 Years"
        case "Y5":
            return "5 Years"
        
def description(d):
    if d['buyout_type_c'] == 'Motor':
        text1 = d['motor_standard_c'] + ' ' + d['motor_power_rating_c'] + ' ' + d['motor_speed_c'] + ' ' + d['motor_voltage_c']
        text2 = d['motor_frequency_c'] + ' ' + d['motor_mounting_c'] + ' ' + d['motor_frame_size_c'] + ' ' + d['motor_area_class_c']
        text3 = d['motor_enclosure_c'] + ' ' + d['motor_ip_rating_c'] + ' ' + d['motor_cooling_method_c'] + ' ' + d['motor_insulation_class_c']
        text4 = d['motor_service_factor_c'] + ' ' + d['motor_efficiency_c'] + ' ' + d['motor_starting_method_c'] + ' ' + d['motor_norm_c'] + ' ' + d['motor_certification_c']
        return  text1 + ' ' + text2 + ' ' + text3 + ' ' + text4
    if d['buyout_type_c'] == 'Coupling':
        text1 = d['coup_type_c']  + ' ' +  d['coup_nomenclature_c']  + ' ' +  d['coup_dbse_c']
        text2 = d['coup_balancing_c']  + ' ' +  d['coup_material_c']  + ' ' +  d['coup_description_c']
        return  text1 + ' ' + text2
    if d['buyout_type_c'] == 'Seal':
        text1 = d['mechseal_type_c'] + ' ' + d['mechseal_arrangement_c'] + ' ' + d['mechseal_shaft_size_c']
        text2 = d['mechseal_material_code_c'] + ' ' + d['mechseal_api_code_c'] + ' ' + d['mechseal_api_plan_c'] + ' ' + d['mechseal_description_c']
        return text1 + ' ' + text2
    if d['buyout_type_c'] == 'Seal System':
        text1 = d['sealsys_api_plan_c'] + ' ' + d['sealsys_capacity_c'] + ' ' + d['sealsys_material_c']
        text2 = d['sealsys_api_edition_c'] + ' ' + d['sealsys_description_c']
        return text1 + ' ' + text2

def incoterms(it):
    match it:
        case "1":
            return "ALR"
        case "2":
            return "CIF"
        case "3":
            return "CRF"
        case "4":
            return "CIF"
        case "5":
            return "CIP"
        case "6":
            return "CTP"
        case "7":
            return "DAF"
        case "8":
            return "DAP"
        case "9":
            return "DAT"
        case "10":
            return "DDP"
        case "11":
            return "DDU"
        case "12":
            return "DEQ"
        case "13":
            return "DES"
        case "14":
            return "EXW"
        case "15":
            return "FA1"
        case "16":
            return "FAD"
        case "17":
            return "FAS"
        case "18":
            return "FOB"
        case "19":
            return "FCA"
        case "20":
            return "FDC"
        case "21":
            return "FDP"
        case "22":
            return "FOB"
        case "23":
            return "FOC"
        case "24":
            return "FOP"
        case "25":
            return "OTH"
        case "26":
            return "EXW"
        case "27":
            return "FCA"
        case "28":
            return "FAS"
        case "29":
            return "FOB"
        case "30":
            return "CFR"
        case "31":
            return "CIF"
        case "32":
            return "CPT"
        case "33":
            return "CIP"
        case "34":
            return "DAP"
        case "35":
            return "DPU"
        case "36":
            return "DDP"
        case "":
            return ""
        
def index(column,start):
    diff = 0
    orden =0

    if start == True :
        x = 64
        if column in range(27,53): #Column A
            diff = 26
            orden = 1
        elif column in range(53,79): #Column B
            diff = 52
            orden = 2
        elif column in range(79,105): #Column C
            diff = 78
            orden = 3
        elif column in range(105,131): #Column D
            diff = 104
            orden = 4
        elif column in range(131,157): #Column E
            diff = 130
            orden = 5
        elif column in range(157,182): #Column F
            diff = 156
            orden = 6

        if column < 27:
            return chr(x + column)
        else:
            return chr(x + orden) + chr(x + (column - diff))
    else :
        x = 72
        if column in range(20,46): #Column A
            diff = 26
            orden = 1
        elif column in range(46,72): #Column B
            diff = 52
            orden = 2
        elif column in range(72,98): #Column C
            diff = 78
            orden = 3
        elif column in range(98,124): #Column D
            diff = 104
            orden = 4
        elif column in range(124,150): #Column E
            diff = 130
            orden = 5
        elif column in range(150,175): #Column F
            diff = 156
            orden = 6
        
        if column < 20:
            return chr(x + column)
        else:
            return chr(x + orden - 8) + chr(x + (column - diff))


options = ["Motor",
	"Coupling",
	"Turbine",
	"VFD",
	"Seal",
	"Seal System",
	"Other"]

Byt = StringVar()
Byt.set( "Motor" )

#ITEMS
gps_label = ttk.Label(frame, text="GPS")
buyout_type_label = ttk.Label(frame, text="Type")

entry = ttk.Entry(frame, width=15)
drop = OptionMenu( frame , Byt , *options )

generate_button = ttk.Button(frame, text="Generate", width=15,command = lambda:generate_report(entry.get(),Byt.get()))
close_button = ttk.Button(frame, text="Close", width=15,command=root.destroy)

#GRID
gps_label.grid(pady=0,padx=5,column=0, row=0)
buyout_type_label.grid(pady=0,padx=5,column=1, row=0)

entry.grid(pady=5,padx=5,column=0, row=1)
drop.grid(pady=5,padx=5,column=1, row=1)

generate_button.grid(column=0, row=2)
close_button.grid(column=1, row=2)

root.mainloop()
