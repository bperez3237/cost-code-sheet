from numpy import number
import xlsxwriter as xl
import pandas as pd
from datetime import date
import pprint as pp
import json
from formats import *

cost_rprt_xls = pd.ExcelFile(r'C:\Users\bperez\Iovino Enterprises, LLC\M007-NYCHA-Coney Island Sites - Documents\General\08 - BUDGET & COST\Cost Codes\Contract Forecasting Spreadsheet\Period 9 Export 09.19.22.xlsx')
cost_rprt = pd.read_excel(cost_rprt_xls)


with open('cc_db.json') as f:
    cc_db = json.load(f)

    
row = 5

def phase_codes():
    codes = cost_rprt.loc[:,'Phase'].unique()
    return codes

def production_rate(quantity, manhours):
    return 1 if manhours == 0 else quantity/manhours

def labor_rate(code):
    if cost_rprt[cost_rprt['Phase']==code].iloc[0]['Input Completed Qty'] != 0:
        return cost_rprt[cost_rprt['Phase']==code].iloc[0]['Actual Cost']/cost_rprt[cost_rprt['Phase']==code].iloc[0]['Input Completed Qty']
    else:
        return 0

def create_heading(workbook, worksheet):
    worksheet.merge_range('A1:C1','M007 - GR1904063 - Coney Island Sites', string_format(workbook,'#366092', True ))

    worksheet.merge_range('E1:F1', 'Contract Value', string_format(workbook,'#366092', True ))
    worksheet.merge_range('E2:F2', 'Projected Cost', string_format(workbook,'#366092', True ))
    worksheet.merge_range('J1:K1', 'Total Mhs', string_format(workbook,'#366092', True ))
    worksheet.merge_range('J2:K2', 'Mhs to Date', string_format(workbook,'#366092', True ))

    worksheet.merge_range('O3:T3', 'Category Total', string_format(workbook,'#366092', True )) 

    worksheet.write('A5', 'Code', string_format(workbook,'#366092', True ))
    worksheet.write('B5', 'Name', string_format(workbook,'#366092', True ))
    worksheet.write('C5', 'Qty', string_format(workbook,'#366092', True ))
    worksheet.write('D5', 'UOM', string_format(workbook,'#366092', True ))
    worksheet.write('E5', 'Mhs', string_format(workbook,'#366092', True ))
    worksheet.write('F5', 'Qty/MH', string_format(workbook,'#366092', True ))
    worksheet.write('G5', 'Projected Forecast', string_format(workbook,'#366092', True ))

    worksheet.write('I5', 'Spent to Date', string_format(workbook,'#366092', True ))
    worksheet.write('J5', 'Committed to Date', string_format(workbook,'#366092', True ))
    worksheet.write('K5', 'Qty to Date', string_format(workbook,'#366092', True ))
    worksheet.write('L5', 'Mhs to Date', string_format(workbook,'#366092', True ))
    worksheet.write('M5', 'Labor Rate', string_format(workbook,'#366092', True ))

    worksheet.write('O5', 'Labor', string_format(workbook,'#366092', True ))
    worksheet.write('P5', 'Subcontract', string_format(workbook,'#366092', True ))
    worksheet.write('Q5', 'Consumable', string_format(workbook,'#366092', True ))
    worksheet.write('R5', 'Permanent Material', string_format(workbook,'#366092', True ))
    worksheet.write('S5', 'Equipment', string_format(workbook,'#366092', True ))
    worksheet.write('T5', 'Other', string_format(workbook,'#366092', True ))
    
    worksheet.write('V5', 'System Projected Cost', string_format(workbook,'#366092', True ))
    worksheet.write('W5', 'Variance', string_format(workbook,'#366092', True ))
    worksheet.write('X5', '& Variance', string_format(workbook,'#366092', True ))


def write_sub_code(workbook, worksheet,code):
    global row
    row +=1

    color = '#DCE6F1'
    for index, area in enumerate(cc_db[code]):
        sub_code = code+'-'+area
        sub_name = cost_rprt[cost_rprt['Phase']==code].iloc[0]['Name']+'-'+area
        sub_forecast = cost_rprt[cost_rprt['Phase']==code]['Projected Cost Forecast'].sum()
        # print(cost_rprt[(cost_rprt['Phase']==code) & (cost_rprt['Category']=='L')])

        worksheet.write(row,0,sub_code, string_format(workbook,color))
        worksheet.write(row,1,sub_name, string_format(workbook,color))
        worksheet.write(row,2,cc_db[code][area]['forecast_qty'], number_format(workbook, color))
        worksheet.write(row,3,'NA' if pd.isnull(cost_rprt[cost_rprt['Phase']==code]['Output WM Code'].iloc[0]) else cost_rprt[cost_rprt['Phase']==code]['Output WM Code'].iloc[0], string_format(workbook,color))
        worksheet.write(row,4,cc_db[code][area]['forecast_mhs'], number_format(workbook, color))
        worksheet.write(row,5,1 if cc_db[code][area]['forecast_mhs']==0 else f"=C{row+1}/E{row+1}", number_format(workbook, color))
        worksheet.write(row,6,sub_forecast, currency_format(workbook, color))
        worksheet.write(row,8,cc_db[code][area]['current_mhs']*labor_rate(code), currency_format(workbook, color))
        worksheet.write(row,9,0, currency_format(workbook, color)) #commited amount N/A
        worksheet.write(row,10,cc_db[code][area]['current_qty'], number_format(workbook,color))
        worksheet.write(row,11,cc_db[code][area]['current_mhs'], number_format(workbook, color))
        worksheet.write(row,12,f"=M{row-index}", currency_format(workbook, color))

    
        worksheet.write(row,14,0,currency_format(workbook, color))
        worksheet.write(row,15,0,currency_format(workbook, color))
        worksheet.write(row,16,0,currency_format(workbook, color))
        worksheet.write(row,17,0,currency_format(workbook, color))
        worksheet.write(row,18,0,currency_format(workbook, color))
        worksheet.write(row,19,0,currency_format(workbook, color))

        worksheet.write(row,21,0,currency_format(workbook, color))
        worksheet.write(row,22,0,currency_format(workbook, color))
        worksheet.write(row,23,0,currency_format(workbook, color))
        
        worksheet.set_row(row, None, None, {'level':1, 'hidden': True})
        row +=1 


def write_categories(workbook, worksheet, code_df):
    worksheet.write(row,14,code_df[code_df['Category']=='L']['Projected Cost Forecast'] if 'L' in code_df['Category'].values else 0 ,currency_format(workbook, 'white'))
    worksheet.write(row,15,code_df[code_df['Category']=='S']['Projected Cost Forecast'] if 'S' in code_df['Category'].values else 0 ,currency_format(workbook, 'white'))
    worksheet.write(row,16,code_df[code_df['Category']=='C']['Projected Cost Forecast'] if 'C' in code_df['Category'].values else 0 ,currency_format(workbook, 'white'))
    worksheet.write(row,17,code_df[code_df['Category']=='M']['Projected Cost Forecast'] if 'M' in code_df['Category'].values else 0 ,currency_format(workbook, 'white'))
    worksheet.write(row,18,code_df[code_df['Category']=='E']['Projected Cost Forecast'] if 'E' in code_df['Category'].values else 0,currency_format(workbook, 'white') )
    worksheet.write(row,19,f"=G{row+1}-SUM(O{row+1}:S{row+1})",currency_format(workbook, 'white'))

def write_system_projection(workbook,worksheet):
    worksheet.write(row,21,f"=I{row+1} + M{row+1}*((C{row+1}-K{row+1})/F{row+1})",currency_format(workbook, 'white'))
    worksheet.write(row,22,f"=V{row+1}-G{row+1}",currency_format(workbook, 'white'))
    worksheet.write(row,23,f"=IF(G{row+1}=0,0,W{row+1}/G{row+1})", number_format(workbook, 'white'))

def write_code_data(workbook, worksheet, code):
    global row

    code_df = cost_rprt.loc[cost_rprt['Phase'] == code]
    uom = 'NA' if pd.isnull(code_df['Output WM Code'].iloc[0]) else code_df['Output WM Code'].iloc[0]

    worksheet.write(row,0,str(code), string_format(workbook,'white'))
    worksheet.write(row,1,code_df['Name'].iloc[0], string_format(workbook,'white'))
    worksheet.write(row,2,code_df['Output Projected Qty'].iloc[0],number_format(workbook, 'white'))
    worksheet.write(row,3, uom, string_format(workbook,'white'))
    worksheet.write(row,4, code_df[code_df['Category']=='L']['Input Projected Qty'] if 'L' in code_df['Category'].values else 0 ,number_format(workbook, 'white')) 
    worksheet.write(row,5, f"=IF(OR(E{row+1}=0,C{row+1}=0),1,C{row+1}/E{row+1})",number_format(workbook, 'white'))
    worksheet.write(row,6,code_df['Projected Cost Forecast'].sum(),currency_format(workbook, 'white'))

    worksheet.write(row,8,code_df['Actual Cost'].sum(),currency_format(workbook, 'white'))
    worksheet.write(row,9,code_df['Spent/Committed Total'].sum(),currency_format(workbook, 'white'))
    worksheet.write(row,10,code_df['Output Completed Qty'].iloc[0],number_format(workbook, 'white'))
    worksheet.write(row,11,code_df[code_df['Category']=='L']['Input Completed Qty'] if 'L' in code_df['Category'].values else 0,number_format(workbook, 'white'))
    worksheet.write(row,12,code_df[code_df['Category']=='L']['Actual Cost']/code_df[code_df['Category']=='L']['Input Completed Qty'] if 'L' in code_df['Category'].values and code_df[code_df['Category']=='L']['Input Completed Qty'].iloc[0] != 0 else 108, currency_format(workbook,'white'))

    write_categories(workbook, worksheet, code_df)
    write_system_projection(workbook,worksheet)

    #write sub code data if exists
    if code in cc_db:
        write_sub_code(workbook, worksheet,code)
    else:
        row +=1

def add_body_data(workbook, worksheet):
    for code in phase_codes():
        write_code_data(workbook, worksheet, code)


def add_heading_data(workbook, worksheet):    
    worksheet.write('G1', 189_859_018.31,currency_format(workbook, 'white'))
    # worksheet.write('G2', cost_rprt['Projected Cost Forecast'].sum(),currency_format(workbook, 'white'))
    worksheet.write('G2',  f'=SUMIF($A${6}:$A${row},"**-****",G{6}:G{row})-SUMIF($A${6}:$A${row},"**-****-*",G{6}:G{row})', currency_format(workbook, 'white'))
    
    worksheet.write('L1',  f'=SUMIF($A${6}:$A${row},"**-****",E{6}:E{row})-SUMIF($A${6}:$A${row},"**-****-*",E{6}:E{row})', number_format(workbook, 'white'))
    worksheet.write('L2',  f'=SUMIF($A${6}:$A${row},"**-****",L{6}:L{row})-SUMIF($A${6}:$A${row},"**-****-*",L{6}:L{row})', number_format(workbook, 'white'))

    worksheet.write('O4', f'=SUM(O{6}:O{row})',currency_format(workbook, 'white'))
    worksheet.write('P4', f'=SUM(P{6}:P{row})',currency_format(workbook, 'white'))
    worksheet.write('Q4', f'=SUM(Q{6}:Q{row})',currency_format(workbook, 'white'))
    worksheet.write('R4', f'=SUM(R{6}:R{row})',currency_format(workbook, 'white'))
    worksheet.write('S4', f'=SUM(S{6}:S{row})',currency_format(workbook, 'white'))
    worksheet.write('T4', f'=SUM(T{6}:T{row})',currency_format(workbook, 'white'))


def write_sheet(workbook, worksheet):
    add_body_data(workbook, worksheet)
    create_heading(workbook, worksheet)
    add_heading_data(workbook, worksheet)

def format_workbook(worksheet):
    gap_width = 2.5
    worksheet.set_column(1,1 ,45)
    worksheet.set_column(2,2 ,14)
    worksheet.set_column(3,3 ,5)
    worksheet.set_column(4,5 ,11.5)
    worksheet.set_column(6,6 ,17.5)

    worksheet.set_column(7,7 ,gap_width)

    worksheet.set_column(8,9 ,16.5)
    worksheet.set_column(10,12 ,14)

    worksheet.set_column(13,13 ,gap_width)

    worksheet.set_column(14,19 ,20)
    
    worksheet.set_column(20,20 ,gap_width)
    worksheet.set_column(21,23 ,20)

    
    worksheet.set_column("M:M",None, None, {'level':1,'hidden': True})
    worksheet.set_column("O:T",None, None, {'level':1,'hidden': True})
    worksheet.set_column("V:X",None, None, {'level':1,'hidden': True})

def create_report():
    workbook = xl.Workbook('Cost Codes '+str(date.today())+'.xlsx')
    worksheet = workbook.add_worksheet('Code Codes')
    write_sheet(workbook, worksheet)
    format_workbook(worksheet)

    workbook.close()

create_report()

