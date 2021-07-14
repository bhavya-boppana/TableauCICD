# -*- coding: utf-8 -*-
"""
Created on Mon Jul 12 19:10:13 2021

@author: bhavya boppana
"""

import tableau_api_lib
from tableau_api_lib import TableauServerConnection
from tableau_api_lib.utils.querying import get_projects_dataframe,get_views_dataframe, get_view_data_dataframe
from tableau_api_lib.utils.common import flatten_dict_column
import random
import math
import pandas as pd
import Move_workbook
import sys
import os
sys.stdout.flush()


tableau_server_config = {
        'my_env': {
                'server': 'http://tableauserver.eastus2.cloudapp.azure.com',
                'api_version': '3.12',
                'username': 'Madhes',
                'password': 'Admin@123',
                'site_name': 'TableauDEV',
                'site_url': 'TableauDEV'
        }
}
        
        
conn = TableauServerConnection(tableau_server_config, env='my_env')
res=conn.sign_in()
print("sign in:",res)
ResDetailsFile = open("testing_details.txt","w")


def replace_chars(filter_name):
    new=filter_name.replace(' ','%20')
    return new

def filter_utility(sheet_id,filter_df,sheet_name):
    filter_name=replace_chars(filter_df.columns[0])
    correct_col=filter_df.columns[1]
    identifier_col=filter_df.columns[2]
    for (filter_val,correct_val,identifier_val) in zip(filter_df[filter_df.columns[0]],filter_df[correct_col],filter_df[identifier_col]):
        params_dict={"filter":f"vf_{filter_name}={filter_val}"}
        sheet_df=get_view_data_dataframe(conn, view_id=sheet_id,parameter_dict=params_dict)
        print(sheet_df[identifier_col])
        print('\n\n\n')
        print(identifier_val)
        record=sheet_df.loc[sheet_df[identifier_col]==identifier_val][correct_col]
        key=record.keys()[0]
        if(record[key]!=correct_val):
            ResDetailsFile.write(f"{filter_df.columns[0]} filter test did not pass on {sheet_name} because for filter value:{filter_val},{identifier_val} value is returned as {record[key]}, when it should be {correct_val} ")
            return False
    return True


def filter_test(sheet_id,excel,location_string,sheet_name):
    locations=location_string.split(';')
    res = True
    for filter_num in range(len(locations)):
        coordinates=locations[filter_num].split(',')
        xl_sheet_num=int(coordinates[0])
        start_col=int(coordinates[1])
        end_col=start_col+3
        if filter_utility(sheet_id,excel[xl_sheet_num].iloc[:,start_col:end_col], sheet_name)==False:
            res=False
    return res


def expected_val_utility(sheet_id,checking_df,sheet_name):
    identifier_col=checking_df.columns[0]
    checking_col=checking_df.columns[1]
    sheet_df=get_view_data_dataframe(conn,view_id=sheet_id)
    for (identifier_val,checking_val) in zip(checking_df[identifier_col],checking_df[checking_col]):
        record=sheet_df.loc[sheet_df[identifier_col]==identifier_val][checking_col]
        key=record.keys()[0]
        if(record[key]!=checking_val):
            ResDetailsFile.write(f"expected value test did not pass on {sheet_name} because for {identifier_val}, {checking_col} value is returned as {record[key]}, when it should be {checking_val} ")
            return False
    return True


def expected_val_test(sheet_id,excel,location_string,sheet_name):
    locations=location_string.split(';')
    res=True
    for test_num in range(len(locations)):
        coordinates=locations[test_num].split(',')
        xl_sheet_num=int(coordinates[0])
        start_col=int(coordinates[1])
        end_col=start_col+2
        if expected_val_utility(sheet_id,excel[xl_sheet_num].iloc[:,start_col:end_col], sheet_name)==False:
            res=False
    return res


def divide_by_zero(sheet_id,sheet_name):
    sheet_df=get_view_data_dataframe(conn,view_id=sheet_id)
    for col in sheet_df.columns:
        for val in sheet_df[col]:
            if not isinstance(val,str) and math.isnan(val):
                ResDetailsFile.write(f"divide by zero test did not pass on {sheet_name} because there are one or more divide by zero cases found in the column:{col}")
                return False
    return True


def Null_checking(sheet_id,sheet_name):
    sheet_df=get_view_data_dataframe(conn,view_id=sheet_id)
    for col in sheet_df.columns:
        for val in sheet_df[col]:
            if not isinstance(val,str) and math.isnan(val):
                ResDetailsFile.write(f"null value checking test did not pass on {sheet_name} because there are one or more null values found in the column:{col}")
                return False
    return True


def test():
   
    NameFile=open("workbookname.txt","r+")
    wbname = NameFile.read()
    print("wbname: ",wbname)
    site_views_df = get_views_dataframe(conn)
    site_views_detailed_df = flatten_dict_column(site_views_df, keys=['name', 'id'], col_name='workbook')
    df = site_views_detailed_df[site_views_detailed_df['workbook_name'] == wbname]
    
    excel_name = str(str(wbname) + ".xlsx")
    path = "C:\\TableauTestResults\\TestCaseDetails"
    xl_sheet_count=len(pd.ExcelFile(os.path.join(path,excel_name)).sheet_names)
    excel=pd.read_excel(os.path.join(path,excel_name),list(range(xl_sheet_count)))
    xl_sheet1=excel[0]
    sheet_names=xl_sheet1['Sheet name']
    sheet_ids=[]
    for sheet_name in sheet_names:
        print(sheet_name)
        sheet_row=df.loc[df['name']==sheet_name]['id']
        key=sheet_row.keys()[0]
        sheet_ids.append(sheet_row[key])
   
    all_passed=True
    tests_done=False
    res_df=pd.DataFrame(columns=['     Sheet Name            ',' Filter functionality checking ',' expected value checking ',' divide by zero checking ',' Null value checking'])
    row=["----         "]*int(5)
    for sheet_name in sheet_names:
        row[0]=sheet_name+'      '
        res_df.loc[len(res_df)]=row
   
    #filter_checking
    for (i,sheet_id) in zip(range(len(sheet_ids)),sheet_ids):
        if(xl_sheet1.iloc[i,1].lower()=="yes"):
            tests_done=True
            filter_test_res=filter_test(sheet_id,excel,xl_sheet1.iloc[i,2],xl_sheet1.iloc[i,0])
            if filter_test_res==False:
                all_passed=False
                res_df.iloc[i,1]="Failed"
            else:
                res_df.iloc[i,1]="Passed"

    #expected value checking
    for (i,sheet_id) in zip(range(len(sheet_ids)),sheet_ids):
        if(xl_sheet1.iloc[i,3].lower()=="yes"):
            tests_done=True
            expected_val_res=expected_val_test(sheet_id,excel,xl_sheet1.iloc[i,4],xl_sheet1.iloc[i,0])
            if expected_val_res==False:
                all_passed=False
                res_df.iloc[i,2]="Failed"
            else:
                res_df.iloc[i,2]="Passed"
    #divide by zero checking
    for (i,sheet_id) in zip(range(len(sheet_ids)),sheet_ids):
        if(xl_sheet1.iloc[i,5].lower()=="yes"):
            tests_done=True
            check_res=divide_by_zero(sheet_id,xl_sheet1.iloc[i,0])
            if check_res==False:
                all_passed=False
                res_df.iloc[i,3]="Failed"
            else:
                res_df.iloc[i,3]="Passed"
               
    #Null value checking
    for (i,sheet_id) in zip(range(len(sheet_ids)),sheet_ids):
        if(xl_sheet1.iloc[i,6].lower()=="yes"):
            tests_done=True
            check_res=Null_checking(sheet_id, xl_sheet1.iloc[i, 0])
            if check_res==False:
                all_passed=False
                res_df.iloc[i,4]="Failed"
            else:
                res_df.iloc[i,4]="Passed"
                
    if tests_done == false:
        ResDetailsFile.write("No test cases were done please mention some")
        return
    with open("test_123.txt",'w') as outfile:
        res_df.to_string(outfile)
        outfile.write('\n\n\n')
        if(all_passed==True):
            outfile.write("All test cases were passed and workbook has been pushed to the Production server!!")
            Move_workbook.execute()
            ResDetailsFile.write("All test cases were passed and workbook has been pushed to the Production server!!")
        else:
            outfile.write("workbook could not be pushed to Production server as some of the test cases were not passed")
test()
