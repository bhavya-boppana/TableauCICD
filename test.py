# -*- coding: utf-8 -*-
"""
Created on Mon Jun 21 16:31:37 2021

@author: bhavya boppana
"""
import sys
sys.stdout.flush()
import tableau_api_lib
from tableau_api_lib import TableauServerConnection
from tableau_api_lib.utils.querying import get_projects_dataframe,get_views_dataframe, get_view_data_dataframe
from tableau_api_lib.utils.common import flatten_dict_column
import random
import math
import pandas as pd
import Move_workbook

tableau_server_config = {
        'my_env': {
                'server': 'http://tableauserver.eastus2.cloudapp.azure.com',
                'api_version': '3.11',
                'username': 'Madheswaran',
                'password': 'Admin1234@',
                'site_name': 'TableauDEV',
                'site_url': 'TableauDEV'
        }
}
   
conn = TableauServerConnection(tableau_server_config, env='my_env')
res=conn.sign_in()
print("sign in:",res)

def replace_chars(filter_name):
    new=filter_name.replace(' ','%20')
    return new

def filter_test(conn,sheet_id,filter_df):
    filter_name=replace_chars(filter_df.columns[0])
    field_name=filter_df.columns[1]
    row=filter_df.columns[2]
    for (filter_val,correct_val,row_val) in zip(filter_df[filter_df.columns[0]],filter_df[field_name],filter_df[row]):
        params_dict={"filter":f"vf_{filter_name}={filter_val}"}
        sheet_df=get_view_data_dataframe(conn, view_id=sheet_id,parameter_dict=params_dict)
        record=sheet_df.loc[sheet_df[row]==row_val][field_name]
        key=record.keys()[0]
        if(record[key]!=correct_val):
            return False
    return True 

def expected_val_checking(conn,sheet_id,checking_df):
    ref_col=checking_df.columns[0]
    checking_col=checking_df.columns[1]
    sheet_df=get_view_data_dataframe(conn,view_id=sheet_id)
    for (ref_val,checking_val) in zip(checking_df[ref_col],checking_df[checking_col]):
        record=sheet_df.loc[sheet_df[ref_col]==ref_val][checking_col]
        key=record.keys()[0]
        if(record[key]!=checking_val):
            return False
    return True
    
def divide_by_zero(conn,sheet_id):
    sheet_df=get_view_data_dataframe(conn,view_id=sheet_id)
    for col in sheet_df.columns:
        for val in sheet_df[col]:
            if not isinstance(val,str) and math.isnan(val):
                return False
    return True
    
def Null_checking(conn,sheet_id):
    sheet_df=get_view_data_dataframe(conn,view_id=sheet_id)
    for col in sheet_df.columns:
        for val in sheet_df[col]:
            if not isinstance(val,str) and math.isnan(val):
                return False
    return True

def test():
    file1 = open("testing_results.txt","w")

    NameFile=open("workbookname.txt","r+")
    s=NameFile.read()

    site_views_df = get_views_dataframe(conn)
    site_views_detailed_df = flatten_dict_column(site_views_df, keys=['name', 'id'], col_name='workbook')
    df = site_views_detailed_df[site_views_detailed_df['workbook_name'] == s]
        
    xl=pd.ExcelFile(r'tableau2.xlsx')
    sheet_count=len(xl.sheet_names)
    excel=pd.read_excel(r'tableau2.xlsx',list(range(sheet_count)))
    sheet_details=excel[0]
    sheet_names=sheet_details['Sheet name']
    test_vals=sheet_details['filters']
    ids=[]
    test_dict={1:"filter_test",2:"expected_value_checking",3:"divide_by_zero_checking",4:"Null_checking"}
    test_results=[]
    res_df=pd.DataFrame(columns=['     Sheet Name            ',' Filter functionality checking ',' expected value checking ',' divide by zero checking ',' Null value checking'])
    for (sheet_name,test_val) in zip(sheet_names,test_vals):
        row=["----         "]*int(5)
        row[0]=sheet_name+'      '
        record=df.loc[df['name']==sheet_name]['id']
        key=record.keys()[0]
        ids.append(record[key])
        l=[]
        if (isinstance(test_val, str)):
            l=test_val.split(',')
            l=[int(i) for i in l]
        else:
            l.append(test_val)
        for i in l:
            test_res=False
            if(i==1):
                if(len(excel)<2):
                    file1.write(f"cannot do filter test on {sheet_name} as the excel sheet does not have required sheet \n")
                    print(f"cannot do filter test on {sheet_name} as the excel sheet does not have required sheet")
                else:  
                    test_res=filter_test(conn,ids[-1],excel[1])
            elif (i==2):
                if(len(excel)<3):
                    file1.write(f"cannot do expected value checking test on {sheet_name} as the excel does not have required sheet \n")
                    print(f"cannot do expected value checking test on {sheet_name} as the excel does not have required sheet")
                else:  
                    test_res=expected_val_checking(conn,ids[-1],excel[2])
            elif (i==3):
                test_res=divide_by_zero(conn,ids[-1])
            elif(i==4):
                test_res=Null_checking(conn,ids[-1])
            else:
                file1.write(f"{i} is invalid test number \n")
                print(f"{i} is invalid test number")
            if i<=4 and i>=1:
                if test_res:
                    file1.write(f"{test_dict[i]} passed on {sheet_name} \n")
                    space_count=int((len(df.columns[i])-6)/2)
                    row[i]="passed         "
                    print(f"{test_dict[i]} passed on {sheet_name}")
                else:
                    file1.write(f"{test_dict[i]} did not pass on {sheet_name} \n")
                    space_count=int((len(df.columns[i])-6)/2)
                    row[i]="Failed         "
                    print(f"{test_dict[i]} did not pass on {sheet_name}")
                test_results.append(test_res)
        res_df.loc[len(res_df)]=row
    with open("test_123.txt",'w') as outfile:
        res_df.to_string(outfile)
            
    for i in test_results:
        if not i:
            file1.write("Workbook could not be pushed as some of the test cases have not been passed \n")
            file1.close() 
            print("Workbook could not be pushed as some of the test cases have not been passed")
            return
        
    if len(test_results)==0:
        file1.write("No tests were done on this workbook. Please mention the cases to check")
        file1.close()
        return
    Move_workbook.execute()
    file1.write("All test cases were passed!! \n")
    print("All test cases were passed!!")
    file1.close() 
        
    
test()
