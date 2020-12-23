"""
Program that will write cassandra table in excel workbook

"""
from openpyxl import Workbook
from openpyxl import load_workbook
import os
import database as bd
from InternalControl import cInternalControl

objControl= cInternalControl()
dir_excel=objControl.excel_dir+objControl.excel_file

def main():
    dexist=False
    dexist=os.path.exists(objControl.excel_dir)
    if dexist==False:
        os.mkdir(objControl.excel_dir)
    wb = Workbook()
    ws = wb.active
    title=objControl.excel_sheet
    ws.title = title
    lsFields=[]
    #Get cassandra columns
    keyspace=objControl.keyspace
    table=objControl.table
    query="select column_name from system_schema.columns WHERE keyspace_name = '"+keyspace+"' AND table_name = '"+table+"';"
    columns_list=''
    columns_list=bd.getShortQuery(query)
    coln=1
    for col in columns_list:
        #Write(row,column)
        #Headers (h1,...)
        h1 = ws.cell(row = 1, column = coln)
        h1.value = col[0]
        lsFields.append(col[0])
        coln=coln+1
        wb.save(dir_excel) 
    #Reading list of fields into commas for the next query
    fieldsForQuery=','.join(lsFields)    
    #Starts  the reading of periods    
    flag=os.path.isfile(dir_excel)
    if flag:
        #Expedient xls already exists
        print('Printing excel... ')
        query="select "+fieldsForQuery+" from "+table+" where year>0 ALLOW FILTERING"
        bd.getLargeQueryAndPrintToExcel(query,dir_excel,title)


    print('The excel is ready!')        
                         
            
                    

if __name__=='__main__':
    main()    
   