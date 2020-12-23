"""
Program that will write cassandra table in excel workbook

"""
from openpyxl import Workbook
from openpyxl import load_workbook
import os
import database as bd


dir_excel='C:\\quartExcel\\impi1.xlsx'
current_dir = os.getcwd()

def main():
    dexist=False
    dexist=os.path.exists('C:\\quartExcel')
    if dexist==False:
        os.mkdir('C:\\quartExcel')
    wb = Workbook()
    ws = wb.active
    ws.title = "Impi1" 
    lsFields=[]
    #Get cassandra columns
    table='thesis.impi_docs'
    query="select column_name from system_schema.columns WHERE keyspace_name = 'thesis' AND table_name = 'impi_docs';"
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
        resultSet=''
        query="select "+fieldsForQuery+" from "+table+" where period_number="+str(i)+""
        bd.getLargeQueryAndPrintToExcel(query)


    print('The excel is ready!')        
                         
            
                    

if __name__=='__main__':
    main()    
   