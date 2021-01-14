from cassandra.cluster import Cluster
from cassandra.auth import PlainTextAuthProvider
from cassandra.query import SimpleStatement
from openpyxl import Workbook
from openpyxl import load_workbook
from InternalControl import cInternalControl

objControl= cInternalControl()
cloud_config= {
        'secure_connect_bundle':'secure-connect-'+objControl.db+'.zip'
    }


def getCluster():
    #Connect to Cassandra
    objCC=CassandraConnection()
    user=''
    password=''
    if objControl.db=='dbquart':
        user=objCC.cc_user
        password=objCC.cc_pwd
    else:
        user=objCC.cc_user_test
        password=objCC.cc_pwd_test

    auth_provider = PlainTextAuthProvider(user,password)
    cluster = Cluster(cloud=cloud_config, auth_provider=auth_provider)

    return cluster

def getLargeQueryAndPrintToExcel(query,dir_excel,title):
    cluster = getCluster()
    session = cluster.connect()
    session.default_timeout=70     
    statement = SimpleStatement(query, fetch_size=1000)
    wb = load_workbook(dir_excel)
    ws = wb[title]
        
    for row in session.execute(statement):
        ls=[]
        for col in row:
            ls.append(str(col))
        ws.append(ls)
        
    print('Total rows:',str(count_row))    
    wb.save(dir_excel) 
    cluster.shutdown() 


def getShortQuery(query):
    res=''
    cluster=getCluster()
    session = cluster.connect()
    session.default_timeout=70
    #Check wheter or not the record exists      
    future = session.execute_async(query)
    res=future.result()
    cluster.shutdown()

    return res 


                

     
class CassandraConnection():
    cc_user='quartadmin'
    cc_pwd='P@ssw0rd33'
    cc_user_test='test'
    cc_pwd_test='testquart'