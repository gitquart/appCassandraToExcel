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
    auth_provider = PlainTextAuthProvider(objCC.cc_user,objCC.cc_pwd)
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
    cc_keyspace='thesis'
    cc_pwd='P@ssw0rd33'
    cc_databaseID='9de16523-0e36-4ff0-b388-44e8d0b1581f'