import pyodbc
print(pyodbc.dataSources())
nomedb="c:\\HTS\\wcArchivi.mdb"
try:
    #global conn
    conn=pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb)};);Dbq=' + nomedb + ';Uid=;Pwd=;')
    print("ok")
except:
    print(conn.__doc__)

