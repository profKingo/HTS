import os
import pyodbc
import os
print("Ciao")
nomedb="C:\\THS\\THS32Env\\wcArchivi.mdb"

connstr=f'Driver={{Microsoft Access Driver (*.mdb)}};Dbq=' + nomedb + ';Uid=;Pwd=;'
try:
    conn=pyodbc.connect(connstr)
    strSQL = "SELECT * FROM CARVEI;"
    carvei_cursor = conn.cursor()
    carvei_cursor.execute(strSQL)  #creo un cursore/recordset(cursor) da una query
    rows = carvei_cursor.fetchall() #recupera il risultato della query in execute e lo mette in una lista 
        
except:
    print("Errore di connessione al database!", "Errore accesso")


