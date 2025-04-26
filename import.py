import tkinter as tk
from tkinter.ttk import *
from tkinter import *
import json
import pyodbc
import os
from datetime import datetime
import struttura as s
    
'''
l'utente parte da file pdf da ocrizzare e quindi da inviare al portale di AIDA (tramite uno dei metodi tra cui WS)
i file pdf sono di tipo MECC o CARR e vengono nominati ad hoc...
pensare un batch che chiami appena generato il file la conversione e poi avvii la lettura del JSON risultante
il JSON viene rimandato all'utente come???? Webhook - FTP - AIDA LINK
'''
conn=pyodbc.Connection
nomedb="C:\\THS\\THS32Env\\wcArchivi.mdb"
filejson="export.json"
tipoPratica="M" #Meccanica #Carrozzieria
#stringa di connessione al db
def connetti():
    connstr=f'Driver={{Microsoft Access Driver (*.mdb)}};Dbq=' + nomedb + ';Uid=;Pwd=;'
    try:
        conn=pyodbc.connect(connstr)
    except:
        tk.messagebox.showerror("Errore di connessione al database!", "Errore accesso")

def leggi_par_ini():
    NOMEF="parametri.ini"
    if not os.path.exists(NOMEF):
        tk.messagebox.showerror("Errore: File dei parametri non esiste")
    else:
        archivio=open(NOMEF, 'r')
        riga=archivio.readline()
        while(riga!=""):
            nome = riga
            riga=archivio.readline()
        '''
        s.pratica.F_CODCLI = Modulo2.ini_manager("r", "CLIENTE", "f_codli")
        s.pratica.F_RAGSOC = Modulo2.ini_manager("r", "CLIENTE", "f_ragsoc")
        s.pratica.F_VIACLI = Modulo2.ini_manager("r", "CLIENTE", "f_viacli")
        s.pratica.F_CITTAC = Modulo2.ini_manager("r", "CLIENTE", "f_cittac")
        s.pratica.F_CAPCLI = Modulo2.ini_manager("r", "CLIENTE", "f_capcli")
        s.pratica.F_PROCLI = Modulo2.ini_manager("r", "CLIENTE", "f_procli")
        s.pratica.F_PARIVA = Modulo2.ini_manager("r", "CLIENTE", "f_pariva")
        s.pratica.F_TELEFO = Modulo2.ini_manager("r", "CLIENTE", "f_telefo")
        '''
        archivio.close()

def leggijson():
    data = json.load(open(filejson))
    if len(data)>=1:
        s.pratica.desc="Pratica"
        for pre in data:
            for x in s.s_header:    #INTESTAZIONE PREVENTIVO
                #qui dovrei leggere l'intestazione se si tratta di ALD o meno
                a=pre[x]
                if isinstance(a, str):
                    com="s.prev.%s = '%s'" % (x.replace(" ", "_").replace(".",""), pre[x].replace("'",""))
                else:
                    com="s.prev.%s = '%s'" % (x.replace(" ", "_").replace(".",""), pre[x])
                exec(com)
                #leggo le righe nel file e creo dei campi in una struttura che memorizzi i dati del preventivo
            #RIGHE PREVENTIVO
            try:
                tab=pre["Tabella Interventi Meccanica"]
            except:
                tab=[]
            righe=[]
            i=0   
            for e in tab:
                el=s.riga()
                for x in s.s_elem:
                    val=pre["Tabella Interventi Meccanica"][i][x]
                    if isinstance(val, str):
                        com='el.%s = "%s"' % (x.replace(" ", "_").replace(".",""), val.replace("'",""))
                    else:
                        com='el.%s = "%s"' % (x.replace(" ", "_").replace(".",""), val)
                    exec(com)
                i+=1
                s.prev.addriga(s.prev, el)
                righe.append(el)
            s.pratica.addprev(s.pratica, s.prev)
        #in righe ho tutte le righe del preventivo (s.elem) e s.prev la testata/piede

def ScriviLog(messaggio):
    filePath = "C:\\hts\\logfile.txt"
    if filePath == "":
        tk.messagebox.showerror("Errore nel percorso del file di log txt", "Errore file log")  
  
    dataOra = datetime.now()
    filelog=open(filePath,"r",encoding="UTF-8")
    # Scrivi la data, l'ora e il messaggio nel file di log
    filelog.write(dataOra +  " - " + messaggio)
    # Chiudi il file
    filelog.close()
      
def nuovaPratica(idpra, filename):
    #trovo ultimo numero pratica e imposto il numero di pratica per nuovo record
    ScriviLog( "import.py - inserimento nuova pratica - insert carvei")
    
    strSQL = "SELECT Count(CARVEI.F_NUMPRA) AS ConteggioDiF_NUMPRA FROM CARVEI;"
    carvei_cursor = conn.cursor()
    carvei_cursor.execute(strSQL)  #creo un cursore/recordset(cursor) da una query
    rows = carvei_cursor.fetchall() #recupera il risultato della query in execute e lo mette in una lista 
    ScriviLog ("Il tot. pratiche è: " + len(rows) + "ConteggioDiF_NUMPRA")
    if len(rows) == 0:  #se non ci sono pratiche, parto da 1
        idPratica = 1
        ScriviLog ("C: nuova pratica con numero " + idPratica)
    else:
        #cerco il numero massimo di pratica
        # definisco query
        strSQL = "SELECT Max(CARVEI.F_NUMPRA) AS MaxDiF_NUMPRA FROM CARVEI;"
        carvei_cursor = conn.cursor()
        carvei_cursor.execute(strSQL)  #creo un cursore/recordset(cursor) da una query
        rows = carvei_cursor.fetchall() #recupera il risultato della query in execute e lo mette in una lista 
        MyVal = int(rows["MaxDiF_NUMPRA"])
        ScriviLog ("Ultimo num. pra: " + MyVal)
        idPratica = MyVal + 1
        ScriviLog ("import.py - Creata nuova pratica con numero " + idPratica)
    
    idPratica2 = idPratica      #per restituire il numero di pratica nuovo
    # inizio INSERT record in CARVEI
    # Ottieni la data corrente
    dataAttuale = datetime.now()
    # Converti la data nel formato yyyymmdd
    dataFormattata = (dataAttuale.__format__("yyyymmdd"))
    StringaTraParentesi=""
    
    if tipoPratica=="C":
        campi_carvei = """(f_numpra, f_targav, f_dataca, f_desmod, f_telaio, f_kimvei, f_datimm, 
            F_CODCLI, F_RAGSOC, F_VIACLI, F_CITTAC, F_CAPCLI, F_PROCLI, F_PARIVA, F_TELEFO, 
            f_tipove, f_tpreve, f_idmess, f_datcre, F___GUID, F_DESCOL)"""
        f_numpra = idPratica
        f_targav = prev0.Targa_Veicolo          #arrDati(2, 8)                
        f_dataca = prev0.Data_preventivo        #DateValue(arrDati(2, 2))
        f_desmod = prev0.Descrizione_Veicolo    #Left(arrDati(2, 4), 70)
        f_telaio = prev0.Telaio                 #arrDati(2, 5)
        f_kimvei = prev0.Km                     #arrDati(2, 6)
        f_datimm = prev0.Data_Immatricolazione  #arrDati(2, 7)
         #f_tipove  per il tipo vernice controllo le iniziali della descrizione
        if prev0.Tipo_smalto[0,3] == "OPA":
            f_tipove = "O"
        if prev0.Tipo_smalto[0,3] == "TRI":
            f_tipove = "T"
        if prev0.Tipo_smalto[0,3] == "PER":
            f_tipove = "L"
        if prev0.Tipo_smalto[0,3] == "MIC":
            f_tipove = "I"
        if prev0.Tipo_smalto[0,3] == "MET":
            f_tipove = "M"
        if prev0.Tipo_smalto[0,3] == "DOP":
            f_tipove = "O"
        if prev0.Tipo_smalto[0,3] == "PAS":
            f_tipove = "P"
        #per i dati del cliente sono impostati su parametri.ini
        leggi_par_ini()
        f_tpreve = "C"   #tipo logo C per carrozzeria
        f_idmess = dataFormattata + idPratica   #id mess
        f_datcre = datetime.now()   #data e ora creazione pratica interna
        F___GUID = prev0.Id_riparazione[0,36]
        # codice colore vernice
        if prev0.Colore != "":
            codcol = "Bianco Lunare (764/A)" # DEBUG prev0.Colore
            if codcol.index("(")>=0:
                StringaTraParentesi=codcol[codcol.index("(")+1, codcol.index(")")-1]
                fine=min(codcol.index("(")-1, 20)
                F_DESCOL = codcol[0, fine] #tronco stringa a 20caratteri    DESCRIZIONE COLORE
            else:
                F_DESCOL = codcol
        valori_carvei = "('" + f_numpra + "','" + f_targav + "','" + f_dataca + "','" + f_desmod + "','" + f_telaio + "','" + f_kimvei + "','" + f_datimm + "', " 
        valori_carvei = valori_carvei + "'" + s.pratica.F_CODCLI + "','" + s.pratica.F_RAGSOC + "','" + s.pratica.F_VIACLI + "','" + s.pratica.F_CITTAC + "','"
        valori_carvei = valori_carvei + s.pratica.F_CAPCLI + "','" + s.pratica.F_PROCLI + "','" + s.pratica.F_PARIVA + "','" + s.pratica.F_TELEFO + "', " 
        valori_carvei = valori_carvei + "'" + f_tipove + "','" + f_tpreve + "','" + f_idmess + "','" + f_datcre + "','" + F___GUID + "', '" + F_DESCOL + "')"
        #fine case Carr 
         

    elif tipoPratica == "M":
        campi_carvei = """(f_numpra, f_targav, f_dataca, f_desmod, f_telaio, f_kimvei, f_datimm, 
                    F_CODCLI, F_RAGSOC, F_VIACLI, F_CITTAC, F_CAPCLI, F_PROCLI, F_PARIVA, F_TELEFO, _
                    f_nummot, f_tipove, f_tpreve, f_idmess, f_datcre, F___GUID)"""
        f_numpra = idPratica
        prev0 = s.pratica.listaprev[0]
        f_targav = prev0.Targa_Veicolo.replace("Targa Veicolo ","")
        f_dataca = datetime(prev0.Data_preventivo)
        f_desmod = prev0.Descrizione_Veicolo.replace("'", "''")[0, 70]
        f_telaio = prev0.Telaio
        f_kimvei = prev0.Km
        f_datimm = datetime(prev0.Data_Immatricolazione)
        #modifica del 24/03/2025 per i clienti privati oltre ALD
        #per i dati del cliente sono impostati su parametri.ini
        if prev0.Id_riparazione != "":    #se la colonna IdRip. contiene testo, è ALD
            leggi_par_ini()
        else:   #non faccio nulla perchè non devo valorizzare i dati del Ciente
            pass

        f_nummot = prev0.Km  #arrDati(2, 6)
        f_tipove = "O"   #tipo vernice di default metto doppio strato
        f_tpreve = "M"   #tipo logo M per meccanica
        f_idmess = dataFormattata + idPratica   #id mess
        f_datcre = datetime.now()   #data e ora creazione pratica interna
        F___GUID = idPratica
    
        valori_carvei = "('" + f_numpra + "','" + f_targav + "','" + f_dataca + "','" + f_desmod + "','" + f_telaio + "','" + f_kimvei + "','" + f_datimm + "', " 
        valori_carvei = valori_carvei + "'" + s.pratica.F_CODCLI + "','" + s.pratica.F_RAGSOC + "','" + s.pratica.F_VIACLI + "','" + s.pratica.F_CITTAC + "','"
        valori_carvei = valori_carvei + s.pratica.F_CAPCLI + "','" + s.pratica.F_PROCLI + "','" + s.pratica.F_PARIVA + "','" + s.pratica.F_TELEFO + "', " 
        valori_carvei = valori_carvei + "'" + f_nummot + "','" + f_tipove + "','" + f_tpreve + "','" + f_idmess + "','" + f_datcre + "','" + F___GUID + "')"
    #fine case M 

    strSQL = "insert into carvei " + campi_carvei + " values " + valori_carvei
    #Debug.Print strSQL   
    # Crea oggetti CURSOR
    carvei_ins_cursor = conn.cursor()
    # Esegui la query
    carvei_ins_cursor.execute(strSQL)
    ScriviLog ("import.py - Fine Line0 - insert carvei")
     
    #fine insert CARVEI

    #insert Pratica2
    if tipoPratica== "C":
        strSQL = "INSERT INTO Pratica2 (f_numpra, f_tippra, F_CODCOL) values (" + idPratica + ", '1', '" + StringaTraParentesi + "')"
    elif tipoPratica == "M":
        strSQL = "INSERT INTO Pratica2 (f_numpra, f_tippra) values (" + idPratica + ", '2')"

    ScriviLog("import.py - insert pratica2")
    # Esegui la query
    carvei_ins_cursor.execute(strSQL)
    #fine insert pratica2

def vediprev():    
    my_cursor = conn.cursor()
    my_cursor.execute("SELECT * FROM TESPRE")  #creo un cursore/recordset(cursor) da una query
    rows = my_cursor.fetchall() #recupera il risultato della query in execute e lo mette in una lista 
    print("Num records: ", len(rows))
    i=0
    #Lb1.delete(0,Lb1.size()-1)
    tv['columns'] = ('IDPrev', 'DataPr', 'RagSoc', 'Totale')
    tv.heading("#0", text='Preventivo', anchor='w')
    tv.column("#0", anchor="w", width=20)
    tv.heading('IDPrev', text='ID Prev.')
    tv.column('IDPrev', anchor='center', width=30)
    tv.heading('DataPr', text='Data')
    tv.column('DataPr', anchor='center', width=100)
    tv.heading('RagSoc', text='Rag. Sociale')
    tv.column('RagSoc', anchor='center', width=200)
    tv.heading('Totale', text='Totale')
    tv.column('Totale', anchor='e', width=100)
    for row in rows:
        #print(row.ID_CODPRE, row.F_DATAPR)
        i=i+1
        tv.insert('', 'end', values=(row.ID_CODPRE, row.F_DATAPR, row.F_RAGSOC, f"€ {int(row.F_TOTRIC):.2f}"))

def cerca():
    #for i in Lb1.curselection():
    #    print(Lb1.get(i))
    for i in tv.selection():
        print(i)
        print(tv.item(i))
        print(tv.item(i).values())
        l=tv.item(i).values()
        print(l)

def print_selection():
    # Create a vertical scrollbar
    v_scrollbar = tk.ttk.Scrollbar(window, orient=tk.VERTICAL, command=tv.yview)
    if (var1.get() == 0):
        label.config(text='Solo Nuove pratiche ')
        tv.grid_forget()
        btnScelta.grid_forget()
        v_scrollbar.grid_forget()
    else:
        label.config(text='Apri vecchie pratiche ')
        tv.grid(row=2, column=0, columnspan=2, sticky="W", padx=10, pady=20)
        #tv.configure(yscrollcommand=v_scrollbar.set)
        #v_scrollbar.grid(row=2, column=2, sticky="E")
        btnScelta.grid(row=3, column=0,sticky="WE", padx=20, pady=30)

#creo la maschera principale 1000x800
window = tk.Tk()
window.geometry("1000x800")
window.title("THS Interfaccia per WinCar")
window.resizable(False, False)

Font_tuple = ("Calibri", 18, "bold")
Font_tab = ("Calibri", 14, "normal")

#Pulsante per avvio procedura e per selezionare pratica esistente
btnAvvia=tk.Button(text="Vai", command=leggijson, font=Font_tuple, fg="yellow", bg="blue")
btnScelta=tk.Button(text="Scegli", command=cerca, font=Font_tuple, fg="yellow", bg="blue") 
        #command definisce il metodo da chiamare alla pressione del tasto
#inserisco una tabella 
tv = Treeview(window)
tv.grid_rowconfigure(0, weight = 1)
tv.grid_columnconfigure(0, weight = 1)
label = tk.Label(window, bg='white', width=20, text='Nuove Pratiche')
var1 = tk.IntVar()
c1 = tk.Checkbutton(window, text='Nuove Pratiche',variable=var1, onvalue=1, offvalue=0, command=print_selection, font=Font_tuple)
#posiziono tutti gli elementi a griglia nella finestra
btnAvvia.grid(row=1, column=0, sticky="WE", padx=20, pady=30)
label.grid(row=0, column=0,sticky="WE", padx=10, pady=20)
c1.grid(row=0, column=1,sticky="WE", padx=10, pady=20)

if __name__ == "__main__":
    window.mainloop()