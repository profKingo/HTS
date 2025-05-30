import tkinter as tk
from tkinter.ttk import *
from tkinter import *
import json
import pyodbc
import os
from datetime import datetime
import struttura as s
import configparser
from tkinter import filedialog
    
'''
l'utente parte da file pdf da ocrizzare e quindi da inviare al portale di AIDA (tramite uno dei metodi tra cui WS)
i file pdf sono di tipo MECC o CARR e vengono nominati ad hoc...
pensare un batch che chiami appena generato il file la conversione e poi avvii la lettura del JSON risultante
il JSON viene rimandato all'utente come???? Webhook - FTP - AIDA LINK
'''
conn=pyodbc.Connection
'''
nomedb="C:\\THS\\THS32Env\\wcArchivi.mdb"
filejson="export.json"
global targetFolder
targetFolder=""
mylog=""
lockFilePath=""
tipoPratica="M" #Meccanica #Carrozzieria
'''
def ScriviLog(messaggio):
    filePath = "C:\\hts\\logfile.txt"
    if filePath == "":
        tk.messagebox.showerror("Errore nel percorso del file di log txt", "Errore file log")  
    
    dataOra = datetime.now()
    filelog=open(filePath, "w", encoding="UTF-8")
    # Scrivi la data, l'ora e il messaggio nel file di log
    filelog.write(str(dataOra) +  s.tipoPratica +" - " + messaggio)
    # Chiudi il file
    filelog.close()


#stringa di connessione al db
def connetti():
    connstr=f'Driver={{Microsoft Access Driver (*.mdb)}};Dbq=' + s.nomedb + ';Uid=;Pwd=;'
    try:
        conn=pyodbc.connect(connstr)
    except:
        tk.messagebox.showerror("Errore di connessione al database!", "Errore accesso")

def leggi_par_ini():
    NOMEF="parametri.ini"
    if not os.path.exists(NOMEF):
        tk.messagebox.showerror("Errore: File dei parametri non esiste")
    else:
        config = configparser.ConfigParser()
        config.read(NOMEF)
        s.pratica.F_CODCLI = config["CLIENTE"]["f_codli"]
        s.pratica.F_RAGSOC = config["CLIENTE"]["f_ragsoc"]
        s.pratica.F_VIACLI = config["CLIENTE"]["f_viacli"]
        s.pratica.F_CITTAC = config["CLIENTE"]["f_cittac"]
        s.pratica.F_CAPCLI = config["CLIENTE"]["f_capcli"]
        s.pratica.F_PROCLI = config["CLIENTE"]["f_procli"]
        s.pratica.F_PARIVA = config["CLIENTE"]["f_pariva"]
        s.pratica.F_TELEFO = config["CLIENTE"]["f_telefo"]

        s.nomedb=config["PATH"]["MyPath"] + "\\" + config["PATH"]["DBName"]
        s.targetFolder=config["PATH"]["targetFolder"]
        s.mylog=config["PATH"]["mylog"]
        s.tipoPratica=config["PATH"]["tipopratica"]
        s.lockFilePath=config["PATH"]["lockFilePath"]

def elaboraPratica(numPratica, Y, lFileName, tipoPratica): 
    ScriviLog ("inizio import")
    connetti()
    #MODIFICA DEL 29/01/2025
        #'controllo se è stato indicato un NUMERO PRATICA nel nome file
    if numPratica > 0:
        strSQL = "SELECT CARVEI.F_NUMPRA FROM CARVEI WHERE (((CARVEI.F_NUMPRA)=" + numPratica + "));"
        # Apri il recordset
        prat_cursor = conn.cursor()
        prat_cursor.execute(strSQL)  #creo un cursore/recordset(cursor) da una query
        rows = prat_cursor.fetchall() #recupera il risultato della query in execute e lo mette in una lista 
        MyVal = int(rows["MaxDiF_NUMPRA"])
        
        if len(rows) > 0:
            ScriviLog("pratica trovata")
            #controllo quanti NUM PREVENTIVI ci sono per questa pratica
            idPratica = numPratica
            cercaNumPre(idPratica, idPreventivo)
            #se 1 preventivo vuoto, allora CREO NUOVO PREVENTIVO
            #se più di 1, faccio selezionare il preventivo da SOVRASCRIVERE oppure CREO NUOVO PREVENTIVO
        else:
            ScriviLog ("non ci sono pratiche")
            corrispondenzaTarga(idPratica, idPreventivo, lFileName, tipoPratica)
            #controllo la corrispondenza con la TARGA
            #se non c'è corrispondenza, creo nuova pratica e importo preventivo
            #se c'è corrispondenza, controllo quanti NUM PREVENTIVI  ci sono per questa pratica
            #se 1 preventivo vuoto, allora CREO NUOVO PREVENTIVO
            #se più di 1, faccio selezionare il preventivo da SOVRASCRIVERE oppure CREO NUOVO PREVENTIVO

    else:
        ScriviLog("pratica KO")
        if var3!=0:
            corrispondenzaTarga(idPratica, idPreventivo, lFileName, tipoPratica)
            #controllo la corrispondenza con la TARGA
            #se non c'è corrispondenza, creo nuova pratica e importo preventivo
            #se c'è corrispondenza, controllo quanti NUM PREVENTIVI  ci sono per questa pratica
            #se 1 preventivo vuoto, allora CREO NUOVO PREVENTIVO
            #se più di 1, faccio selezionare il preventivo da SOVRASCRIVERE oppure CREO NUOVO PREVENTIVO
        else:
            ScriviLog("Inizio inserimento pratica")
            idPratica=nuovaPratica(idPratica, lFileName) 
            #NB per usare una sub e restituire un valore, usare ByRef, in questo caso mi serve il nuovo numero pratica
            #NON CAPISCO L'USO DI idPratica2 => sostituito con return in funzione nuovaPratica
            idPreventivo = 1
            #CICCOLONE    
            # idPratica = idPratica2
            inserisciNuovoPreventivo_TesPre(arrDati, idPratica, idPreventivo, tipoPratica)
            inserisciNuovoPreventivo_RigPre(arrDati, idPratica, idPreventivo, tipoPratica)
            ScriviLog("Inserita nuova pratica N. " + idPratica)
            feedback2 = "inserita nuova"
    termina(Y, lFileName, feedback2, idPratica)

def nuovaPratica(idpra, filename):
    #trovo ultimo numero pratica e imposto il numero di pratica per nuovo record
    ScriviLog("import.py - inserimento nuova pratica - insert carvei")
    
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
    
    #idPratica2 = idPratica      #per restituire il numero di pratica nuovo
    #CICCOLONE Commentata linea sopra non serve per ritornare n° pratica nuovo (faccio return sotto)
    
    # inizio INSERT record in CARVEI
    # Ottieni la data corrente
    dataAttuale = datetime.now()
    # Converti la data nel formato yyyymmdd
    dataFormattata = (dataAttuale.__format__("yyyymmdd"))
    StringaTraParentesi=""
    
    if s.tipoPratica=="C":
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
         
    elif s.tipoPratica == "M":
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
    return idPratica 
    #fine insert CARVEI

    #insert Pratica2
    if s.tipopratica == "C":
        strSQL = "INSERT INTO Pratica2 (f_numpra, f_tippra, F_CODCOL) values (" + idPratica + ", '1', '" + StringaTraParentesi + "')"
    elif s.tipoPratica == "M":
        strSQL = "INSERT INTO Pratica2 (f_numpra, f_tippra) values (" + idPratica + ", '2')"

    ScriviLog("import.py - insert pratica2")
    # Esegui la query
    carvei_ins_cursor.execute(strSQL)
    #fine insert pratica2

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
label2 = tk.Label(window, bg='white', width=20, text='')
label3 = tk.Label(window, bg='white', width=20, text='')
var1 = tk.IntVar()
var2 = tk.IntVar()
var3 = tk.IntVar()
c1 = tk.Checkbutton(window, text='Nuove Pratiche',variable=var1, onvalue=1, offvalue=0, command=print_selection, font=Font_tuple)
c2 = tk.Checkbutton(window, text='Elimina files ',variable=var2, onvalue=1, offvalue=0, command=print_selection, font=Font_tuple)
c3 = tk.Checkbutton(window, text='Controllo corrispondeza Targa',variable=var3, onvalue=1, offvalue=0, command=print_selection, font=Font_tuple)
#posiziono tutti gli elementi a griglia nella finestra
btnAvvia.grid(row=1, column=0, sticky="WE", padx=20, pady=30)
label.grid(row=0, column=0,sticky="WE", padx=10, pady=20)
c1.grid(row=0, column=1,sticky="WE", padx=10, pady=20)
c2.grid(row=1, column=1,sticky="WE", padx=10, pady=20)
c3.grid(row=2, column=1,sticky="WE", padx=10, pady=20)

if __name__ == "__main__":
    window.mainloop()
