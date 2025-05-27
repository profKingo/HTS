import tkinter as tk
from tkinter.ttk import *
from tkinter import *
import json
import pyodbc
import os
import datetime
import struttura as s
import configparser
from tkinter import filedialog
from tkinter import messagebox as mb

'''
l'utente parte da file pdf da ocrizzare e quindi da inviare al portale di AIDA (tramite uno dei metodi tra cui WS)
i file pdf sono di tipo MECC o CARR e vengono nominati ad hoc...
pensare un batch che chiami appena generato il file la conversione e poi avvii la lettura del JSON risultante
il JSON viene rimandato all'utente come???? Webhook - FTP - AIDA LINK
'''
conn=pyodbc.Connection
files_file=[]
arrDati=[]  #######  DA ELIMINARE  #########################


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
    
    dataOra = datetime.datetime.now()
    filelog=open(filePath, "w", encoding="UTF-8")
    # Scrivi la data, l'ora e il messaggio nel file di log
    filelog.write(str(dataOra) +  s.tipoPratica +" - " + messaggio)
    # Chiudi il file
    filelog.close()


#stringa di connessione al db
def connetti():
    connstr=f'Driver={{Microsoft Access Driver (*.mdb)}};Dbq=' + s.nomedb + ';Uid=;Pwd=;'
    try:
        global conn 
        conn=pyodbc.connect(connstr)
    except:
        tk.messagebox.showerror("Errore di connessione al database!", "Errore accesso")
    if conn:
        print("ok")

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

def leggijson(pfile):
    leggi_par_ini()
    connetti()
    #print(s.targetFolder)
    if pfile is None:
        filetypes = (
            ('JSON files', '*.json'),
            ('All files', '*.*')
        )
        file1 = filedialog.askopenfilename(title='Apri un file',
            initialdir="c:/hts/",
            filetypes=filetypes)
    else:   
        file1=pfile
    data = json.load(open(file1))
    ScriviLog( "import.py - lettura file json")
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

def StartClock():
    pass
def StopClock():
    pass

def Import_Dati():
    ####GC########## disabilitaTutto()
    
    if datetime.datetime.now() > datetime.datetime(datetime.datetime.today().year, 12, 31): #
        tk.messagebox.showerror("Error", "Contattare HTS GROUP per l'assistenza.")
        exit
    
    '''' controllo se è finito il processo di conversione 
    Dim pathOCR As String
    pathOCR = Modulo2.ini_manager("r", "PATH", "OCRPath")
    If FileExist(pathOCR & "semaforo") Then
        MsgBox "I file da importare non sono ancora pronti." & vbCrLf & "Attendere... e riprovare.", vbCritical, "Conversione ancora in corso..."
        Exit Sub
    End If
    NON MI SERVE FARE l'OCR
    '''
    count_file_importati = 0
    Y = 12
    numPratica = 0
      
    #prendo il percorso di lavoro file da parametri.ini
    #separatore numero pratica dei file carrozzeria
    #targetFolder = Modulo2.ini_manager("r", "PATH", "targetFolder")
    #separatore = Modulo2.ini_manager("r", "PARAMETRI", "separatore")
    leggi_par_ini()
    
    StartClock()  #avvia contatore runtime
    
    ############# INIZIO ELABORAZIONE FILE PDF 
    # check box 3 controllo sviluppo; check box 4 controllo software conversione OCR
    ####CICCOLONE######## Non serve perchè chiamiamo un'altra applicazione per OCR
    '''
    If ActiveSheet.CheckBoxes("Check Box 3").value <> 1 Then    'se non è 1 vuol dire che non è attivo la modalità sviluppo
        Debug.Print "mod.sviluppo OFF"
        If ActiveSheet.CheckBoxes("Check Box 4").value <> 1 Then    ' se non è 1 vuol dire che non è stata disattivata la conversione OCR
            Debug.Print "OCR ON"
            # controllo se ci sono file pdf da elaborare
            If FileExist(targetFolder & "*.pdf") Then       ' se ci sono file pdf da elaborare
                AvvioFileBatch
            Else
                ScriviLog "Non ci sono file pdf da elaborare nella cartella " & targetFolder
                MsgBox "Non ci sono file pdf da elaborare nella cartella " & targetFolder
                Exit Sub
            End If
        Else
            Debug.Print "OCR OFF"
        End If
        '
    Else
        Debug.Print "mod.sviluppo ON"
    End If
    # FINE ELABORAZIONE FILE PDF '''

    files = os.listdir(s.targetFolder)
    filesno = 0
    global files_file
    files_file = [f for f in files if os.path.isfile(os.path.join(s.targetFolder, f))]
    for file in files_file:
        if file.endswith(".json"):
            filesno = filesno + 1
    if filesno>0:
        ScriviLog("Presenti file *.json da elaborare")     
    else:
        tk.messagebox.showerror("Non ci sono file da elaborare nella cartella " + s.targetFolder, "Mancano files")
        ScriviLog("Non ci sono file da elaborare nella cartella " + s.targetFolder)

    res=mb.askquestion('Importazione su DB', 'Vuoi importare sul DB?')
    if res == 'no':
        tk.messagebox.showinfo("Nessun dato importato","Processo completato senza importazione")
    else:
        importaSuDB()

def importaSuDB():
    #continua:
    ####GC#### ClearCellsFromB12ToLastRow
    '''ciclo per tutti i file excel presenti nella cartella '''
    filesno = 0
    for file in files_file:
        #controllo se è file excel elaboro solo quelli
        if file.endswith(".json"):
            filesno = filesno + 1
            ScriviLog("--------------------------------------------------")
            ScriviLog("inizio elaborazione : " + file)
            leggijson(file) #Sostituisco tutta la parte sotto con questa chiamata 
            prev1 = s.pratica.listaprev[0] 
            if prev1.F_TIPPRE == "Mecc":
                tipoPratica = "M"
            else:
                tipoPratica = "C"

            #### Se gli passo il numero di pratica lo avrei inserito in s.pratica.... AL MOMENTO NON IN USO!!!
            try:
                if s.pratica.NumPratica!=None:
                    numPratica=0
                else:
                    numPratica=s.pratica.NumPratica
            except:
                numPratica=0
            elaboraPratica(numPratica, Y, file, tipoPratica)
        if var2==1:
            #cancello file json elaborato
            os.remove(file)             
        
            #Apro File Excels e conto tot. righe e tot. colonne
            '''COMMENTATO CICCOLONE  Non serve più
            Workbooks.Open targetFolder & lFileName
            lastRow = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
            LastCol = Cells(1, Columns.Count).End(xlToLeft).Column
            'Debug.Print "ultima.riga: " & LastRow & " tot colonne: " & LastCol
            ReDim arrDati(lastRow, LastCol)
            'assegnamo a arr(x, y) tutti i valori del nostro foglio, dove x è l'indice delle righe e y quello delle colonne
            arrDati = Range(Cells(1, 1), Cells(lastRow, LastCol))
            Workbooks(lFileName).Close SaveChanges:=False
            '
            If ActiveSheet.CheckBoxes("Check Box 5").value <> 1 Then
            'controllo se il nome file excel ha numero pratica prima di separatore
                If InStr(File.Name, separatore) > 0 Then    'controllo se il file ha un separatore
                    numeroStringa = Left(File.Name, (InStr(File.Name, separatore) - 1))
                    'Debug.Print "Numero pratica ipotetico: " & numeroStringa
                    If IsNumeric(numeroStringa) Then
                        numPratica = CLng(numeroStringa)
                        'Debug.Print "Il valore è un numero: " & numPratica
                    Else
                        numPratica = 0
                        'Debug.Print "Il valore è una stringa: " & numPratica
                    End If
                    ScriviLog "Numero pratica ipotetico: " & numPratica
                Else
                    numPratica = 0
                End If
            Else
                numPratica = 0
            End If
            '
            'controllo se il file excel è di meccanica o carozzeria
            'modifica del 20/02/2025
            If InStr(1, File.Name, "mecc", vbTextCompare) > 0 Then
                'chiamo la routine di meccanica passando le variabili che mi servono
                ScriviLog "Inizio routine meccanica ALD"
                tipoPratica = "M"
                'Call Modulo4.elaboraMeccanica(arrDati, lFileName, LastRow, y, count_file_importati, numPratica)               'MECCANICA
                'Call Modulo4.elaboraMeccanica(numPratica, Y, lFileName)
            Else
                If InStr(1, File.Name, "carr", vbTextCompare) > 0 Then
                    'chiamo la routine di carrozzeria passando le variabili che mi servono
                    ScriviLog "Inizio routine carrozzeria ALD"
                    tipoPratica = "C"
                    'Call Modulo3.elaboraCarrozzeria(arrDati, lFileName, LastRow, y, count_file_importati, numPratica)            'CARROZZERIA
                    'Call Modulo3.elaboraCarrozzeria(numPratica, lFileName)              'CARROZZERI
                End If
            End If
            
        End If
        '
        End If
    Next'''
    #fine ciclo file json 
    ScriviLog("File correttamente importati: " + str(filesno))
    
    tk.messagebox.showinfo("Processo completato.", "Processo completato")
    StopClock()

###################################################################################
def cercaNumPre(idPratica, idPreventivo):
    pass
def inserisciNuovoPreventivo_TesPre(idPratica, idPreventivo, tipoPratica):
    pass
def inserisciNuovoPreventivo_RigPre(idPratica, idPreventivo, tipoPratica):
    pass
def termina(Y, lFileName, feedback2, idPratica):
    pass
###################################################################################

def elaboraPratica(idPratica, Y, lFileName, tipoPratica): 
    ScriviLog ("inizio import")
    connetti()
    #MODIFICA DEL 29/01/2025
    #'controllo se è stato indicato un NUMERO PRATICA nel nome file
    if idPratica > 0:
        strSQL = "SELECT CARVEI.F_NUMPRA FROM CARVEI WHERE (((CARVEI.F_NUMPRA)=" + str(idPratica) + "));"
        # Apri il recordset
        prat_cursor = conn.cursor()
        prat_cursor.execute(strSQL)  #creo un cursore/recordset(cursor) da una query
        rows = prat_cursor.fetchall() #recupera il risultato della query in execute e lo mette in una lista 
        MyVal = int(rows["MaxDiF_NUMPRA"])
        
        if len(rows) > 0:
            ScriviLog("pratica trovata")
            #controllo quanti NUM PREVENTIVI ci sono per questa pratica
            idPratica = idPratica
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
        if var1==0:
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
            inserisciNuovoPreventivo_TesPre(idPratica, idPreventivo, tipoPratica)
            inserisciNuovoPreventivo_RigPre(idPratica, idPreventivo, tipoPratica)
            ScriviLog("Inserita nuova pratica N. " + str(idPratica))
            feedback2 = "inserita nuova"
    termina(Y, lFileName, feedback2, idPratica)

def vediprev(targa):    
    my_cursor = conn.cursor()
    strSQL = "SELECT CARVEI.F_NUMPRA, CARVEI.F_DATACA, CARVEI.F_TARGAV, CARVEI.F_RAGSOC, CARVEI.F_IMPPRE, " \
        "CARVEI.F_TPREVE FROM CARVEI WHERE (((CARVEI.F_TARGAV) like '" + targa + "') AND ((CARVEI.F_CHIUS2)<>80)) "\
        "ORDER BY CARVEI.F_NUMPRA DESC ,CARVEI.F_DATACA DESC;"
        # Esegui la query
    my_cursor.execute(strSQL)  #creo un cursore/recordset(cursor) da una query
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
    return rows

def nuovaPratica(idpra, filename):
    # Ottieni la data corrente
    dataAttuale = datetime.datetime.now()
    # Converti la data nel formato yyyymmdd
    dataFormattata = (dataAttuale.__format__("yyyymmdd"))
    
    #trovo ultimo numero pratica e imposto il numero di pratica per nuovo record
    ScriviLog( "import.py - inserimento nuova pratica - insert carvei")
    
    strSQL = "SELECT Count(CARVEI.F_NUMPRA) AS ConteggioDiF_NUMPRA FROM CARVEI;"
    carvei_cursor = conn.cursor()
    carvei_cursor.execute(strSQL)  #creo un cursore/recordset(cursor) da una query
    rows = carvei_cursor.fetchall() #recupera il risultato della query in execute e lo mette in una lista 
    ScriviLog ("Il tot. pratiche è: " + str(len(rows)))
    if len(rows) == 0:  #se non ci sono pratiche, parto da 1
        idPratica = 1
        ScriviLog ("C: nuova pratica con numero " + str(idPratica))
    else:
        #cerco il numero massimo di pratica
        # definisco query
        strSQL = "SELECT Max(CARVEI.F_NUMPRA) AS MaxDiF_NUMPRA FROM CARVEI;"
        carvei_cursor = conn.cursor()
        carvei_cursor.execute(strSQL)  #creo un cursore/recordset(cursor) da una query
        maxnp = carvei_cursor.fetchone() #recupera il risultato della query in execute e lo mette in una lista 
        MyVal = maxnp.MaxDiF_NUMPRA
        ScriviLog ("Ultimo num. pratica: " + str(MyVal))
        idPratica = MyVal + 1
        ScriviLog ("import.py - Creata nuova pratica con numero " + str(idPratica))
    
    #idPratica2 = idPratica      #per restituire il numero di pratica nuovo
    #CICCOLONE Commentata linea sopra non serve per ritornare n° pratica nuovo (faccio return sotto)
    
    # inizio INSERT record in CARVEI
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
        f_dataca = datetime.fromtimestamp(prev0.Data_preventivo/1000)
        from datetime import datetime
        f_desmod = prev0.Descrizione_Veicolo.replace("'", "''")[0, 70]
        f_telaio = prev0.Telaio
        f_kimvei = prev0.Km
        f_datimm = datetime.fromtimestamp(prev0.Data_Immatricolazione/1000)
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
    if s.tipopratica == "C":
        strSQL = "INSERT INTO Pratica2 (f_numpra, f_tippra, F_CODCOL) values (" + idPratica + ", '1', '" + StringaTraParentesi + "')"
    elif s.tipoPratica == "M":
        strSQL = "INSERT INTO Pratica2 (f_numpra, f_tippra) values (" + idPratica + ", '2')"

    ScriviLog("import.py - insert pratica2")
    # Esegui la query
    carvei_ins_cursor.execute(strSQL)
    #fine insert pratica2
    return idPratica 
    
def corrispondenzaTarga(idPratica, idPreventivo, lFileName, tipoPratica):
    prev0 = s.pratica.listaprev[0]
    targa = prev0.Targa_Veicolo.replace("Targa Veicolo ","")
    #controllo se la targa non è nulla
    if targa != None:
        ScriviLog("LA TARGA E': " + targa)
        # controllo se esistono altre pratiche fatte con questa targa
        '''
        strSQL = "SELECT CARVEI.F_NUMPRA, CARVEI.F_DATACA, CARVEI.F_TARGAV, CARVEI.F_RAGSOC, CARVEI.F_IMPPRE, " \
        "CARVEI.F_TPREVE FROM CARVEI WHERE (((CARVEI.F_TARGAV) like '" + targa + "') AND ((CARVEI.F_CHIUS2)<>80)) "
        "ORDER BY CARVEI.F_NUMPRA DESC ,CARVEI.F_DATACA DESC;"
        # Esegui la query
        prat_cursor = conn.cursor()
        prat_cursor.execute(strSQL)  #creo un cursore/recordset(cursor) da una query
        # Apri il recordset
        rows = prat_cursor.fetchall() #recupera il risultato della query in execute e lo mette in una lista 
        '''
        #Chiamo la funzione vediprev per riempire la tabella e avere il recordset dei preventivi/pratiche
        rows=vediprev(targa)
        if len(rows)==0:
            #Controlla se il recordset ha dati
            #se l'array è vuoto - NUOVA PRATICA
            ScriviLog("L'array2 è vuoto!")
            ####### nuovaPratica(idPratica2, lFileName) # SOSTITUITO CON SOTTO
            idPratica2=nuovaPratica(idPratica, lFileName) #NB per usare una sub e restituire un valore, usare ByRef, in questo caso mi serve il nuovo numero pratica
            idPreventivo = 1
            idPratica = idPratica2
            inserisciNuovoPreventivo_TesPre(arrDati, idPratica, idPreventivo, tipoPratica)
            inserisciNuovoPreventivo_RigPre(arrDati, idPratica, idPreventivo, tipoPratica)
            ScriviLog("Inserita nuova pratica N. " + str(idPratica))
            feedback2 = "inserita nuova"
        else:
            #se non è vuoto Determina il numero di righe e colonne
            numColonne = len(rows[0])  # NON SERVE!!!!
            numRighe = len(rows)

            if numRighe == 1:
                controllaSingolaPratica(lFileName)
                if selectedValue3 > 0:
                    idPratica = arrDati2(0, 0)  # assegno il numero pratica
                    cercaNumPre(idPratica, idPreventivo)
                else:    #se uguale a zero
                    if selectedValue == 0:
                        idPratica2=nuovaPratica(idPratica2, lFileName) # NB per usare una sub e restituire un valore, usare ByRef, in questo caso mi serve il nuovo numero pratica
                        idPreventivo = 1
                        idPratica = idPratica2
                        inserisciNuovoPreventivo_TesPre(arrDati, idPratica, idPreventivo, tipoPratica)
                        inserisciNuovoPreventivo_RigPre(arrDati, idPratica, idPreventivo, tipoPratica)
                        ScriviLog("Inserita nuova pratica N. " + str(idPratica))
                        feedback1 = "inserita nuova"
                    else:
                        exit()
            else:
                ''' DA CAPIRE......
                arrDati2 = Application.WorksheetFunction.Transpose(arrDati2)
                # Popola la ListBox con l'array
                # Configura il UserForm
                ######################################################################################
                Set UserForm = New UserForm1
                UserForm.lstRecords.ColumnCount = 6         ' Imposta il numero di colonne
                '        UserForm.lstRecords.BoundColumn = 1
                UserForm.lstRecords.ColumnWidths = "80;100;100;200;100;80" ' Larghezza delle colonne
                '        UserForm.lstRecords.Width = 1000
                UserForm.lstRecords.List = arrDati2                ' Assegna l'array alla ListBox
                UserForm.lblHeaders.Caption = "N.Pratica                Data                               Targa                       Cliente                                       Imp.Preventivo           Tipo Pratica"
                ' Mostra il UserForm
                UserForm.txtTarga = targa
                UserForm.txtFilePDF = Left(lFileName, InStrRev(lFileName, ".") - 1)
                UserForm.Show
                ######################################################################################
                '''
                ScriviLog("Hai selezionato: " + selectedValue)
                if selectedValue == 0:       #se 0 creo nuova pratica con preventivo
                    idPreventivo = 1
                    idPratica2=nuovaPratica(idPratica2, lFileName)
                    idPratica = idPratica2
                    idPratica2=nuovaPratica(idPratica2, lFileName)
                    inserisciNuovoPreventivo_TesPre(arrDati, idPratica, idPreventivo, tipoPratica)
                    inserisciNuovoPreventivo_RigPre(arrDati, idPratica, idPreventivo, tipoPratica)
                    ScriviLog("Inserita nuova pratica N. " + str(idPratica))
                    feedback2 = "inserita nuova"
                else:                            #se scelgo una pratica, controllo i numeri preventivi
                    if selectedValue > 0:
                        idPratica = selectedValue
                        cercaNumPre(idPratica, idPreventivo)
                    else:
                        exit()
    else:
        nuovaPratica(idPratica2, lFileName)


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
        tv.grid(row=4, column=0, columnspan=2, sticky="W", padx=10, pady=10)
        #tv.configure(yscrollcommand=v_scrollbar.set)
        #v_scrollbar.grid(row=2, column=2, sticky="E")
        btnScelta.grid(row=5, column=0,sticky="WE", padx=10, pady=10)
    if (var2.get() == 0):
        label2.config(text='')
        #tv.grid_forget()
        #btnScelta.grid_forget()
    else:
        label2.config(text='Elimina files ')
        #tv.grid(row=2, column=0, columnspan=2, sticky="W", padx=10, pady=20)
        #tv.configure(yscrollcommand=v_scrollbar.set)
        #v_scrollbar.grid(row=2, column=2, sticky="E")
        #btnScelta.grid(row=3, column=0,sticky="WE", padx=20, pady=30)

#creo la maschera principale 1000x800
window = tk.Tk()
window.geometry("1000x800")
window.title("THS Interfaccia per WinCar")
window.resizable(False, False)

Font_tuple = ("Calibri", 18, "bold")
Font_tab = ("Calibri", 14, "normal")

#Pulsante per avvio procedura e per selezionare pratica esistente
btnAvvia=tk.Button(text="Leggi files", command=Import_Dati  , font=Font_tuple, fg="yellow", bg="blue")
btnScelta=tk.Button(text="Scegli", command=cerca, font=Font_tuple, fg="yellow", bg="blue") 
        #command definisce il metodo da chiamare alla pressione del tasto
#inserisco una tabella 
tv = Treeview(window)
tv.grid_rowconfigure(0, weight = 1)
tv.grid_columnconfigure(0, weight = 1)
label = tk.Label(window, bg='white', width=20, text='Apri vecchie Pratiche')
label2 = tk.Label(window, bg='white', width=20, text='')
label3 = tk.Label(window, bg='white', width=20, text='')
var1 = tk.IntVar()
var2 = tk.IntVar()
c1 = tk.Checkbutton(window, text='Apri vecchie Pratiche',variable=var1, onvalue=1, offvalue=0, command=print_selection, font=Font_tuple)
#c3 = tk.Checkbutton(window, text='Controllo corrispondeza Targa',variable=var3, onvalue=1, offvalue=0, command=print_selection, font=Font_tuple)
#sottinteso con il checkbox sopra
c2 = tk.Checkbutton(window, text='Elimina files ',variable=var2, onvalue=1, offvalue=0, command=print_selection, font=Font_tuple)
#posiziono tutti gli elementi a griglia nella finestra
img = PhotoImage(file="C:\\HTS\\THS32Env\\logoHTS.png")
labelimg = tk.Label(window, image=img)
labelimg.grid(row=0, column=1,sticky="W", padx=10, pady=20)
label.grid(row=1, column=0,sticky="WE", padx=10, pady=0)
c1.grid(row=1, column=1,sticky="W", padx=10, pady=0)
label2.grid(row=2, column=0,sticky="WE", padx=10, pady=0)
c2.grid(row=2, column=1,sticky="W", padx=10, pady=0)
btnAvvia.grid(row=3, column=0, sticky="WE", padx=20, pady=30)
files_file=[]

if __name__ == "__main__":
    window.mainloop()
