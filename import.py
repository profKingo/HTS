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
        file1=s.targetFolder + pfile
    data = json.load(open(file1))
    ScriviLog( "import.py - lettura file json")
    if len(data)>=1:
        s.pratica.desc="Pratica"
        for pre in data:
            for x in s.s_header:    #INTESTAZIONE PREVENTIVO
                #qui dovrei leggere l'intestazione se si tratta di ALD o meno
                com=""
                try:
                    if type(data) is dict:
                        val=data[x]
                    else:
                        val=pre[x]
                    if isinstance(val, str):
                        com="s.prev.%s = '%s'" % (x.replace(" ", "_").replace(".",""), val.replace("'",""))
                    else:
                        com="s.prev.%s = '%s'" % (x.replace(" ", "_").replace(".",""), val)
                    exec(com)
                except:
                    print(file1, ":", x, "--->", com)
                #leggo le righe nel file e creo dei campi in una struttura che memorizzi i dati del preventivo
            #RIGHE PREVENTIVO
            try:
                if type(data) is dict:
                    tab=data["Tabella Interventi Meccanica"]
                else:
                    tab=pre["Tabella Interventi Meccanica"]   
            except:
                tab=[]
            righe=[]
            i=0   
            for e in tab:
                el=s.riga()
                for x in s.s_elem:
                    if type(data) is dict:
                        listaint=data["Tabella Interventi Meccanica"]
                        interv=listaint[i]
                        val=interv[x] 
                    else:
                        val=pre["Tabella Interventi Meccanica"][i][x] 
                    if isinstance(val, str):
                        com='el.%s = "%s"' % (x.replace(" ", "_").replace(".",""), val.replace("'",""))
                    else:
                        com='el.%s = "%s"' % (x.replace(" ", "_").replace(".",""), val)
                    exec(com)
                i+=1
                s.prev.addriga(s.prev, el)
                righe.append(el)
            s.prev.Targa_Veicolo=(s.prev.Targa_Veicolo).replace("Targa Veicolo ","")
            s.prev.Telaio=(s.prev.Telaio).replace("Telaio ","")
            s.pratica.addprev(s.pratica, s.prev)
            if type(data) is dict:
                break
        #in righe ho tutte le righe del preventivo (s.elem) e s.prev la testata/piede

def StartClock():
    pass
def StopClock():
    pass

def Import_JSON():
    if datetime.datetime.now() > datetime.datetime(2025, 12, 31): #
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

    #fine ciclo file json 
    ScriviLog("File correttamente importati: " + str(filesno))
    tk.messagebox.showinfo("Processo completato.", "Processo completato")
    StopClock()    
    if var1.get()==1:
        #Se non faccio nuove pratiche per forza (checkbox Nuove Pratiche SI) mi devo chiedere 
        #per ogni targa (di ogni preventivo) se devo cercare le pratiche/preventivi vecchi da associare
        for i in range(len(s.pratica.listaprev)):
            prevtmp=s.pratica.listaprev[i]
            listaoldprev=vediprev(prevtmp.Targa_Veicolo)    
            if len(listaoldprev)>0:
                #gestire la selezione della pratica/preventivo da collegare al preventivo attuale
                pass
            else:
                res = tk.messagebox.askquestion("Nessun preventivo con questa targa. Inserisco nuova pratica","Ricerca completata")
                if res=="yes":
                    Import_Dati()
    print(s.pratica.NumPratica)

def Import_Dati():
    res=mb.askquestion('Importazione su DB', 'Vuoi importare sul DB?')
    if res == 'no':
        tk.messagebox.showinfo("Nessun dato importato","Processo completato senza importazione")
    else:
        importaSuDB()

def importaSuDB():
    #continua:
    ####GC#### ClearCellsFromB12ToLastRow
    files = os.listdir(s.targetFolder)
    filesno = 0
    global files_file
    files_file = [f for f in files if os.path.isfile(os.path.join(s.targetFolder, f))]
    for file in files_file:
        if file.endswith(".json"):
            filesno = filesno + 1

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

###################################################################################
def cercaNumPre(idPratica, idPreventivo):
    pass

def inserisciNuovoPreventivo_TesPre(idPratica, idPreventivo, tipoPratica):
    ScriviLog("Avvio - insert tespre - pratica n : " + str(idPratica))
    if tipoPratica=="C":
        # inizio parametrizzazione campi tabella TESPRE
        ID_CODPRE_tespre = idPratica        
        prevcurr=s.prev
        for pr in s.pratica.listaprev:
            if pr.numprev==idPreventivo:
                prevcurr=pr
        F_DATAPR = prevcurr.Data_Preventivo        #data excel mecc
        if prevcurr.KM is not None:
            F_KMVEI = 0
        else:
            F_KMVEI = prevcurr.KM
        F_SUPPLE = 15                   
        F_FINITU = 10                   
        if prevcurr.Tempo_agg_vern == "":
            F_COMPLE = 0
        else:
            F_COMPLE = prevcurr.Tempo_agg_vern      # TEMPO AGG VERNIC
        #F_MATAUT = -1                                               ' calcolo automatico mat.cons.
        # MATERIALI DI CONSUMO
        
        strmatcon = prevcurr.F_MATCON.split(" ")
        F_MATCON = strmatcon[2]
        if prevcurr.Mat_consumo_iva == "":
            F_TOTMAT = 0
        else:
            F_TOTMAT = prevcurr.Mat_consumo_iva     #IVA MATERIALI DI CONSUMO
        F_COSTOR = prevcurr.Manodopera_carr         #IMPORTO MANODOPERA ORARIA CARROZZERIA
        F_COSTO2 = prevcurr.Manodopera_mecc         #IMPORTO MANODOPERA ORARIA MECCANICA
        F_MANCAR = prevcurr.Manodopera_carr_iva     #importo iva su manodopera carr.
        strF_MANMEC = prevcurr.Manodopera_mecc_imp
        F_TOTSR = prevcurr.Sr
        F_TOTLA = prevcurr.La
        F_TOTVE = prevcurr.Ve
        F_TOTME = prevcurr.Me
        F_TOTRIC = prevcurr.Ric_imp
        F_IVARIC = 22           #% IVA su pezzi
        F_IVAMAN = 22           #% IVA su manodopera
        F_IVAMAT = 22           #% IVA su materiali
        F_IVAVAR = 22           #% IVA su varie
        F_CIVRIC = 22
        F_CIVMAN = 22
        F_CIVMAT = 22
        F_CIVVAR = 22
        F_IIVARIC = prevcurr.Ric_imp                      #Imposta su pezzi
        F_IIVACAR = prevcurr.Manodopera_carr_iva          #Imposta su manodopera carrozzeria
        F_IIVAMEC = prevcurr.Manodopera_mecc_iva          #Imposta su manodopera meccanica
        F_IIVAMAT = prevcurr.Mat_consumo_iva              #Imposta su materiali
        F_IIVAVAR = prevcurr.Smalt_rif_iva                #Imposta su varie
        F_TIVARIC = prevcurr.Ric_tot                      #Totale (IVA compresa) pezzi
        F_TIVACAR = prevcurr.Manodopera_carr_tot          #Totale (IVA compresa) manodopera carrozzeria
        F_TIVAMEC = prevcurr.Manodopera_mecc_tot          #Totale (IVA compresa) manodopera meccanica
        F_TIVAMAT = prevcurr.Mat_consumo_tot              #Totale (IVA compresa) materiali
        F_TIVAVAR = prevcurr.Smalt_rif_tot                #Totale (IVA compresa) varie
        F_TOTPRE = prevcurr.Oss_totale_imp                #Totale preventivo IVA escl.
        F_TOTIVA = prevcurr.Oss_totale_iva                #Totale imposta
        F_TOTALE = prevcurr.Oss_totale_tot                #Totale preventivo IVA incl.
        F_TFINIT = prevcurr.Suppl_finitura                #Tempo per la finitura
        F_TSUPPL = prevcurr.Suppl_doppiostrato            #Tempo per il supplemento
        F_TCOMPL = prevcurr.Tempo_agg_vern                #Tempo per il completamento
        F_VEOPER = prevcurr.Tot_tempi_ve                  #Tempo VE operativo
        F_TEMSUP = prevcurr.Tot_tempi_suppl               #Totale tempi supplementari
        F_VALUTA_tespre = "Euro"           
        F_NUMPRE_tespre = idPreventivo                    #Numero Preventivo
        F_SMAMAX = 0                #IMPORTO MAX APPLICABILE SMALTIM. RIFIUTI    
        #Verifica se il valore è vuoto o nullo
        if prevcurr.Ric_imp is None or prevcurr.Ric_imp == "":
            strF_RICNET = 0
        else:
            # Converti il valore in decimale
            strF_RICNET = prevcurr.Ric_imp
        #fine parametrizzazione campi TESPRE
                
        # utilizzo una stringa di appoggio per i campi per semplificare la scrittura della query
            campi_tespre = """(ID_CODPRE  ,   F_DATAPR  ,   F_KMVEI  ,   F_SUPPLE  ,   F_FINITU  ,   F_COMPLE  , 
                F_MATCON  ,   F_TOTMAT  ,   F_COSTOR  ,   F_COSTO2  ,   F_MANCAR  ,   F_MANMEC  ,  
                F_TOTSR  ,   F_TOTLA  ,     F_TOTVE  ,   F_TOTME  , F_TOTRIC  ,   F_IIVARIC  ,   F_IIVACAR  , 
                F_IIVAMEC  ,   F_IIVAMAT  ,   F_IIVAVAR  , F_TIVARIC  ,   F_TIVACAR  ,   F_TIVAMEC  ,   F_TIVAMAT  , 
                F_TIVAVAR  ,   F_TOTPRE  ,   F_TOTIVA  ,   F_TOTALE  , F_TFINIT  ,   F_TSUPPL  ,   F_TCOMPL  , 
                F_VEOPER  ,   F_TEMSUP  ,  F_VALUTA  ,   F_NUMPRE  ,   F_SMAMAX  , F_RICNET, 
                F_IVARIC, F_IVAMAN, F_IVAMAT, F_IVAVAR, F_CIVRIC, F_CIVMAN, F_CIVMAT, F_CIVVAR )"""
        
            # utilizzo una stringa di appoggio per i valori da copiare per semplificare la scrittura della query
            valori_tespre = "(" + ID_CODPRE_tespre + ", " + F_DATAPR + ", " + F_KMVEI + ",  " + F_SUPPLE + ", " + F_FINITU + ", " + F_COMPLE + ", " 
            valori_tespre = valori_tespre + F_MATCON + ", " + F_TOTMAT + ", " + F_COSTOR + ", " + F_COSTO2 + ", " + F_MANCAR + ", " + strF_MANMEC + ", "
            valori_tespre = valori_tespre + F_TOTSR + ", " + F_TOTLA + ", " + F_TOTVE + ", " + F_TOTME + ", " + F_TOTRIC + ", " + F_IIVARIC + ", " + F_IIVACAR + "," 
            valori_tespre = valori_tespre + F_IIVAMEC + ", " + F_IIVAMAT + ", " + F_IIVAVAR + ", " + F_TIVARIC + ", " + F_TIVACAR + ", " + F_TIVAMEC + ", " + F_TIVAMAT + ", " 
            valori_tespre = valori_tespre + F_TIVAVAR + ", " + F_TOTPRE + ", " + F_TOTIVA + ", " + F_TOTALE + ", " + F_TFINIT + ", " + F_TSUPPL + ", " + F_TCOMPL + ", " 
            valori_tespre = valori_tespre + F_VEOPER + ", " + F_TEMSUP + ", '" + F_VALUTA_tespre + "', " + F_NUMPRE_tespre + ", " + F_SMAMAX + ", " + strF_RICNET + ", " 
            valori_tespre = valori_tespre + F_IVARIC + ", " + F_IVAMAN + ", " + F_IVAMAT + ", " + F_IVAVAR + ", " + F_CIVRIC + ", " + F_CIVMAN + ", " + F_CIVMAT + ", " + F_CIVVAR + ")"
        #fine case C
        #MECCANICA
    elif tipoPratica=="M":
        '''
        Costo Tot. Varie
        Piva Riparatore
        Numero ore lavorate
        Tariffa Manodopera
        Costo Tot. MDO
        SmaltRif
        ES. Iva
        Cod.Motore'''
        #inizio parametrizzazione campi tabella TESPRE
        ID_CODPRE_tespre = idPratica         
        F_DATAPR = s.pratica.Data_Preventivo        #data excel mecc
        if prevcurr.KM is not None:
            F_KMVEI = 0
        else:
            F_KMVEI = prevcurr.KM
        F_SUPPLE = 15           
        F_FINITU = 10           
        F_COMPLE = 1.6   
        F_MATCON = prevcurr.Materiale(0, 2)              #IMPORTO MATERIALI CONSUMO
        F_TOTMAT = prevcurr.Materiale                    #tot. materiali di consumo mecc
        F_COSTO2 = prevcurr.Tariffa_Manodopera           #tariffa manodopera mecc
        strF_MANMEC = (ReplaceApostrofo(arrDati(2, 29))) #
        F_TOTME = prevcurr.Numero_ore_lavorate           #tot ore manodopera mecc
        F_TOTRIC = prevcurr.Costo_Tot__Ricambi           #importo netto ricambi
        F_IVARIC = 22                                    #% IVA su pezzi
        F_IVAMAN = 22                                    #% IVA su manodopera
        F_IVAMAT = 22                                    #% IVA su materiali
        F_IVAVAR = 22                                    #% IVA su varie
        F_TOTPRE = prevcurr.Totale_Imponibile            #Totale preventivo IVA escl.
        F_TOTIVA =  prevcurr.Totale_Iva                  #Totale imposta
        F_TOTALE =  prevcurr.F_TOTPRE                    #Totale preventivo IVA incl.
        F_SCORIC = (arrDati(2, 21))                      #% Sconto riservato sui ricambi
        F_SCOMAN = (arrDati(2, 28))                      #% Sconto riservato sulla manodopera
        F_SCOVAR = (arrDati(2, 24))                      #% Sconto riservato sulle varie
        F_CALCOL = -1
        F_PERRIF = 0                                                        #% Smaltimento rifiuti da calcolare su manodopera VE + materiali di consumo IMPOSTATA A ZERO PERCHè INSERITA MANUALMENTE arrDati(2, 30)
        F_IMPRIF = 0                                                        #Importo derivato da  manodopera VE + materiali di consumo per calcolo smalt.rif.
        F_VALUTA_tespre = "Euro"           
        F_NUMPRE_tespre = idPreventivo           #Numero Preventivo
        F_FTSABA = -1         
        F_FTDOME = -1         
        F_CIVRIC = 22         
        F_CIVMAN = 22         
        F_CIVMAT = 22         
        F_CIVVAR = 22         
        F_SR_RIC = -1         
        F_SR_DIM = -1         
        F_SR_MAT = -1         
        F_SR_TSR = -1         
        F_SR_TLA = -1         
        F_SR_TVE = -1         
        F_SR_TME = -1         
        F_IMPRIC = 0          
        F_TEMPAR = 99         
        F_ESERIC = (arrDati(2, 32))                              #ci vanno gli importi esenti iva se presenti
        strF_RICNET = (ReplaceApostrofo(arrDati(2, 36)))         #tot imponibile
        #fine parametrizzazione campi TESPRE
            
        #utilizzo una stringa di appoggio per i campi per semplificare la scrittura della query
        campi_tespre = "(ID_CODPRE, F_DATAPR, F_KMVEI, F_SUPPLE, F_FINITU, F_COMPLE, F_MATCON, F_TOTMAT, " 
        campi_tespre = campi_tespre + "F_COSTO2, F_MANMEC, F_TOTME, F_TOTRIC, F_IVARIC, F_IVAMAN, F_IVAMAT, F_IVAVAR, " 
        campi_tespre = campi_tespre + "F_TOTPRE, F_TOTIVA, F_TOTALE, F_SCORIC, F_SCOMAN, F_SCOVAR, F_CALCOL, F_PERRIF, F_IMPRIF, " 
        campi_tespre = campi_tespre + "F_VALUTA, F_NUMPRE, F_FTSABA, F_FTDOME, F_CIVRIC, F_CIVMAN, F_CIVMAT, F_CIVVAR, " 
        campi_tespre = campi_tespre + "F_SR_RIC, F_SR_DIM, F_SR_MAT, F_SR_TSR, F_SR_TLA, F_SR_TVE, F_SR_TME, F_IMPRIC, " 
        campi_tespre = campi_tespre + "F_TEMPAR, F_ESERIC, F_RICNET )"
    
        #utilizzo una stringa di appoggio per i valori da copiare per semplificare la scrittura della query
        valori_tespre = "(" + ID_CODPRE_tespre + ", " + F_DATAPR + ", " + F_KMVEI + ",  " + F_SUPPLE + ", " + F_FINITU + ", " + F_COMPLE + ", " + F_MATCON + ", " + F_TOTMAT + ", " 
        valori_tespre = valori_tespre + "" + F_COSTO2 + ", " + strF_MANMEC + ", " + F_TOTME + ", " + F_TOTRIC + ", " + F_IVARIC + ", " + F_IVAMAN + ", " + F_IVAMAT + ", " + F_IVAVAR + ", " 
        valori_tespre = valori_tespre + "" + F_TOTPRE + ", " + F_TOTIVA + ", " + F_TOTALE + ", " + F_SCORIC + ", " + F_SCOMAN + ", " + F_SCOVAR + ", " + F_CALCOL + ", " + F_PERRIF + ", " + F_IMPRIF + ", " 
        valori_tespre = valori_tespre + "'" + F_VALUTA_tespre + "', " + F_NUMPRE_tespre + ", " + F_FTSABA + ", " + F_FTDOME + ", " + F_CIVRIC + ", " + F_CIVMAN + ", " + F_CIVMAT + ", " + F_CIVVAR + ",  " 
        valori_tespre = valori_tespre + "" + F_SR_RIC + ", " + F_SR_DIM + ", " + F_SR_MAT + ", " + F_SR_TSR + ", " + F_SR_TLA + ", " + F_SR_TVE + ", " + F_SR_TME + ", " + F_IMPRIC + ", " 
        valori_tespre = valori_tespre + "" + F_TEMPAR + ", " + F_ESERIC + ", " + strF_RICNET + ")"
        #fine case M
    
    strSQL = "insert into TESPRE " + campi_tespre + " values " + valori_tespre + ";"
    # Esegui la query
    tespre_cursor = conn.cursor()
    tespre_cursor.execute(strSQL)  #creo un cursore/recordset(cursor) da una query
    ScriviLog("Fine - insert tespre - pratica n : " + idPratica)
    # fine insert TESPRE

def ReplaceApostrofo(str):
    return str.replace("'","")

def inserisciNuovoPreventivo_RigPre(idPratica, idPreventivo, tipoPratica):
    ScriviLog("Avvio - insert rigpre - pratica n : " + idPratica)
    for i in range(len(s.pratica.listaprev)):
        if tipoPratica=='C':
            #CARROZZERIA
            ID_CODPRE = idPratica
            prevcurr=s.prev
            for pr in s.pratica.listaprev:
                if pr.numprev==idPreventivo:
                    prevcurr=pr
            if len(prevcurr.listrighe)==0:
                ScriviLog("Preventivo vuoto.")
                ScriviLog("Fine Line3 - insert rigpre - error - pdf vuoto preventivo non compilato, n : " + idPratica)
                #GoTo prevVuoto #se non c'è intestagione Righe vuol dire che il pdf è vuoto, allora non compilo rig-pre
                return
            ScriviLog("Imposta valori.")
            F_DATRIG = prevcurr.Data_Preventivo 
            F_ORDINE = i - 1
            F_CITFON = arrDati(i, 12)
            #descrizione articolo.
            originalString = arrDati(i, 13)
            modifiedString = prevcurr.Descrizione.replace("'", "''") # Removes all apostrophes
            F_DESART = modifiedString(0, 50)                 # descrizione articolo ha 50 caratteri max
            # quantità
            if prevcurr.listrighe[i].Qta == 0 or prevcurr.listrighe[i].Qta is None or prevcurr.listrighe[i].Qta == "":
                F_QUANTI = 1   # QUANITA' zero imposto 1
            else:
                F_QUANTI = prevcurr.listrighe[i].Qta
            F_DANNSR = (arrDati(i, 14))
            # h SR
            if arrDati(i, 15) == None or arrDati(i, 15) == "0":
                F_TEMPSR = 0
            else:
                if F_QUANTI > 1:
                    F_TEMPSR = CDbl(arrDati(i, 15)) / F_QUANTI
                else:
                    F_TEMPSR = CDbl(arrDati(i, 15))
            F_TEMPSR = ConvertToSQLDecimal(F_TEMPSR)
            F_DANNLA = (arrDati(i, 16))
            # h LA
            if arrDati(i, 17) is None or arrDati(i, 17) == "0":
                F_TEMPLA = 0
            else:
                if F_QUANTI > 1:
                    F_TEMPLA = CDbl(arrDati(i, 17)) / F_QUANTI
                else:
                    F_TEMPLA = CDbl(arrDati(i, 17))
            F_TEMPLA = ConvertToSQLDecimal(F_TEMPLA)
            F_DANNVE = arrDati(i, 18)
            # h VE
            if arrDati(i, 19) == vbNullString or arrDati(i, 19) == "0":
                F_TEMPVE = 0
            else:
                if F_QUANTI > 1:
                    F_TEMPVE = CDbl(arrDati(i, 19)) / F_QUANTI
                else:
                    F_TEMPVE = CDbl(arrDati(i, 19))
            F_TEMPVE = ConvertToSQLDecimal(F_TEMPVE)
            # h ME
            if arrDati(i, 20) == vbNullString or arrDati(i, 20) == "0":
                F_TEMPME = 0
            else:
                if F_QUANTI > 1:
                    F_TEMPME = CDbl(arrDati(i, 20)) / F_QUANTI
                else:
                    F_TEMPME = CDbl(arrDati(i, 20))
            F_TEMPME = ConvertToSQLDecimal(F_TEMPME)
            # prezzo
            if arrDati(i, 21) == vbNullString: 
                F_PREZZO = 0 
            else: 
                F_PREZZO = arrDati(i, 21)
            if arrDati(i, 23) == vbNullString: 
                F_SCONTO = 0 
            else: 
                F_SCONTO = arrDati(i, 23)
            F_PREZZO = ConvertToSQLDecimal(F_PREZZO)
            F_SCONTO = ConvertToSQLDecimal(F_SCONTO)
            F_VALUTA = "Euro"
            F_NUMPRE = idPreventivo
            F_IDRIGO = i
            F_CODGUI = Right(arrDati(2, 3), 2)
            F_QUANTI = ConvertToSQLDecimal(F_QUANTI)
            
            #stringa di appoggio per campi query rigpre
            campi_rigpre = """(ID_CODPRE, F_DATRIG, F_ORDINE, F_CITFON, F_DESART, F_QUANTI, 
                        F_DANNSR , F_TEMPSR, F_DANNLA, F_TEMPLA, F_DANNVE, F_TEMPVE, 
                        F_TEMPME , F_PREZZO, F_SCONTO, F_VALUTA, 
                        F_NUMPRE , F_IDRIGO, F_CODGUI )"""
                                
            #stringa di appoggo per valori query rigpre
            valori_rigpre = "('" + ID_CODPRE + "', " + F_DATRIG + ", '" + F_ORDINE + "', '" + F_CITFON + "', '" + F_DESART + "', '" + F_QUANTI + "'," 
            valori_rigpre = valori_rigpre + "'" + F_DANNSR + "', " + F_TEMPSR + ",  '" + F_DANNLA + "', " + F_TEMPLA + ", '" + F_DANNVE + "', " + F_TEMPVE + ", " 
            valori_rigpre = valori_rigpre + F_TEMPME + ", " + F_PREZZO + ", " + F_SCONTO + ", '" + F_VALUTA + "', " + F_NUMPRE + ", '" + F_IDRIGO + "', " 
            valori_rigpre = valori_rigpre + "'" + F_CODGUI + "');"
            #fine case C
        elif tipoPratica=="M":  #MECCANICA
            # inizio parametrizzazione variabili
            ID_CODPRE = idPratica
            prevcurr=s.prev
            for pr in s.pratica.listaprev:
                if pr.numprev==idPreventivo:
                    prevcurr=pr
            F_DATRIG = prevcurr.Data_Preventivo
            F_ORDINE = i - 1
            F_CITFON = arrDati(i, 14)   #COD ART PREVENTIVO MECCANICA
            #descrizione articolo.
            originalString = arrDati(i, 15)
            modifiedString = Replace(originalString, "'", " ") 
            F_DESART = Left(modifiedString, 50) 
            if arrDati(i, 13) == 0 or IsNull(arrDati(i, 13)) or arrDati(i, 13) == "":
                F_QUANTI = 1   # QUANTITA' zero imposto 1
            else:
                F_QUANTI = ConvertToSQLDecimal(arrDati(i, 13))   # QUANTITA'
            F_DANNSR = "S"
            F_DANNLA = "S"          #
            '''nei preventivi meccanica ALD le h di MDO inserite nelle righe sono totali,
            'mentre su WinCar vengono moltiplicate per le quantità
            'per ovviare a questo occorre dividere le ore manodopera per le quantità, ove le quantità sono maggiori di 1'''
            if arrDati(i, 16) == "" or arrDati(i, 16) == "0":
                F_TEMPME = 0
            else:
                if F_QUANTI > 1:
                    F_TEMPME = (arrDati(i, 16)) / F_QUANTI
                else:
                    F_TEMPME = (arrDati(i, 16))
            F_TEMPME = ConvertToSQLDecimal(F_TEMPME)
            if arrDati(i, 17) == vbNullString: 
                F_PREZZO = 0 
            else: 
                F_PREZZO = ConvertToSQLDecimal(arrDati(i, 17))
            F_FLAGPR = vbNullString
            if arrDati(i, 18) == vbNullString: 
                F_SCONTO = 0 
            else: 
                F_SCONTO = ConvertToSQLDecimal(arrDati(i, 18))
            F_TIPRIC = "S"         #S di sostituzione per preventivo meccanica
            if arrDati(i, 14) == vbNullString: 
                F___TIPO = vbNullString
            F_VALUTA = "Euro"
            F_NUMPRE = idPreventivo
            F_IDRIGO = i - 1
            F_CODGUI = arrDati(2, 6)
            F_CODIVA = 0
            #controllare se ci sono righe esenti iva, esempio "bolletino postale per revisioni"
            if StrComp((Left((arrDati(i, 15)), 10)), "BOLLETTINO", vbTextCompare) == 0: 
                F_CODIVA = -1    #-1 è il codice per esente iva
            # fine parametrizzazione varibili
                               
            #stringa di appoggio per campi query rigpre
            campi_rigpre = " (ID_CODPRE, F_DATRIG, F_ORDINE, F_CITFON, F_DESART, F_QUANTI, " 
            campi_rigpre = campi_rigpre + "F_DANNSR, F_DANNLA, F_DANNVE, F_TEMPME, F_PREZZO, F_FLAGPR, " 
            campi_rigpre = campi_rigpre + "F_SCONTO, F_TIPRIC, F___TIPO, F_VALUTA, F_NUMPRE, F_IDRIGO, " 
            campi_rigpre = campi_rigpre + "F_CODGUI, F_CODIVA) "
                                
            #stringa di appoggo per valori query rigpre
            valori_rigpre = "(" + ID_CODPRE + ", " + F_DATRIG + ", " + F_ORDINE + ", '" + F_CITFON + "', '" + F_DESART + "', " + F_QUANTI + ", " 
            valori_rigpre = valori_rigpre + "'" + F_DANNSR + "', '" + F_DANNLA + "', '" + F_DANNVE + "', " + F_TEMPME + ", " + F_PREZZO + ", '" + F_FLAGPR + "', " 
            valori_rigpre = valori_rigpre + "" + F_SCONTO + ", '" + F_TIPRIC + "', '" + F___TIPO + "', '" + F_VALUTA + "', " + F_NUMPRE + ", " + F_IDRIGO + ", " 
            valori_rigpre = valori_rigpre + "'" + F_CODGUI + "', " + F_CODIVA + ")"
            #fine CASE M
       
        strSQL = "insert into RIGPRE " + campi_rigpre + " values " + valori_rigpre
        # Esegui la query
        tespre_cursor = conn.cursor()
        tespre_cursor.execute(strSQL)  #creo un cursore/recordset(cursor) da una query
        #dopo la query di inserimento - inserisco una nuova riga su rigpre
        #controllo se il ciclo è alla fine
    if tipoPratica == "C":
        #INSERIMENTO RIGA SMALTIMENTO RIFIUTI
        if arrDati(2, 51) != "":
            strSQL = "insert into RIGPRE (ID_CODPRE, F_DESART, F_QUANTI, F_PREZZO , F_ORDINE, F_NUMPRE, F_IDRIGO, F___TIPO) " 
            strSQL = strSQL + "values ('" + idPratica + "', 'Smaltimento rifiuti', '1', '" + arrDati(2, 51) + "', '" + (i + 1) + "', '" + idPreventivo + "', '" + (i + 1) + "', 'VC')"
            smalt_cursor = conn.cursor()
            smalt_cursor.execute(strSQL) 
    elif tipoPratica=="M":
        #INSERIMENTO RIGA MATERIALI DI CONSUMO
        if arrDati(2, 30) is not None or arrDati(2, 30) != 0 or arrDati(2, 30) != "":
            strSQL = "insert into RIGPRE (ID_CODPRE, F_DESART, F_QUANTI, F_PREZZO , F_ORDINE, F_NUMPRE, F_IDRIGO, F___TIPO) " 
            strSQL = strSQL + "values ('" + idPratica + "', 'Materiali di uso e consumo', '1', '" + arrDati(2, 30) + "', '" + (i + 1) + "', '" + idPreventivo + "', '" + (i + 1) + "', 'VC')"
            # Esegui la query
            matcon_cursor = conn.cursor()
            matcon_cursor.execute(strSQL) 
        #INSERIMENTO RIGA SMALTIMENTO RIFIUTI
        if arrDati(2, 51) != "":
            strSQL = "insert into RIGPRE (ID_CODPRE, F_DESART, F_QUANTI, F_PREZZO , F_ORDINE, F_NUMPRE, F_IDRIGO, F___TIPO) " 
            strSQL = strSQL + "values ('" + idPratica + "', 'Smaltimento rifiuti', '1', '" + arrDati(2, 51) + "', '" + (i + 1) + "', '" + idPreventivo + "', '" + (i + 1) + "', 'VC')"
            smalt_cursor = conn.cursor()
            smalt_cursor.execute(strSQL) 
    #FINE FOR
    ScriviLog("Fine - insert rigpre - pratica n : " + idPratica)
    #fine query import RIPRE   

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
            #idPratica = idPratica
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
        tv.insert('', 'end', values=(row.F_NUMPRA, row.F_DATACA, row.F_RAGSOC, row.F_TARGAV, row.F_IMPPRE))
                  #, f"€ {int(row.F_TOTRIC):.2f}"))
    return rows

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

def nuovaPratica(idpratica, filename):
    dataAttuale = datetime.datetime.now() # Ottieni la data corrente
    # Converti la data nel formato yyyymmdd
    dataFormattata = dataAttuale.strftime("%Y%m%d") #### (dataAttuale.__format__("yyyymmdd"))
    
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
        prev0 = s.pratica.listaprev[0]
        f_numpra = idPratica
        f_targav = prev0.Targa_Veicolo          #                       arrDati(2, 8)                
        datatemp = datetime.datetime.fromtimestamp(int(prev0.Data_preventivo)/1000)
        f_dataca = datatemp.strftime("%d/%m/%Y")    #####datetime.datetime.fromtimestamp(datapr/1000)
        datatemp = datetime.datetime.fromtimestamp(int(prev0.Data_Immatricolazione)/1000)
        f_datcre = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")   #data e ora creazione pratica interna
        f_datimm = datatemp.strftime("%d/%m/%Y %H:%M:%S")
        f_desmod = prev0.Descrizione_Veicolo    #                       Left(arrDati(2, 4), 70)
        f_telaio = prev0.Telaio.replace("Telaio ","")
        f_kimvei = prev0.Km                     #                       arrDati(2, 6)
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
        valori_carvei = "(" + str(f_numpra) + ",'" + f_targav + "','" + f_dataca + "','" + f_desmod 
        valori_carvei = valori_carvei + "','" + f_telaio + "','" + str(f_kimvei) + "','" + f_datimm + "', " 
        valori_carvei = valori_carvei + s.pratica.F_CODCLI + ",'" + s.pratica.F_RAGSOC + "','" + s.pratica.F_VIACLI + "','"
        valori_carvei = valori_carvei + s.pratica.F_CITTAC + "','" + s.pratica.F_CAPCLI + "','" + s.pratica.F_PROCLI + "','" + s.pratica.F_PARIVA
        valori_carvei = valori_carvei +  + "','" + s.pratica.F_TELEFO + "', " + "'" + f_tipove + "','" + f_tpreve + "','" + f_idmess
        valori_carvei = valori_carvei  + "','" + f_datcre + "','" + F___GUID + "', '" + F_DESCOL + "')"
        #fine case Carr 
    elif s.tipoPratica == "M":
        campi_carvei = "(f_numpra, f_targav, f_dataca, f_desmod, f_telaio, f_kimvei, f_datimm, "
        campi_carvei = campi_carvei + "F_CODCLI, F_RAGSOC, F_VIACLI, F_CITTAC, F_CAPCLI, F_PROCLI, F_PARIVA, F_TELEFO, "
        campi_carvei = campi_carvei + "f_nummot, f_tipove, f_tpreve, f_idmess, f_datcre, F___GUID)"""
         
        f_numpra = idPratica
        prev0 = s.pratica.listaprev[0]
        f_targav = prev0.Targa_Veicolo.replace("Targa Veicolo ","")
        datatemp = datetime.datetime.fromtimestamp(int(prev0.Data_preventivo)/1000)
        f_dataca = datatemp.strftime("%d/%m/%Y")    #####datetime.datetime.fromtimestamp(datapr/1000)
        datatemp = datetime.datetime.fromtimestamp(int(prev0.Data_Immatricolazione)/1000)
        f_datcre = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")   #data e ora creazione pratica interna
        f_datimm = datatemp.strftime("%d/%m/%Y %H:%M:%S")
        f_desmod = prev0.Descrizione_Veicolo.replace("'", "''")
        f_desmod = f_desmod[0:70]
        f_telaio = prev0.Telaio.replace("Telaio ","")
        f_kimvei = prev0.Km
        #modifica del 24/03/2025 per i clienti privati oltre ALD
        #per i dati del cliente sono impostati su parametri.ini
        ####NON SERVE
        # if prev0.Id_riparazione != "":    #se la colonna IdRip. contiene testo, è ALD
        #    leggi_par_ini()

        f_nummot = prev0.Km  #arrDati(2, 6)
        f_tipove = "O"   #tipo vernice di default metto doppio strato
        f_tpreve = "M"   #tipo logo M per meccanica
        f_idmess = dataFormattata + str(idPratica)   #id mess
        F___GUID = idPratica
    
        valori_carvei = "(" + str(f_numpra) + ",'" + f_targav 
        valori_carvei = valori_carvei + "','" + f_dataca + "','" + f_desmod
        valori_carvei = valori_carvei + "','" + f_telaio + "','" + str(f_kimvei)
        valori_carvei = valori_carvei + "','" + f_datimm + "', " 
        valori_carvei = valori_carvei + "'" + s.pratica.F_CODCLI + "','" + s.pratica.F_RAGSOC + "','" + s.pratica.F_VIACLI + "','" + s.pratica.F_CITTAC + "','"
        valori_carvei = valori_carvei + s.pratica.F_CAPCLI + "','" + s.pratica.F_PROCLI + "','" + s.pratica.F_PARIVA + "','" + s.pratica.F_TELEFO + "', " 
        valori_carvei = valori_carvei + "'" + f_nummot + "','" + f_tipove + "','" + f_tpreve
        valori_carvei = valori_carvei + "','" + f_idmess + "','" + str(f_datcre) + "','" + str(F___GUID) + "')"
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
    if s.tipoPratica == "C":
        strSQL = "INSERT INTO Pratica2 (f_numpra, f_tippra, F_CODCOL) values (" + str(idPratica) + ", '1', '" + StringaTraParentesi + "')"
    elif s.tipoPratica == "M":
        strSQL = "INSERT INTO Pratica2 (f_numpra, f_tippra) values (" + str(idPratica) + ", '2')"

    ScriviLog("import.py - insert pratica2")
    # Esegui la query
    carvei_ins_cursor.execute(strSQL)
    #fine insert pratica2
    return idPratica 
    
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
btnAvvia=tk.Button(text="Leggi files", command=Import_JSON, font=Font_tuple, fg="yellow", bg="blue")
btnImporta=tk.Button(text="Importa su WinCar", command=Import_Dati, font=Font_tuple, fg="yellow", bg="blue", state="disabled")
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
btnImporta.grid(row=3, column=1, sticky="WE", padx=20, pady=30)
files_file=[]

if __name__ == "__main__":
    window.mainloop()
