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
conn = pyodbc.Connection
files_file=[]
'''     TESTING CONNECTIONS GICO
print(pyodbc.dataSources())
nomedb="c:\\Users\\cicco\\Downloads\\wcArchivi.mdb"
try:
    #global conn
    conn=pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};);Dbq=' + nomedb + ';Uid=;Pwd=;')
    print("ok")
except:
    print(conn.__doc__)
'''
'''
nomedb="C:\\THS\\THS32Env\\wcArchivi.mdb"
filejson="export.json"
global targetFolder
targetFolder=""
mylog=""
lockFilePath=""
tipoPratica="M" #Meccanica #Carrozzeria
'''
######## Creazione eseguibile ########################################################
######## c:\HTS\THS32Env\.venv\Scripts\pyinstaller C:\HTS\THS32Env\import.py --onefile
######## Creazione eseguibile ########################################################

def ScriviLog(messaggio):
    filePath = s.mylog
    if filePath == "":
        tk.messagebox.showerror("Errore nel percorso del file di log txt", "Errore file log")  
    
    dataOra = datetime.datetime.now()
    filelog=open(filePath, "a", encoding="UTF-8")
    # Scrivi la data, l'ora e il messaggio nel file di log
    filelog.write(str(dataOra) +  s.tipoPratica +" - " + messaggio + "\n")
    # Chiudi il file
    filelog.close()

#stringa di connessione al db
def connetti():
    print(pyodbc.dataSources())
    #Connessione su sistema a 32 bit
    connstr = r'DRIVER={Microsoft Access Driver (*.mdb)};);Dbq=' + s.nomedb + ';Uid=;Pwd=;'
    try:
        global conn
        conn=pyodbc.connect(connstr)
        print("Connection established")
    except:
        try:
            #Connessione su sistema a 64 bit
            print("Secondo tentativo di connessione...")
            connstr = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};);Dbq=' + s.nomedb + ';Uid=;Pwd=;'
            conn=pyodbc.connect(connstr)
            print("Connection established")
        except:
            tk.messagebox.showerror("Errore accesso", "Errore di connessione al database! " + conn.__doc__)
            print(conn.__doc__) 
            exit      

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

def leggijson(pfile, nprev):
    leggi_par_ini()
    pratcurr=s.pratica()
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
        pratcurr.desc="Pratica"
        prev = s.prev()
        for pre in data:
            prev.FILE=file1
            prev.Smalt_Rifiuti=None
            if "CV_F_TIPPRE" in data:
                prev.F_TIPPRE="MECC"
                for x in s.s_header_mecc:    #INTESTAZIONE PREVENTIVO
                    #qui dovrei leggere l'intestazione se si tratta di ALD o meno
                    com=""
                    try:
                        if type(data) is dict:
                            val=data[x]
                        else:
                            val=pre[x]
                        if isinstance(val, str):
                            com="prev.%s = '%s'" % (x.replace(" ", "_").replace(".",""), val.replace("'",""))
                        else:
                            com="prev.%s = '%s'" % (x.replace(" ", "_").replace(".",""), val)
                        exec(com)
                    except:
                        print(file1, ":", x, "--->", com)
                    #leggo le righe nel file e creo dei campi in una struttura che memorizzi i dati del preventivo
            else:
                s.prev.F_TIPPRE="CARR"
                for x in s.s_header_carr:    #INTESTAZIONE PREVENTIVO
                    #qui dovrei leggere l'intestazione se si tratta di ALD o meno
                    com=""
                    try:
                        if type(data) is dict:
                            val=data[x]
                        else:
                            val=pre[x]
                        if isinstance(val, str):
                            com="prev.%s = '%s'" % (x.replace(" ", "_").replace(".",""), val.replace("'",""))
                        else:
                            com="prev.%s = '%s'" % (x.replace(" ", "_").replace(".",""), val)
                        exec(com)
                    except:
                        print(file1, ":", x, "--->", com)
                    #leggo le righe nel file e creo dei campi in una struttura che memorizzi i dati del preventivo
            #RIGHE PREVENTIVO
            try:
                if type(data) is dict:
                    if prev.F_TIPPRE=="MECC":
                        tab=data["Tabella Interventi Meccanica"]
                    else:
                        tab=data["RIGPRE_carr"]
                else:
                    if prev.F_TIPPRE=="MECC":
                        tab=pre["Tabella Interventi Meccanica"]   
                    else:
                        tab=pre["RIGPRE_carr"]   
            except:
                tab=[]
            righe=[]
            i=0  
            prev.listrighe=[] 
            for e in tab:
                el=s.riga()
                if prev.F_TIPPRE=="MECC":
                    for x in s.s_obmecc:
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
                else:
                    for x in s.s_obcarr:
                        if type(data) is dict:
                            listaint=data["RIGPRE_carr"]
                            interv=listaint[i]
                            val=interv[x] 
                        else:
                            val=pre["RIGPRE_carr"][i][x] 
                        if isinstance(val, str):
                            com='el.%s = "%s"' % (x.replace(" ", "_").replace(".",""), val.replace("'",""))
                            com='el.%s = "%s"' % (x.replace("\"", ""), val.replace("\"",""))
                        else:
                            com='el.%s = "%s"' % (x.replace(" ", "_").replace(".",""), val)
                        exec(com)
                i+=1
                prev.addriga(el)
                righe.append(el)
            
            pratcurr.F_TARGAV=(prev.CV_F_TARGAV).replace("Targa Veicolo ","")
            pratcurr.F_TELAIO=(prev.CV_F_TELAIO).replace("Telaio ","")
            pratcurr.F_DATACA=prev.CV_F_DATACA
            pratcurr.F_DATIMM=prev.CV_F_DATIMM
            pratcurr.F_DESMOD=prev.CV_F_DESMOD
            if prev.F_TIPPRE=="MECC":
                pratcurr.F_NUMMOT=prev.CV_F_NUMMOT
                pratcurr.F_KIMVEI=prev.CV_F_KIMVEI
                pratcurr.F_TIPOVE=""
                pratcurr.F_DESCOL=""
            else:
                pratcurr.F_DESCOL=prev.CV_F_DESCOL
                pratcurr.F_TIPOVE=prev.CV_F_TIPOVE
                pratcurr.F_KIMVEI=0
                pratcurr.CV_F_KIMVEI=0
            prev.numprev=nprev
            pratcurr.addprev(prev)
            if type(data) is dict:
                break
        #in righe ho tutte le righe del preventivo (s.elem) e s.prev la testata/piede
        s.importazione.addpra(s.importazione, pratcurr)
        s.importazione.numpra = s.importazione.numpra + 1

        save_path_file = "EXPDANNI_" + pfile.replace(".json","") + ".txt"
        f = open(save_path_file, "w")
        if prev.F_TIPPRE=="MECC":
            for i in range(len(prev.listrighe)):
                txt_str=""  
                txt_str = txt_str + '0' * 11       #riempio con 0
                txt_str = txt_str + format(prev.CV_F_TARGAV,">12") + "   000000000" + ' ' * 19     
                txt_str = txt_str + format(prev.CV_F_TELAIO,">17") + ' ' * 3     
                txt_str = txt_str + "0    1000000000" + format(prev.listrighe[i].RG_F_CITFON if prev.listrighe[i].RG_F_CITFON!="None" else ' ', "21")      
                txt_str = txt_str + "0000" + "S" ####(f"{prev.RG_FLAG_CITFON: <1}") 
                txt_str = txt_str + "0000" +  format(prev.listrighe[i].RG_F_DESART," <50") + ' ' * 30 + "00000"   
                txt_str = txt_str + "000000000000000"
                txt_str = txt_str + f"{prev.listrighe[i].RG_F_TEMPME if prev.listrighe[i].RG_F_TEMPME!="None" else 0:0>4}"
                txt_str = txt_str + format(prev.listrighe[i].RG_F_PREZZO if prev.listrighe[i].RG_F_PREZZO!="None" else 0,"0>9") 
                txt_str = txt_str + "       0       S    00000                    " + ' ' * 110
                #txt_str = txt_str + format(prev.listrighe[i].RG_F_DANNSR," >1") + format(prev.listrighe[i].RG_F_DANNLA," >1") + format(prev.listrighe[i].RG_F_DANNVE," >1")
                #controllare
                txt_str = txt_str + "              0      1    1133 1601 1000                                     " 
                txt_str = txt_str + "1     " + "    " + "RC" # prev.listrighe[i].RG_F___TIPO # (es. RC/VC/MD/OL)  
                txt_str = txt_str + format(prev.listrighe[i].RG_F_QUANTI if prev.listrighe[i].RG_F_QUANTI!="None" else 0," <30") + "\n"
                f.write(txt_str)
        else:
            for i in range(len(prev.listrighe)):
                txt_str=""  
                txt_str = txt_str + format(prev.PIva_Riparatore,'0<11')       #riempio con 0
                txt_str = txt_str + format(prev.CV_F_TARGAV,">12") + "   000000000" + ' ' * 19    
                txt_str = txt_str + format(prev.CV_F_TELAIO,"<17") + ' ' * 3     
                txt_str = txt_str + "0    1000000000" + format(prev.listrighe[i].RG_F_CITFON if prev.listrighe[i].RG_F_CITFON!="None" else ' ', "21")       
                txt_str = txt_str + "0000" + "S" ####(f"{prev.RG_FLAG_CITFON: <1}") 
                txt_str = txt_str + "0000" +  format(prev.listrighe[i].RG_F_DESART," <50") + ' ' * 30  + "00000"   
                txt_str = txt_str + f"{prev.listrighe[i].RG_F_TEMPSR if prev.listrighe[i].RG_F_TEMPSR!="None" else 0:0>4}" + '0'
                txt_str = txt_str + f"{prev.listrighe[i].RG_F_TEMPLA if prev.listrighe[i].RG_F_TEMPLA!="None" else 0:0>4}" + '0'    
                txt_str = txt_str + f"{prev.listrighe[i].RG_F_TEMPVE if prev.listrighe[i].RG_F_TEMPVE!="None" else 0:0>4}" + '0'
                txt_str = txt_str + f"{prev.listrighe[i].RG_F_TEMPME if prev.listrighe[i].RG_F_TEMPME!="None" else 0:0>4}"
                txt_str = txt_str + format(prev.listrighe[i].RG_F_PREZZO if prev.listrighe[i].RG_F_PREZZO!="None" else 0,"0>9") 
                txt_str = txt_str + format(prev.listrighe[i].RG_F_DANNSR if prev.listrighe[i].RG_F_DANNSR!="None" else ' '," >1")
                txt_str = txt_str + format(prev.listrighe[i].RG_F_DANNLA if prev.listrighe[i].RG_F_DANNLA!="None" else ' '," >1") 
                txt_str = txt_str + format(prev.listrighe[i].RG_F_DANNVE if prev.listrighe[i].RG_F_DANNVE!="None" else ' '," >1")
                txt_str = txt_str + "    0       S    00000                    " + ' ' * 110
                #controllare
                txt_str = txt_str + "              0      1    1133 1601 1000                                     " 
                txt_str = txt_str + "1     " + "    " + "RC" # prev.listrighe[i].RG_F___TIPO # (es. RC/VC/MD/OL)  
                txt_str = txt_str + format(prev.listrighe[i].RG_F_QUANTI if prev.listrighe[i].RG_F_QUANTI!="None" else 0," <30") + "\n"
                f.write(txt_str)
        if prev.Smalt_Rifiuti is not None:
            txt_str=""  
            if prev.F_TIPPRE=="MECC":
                txt_str = txt_str + '0' * 11       #riempio con 0
            else:
                txt_str = txt_str + format(prev.PIva_Riparatore,'0<11')
            txt_str = txt_str + format(prev.CV_F_TARGAV,">12") + "   000000000" + ' ' * 19     
            txt_str = txt_str + format(prev.CV_F_TELAIO,">17") + ' ' * 3     
            txt_str = txt_str + "0    1000000000" + ' '*21      
            txt_str = txt_str + "0000S0000" +  ' ' * 80 + "000000000000000000000000"
            txt_str = txt_str + format(prev.Smalt_Rifiuti if prev.Smalt_Rifiuti!="None" else 0,"0>9") 
            txt_str = txt_str + "       0       S    00000                    " + ' ' * 110
            #txt_str = txt_str + format(prev.listrighe[i].RG_F_DANNSR," >1") + format(prev.listrighe[i].RG_F_DANNLA," >1") + format(prev.listrighe[i].RG_F_DANNVE," >1")
            #controllare
            txt_str = txt_str + "              0      1    1133 1601 1000                                     " 
            txt_str = txt_str + "0     " + "SR  " + "VC" # prev.listrighe[i].RG_F___TIPO # (es. RC/VC/MD/OL)  
            txt_str = txt_str + format(1," <30") + "\n"
            f.write(txt_str)    
        f.close()
        
        save_path_file = "EXPDANNI2_" + pfile.replace(".json","") + ".txt"
        f = open(save_path_file, "w")
        if prev.F_TIPPRE=="MECC":
            for i in range(len(prev.listrighe)):
                txt_str=""  
                txt_str = txt_str + '0' * 16       #riempio con 0
                txt_str = txt_str + format(prev.CV_F_TARGAV,">7") + "   00000000"
                txt_str = txt_str + format(prev.CV_F_TELAIO,">17") + ' ' * 4 + " " * 20 #NUMSIN   
                txt_str = txt_str + "0    1000000000" + format(prev.listrighe[i].RG_F_CITFON if prev.listrighe[i].RG_F_CITFON!="None" else ' ', "21")      
                txt_str = txt_str + "0000" + "S" ####(f"{prev.RG_FLAG_CITOFON: <1}") 
                txt_str = txt_str + "0000" +  format(prev.listrighe[i].RG_F_DESART," <50") + ' ' * 30 + "00000"   
                txt_str = txt_str + "000000000000000"
                txt_str = txt_str + f"{prev.listrighe[i].RG_F_TEMPME if prev.listrighe[i].RG_F_TEMPME!="None" else 0:0>4}"
                txt_str = txt_str + format(prev.listrighe[i].RG_F_PREZZO if prev.listrighe[i].RG_F_PREZZO!="None" else 0,"0>9") 
                txt_str = txt_str + "       0       S    00000                    " + ' ' * 110
                #controllare
                txt_str = txt_str + "              0      1    1133 1601 1000                                     " 
                txt_str = txt_str + "1     " + "    " + "RC" # prev.listrighe[i].RG_F___TIPO # (es. RC/VC/MD/OL)  
                txt_str = txt_str + format(prev.listrighe[i].RG_F_QUANTI if prev.listrighe[i].RG_F_QUANTI!="None" else 0," <30") 
                txt_str = txt_str + "0000000000000000      " + format(prev.listrighe[i].RG_F_SCONTO if prev.listrighe[i].RG_F_SCONTO!="None" else 0,"0<4")
                txt_str = txt_str + "0" + format(prev.listrighe[i].RG_F_PREZZO if prev.listrighe[i].RG_F_PREZZO!="None" else 0,"0>9") 
                txt_str = txt_str + "00" +"\n"
                f.write(txt_str)
        else:
            for i in range(len(prev.listrighe)):
                txt_str=""  
                txt_str = txt_str + format(prev.PIva_Riparatore,'0<11') + "00000"     #riempio con 0
                txt_str = txt_str + format(prev.CV_F_TARGAV,">7") + "   000000000"  
                txt_str = txt_str + format(prev.CV_F_TELAIO,"<17") + ' ' * 4 + " " * 20 #NUMSIN    
                txt_str = txt_str + "0    1000000000" + format(prev.listrighe[i].RG_F_CITFON if prev.listrighe[i].RG_F_CITFON!="None" else ' ', "21")       
                txt_str = txt_str + "0000" + "S" ####(f"{prev.RG_FLAG_CITFON: <1}") 
                txt_str = txt_str + "0000" +  format(prev.listrighe[i].RG_F_DESART," <50") + ' ' * 30  + "00000"   
                txt_str = txt_str + f"{prev.listrighe[i].RG_F_TEMPSR if prev.listrighe[i].RG_F_TEMPSR!="None" else 0:0>4}" + '0'
                txt_str = txt_str + f"{prev.listrighe[i].RG_F_TEMPLA if prev.listrighe[i].RG_F_TEMPLA!="None" else 0:0>4}" + '0'    
                txt_str = txt_str + f"{prev.listrighe[i].RG_F_TEMPVE if prev.listrighe[i].RG_F_TEMPVE!="None" else 0:0>4}" + '0'
                txt_str = txt_str + f"{prev.listrighe[i].RG_F_TEMPME if prev.listrighe[i].RG_F_TEMPME!="None" else 0:0>4}"
                txt_str = txt_str + format(prev.listrighe[i].RG_F_PREZZO if prev.listrighe[i].RG_F_PREZZO!="None" else 0,"0>9") 
                txt_str = txt_str + format(prev.listrighe[i].RG_F_DANNSR if prev.listrighe[i].RG_F_DANNSR!="None" else ' '," >1")
                txt_str = txt_str + format(prev.listrighe[i].RG_F_DANNLA if prev.listrighe[i].RG_F_DANNLA!="None" else ' '," >1") 
                txt_str = txt_str + format(prev.listrighe[i].RG_F_DANNVE if prev.listrighe[i].RG_F_DANNVE!="None" else ' '," >1")
                txt_str = txt_str + "    0       S    00000                    " + ' ' * 110
                #controllare
                txt_str = txt_str + "              0      1    1133 1601 1000                                     " 
                txt_str = txt_str + "1     " + "    " + "RC" # prev.listrighe[i].RG_F___TIPO # (es. RC/VC/MD/OL)  
                txt_str = txt_str + format(prev.listrighe[i].RG_F_QUANTI if prev.listrighe[i].RG_F_QUANTI!="None" else 0," <30") 
                txt_str = txt_str + "0000000000000000      " + format(prev.listrighe[i].RG_F_SCONTO if prev.listrighe[i].RG_F_SCONTO!="None" else 0,"0<4")
                txt_str = txt_str + "0" + format(prev.listrighe[i].RG_F_PREZZO if prev.listrighe[i].RG_F_PREZZO!="None" else 0,"0>9") 
                txt_str = txt_str + "00" +"\n"
                f.write(txt_str)
        if prev.Smalt_Rifiuti is not None:
            txt_str=""  
            if prev.F_TIPPRE=="MECC":
                txt_str = txt_str + '0' * 11 + "00000"      #riempio con 0
            else:
                txt_str = txt_str + format(prev.PIva_Riparatore,'0<11') + "00000"
            txt_str = txt_str + format(prev.CV_F_TARGAV,">7") + "   00000000"
            txt_str = txt_str + format(prev.CV_F_TELAIO,">17") + ' ' * 4 + " " * 20 #NUMSIN   
            txt_str = txt_str + "0    1000000000" + " " * 21      
            txt_str = txt_str + "0000" + "S" ####(f"{prev.RG_FLAG_CITOFON: <1}") 
            txt_str = txt_str + "0000" +  format("Smalt.Rifiuti (3,00%) su Imponibile Totale"," <50") + ' ' * 30 + "00000"   
            txt_str = txt_str + "0000000000000000000"
            txt_str = txt_str + format(prev.Smalt_Rifiuti if prev.Smalt_Rifiuti!="None" else 0,"0>9") 
            txt_str = txt_str + "       0       S    00000                    " + ' ' * 110
            #controllare
            txt_str = txt_str + "              0      1    1133 1601 1000                                     " 
            txt_str = txt_str + "1     " + "SR  " + "VC" # prev.listrighe[i].RG_F___TIPO # (es. RC/VC/MD/OL)  
            txt_str = txt_str + format(1," <30") 
            txt_str = txt_str + "0000000000000000      0000"
            txt_str = txt_str + "0" + format(prev.Smalt_Rifiuti if prev.Smalt_Rifiuti!="None" else 0,"0>9") 
            txt_str = txt_str + "00" +"\n"
            f.write(txt_str)        
        f.close()

        save_path_file = "EXPERIZ_" + pfile.replace(".json","") + ".txt"
        f1 = open(save_path_file, "w")
        if prev.F_TIPPRE=="MECC":
            txt_str=""  
            txt_str = txt_str + '0' * 11 + "     "       #riempio con 0
            txt_str = txt_str + format(prev.CV_F_TARGAV,">10") + " " * 8    
            txt_str = txt_str + format(prev.CV_F_TELAIO,"<17") + ' ' * 3 + "0" * 20 + "0    0       "  
            txt_str = txt_str + format(s.pratica.F_RAGSOC,"<40") 
            txt_str = txt_str + format(s.pratica.F_CAPCLI," <5")
            txt_str = txt_str + format(s.pratica.F_CITTAC," <33")
            txt_str = txt_str + format(s.pratica.F_PROCLI," <2") + " " * 20
            txt_str = txt_str + format(s.pratica.F_VIACLI," <80") + "0" + " " * 22
            txt_str = txt_str + format(s.pratica.F_RAGSOC," <110") + f"{datetime.date.today().year:0<4}"
            txt_str = txt_str + " " * 103 + "0" + " " * 7  + "0" + " " * 70 + "00"
            txt_str = txt_str + format(s.pratica.F_RAGSOC," <35") + format(s.pratica.F_VIACLI," <35")
            txt_str = txt_str + format(s.pratica.F_TELEFO," <16") + " " * 142
            txt_str = txt_str + format(pratcurr.F_DESMOD," <145")
            txt_str = txt_str + format(pratcurr.F_KIMVEI,"0<7") + "000"
            txt_str = txt_str + format(pratcurr.F_DESCOL," <20")
            txt_str = txt_str + format(pratcurr.F_TIPOVE," <20") + " " * 83 + "0     0    0    0 0 "
            #preant
            txt_str = txt_str + "00000000000000000.000         0                                       0                 "
            dataimm = datetime.datetime.fromtimestamp(int(pratcurr.F_DATIMM)/1000)
            txt_str = txt_str + dataimm.strftime("%d%m%Y") + " " * 44
            txt_str = txt_str + format(s.pratica.F_RAGSOC," <40") + "0"
            txt_str = txt_str + "000000000000000"
            txt_str = txt_str + f"{prev.TS_F_TOTME if prev.TS_F_TOTME!="None" else 0:0>4}" + '0'
            txt_str = txt_str + f"{prev.TS_F_TOTRIC if prev.TS_F_TOTRIC!="None" else 0:0>10}" + '0'
            try:
                txt_str = txt_str + f"{prev.TS_F_TSUPPL if prev.TS_F_TSUPPL!="None" else 0:0>4}" + '0'
            except:
                txt_str = txt_str + "00000"
            try:
                txt_str = txt_str + f"{prev.TS_F_TFINIT if prev.TS_F_TFINIT!="None" else 0:0>4}" + '0'
            except:
                txt_str = txt_str + "00000"
            try:
                txt_str = txt_str + f"{prev.TS_F_TCOMPL if prev.TS_F_TCOMPL!="None" else 0:0>4}" + '0'
            except:
                txt_str = txt_str + "00000"
            try:
                txt_str = txt_str + f"{prev.TS_F_TEMSUP if prev.TS_F_TEMSUP!="None" else 0:0>4}" + '0'
            except:
                txt_str = txt_str + "00000"
            try:
                txt_str = txt_str + f"{prev.TS_F_VEOPER if prev.TS_F_VEOPER!="None" else 0:0>4}" + '0'
            except:
                txt_str = txt_str + "00000"
            txt_str = txt_str + "         0         0 0 0         0         0         "
            txt_str = txt_str + f"{prev.TS_F_TOTPRE if prev.TS_F_TOTPRE!="None" else 0:0>10}"
            txt_str = txt_str + f"{prev.TS_F_TOTIVA if prev.TS_F_TOTIVA!="None" else 0:0>10}" + "0 0         0"
            txt_str = txt_str + f"{prev.TS_F_TOTIVA if prev.TS_F_TOTIVA!="None" else 0:0>10}" + "0          "
            txt_str = txt_str + f"{prev.TS_F_TOTRIC if prev.TS_F_TOTRIC!="None" else 0:0>10}"
            try:
                txt_str = txt_str + f"{prev.TS_F_IIVARIC if prev.TS_F_IIVARIC!="None" else 0:0>10}" + "0"
                txt_str = txt_str + f"{prev.TS_F_TIVARIC if prev.TS_F_TIVARIC!="None" else 0:0>10}"
            except:
                txt_str = txt_str + "0" * 21
            txt_str = txt_str + f"{prev.TS_F_TOTMAT if prev.TS_F_TOTMAT!="None" else 0:0>10}"
            try:
                txt_str = txt_str + f"{prev.TS_F_IIVAMAT if prev.TS_F_IIVAMAT!="None" else 0:0>10}" + "0"
                txt_str = txt_str + f"{prev.TS_F_TIVAMAT if prev.TS_F_TIVAMAT!="None" else 0:0>10}" 
            except:
                txt_str = txt_str + "0" * 21
            txt_str = txt_str + f"{prev.TS_F_TOTCOM if prev.TS_F_TOTCOM!="None" else 0:0>10}" 
            try:
                txt_str = txt_str + f"{prev.TS_F_IIVAVAR if prev.TS_F_IIVAVAR!="None" else 0:0>10}" + "0"  
                txt_str = txt_str + f"{prev.TS_F_TIVAVAR if prev.TS_F_TIVAVAR!="None" else 0:0>10}"
            except:
                txt_str = txt_str + "0" * 21
            try:
                txt_str = txt_str + f"{prev.TS_F_MANCAR if prev.TS_F_MANCAR!="None" else 0:0>10}"
                txt_str = txt_str + f"{prev.TS_F_IIVACAR if prev.TS_F_IIVACAR!="None" else 0:0>10}" + "0"  
                txt_str = txt_str + f"{prev.TS_F_TIVACAR if prev.TS_F_TIVACAR!="None" else 0:0>10}"  
            except:
                txt_str = txt_str + "0" * 31
            txt_str = txt_str + f"{prev.TS_F_MANMEC if prev.TS_F_MANMEC!="None" else 0:0>10}"
            try:
                txt_str = txt_str + f"{prev.TS_F_IIVAMEC if prev.TS_F_IIVAMEC!="None" else 0:0>10}" + "0"  
                txt_str = txt_str + f"{prev.TS_F_TIVAMEC if prev.TS_F_TIVAMEC!="None" else 0:0>10}" + "0" + " " * 10 
            except:
                txt_str = txt_str + "0" * 21 + " " * 10
            txt_str = txt_str + "NON CONCORDATO0" + "  " + " " * 230    #"CV_F_AUTRIP len="230"
            txt_str = txt_str + " " * 8 + " " * 28 + "00000000C0" + "0000" + "00"  #"CV_F_DTPRCO" len="8" 
            try:
                txt_str = txt_str + f"{prev.TS_F_COSTOR if prev.TS_F_COSTOR!="None" else 0:0<5}" 
            except:
                txt_str = txt_str + "00000" 
            #Con align="L" pad="0" un valore come 0.00 (4 char) viene riempito a destra → 0.000. Con riempimento a sinistra si ottiene 00.00. -->
            #"TS_F_COSTOR" pos="1924" len="5" align="R" pad="0" source="cell" address="//TS_F_COSTOR" format="dec2_dot"/>
            txt_str = txt_str + "0" + f"{prev.TS_F_TOTME if prev.TS_F_TOTME!="None" else 0:0>4}" + "00"
            txt_str = txt_str + f"{prev.TS_F_COSTO2 if prev.TS_F_COSTO2!="None" else 0:0>5}" + "C70011          0000.000000.00EURTE\n"
            f1.write(txt_str)
        else:
            txt_str=""  
            txt_str = txt_str + format(prev.PIva_Riparatore,'0<11') + "     "     
            txt_str = txt_str + format(prev.CV_F_TARGAV,">10") + " " * 8    
            txt_str = txt_str + format(prev.CV_F_TELAIO,"<17") + ' ' * 3 + " " * 20 + "0    0       "  
            txt_str = txt_str + format(s.pratica.F_RAGSOC,"<40") 
            txt_str = txt_str + format(s.pratica.F_CAPCLI," <5")
            txt_str = txt_str + format(s.pratica.F_CITTAC," <33")
            txt_str = txt_str + format(s.pratica.F_PROCLI," <2") + " " * 20
            txt_str = txt_str + format(s.pratica.F_VIACLI," <80") + "0" + " " * 22
            txt_str = txt_str + format(s.pratica.F_RAGSOC," <110") + f"{datetime.date.today().year:0<4}"
            txt_str = txt_str + " " * 103 + "0" + " " * 7  + "0" + " " * 70 + "00"
            txt_str = txt_str + format(s.pratica.F_RAGSOC," <35") + format(s.pratica.F_VIACLI," <35")
            txt_str = txt_str + format(s.pratica.F_TELEFO," <16") + " " * 142
            txt_str = txt_str + format(pratcurr.F_DESMOD," <145")
            txt_str = txt_str + format(pratcurr.F_KIMVEI,"0<7") + "000"
            txt_str = txt_str + format(pratcurr.F_DESCOL," <20")
            txt_str = txt_str + format(pratcurr.F_TIPOVE," <20") + " " * 83 + "0     0    0    0 0 "
            #preant
            txt_str = txt_str + "00000000000000000.000         0                                       0                 "
            dataimm = datetime.datetime.fromtimestamp(int(pratcurr.F_DATIMM)/1000)
            txt_str = txt_str + dataimm.strftime("%d%m%Y") + " " * 44
            txt_str = txt_str + format(s.pratica.F_RAGSOC," <40") + "0"
            txt_str = txt_str + f"{prev.TS_F_TOTSR if prev.TS_F_TOTSR!="None" else 0:0>4}" + '0'
            txt_str = txt_str + f"{prev.TS_F_TOTLA if prev.TS_F_TOTLA!="None" else 0:0>4}" + '0'    
            txt_str = txt_str + f"{prev.TS_F_TOTVE if prev.TS_F_TOTVE!="None" else 0:0>4}" + '0'
            txt_str = txt_str + f"{prev.TS_F_TOTME if prev.TS_F_TOTME!="None" else 0:0>4}" + '0'
            txt_str = txt_str + f"{prev.TS_F_TOTRIC if prev.TS_F_TOTRIC!="None" else 0:0>10}" + '0'
            txt_str = txt_str + f"{prev.TS_F_TSUPPL if prev.TS_F_TSUPPL!="None" else 0:0>4}" + '0'
            txt_str = txt_str + f"{prev.TS_F_TIFINIT if prev.TS_F_TIFINIT!="None" else 0:0>4}" + '0'
            txt_str = txt_str + f"{prev.TS_F_TCOMPL if prev.TS_F_TCOMPL!="None" else 0:0>4}" + '0'
            txt_str = txt_str + f"{prev.TS_F_TEMSUP if prev.TS_F_TEMSUP!="None" else 0:0>4}" + '0'
            try:
                txt_str = txt_str + f"{prev.TS_F_VEOPER if prev.TS_F_VEOPER!="None" else 0:0>4}" + '0'
            except:
                txt_str = txt_str + '00000'
            txt_str = txt_str + "         0         0 0 0         0         0         "
            txt_str = txt_str + f"{prev.TS_F_TOTPRE if prev.TS_F_TOTPRE!="None" else 0:0>10}"
            txt_str = txt_str + f"{prev.TS_F_TOTIVA if prev.TS_F_TOTIVA!="None" else 0:0>10}" + "0 0         0"
            txt_str = txt_str + f"{prev.TS_F_TOTIVA if prev.TS_F_TOTIVA!="None" else 0:0>10}" + "0          "
            txt_str = txt_str + f"{prev.TS_F_TOTRIC if prev.TS_F_TOTRIC!="None" else 0:0>10}"
            txt_str = txt_str + f"{prev.TS_F_IIVARIC if prev.TS_F_IIVARIC!="None" else 0:0>10}" + "0"
            txt_str = txt_str + f"{prev.TS_F_TIVARIC if prev.TS_F_TIVARIC!="None" else 0:0>10}"
            txt_str = txt_str + f"{prev.TS_F_TOTMAT if prev.TS_F_TOTMAT!="None" else 0:0>10}"
            txt_str = txt_str + f"{prev.TS_F_IIVAMAT if prev.TS_F_IIVAMAT!="None" else 0:0>10}" + "0"
            txt_str = txt_str + f"{prev.TS_F_TIVAMAT if prev.TS_F_TIVAMAT!="None" else 0:0>10}" 
            txt_str = txt_str + f"{prev.TS_F_TOTCOM if prev.TS_F_TOTCOM!="None" else 0:0>10}" 
            txt_str = txt_str + f"{prev.TS_F_IIVAVAR if prev.TS_F_IIVAVAR!="None" else 0:0>10}" + "0"  
            txt_str = txt_str + f"{prev.TS_F_TIVAVAR if prev.TS_F_TIVAVAR!="None" else 0:0>10}"
            txt_str = txt_str + f"{prev.TS_F_MANCAR if prev.TS_F_MANCAR!="None" else 0:0>10}"
            txt_str = txt_str + f"{prev.TS_F_IIVACAR if prev.TS_F_IIVACAR!="None" else 0:0>10}" + "0"  
            txt_str = txt_str + f"{prev.TS_F_TIVACAR if prev.TS_F_TIVACAR!="None" else 0:0>10}"  
            txt_str = txt_str + f"{prev.TS_F_MANMEC if prev.TS_F_MANMEC!="None" else 0:0>10}"
            txt_str = txt_str + f"{prev.TS_F_IIVAMEC if prev.TS_F_IIVAMEC!="None" else 0:0>10}" + "0"  
            txt_str = txt_str + f"{prev.TS_F_TIVAMEC if prev.TS_F_TIVAMEC!="None" else 0:0>10}" + "0" + " " * 10 
            txt_str = txt_str + "NON CONCORDATO0" + "  " + " " * 230    #"CV_F_AUTRIP len="230"
            txt_str = txt_str + " " * 8 + " " * 28 + "00000000C0" + "0000" + "00"  #"CV_F_DTPRCO" len="8" 
            txt_str = txt_str + f"{prev.TS_F_COSTOR if prev.TS_F_COSTOR!="None" else 0:0<5}" 
            #Con align="L" pad="0" un valore come 0.00 (4 char) viene riempito a destra → 0.000. Con riempimento a sinistra si ottiene 00.00. -->
            #"TS_F_COSTOR" pos="1924" len="5" align="R" pad="0" source="cell" address="//TS_F_COSTOR" format="dec2_dot"/>
            txt_str = txt_str + "0" + f"{prev.TS_F_TOTME if prev.TS_F_TOTME!="None" else 0:0>4}" + "00"
            txt_str = txt_str + f"{prev.TS_F_COSTO2 if prev.TS_F_COSTO2!="None" else 0:0>5}" + "C70011          0000.000000.00EURTE\n"
            f1.write(txt_str)
        f1.close()

        save_path_file = "EXPERIZ2_" + pfile.replace(".json","") + ".txt"
        f1 = open(save_path_file, "w")
        if prev.F_TIPPRE=="MECC":
            txt_str=""  
            txt_str = txt_str + '0' * 11  + "     "     
            txt_str = txt_str + format(prev.CV_F_TARGAV,">10") + " " * 8    
            txt_str = txt_str + format(prev.CV_F_TELAIO,"<17") + ' ' * 3 + "0" + " " * 20 + "0    000.00000.00SSSSSSS0         0         00  0\n"
            f1.write(txt_str)
        else:
            txt_str=""  
            txt_str = txt_str + format(prev.PIva_Riparatore,'0<11') + "     "     
            txt_str = txt_str + format(prev.CV_F_TARGAV,">10") + " " * 8    
            txt_str = txt_str + format(prev.CV_F_TELAIO,"<17") + ' ' * 3 + "0" + " " * 20 + "0    000.00000.00SSSSSSS0         0         00  0\n"
            f1.write(txt_str)
        f1.close()

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
            leggijson(file, filesno) #Sostituisco tutta la parte sotto con questa chiamata 
            '''prev1 = s.pratica.listaprev[0] 
            if prev1.F_TIPPRE == "Mecc":
                s.tipoPratica = "M"
            else:
                s.tipoPratica = "C"'''

    #fine ciclo file json 
    ScriviLog("File correttamente importati: " + str(filesno))
    tk.messagebox.showinfo("Processo completato.", "Processo di importazione completato")
    StopClock()   

    btnImporta["state"] =NORMAL
    btnAvvia["state"]   =DISABLED
    btnScelta["state"]  =DISABLED

def Import_Dati():
    res=mb.askquestion('Importazione su DB', 'Vuoi importare sul DB?')
    if res == 'no':
        tk.messagebox.showinfo("Nessun dato importato","Processo completato senza importazione.")
    else:
        connetti()
        if var1.get()==1:
            #Se non faccio nuove pratiche per forza (checkbox Nuove Pratiche SI) mi devo chiedere 
            #per ogni targa (di ogni preventivo) se devo cercare le pratiche/preventivi vecchi da associare
            for i in range(len(s.pratica.listaprev)):
                prevtmp=s.pratica.listaprev[i]
                listaoldprev=vediprev(prevtmp.CV_F_TARGAV)    
                if len(listaoldprev)>0:
                    #gestire la selezione della pratica/preventivo da collegare al preventivo attuale
                    selezionato=cerca()
                    if selezionato == "":
                        tk.messagebox.showinfo("Nessuna pratica selezionata","Processo in attesa della selezione della pratica.")
                        break
                    else:
                        numPratica=selezionato[0]
                        file=prevtmp.FILE
                        label.config(text='file=' + file + "; pratica=" + str(numPratica))
                        listaprevpra=vediPre(numPratica)
                        if len(listaprevpra)==0:
                            elaboraPratica(numPratica, 0, file, prevtmp.F_TIPPRE)
                        else:
                            prevSel = cercaPre()
                            if prevSel == "":
                                res = tk.messagebox.askquestion("Nessun preventivo selezionato. Inserisco nuovo?","Processo in attesa della selezione del preventivo.")
                                if res=="yes":
                                    elaboraPratica(numPratica, 0, file, prevtmp.F_TIPPRE)
                                else:
                                    break
                            else:
                                numPrev=prevSel[1]
                                label.config(text='file=' + file + "; pratica=" + str(numPratica)+ "/" + str(numPrev))

                                elaboraPratica(numPratica, numPrev, file, prevtmp.F_TIPPRE)
                                if var2==1:
                                    #cancello file json elaborato
                                    os.remove(file)  
                        s.pratica.listaprev.pop(i)          
                else:
                    res = tk.messagebox.askquestion("Nessun preventivo con questa targa. Inserisco nuova pratica","Ricerca completata")
                    if res=="yes":
                        importaNuovo(prevtmp)
        else:
            for i in range(len(s.pratica.listaprev)):
                prevtmp=s.pratica.listaprev[i]
                importaNuovo(prevtmp)
            #print(s.pratica.NumPratica)
    tk.messagebox.showinfo("Preventivi importati su WinCar","Processo completato con successo.")

def importaNuovo(prev):
    #continua:
    ####GC#### ClearCellsFromB12ToLastRow
    file=prev.FILE
    if prev.F_TIPPRE == "MECC":
        s.tipoPratica = "M"
    else:
        s.tipoPratica = "C"
    #### Se gli passo il numero di pratica lo avrei inserito in s.pratica.... AL MOMENTO NON IN USO!!!
    try:
        if s.pratica.NumPratica!=None:
            numPratica=0
        else:
            numPratica=s.pratica.NumPratica
    except:
        numPratica=0
    
    elaboraPratica(numPratica, 0, prev, s.tipoPratica)

    if var2==1:
        #cancello file json elaborato
        os.remove(file)             

def elaboraPratica(idPratica, idPreventivo, preventivo, tipoPratica): 
    ScriviLog ("inizio import")
    #MODIFICA DEL 29/01/2025
    #'controllo se è stato indicato un NUMERO PRATICA nel nome file
    if idPratica > 0:
        strSQL = "SELECT CARVEI.F_NUMPRA FROM CARVEI WHERE CARVEI.F_NUMPRA=" + str(idPratica) + ";"
        # Apri il recordset
        prat_cursor = conn.cursor()
        prat_cursor.execute(strSQL)  #creo un cursore/recordset(cursor) da una query
        rows = prat_cursor.fetchall() #recupera il risultato della query in execute e lo mette in una lista 
        #MyVal = int(rows["MaxDiF_NUMPRA"])
        
        if len(rows) > 0:
            ScriviLog("pratica trovata")
            #controllo quanti NUM PREVENTIVI ci sono per questa pratica
            #idPratica = idPratica
            strSQL = "SELECT TESPRE.ID_CODPRE, TESPRE.F_NUMPRE FROM TESPRE WHERE TESPRE.ID_CODPRE=" + str(idPratica) + " AND TESPRE.F_NUMPRE=" + str(idPreventivo) + ";"
            # Apri il recordset
            prev_cursor = conn.cursor()
            prev_cursor.execute(strSQL)  #creo un cursore/recordset(cursor) da una query
            rowspre = prev_cursor.fetchall() #recupera il risultato della query in execute e lo mette in una lista 
            if len(rowspre)==0:
                inserisciNuovoPreventivo_TesPre(idPratica, idPreventivo, preventivo, tipoPratica)
                inserisciNuovoPreventivo_RigPre(idPratica, idPreventivo, preventivo, tipoPratica)
                ScriviLog("Inserita nuovo preventivo su pratica N. " + str(idPratica))
                #se preventivo vuoto, allora CREO NUOVO PREVENTIVO
            else:
                #se più di 1, aggiorno preventivo
                pass
                # devo fare update su TESPRE/RIGPRE
        else:
            ScriviLog ("non ci sono pratiche")
            corrispondenzaTarga(idPratica, idPreventivo, preventivo, tipoPratica)
            #controllo la corrispondenza con la TARGA
            #se non c'è corrispondenza, creo nuova pratica e importo preventivo
            #se c'è corrispondenza, controllo quanti NUM PREVENTIVI  ci sono per questa pratica
            #se 1 preventivo vuoto, allora CREO NUOVO PREVENTIVO
            #se più di 1, faccio selezionare il preventivo da SOVRASCRIVERE oppure CREO NUOVO PREVENTIVO
    else:
        ScriviLog("pratica KO")
        if var1.get()==1:
            corrispondenzaTarga(idPratica, idPreventivo, preventivo, tipoPratica)
            #controllo la corrispondenza con la TARGA
            #se non c'è corrispondenza, creo nuova pratica e importo preventivo
            #se c'è corrispondenza, controllo quanti NUM PREVENTIVI  ci sono per questa pratica
            #se 1 preventivo vuoto, allora CREO NUOVO PREVENTIVO
            #se più di 1, faccio selezionare il preventivo da SOVRASCRIVERE oppure CREO NUOVO PREVENTIVO
        else:
            ScriviLog("Inizio inserimento pratica")
            idPratica=nuovaPratica(idPratica, preventivo, tipoPratica) 
            #NB per usare una sub e restituire un valore, usare ByRef, in questo caso mi serve il nuovo numero pratica
            #NON CAPISCO L'USO DI idPratica2 => sostituito con return in funzione nuovaPratica
            idPreventivo = 1
            #CICCOLONE    
            # idPratica = idPratica2
            inserisciNuovoPreventivo_TesPre(idPratica, idPreventivo, preventivo, tipoPratica)
            inserisciNuovoPreventivo_RigPre(idPratica, idPreventivo, preventivo, tipoPratica)
            ScriviLog("Inserita nuova pratica N. " + str(idPratica))
            feedback2 = "inserita nuova"
    termina(feedback2, idPratica)

def corrispondenzaTarga(idPratica, idPreventivo, preventivo, tipoPratica):
    prev0 = s.pratica.listaprev[0]
    targa = s.pratica.F_TARGAV.replace("Targa Veicolo ","")
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
            idPratica2=nuovaPratica(idPratica, preventivo, tipoPratica) #NB per usare una sub e restituire un valore, usare ByRef, in questo caso mi serve il nuovo numero pratica
            idPreventivo = 1
            idPratica = idPratica2
            inserisciNuovoPreventivo_TesPre(idPratica, idPreventivo, preventivo, tipoPratica)
            inserisciNuovoPreventivo_RigPre(idPratica, idPreventivo, preventivo, tipoPratica)
            ScriviLog("Inserita nuova pratica N. " + str(idPratica))
            feedback2 = "inserita nuova"
        else:
            #se non è vuoto Determina il numero di righe e colonne
            numColonne = len(rows[0])  # NON SERVE!!!!
            numRighe = len(rows)

            if numRighe == 1:
                controllaSingolaPratica(lFileName)
                ###SELEZIONO LA PRATICA DALLA LISTA DELLA FINESTRA
                selectedValue=cerca()
                if selectedValue > 0:
                    idPratica = rows[selectedValue].NumPratica  # assegno il numero pratica
                    cercaPre(idPratica, idPreventivo)
                else:    #se uguale a zero
                    if selectedValue == 0:
                        idPratica2=nuovaPratica(idPratica2, preventivo, tipoPratica) # NB per usare una sub e restituire un valore, usare ByRef, in questo caso mi serve il nuovo numero pratica
                        idPreventivo = 1
                        idPratica = idPratica2
                        inserisciNuovoPreventivo_TesPre(idPratica, idPreventivo, preventivo, tipoPratica)
                        inserisciNuovoPreventivo_RigPre(idPratica, idPreventivo, preventivo, tipoPratica)
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
                    idPratica2=nuovaPratica(idPratica2, preventivo, tipoPratica)
                    idPratica = idPratica2
                    idPratica2=nuovaPratica(idPratica2, preventivo, tipoPratica)
                    inserisciNuovoPreventivo_TesPre(idPratica, idPreventivo, preventivo, tipoPratica)
                    inserisciNuovoPreventivo_RigPre(idPratica, idPreventivo, preventivo, tipoPratica)
                    ScriviLog("Inserita nuova pratica N. " + str(idPratica))
                    feedback2 = "inserita nuova"
                else:                            #se scelgo una pratica, controllo i numeri preventivi
                    if selectedValue > 0:
                        idPratica = selectedValue
                        cercaPre(idPratica, idPreventivo)
                    else:
                        exit()
    else:
        nuovaPratica(idPratica2, preventivo, tipoPratica)

def nuovaPratica(idpratica, preventivo, tipoPratica):
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
    prev0 = preventivo
    if tipoPratica=="C":
        campi_carvei = """(f_numpra, f_targav, f_dataca, f_desmod, f_telaio, f_kimvei, f_datimm, 
            F_CODCLI, F_RAGSOC, F_VIACLI, F_CITTAC, F_CAPCLI, F_PROCLI, F_PARIVA, F_TELEFO, 
            f_tipove, f_tpreve, f_idmess, f_datcre, F___GUID, F_DESCOL)"""
        f_numpra = idPratica
        f_targav = prev0.CV_F_TARGAV          #                       arrDati(2, 8)                
        datatemp = datetime.datetime.fromtimestamp(int(prev0.CV_F_DATACA)/1000)
        f_dataca = datatemp.strftime("%d/%m/%Y")    #####datetime.datetime.fromtimestamp(datapr/1000)
        datatemp = datetime.datetime.fromtimestamp(int(prev0.CV_F_DATIMM)/1000)
        f_datcre = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")   #data e ora creazione pratica interna
        f_datimm = datatemp.strftime("%d/%m/%Y %H:%M:%S")
        f_desmod = prev0.CV_F_DESMOD    #                       Left(arrDati(2, 4), 70)
        f_telaio = prev0.CV_F_TELAIO
        f_kimvei = 0 
        f_tipove = prev0.CV_F_TIPOVE                    #                       arrDati(2, 6)
         #f_tipove  per il tipo vernice controllo le iniziali della descrizione
        if prev0.CV_F_TIPOVE[0:3] == "OPA":
            f_tipove = "O"
        if prev0.CV_F_TIPOVE[0:3] == "TRI":
            f_tipove = "T"
        if prev0.CV_F_TIPOVE[0:3] == "PER":
            f_tipove = "L"
        if prev0.CV_F_TIPOVE[0:3] == "MIC":
            f_tipove = "I"
        if prev0.CV_F_TIPOVE[0:3] == "MET":
            f_tipove = "M"
        if prev0.CV_F_TIPOVE[0:3] == "DOP":
            f_tipove = "O"
        if prev0.CV_F_TIPOVE[0:3] == "PAS":
            f_tipove = "P"
        #per i dati del cliente sono impostati su parametri.ini
        leggi_par_ini()
        f_tpreve = "C"   #tipo logo C per carrozzeria
        f_idmess = dataFormattata + str(idPratica)   #id mess
        F___GUID = prev0.ID_Riparazione[0:36]
        # codice colore vernice
        if prev0.CV_F_DESCOL != "":
            codcol = prev0.CV_F_DESCOL
            if codcol.index("(")>=0:
                StringaTraParentesi=codcol[codcol.index("(")+1: codcol.index(")")-1]
                fine=min(codcol.index("(")-1, 20)
                F_DESCOL = codcol[0: fine] #tronco stringa a 20 caratteri    DESCRIZIONE COLORE
            else:
                F_DESCOL = codcol
        valori_carvei = "(" + str(f_numpra) + ",'" + f_targav + "','" + str(f_dataca) + "','" + f_desmod 
        valori_carvei = valori_carvei + "','" + f_telaio + "','" + str(f_kimvei) + "','" + f_datimm + "', " 
        valori_carvei = valori_carvei + s.pratica.F_CODCLI + ",'" + s.pratica.F_RAGSOC + "','" + s.pratica.F_VIACLI + "','"
        valori_carvei = valori_carvei + s.pratica.F_CITTAC + "','" + s.pratica.F_CAPCLI + "','" + s.pratica.F_PROCLI + "','" + s.pratica.F_PARIVA
        valori_carvei = valori_carvei + "','" + s.pratica.F_TELEFO + "', " + "'" + f_tipove + "','" + f_tpreve + "','" + f_idmess
        valori_carvei = valori_carvei  + "','" + f_datcre + "','" + F___GUID + "', '" + F_DESCOL + "')"
        #fine case Carr 
    elif tipoPratica == "M":
        campi_carvei = "(f_numpra, f_targav, f_dataca, f_desmod, f_telaio, f_kimvei, f_datimm, "
        campi_carvei = campi_carvei + "F_CODCLI, F_RAGSOC, F_VIACLI, F_CITTAC, F_CAPCLI, F_PROCLI, F_PARIVA, F_TELEFO, "
        campi_carvei = campi_carvei + "f_nummot, f_tipove, f_tpreve, f_idmess, f_datcre, F___GUID)"""
         
        f_numpra = idPratica
        f_targav = prev0.CV_F_TARGAV.replace("Targa Veicolo ","")
        datatemp = datetime.datetime.fromtimestamp(int(prev0.CV_F_DATACA)/1000)
        f_dataca = datatemp.strftime("%d/%m/%Y")    #####datetime.datetime.fromtimestamp(datapr/1000)
        datatemp = datetime.datetime.fromtimestamp(int(prev0.CV_F_DATIMM)/1000)
        f_datcre = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")   #data e ora creazione pratica interna
        f_datimm = datatemp.strftime("%d/%m/%Y %H:%M:%S")
        f_desmod = prev0.CV_F_DESMOD.replace("'", "''")
        f_desmod = f_desmod[0:70]
        f_telaio = prev0.CV_F_TELAIO.replace("Telaio ","")
        f_kimvei = s.pratica.F_KIMVEI

        #modifica del 24/03/2025 per i clienti privati oltre ALD
        #per i dati del cliente sono impostati su parametri.ini
        ####NON SERVE
        # if prev0.Id_riparazione != "":    #se la colonna IdRip. contiene testo, è ALD
        #    leggi_par_ini()

        f_nummot = s.pratica.F_KIMVEI  #arrDati(2, 6)
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
        valori_carvei = valori_carvei + "'" + str(f_nummot) + "','" + f_tipove + "','" + f_tpreve
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
    if tipoPratica == "C":
        strSQL = "INSERT INTO Pratica2 (f_numpra, f_tippra, F_CODCOL) values (" + str(idPratica) + ", '1', '" + StringaTraParentesi + "')"
    elif tipoPratica == "M":
        strSQL = "INSERT INTO Pratica2 (f_numpra, f_tippra) values (" + str(idPratica) + ", '2')"

    ScriviLog("import.py - insert pratica2")
    # Esegui la query
    carvei_ins_cursor.execute(strSQL)
    #conn.commit()
    #fine insert pratica2
    return idPratica 
    
###################################################################################
def inserisciNuovoPreventivo_TesPre(idPratica, idPreventivo, prev, tipoPratica):
    #ID_CODPRE F_NUMPRE identificano preventivo in pratica su TESPRE 
    ScriviLog("Avvio - insert tespre - pratica n : " + str(idPratica))
    ID_CODPRE_tespre = idPratica        
    prevcurr=prev
    if prevcurr.F_TIPPRE=="CARR":
        # inizio parametrizzazione campi tabella TESPRE
        F_DATAPR = datetime.datetime.fromtimestamp(int(prevcurr.CV_F_DATACA)/1000).strftime("%d/%m/%Y") 
        #prevcurr.CV_F_DATACA        #data excel mecc
        if s.pratica.F_KIMVEI is None:
            F_KMVEI = 0
        else:
            F_KMVEI = s.pratica.F_KIMVEI
        F_SUPPLE = 15                   
        F_FINITU = 10                   
        if prevcurr.TS_F_TCOMPL == "":
            F_COMPLE = 0
        else:
            F_COMPLE = prevcurr.TS_F_TCOMPL      # TEMPO AGG VERNIC
        #F_MATAUT = -1                                               ' calcolo automatico mat.cons.
        # MATERIALI DI CONSUMO
        
        strmatcon = prevcurr.TS_F_MATCON.split(" ")
        if len(strmatcon)>1:
            F_MATCON = strmatcon[1]
        else:
            F_MATCON = strmatcon[0]
        if prevcurr.TS_F_TOTMAT == "":
            F_TOTMAT = 0
        else:
            F_TOTMAT = prevcurr.TS_F_TOTMAT     #Mat_consumo_iva     #IVA MATERIALI DI CONSUMO
        F_COSTOR = prevcurr.TS_F_MANCAR         #IMPORTO MANODOPERA ORARIA CARROZZERIA
        F_COSTO2 = prevcurr.TS_F_MANMEC         #IMPORTO MANODOPERA ORARIA MECCANICA
        F_MANCAR = prevcurr.TS_F_TIVACAR        #importo iva su manodopera carr.
        strF_MANMEC = prevcurr.TS_F_TIVAMEC
        F_TOTSR = prevcurr.TS_F_TOTSR
        F_TOTLA = prevcurr.TS_F_TOTLA
        F_TOTVE = prevcurr.TS_F_TOTVE
        F_TOTME = prevcurr.TS_F_TOTME
        F_TOTRIC = prevcurr.TS_F_TOTRIC
        F_IVARIC = 22           #% IVA su pezzi
        F_IVAMAN = 22           #% IVA su manodopera
        F_IVAMAT = 22           #% IVA su materiali
        F_IVAVAR = 22           #% IVA su varie
        F_CIVRIC = 22
        F_CIVMAN = 22
        F_CIVMAT = 22
        F_CIVVAR = 22
        F_IIVARIC = prevcurr.TS_F_IIVARIC.replace(",",".")         #Imposta su pezzi
        F_IIVACAR = prevcurr.TS_F_IIVACAR.replace(",",".")         #Imposta su manodopera carrozzeria
        F_IIVAMEC = prevcurr.TS_F_IIVAMEC.replace(",",".")         #Imposta su manodopera meccanica
        F_IIVAMAT = prevcurr.TS_F_IIVAMAT.replace(",",".")         #Imposta su materiali
        F_IIVAVAR = prevcurr.TS_F_IIVAVAR.replace(",",".")         #Imposta su varie
        F_TIVARIC = prevcurr.TS_F_TIVARIC.replace(",",".")         #Totale (IVA compresa) pezzi
        F_TIVACAR = prevcurr.TS_F_TIVACAR.replace(",",".")         #Totale (IVA compresa) manodopera carrozzeria
        F_TIVAMEC = prevcurr.TS_F_TIVAMEC.replace(",",".")         #Totale (IVA compresa) manodopera meccanica
        F_TIVAMAT = prevcurr.TS_F_TIVAMAT.replace(",",".")         #Totale (IVA compresa) materiali
        F_TIVAVAR = prevcurr.TS_F_TIVAVAR.replace(",",".")         #Totale (IVA compresa) varie
        F_TOTPRE = prevcurr.TS_F_TOTPRE.replace(",",".")           #Totale preventivo IVA escl.
        F_TOTIVA = prevcurr.TS_F_TOTIVA.replace(",",".")           #Totale imposta
        F_TOTALE = prevcurr.TS_F_TOTALE.replace(",",".")           #Totale preventivo IVA incl.
        F_TFINIT = prevcurr.TS_F_TIFINIT.replace(",",".")          #Tempo per la finitura
        F_TSUPPL = prevcurr.TS_F_TSUPPL.replace(",",".")           #Tempo per il supplemento
        F_TCOMPL = prevcurr.TS_F_TCOMPL.replace(",",".")           #Tempo per il completamento
        try:
            F_VEOPER = prevcurr.TS_F_VEOPER             #Tempo VE operativo
        except:
            F_VEOPER = 0
        F_TEMSUP = prevcurr.TS_F_TEMSUP                 #Totale tempi supplementari
        F_VALUTA_tespre = "Euro"           
        F_NUMPRE_tespre = prevcurr.numprev              #Numero Preventivo
        F_SMAMAX = 0
        F_KMVEI = '0'                                   #IMPORTO MAX APPLICABILE SMALTIM. RIFIUTI    
        #Verifica se il valore è vuoto o nullo
        if prevcurr.TS_F_TOTRIC is None or prevcurr.TS_F_TOTRIC == "":
            strF_RICNET = 0
        else:
            # Converti il valore in decimale
            strF_RICNET = prevcurr.TS_F_TOTRIC
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
            valori_tespre = "(" + str(ID_CODPRE_tespre) + ", '" + str(F_DATAPR) + "', " + F_KMVEI + ",  " + str(F_SUPPLE) + ", " + str(F_FINITU) + ", " + F_COMPLE + ", " 
            valori_tespre = valori_tespre + F_MATCON + ", " + F_TOTMAT + ", " + F_COSTOR + ", " + F_COSTO2 + ", " + F_MANCAR + ", " + strF_MANMEC + ", "
            valori_tespre = valori_tespre + F_TOTSR + ", " + F_TOTLA + ", " + F_TOTVE + ", " + F_TOTME + ", " + F_TOTRIC + ", " + F_IIVARIC + ", " + F_IIVACAR + "," 
            valori_tespre = valori_tespre + F_IIVAMEC + ", " + F_IIVAMAT + ", " + F_IIVAVAR + ", " + F_TIVARIC + ", " + F_TIVACAR + ", " + F_TIVAMEC + ", " + F_TIVAMAT + ", " 
            valori_tespre = valori_tespre + F_TIVAVAR + ", " + F_TOTPRE + ", " + F_TOTIVA + ", " + F_TOTALE + ", " + F_TFINIT + ", " + F_TSUPPL + ", " + F_TCOMPL + ", " 
            valori_tespre = valori_tespre + str(F_VEOPER) + ", " + F_TEMSUP + ", '" + F_VALUTA_tespre + "', " + str(F_NUMPRE_tespre) + ", " + str(F_SMAMAX) + ", " + strF_RICNET + ", " 
            valori_tespre = valori_tespre + str(F_IVARIC) + ", " + str(F_IVAMAN) + ", " + str(F_IVAMAT) + ", " + str(F_IVAVAR) + ", " + str(F_CIVRIC) + ", " + str(F_CIVMAN) + ", " + str(F_CIVMAT) + ", " + str(F_CIVVAR) + ")"
        #fine case C
        #MECCANICA
    elif prevcurr.F_TIPPRE=="MECC":
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
        F_DATAPR = datetime.datetime.fromtimestamp(int(prevcurr.CV_F_DATACA)/1000).strftime("%d/%m/%Y")
        if prevcurr.CV_F_KIMVEI is not None:
            F_KMVEI = 0
        else:
            F_KMVEI = prevcurr.CV_F_KIMVEI
        F_SUPPLE = 15           
        F_FINITU = 10           
        F_COMPLE = 1.6   
        F_MATCON = prevcurr.TS_F_TOTMAT[0:2]             #IMPORTO MATERIALI CONSUMO
        F_TOTMAT = prevcurr.TS_F_TOTMAT.replace(",",".") #tot. materiali di consumo mecc
        F_COSTO2 = prevcurr.TS_F_COSTO2.replace(",",".") #tariffa manodopera mecc
        strF_MANMEC = (ReplaceApostrofo(prevcurr.TS_F_MANMEC))#
        F_TOTME = prevcurr.TS_F_TOTME.replace(",",".")   #tot ore manodopera mecc
        F_TOTRIC = prevcurr.TS_F_TOTRIC.replace(",",".") #importo netto ricambi
        F_IVARIC = '22'                                  #% IVA su pezzi
        F_IVAMAN = '22'                                  #% IVA su manodopera
        F_IVAMAT = '22'                                  #% IVA su materiali
        F_IVAVAR = '22'                                  #% IVA su varie
        F_TOTPRE = prevcurr.TS_F_TOTPRE.replace(",",".") #Totale preventivo IVA escl.
        F_TOTIVA =  prevcurr.TS_F_TOTIVA.replace(",",".")#Totale imposta
        F_TOTALE =  prevcurr.TS_F_TOTALE.replace(",",".")#Totale preventivo IVA incl.
        F_SCORIC = prevcurr.TS_F_SCORIC.replace(",",".") #% Sconto riservato sui ricambi
        F_SCOMAN = prevcurr.TS_F_SCOMAN.replace(",",".") #% Sconto riservato sulla manodopera
        F_SCOVAR = prevcurr.TS_F_SCOVAR.replace(",",".") #% Sconto riservato sulle varie
        F_CALCOL = '-1'
        F_PERRIF = '0'                                   #% Smaltimento rifiuti da calcolare su manodopera VE + materiali di consumo IMPOSTATA A ZERO PERCHè INSERITA MANUALMENTE arrDati(2, 30)
        F_IMPRIF = '0'                                   #Importo derivato da  manodopera VE + materiali di consumo per calcolo smalt.rif.
        F_VALUTA_tespre = "Euro"           
        F_NUMPRE_tespre = prevcurr.numprev               #Numero Preventivo
        F_FTSABA = '-1'         
        F_FTDOME = '-1'         
        F_CIVRIC = '22'         
        F_CIVMAN = '22'         
        F_CIVMAT = '22'         
        F_CIVVAR = '22'         
        F_SR_RIC = '-1'         
        F_SR_DIM = '-1'         
        F_SR_MAT = '-1'         
        F_SR_TSR = '-1'         
        F_SR_TLA = '-1'         
        F_SR_TVE = '-1'         
        F_SR_TME = '-1'         
        F_IMPRIC = '0'         
        F_TEMPAR = '99'         
        F_ESERIC = prevcurr.TS_F_ESERIC.replace(",",".")       #ci vanno gli importi esenti iva se presenti
        strF_RICNET = (ReplaceApostrofo(prevcurr.TS_F_TOTPRE)) #tot imponibile
        #fine parametrizzazione campi TESPRE
            
        #utilizzo una stringa di appoggio per i campi per semplificare la scrittura della query
        campi_tespre = "(ID_CODPRE, F_DATAPR, F_KMVEI, F_SUPPLE, F_FINITU, F_COMPLE, F_MATCON, F_TOTMAT, " 
        campi_tespre = campi_tespre + "F_COSTO2, F_MANMEC, F_TOTME, F_TOTRIC, F_IVARIC, F_IVAMAN, F_IVAMAT, F_IVAVAR, " 
        campi_tespre = campi_tespre + "F_TOTPRE, F_TOTIVA, F_TOTALE, F_SCORIC, F_SCOMAN, F_SCOVAR, F_CALCOL, F_PERRIF, F_IMPRIF, " 
        campi_tespre = campi_tespre + "F_VALUTA, F_NUMPRE, F_FTSABA, F_FTDOME, F_CIVRIC, F_CIVMAN, F_CIVMAT, F_CIVVAR, " 
        campi_tespre = campi_tespre + "F_SR_RIC, F_SR_DIM, F_SR_MAT, F_SR_TSR, F_SR_TLA, F_SR_TVE, F_SR_TME, F_IMPRIC, " 
        campi_tespre = campi_tespre + "F_TEMPAR, F_ESERIC, F_RICNET )"
    
        #utilizzo una stringa di appoggio per i valori da copiare per semplificare la scrittura della query
        valori_tespre = "(" + str(ID_CODPRE_tespre) + ", '" + str(F_DATAPR) + "', " + str(F_KMVEI) + ",  " + str(F_SUPPLE) + ", " + str(F_FINITU) + ", " + str(F_COMPLE) + ", " + F_MATCON + ", " + F_TOTMAT + ", " 
        valori_tespre = valori_tespre + "" + F_COSTO2 + ", " + strF_MANMEC + ", " + F_TOTME + ", " + F_TOTRIC + ", " + F_IVARIC + ", " + F_IVAMAN + ", " + F_IVAMAT + ", " + F_IVAVAR + ", " 
        valori_tespre = valori_tespre + "" + F_TOTPRE + ", " + F_TOTIVA + ", " + F_TOTALE + ", " + F_SCORIC + ", " + F_SCOMAN + ", " + F_SCOVAR + ", " + F_CALCOL + ", " + F_PERRIF + ", " + F_IMPRIF + ", " 
        valori_tespre = valori_tespre + "'" + F_VALUTA_tespre + "', " + str(F_NUMPRE_tespre) + ", " + F_FTSABA + ", " + F_FTDOME + ", " + F_CIVRIC + ", " + F_CIVMAN + ", " + F_CIVMAT + ", " + F_CIVVAR + ",  " 
        valori_tespre = valori_tespre + "" + F_SR_RIC + ", " + F_SR_DIM + ", " + F_SR_MAT + ", " + F_SR_TSR + ", " + F_SR_TLA + ", " + F_SR_TVE + ", " + F_SR_TME + ", " + F_IMPRIC + ", " 
        valori_tespre = valori_tespre + "" + F_TEMPAR + ", " + F_ESERIC + ", " + strF_RICNET + ")"
        #fine case M
    
    strSQL = "insert into TESPRE " + campi_tespre + " values " + valori_tespre + ";"
    # Esegui la query
    tespre_cursor = conn.cursor()
    #print(strSQL)
    tespre_cursor.execute(strSQL)  #creo un cursore/recordset(cursor) da una query
    ScriviLog("Fine - insert tespre - pratica n : " + str(idPratica))
    # fine insert TESPRE

def ReplaceApostrofo(str):
    return str.replace("'","")

def inserisciNuovoPreventivo_RigPre(idPratica, idPreventivo, prev, tipoPratica):
    ScriviLog("Avvio - insert rigpre - pratica n : " + str(idPratica))
    if prev.F_TIPPRE=='CARR':
        #CARROZZERIA
        ID_CODPRE = idPratica
        prevcurr=prev
        if len(prevcurr.listrighe)==0:
            ScriviLog("Preventivo vuoto.")
            ScriviLog("Fine Line3 - insert rigpre - error - pdf vuoto preventivo non compilato, n : " + idPratica)
            #GoTo prevVuoto #se non c'è intestagione Righe vuol dire che il pdf è vuoto, allora non compilo rig-pre
            return
        ScriviLog("Imposta valori.")
        F_DATRIG = datetime.datetime.fromtimestamp(int(prevcurr.CV_F_DATACA)/1000).strftime("%d/%m/%Y")
        for i in range(len(prevcurr.listrighe)):
            F_ORDINE = i + 1
            F_CITFON = prevcurr.listrighe[i].RG_F_CITFON
            #descrizione articolo.
            originalString = prevcurr.listrighe[i].RG_F_DESART
            modifiedString = prevcurr.listrighe[i].RG_F_DESART.replace("'", "''") # Removes all apostrophes
            F_DESART = modifiedString[0: 50]                 # descrizione articolo ha 50 caratteri max
            # quantità
            if prevcurr.listrighe[i].RG_F_QUANTI == 0 or prevcurr.listrighe[i].RG_F_QUANTI == "None" or prevcurr.listrighe[i].RG_F_QUANTI is None or prevcurr.listrighe[i].RG_F_QUANTI == "":
                F_QUANTI = 1   # QUANITA' zero imposto 1
            else:
                F_QUANTI = float(prevcurr.listrighe[i].RG_F_QUANTI)
            F_DANNSR = prevcurr.listrighe[i].RG_F_DANNSR
            if prevcurr.listrighe[i].RG_F_DANNSR == 'None':
                F_DANNSR = ''
            # h SR
            if prevcurr.listrighe[i].RG_F_TEMPSR == None or prevcurr.listrighe[i].RG_F_TEMPSR == 'None' or prevcurr.listrighe[i].RG_F_TEMPSR == "0":
                F_TEMPSR = 0
            else:
                if F_QUANTI > 1:
                    F_TEMPSR = prevcurr.listrighe[i].RG_F_TEMPSR / F_QUANTI
                else:
                    F_TEMPSR = prevcurr.listrighe[i].RG_F_TEMPSR
            F_DANNLA = prevcurr.listrighe[i].RG_F_DANNLA
            if prevcurr.listrighe[i].RG_F_DANNLA == 'None':
                F_DANNLA = ''
            # h LA
            if prevcurr.listrighe[i].RG_F_TEMPLA is None or prevcurr.listrighe[i].RG_F_TEMPLA == 'None' or prevcurr.listrighe[i].RG_F_TEMPLA == "0":
                F_TEMPLA = 0
            else:
                if F_QUANTI > 1:
                    F_TEMPLA = prevcurr.listrighe[i].RG_F_TEMPLA / F_QUANTI
                else:
                    F_TEMPLA = prevcurr.listrighe[i].RG_F_TEMPLA
            F_DANNVE = prevcurr.listrighe[i].RG_F_DANNVE
            if prevcurr.listrighe[i].RG_F_DANNVE == 'None':
                F_DANNVE = ''
            # h VE
            if prevcurr.listrighe[i].RG_F_TEMPVE == None or prevcurr.listrighe[i].RG_F_TEMPVE == 'None' or prevcurr.listrighe[i].RG_F_TEMPVE == "0":
                F_TEMPVE = 0
            else:
                if F_QUANTI > 1:
                    F_TEMPVE = prevcurr.listrighe[i].RG_F_TEMPVE / F_QUANTI
                else:
                    F_TEMPVE = prevcurr.listrighe[i].RG_F_TEMPVE
            # h ME
            if prevcurr.listrighe[i].RG_F_TEMPME == None or prevcurr.listrighe[i].RG_F_TEMPME == 'None' or prevcurr.listrighe[i].RG_F_TEMPME == "0":
                F_TEMPME = 0
            else:
                if F_QUANTI > 1:
                    F_TEMPME = prevcurr.listrighe[i].RG_F_TEMPME / F_QUANTI
                else:
                    F_TEMPME = prevcurr.listrighe[i].RG_F_TEMPME
            F_IDRIGO = i
            # prezzo
            if prevcurr.listrighe[i].RG_F_PREZZO == 'None' or prevcurr.listrighe[i].RG_F_PREZZO is None: 
                F_PREZZO = 0 
            else: 
                F_PREZZO = prevcurr.listrighe[i].RG_F_PREZZO
            if prevcurr.listrighe[i].RG_F_SCONTO == 'None' or prevcurr.listrighe[i].RG_F_SCONTO is None: 
                F_SCONTO = 0 
            else: 
                F_SCONTO = prevcurr.listrighe[i].RG_F_SCONTO 
            F_VALUTA = "Euro"
            F_NUMPRE = prevcurr.numprev
            F_CODGUI = prevcurr.ID_Riparazione[0: 2]
            F_QUANTI = F_QUANTI
            
            #stringa di appoggio per campi query rigpre
            campi_rigpre = """(ID_CODPRE, F_DATRIG, F_ORDINE, F_CITFON, F_DESART, F_QUANTI, 
                        F_DANNSR , F_TEMPSR, F_DANNLA, F_TEMPLA, F_DANNVE, F_TEMPVE, 
                        F_TEMPME , F_PREZZO, F_SCONTO, F_VALUTA, 
                        F_NUMPRE , F_IDRIGO, F_CODGUI )"""
                                
            #stringa di appoggo per valori query rigpre
            valori_rigpre = "('" + str(ID_CODPRE) + "', '" + str(F_DATRIG) + "', '" + str(F_ORDINE) + "', '" + F_CITFON + "', '" + F_DESART + "', '" + str(F_QUANTI) + "'," 
            valori_rigpre = valori_rigpre + "'" + F_DANNSR + "', " + str(F_TEMPSR) + ",  '" + F_DANNLA + "', " + str(F_TEMPLA)
            valori_rigpre = valori_rigpre + ", '" + F_DANNVE + "', " + str(F_TEMPVE) + ", "  + str(F_TEMPME) + ", " + str(F_PREZZO)
            valori_rigpre = valori_rigpre + ", " + str(F_SCONTO) + ", '" + F_VALUTA + "', " + str(F_NUMPRE) + ", '" + str(F_IDRIGO) + "', " 
            valori_rigpre = valori_rigpre + "'" + F_CODGUI + "');"
            #fine case C
            strSQL = "insert into RIGPRE " + campi_rigpre + " values " + valori_rigpre
            # Esegui la query
            rigpre_cursor = conn.cursor()
            #print(strSQL)
            rigpre_cursor.execute(strSQL)  #creo un cursore/recordset(cursor) da una query
    elif prev.F_TIPPRE=="MECC":  
        #MECCANICA
        # inizio parametrizzazione variabili
        ID_CODPRE = idPratica
        prevcurr=prev
        F_DATRIG = datetime.datetime.fromtimestamp(int(prevcurr.CV_F_DATACA)/1000).strftime("%d/%m/%Y") 
        for i in range(len(prevcurr.listrighe)):
            F_ORDINE = i + 1
            rigpre=prevcurr.listrighe[i]
            F_CITFON = rigpre.RG_F_CITFON   #COD ART PREVENTIVO MECCANICA
            if F_CITFON == 'None':
                F_CITFON = ''
            #descrizione articolo.
            originalString = rigpre.RG_F_DESART
            modifiedString = originalString.replace("'", " ") 
            F_DESART = modifiedString[0: 50] 
            if rigpre.RG_F_QUANTI == 0 or rigpre.RG_F_QUANTI is None or rigpre.RG_F_QUANTI == "None" or rigpre.RG_F_QUANTI == "":
                F_QUANTI = 1   # QUANTITA' zero imposto 1
            else:
                F_QUANTI = float(rigpre.RG_F_QUANTI)   # QUANTITA'
            F_DANNSR = "S"
            F_DANNLA = "S"          #
            F_DANNVE = "S"          #
            '''nei preventivi meccanica ALD le h di MDO inserite nelle righe sono totali,
            'mentre su WinCar vengono moltiplicate per le quantità
            'per ovviare a questo occorre dividere le ore manodopera per le quantità, ove le quantità sono maggiori di 1'''
            if rigpre.RG_F_TEMPME == "" or rigpre.RG_F_TEMPME == "0" or rigpre.RG_F_TEMPME == "None":
                F_TEMPME = 0
            else:
                if F_QUANTI > 1:
                    F_TEMPME = rigpre.RG_F_TEMPME / F_QUANTI
                else:
                    F_TEMPME = rigpre.RG_F_TEMPME
            if rigpre.RG_F_PREZZO is None or rigpre.RG_F_PREZZO == "None": 
                F_PREZZO = 0 
            else: 
                F_PREZZO = rigpre.RG_F_PREZZO
            F_FLAGPR = ''
            if rigpre.RG_F_SCONTO is None or rigpre.RG_F_SCONTO == "None": 
                F_SCONTO = 0 
            else: 
                F_SCONTO = rigpre.RG_F_SCONTO
            F_TIPRIC = "S"         #S di sostituzione per preventivo meccanica
            try:
                if rigpre.F___TIPO is None: 
                    F___TIPO = ''
                else:
                    F___TIPO = rigpre.F___TIPO
            except:
                F___TIPO = ''
            F_VALUTA = "Euro"
            F_IDRIGO = i
            F_CODGUI = prevcurr.CV_F_NUMMOT
            F_NUMPRE = prevcurr.numprev
            F_CODIVA = 0
            #controllare se ci sono righe esenti iva, esempio "bolletino postale per revisioni"
            if originalString[0:10].upper() == "BOLLETTINO": 
                F_CODIVA = -1    #-1 è il codice per esente iva
            # fine parametrizzazione varibili
                            
            #stringa di appoggio per campi query rigpre
            campi_rigpre = " (ID_CODPRE, F_DATRIG, F_ORDINE, F_CITFON, F_DESART, F_QUANTI, " 
            campi_rigpre = campi_rigpre + "F_DANNSR, F_DANNLA, F_DANNVE, F_TEMPME, F_PREZZO, F_FLAGPR, " 
            campi_rigpre = campi_rigpre + "F_SCONTO, F_TIPRIC, F___TIPO, F_VALUTA, F_NUMPRE, F_IDRIGO, " 
            campi_rigpre = campi_rigpre + "F_CODGUI, F_CODIVA) "
                                
            #stringa di appoggo per valori query rigpre
            valori_rigpre = "(" + str(ID_CODPRE) + ", '" + str(F_DATRIG) + "', " + str(F_ORDINE) + ", '" + F_CITFON + "', '" + F_DESART + "', " + str(F_QUANTI) + ", " 
            valori_rigpre = valori_rigpre + "'" + F_DANNSR + "', '" + F_DANNLA + "', '" + F_DANNVE + "', " + str(F_TEMPME) + ", " + str(F_PREZZO) + ", '" + F_FLAGPR + "', " 
            valori_rigpre = valori_rigpre + "" + str(F_SCONTO) + ", '" + F_TIPRIC + "', '" + F___TIPO + "', '" + F_VALUTA + "', " + str(F_NUMPRE) + ", " + str(F_IDRIGO) + ", " 
            valori_rigpre = valori_rigpre + "'" + str(F_CODGUI) + "', " + str(F_CODIVA) + ")"
            #fine CASE M
            strSQL = "insert into RIGPRE " + campi_rigpre + " values " + valori_rigpre
            # Esegui la query
            rigpre_cursor = conn.cursor()
            rigpre_cursor.execute(strSQL)  #creo un cursore/recordset(cursor) da una query       
        #dopo la query di inserimento - inserisco una nuova riga su rigpre
        #controllo se il ciclo è alla fine
    if prev.F_TIPPRE == "MECC":
        F_ORDINE=F_ORDINE+1
        #INSERIMENTO RIGA SMALTIMENTO RIFIUTI
        if prevcurr.Smalt_Rifiuti != "":
            strSQL = "insert into RIGPRE (ID_CODPRE, F_DATRIG, F_DESART, F_QUANTI, F_PREZZO , F_ORDINE, F_NUMPRE, F_IDRIGO, F___TIPO) " 
            strSQL = strSQL + "values ('" + str(idPratica) + "', '" + str(F_DATRIG) + "', 'Smaltimento rifiuti', '1', '" + str(prevcurr.Smalt_Rifiuti) 
            strSQL = strSQL + "', '" + str(F_ORDINE) + "', '" + str(prevcurr.numprev) + "', '" + str((i + 1)) + "', 'VC')"
            smalt_cursor = conn.cursor()
            smalt_cursor.execute(strSQL) 
    elif prev.F_TIPPRE=="CARR":
        F_ORDINE=F_ORDINE+1
        #INSERIMENTO RIGA MATERIALI DI CONSUMO
        if prevcurr.TS_F_MATCON is not None or prevcurr.TS_F_MATCON != 0 or prevcurr.TS_F_MATCON != "":
            strSQL = "insert into RIGPRE (ID_CODPRE, F_DATRIG, F_DESART, F_QUANTI, F_PREZZO , F_ORDINE, F_NUMPRE, F_IDRIGO, F___TIPO) " 
            strSQL = strSQL + "values ('" + str(idPratica) + "', '" + str(F_DATRIG) + "', 'Materiali di uso e consumo', '1', '" + prevcurr.TS_F_MATCON 
            strSQL = strSQL + "', '" + str(F_ORDINE) + "', '" + str(prevcurr.numprev) + "', '" + str((i + 1)) + "', 'VC')"
            # Esegui la query
            matcon_cursor = conn.cursor()
            matcon_cursor.execute(strSQL) 
        #INSERIMENTO RIGA SMALTIMENTO RIFIUTI  ????????
        try:
            if prevcurr.Smalt_Rifiuti != "":
                F_ORDINE=F_ORDINE+1
                strSQL = "insert into RIGPRE (ID_CODPRE, F_DATRIG, F_DESART, F_QUANTI, F_PREZZO , F_ORDINE, F_NUMPRE, F_IDRIGO, F___TIPO) " 
                strSQL = strSQL + "values ('" + idPratica + "', '" + str(F_DATRIG) + "', 'Smaltimento rifiuti', '1', '" + prevcurr.Smalt_Rifiuti 
                strSQL = strSQL + "', '" + str(F_ORDINE) + "', '" + str(prevcurr.numprev) + "', '" + str(i + 1) + "', 'VC')"
                smalt_cursor = conn.cursor()
                smalt_cursor.execute(strSQL) 
        except:
            pass
    #FINE FOR
    ScriviLog("Fine - insert rigpre - pratica n : " + str(idPratica))
    conn.commit()
    #fine query import RIPRE   

def termina(feedback2, idPratica):
    pass
###################################################################################

def controllaSingolaPratica(file):
    pass
###################################DA IMPLEMENTARE###################################
###################################DA IMPLEMENTARE###################################
###################################DA IMPLEMENTARE###################################
###################################DA IMPLEMENTARE###################################
###################################DA IMPLEMENTARE###################################

#Cerca le pratiche di una targa 
def vediprev(targa):  
    #connetti()  
    my_cursor = conn.cursor()
    strSQL = "SELECT CARVEI.F_NUMPRA, CARVEI.F_DATACA, CARVEI.F_TARGAV, CARVEI.F_RAGSOC, CARVEI.F_IMPPRE, " \
        "CARVEI.F_TPREVE FROM CARVEI WHERE (((CARVEI.F_TARGAV) like '" + targa + "') AND ((CARVEI.F_CHIUS2)<>80)) "\
        "ORDER BY CARVEI.F_NUMPRA DESC ,CARVEI.F_DATACA DESC;"
        # Esegui la query
    my_cursor.execute(strSQL)  #creo un cursore/recordset(cursor) da una query
    rows = my_cursor.fetchall() #recupera il risultato della query in execute e lo mette in una lista 
    #print("Num records: ", len(rows))
    i=0
    #Lb1.delete(0,Lb1.size()-1)
    tv['columns'] = ('IDPrev', 'DataPr', 'RagSoc', 'Totale')
    tv.heading("#0", text='Preventivo', anchor='w')
    tv.column("#0", anchor="w", width=10)
    tv.heading('IDPrev', text='ID Prev.')
    tv.column('IDPrev', anchor='center', width=20)
    tv.heading('DataPr', text='Data')
    tv.column('DataPr', anchor='center', width=80)
    tv.heading('RagSoc', text='Rag. Sociale')
    tv.column('RagSoc', anchor='center', width=150)
    tv.heading('RagSoc', text='Targa')
    tv.column('RagSoc', anchor='center', width=80)
    tv.heading('Totale', text='Totale')
    tv.column('Totale', anchor='center', width=100)
    for row in rows:
        #print(row.ID_CODPRE, row.F_DATAPR)
        i=i+1
        tv.insert('', 'end', values=(row.F_NUMPRA, row.F_DATACA, row.F_RAGSOC, row.F_TARGAV, row.F_IMPPRE))
                  #, f"€ {int(row.F_TOTRIC):.2f}"))
    return rows

def cerca():
    l=""
    for i in tv.selection():
        l=tv.item(i)["values"]
    return l

#Cerca i preventivi collegati alla pratica
def vediPre(idPratica):
    my_cursor = conn.cursor()
    strSQL = "SELECT TESPRE.ID_CODPRE, TESPRE.F_NUMPRE FROM TESPRE WHERE TESPRE.ID_CODPRE=" + str(idPratica) + " ORDER BY TESPRE.ID_CODPRE DESC;"
        # Esegui la query
    my_cursor.execute(strSQL)  #creo un cursore/recordset(cursor) da una query
    rows = my_cursor.fetchall() #recupera il risultato della query in execute e lo mette in una lista 
    #print("Num records: ", len(rows))
    #inserisco una tabella 
    tvpre['columns'] = ('IDPrev', 'NPrev')
    tvpre.heading("#0", text='Preventivo', anchor='w')
    tvpre.column("#0", anchor="w", width=10)
    tvpre.heading('IDPrev', text='ID Prev.')
    tvpre.column('IDPrev', anchor='center', width=20)
    tvpre.heading('NPrev', text='N. Prev.')
    tvpre.column('NPrev', anchor='center', width=20)
    tvpre.grid_rowconfigure(0, weight = 1)
    tvpre.grid_columnconfigure(0, weight = 1)
    tvpre.grid(row=5, column=2, columnspan=1, sticky="W", padx=10, pady=10)
    i=0

    for row in rows:
        #print(row.ID_CODPRE, row.F_DATAPR)
        i=i+1
        tvpre.insert('', 'end', values=(row.ID_CODPRE, row.F_NUMPRA))   
    return rows

def cercaPre():
    l=""
    for i in tvpre.selection():
        #print(tv.item(i).values())
        l=tvpre.item(i)["values"]
        #print(l)
    return l

def print_selection():
    # Create a vertical scrollbar
    #v_scrollbar = tk.ttk.Scrollbar(window, orient=tk.VERTICAL, command=tv.yview)
    if (var1.get() == 0):
        label.config(text='Solo Nuove pratiche ')
        tv.grid_forget()
        btnScelta.grid_forget()
        #v_scrollbar.grid_forget()
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
window.geometry("1000x1000")
window.title("THS Interfaccia per WinCar")
window.resizable(False, False)

Font_tuple = ("Calibri", 18, "bold")
Font_tab = ("Calibri", 14, "normal")

#Pulsante per avvio procedura e per selezionare pratica esistente
btnAvvia=tk.Button(text="Leggi files", command=Import_JSON, font=Font_tuple, fg="yellow", bg="blue")
btnImporta=tk.Button(text="Importa su WinCar", command=Import_Dati, font=Font_tuple, fg="yellow", bg="blue", state="disabled")
btnScelta=tk.Button(text="Scegli", command=cercaPre, font=Font_tuple, fg="yellow", bg="blue") 
        #command definisce il metodo da chiamare alla pressione del tasto
#inserisco una tabella 
tv = Treeview(window)
tvpre = Treeview(window)
tv.grid_rowconfigure(0, weight = 1)
tv.grid_columnconfigure(0, weight = 1)
label = tk.Label(window, bg='white', width=20, text='Solo Nuove pratiche')
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
