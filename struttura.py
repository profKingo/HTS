s_header=["F_RAGSOC","F_TOTPRE","Costo Tot. Ricambi","Data preventivo","Costo Tot. Varie",
          "Totale Imponibile","Piva Riparatore","Numero ore lavorate","Totale Iva","F_TIPPRE",
          "Tariffa Manodopera","Costo Tot. MDO","Descrizione Veicolo","Targa Veicolo",
          "Telaio","Km","Data Immatricolazione","Cod.Motore"]

s_elem=["Quantita","Codice","Descrizione","Ore","Prezzo unitario","Sconto","Ammontare"]

class pratica:
    listaprev=[]
    F_CODCLI = 0
    F_RAGSOC = ""
    F_VIACLI = ""
    F_CITTAC = ""
    F_CAPCLI = ""
    F_PROCLI = ""
    F_PARIVA = ""
    F_TELEFO = ""
    NumPratica = 0
    def __init__(self):
        pass
    def addprev(self, prev):
        self.listaprev.append(prev)
class prev:
    listrighe=[]
    #Targa_Veicolo=""
    def __init__(self):
        listrighe=[]
    def addriga(self, riga):
        self.listrighe.append(riga)
class riga:
    def __init__(self):
        pass
'''
nomedb="C:\\THS\\THS32Env\\wcArchivi.mdb"
filejson="export.json"
targetFolder=""
mylog=""
lockFilePath=""
tipoPratica="M" #Meccanica #Carrozzieria
'''
