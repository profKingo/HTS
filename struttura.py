s_header_mecc=["Totale_Ricambi"
,"TS_F_TOTPRE"
,"TS_F_IVARIC_IVAMAN_IVAMAT_IVAVAR"
,"TS_F_SCORIC"
,"TS_F_TOTALE"
,"TS_F_TOTRIC"
,"Totale_Varie"
,"TS_F_TOTIVA"
,"TS_F_SCOVAR"
,"TS_F_TOTCOM"
,"TS_F_TOTME"
,"TS_F_COSTO2"
,"TS_F_SCOMAN"
,"TS_F_MANMEC"
,"TS_F_TOTMAT"
,"Smalt_Rifiuti"
,"TS_F_ESERIC"
,"CV_F_RAGSOC"
,"CV_F_PARIVA"
,"CV_F_DATACA"
,"CV_F_DESMOD"
,"CV_F_TARGAV"
,"CV_F_TELAIO"
,"CV_F_KIMVEI"
,"CV_F_DATIMM"
,"CV_F_TIPPRE"
,"CV_F_NUMMOT"
,"Data Stampa"]    
    
 #   "F_RAGSOC","F_TOTPRE","Costo Tot. Ricambi","Data preventivo","Costo Tot. Varie",
 #   "Totale Imponibile","Piva Riparatore","Numero ore lavorate","Totale Iva","F_TIPPRE",
 #   "Tariffa Manodopera","Costo Tot. MDO","Descrizione Veicolo","Targa Veicolo",
 #   "Telaio","Km","Data Immatricolazione","Cod.Motore"


s_header_carr=["CV_F_DATACA"
,"TS_F_TOTPRE"
,"TS_F_TOTSR"
,"TS_F_TOTRIC"
,"TS_F_TOTSR_F_TOTLA_F_TOTVE"
,"TS_F_TOTME"
,"TS_F_TOTCOM"
,"TS_F_MATCON"
,"ID_Riparazione"
,"CV_F_RAGSOC"
,"TS_F_TOTIVA"
,"TS_F_IIVAVAR"
,"check_SR_rigpre"
,"check_tariffa_mat_con"
,"check_imponibile_ric"
,"check_tot_h_carr"
,"Ore Mano D'Opera Mecc Check"
,"TS_F_TOTALE"
,"TS_F_TOTLA"
,"TS_F_IIVARIC"
,"TS_F_TOTMAT"
,"TS_F_COSTOR"
,"TS_F_COSTO2"
,"TS_F_TIVAVAR"
,"TS_F_TIVARIC"
,"check_LA_rigpre"
,"check_imponibile_mat_cons"
,"check_tariffa_man_carr"
,"check_tariffa_manodop_mecc"
,"VE" #capire che campo è in TESPRE
,"TS_F_IIVAMAT"
,"TS_F_MANCAR"
,"TS_F_MANMEC"
,"check_tot_ric"
,"TS_F_TIVAMAT"
,"chek_VE_rigpre"
,"check_imponibile_man_carr"
,"check_imponibile_manodop_mecc"
,"TS_F_IIVACAR"
,"TS_F_IIVAMEC"
,"check_tot_mat_cons"
,"TS_F_TIVACAR"
,"TS_F_TIVAMEC"
,"check_ME_rigpre"
,"TS_F_TSUPPL"
,"check_tot_manodop_mecc"
,"TS_F_TIFINIT"
,"TS_F_TCOMPL"
,"TS_F_TEMSUP"
,"check_temp_suppl"
,"TS_F_TOTVE"
,"check_tot_VE"
,"Ragione_Sociale_Riparatore"
,"CV_F_TARGAV"
,"CV_F_TELAIO"
,"CV_F_KIMVEI"
,"CV_F_DATIMM"
,"CV_F_PREANT"
,"PIva_Riparatore"
,"CV_F_DESMOD"
,"CV_F_DESCOL"
,"CV_F_TIPOVE"
,"TS_F___FTGG"]

s_obmecc=["RG_F_QUANTI","RG_F_CITFON","RG_F_DESART","RG_F_TEMPME","RG_F_PREZZO","RG_F_SCONTO","Totale_Riga"]

s_obcarr=["RG_F_QUANTI","RG_F_CITFON","RG_F_DESART","RG_F_DANNSR","RG_F_TEMPSR","RG_F_DANNLA","RG_F_TEMPLA","RG_F_DANNVE","RG_F_TEMPVE","RG_F_TEMPME","RG_F_PREZZO","RG_F_SCONTO","Netto_Riga"]
#  "Quantita","Codice","Descrizione","SR Dif.","SR Tempo","LA Dif.","LA Tempo","VE Dif.","VE Tempo","ME Tempo","Prezzo Unitario", "D.S.M.C.", "Ammontare"]

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
    F_TELAIO = ""
    F_TARGAV = 0
    F_DATACA = 0
    F_DATIMM = 0
    F_KIMVEI = ""
    F_DESCOL = ""
    F_DESMOD = ""
    F_TIPOVE = ""
    F_NUMMOT = ""
    def __init__(self):
        pass
    def addprev(self, prev):
        self.listaprev.append(prev)
class prev:
    listrighe=[]
    #Targa_Veicolo=""
    numprev=0
    def __init__(self):
        listrighe=[]
    def addriga(self, riga):
        self.listrighe.append(riga)
class riga:
    def __init__(self):
        pass

class importazione:
    listprat=[]
    #Targa_Veicolo=""
    numpra=0
    def __init__(self):
        listprat=[]
    def addpra(self, pra):
        self.listprat.append(pra)
'''
nomedb="C:\\THS\\THS32Env\\wcArchivi.mdb"
filejson="export.json"
targetFolder=""
mylog=""
lockFilePath=""
tipoPratica="M" #Meccanica #Carrozzieria
'''