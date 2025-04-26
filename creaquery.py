import tkinter as tk
import json
import pyodbc
import struttura as s
from tkinter.ttk import *
from tkinter import *

'''
l'utente parte da file pdf da ocrizzare e quindi da inviare al portale di AIDA (tramite uno dei metodi tra cui WS)
i file pdf sono di tipo MECC o CARR e vengono nominati ad hoc...
pensare un batch che chiami appena generato il file la conversione e poi avvii la lettura del JSON risultante
il JSON viene rimandato all'utente come???? Webhook - FTP - AIDA LINK
'''
#stringa di connessione al db
connstr=f'Driver={{Microsoft Access Driver (*.mdb)}};Dbq=C:\\THS\\THS32Env\\wcArchivi.mdb;Uid=;Pwd=;'
conn=pyodbc.connect(connstr)
filejson="export.json"

def leggijson():
    data = json.load(open(filejson))
    if len(data)>=1:
        s.pratica.desc="Pratica"
        for pre in data:
            for x in s.s_header:    #INTESTAZIONE PREVENTIVO
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

    
    my_cursor = conn.cursor()
    my_cursor.execute("SELECT * FROM TESPRE")  #creo un cursore/recordset(cursor) da una query
    rows = my_cursor.fetchall() #recupera il risultato della query in execute e lo mette in una lista 
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
    print("Num records: ", len(rows))
    for row in rows:
        #print(row.ID_CODPRE, row.F_DATAPR)
        i=i+1
        tv.insert('', 'end', values=(row.ID_CODPRE, row.F_DATAPR, row.F_RAGSOC, f"€ {int(row.F_TOTRIC):.2f}"))
    
    #my_cursor.execute("insert into tabella(Nome, Cognome) values('spina','rosa')")

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



'''Lb1 = tk.Listbox(window)
Lb1.insert(1, "Python")
Lb1.insert(2, "Perl")
Lb1.insert(3, "C")
Lb1.insert(4, "PHP")
Lb1.insert(5, "JSP")
Lb1.insert(6, "Ruby")
Lb1.grid(row=1, column=0,sticky="WE", padx=10, pady=20)   '''

'''txt=tk.Entry()
txt.grid(row=1, column=0, sticky="WE")
gender = Combobox(
    window,
    values=["Male", "Female", "Other"],
    state="readonly",
)
gender.grid(row=2, column=1, padx=5, pady=5)'''
