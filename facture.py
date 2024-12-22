from docxtpl import DocxTemplate
import tkinter
from tkinter import ttk
import openpyxl
import ttkbootstrap as ttk
from tkinter import *

entry1 = None
raison_sociale = None
adresse = None
client_data = None
entry2 = None
entry3= None
window2 = None
totalprices = []
totale = None
prix = None
totremise = None
remise = None
percremise = None
Tva = None
Tvv = None
topay =None
payy = None
treeview = None

doc = DocxTemplate("C:/Users/Hamza/Desktop/python tkinter/invoice_template.docx")




def find_data(sheet_name, value):
    path = "C:/Users/Hamza/Desktop/python tkinter/tk.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook[sheet_name]

    for row in sheet.iter_rows(min_row=2, values_only=True):
        if str(row[0]) == str(value):
            return row[1], row[2]

    return None, None



def get_data():
    global entry1, raison_sociale, adresse
    data = entry1.get()
    try:
        data = int(data)  # Convert the input value to an integer
        raison_sociale, adresse = find_data('client', data)
        if raison_sociale and adresse:
            print("Raison sociale:", raison_sociale)
            print("Adresse:", adresse)
        else:
            print("No data found for the given value.")
    except ValueError:
        print("Invalid input. Please enter a valid integer.")
    
    update_label()



def facturation():
    global entry1, raison_sociale, adresse, client_data, entry2, entry3, treeview, prix, totremise,  remise, percremise, Tvv, payy, doc
    window2 = ttk.Window(themename='darkly')
    #window2 = tkinter.Tk()
    window2.title("Facturation")

    frame2 = ttk.Frame(master=window2)
    frame2.pack(padx=20, pady=10)

    numberclient = ttk.Label(frame2, text="Entrer le numero du client")
    numberclient.grid(row=0, column=0, padx=10, pady=5, sticky="w")

    entry1 = ttk.Entry(frame2)
    entry1.insert(0, "Numero client")
    entry1.bind("<FocusIn>", lambda e: entry1.delete('0', 'end'))
    entry1.grid(row=1, column=0, padx=10, pady=5, sticky="w")

    inser_bt = ttk.Button(master=frame2, text='Ajouter le client', command=get_data)
    inser_bt.grid(row=2, column=0, padx=10, pady=5)

    numberproduit = ttk.Label(frame2, text="Ajout d'un produit:")
    numberproduit.grid(row=3, column=0, padx=10, pady=5, sticky="w")

    entry2 = ttk.Entry(frame2)
    entry2.insert(0, "Numero produit")
    entry2.bind("<FocusIn>", lambda b: entry2.delete('0', 'end'))
    entry2.grid(row=4, column=0, padx=10, pady=5, sticky="w")

    inser_bt1 = ttk.Button(master=frame2, text='Ajouter le produit', command=add_table)
    inser_bt1.grid(row=5, column=0, padx=10, pady=5)

    entry3 = ttk.Entry(frame2)
    entry3.insert(0, "La QTE")
    entry3.bind("<FocusIn>", lambda b: entry3.delete('0', 'end'))
    entry3.grid(row=4, column=1, padx=10, pady=5, sticky="w")

    client_data = ttk.Label(frame2, text='')
    client_data.grid(row=1, column=1, padx=10, pady=5, sticky="w")

    headers = ["Number", "Description", "Qte", "Prix UNIT", "Total"]
    treeview = ttk.Treeview(frame2, columns=headers, show="headings")

    treeview.column("Number", width=100)
    treeview.column("Description", width=200)
    treeview.column("Qte", width=100)
    treeview.column("Prix UNIT", width=100)
    treeview.column("Total", width=100)

    for header in headers:
        treeview.heading(header, text=header)

    treeview.grid(row=6, column=0, columnspan=2, padx=10, pady=10)

    prix = ttk.Label(frame2, text= f'Le Total HT : {totale}')
    prix.grid(row=7, column=1, padx=10, pady=5, sticky="w")

    remise = ttk.Label(frame2, text= f'pourcentage de la remise est: {percremise}, avec un totale de :{totremise}')
    remise.grid(row=8, column=1, padx=10, pady=5, sticky="w")

    Tvv = ttk.Label(frame2, text= f'TVA a 19,6 % : {Tva}')
    Tvv.grid(row=9, column=1, padx=10, pady=5, sticky="w")

    payy = ttk.Label(frame2, text= f'Montant TTC a payee est : {topay}')
    payy.grid(row=10, column=1, padx=10, pady=5, sticky="w")

    def generate_facture(treeview):
        global doc, entry1, raison_sociale, adresse, client_data, entry2, entry3, prix, totremise, remise, percremise, Tvv, payy

        table_rows = []
        if treeview:
            for item in treeview.get_children():
                values = treeview.item(item)["values"]
                table_rows.append(values)
        
        context = {
            "invoice_list": table_rows,
            "name": raison_sociale,
            "phone": adresse,
            "total": totale,
            "remise_percentage": percremise,
            "remise_total": totremise,
            "tva": Tva,
            "pay_amount": topay
        }

        doc.render(context)
        doc.save("C:/Users/Hamza/Desktop/python tkinter/facture.docx")

    inser_bt3 = ttk.Button(master=frame2, text='Generer la facture', command=lambda: generate_facture(treeview))
    inser_bt3.grid(row=11, column=1, padx=10, pady=5)

    window2.mainloop()


def add_table():
    global entry2, entry3, treeview, totalprices, totale, totremise, percremise, Tva, topay
    number = entry2.get()
    qte = entry3.get()
    description, price = find_data('catalogue', number)
    if description and price:
        total =  int(qte) * int(price)
        totalprices.append(total)
        totale =  sum(totalprices)
        #totremise = totale * 2 / 100
        if totale <= 500 :
            percremise = '0%'
            totremise = 0
        elif totale >= 500 and totale <= 700:
            percremise = '8%'
            totremise = totale * 8 / 100
        elif totale >= 700 and totale <= 900:
            percremise = '15%'
            totremise = totale * 15 / 100
        elif totale <= 900:
            percremise = '20%'
            totremise = totale * 20 / 100

        Tva = round(float(totale) * 19.6 / 100, 2)
        topay = totale + Tva - totremise
        treeview.insert("", "end", values=(number, description, qte, price,total))
        update_label()
    else:
        print("No data found for the given product number.")
    entry2.delete(0, 'end')
    entry3.delete(0, 'end')

def update_label():
        global entry1, raison_sociale, adresse, client_data, entry2, prix, remise, Tvv, topay
        client_data.config(text=f'Raison sociale: {raison_sociale}\nAdresse: {adresse}')
        prix.config(text=f'Le Total HT : {totale}')
        remise.config(text=f'pourcentage de la remise est : {percremise}, avec un totale de : {totremise}')
        Tvv.config(text=f'TVA a 19,6 % : {Tva}')
        payy.config(text=f'Montant TTC a payee est : {topay}')




facturation()
print(Tva)
print(totalprices)
print(totale)

#print(find_data('client', 3))
