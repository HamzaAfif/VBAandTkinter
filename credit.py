import ttkbootstrap as ttk
from tkinter import *


def calcule():
    loan_amount = float(entry_loan.get())
    interest_rate = float(entry_interest.get()) / 100
    months = int(entry_months.get())

    monthly_interest = interest_rate / 12
    monthly_payment = (loan_amount * monthly_interest) / (1 - (1 + monthly_interest) ** -months)

    label_result.config(text=f"Mensualité: MAD{monthly_payment:.2f}")


win = ttk.Window(themename='darkly')

win.geometry("400x300")

label_loan = ttk.Label(win, text="Montant:")
label_loan.grid(row=0, column=0, padx=10, pady=10, sticky="e")

label_interest = ttk.Label(win, text="Taux (%):")
label_interest.grid(row=1, column=0, padx=10, pady=10, sticky="e")

label_months = ttk.Label(win, text="Duration (mois):")
label_months.grid(row=2, column=0, padx=10, pady=10, sticky="e")


entry_loan = ttk.Entry(win, width=20)
entry_loan.grid(row=0, column=1, padx=10, pady=10)


entry_interest = ttk.Entry(win, width=20)
entry_interest.grid(row=1, column=1, padx=10, pady=10)


entry_months = ttk.Entry(win, width=20)
entry_months.grid(row=2, column=1, padx=10, pady=10)


button_calculate = ttk.Button(win, text="Calculer", command=calcule)
button_calculate.grid(row=3, column=1, padx=10, pady=10, sticky="w")


label_result = ttk.Label(win, text="Mensualité: MAD 0.00", font=("Arial", 12, "bold"))
label_result.grid(row=4, column=0, columnspan=2, padx=10, pady=10)

win.mainloop()
