# -*- coding: utf-8 -*-
"""
Created on Wed Jul 23 17:39:13 2025

@author: User
"""

def create_picking_interface(tab, conn, cursor):
    def check_entry_length(event):
        num_cell = numero_serie_cell_entry.get()
        if len(num_cell) >= 12:
            numero_serie_cell_entry.unbind('<KeyRelease>')
            try:
                cursor.execute("SELECT affectation_produit FROM cellule WHERE numero_serie_cellule = %s", (num_cell,))
                produit = cursor.fetchall()[0][0]
                numero_serie_batt_entry.delete(0, tk.END)
                numero_serie_batt_entry.insert(0, str(produit))
            finally:
                numero_serie_cell_entry.bind('<KeyRelease>', check_entry_length)

    def etape_picking():
        num_batt = numero_serie_batt_entry.get()
        commentaire = text_box.get("1.0", "end-1c")
        valeur_control = ctrl_combobox.get()
        date = datetime.now()
        etat = 1 if valeur_control == 'Yes' else 0

        if commentaire:
            query = "UPDATE suivi_production set picking=%s, date_picking=%s,commentaire=%s where numero_serie_batterie=%s"
            param = (etat, date, commentaire, num_batt)
        else:
            query = "UPDATE suivi_production set picking=%s, date_picking=%s where numero_serie_batterie=%s"
            param = (etat, date, num_batt)

        cursor.execute(query, param)
        conn.commit()

        if etat == 1:
            messagebox.showinfo("Succès!", "Point de contrôle picking OK!")
        else:
            messagebox.showerror("Échec!", "Point de contrôle picking NON OK!")

        numero_serie_cell_entry.delete(0, tk.END)
        numero_serie_batt_entry.delete(0, tk.END)
        text_box.delete("1.0", tk.END)

    # Interface
    picking_frame = tk.Frame(tab, bg="#999999")
    picking_frame.pack(fill=tk.BOTH, expand=True)

    # === N° série cellule & batterie ===
    tk.Label(picking_frame, text="N° serie cellule", bg="#999999", fg="white", font=("Arial", 10)).place(x=50, y=40)
    numero_serie_cell_entry = tk.Entry(picking_frame, font=("Arial", 10))
    numero_serie_cell_entry.place(x=50, y=65)
    numero_serie_cell_entry.bind('<KeyRelease>', check_entry_length)

    tk.Label(picking_frame, text="N° serie batterie", bg="#999999", fg="white", font=("Arial", 10)).place(x=50, y=100)
    numero_serie_batt_entry = tk.Entry(picking_frame, font=("Arial", 10))
    numero_serie_batt_entry.place(x=50, y=125)

    # === Modèle batterie ===
    tk.Label(picking_frame, text="modele batterie", bg="#999999", fg="white", font=("Arial", 10)).place(x=50, y=160)
    modele_entry = tk.Entry(picking_frame, font=("Arial", 10))
    modele_entry.place(x=50, y=185)

    # === Liste batterie ===
    tk.Label(picking_frame, text="Liste batterie", bg="#999999", fg="white", font=("Arial", 10)).place(x=50, y=220)
    liste_batterie = tk.Listbox(picking_frame, height=8)
    liste_batterie.place(x=50, y=245, width=150)

    scrollbar = tk.Scrollbar(picking_frame, command=liste_batterie.yview)
    scrollbar.place(x=200, y=245, height=130)
    liste_batterie.config(yscrollcommand=scrollbar.set)

    # === Non conforme ===
    tk.Button(picking_frame, text="Non conforme", bg="red", fg="white", font=("Arial", 12)).place(x=50, y=400)

    # === Date ===
    tk.Label(picking_frame, text="Date", bg="#999999", fg="white", font=("Arial", 10)).place(x=350, y=40)
    date_entry = tk.Entry(picking_frame, font=("Arial", 10))
    date_entry.place(x=350, y=65)

    # === Remplacement cellules ===
    tk.Label(picking_frame, text="Remplacement cellules", bg="#999999", fg="white", font=("Arial", 10)).place(x=600, y=10)
    tk.Label(picking_frame, text="N° série cellule", bg="#999999", fg="white", font=("Arial", 10)).place(x=600, y=40)
    remplacement_cell_entry = tk.Entry(picking_frame, font=("Arial", 10))
    remplacement_cell_entry.place(x=600, y=65)

    # === Défaut (Combobox) ===
    tk.Label(picking_frame, text="Défaut", bg="#999999", fg="white", font=("Arial", 10)).place(x=600, y=100)
    defaut_combobox = ttk.Combobox(picking_frame, values=["Corrosion/tension etc"], font=("Arial", 10))
    defaut_combobox.place(x=600, y=125)

    # === Bouton Remplacement ===
    tk.Button(picking_frame, text="Remplacement", bg="red", fg="white", font=("Arial", 12)).place(x=600, y=170)

    # === Écart tension ===
    tk.Label(picking_frame, text="Ecart maximum de 0.05V", bg="lightgreen", font=("Arial", 14, "bold")).place(x=450, y=260)

    # === Contrôle OK ===
    tk.Button(picking_frame, text="Controle OK", bg="blue", fg="white", font=("Arial", 14, "bold"), width=15, command=etape_picking).place(x=500, y=330)

