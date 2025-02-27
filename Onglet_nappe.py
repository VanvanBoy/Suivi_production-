# -*- coding: utf-8 -*-
"""
Created on Tue Jan 21 15:37:09 2025

@author: User
"""

import mysql.connector 
import tkinter as tk 
from tkinter import ttk 
from tkinter import messagebox 
from tkcalendar import DateEntry 
import sys
from datetime import datetime 
import traceback


#Fonction de creation de l'interface dans l'ong tab, passage en argument de la connexion BDD et du curseur
def create_nappe_interface(tab,conn,cursor):
    
    def check_entry_length(event):
        num_cell=numero_serie_cell_entry.get()
        if len(num_cell) >= 12:
            numero_serie_cell_entry.unbind('<KeyRelease>')
            try:
                cursor.execute("SELECT affectation_produit FROM cellules WHERE numero_serie_cellule = %s", (num_cell,))
                produit = cursor.fetchall()[0][0]
                num_produit = str(produit)
                numero_serie_batt_entry.delete(0, tk.END)  # Effacer l'entrée avant d'insérer
                numero_serie_batt_entry.insert(0, num_produit)
            finally:
                # Réactiver l'événement
                numero_serie_cell_entry.bind('<KeyRelease>', check_entry_length)
      
        
    def etape_nappe():
        num_batt=numero_serie_batt_entry.get()
        commentaire=text_box.get("1.0", "end-1c")
        valeur_control=ctrl_combobox.get()
        tension=float(tension_entry.get())
        date=datetime.now()
        
        if valeur_control=='Yes' and tension <0.051:
            etat=1
        else:
            etat=0
           
        if commentaire:
            query="UPDATE suivi_production set soudure_nappe=%s, date_soudure_nappe=%s,commentaire=%s where numero_serie_batterie=%s"
            param=(etat,date,commentaire,num_batt)
        else : 
            query="UPDATE suivi_production set soudure_nappe=%s, date_soudure_nappe=%s where numero_serie_batterie=%s"
            param=(etat,date,num_batt)
        if etat==1:
            messagebox.showinfo("Succès!","Point de controle nappe OK! Passage à l'etape suivante")
        else : 
            messagebox.showerror("Fail!","Point de controle nappe NON OK!")
            
        cursor.execute(query,param)
        
        numero_serie_cell_entry.delete(0,tk.END)
        numero_serie_batt_entry.delete(0,tk.END)
        text_box.delete("1.0",tk.END)
        tension_entry.delete(0,tk.END)
        
        conn.commit()
        

    # Créer la fenêtre principale
    picking_frame = ttk.Frame(tab)

    # Créer les cadres pour les deux colonnes
    left_frame = tk.Frame(picking_frame)
    left_frame.pack(side=tk.LEFT, padx=10, pady=10)

    right_frame = tk.Frame(picking_frame)
    right_frame.pack(side=tk.RIGHT, padx=10, pady=10)

    # Labels et champs de saisie 
    numero_serie_cell_label = tk.Label(left_frame, text="Numéro de série d'une cellule du produit:", font=('Arial', 12))
    numero_serie_cell_label.pack(pady=5)
    numero_serie_cell_entry = tk.Entry(left_frame, font=('Arial', 12))
    numero_serie_cell_entry.pack(pady=5)
    numero_serie_cell_entry.bind('<KeyRelease>', check_entry_length)

    numero_serie_batt_label = tk.Label(left_frame, text="Numéro de série produit:", font=('Arial', 12))
    numero_serie_batt_label.pack(pady=5)
    numero_serie_batt_entry = tk.Entry(left_frame, font=('Arial', 12))
    numero_serie_batt_entry.pack(pady=5)
    
    ctrl_label = tk.Label(right_frame, text="Controle soudure nappes:", font=('Arial', 12))
    ctrl_label.pack(pady=5)
    ctrl_combobox = ttk.Combobox(right_frame,values=['Yes','No'], font=('Arial', 12))
    ctrl_combobox.set('Yes')
    ctrl_combobox.pack(pady=5)
    
    ctrl2_label = tk.Label(right_frame, text="Ecart tension modules:", font=('Arial', 12))
    ctrl2_label.pack(pady=5)
    tension_entry = tk.Entry(right_frame, font=('Arial', 12))
    tension_entry.pack(pady=5)
        
    label = tk.Label(right_frame, text="Commentaire :",font=('Arial', 12))
    label.pack(pady=10)
    text_box = tk.Text(right_frame, height=10, width=50)
    text_box.pack(pady=10)

    
    # Bouton Sauvegarder
    submit_button = tk.Button(right_frame, text="Process", command=etape_nappe, font=('Arial', 12), bg='blue', fg='white')
    submit_button.pack(pady=30, side=tk.BOTTOM)

    # Lancer la boucle principale
    picking_frame.pack()