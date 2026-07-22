# -*- coding: utf-8 -*-
"""
Module extrait automatiquement de Suivi_de_production_prod_V5_5_claude.py
Regroupe les méthodes liées a : TabEmballage
"""
import os, re, shutil
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from PIL import Image, ImageTk
import pandas as pd
import mysql.connector
import webbrowser
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

from config import EXCEL_PATH, NON_CONFORMITE_FORM_URL


class TabEmballageMixin:
    def setup_emb(self, frame):
        left_frame = ttk.Frame(frame)
        left_frame.pack(side="left", fill='both', expand=True, padx=20, pady=20)
        
        # --- Frame droite ---
        right_frame = ttk.Frame(frame)
        right_frame.pack(side="right", fill="both", expand=True, padx=20, pady=20)
        
        ttk.Label(left_frame, text="N° série produit:").pack(pady=5)
        self.emb_numero_serie_batt_entry = ttk.Entry(left_frame)
        self.emb_numero_serie_batt_entry.pack(pady=5)
        self.emb_numero_serie_batt_entry.bind("<KeyRelease>", self.emb_check_entry_length)
        
        ttk.Label(right_frame, text="Liste des batteries du modèle:").pack(pady=5)
        
        # --- Bloc Listbox + Scrollbar ---
        listbox_frame = tk.Frame(right_frame)
        listbox_frame.pack(fill="both", expand=True, pady=5)
        
        self.emb_listbox_batt = tk.Listbox(
            listbox_frame,
            font=('Segoe UI', 11),
            height=10
        )
        self.emb_listbox_batt.pack(side="left", fill="both", expand=True)
        
        scrollbar = tk.Scrollbar(listbox_frame, orient="vertical", command=self.emb_listbox_batt.yview)
        scrollbar.pack(side="right", fill="y")
        
        self.emb_listbox_batt.config(yscrollcommand=scrollbar.set)
        self.emb_listbox_batt.bind("<<ListboxSelect>>", self.emb_on_select_batt)
        # --- fin bloc listbox ---
        
        ttk.Label(left_frame, text="Modèle batterie").pack(pady=5)
        
        cb_var_m = tk.StringVar()
        self.cb_emb = ttk.Combobox(
            left_frame,
            textvariable=cb_var_m,
            values=self.models,
            state="disabled",
            width=40
        )
        self.cb_emb.pack(pady=5)
        
        self.chk_var_emb = tk.BooleanVar(value=False)
        
        def toggle_combobox():
            self.cb_emb.configure(state="readonly" if self.chk_var_emb.get() else "disabled")
        
        self.chk_emb = ttk.Checkbutton(
            left_frame,
            text="Changer le modèle",
            variable=self.chk_var_emb,
            command=toggle_combobox
        )
        self.chk_emb.pack(pady=5)
        
        ttk.Button(
            left_frame, text="❌ Non conforme",
            command=self.add_non_conf_batterie_emb,
            style="Danger.TButton"
        ).pack(pady=10)
        
        ttk.Button(
            left_frame, text="✅ Contrôle OK",
            command=self.valider_emballage,
            style="Good.TButton"
        ).pack(pady=10)
        
        self.display_model_list_emballage()

    def emb_check_entry_length(self, event=None):
        """
        Recherche la batterie associée dès que le champ (12 car.) est saisi.
        NB: contrairement aux autres onglets, l'emballage n'a qu'un seul champ
        (emb_numero_serie_batt_entry) qui sert à la fois de saisie "cellule"
        et de destination pour le numéro de batterie trouvé.
        """
        self._lookup_batterie_from_cellule(self.emb_numero_serie_batt_entry, self.emb_numero_serie_batt_entry)
    
    def display_model_list_emballage(self):
        stage_act='emb'
        modele=str(self.selected_model)
        conn = self.db_manager.connect()
        if not conn:
            return
        try: 
            cursor= conn.cursor()
            query=self.build_stage_query(stage_act)
            param=(modele,)
            cursor.execute(query, param)
            rows = cursor.fetchall()  
    
            # Transforme en liste simple
            liste_batteries = [str(r[0]) for r in rows]
    
            # Vide la Listbox
            self.emb_listbox_batt.delete(0, tk.END)
            
            # Ajoute chaque batterie dans la Listbox
            for batt in liste_batteries:
                self.emb_listbox_batt.insert(tk.END, batt)
    
        except Exception as e:
            messagebox.showerror("Erreur SQL", f"Impossible de récupérer les données :\n{e}")
        finally:
            try:
                cursor.close()
            except:
                pass
            conn.close()

    def valider_emballage(self):
        num_batt=str(self.emb_numero_serie_batt_entry.get())
        stage_col='emb'
        if not self._check_prereqs_and_warn(num_batt, stage_col):
            return
        if self.chk_var_emb.get():
            new_ref=str(self.cb_emb.get())
            self.changer_ref_batterie(new_ref,num_batt)
        else:
            directive=self.verfier_coherence_ref(num_batt)
            if directive=='stop':
                return
        conn = self.db_manager.connect()
        if not conn:
            return
        try:
            cursor = conn.cursor()
            query = "UPDATE suivi_production SET emballage = 1, date_emballage= NOW() where numero_serie_batterie = %s "
            param = (num_batt,)  
            cursor.execute(query, param)
        except Exception as e:
            messagebox.showerror("Erreur SQL", f"Impossible de récupérer les données :\n{e}")
        finally:
            try:
                conn.commit()
                cursor.close()
            except:
                pass
            conn.close()
            messagebox.showinfo("Controle OK",f"Batterie {num_batt} controlée")
            self.emb_numero_serie_batt_entry.delete(0, tk.END)
            for f in self.funcs_to_run:
                f()
            
    def add_non_conf_batterie_emb(self):
        
        reponse = messagebox.askyesno("Non conformité", "Ouvrir une non-conformité ?")
                
        if reponse:
        
            gg_from = NON_CONFORMITE_FORM_URL
            webbrowser.open_new_tab(gg_from) 
        
        self.emb_numero_serie_batt_entry.delete(0, tk.END)
    
    def emb_on_select_batt(self, event=None):
        """Quand on sélectionne une batterie dans la listbox, la mettre dans l'Entry produit."""
        self._fill_entry_from_listbox_selection(self.emb_listbox_batt, self.emb_numero_serie_batt_entry)

    #------------------------------ Onglet capa -----------------------------------------------   

