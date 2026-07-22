# -*- coding: utf-8 -*-
"""
Module extrait automatiquement de Suivi_de_production_prod_V5_5_claude.py
Regroupe les méthodes liées a : TabWrap
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


class TabWrapMixin:
    def setup_wrap(self, frame):
        left_frame = ttk.Frame(frame)
        left_frame.pack(side="left", fill='both', expand=True, padx=20, pady=20)
    
        list_of = self.show_of_in_process()
        ttk.Label(left_frame, text="N° of :").pack(pady=5)
        self.combo_of_wrap = ttk.Combobox(left_frame)
        self.combo_of_wrap.pack(pady=5)
        self.combo_of_wrap['values'] = list_of

        ttk.Label(left_frame, text="N° série d'une cellule du produit:").pack(pady=5)
        self.w_numero_serie_cell_entry = ttk.Entry(left_frame)
        self.w_numero_serie_cell_entry.pack(pady=5)
        self.w_numero_serie_cell_entry.bind("<KeyRelease>", self.w_check_entry_length)
    
        ttk.Label(left_frame, text="N° série produit:").pack(pady=5)
        self.w_numero_serie_batt_entry = ttk.Entry(left_frame)
        self.w_numero_serie_batt_entry.pack(pady=5)
        self.w_numero_serie_batt_entry.bind("<KeyRelease>", self.check_entry_length_batt)
    
        ttk.Button(
            left_frame, text="❌ Non conforme",
            command=self.add_non_conf_batterie_wrap,
            style="Danger.TButton"
        ).pack(pady=10)
    
        # --- Frame droite ---
        right_frame = ttk.Frame(frame)
        right_frame.pack(side="right", fill='both', expand=True, padx=20, pady=20)
    
        self.avancement_of_wrap = ttk.Label(right_frame, text="Avancement of :")
        self.avancement_of_wrap.pack(pady=5)

        ttk.Label(right_frame, text="Liste des batteries du modèle:").pack(pady=5)
    
        # --- Bloc Listbox + Scrollbar ---
        listbox_frame = tk.Frame(right_frame)
        listbox_frame.pack(fill="both", expand=True, pady=5)
    
        self.w_listbox_batt = tk.Listbox(
            listbox_frame,
            font=('Segoe UI', 11),
            height=10
        )
        self.w_listbox_batt.pack(side="left", fill="both", expand=True)
    
        scrollbar = tk.Scrollbar(listbox_frame, orient="vertical", command=self.w_listbox_batt.yview)
        scrollbar.pack(side="right", fill="y")
    
        self.w_listbox_batt.config(yscrollcommand=scrollbar.set)
        self.w_listbox_batt.bind("<<ListboxSelect>>", self.w_on_select_batt)
        # --- fin bloc listbox ---
    
        self.btn_valider_wrap=ttk.Button(
            left_frame, text="✅ Contrôle OK",
            command=self.valider_wrap,
            style="Good.TButton"
        )
        self.btn_valider_wrap.pack(pady=10)
        
        self.make_tab_chain(
            [
                self.w_numero_serie_batt_entry,
                self.btn_valider_wrap  # placer le bouton en dernier si tu veux que Tab atteigne le bouton
            ],
            submit_button=self.btn_valider_wrap,
            ring=True,
            enter_from_fields=False  # sécurité : Enter depuis les champs ne déclenche pas le bouton
        )
    
    
        self.display_model_list_wrap()
        
        return self.w_numero_serie_batt_entry

        
    def w_check_entry_length(self, event=None):
        """Recherche la batterie associée dès que le n° de série cellule (12 car.) est saisi."""
        self._lookup_batterie_from_cellule(self.w_numero_serie_cell_entry, self.w_numero_serie_batt_entry)
    
    def display_model_list_wrap(self):
        stage_act='wrap'
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
            self.w_listbox_batt.delete(0, tk.END)
            
            # Ajoute chaque batterie dans la Listbox
            for batt in liste_batteries:
                self.w_listbox_batt.insert(tk.END, batt)
    
        except Exception as e:
            messagebox.showerror("Erreur SQL", f"Impossible de récupérer les données :\n{e}")
        finally:
            try:
                cursor.close()
            except:
                pass
            conn.close()

    def valider_wrap(self):
        num_batt=str(self.w_numero_serie_batt_entry.get())
        stage_col='wrap'
        if not self.verif_etape_act_non_ok(stage_col, num_batt):
            return
        if not self._check_prereqs_and_warn(num_batt, stage_col):
            return
        conn = self.db_manager.connect()
        visa=self.db_manager.user
        directive=self.verfier_coherence_ref(num_batt)
        if directive=='stop':
            return
        if not conn:
            return
        try:
            cursor = conn.cursor()
            query = "UPDATE suivi_production SET wrap = 1, date_wrap = NOW(), visa_wrap= %s where numero_serie_batterie = %s "
            param = (visa,num_batt)  
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
            self.of_avancement()
            self.w_numero_serie_batt_entry.delete(0, tk.END)
            self._focus_active_tab()
            for f in self.funcs_to_run:
                f()
            
            
    def add_non_conf_batterie_wrap(self):
        
        reponse = messagebox.askyesno("Non conformité", "Ouvrir une non-conformité ?")
                
        if reponse:
        
            gg_from = NON_CONFORMITE_FORM_URL
            webbrowser.open_new_tab(gg_from) 
        
        num_batt=str(self.w_numero_serie_batt_entry.get())
        
        self.verfier_coherence_ref(num_batt)
        
        conn = self.db_manager.connect()
        if not conn:
            return
        try:
            cursor = conn.cursor()
            query = "UPDATE suivi_production SET wrap_fail = wrap_fail + 1 where numero_serie_batterie = %s "
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
            self.w_numero_serie_batt_entry.delete(0, tk.END)
    
    def w_on_select_batt(self, event=None):
        """Quand on sélectionne une batterie dans la listbox, la mettre dans l'Entry produit."""
        self._fill_entry_from_listbox_selection(self.w_listbox_batt, self.w_numero_serie_batt_entry)
        self.of_avancement()
        
        
    #------------------------------ Onglet fermeture -----------------------------------------------   
