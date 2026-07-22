# -*- coding: utf-8 -*-
"""
Module extrait automatiquement de Suivi_de_production_prod_V5_5_claude.py
Regroupe les méthodes liées a : TabFinligne
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


class TabFinligneMixin:
    def setup_finligne(self,frame):
        left_frame = ttk.Frame(frame)
        left_frame.pack(side="left", fill='both', expand=True, padx=20, pady=20)
    
        list_of = self.show_of_in_process()
        ttk.Label(left_frame, text="N° of:").pack(pady=5)
        self.combo_of_fin_ligne = ttk.Combobox(left_frame)
        self.combo_of_fin_ligne.pack(pady=5)
        self.combo_of_fin_ligne['values'] = list_of

        ttk.Label(left_frame, text="N° série d'une cellule du produit:").pack(pady=5)
        self.fl_numero_serie_cell_entry = ttk.Entry(left_frame)
        self.fl_numero_serie_cell_entry.pack(pady=5)
        self.fl_numero_serie_cell_entry.bind("<KeyRelease>", self.fl_check_entry_length)
    
        ttk.Label(left_frame, text="N° série produit:").pack(pady=5)
        self.fl_numero_serie_batt_entry = ttk.Entry(left_frame)
        self.fl_numero_serie_batt_entry.pack(pady=5)
        self.fl_numero_serie_batt_entry.bind("<KeyRelease>", self.check_entry_length_batt)
    
        ttk.Button(
            left_frame, text="❌ Non conforme",
            command=self.add_non_conf_batterie_fl,
            style="Danger.TButton"
        ).pack(pady=10)
        
        """
        ttk.Label(left_frame, text="Tension en fin de test:").pack(pady=5)
        self.tension_end_entry = ttk.Entry(left_frame)
        self.tension_end_entry.pack(pady=5)
        self.tension_end_entry.bind("<KeyRelease>", self.convert_comma_to_dot)
        """
        
        ttk.Label(left_frame, text="Controle de charge:").pack(pady=5)
        self.combobox_charge = ttk.Combobox(
            left_frame, state="readonly",
            values=["OK","NOK"]
        )
        self.combobox_charge.pack(pady=5)
        
        ttk.Label(left_frame, text="Controle fonctionnel:").pack(pady=5)
        self.combobox_fonction = ttk.Combobox(
            left_frame, state="readonly",
            values=["OK","NOK"]
        )
        self.combobox_fonction.pack(pady=5)
        
        
        ttk.Label(left_frame, text="Tension fin de ligne (V)").pack(pady=5)
        self.fl_tension_fin_entry = ttk.Entry(left_frame)
        self.fl_tension_fin_entry.pack(pady=5)
        
        if self.selected_model=='EMBR036AG':
            ttk.Label(left_frame, text="Ecart max tension modules (V):").pack(pady=5)
            self.ecart_entry = ttk.Entry(left_frame)
            self.ecart_entry.pack(pady=5)
            self.ecart_entry.bind("<KeyRelease>", self.convert_comma_to_dot)
            
    
        self.btn_valider_fl=ttk.Button(
            left_frame, text="✅ Contrôle OK",
            command=self.valider_fl,
            style="Good.TButton"
        )
        self.btn_valider_fl.pack(pady=10)
        
        self.make_tab_chain(
            [
                self.fl_numero_serie_batt_entry,
                self.btn_valider_fl  # placer le bouton en dernier si tu veux que Tab atteigne le bouton
            ],
            submit_button=self.btn_valider_fl,
            ring=True,
            enter_from_fields=False  # sécurité : Enter depuis les champs ne déclenche pas le bouton
        )
        
        # --- Frame droite ---
        right_frame = ttk.Frame(frame)
        right_frame.pack(side="right", fill='both', expand=True, padx=20, pady=20)
    
        self.avancement_of_fin_ligne = ttk.Label(right_frame, text="Avancement of :")
        self.avancement_of_fin_ligne.pack(pady=5)

        ttk.Label(right_frame, text="Liste des batteries du modèle:").pack(pady=5)
    
        # --- Bloc Listbox + Scrollbar ---
        listbox_frame = tk.Frame(right_frame)
        listbox_frame.pack(fill="both", expand=True, pady=5)
    
        self.fl_listbox_batt = tk.Listbox(
            listbox_frame,
            font=('Segoe UI', 11),
            height=10
        )
        self.fl_listbox_batt.pack(side="left", fill="both", expand=True)
    
        scrollbar = tk.Scrollbar(listbox_frame, orient="vertical", command=self.fl_listbox_batt.yview)
        scrollbar.pack(side="right", fill="y")
    
        self.fl_listbox_batt.config(yscrollcommand=scrollbar.set)
        self.fl_listbox_batt.bind("<<ListboxSelect>>", self.fl_on_select_batt)
        # --- fin bloc listbox ---
    
        self.display_model_list_fl()
        
        return self.fl_numero_serie_batt_entry
        
        
    def fl_check_entry_length(self, event=None):
        """Recherche la batterie associée dès que le n° de série cellule (12 car.) est saisi."""
        self._lookup_batterie_from_cellule(self.fl_numero_serie_cell_entry, self.fl_numero_serie_batt_entry)
    
    def display_model_list_fl(self):
        stage_act='fin_ligne'
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
            self.fl_listbox_batt.delete(0, tk.END)
            
            # Ajoute chaque batterie dans la Listbox
            for batt in liste_batteries:
                self.fl_listbox_batt.insert(tk.END, batt)
    
        except Exception as e:
            messagebox.showerror("Erreur SQL", f"Impossible de récupérer les données :\n{e}")
        finally:
            try:
                cursor.close()
            except:
                pass
            conn.close()

    def valider_fl(self):
        num_batt=str(self.fl_numero_serie_batt_entry.get())
        stage_col='fin_ligne'
        if not self.verif_etape_act_non_ok(stage_col, num_batt):
            return
        if not self._check_prereqs_and_warn(num_batt, stage_col):
            return
        directive=self.verfier_coherence_ref(num_batt)
        if directive=='stop':
            return
        conn = self.db_manager.connect()
        visa=self.db_manager.user
        
        tension_fin=self.fl_tension_fin_entry.get()
        if tension_fin:
            tension_f=tension_fin.replace(",",".")
            tension_end=float(tension_f)
        else :
            tension_end=''
        control_charge=self.combobox_fonction.get()
        control_fonc=self.combobox_fonction.get()
        
        if str(self.selected_model)== 'EMBR036AG':
        
            if self.ecart_entry.get():
                delta_tension_str=self.ecart_entry.get()
                delta_t=delta_tension_str.replace(",", ".")
                delta_tension=float(delta_t)
            else: 
                messagebox.showerror("Saisie incompléte !","Renseigner l'ecart de tension entre le module min et le module max")
                
        
        if str(self.selected_model)== 'EMBR036AG':
            lim_t=0.05
            if delta_tension > lim_t:
                reponse = messagebox.askyesno("Non conformité", "Ecart de tension trop élevé, ouvrir une non-conformité ?")
                if not reponse:
                    return
                
                gg_from = NON_CONFORMITE_FORM_URL
                webbrowser.open_new_tab(gg_from)
        
        if not conn:
            return
        try:
            cursor = conn.cursor()
            if control_charge == "OK" and control_fonc == "OK" and tension_end  != "":
                query = "UPDATE suivi_production SET fin_ligne = 1, test_fonction= 1, test_charge= 1, test_tension_finale= %s, visa_fin_ligne= %s, date_fin_ligne = NOW() where numero_serie_batterie = %s "
                param = (tension_end,visa,num_batt)  
                cursor.execute(query, param)
                
                if str(self.selected_model)== 'EMBR036AG':
                    query_delta=("UPDATE suivi_production SET delta_tension_module = %s where numero_serie_batterie = %s")
                    param_delta=(delta_tension,num_batt)
                    cursor.execute(query_delta, param_delta)
                
                
                cursor.execute("UPDATE produit_voltr SET statut = %s where numero_serie_produit= %s",('stock',num_batt))
                messagebox.showinfo("Controle OK",f"Batterie {num_batt} controlée")
            else:
                if control_charge == 'NOK':
                    cursor.execute("UPDATE suivi_production SET charge_fail = charge_fail + 1 where numero_serie_batterie=%s",(num_batt,))
                elif control_fonc==' NOK':
                    cursor.execute("UPDATE suivi_production SET fonction_fail = fonction_fail + 1 where numero_serie_batterie=%s",(num_batt,))
                messagebox.showwarning("Conditions non remplies",
                                       "Les 2 tests doivent être OK et la tension doit être renseignée.")
        except Exception as e:
            messagebox.showerror("Erreur SQL", f"Impossible de récupérer les données :\n{e}")
        finally:
            try:
                conn.commit()
                cursor.close()
            except:
                pass
            conn.close()
            self.of_avancement()
            self.fl_numero_serie_batt_entry.delete(0, tk.END)
            self._focus_active_tab()
            for f in self.funcs_to_run:
                f()
            
            
    def add_non_conf_batterie_fl(self):
        
        reponse = messagebox.askyesno("Non conformité", "Ouvrir une non-conformité ?")
                
        if reponse:
        
            gg_from = NON_CONFORMITE_FORM_URL
            webbrowser.open_new_tab(gg_from) 
        
        num_batt=str(self.fl_numero_serie_batt_entry.get())
        
        self.verfier_coherence_ref(num_batt)
        
        conn = self.db_manager.connect()
        if not conn:
            return
        try:
            cursor = conn.cursor()
            query = "UPDATE suivi_production SET fermeture_fail = fermeture_fail + 1 where numero_serie_batterie = %s "
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
            self.fl_numero_serie_batt_entry.delete(0, tk.END)
    
    def fl_on_select_batt(self, event=None):
        """Quand on sélectionne une batterie dans la listbox, la mettre dans l'Entry produit."""
        self._fill_entry_from_listbox_selection(self.fl_listbox_batt, self.fl_numero_serie_batt_entry)
        self.of_avancement()
        
    #------------------------------ Onglet recyclage -----------------------------------------------   
