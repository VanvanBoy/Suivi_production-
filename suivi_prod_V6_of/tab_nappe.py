# -*- coding: utf-8 -*-
"""
Module extrait automatiquement de Suivi_de_production_prod_V5_5_claude.py
Regroupe les méthodes liées a : TabNappe
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

from config import EXCEL_PATH, VOLTR_LOGO_PATH, NON_CONFORMITE_FORM_URL


class TabNappeMixin:
    def setup_nappe(self, frame):
        
        left_frame = ttk.Frame(frame)
        left_frame.pack(side="left", fill='both', expand=True, padx=20, pady=20)
    
        list_of = self.show_of_in_process()
        ttk.Label(left_frame, text="N° Of:").pack(pady=5)
        self.combo_of_nappe = ttk.Combobox(left_frame)
        self.combo_of_nappe.pack(pady=5)
        self.combo_of_nappe['values'] = list_of

        ttk.Label(left_frame, text="N° série d'une cellule du produit:").pack(pady=5)
        self.n_numero_serie_cell_entry = ttk.Entry(left_frame)
        self.n_numero_serie_cell_entry.pack(pady=5)
        self.n_numero_serie_cell_entry.bind("<KeyRelease>", self.n_check_entry_length)
    
        ttk.Label(left_frame, text="N° série produit:").pack(pady=5)
        self.n_numero_serie_batt_entry = ttk.Entry(left_frame)
        self.n_numero_serie_batt_entry.pack(pady=5)
        self.n_numero_serie_batt_entry.bind("<KeyRelease>", self.check_entry_length_batt)
    
        ttk.Button(
            left_frame, text="❌ Non conforme",
            command=self.add_non_conf_batterie_nappe,
            style="Danger.TButton"
        ).pack(pady=10)
    
        ttk.Label(left_frame, text="Liste des batteries du modèle:").pack(pady=5)
    
        # --- Bloc Listbox + Scrollbar ---
        listbox_frame = tk.Frame(left_frame)
        listbox_frame.pack(fill="both", expand=True, pady=5)
    
        self.n_listbox_batt = tk.Listbox(
            listbox_frame,
            font=('Segoe UI', 11),
            height=10
        )
        self.n_listbox_batt.pack(side="left", fill="both", expand=True)
    
        scrollbar = tk.Scrollbar(listbox_frame, orient="vertical", command=self.n_listbox_batt.yview)
        scrollbar.pack(side="right", fill="y")
    
        self.n_listbox_batt.config(yscrollcommand=scrollbar.set)
        self.n_listbox_batt.bind("<<ListboxSelect>>", self.n_on_select_batt)
        # --- fin bloc listbox ---
    
        # --- Frame droite ---
        right_frame = ttk.Frame(frame)
        right_frame.pack(side="right", fill='both', expand=True, padx=20, pady=20)
    
        self.avancement_of_nappe = ttk.Label(right_frame, text="Avancement of :")
        self.avancement_of_nappe.pack(pady=5)

        ttk.Label(right_frame, text="Ecart max tension modules (V):").pack(pady=5)
        self.ecart_t_entry = ttk.Entry(right_frame)
        self.ecart_t_entry.pack(pady=5)
        self.ecart_t_entry.bind("<KeyRelease>", self.convert_comma_to_dot)
    
        self.n_label_photo = tk.Label(
            right_frame,
            bg="#e0e0e0",
            width=200,
            height=200,
            text="Aperçu photo",
            anchor='center'
        )
        self.n_label_photo.pack(pady=10)
        self.set_photo(
            self.n_label_photo,
            VOLTR_LOGO_PATH
        )
    
        self.btn_valider_nappe = ttk.Button(
            right_frame, text="✅ Contrôle OK",
            command=self.valider_soudure_nappe,
            style="Good.TButton"
        )
        self.btn_valider_nappe.pack(pady=10)
        
        self.make_tab_chain(
            [
                self.n_numero_serie_batt_entry,
                self.btn_valider_nappe  # placer le bouton en dernier si tu veux que Tab atteigne le bouton
            ],
            submit_button=self.btn_valider_nappe,
            ring=True,
            enter_from_fields=False  # sécurité : Enter depuis les champs ne déclenche pas le bouton
        )
    
        self.display_model_list_nappe()
        
        return self.n_numero_serie_batt_entry
            
        #Back
        
    def n_check_entry_length(self, event=None):
        """Recherche la batterie associée dès que le n° de série cellule (12 car.) est saisi."""
        self._lookup_batterie_from_cellule(self.n_numero_serie_cell_entry, self.n_numero_serie_batt_entry)

    def display_model_list_nappe(self):
        stage_act='nappe'
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
            
            try:
                first_frac, last_frac = self.n_listbox_batt.yview()
            except Exception:
                first_frac = 0.0

    
            # Vide la Listbox
            self.n_listbox_batt.delete(0, tk.END)
            
            # Ajoute chaque batterie dans la Listbox
            for batt in liste_batteries:
                self.n_listbox_batt.insert(tk.END, batt)
                
            # Restaure la position de scroll (clamped automatiquement par Tk)
            try:
                self.n_listbox_batt.yview_moveto(first_frac)
            except Exception:
                pass
    
        except Exception as e:
            messagebox.showerror("Erreur SQL", f"Impossible de récupérer les données :\n{e}")
        finally:
            try:
                cursor.close()
            except:
                pass
            conn.close()

    def valider_soudure_nappe(self):
        modele=str(self.selected_model)
        num_batt=str(self.n_numero_serie_batt_entry.get())
        # 1) Pré-requis
        stage_col='nappe'
        if not self.verif_etape_act_non_ok(stage_col, num_batt):
            return
        if not self._check_prereqs_and_warn(num_batt, stage_col):
            return
        directive=self.verfier_coherence_ref(num_batt)
        if directive=='stop':
            return
        if self.ecart_t_entry.get():
            delta_tension_str=self.ecart_t_entry.get()
            delta_t=delta_tension_str.replace(",", ".")
            delta_tension=float(delta_t)
        else: 
            messagebox.showerror("Saisie incompléte !","Renseigner l'ecart de tension entre le module min et le module max")
            return
        conn = self.db_manager.connect()
        visa=self.db_manager.user
        if not conn:
            return
        
        if modele=="EPDR011AA":
            lim_t=0.025
        else :
            lim_t=0.05
        if delta_tension > lim_t:
            reponse = messagebox.askyesno("Non conformité", "Ecart de tension trop élevé, ouvrir une non-conformité ?")
            if not reponse:
                return
            
            gg_from = NON_CONFORMITE_FORM_URL
            webbrowser.open_new_tab(gg_from) 
            
            num_batt=str(self.n_numero_serie_batt_entry.get())
            
            self.verfier_coherence_ref(num_batt)
            
            try:
                cursor = conn.cursor()
                query = "UPDATE suivi_production SET soudure_nappe_fail = soudure_nappe_fail + 1 where numero_serie_batterie = %s "
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
                self.n_numero_serie_batt_entry.delete(0, tk.END)
                self.ecart_t_entry.delete(0, tk.END)
                    
        else :
            try:
                cursor = conn.cursor()
                query = "UPDATE suivi_production SET soudure_nappe = 1,delta_tension_module = %s, date_soudure_nappe = NOW(), visa_soudure_nappe= %s where numero_serie_batterie = %s "
                param = (delta_tension,visa,num_batt)  
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
                self.n_numero_serie_batt_entry.delete(0, tk.END)
                self.ecart_t_entry.delete(0, tk.END)
                self._focus_active_tab()
                for f in self.funcs_to_run:
                    f()
            
    def add_non_conf_batterie_nappe(self):
        reponse = messagebox.askyesno("Non conformité", "Ouvrir une non-conformité ?")
        
        if reponse:
        
            gg_from = NON_CONFORMITE_FORM_URL
            webbrowser.open_new_tab(gg_from) 
        
        num_batt=str(self.n_numero_serie_batt_entry.get())
        
        self.verfier_coherence_ref(num_batt)
        
        conn = self.db_manager.connect()
        if not conn:
            return
        try:
            cursor = conn.cursor()
            query = "UPDATE suivi_production SET soudure_nappe_fail = soudure_nappe_fail + 1 where numero_serie_batterie = %s "
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
            self.n_numero_serie_batt_entry.delete(0, tk.END)
            self.ecart_t_entry.delete(0, tk.END)
        
    def n_on_select_batt(self, event=None):
        """Quand on sélectionne une batterie dans la listbox, la mettre dans l'Entry produit."""
        self._fill_entry_from_listbox_selection(self.n_listbox_batt, self.n_numero_serie_batt_entry)
        self.of_avancement()
        
    #------------------------------ Onglet bms -----------------------------------------------   

