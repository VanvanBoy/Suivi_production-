# -*- coding: utf-8 -*-
"""
Module extrait automatiquement de Suivi_de_production_prod_V5_5_claude.py
Regroupe les méthodes liées a : TabPack
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


class TabPackMixin:
    def setup_pack(self, frame):
        left_frame = ttk.Frame(frame)
        left_frame.pack(side="left", fill='both', expand=True, padx=20, pady=20)

        list_of = self.show_of_in_process()
        ttk.Label(left_frame, text="Entrer un numero d'of :").pack(pady=5)
        self.s_numero_of = ttk.Combobox(left_frame)
        self.s_numero_of.pack(pady=5)
        self.s_numero_of["values"] = list_of
        self.s_numero_of.bind("<<ComboboxSelected>>", self.of_avancement)
    
        ttk.Label(left_frame, text="N° série d'une cellule du produit:").pack(pady=5)
        self.s_numero_serie_cell_entry = ttk.Entry(left_frame)
        self.s_numero_serie_cell_entry.pack(pady=5)
        self.s_numero_serie_cell_entry.bind("<KeyRelease>", self.s_check_entry_length)
    
        ttk.Label(left_frame, text="N° série produit:").pack(pady=5)
        self.s_numero_serie_batt_entry = ttk.Entry(left_frame)
        self.s_numero_serie_batt_entry.pack(pady=5)
        
        if str(self.selected_model)[:8]=='PPTR018A':
        
            ttk.Label(left_frame, text="Choisir une reference EOP").pack(pady=5)
            
            VALS = ["PPTR018AA", "PPTR018AB", "PPTR018AC","PPTR018AD"]

            # Variable liée
            self.choice = tk.StringVar(value=VALS[0])
                
            self.s_mod_combobox=ttk.Combobox(left_frame, textvariable=self.choice, values=VALS, state="readonly", width=20)
            self.s_mod_combobox.insert(0,str(self.selected_model))  # sélectionne l'élément par défaut
            self.s_mod_combobox.pack(pady=5)
            
            ttk.Label(left_frame, text="Mesure d'impedance (Ohms)").pack(pady=5)
            self.impedance_entry= ttk.Entry(left_frame)
            self.impedance_entry.pack(pady=5)
            self.impedance_entry.bind("<KeyRelease>", self.convert_comma_to_dot)
            
            
            ttk.Label(left_frame, text="Mesure tension (V)").pack(pady=5)
            self.tension_eop_entry= ttk.Entry(left_frame)
            self.tension_eop_entry.pack(pady=5)
            self.tension_eop_entry.bind("<KeyRelease>", self.convert_comma_to_dot)
        
        if str(self.selected_model)=='EPDR011AA' or str(self.selected_model)== 'EMBR036AG':
            ttk.Label(left_frame, text="Mesure d'impedance (mOhm)").pack(pady=5)
            self.impedance_entry= ttk.Entry(left_frame)
            self.impedance_entry.pack(pady=5)
            self.impedance_entry.bind("<KeyRelease>", self.convert_comma_to_dot)
            
            
        ttk.Button(
            left_frame, text="❌ Non conforme",
            command=self.add_non_conf_batterie_pack,
            style="Danger.TButton"
        ).pack(pady=10)
    
        # --- Frame droite ---
        right_frame = ttk.Frame(frame)
        right_frame.pack(side="right", fill='both', expand=True, padx=20, pady=20)
    
        self.avancement_of_pack = ttk.Label(right_frame, text="Avancement of :")
        self.avancement_of_pack.pack(pady=5)

        ttk.Label(right_frame, text="Liste des batteries du modèle:").pack(pady=5)
    
        # --- Bloc Listbox + Scrollbar ---
        listbox_frame = tk.Frame(right_frame)
        listbox_frame.pack(fill="both", expand=True, pady=5)
    
        self.s_listbox_batt = tk.Listbox(
            listbox_frame,
            font=('Segoe UI', 11),
            height=10
        )
        self.s_listbox_batt.pack(side="left", fill="both", expand=True)
    
        scrollbar = tk.Scrollbar(listbox_frame, orient="vertical", command=self.s_listbox_batt.yview)
        scrollbar.pack(side="right", fill="y")
    
        self.s_listbox_batt.config(yscrollcommand=scrollbar.set)
        self.s_listbox_batt.bind("<<ListboxSelect>>", self.s_on_select_batt)
        # --- fin bloc listbox ---
    
        self.btn_valider_pack = ttk.Button(
            left_frame, text="✅ Contrôle OK",
            command=self.valider_soudure_pack,
            style="Good.TButton"
        )
        self.btn_valider_pack.pack(pady=10)
        
        self.display_model_list_pack()
        
        self.make_tab_chain(
            [
                self.s_numero_serie_batt_entry,
                self.btn_valider_pack  # placer le bouton en dernier si tu veux que Tab atteigne le bouton
            ],
            submit_button=self.btn_valider_pack,
            ring=True,
            enter_from_fields=False  # sécurité : Enter depuis les champs ne déclenche pas le bouton
        )   
        
        return self.s_numero_serie_batt_entry

    #Back
    
    def s_check_entry_length(self, event=None):
        """Recherche la batterie associée dès que le n° de série cellule (12 car.) est saisi."""
        self._lookup_batterie_from_cellule(self.s_numero_serie_cell_entry, self.s_numero_serie_batt_entry)
    
    def display_model_list_pack(self):
        stage_act='pack'
        modele=str(self.selected_model)
        conn = self.db_manager.connect()
        if not conn:
            return
        try: 
            cursor= conn.cursor()
            if modele[:8]=="PPTR018A":
                query ="""
                    SELECT sp.numero_serie_batterie
                    FROM suivi_production AS sp
                    JOIN produit_voltr AS pv
                      ON sp.numero_serie_batterie = pv.numero_serie_produit
                    WHERE sp.picking_tension = 1
                      AND (sp.soudure_pack=0 or sp.soudure_pack is null)
                      AND (sp.recyclage = 0 or sp.recyclage is null)
                      AND pv.reference_produit_voltr like %s
                     """
                param=(modele[:8]+"%",)
                
                
            else :
                
                query=self.build_stage_query(stage_act)
                param=(modele,)
            cursor.execute(query, param)
            rows = cursor.fetchall()  
    
            # Transforme en liste simple
            liste_batteries = [str(r[0]) for r in rows]
    
            # Vide la Listbox
            self.s_listbox_batt.delete(0, tk.END)
            
            # Ajoute chaque batterie dans la Listbox
            for batt in liste_batteries:
                self.s_listbox_batt.insert(tk.END, batt)
    
        except Exception as e:
            messagebox.showerror("Erreur SQL", f"Impossible de récupérer les données :\n{e}")
        finally:
            try:
                cursor.close()
            except:
                pass
            conn.close()

    def valider_soudure_pack(self):
        num_of = self.s_numero_of.get()
        if num_of == '':
            messagebox.showerror("Num of manquant !", "Veuillez renseigner l'of")
            return

        modele=str(self.selected_model)
        if modele[:8]=='PPTR018A':
            new_modele=self.s_mod_combobox.get()
            if new_modele=="— choisir —":
                messagebox.showerror("Modele EOP","Choisir un modele d'EOP")
                return
        num_batt=str(self.s_numero_serie_batt_entry.get())
        # 1) Pré-requis
        stage_col='pack'
        if not self.verif_etape_act_non_ok(stage_col, num_batt):
            return
        if not self._check_prereqs_and_warn(num_batt, stage_col):
            return
        conn = self.db_manager.connect()
        visa=self.db_manager.user
        if not conn:
            return
        if modele[:8]!='PPTR018A':
            directive=self.verfier_coherence_ref(num_batt)
            if directive=='stop':
                return
            
        try:
            cursor = conn.cursor()
            if modele[:8]=='PPTR018A':
                impedance=self.impedance_entry.get()
                controle_tension=self.tension_eop_entry.get()
                if impedance:
                    impedance=float(impedance)
                else :
                    messagebox.showerror("Impedance non renseignée", "Veuillez renseigner l'impedance")
                    
                if controle_tension:
                    controle_tension=float(controle_tension)
                else :
                    messagebox.showerror("Tension non renseignée", "Veuillez renseigner la tension")
                df_cyclage = pd.read_excel(EXCEL_PATH,sheet_name="Cyclage",header=1)
                try:
                    query = """
                        SELECT reference_cellule
                        FROM cellule
                        WHERE affectation_produit = %s
                        LIMIT 1
                    """
                    cursor.execute(query, (num_batt,))
                    row_db = cursor.fetchone()
                except Exception as e:
                    # Erreur SQL => échec de traitement
                    messagebox.showerror("Reference cellule introuvable !", f"Pas de reference cellule pour la batterie {num_batt}")
                    row_db = None
                    return
                
                if not row_db or not row_db[0]:
                    messagebox.showerror("Reference cellule introuvable !", f"Pas de reference cellule pour la batterie {num_batt}")   
                    return
                    
                ref_cell = row_db[0]
                # --- Seuils (df_cyclage) ---
                row = df_cyclage[
                    (df_cyclage["Nom_modele"] == new_modele) &
                    (df_cyclage["Ref cellule"] == ref_cell)
                ]

                if row.empty:
                    messagebox.showerror("Duo batterie/cellule introuvable !", f"Pas de batterie {num_batt} avec la cellule {ref_cell} dans le suivi de production")   
                    return
                    
                seuils = row.iloc[0]
                try:
                    imp_min = float(seuils["Impedance borne min"])
                except Exception:
                    imp_min = float(str(seuils["Impedance borne min"]).replace(",", "."))
                try:
                    imp_max = float(seuils["Impedance borne max"])
                except Exception:
                    imp_max = float(str(seuils["Impedance borne max"]).replace(",", "."))
                
                if imp_min <= impedance <= imp_max :
                    query = "UPDATE suivi_production SET soudure_pack = 1, date_soudure_pack = NOW(),test_impedance= %s, date_impedance=NOW(), test_tension_pack=%s, date_test_tension= NOW(), visa_soudure_pack= %s where numero_serie_batterie = %s "
                    param = (impedance,controle_tension,visa,num_batt)  
                    cursor.execute(query, param)
                    messagebox.showinfo("Controle OK",f"Batterie {num_batt} controlée")
                else :
                    messagebox.showerror("Erreur impedance", "La valeure d'impedance est NOK")
                    return
            
            elif modele=='EPDR011AA' or modele=='EMBR036AG':
                impedance=self.impedance_entry.get()
                if impedance:
                    impedance=float(impedance)
                else :
                    messagebox.showerror("Impedance non renseignée", "Veuillez renseigner l'impedance")
                    
                df_cyclage = pd.read_excel(EXCEL_PATH,sheet_name="Cyclage",header=1)
                try:
                    query = """
                        SELECT reference_cellule
                        FROM cellule
                        WHERE affectation_produit = %s
                        LIMIT 1
                    """
                    cursor.execute(query, (num_batt,))
                    row_db = cursor.fetchone()
                except Exception as e:
                    # Erreur SQL => échec de traitement
                    messagebox.showerror("Reference cellule introuvable !", f"Pas de reference cellule pour la batterie {num_batt}")
                    row_db = None
                    return
                
                if not row_db or not row_db[0]:
                    messagebox.showerror("Reference cellule introuvable !", f"Pas de reference cellule pour la batterie {num_batt}")   
                    return
                    
                ref_cell = row_db[0]
                # --- Seuils (df_cyclage) ---
                row = df_cyclage[
                    (df_cyclage["Nom_modele"] == modele) &
                    (df_cyclage["Ref cellule"] == ref_cell)
                ]

                if row.empty:
                    messagebox.showerror("Duo batterie/cellule introuvable !", f"Pas de batterie {num_batt} avec la cellule {ref_cell} dans le suivi de production")   
                    return
                    
                seuils = row.iloc[0]
                try:
                    imp_min = float(seuils["Impedance borne min"])
                except Exception:
                    imp_min = float(str(seuils["Impedance borne min"]).replace(",", "."))
                try:
                    imp_max = float(seuils["Impedance borne max"])
                except Exception:
                    imp_max = float(str(seuils["Impedance borne max"]).replace(",", "."))
                
                if imp_min <= impedance <= imp_max :
                    query = "UPDATE suivi_production SET soudure_pack = 1, date_soudure_pack = NOW(),test_impedance= %s, date_impedance=NOW(), visa_soudure_pack= %s where numero_serie_batterie = %s "
                    param = (impedance,visa,num_batt)  
                    cursor.execute(query, param)
                    messagebox.showinfo("Controle OK",f"Batterie {num_batt} controlée")
                else :
                    messagebox.showerror("Erreur impedance", "La valeure d'impedance est NOK")
                    return
                    
            else :    
                query = "UPDATE suivi_production SET soudure_pack = 1, date_soudure_pack = NOW(), visa_soudure_pack= %s where numero_serie_batterie = %s "
                param = (visa,num_batt)  
                cursor.execute(query, param)
                messagebox.showinfo("Controle OK",f"Batterie {num_batt} controlée")
                
            if modele[:8]=='PPTR018A':
            
                query_modele='UPDATE produit_voltr set reference_produit_voltr= %s where numero_serie_produit = %s'
                param_modele=(new_modele,num_batt)
                cursor.execute(query_modele, param_modele)

            query_of = 'UPDATE produit_voltr set n_of= %s where numero_serie_produit = %s'
            param_of = (num_of, num_batt)
            cursor.execute(query_of, param_of)

            # Si l'OF était encore "en attente", on le passe "en cours" dès que
            # la 1ère batterie lui est affectée (no-op si déjà en cours/terminé).
            cursor.execute(
                "UPDATE ref_of SET etat_fabrication = %s WHERE n_of = %s AND etat_fabrication = %s",
                ("en cours", num_of, "en attente")
            )

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
            self.s_numero_serie_batt_entry.delete(0, tk.END)
            self._focus_active_tab()
            for f in self.funcs_to_run:
                f()
            
            
    def add_non_conf_batterie_pack(self):
        
        reponse = messagebox.askyesno("Non conformité", "Ouvrir une non-conformité ?")
        if reponse:
            gg_from = NON_CONFORMITE_FORM_URL
            webbrowser.open_new_tab(gg_from) 
           
        
        num_batt=str(self.s_numero_serie_batt_entry.get())
        
        self.verfier_coherence_ref(num_batt)
        
        conn = self.db_manager.connect()
        if not conn:
            return
        try:
            cursor = conn.cursor()
            query = "UPDATE suivi_production SET soudure_pack_fail = soudure_pack_fail + 1 where numero_serie_batterie = %s "
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
            self.s_numero_serie_batt_entry.delete(0, tk.END)
    
    def s_on_select_batt(self, event=None):
        """Quand on sélectionne une batterie dans la listbox, la mettre dans l'Entry produit."""
        self._fill_entry_from_listbox_selection(self.s_listbox_batt, self.s_numero_serie_batt_entry)
            
    
        
    #------------------------------ Onglet soudure nappe -----------------------------------------------   

