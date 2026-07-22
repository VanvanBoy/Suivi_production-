# -*- coding: utf-8 -*-
"""
Module extrait automatiquement de Suivi_de_production_prod_V5_5_claude.py
Regroupe les méthodes liées a : TabBancSomfy
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

from config import EXCEL_PATH, BANC_TEST_SOMFY_XLSM_PATH


class TabBancSomfyMixin:
    def setup_banc_somfy(self,frame):
        """
        style = ttk.Style()
        style.theme_use("clam")
        """
        style = ttk.Style()
        BG_MAIN = "#1f2a38"
        FRAME_BG = "#263544"
        ACCENT = "#2e86de"
        SUCCESS = "#27ae60"
        ERROR = "#e74c3c"
        LABEL_GREY = "#6c7a89"
        
        style.configure("Section.TLabel",
                        background=ACCENT,
                        foreground="white",
                        font=("Segoe UI", 10, "bold"),
                        padding=6)
        
        style.configure("Grey.TLabel",
                        background=LABEL_GREY,
                        foreground="white",
                        padding=6)

        style.configure("Result.TLabel",
                        background=SUCCESS,
                        foreground="white",
                        font=("Segoe UI", 10, "bold"),
                        padding=6)

        style.configure("TEntry",
                        padding=6,
                        font=("Segoe UI", 10))

        style.configure("TCombobox",
                        padding=4)

        style.configure("Calc.TButton",
                        background=SUCCESS,
                        foreground="white",
                        font=("Segoe UI", 12, "bold"),
                        padding=12)

        style.configure("Save.TButton",
                        background=SUCCESS,
                        foreground="black",
                        font=("Segoe UI", 12, "bold"),
                        padding=12)
        style.map("Save.TButton",
                  background=[("active", "#1e8449")])

        """
        # STYLES

        style.configure("Main.TFrame", background=FRAME_BG)

        style.configure("Section.TLabel",
                        background=ACCENT,
                        foreground="white",
                        font=("Segoe UI", 10, "bold"),
                        padding=6)

        style.configure("Grey.TLabel",
                        background=LABEL_GREY,
                        foreground="white",
                        padding=6)

        style.configure("Result.TLabel",
                        background=SUCCESS,
                        foreground="white",
                        font=("Segoe UI", 10, "bold"),
                        padding=6)

        style.configure("TEntry",
                        padding=6,
                        font=("Segoe UI", 10))

        style.configure("TCombobox",
                        padding=4)

        style.configure("Calc.TButton",
                        background=SUCCESS,
                        foreground="white",
                        font=("Segoe UI", 12, "bold"),
                        padding=12)

        style.configure("Save.TButton",
                        background=SUCCESS,
                        foreground="white",
                        font=("Segoe UI", 12, "bold"),
                        padding=12)
        style.map("Save.TButton",
                  background=[("active", "#1e8449")])
        """
        # connexion bdd
        def conn_bdd():
            config= {
                'user': 'Vanvan',
                'password': 'VoltR99!',
                'host': '34.77.226.40',
                'database': 'cellules_batteries_cloud',
                'auth_plugin': 'mysql_native_password'
            }
            try:
                return mysql.connector.connect(**config)
            except mysql.connector.Error as err:
                messagebox.showerror("Erreur DB", f"Erreur lors de la connexion : {err}")
                return None
            

        # VALIDATION NUMERIQUE

        def validate_float(P):
            if P == "":
                return True
            try:
                float(P.replace(",", "."))
                return True
            except ValueError:
                return False

        vcmd = (frame.register(validate_float), "%P")


        #BACK

        def on_tension_change(*args):
            tension=float(tension_var.get().replace(",", "."))
            print("Tension lue (V):", tension_var.get())
            if not tension_var.get():
                messagebox.showerror("Pas de tension !", "Veuillez renseigner une tension")
            if len(tension_var.get())==5:
                soc_batt=(54.25 * tension) - 552.58
                soc_display.config(state="normal")
                soc_display.insert(0,soc_batt)
                soc_display.config(state="readonly")
                
                if soc_batt >40 or soc_batt <20:
                    messagebox.showerror("Hors range !", " SOC calculé hors range, donnée moins fiable !")
            

        def dipslay_values(*args):
            try:
                tension = float(tension_var.get())
                current1 = float(current_15w_var.get())
                current2 = float(current_5w_var.get())

            except:
                pass

        def charger_test():
            """Charge automatiquement les 3 mesures (charge 15W, charge 5W,
            charge solaire) depuis le fichier Excel du banc de test Somfy,
            au lieu d'une saisie manuelle."""
            fichier_excel = BANC_TEST_SOMFY_XLSM_PATH
            # Lecture de la feuille "Simple Data" sans utiliser la première ligne comme en-tête
            df = pd.read_excel(fichier_excel, sheet_name="Simple Data", header=None)

            # Récupération des valeurs
            current_15w = df.iloc[3, 5]      # F4
            current_5w = df.iloc[1, 5]       # F2
            charge_solaire = df.iloc[2, 6]   # G3

            # Affectation aux variables Tkinter
            current_15w_var.set(current_15w)
            current_5w_var.set(current_5w)
            charge_solaire_var.set(charge_solaire)

        def enregistrer():
            
            conn = self.db_manager.connect()
            cursor=conn.cursor()
            new_modele='EPDR011AA'
        
            num_batt=battery_entry.get()
            tension_batt=float(tension_var.get().replace(",", "."))
            p_15=float(current_15w_var.get().replace(",", "."))
            p_5=float(current_5w_var.get().replace(",", "."))
            charge_sol=combo15.get()
            soc=soc_display.get()
            
            if str(charge_sol)=="OK":
                charge_sol_int=1
            else :
                charge_sol_int=0
            
            
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
                c15_min = float(seuils["c15_min"])
            except Exception:
                c15_min = float(str(seuils["c15_min"]).replace(",", "."))
            try:
                c15_max = float(seuils["c15_max"])
            except Exception:
                c15_max = float(str(seuils["c15_max"]).replace(",", "."))
            
            try:
                c5_min = float(seuils["c5_min"])
            except Exception:
                c5_min = float(str(seuils["c5_min"]).replace(",", "."))
            try:
                c5_max = float(seuils["c5_max"])
            except Exception:
                c5_max = float(str(seuils["c5_max"]).replace(",", "."))
            
            if c15_min <= p_15 <= c15_max and c5_min <= p_5 <= c5_max and charge_sol_int==1:
                query="update suivi_production set banc_somfy=1, test_tension_finale= %s,test_charge_15W= %s,test_charge_5W= %s,test_charge_sol= %s,soc_calcule= %s,fin_ligne= %s,date_fin_ligne=NOW() where numero_serie_batterie= %s"
                param = (tension_batt, p_15, p_5,charge_sol_int, soc,1,num_batt)  
                cursor.execute(query, param)
                messagebox.showinfo("Controle OK",f"Batterie {num_batt} controlée")
            else :
                messagebox.showerror("Erreur valeurs !", "Valeur de controle NOK")
                cursor.close()
                conn.close()
                return
                
            battery_entry.delete(0,tk.END)
            tension_var.set('')
            current_15w_var.set('')
            current_5w_var.set('')
            soc_display.config(state="normal")
            soc_display.delete(0,tk.END)
            soc_display.config(state="readonly")
            
            conn.commit()
            cursor.close()
            conn.close()


        # VARIABLES

        battery_var = tk.StringVar()

        tension_var = tk.StringVar()
        tension_var.trace_add("write", on_tension_change)
        current_15w_var = tk.StringVar()
        current_5w_var = tk.StringVar()

        charge_solaire_var = tk.StringVar(value="OK")

        soc_var= tk.StringVar()

        """
        # recalcul auto si modification
        tension_var.trace_add("write", auto_calculate)
        current_15w_var.trace_add("write", auto_calculate)
        current_5w_var.trace_add("write", auto_calculate)

        """

        # TITRE

        title = tk.Label(frame,
                         text="BANC DE TEST SOMFY",
                         bg=BG_MAIN,
                         fg="white",
                         font=("Segoe UI", 18, "bold"))
        title.pack(pady=20)

        # FRAME 

        main = ttk.Frame(frame, style="Main.TFrame", padding=30)
        main.pack(padx=40, pady=20, fill="both", expand=True)


        # L1

        ttk.Label(main, text="N° BATTERIE", style="Section.TLabel").grid(row=0, column=0, sticky="ew", pady=10)
        battery_entry = ttk.Entry(main, textvariable=battery_var, width=25)
        battery_entry.grid(row=0, column=1, padx=15)

        ttk.Label(main, text="Tension mesurée (V)", style="Section.TLabel").grid(row=0, column=2, sticky="ew")
        ttk.Entry(main, textvariable=tension_var, width=15, validate="key", validatecommand=vcmd)\
            .grid(row=0, column=3, padx=15)

        # L2

        ttk.Label(main, text="Test Charge 15W", style="Grey.TLabel").grid(row=1, column=0, sticky="ew", pady=15)


        ttk.Label(main, text="Puissance (W)", style="Section.TLabel").grid(row=1, column=1)
        ttk.Entry(main, textvariable=current_15w_var, width=15,
                  validate="key", validatecommand=vcmd)\
            .grid(row=1, column=2)


        # L3

        ttk.Label(main, text="Test Charge 5W", style="Grey.TLabel").grid(row=2, column=0, sticky="ew", pady=15)

        ttk.Label(main, text="Puissance (W)", style="Section.TLabel").grid(row=2, column=1)
        ttk.Entry(main, textvariable=current_5w_var, width=15,
                  validate="key", validatecommand=vcmd)\
            .grid(row=2, column=2)


        # L4
        ttk.Label(main, text="Test Charge Solaire", style="Grey.TLabel").grid(row=3, column=0, sticky="ew", pady=15)

        combo15 = ttk.Combobox(main,
                               textvariable=charge_solaire_var,
                               values=["OK", "NOK"],
                               state="readonly",
                               width=10)
        combo15.grid(row=3, column=1)

        # L5

        ttk.Button(main, text="Charger les valeurs de test", command=charger_test).grid(row=4, column=2)

        # L6

        ttk.Label(main, text="SOC Calculé (%)", style="Result.TLabel").grid(row=4, column=0) 
        soc_display = ttk.Entry(main, textvariable=soc_var, state="readonly", width=15)
        soc_display.grid(row=4,column=1)

        # Enregistrement 

        ttk.Button(frame, text="ENREGISTRER", style="Save.TButton",command=enregistrer).pack(pady=30)

        battery_entry.focus()
        
    
    #Front

