# -*- coding: utf-8 -*-
"""
Module extrait automatiquement de Suivi_de_production_prod_V5_5_claude.py
Regroupe les méthodes liées a : TabTriTest
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

from config import EXCEL_PATH


class TabTriTestMixin:
    def setup_tri_test(self,frame):
            
        # ===== Frame principale =====
        main_frame = ttk.Frame(frame)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
    
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=0)  # top_frame
        main_frame.rowconfigure(1, weight=0)  # titre_frame
        main_frame.rowconfigure(2, weight=1)  # bottom_frame (treeviews)
        
    
        # ===== Frame du haut (entrée) =====
        top_frame = ttk.Frame(main_frame)
        top_frame.grid(row=0, column=0, pady=(0, 20))
    
        ttk.Label(top_frame, text="N° série produit :").grid(row=0, column=0, padx=5)
        self.t_numero_serie_batt_entry = ttk.Entry(top_frame, width=30)
        self.t_numero_serie_batt_entry.grid(row=0, column=1, padx=5)
    
        self.t_numero_serie_batt_entry.bind("<KeyRelease>", self.check_entry_length_tri)
        
        # mid france pour titre 
        titre_frame = ttk.Frame(main_frame)
        titre_frame.grid(row=1, column=0, pady=(0, 5), sticky="n")
        
        left_ti_frame = ttk.Frame(titre_frame)
        left_ti_frame.grid(row=0, column=0, padx=(0, 10))
        self.title_left=ttk.Label(left_ti_frame,text="Test DCIR only (2.0)",font=("Segoe UI", 10, "bold"))
        self.title_left.grid(row=0, column=0, padx=5)
        
        right_ti_frame = ttk.Frame(titre_frame)
        right_ti_frame.grid(row=0, column=1, padx=(0, 10))
        self.title_right=ttk.Label(right_ti_frame,text="Test complet (0.1 ou 1.0)",font=("Segoe UI", 10, "bold"))
        self.title_right.grid(row=0, column=1, padx=5)
 
        # ===== Frame du bas (treeviews) =====
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.grid(row=2, column=0, sticky="nsew")
    
        bottom_frame.columnconfigure(0, weight=1)
        bottom_frame.columnconfigure(1, weight=1)
        bottom_frame.rowconfigure(0, weight=1)
    
        # ===== Treeview gauche ====
        
        left_tv_frame = ttk.Frame(bottom_frame)
        left_tv_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
    
        columns = ("numero_serie", "reference", "date_fin_ligne")
    
        self.tree_left = ttk.Treeview(
            left_tv_frame,
            columns=columns,
            show="headings"
        )
    
        for col in columns:
            self.tree_left.heading(col, text=col.replace("_", " ").title())
            self.tree_left.column(col, anchor="center", width=140)
    
        self.tree_left.pack(fill="both", expand=True)
    
        # ===== Treeview droite =====
        right_tv_frame = ttk.Frame(bottom_frame)
        right_tv_frame.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
    
        self.tree_right = ttk.Treeview(
            right_tv_frame,
            columns=columns,
            show="headings"
        )
    
        for col in columns:
            self.tree_right.heading(col, text=col.replace("_", " ").title())
            self.tree_right.column(col, anchor="center", width=140)
    
        self.tree_right.pack(fill="both", expand=True)
        

    def check_entry_length_tri(self, event):
        
        numero_serie = self.t_numero_serie_batt_entry.get().strip()

        # On ne fait rien si le numéro n'a pas exactement 9 chiffres
        if len(numero_serie) != 9:
            return
    
        try:
            # Connexion à la BDD via ton db_manager
            conn = self.db_manager.connect()
            cursor = conn.cursor()
    
            # Requête pour récupérer reference et type de cyclage
            query = """
                SELECT reference_produit_voltr, type_cyclage, date_fin_ligne
                FROM produit_voltr
                join suivi_production 
                on numero_serie_produit=numero_serie_batterie
                WHERE numero_serie_produit = %s
            """
            cursor.execute(query, (numero_serie,))
            result = cursor.fetchone()  # On récupère une seule ligne
    
            if result:
                reference, type_cyclage,date = result
                # Préparer la ligne à insérer
                row_values = (numero_serie, reference, str(date))
    
                # Insérer dans le Treeview correspondant
                if str(type_cyclage) == "2.0":
                    self.tree_left.insert("", "end", values=row_values)
                elif str(type_cyclage) == "1.0":
                    self.tree_right.insert("", "end", values=row_values)
                else:
                    print(f"Type de cyclage inconnu: {type_cyclage}")
            else:
                print(f"Aucune donnée trouvée pour le numéro de série {numero_serie}")
    
        except Exception as e:
            print(f"Erreur lors de la récupération du type de cyclage: {e}")
    
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()
                
            self.t_numero_serie_batt_entry.delete(0,tk.END)
        
    #------------------------------ Onglet picking -----------------------------------------------
    
