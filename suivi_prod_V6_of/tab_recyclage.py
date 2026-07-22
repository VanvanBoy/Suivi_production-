# -*- coding: utf-8 -*-
"""
Module extrait automatiquement de Suivi_de_production_prod_V5_5_claude.py
Regroupe les méthodes liées a : TabRecyclage
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


class TabRecyclageMixin:
    def setup_recyclage(self,frame):

        left = ttk.Frame(frame)
        left.pack(side="left", fill='both', expand=True, padx=20, pady=20)
        
        right = ttk.Frame(frame)
        right.pack(side="right", fill='both', expand=True, padx=20, pady=20)
        
        ttk.Label(left, text="N° série cellule").pack(pady=5)
        self.r_entry_cell = ttk.Entry(left, width=28)
        self.r_entry_cell.pack(pady=5)
        # Remplissage auto du n° batterie quand l'entry cellule atteint 12 chars
        self.r_entry_cell.bind("<KeyRelease>", self.r_on_cell_entry)
    
        ttk.Label(left, text="N° série batterie").pack(pady=5)
        self.r_entry_batt = ttk.Entry(left, width=28)
        self.r_entry_batt.pack(pady=5)
    
        ttk.Label(left, text="Référence batterie").pack(pady=5)
        # valeur par défaut nulle (vide)
        self.r_model_var = tk.StringVar(value="")
        self.r_combo = ttk.Combobox(left, textvariable=self.r_model_var,
                                       values=(self.models or []), state="readonly", width=30)
        self.r_combo.pack(pady=5)
        self.r_combo.bind("<<ComboboxSelected>>", lambda e: self.r_on_model_change())
        
        ttk.Label(left, text="Cause recyclage:").pack(pady=5)
        
        self.cause_var = tk.StringVar(value="")
        self.cause_combo = ttk.Combobox(left, textvariable=self.cause_var,
                                       values=("erreur_soudure","choc","autres"), state="readonly", width=30)
        self.cause_combo.pack(pady=5)
        
        ttk.Button(
            left, text="🔄 Recyler la batterie",
            command=self.recycle_batterie_sp, style="Danger.TButton"
        ).pack(pady=10)
    
        ttk.Label(right, text="Liste batterie").pack(pady=5)
        
        self.r_listbox = tk.Listbox(right, height=10, activestyle="dotbox", selectmode="extended")
        yscroll = ttk.Scrollbar(right, orient="vertical", command=self.r_listbox.yview)
        self.r_listbox.configure(yscrollcommand=yscroll.set)
        self.r_listbox.pack(side="right", fill="both", expand=True)
        yscroll.pack(side="right", fill="y")
        self.r_listbox.bind("<<ListboxSelect>>", self.on_r_listbox_select)
        
    def on_r_listbox_select(self, event):
        # obtenir indices sélectionnés (peut être plusieurs si selectmode="extended")
        sel = self.r_listbox.curselection()
        if not sel:
            return
        # on prend le premier sélectionné
        idx = sel[0]
        text = self.r_listbox.get(idx)
    
        # --- si text est déjà le numero de série simple ---
        # numero = text.strip()
    
        # --- OU : si text contient d'autres champs, extraire le numéro ---
        # Exemples d'extraction (choisis celle qui correspond à ton format)
        # 1) format "123456789012" => direct
        # 2) format "1;123456789012;moduleA" => split par ';' et prendre le 2ème
        # 3) format "1 | 123456789012 | module A" => split par '|' et strip
        numero = None
        if ";" in text:
            parts = [p.strip() for p in text.split(";")]
            # si le numéro est en 2e position
            if len(parts) >= 2:
                numero = parts[1]
        elif "|" in text:
            parts = [p.strip() for p in text.split("|")]
            # chercher la première partie qui ressemble à un n° (ex: longueur 12, chiffres)
            for p in parts:
                if p and any(ch.isdigit() for ch in p):
                    numero = p
                    break
        else:
            # par défaut on prend toute la chaîne
            numero = text.strip()
    
        # si tu utilises une StringVar pour l'entry, mets-la ; sinon delete/insert
        try:
            self.r_entry_batt.delete(0, "end")
            if numero:
                self.r_entry_batt.insert(0, numero)
        except Exception as e:
            # fallback si self.r_entry_batt est une StringVar
            if hasattr(self, "r_model_var") and isinstance(self.r_model_var, tk.StringVar):
                self.r_model_var.set(numero or "")
            else:
                raise
    
    def r_on_cell_entry(self, event=None):
        """Quand l'entry cellule atteint 12 chars, on cherche la batterie associée et on remplit l'entry batterie."""
        numero_serie_cell = self.r_entry_cell.get().strip()
        if len(numero_serie_cell) != 12:
            return
        conn = self.db_manager.connect()
        if not conn:
            return
        try:
            cur = conn.cursor()
            # affectation_produit = numero_serie_batterie (selon ta logique existante)
            cur.execute("SELECT affectation_produit FROM cellule WHERE numero_serie_cellule = %s", (numero_serie_cell,))
            row = cur.fetchone()
            self.r_entry_batt.delete(0, tk.END)
            if row and row[0]:
                self.rech_entry_batt.insert(0, str(row[0]))
        except Exception as e:
            messagebox.showerror("Erreur SQL", f"Lookup cellule→batterie impossible :\n{e}")
        finally:
            try: cur.close()
            except: pass
            conn.close()
        
    def r_on_model_change(self):
        """Quand on choisit un modèle, on alimente la liste des n° batteries via la jointure demandée."""
        ref = self.r_model_var.get().strip()
        self.r_listbox.delete(0, tk.END)
        if not ref:
            return
        conn = self.db_manager.connect()
        if not conn:
            return
        try:
            cur = conn.cursor()
            # Liste des NUMÉROS DE SÉRIE BATTERIE pour la référence choisie
            # sp = suivi_production / p = produit
            sql = ("""
                SELECT DISTINCT sp.numero_serie_batterie
                FROM suivi_production sp
                JOIN produit_voltr p
                  ON sp.numero_serie_batterie = p.numero_serie_produit
                WHERE p.reference_produit_voltr = %s
                AND recyclage is null 
                ORDER BY sp.numero_serie_batterie
            """)
            cur.execute(sql, (ref,))
            for (num_batt,) in cur.fetchall():
                if num_batt:
                    self.r_listbox.insert(tk.END, str(num_batt))
        except Exception as e:
            messagebox.showerror("Erreur SQL", f"Chargement liste batteries impossible :\n{e}")
        finally:
            try: cur.close()
            except: pass
            conn.close()
    
    def recycle_batterie_sp(self):
        type_obj = "batterie"
        cause = self.cause_combo.get()
        if not cause:
            messagebox.showerror("Pas de cause", "Veuillez sélectionner une cause !")
            return
    
        numero_serie_batt = self.r_entry_batt.get().strip()
        if not numero_serie_batt:
            messagebox.showerror("Erreur", "Veuillez renseigner le numéro de série de la batterie.")
            return
    
        conn = None
        cursor = None
        try:
            conn = self.db_manager.connect()
            cursor = conn.cursor()
    
            # 1) Récupérer référence et poids de la batterie
            query = "SELECT pv.reference_produit_voltr, rv.poids FROM produit_voltr as pv join ref_batterie_voltr as rv on pv.reference_produit_voltr=rv.reference_batterie_voltr WHERE numero_serie_produit = %s"
            cursor.execute(query, (numero_serie_batt,))
            row_prod = cursor.fetchone()
            if row_prod is None:
                messagebox.showerror("Non trouvé", f"Aucune batterie trouvée pour le n° {numero_serie_batt}")
                return
    
            reference_batt = row_prod[0]
            poids_batt = row_prod[1] or 0
    
            # 2) Lire la feuille Excel et trouver le dest_recyclage
            df_cyclage = pd.read_excel(EXCEL_PATH, sheet_name="Cyclage", header=1)
            sel = df_cyclage[df_cyclage["Nom_modele"] == reference_batt]
            if sel.empty:
                messagebox.showerror("Erreur Excel", f"Modèle {reference_batt} introuvable dans {EXCEL_PATH} sheet Cyclage.")
                return
            seuils = sel.iloc[0]
            dest_recyclage = str(seuils.get("Recyclage", "")).strip()
            if not dest_recyclage:
                messagebox.showerror("Erreur", f"Pas de destination de recyclage définie pour {reference_batt}.")
                return
            type_fut = dest_recyclage
    
            # 3) Chercher un fut ouvert pour ce type
            cursor.execute(
                "SELECT id_fut, poids FROM fut_recyclage WHERE exutoire = %s AND etat_fut = %s LIMIT 1",
                (type_fut, "en cours")
            )
            fut_row = cursor.fetchone()
            if fut_row is None:
                messagebox.showerror("Erreur !", "Aucun fut d'exutoire eo_org_mtl n'est ouvert")
                return
    
            id_fut, poids_fut = fut_row[0], fut_row[1] or 0
    
            # 4) Mettre à jour le poids du fut
            poids_tot = poids_fut + poids_batt
            cursor.execute("UPDATE fut_recyclage SET poids = %s WHERE id_fut = %s", (poids_tot, id_fut))
    
            # 5) Insérer la ligne de recyclage (remarquer VALUES (...) et les colonnes explicitement)
            query_recy = """
                INSERT INTO recyclage
                    (numero_serie, type_objet, id_fut, sur_site, date_rebut, cause)
                VALUES (%s, %s, %s, %s, NOW(), %s)
            """
            param_recy = (numero_serie_batt, type_obj, id_fut, "oui", cause)
            cursor.execute(query_recy, param_recy)
    
            # 6) Mettre à jour le suivi_production
            query_sp = "UPDATE suivi_production SET recyclage = 1, date_recyclage = NOW() WHERE numero_serie_batterie = %s"
            cursor.execute(query_sp, (numero_serie_batt,))
            cursor.execute("Update produit_voltr set statut=%s where numero_serie_produit=%s",("recyclee",numero_serie_batt))
    
            # 7) Commit et message utilisateur
            conn.commit()
            emplacement = f"fut {id_fut}"
            messagebox.showinfo("Recyclage réussi",
                                f"La batterie {numero_serie_batt} recyclée dans un fut {type_fut} : {emplacement}")
    
        except Exception as e:
            # rollback en cas d'erreur et message
            if conn:
                try:
                    conn.rollback()
                except Exception:
                    pass
            messagebox.showerror("Erreur BDD", f"Erreur lors du recyclage : {e}")
        finally:
            if cursor:
                try:
                    cursor.close()
                except Exception:
                    pass
            if conn:
                try:
                    conn.close()
                except Exception:
                    pass
                
    # ===================== Handlers & utilitaires =====================
    
