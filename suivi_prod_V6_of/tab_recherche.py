# -*- coding: utf-8 -*-
"""
Module extrait automatiquement de Suivi_de_production_prod_V5_5_claude.py
Regroupe les méthodes liées a : TabRecherche
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


class TabRechercheMixin:
    def setup_recherche(self, frame):
        # ----- Layout principal : gauche | boutons | droite -----
        container = ttk.Frame(frame); container.pack(fill="both", expand=True, padx=12, pady=12)
        container.columnconfigure(0, weight=1)
        container.columnconfigure(1, weight=0)
        container.columnconfigure(2, weight=2)
        container.rowconfigure(0, weight=1)
    
        # ------- Colonne gauche : entrées + combobox + listbox -------
        left = ttk.LabelFrame(container, text="Recherche", padding=10)
        left.grid(row=0, column=0, sticky="nsew", padx=(0,8))
       
    
        ttk.Label(left, text="N° série cellule").grid(row=0, column=0, sticky="w")
        self.rech_entry_cell = ttk.Entry(left, width=28)
        self.rech_entry_cell.grid(row=1, column=0, sticky="we", pady=(0,8))
        # Remplissage auto du n° batterie quand l'entry cellule atteint 12 chars
        self.rech_entry_cell.bind("<KeyRelease>", self._rech_on_cell_entry)
    
        ttk.Label(left, text="N° série batterie").grid(row=2, column=0, sticky="w")
        self.rech_entry_batt = ttk.Entry(left, width=28)
        self.rech_entry_batt.grid(row=3, column=0, sticky="we", pady=(0,8))
    
        ttk.Label(left, text="Référence batterie").grid(row=4, column=0, sticky="w")
        # valeur par défaut nulle (vide)
        self.rech_model_var = tk.StringVar(value="")
        self.rech_combo = ttk.Combobox(left, textvariable=self.rech_model_var,
                                       values=(self.models or []), state="readonly", width=30)
        self.rech_combo.grid(row=5, column=0, sticky="we", pady=(0,8))
        self.rech_combo.bind("<<ComboboxSelected>>", lambda e: self._rech_on_model_change())
    
        ttk.Label(left, text="Liste batterie").grid(row=6, column=0, sticky="w")
        lb_frame = ttk.Frame(left); lb_frame.grid(row=7, column=0, sticky="nsew")
        left.rowconfigure(7, weight=1)
    
        # multi-sélection
        self.rech_listbox = tk.Listbox(lb_frame, height=10, activestyle="dotbox", selectmode="extended")
        yscroll = ttk.Scrollbar(lb_frame, orient="vertical", command=self.rech_listbox.yview)
        self.rech_listbox.configure(yscrollcommand=yscroll.set)
        self.rech_listbox.pack(side="left", fill="both", expand=True)
        yscroll.pack(side="right", fill="y")
    
        # Double-clic => déplacer à droite
        self.rech_listbox.bind("<Double-1>", lambda e: self._rech_move_right())
    
        # --------- Colonne boutons centraux ----------
        mid = ttk.Frame(container); mid.grid(row=0, column=1, sticky="ns")
        for i in range(3): mid.rowconfigure(i, weight=1)
        ttk.Button(mid, text="→", width=3, command=self._rech_move_right).grid(row=0, column=0, pady=4)
        ttk.Button(mid, text="←", width=3, command=self._rech_remove_selected_right).grid(row=1, column=0, pady=4)
    
        # --------- Colonne droite : table dynamique ----------
        right = ttk.LabelFrame(container, text="Sélection / Détails", padding=10)
        right.grid(row=0, column=2, sticky="nsew", padx=(8,0))
        right.rowconfigure(1, weight=1)
        right.columnconfigure(0, weight=1)
        
        # largeur fixe (par ex. 500 px, ajuste comme tu veux)
        right.configure(width=800)
        right.grid_propagate(False)   # bloque l’expansion auto
        
        self.rech_right_title = ttk.Label(right, text="", font=("", 10, "bold"))
        self.rech_right_title.grid(row=0, column=0, sticky="w", pady=(0,6))
    
        tv_frame = ttk.Frame(right); tv_frame.grid(row=1, column=0, sticky="nsew")
        # + scroll vertical & horizontal
        self.rech_tree = ttk.Treeview(tv_frame, columns=(), show="headings", selectmode="extended")
        y2 = ttk.Scrollbar(tv_frame, orient="vertical", command=self.rech_tree.yview)
        x2 = ttk.Scrollbar(tv_frame, orient="horizontal", command=self.rech_tree.xview)
        self.rech_tree.configure(yscrollcommand=y2.set, xscrollcommand=x2.set)
        self.rech_tree.grid(row=0, column=0, sticky="nsew")
        y2.grid(row=0, column=1, sticky="ns")
        x2.grid(row=1, column=0, sticky="ew")
        tv_frame.rowconfigure(0, weight=1); tv_frame.columnconfigure(0, weight=1)
    
        # set utilisé pour éviter les doublons à droite (clé = numero_serie_batterie)
        self._rech_right_keys = set()
        
    def _rech_on_cell_entry(self, event=None):
        """Quand l'entry cellule atteint 12 chars, on cherche la batterie associée et on remplit l'entry batterie."""
        numero_serie_cell = self.rech_entry_cell.get().strip()
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
            self.rech_entry_batt.delete(0, tk.END)
            if row and row[0]:
                self.rech_entry_batt.insert(0, str(row[0]))
        except Exception as e:
            messagebox.showerror("Erreur SQL", f"Lookup cellule→batterie impossible :\n{e}")
        finally:
            try: cur.close()
            except: pass
            conn.close()
    
    def _rech_on_model_change(self):
        """Quand on choisit un modèle, on alimente la liste des n° batteries via la jointure demandée."""
        ref = self.rech_model_var.get().strip()
        self.rech_listbox.delete(0, tk.END)
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
                ORDER BY sp.numero_serie_batterie
            """)
            cur.execute(sql, (ref,))
            for (num_batt,) in cur.fetchall():
                if num_batt:
                    self.rech_listbox.insert(tk.END, str(num_batt))
            self.rech_right_title.config(text=f"Modèle sélectionné : {ref}")
        except Exception as e:
            messagebox.showerror("Erreur SQL", f"Chargement liste batteries impossible :\n{e}")
        finally:
            try: cur.close()
            except: pass
            conn.close()
    
    def _rech_move_right(self):
        """Ajoute à droite : 1) le n° saisi dans l’entry batterie (si présent),
        2) tous les n° sélectionnés dans la liste. Charge les LIGNES de suivi_production correspondantes."""
        # rassembler les cibles
        targets = set()
        batt_from_entry = self.rech_entry_batt.get().strip()
        if batt_from_entry:
            targets.add(batt_from_entry)
        for idx in self.rech_listbox.curselection():
            targets.add(self.rech_listbox.get(idx))
    
        if not targets:
            return
    
        # Charger les lignes depuis suivi_production
        rows, colnames = self._rech_load_suivi_rows(list(targets))
        if not rows:
            return
    
        # Configurer le tree si c'est la 1ère fois ou si colonnes différentes
        self._rech_configure_tree_for_columns(colnames)
    
        # Insérer sans doublon (clé = numero_serie_batterie)
        try:
            k_idx = colnames.index("numero_serie_batterie")  # l’utilisateur veut dédupliquer sur ce champ
        except ValueError:
            # si absent (peu probable), on déduplique sur la 1ère colonne
            k_idx = 0
    
        added = 0
        for r in rows:
            key = str(r[k_idx]) if r[k_idx] is not None else ""
            if key and key not in self._rech_right_keys:
                self.rech_tree.insert("", tk.END, values=[("" if v is None else v) for v in r])
                self._rech_right_keys.add(key)
                added += 1
    
        if added == 0:
            self.rech_right_title.config(text="Aucun nouvel élément (déduplication active)")
    
    def _rech_remove_selected_right(self):
        """Supprimer les lignes sélectionnées dans la table de droite."""
        sel = self.rech_tree.selection()
        if not sel:
            return
        # identifier l'index de numero_serie_batterie pour nettoyer le set
        cols = self.rech_tree.cget("columns")
        try:
            k_idx = cols.index("numero_serie_batterie")
        except ValueError:
            k_idx = 0
        for iid in sel:
            vals = self.rech_tree.item(iid, "values")
            # protéger si vide
            if vals:
                key = str(vals[k_idx])
                if key in self._rech_right_keys:
                    self._rech_right_keys.remove(key)
            self.rech_tree.delete(iid)
    
    def _rech_load_suivi_rows(self, numero_serie_batteries):
        """Retourne (rows, colnames) pour les lignes de suivi_production filtrées par n° série batterie."""
        if not numero_serie_batteries:
            return [], []
        conn = self.db_manager.connect()
        if not conn:
            return [], []
        try:
            cur = conn.cursor()
            # On charge TOUTE la ligne de suivi_production (c’est ce que tu as demandé)
            # IN sécurisé : on fabrique la liste de %s
            placeholders = ", ".join(["%s"] * len(numero_serie_batteries))
            sql = f"SELECT * FROM suivi_production WHERE numero_serie_batterie IN ({placeholders})"
            cur.execute(sql, tuple(numero_serie_batteries))
            rows = cur.fetchall()
            colnames = [d[0] for d in cur.description]
            return rows, colnames
        except Exception as e:
            messagebox.showerror("Erreur SQL", f"Chargement suivi_production impossible :\n{e}")
            return [], []
        finally:
            try: cur.close()
            except: pass
            conn.close()
    
    def _rech_configure_tree_for_columns(self, colnames):
        """Configure le Treeview (colonnes, largeur, ancrage, headings) + stretch + scroll horiz/vert déjà mis côté UI."""
        # si colonnes déjà identiques, ne rien faire
        current = self.rech_tree.cget("columns")
        if tuple(current) == tuple(colnames):
            return
    
        # reset
        for c in current:
            try: self.rech_tree.heading(c, text="")
            except: pass
        self.rech_tree.delete(*self.rech_tree.get_children())
        self._rech_right_keys.clear()
    
        self.rech_tree["columns"] = colnames
        # headings & options
        for c in colnames:
            self.rech_tree.heading(c, text=c)
            # largeur heuristique : min 100, sinon ~len*9 px
            width = max(100, min(380, int(len(c) * 9)))
            self.rech_tree.column(c, width=width, minwidth=80, stretch=True, anchor="w")

    
    #------------------------------ Onglet expedition -----------------------------------------------   

