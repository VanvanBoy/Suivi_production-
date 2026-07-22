# -*- coding: utf-8 -*-
"""
Module extrait automatiquement de Suivi_de_production_prod_V5_5_claude.py
Regroupe les méthodes liées a : TabExpedition
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


class TabExpeditionMixin:
    def setup_exp(self, frame):
        # ---------- Layout principal : gauche | droite ----------
        container = ttk.Frame(frame); container.pack(fill="both", expand=True, padx=12, pady=12)
        container.columnconfigure(0, weight=1)
        container.columnconfigure(1, weight=1)
        container.rowconfigure(0, weight=1)
    
        # ========== Colonne gauche : entrées & combos ==========
        left = ttk.LabelFrame(container, text="Expédition", padding=10)
        left.grid(row=0, column=0, sticky="nsew", padx=(0,8))
    
        # N° série produit + ajout auto si 9 chiffres
        ttk.Label(left, text="N° série produit:").pack(pady=(0,4), anchor="w")
        self.exp_numero_serie_batt_entry = ttk.Entry(left, width=30)
        self.exp_numero_serie_batt_entry.pack(pady=(0,8), fill="x")
        self.exp_numero_serie_batt_entry.bind("<KeyRelease>", self._exp_on_entry_change)
    
        ttk.Button(left, text="❌ Non conforme", command=self.add_non_conf_batterie, style="Danger.TButton").pack(pady=6, anchor="w")
    
        # Client
        ttk.Label(left, text="Client").pack(pady=(8,4), anchor="w")
        cb_var_client = tk.StringVar()
        self.cb_cl = ttk.Combobox(left, textvariable=cb_var_client, state="readonly", width=40)
        self.cb_cl.pack(fill="x")
    
        # Alimentation des clients
        conn = self.db_manager.connect()
        if conn:
            try:
                cursor = conn.cursor()
                cursor.execute("SELECT nom_client FROM client ORDER BY nom_client")
                mots = [row[0].replace(" ", "-") for row in cursor.fetchall()]
                self.cb_cl["values"] = mots
            except Exception as e:
                messagebox.showerror("Erreur SQL", f"Chargement clients :\n{e}")
            finally:
                try: cursor.close()
                except: pass
                conn.close()
    
        # Projet
        ttk.Label(left, text="Projet").pack(pady=(8,4), anchor="w")
        self.cb_pr = ttk.Combobox(left, state="readonly", width=40)
        self.cb_pr.pack(fill="x")
        
        # Alimentation des clients
        conn = self.db_manager.connect()
        if conn:
            try:
                cursor = conn.cursor()
                cursor.execute("SELECT nom_projet FROM projet ORDER BY nom_projet")
                mots = [row[0] for row in cursor.fetchall()]
                self.cb_pr["values"] = mots
            except Exception as e:
                messagebox.showerror("Erreur SQL", f"Chargement projet :\n{e}")
            finally:
                try: cursor.close()
                except: pass
                conn.close()
    
        # Modèle + checkbox d’activation
        ttk.Label(left, text="Modèle batterie").pack(pady=(8,4), anchor="w")
        cb_var_m_exp = tk.StringVar()
        self.cb_exp = ttk.Combobox(left, textvariable=cb_var_m_exp, values=getattr(self, "models", []),
                                   state="disabled", width=40)
        self.cb_exp.pack(fill="x")
    
        self.chk_exp_var = tk.BooleanVar(value=False)
        def toggle_combobox_exp():
            self.cb_exp.configure(state="readonly" if self.chk_exp_var.get() else "disabled")
            
        ttk.Checkbutton(left, text="Changer le modèle", variable=self.chk_exp_var,
                        command=toggle_combobox_exp).pack(pady=(6,0), anchor="w")
    
        # Commentaire
        ttk.Label(left, text="Commentaire").pack(pady=(8,4), anchor="w")
        self.exp_comm_entry = ttk.Entry(left)
        self.exp_comm_entry.pack(fill="x")
    
        # ========= Colonne droite : 2 Listbox (disponibles | sélectionnées) =========
        right = ttk.Frame(container)
        right.grid(row=0, column=1, sticky="nsew", padx=(8,0))
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)
        right.rowconfigure(4, weight=1)
    
        # Liste des batteries disponibles (emballées)
        ttk.Label(right, text="Batteries emballées (disponibles)").grid(row=0, column=0, sticky="w", pady=(0,4))
        self.exp_listbox_batt = tk.Listbox(right, font=('Segoe UI', 11), height=10)
        self.exp_listbox_batt.grid(row=1, column=0, sticky="nsew")
        self.exp_listbox_batt.bind("<<ListboxSelect>>", lambda e: None)  # neutre
        self.exp_listbox_batt.bind("<Double-Button-1>", self._exp_on_available_double_click)
    
        ttk.Separator(right, orient="horizontal").grid(row=2, column=0, sticky="ew", pady=10)
    
        # Sélection (avec compteur)
        header_sel = ttk.Frame(right); header_sel.grid(row=3, column=0, sticky="ew", pady=(0,4))
        ttk.Label(header_sel, text="Batteries sélectionnées").pack(side="left")
        ttk.Label(header_sel, text="Quantité:").pack(side="right")
        self._exp_count_var = tk.IntVar(value=0)
        self._exp_count_lbl = ttk.Label(header_sel, textvariable=self._exp_count_var)
        self._exp_count_lbl.pack(side="right", padx=(0,8))
    
        self.send_listbox_batt = tk.Listbox(right, font=('Segoe UI', 11), height=10)
        self.send_listbox_batt.grid(row=4, column=0, sticky="nsew")
        self.send_listbox_batt.bind("<Double-Button-1>", self._exp_on_selected_double_click)
    
        ttk.Button(right, text="✅ Contrôle OK", command=self.valider_exp, style="Good.TButton").grid(
            row=5, column=0, pady=10, sticky="e"
        )
    
        # Charge la liste des disponibles
        self.display_model_list_exp()
    
    
    def display_model_list_exp(self):
        """Alimente la listbox des batteries emballées (disponibles)."""
        conn = self.db_manager.connect()
        stage_act='exp'
        if not conn:
            return
        try:
            cursor = conn.cursor()
            query=self.build_stage_query(stage_act)
            cursor.execute(query)
            rows = cursor.fetchall()
            liste_batteries = [str(r[0]) for r in rows]
            self.exp_listbox_batt.delete(0, tk.END)
            for batt in liste_batteries:
                self.exp_listbox_batt.insert(tk.END, batt)
        except Exception as e:
            messagebox.showerror("Erreur SQL", f"Impossible de récupérer les données :\n{e}")
        finally:
            try: cursor.close()
            except: pass
            conn.close()
            
    #------------------------------ Onglet fin de ligne  -----------------------------------------------    
    def _exp_on_entry_change(self, event):
        """Ajoute la batterie dans 'sélectionnées' quand l'entrée atteint 9 chiffres."""
        txt = self.exp_numero_serie_batt_entry.get().strip()
        if len(txt) == 9:
            self._exp_add_to_selection(txt)
            self.exp_numero_serie_batt_entry.delete(0, tk.END)
    
    def _exp_on_available_double_click(self, event):
        """Double-clic sur une batterie disponible -> ajoute à la sélection."""
        sel = self.exp_listbox_batt.curselection()
        if not sel:
            return
        value = self.exp_listbox_batt.get(sel[0])
        self._exp_add_to_selection(value)
    
    def _exp_on_selected_double_click(self, event):
        """Double-clic sur une batterie sélectionnée -> la retire de la sélection."""
        sel = self.send_listbox_batt.curselection()
        if not sel:
            return
        value = self.send_listbox_batt.get(sel[0])
        # Retire la première occurrence (il n'y a pas de doublon, donc OK)
        self.send_listbox_batt.delete(sel[0])
        self._exp_update_counter()
    
    def _exp_add_to_selection(self, numero):
        """Ajoute sans doublon à la listbox sélectionnée puis met à jour le compteur."""
        # Anti-doublon
        current = set(self.send_listbox_batt.get(0, tk.END))
        if numero in current:
            return
        self.send_listbox_batt.insert(tk.END, numero)
        self._exp_update_counter()
    
    def _exp_update_counter(self):
        """Mise à jour du compteur de batteries sélectionnées."""
        self._exp_count_var.set(self.send_listbox_batt.size())
        
    def _exp_get_selected_batteries(self):
        """Retourne toutes les batteries présentes dans la listbox 'sélectionnées'."""
        return list(self.send_listbox_batt.get(0, tk.END))
    
    def exp_add_non_conf_batterie(self):
        reponse = messagebox.askyesno("Non conformité", "Ouvrir une non-conformité ?")
                
        if reponse:
        
            gg_from = NON_CONFORMITE_FORM_URL
            webbrowser.open_new_tab(gg_from)  
        
        self.exp_numero_serie_batt_entry.delete(0, tk.END)
        
        
    def valider_exp(self):
        
        # 0) Récup sélection
        numeros = list(self.send_listbox_batt.get(0, tk.END))
        if not numeros:
            messagebox.showwarning("Avertissement", "Aucune batterie sélectionnée.")
            return
        
        # 1) Référence cible (combobox si case cochée, sinon self.modele)
        target_ref = ""
        if getattr(self, "chk_exp_var", None) and self.chk_exp_var.get():
            target_ref = self.cb_exp.get().strip()
        if not target_ref:
            target_ref = getattr(self, "selected_model", "").strip()
        
        # 2) Lecture des références actuelles via IN (...)
        placeholders = ", ".join(["%s"] * len(numeros))
        conn = self.db_manager.connect()
        if not conn:
            return
        try:
            cur = conn.cursor()
        
            # -- Ajuste ce SELECT selon ton schéma (JOIN si id_produit, sinon colonne directe)
            sql_sel = f"""
                SELECT sp.numero_serie_batterie, p.reference_produit_voltr
                FROM suivi_production AS sp
                JOIN produit_voltr AS p ON p.numero_serie_produit = sp.numero_serie_batterie
                WHERE sp.numero_serie_batterie IN ({placeholders})
            """
            cur.execute(sql_sel, tuple(numeros))
            rows = cur.fetchall()
            num2ref = {str(n): str(r) for (n, r) in rows}
        
            # Vérifs de base (numéros non trouvés)
            manquants = [n for n in numeros if n not in num2ref]
            if manquants:
                messagebox.showerror("Erreur",
                                     f"Références introuvables pour : {', '.join(manquants)}")
                return
        
            # 3) Regroupement par référence
            groups = {}
            for n, ref in num2ref.items():
                groups.setdefault(ref, []).append(n)
        
            # 4) Cohérence vs target_ref
            #    - si target_ref définie : tout ce qui n'a pas cette ref est "à corriger"
            #    - si target_ref vide : on ne corrige pas, mais on exige l'homogénéité
            to_fix = []
            if target_ref:
                for ref, nums in groups.items():
                    if ref != target_ref:
                        to_fix.extend(nums)
        
                if to_fix:
                    # 5) Proposition de correction
                    if not messagebox.askyesno(
                        "Référence différente",
                        f"{len(to_fix)} batterie(s) ne sont pas en '{target_ref}'.\n"
                        "Voulez-vous les mettre toutes à cette référence ?"
                    ):
                        # L’utilisateur refuse -> on arrête le process
                        return
        
                    # 6) Mise à jour des références
                    placeholders_fix = ", ".join(["%s"] * len(to_fix))
                    # -- VERSION avec id_produit (JOIN sur reference_produit)
                    sql_upd = f"""
                        UPDATE produit_voltr 
                        SET reference_produit_voltr = %s
                        WHERE produit_voltr.numero_serie_produit IN ({placeholders_fix})
                    """
                    cur.execute(sql_upd, (target_ref, *to_fix))
                    conn.commit()
        
                    # 7) Re-lecture pour contrôle
                    cur.execute(sql_sel, tuple(numeros))
                    rows = cur.fetchall()
                    num2ref = {str(n): str(r) for (n, r) in rows}
                    groups = {}
                    for n, ref in num2ref.items():
                        groups.setdefault(ref, []).append(n)
        
            # 8) Garde-fou : exiger une seule référence finale
            if len(groups) > 1:
                refs = ", ".join(groups.keys())
                messagebox.showerror("Références multiples",
                                     f"Plusieurs références restent présentes : {refs}\n"
                                     f"Process interrompu.")
                return
        
            # 9) Suite du process OK (une seule ref, corrigée si besoin)
            ref_unique = next(iter(groups.keys()))
            numeros_final = groups[ref_unique]
            
            # 9) FLUX MÉTIER : marquer expédiées
            comment = self.exp_comm_entry.get().strip()
            placeholders_final = ", ".join(["%s"] * len(numeros))
            projet=self.cb_pr.get().strip()
            nom_client=self.cb_cl.get().strip()
            
            conn2 = self.db_manager.connect()
            if conn2:
                try:
                    cursor = conn2.cursor()
                    cursor.execute("SELECT nom_client FROM client ORDER BY nom_client")
                    mots = [row[0].replace(" ", "-") for row in cursor.fetchall()]
                    self.cb_cl["values"] = mots
                except Exception as e:
                    messagebox.showerror("Erreur SQL", f"Chargement clients :\n{e}")
                finally:
                    try: cursor.close()
                    except: pass
                    conn2.close()
            
            #Obtenir id_client
            if nom_client:
                nom_c_bdd=nom_client.replace("-"," ")
                conn2 = self.db_manager.connect()
                if conn2:
                    try:
                        cursor = conn2.cursor()
                        cursor.execute("SELECT id_client FROM client where nom_client = %s",(nom_c_bdd,))
                        id_client = [row[0] for row in cursor.fetchall()]
                    except Exception as e:
                        messagebox.showerror("Erreur SQL", f"Chargement id_client :\n{e}")
                    finally:
                        try: cursor.close()
                        except: pass
                        conn2.close()
    
            if comment:
                
                sql_sp = f"""
                    UPDATE suivi_production
                    SET expedition = 1,
                        date_expedition = NOW(),
                        commentaire = %s
                    WHERE numero_serie_batterie IN ({placeholders_final})
                """
                params_sp = (comment, *numeros)
                
            else:
                
                sql_sp = f"""
                    UPDATE suivi_production
                    SET expedition = 1,
                        date_expedition = NOW()
                    WHERE numero_serie_batterie IN ({placeholders_final})
                """
                params_sp = tuple(numeros)
    
            cur.execute(sql_sp, params_sp)
            
            set_parts = ["statut = 'expediee'"]
            params_mark = []
            
            if nom_client:
                if id_client:
                    set_parts.append("id_client = %s")
                    params_mark.append(id_client)
            
            if projet:
                set_parts.append("numero_projet = %s")
                params_mark.append(projet)
            
            sql_mark = f"""
                UPDATE produit_voltr
                SET {', '.join(set_parts)}
                WHERE numero_serie_produit IN ({placeholders_final})
            """
            
            params_mark = tuple(params_mark) + tuple(numeros)
            
            cur.execute(sql_mark, params_mark)
            
            conn.commit()
    
            # 10) UI : succès + refresh
            messagebox.showinfo("Succès", f"{len(numeros)} batterie(s) marquées expédiées.")
            self.send_listbox_batt.delete(0, tk.END)       # vide la sélection
            self._exp_update_counter()                      # remet le compteur à jour
            for f in self.funcs_to_run:
                f()              # recharge les "disponibles"
            self.exp_comm_entry.delete(0, tk.END)           # vide le commentaire
            
            
        except Exception as e:
            try:
                conn.rollback()
            except:
                pass
            messagebox.showerror("Erreur SQL", f"valider_exp :\n{e}")
        finally:
            try:
                cur.close()
                conn.close()
            except:
                pass
