# -*- coding: utf-8 -*-
"""
Module extrait automatiquement de Suivi_de_production_prod_V5_5_claude.py
Regroupe les méthodes liées a : TabCapa
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

from config import EXCEL_PATH, CAPA_DOSSIER_ENTRANT, CAPA_DOSSIER_EXPLOITES, CAPA_DOSSIER_KO


class TabCapaMixin:
    def setup_capa(self, frame):
        
        # Conteneur principal
        wrap = ttk.LabelFrame(frame, text="Test OK & En attente de cyclage", padding=12)
        wrap.pack(fill="both", expand=True, padx=8, pady=8)
        
        # ====== Tableau PRINCIPAL (tests OK) + mini treeview à droite ======
        table_frame = ttk.Frame(wrap)
        table_frame.pack(fill="both", expand=True, padx=8, pady=(6, 12))
    
        # --- Colonne de droite : mini treeview "N° Série" ---
        right_frame = ttk.Frame(table_frame)
        right_frame.pack(side="right", fill="y", padx=(6, 0))  # se place tout à droite
    
        self.ok_series_tree = ttk.Treeview(
            right_frame,
            columns=("N° Série",),
            show="headings",
            selectmode="browse",
            height=12  # petit format
        )
        self.ok_series_tree.heading("N° Série", text="N° Série")
        self.ok_series_tree.column("N° Série", anchor="w", width=180, stretch=True)
    
        yscroll_ok_series = ttk.Scrollbar(right_frame, orient="vertical", command=self.ok_series_tree.yview)
        self.ok_series_tree.configure(yscrollcommand=yscroll_ok_series.set)
    
        self.ok_series_tree.pack(side="left", fill="y")
        yscroll_ok_series.pack(side="right", fill="y")
    
        # --- Tableau principal (tests OK) ---
        cols_ok = ("N° Série", "Modèle", "Capacité","DCIR", "Tension de fin de test", "Emplacement")
        self.test_tree = ttk.Treeview(table_frame, columns=cols_ok, show="headings", selectmode="browse")
        for c in cols_ok:
            self.test_tree.heading(c, text=c)
            if c == "Tension de fin de test":
                width = 170
            elif c in ("Emplacement", "Modèle"):
                width = 160
            else:
                width = 140
            self.test_tree.column(c, anchor="w", width=width, stretch=True)
    
        yscroll_ok = ttk.Scrollbar(table_frame, orient="vertical", command=self.test_tree.yview)
        self.test_tree.configure(yscrollcommand=yscroll_ok.set)
    
        # ordre des pack: scrollbar (droite), table (gauche)
        yscroll_ok.pack(side="right", fill="y")
        self.test_tree.pack(side="left", fill="both", expand=True)
    
        # ====== Tableau SECONDAIRE (tests défaillants) ======
        failed_box = ttk.LabelFrame(wrap, text="Tests défaillants & Non prêtes", padding=8)
        failed_box.pack(fill="both", expand=True, padx=8, pady=(0, 12))
    
        failed_box.columnconfigure(0, weight=1)
        failed_box.columnconfigure(1, weight=1)
    
        # --- Partie gauche : Tests défaillants ---
        failed_frame = ttk.Frame(failed_box)
        failed_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
    
        cols_fail = ("N° Série", "Emplacement", "Cause")
        self.fail_tree = ttk.Treeview(failed_frame, columns=cols_fail, show="headings", selectmode="browse")
        for c in cols_fail:
            self.fail_tree.heading(c, text=c)
            width = 150 if c != "Cause" else 260
            self.fail_tree.column(c, anchor="w", width=width, stretch=True)
    
        yscroll_fail = ttk.Scrollbar(failed_frame, orient="vertical", command=self.fail_tree.yview)
        self.fail_tree.configure(yscrollcommand=yscroll_fail.set)
    
        self.fail_tree.pack(side="left", fill="both", expand=True)
        yscroll_fail.pack(side="right", fill="y")
    
        # --- Partie droite : Batteries non prêtes ---
        notready_frame = ttk.Frame(failed_box)
        notready_frame.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
    
        cols_notready = ("N° Série",)
        self.notready_tree = ttk.Treeview(
            notready_frame, columns=cols_notready, show="headings", selectmode="browse"
        )
        self.notready_tree.heading("N° Série", text="N° Série")
        self.notready_tree.column("N° Série", anchor="w", width=180, stretch=True)
    
        yscroll_notready = ttk.Scrollbar(notready_frame, orient="vertical", command=self.notready_tree.yview)
        self.notready_tree.configure(yscrollcommand=yscroll_notready.set)
    
        self.notready_tree.pack(side="left", fill="both", expand=True)
        yscroll_notready.pack(side="right", fill="y")
    
        # ====== Gros bouton ovale centré (canvas) ======
        btn_frame = ttk.Frame(wrap)
        btn_frame.pack(fill="x", pady=8)
    
        canvas = tk.Canvas(btn_frame, width=260, height=64, highlightthickness=0, bg=self.cget("background"))
        canvas.pack(pady=6)
        oval = canvas.create_oval(4, 4, 256, 60, fill="#2E62FF", outline="#1b3fb3", width=2)
        txt  = canvas.create_text(130, 32, text="Traiter les fichiers", fill="white", font=("Segoe UI", 11, "bold"))
    
        canvas.tag_bind(oval, "<Button-1>", self._on_click)
        canvas.tag_bind(txt,  "<Button-1>", self._on_click)
        
        self.afficher_numero_en_attente()
        
    def afficher_numero_en_attente(self):
        conn = self.db_manager.connect()

        try:
            modele=self.selected_model
            
            if modele[:8]=='PPTR018A':
                stage_act='capa'
                cursor = conn.cursor()
                query=self.build_stage_query_EOP(stage_act)
                param=(str(modele[:8])+'%',)
                cursor.execute(query, param)
                rows = cursor.fetchall()  
    
                # Transforme en liste simple
                liste_batteries = [str(r[0]) for r in rows]
                
            else:
                stage_act='capa'
                cursor = conn.cursor()
                query=self.build_stage_query(stage_act)
                param=(modele,)
                cursor.execute(query, param)
                rows = cursor.fetchall()  
    
                # Transforme en liste simple
                liste_batteries = [str(r[0]) for r in rows]
            
        except Exception as e:
            # Erreur SQL => échec de traitement
            messagebox.showerror('error!', f"recuperation de la liste des batteries a tester impossible : {e}")
            
        finally:
            try:
                cursor.close()
            except:
                pass
            conn.close()
            self.ok_series_tree.delete(*self.ok_series_tree.get_children())
            for batt in liste_batteries:
                self.ok_series_tree.insert("", "end", values=(batt,))
            
    def _on_click(self,event=None):
        """
        Cœur de l'onglet "capa" (test de capacité / cyclage) : dépouille tous les
        fichiers Excel de résultats de cyclage présents dans CAPA_DOSSIER_ENTRANT,
        les rapproche des batteries en attente de test, vérifie les seuils
        (feuille "Cyclage" du fichier EXCEL_PATH), met à jour la base et déplace
        les fichiers traités vers CAPA_DOSSIER_EXPLOITES / CAPA_DOSSIER_KO.

        Repères pour s'y retrouver (méthode volontairement non découpée pour ne
        pas risquer de régression sans tests) :
          - ~L.180 : récupération de la liste des batteries en attente de capa
          - ~L.220 : boucle sur les fichiers du dossier entrant, parsing du nom
                     de fichier (numéro de série, emplacement, modèle, type de
                     cyclage) et répartition dans les Treeview de l'IHM
          - ~L.244+: branche "modele_test[:8]=='PPTR018A'" (gamme PPTR018A) :
                     récupération ref_cell -> seuils -> lecture mesures -> calcul
                     des indicateurs -> décision OK/KO
          - ~L.580+: même enchaînement (récupération ref_cell -> seuils -> mesures
                     -> indicateurs) pour les autres modèles
        Les deux branches sont structurellement très proches (candidates à une
        factorisation future, cf. rapport joint), mais n'ont pas été fusionnées
        ici pour ne pas modifier un comportement métier non couvert par des tests.
        """
        
        ok_pairs = []   # [(numero_serie_batterie, chemin_fichier), ...]
        ko_files = []   # [chemin_fichier, ...]
        non_test_files=[]
        
        conn = self.db_manager.connect()

        try:
            
            modele=self.selected_model
            
            if modele[:8]=='PPTR018A':
                stage_act='capa'
                cursor = conn.cursor()
                query=self.build_stage_query_EOP(stage_act)
                param=(str(modele[:8])+'%',)
                cursor.execute(query, param)
                rows = cursor.fetchall()  
    
                # Transforme en liste simple
                liste_batteries = [str(r[0]) for r in rows]
                
            else:
                stage_act='capa'
                cursor = conn.cursor()
                query=self.build_stage_query(stage_act)
                param=(modele,)
                cursor.execute(query, param)
                rows = cursor.fetchall()  
    
                # Transforme en liste simple
                liste_batteries = [str(r[0]) for r in rows]
            
        except Exception as e:
            # Erreur SQL => échec de traitement
            messagebox.showerror('error!', f"recuperation de la liste des batteries a tester impossible : {e}")
            
        finally:
            try:
                cursor.close()
            except:
                pass
            conn.close()
        
        
        dossier_path = CAPA_DOSSIER_ENTRANT
        dossier_exploites = CAPA_DOSSIER_EXPLOITES
        dossier_ko = CAPA_DOSSIER_KO

        df_cyclage = pd.read_excel(EXCEL_PATH,sheet_name="Cyclage",header=1)
        
        for fichier in os.listdir(dossier_path):#Traite chaque fichier dans le dossier 
            if fichier.endswith((".xlsx", ".xls")):  # Vérifie si le fichier est de type Excel
                chemin_fichier = os.path.join(dossier_path, fichier)
                numero_serie_test = os.path.splitext(os.path.basename(chemin_fichier))[0]
                # Ex: "MC0002031068-4-3-7-INR18650MH1_A.0.1"
                parties = numero_serie_test.split('-', 5)
            
                # Sécurités basiques sur le nom
                if len(parties) < 5:
                    # Nom inattendu -> on ignore/passe
                    continue
            
                numero_serie_batterie = parties[0]
                emplacement = f"{parties[1]}-{parties[2]}-{parties[3]}"
                modele_test = parties[4]
                modele_test = modele_test.split("_")[0]
             
                if numero_serie_batterie not in liste_batteries:
                    
                    self.notready_tree.insert("", "end", values=(numero_serie_batterie))
                    continue
                
                type_cyclage=parties[5]
                
                if modele_test[:8] == self.selected_model[:8]:
                    
                    if modele_test[:8]=='PPTR018A':
                        
                        try:
                            
                            # --- DB: récupérer ref_cell ---
                            conn = self.db_manager.connect()
                            if not conn:
                                # Pas de connexion = échec de traitement
                                self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "conn BDD"))
                                
                                continue
                
                            try:
                                cursor = conn.cursor()
                                query = """
                                    SELECT reference_cellule
                                    FROM cellule
                                    WHERE affectation_produit = %s
                                    LIMIT 1
                                """
                                cursor.execute(query, (numero_serie_batterie,))
                                row_db = cursor.fetchone()
                            except Exception as e:
                                # Erreur SQL => échec de traitement
                                self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "traitement"))
                                print(f"[SQL] {numero_serie_batterie} -> {e}")
                               
                                row_db = None
                            finally:
                                try:
                                    cursor.close()
                                except:
                                    pass
                                conn.close()
                
                            if not row_db or not row_db[0]:
                                # Pas de ref cellule => impossible de lire les seuils
                                self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "traitement"))
                                
                                continue
                
                            ref_cell = row_db[0]
                
                            # --- Seuils (df_cyclage) ---
                            row = df_cyclage[
                                (df_cyclage["Nom_modele"] == modele_test) &
                                (df_cyclage["Ref cellule"] == ref_cell)
                            ]
                
                            if row.empty:
                                # Seuil introuvable
                                self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "traitement"))
                                
                                continue
                
                            seuils = row.iloc[0]
                            try:
                                capa_min = float(seuils["Capa mini (Ah)"])
                            except Exception:
                                capa_min = float(str(seuils["Capa mini (Ah)"]).replace(",", "."))
                            try:
                                temps_h = float(seuils["Temps de test (h)"])
                            except Exception:
                                temps_h = float(str(seuils["Temps de test (h)"]).replace(",", "."))
                            try:
                                tension_seuil = float(seuils["Tension seuil (V)"])
                            except Exception:
                                tension_seuil = float(str(seuils["Tension seuil (V)"]).replace(",", "."))
                
                            # --- Lecture mesures ---
                            
                            try:
                                DCIR_seuil = float(seuils["DCIR borne max"])
                            except Exception:
                                DCIR_seuil = float(str(seuils["DCIR borne max"]).replace(",", "."))
                            
                            """
                            # --- assurer la conformité du type cyclage ---
                            conn = self.db_manager.connect()
                            if not conn:
                                # Pas de connexion = échec de traitement
                                self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "conn BDD"))
                                
                                continue
                
                            try:
                                cursor = conn.cursor()
                                query = #rajouter triples guillemets
                                    SELECT type_cyclage
                                    FROM produit_voltr
                                    WHERE numero_serie_produit = %s
                                    LIMIT 1
                                #rajouter triples guillemets
                                cursor.execute(query, (numero_serie_batterie,))
                                type_cyclage_bdd = cursor.fetchone()
                            except Exception as e:
                                # Erreur SQL => échec de traitement
                                self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "traitement"))
                                print(f"[SQL] {numero_serie_batterie} -> {e}")
                               
                                type_cyclage_bdd = None
                            
                            finally:
                                try:
                                    cursor.close()
                                except:
                                    pass
                                conn.close()
                                
                            
                                
                            if type_cyclage != type_cyclage_bdd:
                                self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "Mauvais type cyclage"))
                                
                                continue
                                
                            """
                            try:
                                
                                if type_cyclage == '2.0':
                                    
                                    DCIR_seuil = float(seuils["DCIR_cablelong"])
                                    data = pd.read_excel(chemin_fichier, sheet_name='record')
                                    df_1_last = data[data["Step Index"] == 1].tail(1).reset_index(drop=True)
                                    df_2_first =data[data["Step Index"] == 2].head(1).reset_index(drop=True)
                                    df_2_last=data[data["Step Index"] == 2].tail(1).reset_index(drop=True)
                                    
                                    v0=df_1_last["Voltage(V)"][0]
                                    v1=df_2_last["Voltage(V)"][0]
                                    cr=df_2_first["Current(A)"][0]
                                    
                                else :
                                    
                                    # DCIR
                                    data = pd.read_excel(chemin_fichier, sheet_name='record')
                                    df_7_last = data[data["Step Index"] == 7].tail(1).reset_index(drop=True)
                                    df_8_first =data[data["Step Index"] == 8].head(1).reset_index(drop=True)
                                    df_8_last=data[data["Step Index"] == 8].tail(1).reset_index(drop=True)
                                    
                                    v0=df_7_last["Voltage(V)"][0]
                                    v1=df_8_last["Voltage(V)"][0]
                                    cr=df_8_first["Current(A)"][0]
                                
                                DCIR=abs((v0-v1)/cr)*1000
                                DCIR = float(round(DCIR, 3))
                                
                                indic_dcir = "OK" if DCIR < DCIR_seuil else "NOK"

                            except Exception as e:
                                # Si la ligne/col manque -> échec de traitement
                                self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "traitement"))
                                print(f"[STEP DCIR] {numero_serie_batterie} -> {e}")
                                
                                continue
                            
                            if type_cyclage == "2.0":
                                
                                capa_dch=0
                                tension_finale=v1
                                indic_capa = "OK" 
                                indic_v    = "OK" 
                                indic_t    = "OK" 
                                
                            else : 
                            
                                # Capacité (onglet 'step', ligne 4, col 'Capacity(Ah)')
                                step = pd.read_excel(chemin_fichier, sheet_name="step")
                                try:
                                    capa_dch = float(step["Capacity(Ah)"].iloc[3])
                                except Exception as e:
                                    # Si la ligne/col manque -> échec de traitement
                                    self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "traitement"))
                                    print(f"[STEP capa] {numero_serie_batterie} -> {e}")
                                    
                                    continue
                    
                                # Tension fin test (onglet 'record', Step Index == 4, max sur les 10 derniers points)
                                record = pd.read_excel(chemin_fichier, sheet_name="record")
                                df_dch_last = record[record["Step Index"] == 5].head(10)
                                if df_dch_last.empty or "Voltage(V)" not in df_dch_last.columns:
                                    self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "traitement"))
                                    continue
                                try:
                                    max_voltage = float(df_dch_last["Voltage(V)"].max())
                                except Exception as e:
                                    self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "traitement"))
                                    print(f"[RECORD volt] {numero_serie_batterie} -> {e}")
                                    
                                    continue
                    
                                # Temps test (dernière valeur 'Total Time')
                                if "Total Time" not in record.columns or record["Total Time"].empty:
                                    self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "traitement"))
                                    continue
                    
                                last_time_val = record["Total Time"].iloc[-1]
                                tension_finale = record["Voltage(V)"].iloc[-1]
                                # Robustifier la conversion en timedelta
                                try:
                                    # Si déjà de type timedelta/chaîne "HH:MM:SS"
                                    time_obj = pd.to_timedelta(last_time_val)
                                    if pd.isna(time_obj):
                                        raise ValueError("NaT")
                                except Exception:
                                    # Si Excel time (float en jours)
                                    try:
                                        time_obj = pd.to_timedelta(float(last_time_val), unit="D")
                                    except Exception as e:
                                        self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "traitement"))
                                        print(f"[RECORD time] {numero_serie_batterie} -> {e}")
                                        
                                        continue
                                
                             
                                # --- Indicateurs ---
                                indic_capa = "OK" if capa_dch > capa_min else "NOK"
                                indic_v    = "OK" if max_voltage < tension_seuil else "NOK"
                                indic_t    = "OK" if time_obj > pd.to_timedelta(f"{int(temps_h)} hours") else "NOK"
                            
                            if indic_capa == "OK" and indic_v == "OK" and indic_t == "OK" and indic_dcir == "OK" :
                                # Tous OK -> tree principal (5 colonnes)
                                self.test_tree.insert(
                                    "", "end",
                                    values=(
                                        numero_serie_batterie,
                                        modele_test,
                                        f"{capa_dch:.3f}",
                                        DCIR,
                                        f"{tension_finale:.3f}",
                                        emplacement
                                    )
                                )
                                ok_pairs.append((numero_serie_batterie, chemin_fichier,modele_test))
                                conn = self.db_manager.connect()
                                if not conn:
                                    messagebox.showerror("Erreur DB", "Connexion DB impossible pour l'update.")
                                    return
                            
                                try:
                                    cur = conn.cursor()
                                    
                                    request='Select valeur_test_capa from suivi_production where numero_serie_batterie =%s'
                                    param=(numero_serie_batterie,)
                                    cur.execute(request,param)
                                    result=cur.fetchone()
                                    if result is not None:
                                        cur.execute("update suivi_production set valeur_capa_ko=valeur_test_capa, date_capa_ko=date_test_capa where numero_serie_batterie =%s",(numero_serie_batterie,))
                                        
                                    sql = """
                                        UPDATE suivi_production
                                        SET valeur_test_capa = %s,
                                        DCIR = %s,
                                        date_test_capa = NOW(),
                                        date_DCIR = NOW()
                                        WHERE numero_serie_batterie = %s
                                    """
                            
                                    params = (capa_dch,DCIR,numero_serie_batterie)
                                    cur.execute(sql, params)
                                    
                                    if modele_test== 'PPTR018AC' or modele_test=='PPTR018AB':
                                        cur.execute("update produit_voltr set statut= %s where numero_serie_produit=%s",("stock",numero_serie_batterie))
                                        
                                except Exception as e:
                                    messagebox.showerror("Erreur DB", f"Update échoué : {e}")
                                finally:
                                    try:
                                        cur.close()
                                    except:
                                        pass
                                    conn.commit()
                                    conn.close()
                            else:
                                # Au moins un NOK -> tree défaillants (3 colonnes) avec cause
                                causes = []
                                if indic_capa == "NOK":
                                    causes.append("Capacité")
                                    conn = self.db_manager.connect()
                                    cur = conn.cursor()
                                    request='Select valeur_test_capa from suivi_production where numero_serie_batterie =%s'
                                    param=(numero_serie_batterie,)
                                    cur.execute(request,param)
                                    result=cur.fetchone()
                                    if result is None:
                                        sql = """
                                            UPDATE suivi_production
                                            SET valeur_test_capa = %s,
                                            DCIR = %s,
                                            date_test_capa = NOW(),
                                            date_DCIR = NOW()
                                            WHERE numero_serie_batterie = %s
                                        """
                                
                                        params = (capa_dch,DCIR,numero_serie_batterie)
                                        cur.execute(sql, params)
                                        
                                        
                                    
                                    else :
                                        cur.execute("update suivi_production set valeur_capa_ko = valeur_test_capa, date_capa_ko = date_test_capa where numero_serie_batterie = %s",(numero_serie_batterie,))
                                        sql = """
                                            UPDATE suivi_production
                                            SET valeur_test_capa = %s,
                                            DCIR = %s,
                                            date_test_capa = NOW(),
                                            date_DCIR = NOW()
                                            WHERE numero_serie_batterie = %s
                                        """
                                
                                        params = (capa_dch,DCIR,numero_serie_batterie)
                                        cur.execute(sql, params)                      
                                    
                                    conn.commit()
                                    conn.close()
                                    
                                if indic_v == "NOK":
                                    causes.append("Tension")
                                if indic_t == "NOK":
                                    causes.append("Temps")
                                if indic_dcir =="NOK":
                                    causes.append("DCIR")
                                cause_txt = " & ".join(causes) if causes else "traitement"
                                self.fail_tree.insert(
                                    "", "end",
                                    values=(numero_serie_batterie, emplacement, cause_txt)
                                )
                                ko_files.append(chemin_fichier)
                
                        except Exception as e:
                            # Toute autre erreur après le if modele_test == self.selected_model
                            self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "traitement"))
                            ko_files.append(chemin_fichier)
                            print(f"[TRAITEMENT] {numero_serie_batterie} -> {e}")
                    
                    elif modele_test==self.selected_model:
                        
                        try:
                            
                            # --- DB: récupérer ref_cell ---
                            conn = self.db_manager.connect()
                            if not conn:
                                # Pas de connexion = échec de traitement
                                self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "conn BDD"))
                                
                                continue
                
                            try:
                                cursor = conn.cursor()
                                query = """
                                    SELECT reference_cellule
                                    FROM cellule
                                    WHERE affectation_produit = %s
                                    LIMIT 1
                                """
                                cursor.execute(query, (numero_serie_batterie,))
                                row_db = cursor.fetchone()
                            except Exception as e:
                                # Erreur SQL => échec de traitement
                                self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "traitement"))
                                print(f"[SQL] {numero_serie_batterie} -> {e}")
                               
                                row_db = None
                            finally:
                                try:
                                    cursor.close()
                                except:
                                    pass
                                conn.close()
                
                            if not row_db or not row_db[0]:
                                # Pas de ref cellule => impossible de lire les seuils
                                self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "traitement"))
                                
                                continue
                
                            ref_cell = row_db[0]
                
                            # --- Seuils (df_cyclage) ---
                            row = df_cyclage[
                                (df_cyclage["Nom_modele"] == modele_test) &
                                (df_cyclage["Ref cellule"] == ref_cell)
                            ]
                
                            if row.empty:
                                # Seuil introuvable
                                self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "traitement"))
                                
                                continue
                
                            seuils = row.iloc[0]
                            try:
                                capa_min = float(seuils["Capa mini (Ah)"])
                            except Exception:
                                capa_min = float(str(seuils["Capa mini (Ah)"]).replace(",", "."))
                            try:
                                temps_h = float(seuils["Temps de test (h)"])
                            except Exception:
                                temps_h = float(str(seuils["Temps de test (h)"]).replace(",", "."))
                            try:
                                tension_seuil = float(seuils["Tension seuil (V)"])
                            except Exception:
                                tension_seuil = float(str(seuils["Tension seuil (V)"]).replace(",", "."))
                
                            # --- Lecture mesures ---
                            
                            if modele_test[:7]=="EMBR036" or modele_test=="EPDR011AA":
                                
                                # Capacité (onglet 'step', ligne 4, col 'Capacity(Ah)')
                                step = pd.read_excel(chemin_fichier, sheet_name="step")
                                try:
                                    if modele_test=="EPDR011AA":
                                        capa_dch = float(step["Capacity(Ah)"].iloc[3])
                                    else:
                                        capa_dch = float(step["Capacity(Ah)"].iloc[4])
                                except Exception as e:
                                    # Si la ligne/col manque -> échec de traitement
                                    self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "traitement"))
                                    print(f"[STEP capa] {numero_serie_batterie} -> {e}")
                                    
                                    continue
                    
                                # Tension fin test (onglet 'record', Step Index == 4, max sur les 10 derniers points)
                                record = pd.read_excel(chemin_fichier, sheet_name="record")
                                df_dch_last = record[record["Step Index"] == 6].head(10)
                                if df_dch_last.empty or "Voltage(V)" not in df_dch_last.columns:
                                    self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "traitement"))
                                    continue
                                try:
                                    max_voltage = float(df_dch_last["Voltage(V)"].max())
                                except Exception as e:
                                    self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "traitement"))
                                    print(f"[RECORD volt] {numero_serie_batterie} -> {e}")
                                    
                                    continue
                    
                                # Temps test (dernière valeur 'Total Time')
                                if "Total Time" not in record.columns or record["Total Time"].empty:
                                    self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "traitement"))
                                    continue
                    
                                last_time_val = record["Total Time"].iloc[-1]
                                tension_finale = record["Voltage(V)"].iloc[-1]
                                DCIR=0
                                indic_dcir = "OK"
                                # Robustifier la conversion en timedelta
                                try:
                                    # Si déjà de type timedelta/chaîne "HH:MM:SS"
                                    time_obj = pd.to_timedelta(last_time_val)
                                    if pd.isna(time_obj):
                                        raise ValueError("NaT")
                                except Exception:
                                    # Si Excel time (float en jours)
                                    try:
                                        time_obj = pd.to_timedelta(float(last_time_val), unit="D")
                                    except Exception as e:
                                        self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "traitement"))
                                        print(f"[RECORD time] {numero_serie_batterie} -> {e}")
                                        
                                        continue
                                
                            elif modele_test[:8]=="PPTR018A" or modele_test[:8]=="LNBR008A":
                                
                                try:
                                    DCIR_seuil = float(seuils["DCIR borne max"])
                                except Exception:
                                    DCIR_seuil = float(str(seuils["DCIR borne max"]).replace(",", "."))
                                
                                try:
                                    # DCIR
                                    data = pd.read_excel(chemin_fichier, sheet_name='record')
                                    df_7_last = data[data["Step Index"] == 7].tail(1).reset_index(drop=True)
                                    df_8_first =data[data["Step Index"] == 8].head(1).reset_index(drop=True)
                                    df_8_last=data[data["Step Index"] == 8].tail(1).reset_index(drop=True)
                                    
                                    v0=df_7_last["Voltage(V)"][0]
                                    v1=df_8_last["Voltage(V)"][0]
                                    cr=df_8_first["Current(A)"][0]
                                    
                                    DCIR=abs((v0-v1)/cr)*1000
                                    DCIR = float(round(DCIR, 3))
                                    
                                    indic_dcir = "OK" if DCIR < DCIR_seuil else "NOK"
                                    
                                    
                                except Exception as e:
                                    # Si la ligne/col manque -> échec de traitement
                                    self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "traitement"))
                                    print(f"[STEP DCIR] {numero_serie_batterie} -> {e}")
                                    
                                    continue
                                
                                
                                # Capacité (onglet 'step', ligne 4, col 'Capacity(Ah)')
                                step = pd.read_excel(chemin_fichier, sheet_name="step")
                                try:
                                    capa_dch = float(step["Capacity(Ah)"].iloc[3])
                                except Exception as e:
                                    # Si la ligne/col manque -> échec de traitement
                                    self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "traitement"))
                                    print(f"[STEP capa] {numero_serie_batterie} -> {e}")
                                    
                                    continue
                    
                                # Tension fin test (onglet 'record', Step Index == 4, max sur les 10 derniers points)
                                record = pd.read_excel(chemin_fichier, sheet_name="record")
                                df_dch_last = record[record["Step Index"] == 5].head(10)
                                if df_dch_last.empty or "Voltage(V)" not in df_dch_last.columns:
                                    self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "traitement"))
                                    continue
                                try:
                                    max_voltage = float(df_dch_last["Voltage(V)"].max())
                                except Exception as e:
                                    self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "traitement"))
                                    print(f"[RECORD volt] {numero_serie_batterie} -> {e}")
                                    
                                    continue
                    
                                # Temps test (dernière valeur 'Total Time')
                                if "Total Time" not in record.columns or record["Total Time"].empty:
                                    self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "traitement"))
                                    continue
                    
                                last_time_val = record["Total Time"].iloc[-1]
                                tension_finale = record["Voltage(V)"].iloc[-1]
                                # Robustifier la conversion en timedelta
                                try:
                                    # Si déjà de type timedelta/chaîne "HH:MM:SS"
                                    time_obj = pd.to_timedelta(last_time_val)
                                    if pd.isna(time_obj):
                                        raise ValueError("NaT")
                                except Exception:
                                    # Si Excel time (float en jours)
                                    try:
                                        time_obj = pd.to_timedelta(float(last_time_val), unit="D")
                                    except Exception as e:
                                        self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "traitement"))
                                        print(f"[RECORD time] {numero_serie_batterie} -> {e}")
                                        
                                        continue
                         
                            # --- Indicateurs ---
                            indic_capa = "OK" if capa_dch > capa_min else "NOK"
                            indic_v    = "OK" if max_voltage < tension_seuil else "NOK"
                            indic_t    = "OK" if time_obj > pd.to_timedelta(f"{int(temps_h)} hours") else "NOK"
                            
                            if indic_capa == "OK" and indic_v == "OK" and indic_t == "OK" and indic_dcir == "OK" :
                                # Tous OK -> tree principal (5 colonnes)
                                self.test_tree.insert(
                                    "", "end",
                                    values=(
                                        numero_serie_batterie,
                                        modele_test,
                                        f"{capa_dch:.3f}",
                                        DCIR,
                                        f"{tension_finale:.3f}",
                                        emplacement
                                    )
                                )
                                ok_pairs.append((numero_serie_batterie, chemin_fichier, modele_test))
                                conn = self.db_manager.connect()
                                if not conn:
                                    messagebox.showerror("Erreur DB", "Connexion DB impossible pour l'update.")
                                    return
                            
                                try:
                                    cur = conn.cursor()
                                    
                                    request='Select valeur_test_capa from suivi_production where numero_serie_batterie =%s'
                                    param=(numero_serie_batterie,)
                                    cur.execute(request,param)
                                    result=cur.fetchone()
                                    if result is not None:
                                        cur.execute("update suivi_production set valeur_capa_ko=valeur_test_capa, date_capa_ko=date_test_capa where numero_serie_batterie =%s",(numero_serie_batterie,))
   
                                    
                                    sql = """
                                        UPDATE suivi_production
                                        SET valeur_test_capa = %s,
                                        DCIR = %s,
                                        date_test_capa = NOW(),
                                        date_DCIR = NOW()
                                        WHERE numero_serie_batterie = %s
                                    """
                            
                                    params = (capa_dch,DCIR,numero_serie_batterie)
                                    cur.execute(sql, params)
                                    
                                    if modele_test== 'PPTR018AC' or modele_test=='PPTR018AB':
                                        cur.execute("update produit_voltr set statut= %s where numero_serie_produit=%s",("stock",numero_serie_batterie))
                                        
                                except Exception as e:
                                    messagebox.showerror("Erreur DB", f"Update échoué : {e}")
                                finally:
                                    try:
                                        cur.close()
                                    except:
                                        pass
                                    conn.commit()
                                    conn.close()
                            else:
                                # Au moins un NOK -> tree défaillants (3 colonnes) avec cause
                                causes = []
                                if indic_capa == "NOK":
                                    causes.append("Capacité")
                                    conn = self.db_manager.connect()
                                    cur = conn.cursor()
                                    request='Select valeur_test_capa from suivi_production where numero_serie_batterie =%s'
                                    param=(numero_serie_batterie,)
                                    cur.execute(request,param)
                                    result=cur.fetchone()
                                    if result is None:
                                        sql = """
                                            UPDATE suivi_production
                                            SET valeur_test_capa = %s,
                                            DCIR = %s,
                                            date_test_capa = NOW(),
                                            date_DCIR = NOW()
                                            WHERE numero_serie_batterie = %s
                                        """
                                
                                        params = (capa_dch,DCIR,numero_serie_batterie)
                                        cur.execute(sql, params)
                                    
                                    else :
                                        cur.execute("update suivi_production set valeur_capa_ko=valeur_test_capa, date_capa_ko=date_test_capa where numero_serie_batterie =%s",(numero_serie_batterie,))
                                        sql = """
                                            UPDATE suivi_production
                                            SET valeur_test_capa = %s,
                                            DCIR = %s,
                                            date_test_capa = NOW(),
                                            date_DCIR = NOW()
                                            WHERE numero_serie_batterie = %s
                                        """
                                
                                        params = (capa_dch,DCIR,numero_serie_batterie)
                                        cur.execute(sql, params)  
                                    
                                if indic_v == "NOK":
                                    causes.append("Tension")
                                if indic_t == "NOK":
                                    causes.append("Temps")
                                if indic_dcir =="NOK":
                                    causes.append("DCIR")
                                cause_txt = " & ".join(causes) if causes else "traitement"
                                self.fail_tree.insert(
                                    "", "end",
                                    values=(numero_serie_batterie, emplacement, cause_txt)
                                )
                                ko_files.append(chemin_fichier)
                
                        except Exception as e:
                            # Toute autre erreur après le if modele_test == self.selected_model
                            self.fail_tree.insert("", "end", values=(numero_serie_batterie, emplacement, "traitement"))
                            ko_files.append(chemin_fichier)
                            print(f"[TRAITEMENT] {numero_serie_batterie} -> {e}")
                    else:
                        # Modèle différent => on ignore ce fichier
                        pass                       
                else:
                    # Modèle différent => on ignore ce fichier
                    pass 
            
        # === Après la boucle: mise à jour BDD pour les OK ===
        ok_serials = list({s for (s, _p, _m) in ok_pairs})  # dédoublonné
        if ok_serials:
            self._update_ok_in_db(ok_serials)
    
        # === Déplacement des fichiers ===
        self._move_processed_files(ok_pairs, ko_files, dossier_exploites, dossier_ko)    
        
    def _update_ok_in_db(self, ok_serials):
        """
        Met à jour la base pour tous les numéros série OK.
        Adapte la requête SQL selon ton schéma.
        """
        conn = self.db_manager.connect()
        if not conn:
            messagebox.showerror("Erreur DB", "Connexion DB impossible pour l'update.")
            return
    
        try:
            cur = conn.cursor()
    
            sql = """
                UPDATE suivi_production
                   SET test_capa = 1,
                       date_test_capa = NOW()
                 WHERE numero_serie_batterie = %s
            """
    
            params = [(s,) for s in ok_serials]
            cur.executemany(sql, params)
            
            for s in ok_serials:
                cur.execute("Select fermeture_batt from suivi_production where numero_serie_batterie= %s",(s,))
                ress=cur.fetchall()
                etat_f=[res[0] for res in ress]
                if etat_f:
                    cur.execute("UPDATE produit_voltr SET statut = %s where numero_serie_produit= %s",('stock',s))
            conn.commit()
            print(f"Update OK sur {cur.rowcount} lignes.")
        except Exception as e:
            messagebox.showerror("Erreur DB", f"Update échoué : {e}")
        finally:
            try:
                cur.close()
            except:
                pass
            conn.close()
            self.afficher_numero_en_attente()
    
    
    def _move_processed_files(self, ok_pairs, ko_files, dossier_exploites, dossier_ko):
        """
        Déplace les fichiers:
          - OK -> dossier_exploites/<model>
          - KO/erreurs -> dossier_ko
        """
    
        os.makedirs(dossier_ko, exist_ok=True)
    
        def _safe_move(src_path, dst_dir):
            try:
                base = os.path.basename(src_path)
                dst = os.path.join(dst_dir, base)
    
                # Gestion des doublons
                if os.path.exists(dst):
                    name, ext = os.path.splitext(base)
                    i = 1
                    candidate = os.path.join(dst_dir, f"{name} ({i}){ext}")
                    while os.path.exists(candidate):
                        i += 1
                        candidate = os.path.join(dst_dir, f"{name} ({i}){ext}")
                    dst = candidate
    
                shutil.move(src_path, dst)
                return True
            except Exception as e:
                print(f"[MOVE] {src_path} -> {dst_dir} : {e}")
                return False
    
        # OK -> exploités/<model>
        for _, src, model in ok_pairs:
            dossier_exp_model = os.path.join(dossier_exploites, model)
            os.makedirs(dossier_exp_model, exist_ok=True)
            _safe_move(src, dossier_exp_model)
    
        # KO -> ko
        for src in ko_files:
            _safe_move(src, dossier_ko)
    
    
    #------------------------------ Onglet recherche -----------------------------------------------   
        
