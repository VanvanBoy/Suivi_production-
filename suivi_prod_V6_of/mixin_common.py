# -*- coding: utf-8 -*-
"""
Module extrait automatiquement de Suivi_de_production_prod_V5_5_claude.py
Regroupe les méthodes liées a : Common
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


class CommonMixin:
    def on_close(self):
        self._cancel_tab_refresh()
        self.destroy()
        
    def get_db_credentials(self):
        # Fonction pour obtenir les informations d'identification de l'utilisateur
        user = simpledialog.askstring("Login", "Enter your MySQL username:")
        password = simpledialog.askstring("Login", "Enter your MySQL password:", show='*')
        return user, password
    
    def _required_previous_dbcols(self,current_stage: str):
        """
        Retourne la liste des NOMS DE COLONNES MySQL (dans suivi_production)
        à vérifier (=1) avant de valider current_stage.
    
        On s'appuie sur self.stage_order (rangs > 0 = étape active)
        et sur STAGE_TO_DBCOL pour la correspondance.
        """
        if not hasattr(self, "stage_order") or not self.stage_order:
            return []
    
        if current_stage not in self.ALLOWED_STAGE_KEYS:
            return []
    
        cur_rank = self.stage_order.get(current_stage, 0)
        if cur_rank <= 0:
            return []
    
        prev_pairs = []
        for k, v in self.stage_order.items():
            if k in self.ALLOWED_STAGE_KEYS and v > 0 and v < cur_rank:
                dbcol = self.STAGE_TO_DBCOL.get(k)
                if dbcol:
                    prev_pairs.append((k, v, dbcol))
    
        # tri pour lisibilité/diagnostic
        prev_pairs.sort(key=lambda x: x[1])
    
        # on retourne uniquement les colonnes DB
        return [dbcol for (_, _, dbcol) in prev_pairs]
    
    
    def _check_prereqs_and_warn(self, num_batt: str, current_stage: str) -> bool:
        """
        Vérifie dans suivi_production que TOUTES les colonnes prérequis (=1)
        sont validées pour num_batt, d'après le mapping STAGE_TO_DBCOL.
        """
        prev_dbcols = self._required_previous_dbcols(current_stage)
        if not prev_dbcols:
            return True  # pas de prérequis
    
        conn = self.db_manager.connect()
        if not conn:
            return False
        try:
            cursor = conn.cursor(dictionary=True)
    
            # Sécurisation : on n'insère que des noms de colonnes issus de STAGE_TO_DBCOL
            cols_sql = ", ".join(prev_dbcols)
            sql = f"""
                SELECT {cols_sql}
                FROM suivi_production
                WHERE numero_serie_batterie = %s
                LIMIT 1
            """
            cursor.execute(sql, (num_batt,))
            row = cursor.fetchone()
            if not row:
                messagebox.showwarning("Vérification", "Sélectionner une batterie")
                return False
    
            missing = [c for c in prev_dbcols if (row.get(c) or 0) != 1]
            if missing:
                lis = ", ".join(missing)
                messagebox.showwarning(
                    "Pré-requis manquants",
                    f"Impossible de valider '{current_stage}'. Étape(s) non validée(s) : {lis}"
                )
                return False
    
            return True
        except Exception as e:
            messagebox.showerror("SQL", f"Erreur vérification prérequis:\n{e}")
            return False
        finally:
            try:
                cursor.close(); conn.close()
            except:
                pass
                
    def _show_model_selector_and_build(self):
        try:
            df = pd.read_excel(EXCEL_PATH,sheet_name="Flux test modele")
        except Exception as e:
            messagebox.showerror("Excel", f"Impossible de lire l'Excel:\n{e}")
            self.destroy()
            return

        if "nom_modele" not in df.columns:
            messagebox.showerror("Excel", "La colonne 'nom_modele' est absente du fichier.")
            self.destroy()
            return

        self.models = df["nom_modele"].dropna().astype(str).unique().tolist()
        if not self.models:
            messagebox.showerror("Excel", "Aucune valeur dans 'nom_modele'.")
            self.destroy()
            return

        dlg = tk.Toplevel(self)
        dlg.title("Choisir la référence batterie")
        dlg.geometry("420x160")
        dlg.transient(self)
        dlg.grab_set()

        ttk.Label(dlg, text="Référence batterie :").pack(pady=(18, 6))
        cb_var = tk.StringVar()
        cb = ttk.Combobox(dlg, textvariable=cb_var, values=self.models, state="readonly", width=40)
        cb.pack()
        cb.current(0)

        def on_launch():
            
            ref = cb_var.get()
            row = df.loc[df["nom_modele"].astype(str) == ref]
            if row.empty:
                messagebox.showerror("Sélection", "Référence introuvable.")
                # return

            column_to_stage = {
                "picking": "picking",
                "soudure_pack": "pack",
                "soudure_nappe": "nappe",
                "soudure_bms": "bms",
                "wrap": "wrap",
                "fermeture": "fermeture_batt",
                "test_capa": "capa",
                "fin_ligne": "fin_ligne",
                "emballage": "emb",
                "expedition": "exp",
                "recherche": "recherche",
                "recyclage": "recyclage",
                "tri_test" : "tri_test",
                "banc_somfy" : "banc_somfy"
            }

            stage_order = {}
            for col, key in column_to_stage.items():
                if col in df.columns:
                    try:
                        val = int(row.iloc[0][col])
                    except Exception:
                        # NaN / texte -> traite comme 0
                        val = 0
                    stage_order[key] = val

            self.selected_model = ref
            self.stage_order = stage_order
            
            dlg.destroy()

        ttk.Button(dlg, text="Lancer l'application", style="Good.TButton", command=on_launch).pack(pady=16)

        # centre la boîte
        self.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - 210
        y = self.winfo_y() + (self.winfo_height() // 2) - 80
        dlg.geometry(f"+{x}+{y}")

        self.wait_window(dlg)

        if self.stage_order is None:
            
            self.destroy()
            return

        
        self._create_widgets_with_order()
        
        stage_to_func = {
            'picking': self.display_model_list,
            'pack':  self.display_model_list_pack,
            'bms':  self.display_model_list_bms,
            'fermeture_batt':  self.display_model_list_fermeture,
            'emb':  self.display_model_list_emballage,
            'exp':  self.display_model_list_exp,
            'nappe':  self.display_model_list_nappe,
            'wrap':  self.display_model_list_wrap,
            'fin_ligne': self.display_model_list_fl, 
        }
        
        self.funcs_to_run= [
            stage_to_func[k]
            for k, v in sorted(self.stage_order.items(), key=lambda kv: kv[1])
            if v > 0 and k in stage_to_func
        ]
        
        for f in self.funcs_to_run:
            print(f.__name__)
            
    
    def verif_etape_act_non_ok(self,current_stage,num_prod):
        
        if current_stage not in self.ALLOWED_STAGE_KEYS:
            return False
        
        dbcol = self.STAGE_TO_DBCOL[current_stage]
        
        query=f"select {dbcol} from suivi_production where numero_serie_batterie =%s"
        param=(num_prod,)
        
        conn=self.db_manager.connect()
        cursor=conn.cursor()
        cursor.execute(query,param)
        
        bool_stage=cursor.fetchone()[0]
        if bool_stage:
            bool_stage=int(bool_stage)
            
        if bool_stage==1:
            messagebox.showerror('Etape deja validée',"Etape deja validée !")
            return False
        else :
            return True
        

    def _create_widgets_with_order(self):
        
        self.title("Suivi de production - "+str(self.selected_model))
        stage_defs = {
            "picking":   ("Contrôle de picking", self.setup_picking),
            "pack":      ("Contrôle soudure pack", self.setup_pack),
            "nappe":     ("Contrôle soudure nappe", self.setup_nappe),
            "bms":       ("Contrôle soudure BMS", self.setup_bms),
            "wrap":      ("Contrôle wrap", self.setup_wrap),
            "fermeture_batt": ("Contrôle fermeture", self.setup_fermeture),
            "capa":      ("Test de capacité", self.setup_capa),
            "fin_ligne": ("Test fin de ligne", self.setup_finligne),
            "emb":       ("Contrôle emballage", self.setup_emb),
            "exp":       ("Contrôle expédition", self.setup_exp),
            "recherche": ("Recherche de batterie", self.setup_recherche),
            "recyclage": ("Gestion recyclage", self.setup_recyclage),
            "tri_test": ("Tri test", self.setup_tri_test),
            "banc_somfy": ("Banc somfy", self.setup_banc_somfy)
        }

        
        ordered_keys = [
            k for k, v in sorted(self.stage_order.items(), key=lambda kv: kv[1])
            if v and v > 0 and k in stage_defs
        ]

        if not ordered_keys:
            messagebox.showwarning("Configuration", "Aucun onglet actif pour cette référence.")
          
            ordered_keys = ["picking"]
        
        self.ordered_keys = ordered_keys

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill="both")
        
        self.tab_to_stage = {}   # tab_id (str) -> 'picking' | 'pack' | ...
        self.stage_refreshers = {}  # 'picking' -> fonction reload

        for key in ordered_keys:
            title, setup_fn = stage_defs[key]
            print(title)
            frame = ttk.Frame(self.notebook)
            self.notebook.add(frame, text=title)
            focus_widget=setup_fn(frame)
            if focus_widget is not None:
                self.register_focus_target(key, focus_widget)
            # mémorise le mapping tab -> stage
            self.tab_to_stage[str(frame)] = key
    
        # mappe les fonctions de refresh (remplace par tes vraies fonctions)
        self.stage_refreshers = {
            'picking':         self.display_model_list,
            'pack':            self.display_model_list_pack,
            'nappe':           self.display_model_list_nappe,
            'bms':             self.display_model_list_bms,
            'wrap':            self.display_model_list_wrap,
            'fermeture_batt':  self.display_model_list_fermeture,
            'capa':            self.afficher_numero_en_attente,
            'fin_ligne':       self.display_model_list_fl,
        }
    
        # Bind: lorsque l’onglet change → reload immédiat + restart timer
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
    
        # Premier reload immédiat, puis démarrage du cycle 10 s
        self._refresh_active_tab_now()
        self._schedule_next_tab_tick()
        self.after(0, self._focus_active_tab)

        # Registre stage -> champ "n° série produit" de l'onglet, utilisé par
        # le mécanisme d'avancement OF (check_entry_length_batt / of_avancement).
        mapping = {
            "pack":           "s_numero_serie_batt_entry",
            "picking":        "numero_serie_batt_entry",
            "bms":            "b_numero_serie_batt_entry",
            "nappe":          "n_numero_serie_batt_entry",
            "wrap":           "w_numero_serie_batt_entry",
            "fermeture_batt": "f_numero_serie_batt_entry",
            "fin_ligne":      "fl_numero_serie_batt_entry",
        }

        self.entry_widgets = {
            key: getattr(self, attr)
            for key, attr in mapping.items()
            if key in self.ordered_keys and hasattr(self, attr)
        }
    
    def register_focus_target(self, stage: str, widget):
        """Enregistre le widget qui doit recevoir le focus au changement d'onglet."""
        if not hasattr(self, "focus_targets"):
            self.focus_targets = {}
        if widget is not None:
            self.focus_targets[stage] = widget
            
    def make_tab_chain(self, widgets, submit_button=None, ring=True, enter_from_fields=False):
        """
        Définit l'ordre Tab explicite et fait en sorte que seul le bouton
        déclenche l'action quand il est focalisé.
        - widgets : liste ordonnée des widgets (Entry/Combobox/Button/…).
                    Si tu veux que Tab atteigne le bouton, mets-le en dernier.
        - submit_button : ttk.Button (optionnel). Seul le bouton recevra <Return>.
        - ring : True => Tab boucle sur le 1er widget
        - enter_from_fields : si True, <Return> sur les champs déclenchera submit_button
                              (NE PAS utiliser si dangereux). Par défaut False.
        """
        if not widgets:
            return
    
        # assure takefocus sur les widgets inclus
        for w in widgets:
            try:
                w['takefocus'] = True
            except Exception:
                pass
    
        def next_idx(i):
            return (i + 1) % len(widgets) if ring else min(i + 1, len(widgets) - 1)
        def prev_idx(i):
            return (i - 1) % len(widgets) if ring else max(i - 1, 0)
    
        for i, w in enumerate(widgets):
            # handlers capturant i (closure sûre grâce à i=i)
            def go_next(e, i=i):
                widgets[next_idx(i)].focus_set()
                return "break"
            def go_prev(e, i=i):
                widgets[prev_idx(i)].focus_set()
                return "break"
    
            # Bind Tab / Shift-Tab (généralement suffisants)
            w.bind("<Tab>", go_next)
            w.bind("<Shift-Tab>", go_prev)
    
            # Certains environnements linux/old-tk envoient ISO_Left_Tab pour Shift-Tab.
            # On essaye de binder mais on ignore proprement l'erreur si le keysym n'existe pas.
            try:
                w.bind("<ISO_Left_Tab>", go_prev)
            except tk.TclError:
                # keysym non supporté sur cette plateforme : on ignore silencieusement
                pass
    
            # IMPORTANT: on NE bind pas <Return> sur les champs par défaut (sécurité)
            if enter_from_fields and submit_button is not None:
                w.bind("<Return>", lambda e, b=submit_button: b.invoke())
    
        # Sur le bouton : Enter et Espace déclenchent l'action (mais seulement si le bouton a le focus)
        if submit_button is not None:
            try:
                submit_button['takefocus'] = True
            except Exception:
                pass
            # binding sûr sur le bouton lui-même
            submit_button.bind("<Return>", lambda e, b=submit_button: b.invoke())
            submit_button.bind("<space>", lambda e, b=submit_button: b.invoke())
        
    
    def _get_active_stage(self) -> str | None:
        if not hasattr(self, "notebook"):
            return None
        tab_id = self.notebook.select()
        return self.tab_to_stage.get(tab_id)
    
    def _refresh_active_tab_now(self):
        if self._refreshing:
            return
        stage = self._get_active_stage()
        if not stage:
            return
        ref_fn = self.stage_refreshers.get(stage)
        if not callable(ref_fn):
            return
        self._refreshing = True
        try:
            ref_fn()  # ta fonction qui va en BDD et recharge le Treeview de l’onglet
        except Exception as e:
            print(f"[refresh:{stage}] {e}")
        finally:
            self._refreshing = False
    
    def _on_tab_changed(self, event=None):
        self._cancel_tab_refresh()
        self._refresh_active_tab_now()   # reload immédiat en arrivant sur le nouvel onglet
        self._focus_active_tab()         # <<< donne le focus dans le nouvel ongle
        self._schedule_next_tab_tick()   # redémarre le timer
    
    def _tab_tick(self):
        self._refresh_active_tab_now()
        self._schedule_next_tab_tick()
    
    def _schedule_next_tab_tick(self):
        # petit jitter pour éviter que 10 postes tapent la DB exactement en même temps
        import random
        jitter = random.randint(0, 1500)
        self._tab_refresh_job = self.after(self.refresh_ms + jitter, self._tab_tick)
    
    def _cancel_tab_refresh(self):
        if self._tab_refresh_job is not None:
            try:
                self.after_cancel(self._tab_refresh_job)
            except Exception:
                pass
            self._tab_refresh_job = None
            
    def _find_first_input(self, container):
        """Fallback: cherche récursivement le premier Entry/Combobox/Text dans un frame."""
        for w in container.winfo_children():
            # si c'est un champ saisissable
            if isinstance(w, (ttk.Entry, tk.Entry, ttk.Combobox, tk.Text)):
                return w
            # si c'est un conteneur, on descend
            if isinstance(w, (ttk.Frame, tk.Frame, ttk.LabelFrame, ttk.Labelframe, ttk.Panedwindow)):
                found = self._find_first_input(w)
                if found:
                    return found
        return None
    
    def _focus_active_tab(self):
        """Donne le focus au widget cible de l'onglet actif (ou fallback)."""
        stage = self._get_active_stage()
        if not stage or not hasattr(self, "notebook"):
            return
    
        tab_id = self.notebook.select()
        tab = self.nametowidget(tab_id)
    
        target = self.focus_targets.get(stage)
        if target is None or not target.winfo_exists():
            target = self._find_first_input(tab)
    
        if target and target.winfo_exists():
            # after(0) pour que l'onglet ait fini de s'afficher
            self.after(0, lambda: (
                target.focus_set(),
                # si Entry/Combobox: on met le curseur à la fin + on sélectionne tout (optionnel)
                (hasattr(target, "icursor") and target.icursor('end')),
                (hasattr(target, "selection_range") and target.selection_range(0, 'end'))
            ))

    
    def set_photo(self, label: tk.Label, chemin_image: str, size=(200, 200)):
        """Charge une image et l'affecte AU SEUL label donné."""
        try:
            img = Image.open(chemin_image)
            img = img.resize(size)
            photo = ImageTk.PhotoImage(img)
            label.config(image=photo, text="")
            label.image = photo  # garder une référence pour éviter GC
        except Exception as e:
            messagebox.showerror("Erreur image", f"Impossible de charger l'image : {e}")
    
    def convert_comma_to_dot(self,event):
        # Fonction pour convertir les virgules en points dans la zone de saisie
        widget = event.widget
        text = widget.get()
        text = text.replace(',', '.')
        widget.delete(0, tk.END)
        widget.insert(0, text)
            
    def verfier_coherence_ref(self,num_batt):
        modele=self.selected_model
        conn = self.db_manager.connect()
        if not conn:
            return
        try:
            cursor = conn.cursor()
            query = "Select reference_produit_voltr from produit_voltr where numero_serie_produit = %s "
            param = (num_batt,)
            cursor.execute(query, param)
            modele_act=cursor.fetchone()[0]
            if modele_act != modele:
                reponse=messagebox.askyesno("Modèle de batterie différent",f"Le mdodèle de batterie n'est pas coherent \n Passer du modèle {modele_act} au modèle {modele} pour la batterie {num_batt} ?")
                if not reponse:
                    return 'stop'
                else :
                    cursor.execute("UPDATE produit_voltr SET reference_produit_voltr = %s WHERE numero_serie_produit =%s",(modele,num_batt))
                    messagebox.showinfo("Nouveau modèle", f'la batterie {num_batt} est passé au modèle {modele}.')
                    conn.commit()
                    return 'next'
        except Exception as e:
            messagebox.showerror("Erreur SQL", f"Impossible de récupérer les données :\n{e}")
        finally:
            try:
                cursor.close()
            except:
                pass
            conn.close()  
    
    def changer_ref_batterie(self,new_ref,num_batt):
        modele=self.selected_model
        if modele == new_ref:
            messagebox.showinfo("Réference identique","Le nouveau modele est identique au modele actuel")
            return 
        else :
            conn = self.db_manager.connect()
            if not conn:
                return
            try:
                cursor = conn.cursor()
                cursor.execute("UPDATE produit_voltr SET reference_produit_voltr = %s WHERE numero_serie_produit =%s",(new_ref,num_batt))
                messagebox.showinfo("Nouveau modèle", f'la batterie {num_batt} est passé au modèle {new_ref}.')
                conn.commit()
            except Exception as e:
                messagebox.showerror("Erreur SQL", f"Impossible de récupérer les données :\n{e}")
            finally:
                try:
                    cursor.close()
                except:
                    pass
                conn.close()  
                
    def build_stage_query(self, current_stage):
        # Trouver l’index de l’étape courante
        idx = self.ordered_keys.index(current_stage)
    
        # Étapes précédentes
        previous_stages = self.ordered_keys[:idx]
    
        # Colonnes SQL associées
        prev_cols = [self.STAGE_TO_DBCOL[s] for s in previous_stages]
        current_col = self.STAGE_TO_DBCOL[current_stage]
    
        # Condition : toutes les étapes précédentes doivent être validées (=1)
        prev_conditions = " AND ".join([f"sp.{col} = 1" for col in prev_cols]) if prev_cols else "1=1"
    
        # Condition : l’étape courante doit être non validée (=0 ou NULL)
        current_condition = f"(sp.{current_col} = 0 OR sp.{current_col} IS NULL)"
    
        if current_stage != 'exp':
            # Construire la requête
            query = f"""
                SELECT sp.numero_serie_batterie
                FROM suivi_production AS sp
                JOIN produit_voltr AS pv
                  ON sp.numero_serie_batterie = pv.numero_serie_produit
                WHERE {prev_conditions}
                  AND {current_condition}
                  AND pv.reference_produit_voltr = %s
            """
        else: 
            # Construire la requête (pas de filtre ref produit)
            query = f"""
                SELECT sp.numero_serie_batterie
                FROM suivi_production AS sp
                JOIN produit_voltr AS pv
                  ON sp.numero_serie_batterie = pv.numero_serie_produit
                WHERE {prev_conditions}
                  AND {current_condition}
                  AND (recyclage=0 or recyclage is null)
            """
            
        return query
    
    def build_stage_query_EOP(self, current_stage):
        # Trouver l’index de l’étape courante
        idx = self.ordered_keys.index(current_stage)
    
        # Étapes précédentes
        previous_stages = self.ordered_keys[:idx]
    
        # Colonnes SQL associées
        prev_cols = [self.STAGE_TO_DBCOL[s] for s in previous_stages]
        current_col = self.STAGE_TO_DBCOL[current_stage]
    
        # Condition : toutes les étapes précédentes doivent être validées (=1)
        prev_conditions = " AND ".join([f"sp.{col} = 1" for col in prev_cols]) if prev_cols else "1=1"
    
        # Condition : l’étape courante doit être non validée (=0 ou NULL)
        current_condition = f"(sp.{current_col} = 0 OR sp.{current_col} IS NULL)"
    
        if current_stage != 'exp':
            # Construire la requête
            query = f"""
                SELECT sp.numero_serie_batterie
                FROM suivi_production AS sp
                JOIN produit_voltr AS pv
                  ON sp.numero_serie_batterie = pv.numero_serie_produit
                WHERE {prev_conditions}
                  AND {current_condition}
                  AND pv.reference_produit_voltr like %s
            """
        else: 
            # Construire la requête (pas de filtre ref produit)
            query = f"""
                SELECT sp.numero_serie_batterie
                FROM suivi_production AS sp
                JOIN produit_voltr AS pv
                  ON sp.numero_serie_batterie = pv.numero_serie_produit
                WHERE {prev_conditions}
                  AND {current_condition}
                  AND (recyclage=0 or recyclage is null)
            """
            
        return query
                
    #------------------------------ Helpers génériques communs aux onglets de contrôle -----------------

    def _lookup_batterie_from_cellule(self, cell_entry, batt_entry, expected_len=12):
        """
        Factorisation de la logique "X_check_entry_length" répétée à l'identique
        (à la nomenclature des widgets près) dans les onglets picking, pack,
        nappe, bms, wrap, fermeture, emballage et fin de ligne.

        Dès que `cell_entry` contient un numéro de série de cellule complet
        (longueur `expected_len`, 12 par défaut), on va chercher en base la
        batterie (numero_serie_produit / affectation_produit) à laquelle cette
        cellule est affectée, et on remplit `batt_entry` avec ce numéro.

        cell_entry : widget Entry contenant le numéro de série de la cellule.
        batt_entry : widget Entry à remplir avec le numéro de série batterie trouvé.
        """
        numero_serie_cell = cell_entry.get().strip()
        if len(numero_serie_cell) != expected_len:
            return

        conn = self.db_manager.connect()
        if not conn:
            return
        try:
            cursor = conn.cursor()
            query = "SELECT affectation_produit FROM cellule WHERE numero_serie_cellule = %s"
            cursor.execute(query, (numero_serie_cell,))
            row = cursor.fetchone()
            if row and row[0]:
                batt_entry.delete(0, tk.END)
                batt_entry.insert(0, str(row[0]))
            else:
                # Efface si pas trouvé (optionnel)
                batt_entry.delete(0, tk.END)
        finally:
            try:
                cursor.close()
            except Exception:
                pass
            conn.close()

    def _fill_entry_from_listbox_selection(self, listbox, batt_entry):
        """
        Factorisation de la logique "X_on_select_batt" répétée à l'identique
        dans les onglets picking, pack, nappe, bms, wrap, fermeture, emballage
        et fin de ligne : quand on clique sur une ligne de la listbox des
        batteries du modèle, on recopie la valeur sélectionnée dans l'Entry
        "numéro de série produit" de l'onglet.
        """
        selection = listbox.curselection()
        if not selection:
            return  # rien de sélectionné

        selected_value = listbox.get(selection[0])
        batt_entry.delete(0, tk.END)
        batt_entry.insert(0, selected_value)

    #------------------------------ Ordres de fabrication (OF) --------------------------------------

    def show_of_in_process(self):
        """
        Retourne la liste des n° d'OF pertinents pour l'onglet en cours :
        - uniquement les OF dont la référence batterie (ref_of.reference_batterie)
          correspond à la référence actuellement sélectionnée (self.selected_model),
        - en excluant les OF déjà marqués 'terminé' (ref_of.etat_fabrication).

        Cas particulier PPTR018A (Bosch/Makita/Ryobi/...) : à l'étape pack, la
        variante finale (AA/AB/AC/AD) n'est choisie qu'au moment de la
        validation (cf. s_mod_combobox dans tab_pack.py) - donc, comme pour le
        listbox des batteries (display_model_list_pack), on affiche les OF de
        toute la famille PPTR018A (LIKE 'PPTR018A%') plutôt que la seule
        variante exacte.
        """
        conn = self.db_manager.connect()
        if not conn:
            return []
        try:
            cursor = conn.cursor()
            modele = str(self.selected_model)
            if modele[:8] == "PPTR018A":
                query = (
                    "SELECT n_of FROM ref_of "
                    "WHERE reference_batterie LIKE %s "
                    "AND (etat_fabrication IS NULL OR etat_fabrication != %s)"
                )
                param = (modele[:8] + "%", "terminé")
            else:
                query = (
                    "SELECT n_of FROM ref_of "
                    "WHERE reference_batterie = %s "
                    "AND (etat_fabrication IS NULL OR etat_fabrication != %s)"
                )
                param = (modele, "terminé")
            cursor.execute(query, param)
            result = cursor.fetchall()
            return [row[0] for row in result]
        finally:
            try:
                cursor.close()
            except Exception:
                pass
            conn.close()

    def check_of(self, numero_serie, of_saisi):
        """
        Compare l'OF réellement affecté à `numero_serie` en base (produit_voltr.n_of)
        avec l'OF sélectionné/saisi par l'opérateur (`of_saisi`). En cas d'écart,
        propose d'échanger cette batterie avec une autre pour remettre les OF en
        cohérence (cf. handle_of_mismatch / swap_of).

        ⚠️ `of_saisi` est une chaîne (valeur de combobox) alors que `n_of` peut être
        stocké en base comme un entier : vérifie que la comparaison se comporte
        comme attendu sur ta base avant mise en prod (sinon caster les deux côtés
        en str, ou en int, avant de comparer).
        """
        conn = None
        cursor = None
        try:
            conn = self.db_manager.connect()
            if not conn:
                return
            cursor = conn.cursor()

            query = "SELECT n_of FROM produit_voltr WHERE numero_serie_produit = %s"
            cursor.execute(query, (numero_serie,))
            result = cursor.fetchone()

            if not result:
                messagebox.showerror("Erreur", "Produit introuvable en base")
                return

            of_bdd = result[0]

            if of_saisi != of_bdd:
                self.handle_of_mismatch(numero_serie, of_bdd, of_saisi)

        except Exception as e:
            messagebox.showerror("Erreur BDD", str(e))
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    def handle_of_mismatch(self, numero_serie, of_bdd, of_saisi):
        """Propose à l'opérateur d'échanger la batterie scannée avec une autre
        lorsque l'OF saisi ne correspond pas à celui enregistré en base."""
        response = messagebox.askyesno(
            "OF différent",
            f"L'OF saisi ({of_saisi}) est différent de celui en base ({of_bdd}).\n\n"
            "Voulez-vous remplacer cette batterie dans l'OF ?"
        )
        if response:
            self.swap_of(numero_serie, of_saisi, of_bdd)

    def swap_of(self, numero_serie_initial, nouvel_of, ancien_of):
        """
        Échange l'affectation OF (produit_voltr.n_of) entre `numero_serie_initial`
        et une batterie de remplacement saisie par l'opérateur, à condition que
        cette dernière appartienne bien à `nouvel_of`.
        """
        conn = None
        cursor = None
        try:
            conn = self.db_manager.connect()
            if not conn:
                return
            cursor = conn.cursor()

            new_serial = simpledialog.askstring(
                "Remplacement batterie",
                "Entrer le numéro de série de la batterie de remplacement :"
            )
            if not new_serial:
                return

            if numero_serie_initial == new_serial:
                messagebox.showwarning("Erreur", "Impossible de remplacer par la même batterie")
                return

            query = "SELECT n_of FROM produit_voltr WHERE numero_serie_produit = %s"
            cursor.execute(query, (new_serial,))
            result = cursor.fetchone()

            if not result:
                messagebox.showerror("Erreur", "Batterie de remplacement introuvable")
                return

            of_remplacement = result[0]

            if of_remplacement != nouvel_of:
                messagebox.showerror("Erreur", f"Batterie de remplacement ne fait pas partie du bon of {nouvel_of}")
                return

            update_query = """
            UPDATE produit_voltr 
            SET n_of = CASE 
                WHEN numero_serie_produit = %s THEN %s
                WHEN numero_serie_produit = %s THEN %s
            END
            WHERE numero_serie_produit IN (%s, %s)
            """
            cursor.execute(update_query, (
                numero_serie_initial, of_remplacement,
                new_serial, ancien_of,
                numero_serie_initial, new_serial
            ))

            conn.commit()
            messagebox.showinfo("Succès", "Les OF ont été échangés avec succès")

        except Exception as e:
            if conn:
                conn.rollback()
            messagebox.showerror("Erreur BDD", str(e))
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    def update_avancement(self, stage, valeur, num_of):
        """Met à jour le label 'avancement_of_{stage}' et le combobox 'combo_of_{stage}'
        de l'onglet correspondant. Utilise getattr pour rester silencieux (juste un
        print de diagnostic) sur les onglets qui n'ont pas ces widgets (ex: picking)."""
        widget_name = f"avancement_of_{stage}"
        combo_name = f"combo_of_{stage}"

        widget = getattr(self, widget_name, None)
        combo = getattr(self, combo_name, None)

        if widget is not None:
            widget.config(text=valeur)
        else:
            print(f"Widget {widget_name} introuvable")

        if combo is not None:
            combo.set(num_of)
        else:
            print(f"Widget {combo_name} introuvable")

    # Étapes de fabrication suivies par le mécanisme OF (dans l'ordre du flux).
    # "recherche" n'en fait pas partie : c'est un outil de consultation, pas une
    # étape de fabrication à proprement parler - elle est ignorée ici, ce qui
    # fait que la dernière de ces étapes active pour un modèle correspond bien
    # à "l'onglet qui précède recherche" dans le flux de production.
    OF_TRACKED_STAGES = ("pack", "nappe", "bms", "wrap", "fermeture_batt", "fin_ligne")

    def _last_of_tracked_stage(self):
        """Renvoie la dernière étape suivie par OF active pour le modèle en
        cours (celle juste avant l'onglet recherche dans le flux), ou None."""
        ordered_keys = getattr(self, "ordered_keys", None)
        if not ordered_keys:
            return None
        tracked = [s for s in ordered_keys if s in self.OF_TRACKED_STAGES]
        return tracked[-1] if tracked else None

    def _mark_of_termine_if_complete(self, num_of, nb_faites, total):
        """Marque l'OF `num_of` comme 'terminé' dans ref_of si toutes les
        batteries prévues (quantite_batterie) ont bien validé la dernière
        étape de fabrication suivie par OF."""
        if not num_of or not total or nb_faites is None or nb_faites < total:
            return
        conn = self.db_manager.connect()
        if not conn:
            return
        try:
            cursor = conn.cursor()
            cursor.execute(
                "UPDATE ref_of SET etat_fabrication = %s WHERE n_of = %s",
                ("terminé", num_of)
            )
            conn.commit()
        finally:
            try:
                cursor.close()
            except Exception:
                pass
            conn.close()

    def of_avancement(self, event=None):
        """
        Rafraîchit l'affichage "Avancement of : X/Y" de l'onglet actif.
        - Sur l'onglet pack, l'OF suivi est celui choisi manuellement dans le
          combobox 'N° of' (self.s_numero_of), car c'est à cette étape que la
          batterie est affectée à un OF.
        - Sur les autres onglets (nappe, bms, wrap, fermeture, fin de ligne),
          l'OF suivi est déduit automatiquement de la batterie en cours de
          saisie (self.entry_widgets[stage]), via produit_voltr.n_of.
        Si l'étape active est la dernière étape suivie par OF (juste avant
        recherche) et que toutes les batteries prévues l'ont validée, l'OF est
        automatiquement marqué 'terminé' dans ref_of.
        """
        stage = self._get_active_stage()
        colonne = self.STAGE_TO_DBCOL.get(stage)
        if stage == "pack":
            num_of = self.s_numero_of.get()
        else:
            widget = self.entry_widgets.get(stage)
            if widget is None:
                return
            num_batt = widget.get()
            conn = self.db_manager.connect()
            if not conn:
                return
            cursor = conn.cursor()
            query = "select n_of from produit_voltr where numero_serie_produit=%s"
            param = (num_batt,)
            cursor.execute(query, param)
            rows = cursor.fetchall()
            if not rows:
                cursor.close()
                conn.close()
                return
            num_of = rows[0][0]
            cursor.close()
            conn.close()

        conn = self.db_manager.connect()
        if not conn:
            return
        try:
            cursor = conn.cursor()

            query = "SELECT quantite_batterie FROM ref_of where n_of=%s"
            param = (num_of,)
            cursor.execute(query, param)
            tot_of = cursor.fetchall()[0][0]

            query = (
                "SELECT count(pv.numero_serie_produit) FROM produit_voltr as pv "
                "join suivi_production as sp on sp.numero_serie_batterie=pv.numero_serie_produit "
                f"where pv.n_of=%s and sp.{colonne}=1"
            )
            param = (num_of,)
            cursor.execute(query, param)
            nb_of = cursor.fetchall()[0][0]

            ratio = f"Avancement of {num_of} :\n{nb_of}/{tot_of}"

            self.update_avancement(stage, ratio, num_of)

            if stage == self._last_of_tracked_stage():
                self._mark_of_termine_if_complete(num_of, nb_of, tot_of)

        finally:
            try:
                cursor.close()
            except Exception:
                pass
            conn.close()

    def check_entry_length_batt(self, event=None):
        """
        Bindée sur le champ 'n° série produit' des onglets nappe/bms/wrap/
        fermeture/fin de ligne (et picking, sans effet visible - cf.
        of_avancement) : dès que 9 caractères sont saisis (longueur d'un n°
        de série produit), rafraîchit l'avancement OF de l'onglet actif.
        """
        stage = self._get_active_stage()
        widget = self.entry_widgets.get(stage)
        if widget is None:
            return
        if len(widget.get()) == 9:
            self.of_avancement()

    #------------------------------ Onglet choix DCIR -----------------------------------------------
    
