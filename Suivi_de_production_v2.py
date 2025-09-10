import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import mysql.connector
import pandas as pd

class DBManager:
    def __init__(self):
        self.config = {
            'user': 'Vanvan',
            'password': 'VoltR99!', 
            'host': '34.77.226.40', 
            'database': 'cellules_batteries_cloud',
            'auth_plugin': 'mysql_native_password'
        }
    
    def connect(self):
        try:
            conn = mysql.connector.connect(**self.config)
            return conn
        except mysql.connector.Error as err:
            messagebox.showerror("Erreur DB", f"Erreur lors de la connexion : {err}")
            return None

class StockApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Suivi de production")
        self.geometry("800x600")  
        self.db_manager = DBManager()
        
        self.picking_file_path = None
        
        self.create_widgets()

    def create_widgets(self):
        notebook = ttk.Notebook(self)
        notebook.pack(expand=True, fill="both")

        # Onglet picking
        tab_picking = ttk.Frame(notebook)
        notebook.add(tab_picking, text="Controle de picking")
        self.setup_picking(tab_picking)
        """
        # Onglet soudure pack
        tab_pack = ttk.Frame(notebook)
        notebook.add(tab_pack, text="Controle soudure pack")
        self.setup_pack(tab_pack)

        # Onglet soudure nappe
        tab_nappe = ttk.Frame(notebook)
        notebook.add(tab_nappe, text="Controle soudure nappe")
        self.setup_nappe(tab_nappe)

        # Onglet soudure BMS
        tab_bms = ttk.Frame(notebook)
        notebook.add(tab_bms, text="Controle soudure BMS")
        self.setup_bms(tab_bms)

        # Onglet wrap
        tab_wrap = ttk.Frame(notebook)
        notebook.add(tab_wrap, text="Controle wrap")
        self.setup_wrap(tab_wrap)

        # Onglet fermeture
        tab_fermeture = ttk.Frame(notebook)
        notebook.add(tab_fermeture, text="Controle fermeture")
        self.setup_fermeture(tab_fermeture)

        # Onglet test capa
        tab_capa = ttk.Frame(notebook)
        notebook.add(tab_capa, text="Test de capacité")
        self.setup_capa(tab_capa)
        """
    # -------------------------------
    # Onglet Picking
    # -------------------------------
    def setup_picking(self, frame):

        # --- Partie gauche : cellules a mettre en plateau ---
        left_frame = tk.Frame(frame)
        left_frame.pack(side="left", fill='y', expand=True)

        self.numero_serie_cell_label = tk.Label(left_frame, text="Numéro de série d'une cellule du produit:", font=('Arial', 12))
        self.numero_serie_cell_label.pack(pady=5)
        self.numero_serie_cell_entry = tk.Entry(left_frame, font=('Arial', 12))
        self.numero_serie_cell_entry.pack(pady=5)
        #self.numero_serie_cell_entry.bind('<KeyRelease>', check_entry_length)

        self.numero_serie_batt_label = tk.Label(left_frame, text="Numéro de série produit:", font=('Arial', 12))
        self.numero_serie_batt_label.pack(pady=5)
        self.numero_serie_batt_entry = tk.Entry(left_frame, font=('Arial', 12))
        self.numero_serie_batt_entry.pack(pady=5)

        self.list_batt_label= tk.Label(left_frame, text="Liste des batteries du modèle:", font=('Arial', 12))
        self.list_batt_label.pack(pady=5)
        self.listbox_batt= tk.Listbox(left_frame, font=('Arial', 12), height=10)
        self.listbox_batt.pack(pady=5, fill='both', expand=False)

        self.button_non_conf= tk.Button(left_frame, text="non conforme", command=self.add_non_conf_cellule, bg="#F06F65", fg="black")
        self.button_non_conf.pack(pady=5)

        # --- Partie droite : recherche/changement de place ---
        right_frame = tk.Frame(frame)
        right_frame.pack(side="right", fill='y', expand=True)

        tk.Label(right_frame, text="Remplacement cellule:").pack(pady=5)

        self.numero_cell_r_label = tk.Label(right_frame, text="Numéro série cellule:", font=('Arial', 12))
        self.numero_cell_r_label.pack(pady=5)
        self.numero_cell_r_entry = tk.Entry(right_frame, font=('Arial', 12))
        self.numero_cell_r_entry.pack(pady=5)

        self.default_label = tk.Label(right_frame, text="Défaut:", font=('Arial', 12))
        self.default_label.pack(pady=5)
        self.combobox_default = ttk.Combobox(right_frame, state="readonly", font=('Arial', 12))
        self.combobox_default.pack(pady=5)
        self.combobox_default.values= ["Non trouvée", "Tension", "Corrosion","Déformation"]


        self.button_remplacer = tk.Button(right_frame, text="Remplacer cellule", command=self.replace_cellule, bg="#F06F65", fg="black")
        self.button_remplacer.pack(pady=5)

        # Création d'un cadre carré avec couleur de fond verte
        carre = tk.Frame(right_frame, width=300, height=100, bg='#99ff99')  # vert clair
        carre.pack(padx=20, pady=20)

        # Empêche le redimensionnement du cadre par son contenu
        carre.pack_propagate(False)
        # Ajout du texte centré
        label = tk.Label(carre, text="Ecart maximum de 0.05V", bg='#99ff99', fg='black', font=("Arial", 12))
        label.pack(expand=True)

        self.button_cont_ok = tk.Button(right_frame, text="Controle ok", command=self.valider_picking, bg="#66B2FF", fg="black")
        self.button_cont_ok.pack(pady=5)

    def valider_picking(self):
        print("test")

    def add_non_conf_cellule(self):
        print("test")
    
    def replace_cellule(self):
        print("test")

    # -------------------------------
    # Onglet Picking
    # -------------------------------

    



if __name__ == "__main__":
        app = StockApp()
        app.mainloop()