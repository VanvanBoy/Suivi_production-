import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from ttkthemes import ThemedTk
from PIL import ImageTk, Image
import mysql.connector
import pandas as pd
from PIL import Image, ImageTk  

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

class StockApp(ThemedTk):

    def __init__(self):
        super().__init__(theme="xpnative")
        self.title("Suivi de production")
        self.geometry("900x600")
        self.db_manager = DBManager()
        self.picking_file_path = None

        self.style = ttk.Style()
        self.style.configure("TLabel", font=('Segoe UI', 11), padding=4)
        self.style.configure("TButton", font=('Segoe UI', 11), padding=6)
        self.style.configure("TEntry", font=('Segoe UI', 11))
        self.style.configure("TCombobox", font=('Segoe UI', 11))

        self.style.configure("Danger.TButton", foreground="black", background="#F06F65")
        self.style.map("Danger.TButton",
                    background=[('active', '#e55b50')],
                    foreground=[('disabled', 'black')])
        
        self.style.configure("Good.TButton", foreground="black", background="#0ED329")
        self.style.map("Good.TButton",
                    background=[('active', '#0ED329')],
                    foreground=[('disabled', 'black')])

        self.create_widgets()

    def create_widgets(self):
        notebook = ttk.Notebook(self)
        notebook.pack(expand=True, fill="both")

        tab_picking = ttk.Frame(notebook)
        notebook.add(tab_picking, text="Contrôle de picking")
        self.setup_picking(tab_picking)

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

        #Onglet emballage
        tab_emb = ttk.Frame(notebook)
        notebook.add(tab_emb, text="Controle emballage")
        self.setup_emb(tab_emb)

        #Onglet expedition
        tab_exp = ttk.Frame(notebook)
        notebook.add(tab_exp, text="Controle expedition")
        self.setup_exp(tab_exp)

        # Onglet recherche
        tab_recherche = ttk.Frame(notebook)
        notebook.add(tab_recherche, text="Recherche de batterie")
        self.setup_recherche(tab_capa)
        
    #Onglet picking

    def setup_picking(self, frame):
        # --- Partie gauche : cellules à mettre en plateau ---
        left_frame = ttk.Frame(frame)
        left_frame.pack(side="left", fill='both', expand=True, padx=20, pady=20)

        ttk.Label(left_frame, text="N° série d'une cellule du produit:").pack(pady=5)
        self.numero_serie_cell_entry = ttk.Entry(left_frame)
        self.numero_serie_cell_entry.pack(pady=5)

        ttk.Label(left_frame, text="N° série produit:").pack(pady=5)
        self.numero_serie_batt_entry = ttk.Entry(left_frame)
        self.numero_serie_batt_entry.pack(pady=5)

        ttk.Label(left_frame, text="Liste des batteries du modèle:").pack(pady=5)
        self.listbox_batt = tk.Listbox(left_frame, font=('Segoe UI', 11), height=10)
        self.listbox_batt.pack(pady=5, fill='both', expand=False)

        self.button_non_conf = ttk.Button(left_frame, text="❌ Non conforme", command=self.add_non_conf_cellule,style="Danger.TButton")
        self.button_non_conf.pack(pady=10)

        # --- Partie droite : remplacement de cellule ---
        right_frame = ttk.Frame(frame)
        right_frame.pack(side="right", fill='both', expand=True, padx=20, pady=20)

        ttk.Label(right_frame, text="🔁 Remplacement cellule:").pack(pady=5)

        ttk.Label(right_frame, text="N° série cellule:").pack(pady=5)
        self.numero_cell_r_entry = ttk.Entry(right_frame)
        self.numero_cell_r_entry.pack(pady=5)

        ttk.Label(right_frame, text="Défaut:").pack(pady=5)
        self.combobox_default = ttk.Combobox(right_frame, state="readonly", values=["Non trouvée", "Tension", "Corrosion", "Déformation"])
        self.combobox_default.pack(pady=5)

        self.button_remplacer = ttk.Button(right_frame, text="🔄 Remplacer cellule", command=self.replace_cellule,style="Danger.TButton")
        self.button_remplacer.pack(pady=10)

        # Cadre visuel : Écart tension
        frame_info = tk.Frame(right_frame, width=300, height=100, bg='#D0F5BE')
        frame_info.pack(pady=20)
        frame_info.pack_propagate(False)

        label_info = tk.Label(frame_info, text="⚠ Écart maximum de 0.05V", bg='#D0F5BE', fg='black', font=("Segoe UI", 11, 'bold'))
        label_info.pack(expand=True)

        # Bouton contrôle OK
        self.button_cont_ok = ttk.Button(right_frame, text="✅ Contrôle OK", command=self.valider_picking,style="Good.TButton")
        self.button_cont_ok.pack(pady=10)

    #Onglet Soudure

    def setup_pack(self, frame):
        # --- Partie gauche : cellules à mettre en plateau ---
        left_frame = ttk.Frame(frame)
        left_frame.pack(side="left", fill='both', expand=True, padx=20, pady=20)

        ttk.Label(left_frame, text="N° série d'une cellule du produit:").pack(pady=5)
        self.s_numero_serie_cell_entry = ttk.Entry(left_frame)
        self.s_numero_serie_cell_entry.pack(pady=5)

        ttk.Label(left_frame, text="N° série produit:").pack(pady=5)
        self.s_numero_serie_batt_entry = ttk.Entry(left_frame)
        self.s_numero_serie_batt_entry.pack(pady=5)

        self.s_button_non_conf = ttk.Button(left_frame, text="❌ Non conforme", command=self.add_non_conf_cellule,style="Danger.TButton")
        self.s_button_non_conf.pack(pady=10)

        # --- Partie droite : remplacement de cellule ---
        right_frame = ttk.Frame(frame)
        right_frame.pack(side="right", fill='both', expand=True, padx=20, pady=20)

        ttk.Label(right_frame, text="Liste des batteries du modèle:").pack(pady=5)
        self.s_listbox_batt = tk.Listbox(right_frame, font=('Segoe UI', 11), height=10)
        self.s_listbox_batt.pack(pady=5, fill='both', expand=False)

        # Bouton contrôle OK
        self.s_button_cont_ok = ttk.Button(right_frame, text="✅ Contrôle OK", command=self.valider_picking,style="Good.TButton")
        self.s_button_cont_ok.pack(pady=10)

    #Onglet Soudure nappe

    def setup_nappe(self, frame):

        # --- Partie gauche : cellules à mettre en plateau ---
        left_frame = ttk.Frame(frame)
        left_frame.pack(side="left", fill='both', expand=True, padx=20, pady=20)

        ttk.Label(left_frame, text="N° série d'une cellule du produit:").pack(pady=5)
        self.n_numero_serie_cell_entry = ttk.Entry(left_frame)
        self.n_numero_serie_cell_entry.pack(pady=5)

        ttk.Label(left_frame, text="N° série produit:").pack(pady=5)
        self.n_numero_serie_batt_entry = ttk.Entry(left_frame)
        self.n_numero_serie_batt_entry.pack(pady=5)

        self.n_button_non_conf = ttk.Button(left_frame, text="❌ Non conforme", command=self.add_non_conf_cellule,style="Danger.TButton")
        self.n_button_non_conf.pack(pady=10)

        ttk.Label(left_frame, text="Liste des batteries du modèle:").pack(pady=5)
        self.n_listbox_batt = tk.Listbox(left_frame, font=('Segoe UI', 11), height=10)
        self.n_listbox_batt.pack(pady=5, fill='both', expand=False)

        # --- Partie droite : remplacement de cellule ---
        right_frame = ttk.Frame(frame)
        right_frame.pack(side="right", fill='both', expand=True, padx=20, pady=20)

        ttk.Label(right_frame, text="Ecart tension modules:").pack(pady=5)
        self.ecart_t_entry=ttk.Entry(right_frame).pack(pady=5)

        self.label_photo = tk.Label(right_frame, bg="#e0e0e0", width=200, height=200, text="Aperçu photo", anchor='center')
        self.label_photo.pack(pady=10)
        chemin_image=r"C:\Users\User\Downloads\voltr_logo.jpg"
        self.charger_photo(chemin_image)

        # Bouton contrôle OK
        self.n_button_cont_ok = ttk.Button(right_frame, text="✅ Contrôle OK", command=self.valider_picking,style="Good.TButton")
        self.n_button_cont_ok.pack(pady=10)
    
    def setup_bms(self, frame):
        
        # --- Partie gauche : cellules à mettre en plateau ---
        left_frame = ttk.Frame(frame)
        left_frame.pack(side="left", fill='both', expand=True, padx=20, pady=20)

        ttk.Label(left_frame, text="N° série d'une cellule du produit:").pack(pady=5)
        self.b_numero_serie_cell_entry = ttk.Entry(left_frame)
        self.b_numero_serie_cell_entry.pack(pady=5)

        ttk.Label(left_frame, text="N° série produit:").pack(pady=5)
        self.b_numero_serie_batt_entry = ttk.Entry(left_frame)
        self.b_numero_serie_batt_entry.pack(pady=5)

        self.b_button_non_conf = ttk.Button(left_frame, text="❌ Non conforme", command=self.add_non_conf_cellule,style="Danger.TButton")
        self.b_button_non_conf.pack(pady=10)

        # --- Partie droite : remplacement de cellule ---
        right_frame = ttk.Frame(frame)
        right_frame.pack(side="right", fill='both', expand=True, padx=20, pady=20)

        ttk.Label(left_frame, text="Liste des batteries du modèle:").pack(pady=5)
        self.b_listbox_batt = tk.Listbox(left_frame, font=('Segoe UI', 11), height=10)
        self.b_listbox_batt.pack(pady=5, fill='both', expand=False)

        self.b_label_photo = tk.Label(right_frame, bg="#e0e0e0", width=200, height=200, text="Aperçu photo", anchor='center')
        self.b_label_photo.pack(pady=10)
        chemin_image=r"C:\Users\User\Downloads\voltr_logo.jpg"
        self.charger_photo(chemin_image)

        # Bouton contrôle OK
        self.b_button_cont_ok = ttk.Button(right_frame, text="✅ Contrôle OK", command=self.valider_picking,style="Good.TButton")
        self.b_button_cont_ok.pack(pady=10)
    
    def setup_wrap(self, frame):

        # --- Partie gauche : cellules à mettre en plateau ---
        left_frame = ttk.Frame(frame)
        left_frame.pack(side="left", fill='both', expand=True, padx=20, pady=20)

        ttk.Label(left_frame, text="N° série d'une cellule du produit:").pack(pady=5)
        self.w_numero_serie_cell_entry = ttk.Entry(left_frame)
        self.w_numero_serie_cell_entry.pack(pady=5)

        ttk.Label(left_frame, text="N° série produit:").pack(pady=5)
        self.w_numero_serie_batt_entry = ttk.Entry(left_frame)
        self.w_numero_serie_batt_entry.pack(pady=5)

        self.w_button_non_conf = ttk.Button(left_frame, text="❌ Non conforme", command=self.add_non_conf_cellule,style="Danger.TButton")
        self.w_button_non_conf.pack(pady=10)

        # --- Partie droite : remplacement de cellule ---
        right_frame = ttk.Frame(frame)
        right_frame.pack(side="right", fill='both', expand=True, padx=20, pady=20)

        ttk.Label(right_frame, text="Liste des batteries du modèle:").pack(pady=5)
        self.w_listbox_batt = tk.Listbox(right_frame, font=('Segoe UI', 11), height=10)
        self.w_listbox_batt.pack(pady=5, fill='both', expand=False)

        # Bouton contrôle OK
        self.w_button_cont_ok = ttk.Button(right_frame, text="✅ Contrôle OK", command=self.valider_picking,style="Good.TButton")
        self.w_button_cont_ok.pack(pady=10)

    def setup_fermeture(self, frame):

        # --- Partie gauche : cellules à mettre en plateau ---
        left_frame = ttk.Frame(frame)
        left_frame.pack(side="left", fill='both', expand=True, padx=20, pady=20)

        ttk.Label(left_frame, text="N° série d'une cellule du produit:").pack(pady=5)
        self.f_numero_serie_cell_entry = ttk.Entry(left_frame)
        self.f_numero_serie_cell_entry.pack(pady=5)

        ttk.Label(left_frame, text="N° série produit:").pack(pady=5)
        self.f_numero_serie_batt_entry = ttk.Entry(left_frame)
        self.f_numero_serie_batt_entry.pack(pady=5)

        self.f_button_non_conf = ttk.Button(left_frame, text="❌ Non conforme", command=self.add_non_conf_cellule,style="Danger.TButton")
        self.f_button_non_conf.pack(pady=10)

        # --- Partie droite : remplacement de cellule ---
        right_frame = ttk.Frame(frame)
        right_frame.pack(side="right", fill='both', expand=True, padx=20, pady=20)

        ttk.Label(left_frame, text="Liste des batteries du modèle:").pack(pady=5)
        self.f_listbox_batt = tk.Listbox(left_frame, font=('Segoe UI', 11), height=10)
        self.f_listbox_batt.pack(pady=5, fill='both', expand=False)

        ttk.Label(right_frame, text="Tension en fin de test:").pack(pady=5)
        self.tension_end_entry=ttk.Entry(right_frame)
        self.tension_end_entry.pack(pady=5)

        self.f_label_photo = tk.Label(right_frame, bg="#e0e0e0", width=200, height=200, text="Aperçu photo", anchor='center')
        self.f_label_photo.pack(pady=10)
        chemin_image=r"C:\Users\User\Downloads\voltr_logo.jpg"
        self.charger_photo(chemin_image)

        # Bouton contrôle OK
        self.f_button_cont_ok = ttk.Button(right_frame, text="✅ Contrôle OK", command=self.valider_picking,style="Good.TButton")
        self.f_button_cont_ok.pack(pady=10)



    def setup_emb(self, frame):
        print('tqt')

    def setup_capa(self, frame):
        print("Onglet Test de capacité non implémenté")

    def setup_recherche(self,frame):
        print("Onglet Recherche de batterie non implémenté")
    
    def setup_exp(self,frame):
        print('emma')

    def valider_picking(self):
        print("Contrôle OK validé")

    def add_non_conf_cellule(self):
        print("Cellule ajoutée en non conforme")

    def replace_cellule(self):
        print("Cellule remplacée")

    def charger_photo(self, chemin_image):
        try:
            image = Image.open(chemin_image)
            image = image.resize((200, 200))  # adapter à la taille de ton Label
            photo = ImageTk.PhotoImage(image)
            self.label_photo.config(image=photo, text="")  # enlever le texte
            self.label_photo.image = photo  # éviter le garbage collector
            self.b_label_photo.config(image=photo, text="")  # enlever le texte
            self.b_label_photo.image = photo  # éviter le garbage collector
            self.f_label_photo.config(image=photo, text="")  # enlever le texte
            self.f_label_photo.image = photo  # éviter le garbage collector
        except Exception as e:
            messagebox.showerror("Erreur image", f"Impossible de charger l'image : {e}")


if __name__ == "__main__":
    app = StockApp()
    app.mainloop()
