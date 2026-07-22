# -*- coding: utf-8 -*-
"""
Gestion de la connexion a la base de données MySQL.
"""
from tkinter import messagebox, simpledialog
import mysql.connector


class DBManager:
    def __init__(self):

        def get_db_credentials():
            # Fonction pour obtenir les informations d'identification de l'utilisateur
            user = simpledialog.askstring("Login", "Enter your MySQL username:")
            password = simpledialog.askstring("Login", "Enter your MySQL password:", show='*')
            return user, password

        self.user, self.password = get_db_credentials()

        self.config = {
            'user': self.user,
            'password': self.password,
            'host': '34.77.226.40',
            'database': 'bdd_22072026',
            'auth_plugin': 'mysql_native_password'
        }

    def connect(self):
        try:
            return mysql.connector.connect(**self.config)
        except mysql.connector.Error as err:
            messagebox.showerror("Erreur DB", f"Erreur lors de la connexion : {err}")
            return None
