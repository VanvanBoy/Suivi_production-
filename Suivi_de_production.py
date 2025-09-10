# -*- coding: utf-8 -*-
"""
Created on Fri Jan 17 14:17:34 2025

@author: Vanvan 

IHM suivi de prod 
"""

#Importation des bibliotheques et des fonctions de création d'onglets
import tkinter as tk 
from tkinter import ttk,simpledialog
from Onglet_picking import create_picking_interface 
from Onglet_pack import create_pack_interface 
from Onglet_nappe import create_nappe_interface
from Onglet_fermeture import create_fermeture_interface
from Onglet_bms import create_bms_interface
from Onglet_wrap import create_wrap_interface
from Onglet_test_capa import create_test_interface

import mysql.connector
from mysql.connector.locales.eng import client_error

def get_db_credentials():
    # Fonction pour obtenir les informations d'identification de l'utilisateur
    user = simpledialog.askstring("Login", "Enter your MySQL username:")
    password = simpledialog.askstring("Login", "Enter your MySQL password:", show='*')
    return user, password

#definition de la fonction main()
def main():
    # Obtenir les informations d'identification de l'utilisateur
    user, password = get_db_credentials()
    # Connexion à la base de données
    
    conn = mysql.connector.connect(
        host="34.77.226.40",
        user=user,
        password=password,
        database="cellules_batteries_cloud",
        auth_plugin='mysql_native_password'
    ) 

    #Creation du curseur sql
    cursor = conn.cursor()
    
    #Fonction qui crée le contenu l'onglet selectionnée
    def on_tab_change(event):
        # Récuperation du nom de l'onglet sélectionné
        selected_tab = event.widget.tab(event.widget.select(), "text")
        # Creation de l'interface de l'onglet selectionné 
        if selected_tab == "Controle de picking":
            for widget in tab1.winfo_children():
             widget.destroy() #destruction de l'ancienne strucutre afin de mettre a jour les infos 
            create_picking_interface(tab1, conn, cursor)
            
        elif selected_tab == "Controle soudure pack":
            for widget in tab2.winfo_children():
             widget.destroy()
            create_pack_interface(tab2, conn, cursor)
        elif selected_tab == "Controle soudure nappe":
            for widget in tab3.winfo_children():
             widget.destroy()
            create_nappe_interface(tab3, conn, cursor)  
        elif selected_tab == "Controle soudure BMS":
            for widget in tab4.winfo_children():
             widget.destroy()
            create_bms_interface(tab4, conn, cursor)  
        elif selected_tab == "Controle wrap":
            for widget in tab5.winfo_children():
             widget.destroy()
            create_wrap_interface(tab5, conn, cursor)
        elif selected_tab == "Controle fermeture":
            for widget in tab6.winfo_children():
             widget.destroy()
            create_fermeture_interface(tab6, conn, cursor)
        elif selected_tab =="Test capa":
            for widget in tab7.winfo_children():
             widget.destroy()
            create_test_interface(tab7, conn, cursor)
            
            
     
    # Création de la fenêtre principale + titre et dimensions
    root = tk.Tk()
    root.title("Suivi production")
    root.geometry("800x600")

    #Création de l'interface à onglets
    tab_control = ttk.Notebook(root)
    
    # Création des huit onglets 
    tab1 = ttk.Frame(tab_control)
    tab2 = ttk.Frame(tab_control)
    tab3 = ttk.Frame(tab_control)
    tab4 = ttk.Frame(tab_control)
    tab5 = ttk.Frame(tab_control)
    tab6 = ttk.Frame(tab_control)
    tab7 = ttk.Frame(tab_control)
    
    # Ajouter les onglets à l'interface à onglets 
    tab_control.add(tab1, text="Controle de picking")
    tab_control.add(tab2, text="Controle soudure pack")
    tab_control.add(tab3, text="Controle soudure nappe")
    tab_control.add(tab4, text="Controle soudure BMS")
    tab_control.add(tab5, text="Controle wrap")
    tab_control.add(tab6, text="Controle fermeture")
    tab_control.add(tab7, text="Test capa")
    
    # Placer l'interface à onglets dans la fenêtre principale
    tab_control.pack(expand=1, fill="both")
    
    #association entre l'evenement changement d'onglet et la fonction on_tab_change
    tab_control.bind("<<NotebookTabChanged>>", on_tab_change)

    #Boucle principale 
    root.mainloop()
    
    # Fermeture du curseur et de la connexion
    cursor.close()
    conn.close()
    
#Execution de la fonction main lorsque le script est lancé directement 
#Permet de ne pas executere le code lorsqu'il est import" en tant que module 
if __name__ == "__main__":
    main()

