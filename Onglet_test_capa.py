# -*- coding: utf-8 -*-
"""
Created on Wed Feb  5 13:33:33 2025

@author: User
"""

import sys
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import mysql.connector
import pandas as pd
import os
import shutil
from datetime import datetime 
from tkinter import font

def create_test_interface(tab, conn, cursor):
    
    # Fonction pour traiter les fichiers Excel avec les resultats des tests et mettre à jour la base de données
    def resultats_batteries():
        try:
            now = datetime.now()
            # Formater la date et l'heure
            date_heure = now.strftime("%Y-%m-%d %H:%M:%S")
            timing=str(date_heure)
            label_resultat.config(text="Dernier traitement des fichiers :\n"+timing)
            list_baie=[]
            new_series_generated = [] #Liste pour stocker les numeros de series des cellules testees
            new_series_pb=[] #Liste pour stocker les numeros series des cellules dont les tests ont un problemes
            new_series_ko=[]
            df_pb_test=pd.DataFrame(columns=["numero_serie", "emplacement","fichier source","fichier destination"])
            df_pb_ko=pd.DataFrame(columns=["numero_serie", "emplacement","fichier source","fichier destination"])
            
            
            """
            dossier="C:/Users/User/Desktop/MAJ_TEST/Resultats"
            dossier_ok="C:/Users/User/Desktop/MAJ_TEST/Fichiers_exploites"
            dossier_nok="C:/Users/User/Desktop/MAJ_TEST/Fichiers_non_ok/cellues K.O"
            dossier_pb_test="C:/Users/User/Desktop/MAJ_TEST/Fichiers_non_ok/pb test"
            """
            
            dossier = "G:/Drive partagés/VoltR/4_Production/5_Cyclage/1_Résultats de cyclage/Fichier en cours"
            dossier_ok= "G:/Drive partagés/VoltR/4_Production/5_Cyclage/1_Résultats de cyclage/Batteries traitées/Batteries OK"
            dossier_nok= "G:/Drive partagés/VoltR/4_Production/5_Cyclage/1_Résultats de cyclage/Batteries traitées/Batteries NOK"
            dossier_pb_test="G:/Drive partagés/VoltR/4_Production/5_Cyclage/1_Résultats de cyclage/Batteries traitées/Tests defaillants"
            
            
            cursor.execute("SELECT Numero_serie_batterie FROM suivi_production WHERE Test_capa=%s or Test_capa is null",(0,))
            numeros_series = cursor.fetchall() #Variable qui stock tout les numeros de serie de cellules qui doivent etre demantelees
          
            
            
            if not os.listdir(dossier): #cas ou le dossier est vide
                messagebox.showinfo("","Le dossier est vide !")
                return
            # Traitement des fichiers Excel dans le dossier
            for fichier in os.listdir(dossier):#Traite chaque fichier dans le dossier 
                if fichier.endswith('.xlsx') or fichier.endswith('.xls'):#Verifie si le fichier est de type excel
                    chemin_fichier = os.path.join(dossier, fichier) #Cration du chemin vers le fichier 
                    numero_serie_test = os.path.splitext(os.path.basename(chemin_fichier))[0] #Obtenir le nom du fichier puis retirer l'extension '.xls'
                    #je devrais surment remplacer os.path.basename(chemin_fichier) par fichier 
                    numero_serie_batt,id_cycleur1 = numero_serie_test.split('-',1)#[0] #Les noms de fichier des tests sont du type MC000001001_2_32_2 ou MC000001001-2_32_2
                    numero_serie_batt = numero_serie_batt.split('_')[0] # Les 2 lignes nous permette d'extraire le numero serie cellules du nom du fichier test
                    
                    # Trouver les positions des tirets
                    positions = [pos for pos, char in enumerate(id_cycleur1) if char == '-']
                    
                    # Vérifier qu'il y a au moins 3 tirets
                    if len(positions) >= 3:
                        # Extraire la sous-chaîne jusqu'au 3ème tiret
                        id_cycleur = id_cycleur1[:positions[2]]
                    else:
                        # Si moins de 3 tirets, on garde tout le texte
                        id_cycleur = id_cycleur1
                     
                    id_baie,suite=id_cycleur.split('-',1)

                    #Verifier si le numero serie du fichier correspond a celui d'une cellule caracterisees
                    
                    if numero_serie_batt in [item[0] for item in numeros_series]: # Si on trouve un concordance 
                        
                        print(numero_serie_batt)
                    
                    #if numero_serie_cellule==numeros_series:
                    
                        new_series_generated.append(numero_serie_batt) #Ajout du num-serie dans la liste des cellules traitées
                        
                        chemier_destination=os.path.join(dossier_ok, fichier)
  
                        #Date de test
                        
                        date_sheet = pd.read_excel(chemin_fichier, sheet_name='test')
                        date_test = date_sheet.at[0, 'Unnamed: 8']
                        #date_test = date_test.split()[0]

                        # Convertir la chaîne en format de date SQL
                        #date_sql = datetime.strptime(date_test, '%Y-%m-%d').strftime('%Y-%m-%d')
                        
                        #Capacité
                        
                        data = pd.read_excel(chemin_fichier, sheet_name='cycle') #Extraction des donnees de la sheet 'step'
                        col_capa = data['DChg. Cap.(Ah)'] #Extraire la colonne STEP Type
                        
                        capa_dch=float(col_capa[0])
                        
                        #Verification du SOH
                        cursor.execute("SELECT b.Capacite FROM suivi_production sp JOIN ref_batterie_voltr b ON sp.reference_batterie_voltr = b.Reference_batterie_voltr WHERE sp.Numero_serie_batterie = %s", (numero_serie_batt,))
                        
                        capa_cible = cursor.fetchone()[0] if cursor.rowcount else None
                        
                        if capa_dch - capa_cible < 0:
                            chemier_destination=os.path.join(dossier_nok, fichier)
                            new_series_generated.pop()
                            new_series_ko.append(numero_serie_batt)
                            new_row=pd.DataFrame([{"numero_serie": numero_serie_batt,
                                "emplacement": id_cycleur,
                                "fichier source": chemin_fichier,
                                "fichier destination": chemier_destination,
                                "capacité":capa_dch}])
                            df_pb_ko = pd.concat([df_pb_ko, new_row], ignore_index=True)
                            shutil.move(chemin_fichier,chemier_destination)
                            cursor.execute("UPDATE suivi_production SET Test_capa = %s,Valeur_test_capa= %s, Date_test_capa= %s WHERE numero_serie_batterie = %s", (0,capa_dch,date_test,numero_serie_batt))
                            conn.commit()
                            continue 
                        
                        # Mettre à jour la base de données
                        
                        list_baie.append(id_baie)

                        cursor.execute("UPDATE suivi_production SET Test_capa = %s,Valeur_test_capa= %s, Date_test_capa= %s WHERE numero_serie_batterie = %s", (1,capa_dch,date_test,numero_serie_batt))
                        shutil.move(chemin_fichier, os.path.join(dossier_ok, fichier))
                        # Commit les modifications
                        conn.commit()       
                     
            if new_series_generated:
                # Mettre à jour le treeview pour afficher les résultats des nouvelles séries générées
                update_treeview(new_series_generated)
                
                compteur = pd.Series(list_baie).value_counts()

                # Convertir le résultat en DataFrame
                df = compteur.reset_index()
                
                # Renommer les colonnes
                df.columns = ['Baie', 'Nombre de test OK']
                
                df = df.sort_values(by='Baie')
                
                for item in treeview3.get_children():
                   treeview3.delete(item)
                
                for index, row in df.iterrows():
                    treeview3.insert("", "end", text=row['Baie'], values=(row['Nombre de test OK'],))
            
            if new_series_pb:
                for num_serie in new_series_generated:
                    df_pb_test = df_pb_test[df_pb_test["numero_serie"] != num_serie]
                # Effacer les éléments actuels du menu déroulant
                for item in treeview1.get_children():
                   treeview1.delete(item)

                # Ajouter les données à partir de la base de données
                for index,row in df_pb_test.iterrows():
                  numero_serie = row['numero_serie']
                  cycleur = row['emplacement']
                  treeview1.insert("", "end",text=numero_serie, values=(cycleur,))

            if new_series_ko:
                for num_serie in new_series_generated:
                    df_pb_ko = df_pb_ko[df_pb_ko["numero_serie"] != num_serie]
                    # Effacer les éléments actuels du menu déroulant
                    for item in treeview2.get_children():
                       treeview2.delete(item)

                    # Ajouter les données à partir de la base de données
                    for index,row in df_pb_ko.iterrows():
                      numero_serie = row['numero_serie']
                      cycleur = row['emplacement']
                      capa= row["capacité"]
                      treeview2.insert("", "end",text=numero_serie, values=(cycleur,capa))
            
            if new_series_generated or new_series_pb or new_series_ko:
                # Ceci est une expression, mais elle n'a pas d'effet
                messagebox.showinfo("Succès !", "Opération réussie !")
            else:
                # Afficher un message si les deux conditions précédentes sont fausses 
                messagebox.showinfo("Vide !", "Pas de nouvelles batteries à traiter !")

        except Exception as e:
            # Obtenir le numéro de ligne où l'erreur s'est produite
            line_number = sys.exc_info()[2].tb_lineno

            # Afficher un message d'erreur avec le numéro de ligne
            error_message = f"Une erreur s'est produite à la ligne {line_number}: {str(e)}"
            messagebox.showerror("Erreur", error_message)

    # Fonction pour mettre à jour le treewiew avec les resultats
    def update_treeview(new_series_generated):
        try:
            #Extraire les données d'interet
            if new_series_generated:
                cursor.execute("SELECT Numero_serie_batterie, Valeur_test_capa FROM suivi_production WHERE Numero_serie_batterie IN (%s)" % ','.join(['%s'] * len(new_series_generated)), new_series_generated)
                data= cursor.fetchall()
            
            # Effacer les éléments actuels du menu déroulant
            for item in treeview.get_children():
                treeview.delete(item)

            # Ajouter les données à partir de la base de données
            for row in data:
                numero_serie = row[0]
                capacite = row[1]
                treeview.insert("", "end",text=numero_serie, values=(capacite))
                
                

        except Exception as e:
            # Obtenir le numéro de ligne où l'erreur s'est produite
            line_number = sys.exc_info()[2].tb_lineno

            # Afficher un message d'erreur avec le numéro de ligne
            error_message = f"Une erreur s'est produite à la ligne {line_number}: {str(e)}"
            messagebox.showerror("Erreur", error_message)
            
    
    # Créer la fenêtre principale
    test_frame = ttk.Frame(tab)
    # Créer les cadres pour les deux colonnes
    left_frame = tk.Frame(test_frame)
    left_frame.pack(side=tk.LEFT, padx=10, pady=10)

    right_frame = tk.Frame(test_frame)
    right_frame.pack(side=tk.RIGHT, padx=10, pady=10)
    
    label0 = tk.Label(left_frame, text="Test conformes")
    label0.pack()
    
    # Créer l'ecran pour display les informations 
    treeview = ttk.Treeview(left_frame, columns=("Capacité",))
    treeview.heading("#0", text="Numéro de Série")
    treeview.heading("#1", text="Capacité")
    
    # Placer le treeview dans le cadre principal
    treeview.pack(fill="both", expand=True)
    
    # Largeur commune pour toutes les colonnes
    col_width = 150
    col_width1=300
    
    # Définir la largeur des colonnes
    treeview.column("#0", width=col_width)  # Largeur de la colonne Numéro de Série
    treeview.column("#1", width=col_width)  # Largeur de la colonne Capacité
    
    # Créer un Label
    label = tk.Label(right_frame, text="Tests non conformes")
    label.pack()
    
    # Créer le Treeview avec une seule colonne supplémentaire
    treeview1 = ttk.Treeview(right_frame, columns=("Position cycleur",))
    
    # Définir les en-têtes de colonnes
    treeview1.heading("#0", text="Numéro de Série")  # Colonne par défaut
    treeview1.heading("#1", text="Position cycleur")  # Colonne que vous avez ajoutée
    
    # Placer le Treeview dans le cadre principal
    treeview1.pack(fill="both", expand=True)
    
    # Définir la largeur des colonnes
    
    treeview1.column("#0", width=col_width1)  # Largeur de la colonne "Numéro de Série"
    treeview1.column("#1", width=col_width1)  # Largeur de la colonne "Position cycleur"
    
    label1 = tk.Label(right_frame, text="Batteries capacité insuffisante")
    label1.pack()
    
    # Créer l'ecran pour display les informations 
    treeview2 = ttk.Treeview(right_frame, columns=("Position cycleur","Capacité"))
    
    treeview2.heading("#0", text="Numéro de Série")
    treeview2.heading("#1", text="Position cycleur")
    treeview2.heading("#2", text="Capacité")
    
    
    # Placer le treeview dans le cadre principal
    treeview2.pack(fill="both", expand=True)
    
    
    # Définir la largeur des colonnes
    treeview2.column("#0", width=col_width1)  # Largeur de la colonne Numéro de Série
    treeview2.column("#1", width=col_width1)  # Largeur de la colonne Capacité
    treeview2.column("#2", width=col_width1)  # Largeur de la colonne Capacité
    
    label2 = tk.Label(left_frame, text="Bilan test ok par baies")
    label2.pack()
    
    # Créer l'ecran pour display les informations 
    treeview3 = ttk.Treeview(left_frame, columns=("Nombre de test OK",))
    
    treeview3.heading("#0", text="Baie")
    treeview3.heading("#1", text="Nombre de test OK")
    
    # Placer le treeview dans le cadre principal
    treeview3.pack(fill="both", expand=True)
    
    # Définir la largeur des colonnes
    treeview3.column("#0", width=col_width1)  # Largeur de la colonne Numéro de Série
    treeview3.column("#1", width=col_width1)  # Largeur de la colonne Capacité
 
    # Bouton pour traiter les fichiers Excel
    process_button = tk.Button(left_frame, text="Traiter les fichiers Neware", command=resultats_batteries)
    process_button.pack(pady=10)
    
    police_gras = font.Font(family="Helvetica", size=14, weight="bold")
    label_resultat = tk.Label(left_frame, text="", font=police_gras)
    label_resultat.pack(pady=10)
    
    test_frame.pack()