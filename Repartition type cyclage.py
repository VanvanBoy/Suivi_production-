# -*- coding: utf-8 -*-
"""
Created on Wed Dec 17 15:32:04 2025

@author: User
"""

import mysql.connector
from datetime import datetime, timedelta
import random
import math


# --- Connexion à la BDD ---
conn = mysql.connector.connect(
    host="34.77.226.40",
    user="Vanvan",
    password="VoltR99!",
    database="bdd_18122025",
    auth_plugin="mysql_native_password"
)

cursor = conn.cursor()

# --- Trouver la date d'hier ---
hier = datetime.now() - timedelta(days=1)
hier_str = hier.strftime('%Y-%m-%d')  # format compatible avec MySQL

# --- Récupérer les numéros de série ---
query = """
SELECT p.numero_serie_produit
FROM produit_voltr p
JOIN suivi_production s ON p.numero_serie_produit = s.numero_serie_batterie
WHERE s.date_fin_ligne = %s
  AND p.reference_produit_voltr = 'PPTR018AC'
"""
cursor.execute(query, (hier_str,))
result = cursor.fetchall()

# Extraire les numéros de série dans une liste
num_series = [row[0] for row in result]

if not num_series:
    print("Aucune batterie trouvée pour les critères.")
else:
    # --- Déterminer 10% aléatoires ---
    sample_size = math.ceil(len(num_series) * 0.10)  # au moins 10%, arrondi au supérieur
    sample_2_0 = random.sample(num_series, sample_size)
    sample_1_0 = list(set(num_series) - set(sample_2_0))

    # --- Update type_cyclage pour 2.0 ---
    update_query = "UPDATE produit_voltr SET type_cyclage = %s WHERE numero_serie_produit = %s"
    for ns in sample_2_0:
        cursor.execute(update_query, ('2.0', ns))

    # --- Update type_cyclage pour 1.0 ---
    for ns in sample_1_0:
        cursor.execute(update_query, ('1.0', ns))

    # --- Valider les changements ---
    conn.commit()
    print(f"Mise à jour terminée : {len(sample_2_0)} batteries en '2.0', {len(sample_1_0)} en '1.0'.")

# --- Fermeture ---
cursor.close()
conn.close()
