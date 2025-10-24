# -*- coding: utf-8 -*-
"""
Created on Fri Oct 17 16:49:49 2025

@author: User
"""



def recycle_cellule(self):
    type_obj="batterie"
    cause=self.cause_combo.get()
    if not cause:
        messagebox.showerror("Pas de cause","Veuillez selectionner une cause !")
        return
    numero_serie_batt=self.r_entry_batt.get()
    conn=self.db_manager.connect()
    cursor=conn.cursor()
    query="select reference_produit_voltr,poids from produit_voltr where numero_serie_produit =%s "
    param=(numero_serie_batt,)
    cursor.execute(query,param)
    row=cursor.fetchone()
    reference_batt=row[0]
    poids=row[1]
    
    df_cyclage = pd.read_excel(EXCEL_PATH,sheet_name="Cyclage",header=1)
    row = df_cyclage[
        df_cyclage["Nom_modele"] == reference_batt]
    seuils = row.iloc[0]
    dest_recyclage = str(seuils["Recyclage"])
    type_fut=dest_recyclage
    cursor.execute("SELECT id_fut from fut_recyclage WHERE exutoire=%s and etat_fut=%s limit 1",(type_fut,"en cours"))
    fut=cursor.fetchall()
    if not fut : 
        messagebox.showerror("Erreur !", "Aucun fut d'exutoire eo_org_mtl n'est ouvert")
        return
    else :
        cursor.execute("SELECT id_fut,poids from fut_recyclage WHERE exutoire=%s and etat_fut=%s limit 1",(type_fut,"en cours"))
        data_fut=cursor.fetchall()
        fut,poids_fut=data_fut[0]
        poids_tot= poids_fut + poids 
        cursor.execute("UPDATE recyclage set fut_recyclage= SET poids=%s where id_fut=%s",(poids_tot,fut))
        emplacement = 'fut'+' '+str(fut)
    query_recy="Insert into recyclage numero_serie= %s, type_objet= %s, id_fut= %s, sur_site= %s, date_rebut= NOW(),cause= %s "
    param_recy=(numero_serie_batt,type_obj,fut, "oui", cause)
    cursor.execute(query_recy,param_recy)
    query_sp="UPDATE suivi_production set recyclage=1, date_recyclage= NOW() where numero_serie= %s"
    param_sp=(numero_serie_batt,)
    cursor.execute(query_sp,param_sp)
    conn.commit()
    messagebox.showinfo("Recyclage reussi",f"La batterie {numero_serie_batt} recycléé dans un fut {type_fut} : {emplacement}")
    
    cursor.close()
    conn.close()