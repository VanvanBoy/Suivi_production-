# -*- coding: utf-8 -*-
"""
Created on Thu Aug 21 15:18:57 2025

@author: User

Point d'entrée de l'application "Suivi de production".
Ce fichier assemble la fenêtre principale (StockApp) a partir de tous les
modules de fonctionnalités (un module par onglet/étape de production).

C'est CE fichier qu'il faut viser pour générer l'exécutable, par exemple :
    pyinstaller --onefile --noconsole main.py
"""
from tkinter import ttk
from ttkthemes import ThemedTk

from config import STAGE_TO_DBCOL, ALLOWED_STAGE_KEYS
from db_manager import DBManager

from mixin_common import CommonMixin
from tab_tri_test import TabTriTestMixin
from tab_banc_somfy import TabBancSomfyMixin
from tab_picking import TabPickingMixin
from tab_pack import TabPackMixin
from tab_nappe import TabNappeMixin
from tab_bms import TabBmsMixin
from tab_wrap import TabWrapMixin
from tab_fermeture import TabFermetureMixin
from tab_emballage import TabEmballageMixin
from tab_capa import TabCapaMixin
from tab_recherche import TabRechercheMixin
from tab_expedition import TabExpeditionMixin
from tab_finligne import TabFinligneMixin
from tab_recyclage import TabRecyclageMixin


class StockApp(
    ThemedTk,
    CommonMixin,
    TabTriTestMixin,
    TabBancSomfyMixin,
    TabPickingMixin,
    TabPackMixin,
    TabNappeMixin,
    TabBmsMixin,
    TabWrapMixin,
    TabFermetureMixin,
    TabEmballageMixin,
    TabCapaMixin,
    TabRechercheMixin,
    TabExpeditionMixin,
    TabFinligneMixin,
    TabRecyclageMixin,
):

    def __init__(self):
        super().__init__(theme="xpnative")
        self.title("Suivi de production")
        self.geometry("1000x650")

        self.refresh_ms = 10_000  # 10 secondes
        self._tab_refresh_job = None
        self._refreshing = False

        # annuler proprement à la fermeture
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.focus_targets = {}  # stage -> widget qui doit recevoir le focu

        self.db_manager = DBManager()
        self.picking_file_path = None

        # styles
        self.style = ttk.Style()
        self.style.configure("TLabel", font=('Segoe UI', 11), padding=4)
        self.style.configure("TButton", font=('Segoe UI', 11), padding=6)
        self.style.configure("TEntry", font=('Segoe UI', 11))
        self.style.configure("TCombobox", font=('Segoe UI', 11))
        self.style.configure("Danger.TButton", foreground="black", background="#F06F65")
        self.style.map("Danger.TButton", background=[('active', '#e55b50')], foreground=[('disabled', 'black')])
        self.style.configure("Good.TButton", foreground="black", background="#0ED329")
        self.style.map("Good.TButton", background=[('active', '#0ED329')], foreground=[('disabled', 'black')])

        self.selected_model = None
        self.stage_order = None

        self.STAGE_TO_DBCOL = STAGE_TO_DBCOL
        self.ALLOWED_STAGE_KEYS = ALLOWED_STAGE_KEYS

        self._show_model_selector_and_build()


if __name__ == "__main__":
    app = StockApp()
    app.mainloop()
