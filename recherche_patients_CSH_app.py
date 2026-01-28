"""
🏠 RECHERCHE DE PATIENTS - CELLULE SANTÉ HABITAT
================================================
Application graphique de recherche dans les bilans suivi (2015-2026)
Version corrigée - Python 3.11
"""

import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import sys
import os


class RechercheCSH:
    """Classe de recherche de patients dans les bilans CSH"""
    
    def __init__(self, dossier=None):
        self.dossier = Path(dossier) if dossier else None
        self.donnees = pd.DataFrame()
        self.fichiers_charges = []
        self.log_messages = []
        
    def normaliser_colonnes(self, df, nom_fichier):
        """Normalise les noms de colonnes"""
        mapping = {}
        
        for col in df.columns:
            col_str = str(col).lower().replace('\n', ' ').strip()
            
            if 'nom patient' in col_str or col_str == 'nom':
                mapping[col] = 'NOM_PATIENT'
            elif 'nom medecin' in col_str or ('médecin' in col_str and 'demandeur' in col_str):
                mapping[col] = 'MEDECIN'
            elif 'date' in col_str and 'audit' in col_str:
                mapping[col] = 'DATE_AUDIT'
            elif col_str == 'cmei':
                mapping[col] = 'CMEI'
            elif 'audit' in col_str and 'code' not in col_str:
                mapping[col] = 'AUDIT'
            elif 'code' in col_str and 'affaire' in col_str:
                mapping[col] = 'CODE_AFFAIRE'
            elif 'remarques' in col_str:
                mapping[col] = 'REMARQUES'
        
        df = df.rename(columns=mapping)
        
        # Extraire l'année du nom de fichier
        annee = "N/A"
        for a in range(2015, 2030):
            if str(a) in nom_fichier:
                annee = str(a)
                break
        
        df['ANNEE'] = annee
        df['FICHIER'] = nom_fichier
        return df
    
    def charger_fichier(self, chemin_complet):
        """Charge un fichier Excel"""
        chemin = Path(chemin_complet)
        nom_fichier = chemin.name  # FIX: Définir nom_fichier au début
        
        if not chemin.exists():
            self.log_messages.append(f"⚠️ {nom_fichier}: fichier non trouvé")
            return None
        
        try:
            xl = pd.ExcelFile(chemin)
            
            # Chercher la bonne feuille
            feuille_cible = None
            for f in xl.sheet_names:
                f_lower = f.lower()
                if 'bilan csh' in f_lower or 'bilan paris' in f_lower or 'gratuit' in f_lower:
                    feuille_cible = f
                    break
            
            if feuille_cible is None:
                feuille_cible = xl.sheet_names[0]
            
            df = pd.read_excel(chemin, sheet_name=feuille_cible)
            
            # Supprimer les colonnes dupliquées
            df = df.loc[:, ~df.columns.duplicated()]
            
            # Nettoyer les lignes vides
            col_nom = [c for c in df.columns if 'nom' in str(c).lower().replace('\n', ' ')]
            if col_nom:
                df = df.dropna(subset=[col_nom[0]])
            
            df = self.normaliser_colonnes(df, nom_fichier)
            return df
            
        except Exception as e:
            self.log_messages.append(f"❌ Erreur {nom_fichier}: {e}")
            return None
    
    def charger_tous_les_fichiers(self, liste_fichiers):
        """Charge tous les fichiers d'une liste"""
        self.log_messages = []
        self.log_messages.append(f"📁 Chargement de {len(liste_fichiers)} fichier(s)\n")
        
        tous_les_df = []
        
        for fichier in liste_fichiers:
            df = self.charger_fichier(fichier)
            nom_fichier = Path(fichier).name
            if df is not None:
                tous_les_df.append(df)
                self.fichiers_charges.append(nom_fichier)
                self.log_messages.append(f"✅ {nom_fichier}: {len(df)} patients")
            else:
                self.log_messages.append(f"⚠️ {nom_fichier}: non chargé")
        
        if tous_les_df:
            # Garder les colonnes importantes
            colonnes = ['NOM_PATIENT', 'MEDECIN', 'CMEI', 'AUDIT', 'DATE_AUDIT', 
                       'CODE_AFFAIRE', 'REMARQUES', 'ANNEE', 'FICHIER']
            
            dfs_nettoyes = []
            for df in tous_les_df:
                df = df.loc[:, ~df.columns.duplicated()]
                cols_existantes = [c for c in colonnes if c in df.columns]
                dfs_nettoyes.append(df[cols_existantes])
            
            self.donnees = pd.concat(dfs_nettoyes, ignore_index=True)
            
            self.log_messages.append(f"\n✅ BASE CHARGÉE: {len(self.donnees)} interventions")
            self.log_messages.append(f"📁 Fichiers: {len(self.fichiers_charges)}")
            return True
        else:
            self.log_messages.append("\n❌ Aucun fichier valide chargé!")
            return False
    
    def rechercher(self, nom):
        """Recherche un patient par nom"""
        if self.donnees.empty or not nom.strip():
            return pd.DataFrame()
        
        nom_clean = nom.strip().upper()
        mask = self.donnees['NOM_PATIENT'].astype(str).str.upper().str.contains(nom_clean, na=False)
        return self.donnees[mask].sort_values('ANNEE', ascending=False)


class Application(tk.Tk):
    """Interface graphique principale"""
    
    def __init__(self):
        super().__init__()
        
        self.title("🏠 Recherche Patients - Cellule Santé Habitat")
        self.geometry("900x700")
        self.configure(bg='#f0f4f8')
        
        # Variables
        self.dossier = None  # FIX: Initialiser la variable
        self.fichiers_selectionnes = []
        self.app = RechercheCSH()
        
        # Centrer la fenêtre
        self.center_window()
        
        self.create_widgets()
        
    def center_window(self):
        """Centre la fenêtre sur l'écran"""
        self.update_idletasks()
        width = 900
        height = 700
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')
    
    def create_widgets(self):
        """Crée les widgets de l'interface"""
        
        # Style
        style = ttk.Style()
        style.theme_use('clam')
        
        # En-tête
        header_frame = tk.Frame(self, bg='#2c5282', height=100)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(header_frame, text="🏠 Cellule Santé Habitat", 
                              font=('Segoe UI', 20, 'bold'), bg='#2c5282', fg='white')
        title_label.pack(pady=15)
        
        subtitle_label = tk.Label(header_frame, text="Recherche de patients dans les bilans suivi (2015-2026)", 
                                 font=('Segoe UI', 11), bg='#2c5282', fg='#bee3f8')
        subtitle_label.pack()
        
        # Frame principal
        main_frame = tk.Frame(self, bg='#f0f4f8', padx=25, pady=25)
        main_frame.pack(fill='both', expand=True)
        
        # Section 1: Sélection des fichiers
        files_frame = tk.LabelFrame(main_frame, text="📁 Sélection des fichiers", 
                                    font=('Segoe UI', 11, 'bold'), bg='#f0f4f8', 
                                    fg='#2c5282', padx=15, pady=15)
        files_frame.pack(fill='x', pady=(0, 20))
        
        # Bouton pour sélectionner un dossier
        btn_frame = tk.Frame(files_frame, bg='#f0f4f8')
        btn_frame.pack(fill='x', pady=5)
        
        self.btn_dossier = tk.Button(btn_frame, text="📂 Sélectionner un dossier", 
                                     command=self.select_folder, font=('Segoe UI', 10),
                                     bg='#4299e1', fg='white', relief='flat', 
                                     padx=20, pady=10, cursor='hand2')
        self.btn_dossier.pack(side='left', padx=(0, 10))
        
        self.btn_fichiers = tk.Button(btn_frame, text="📄 Sélectionner des fichiers", 
                                      command=self.select_files, font=('Segoe UI', 10),
                                      bg='#48bb78', fg='white', relief='flat', 
                                      padx=20, pady=10, cursor='hand2')
        self.btn_fichiers.pack(side='left')
        
        # Label pour afficher le dossier/fichiers sélectionnés
        self.files_info_label = tk.Label(files_frame, text="Aucun fichier sélectionné", 
                                        font=('Segoe UI', 9), bg='#f0f4f8', 
                                        fg='#718096', anchor='w', justify='left')
        self.files_info_label.pack(fill='x', pady=(10, 0))
        
        # Bouton charger
        self.btn_charger = tk.Button(files_frame, text="📥 Charger les fichiers", 
                                     command=self.charger_fichiers, font=('Segoe UI', 11, 'bold'),
                                     bg='#38a169', fg='white', relief='flat', 
                                     padx=25, pady=12, state='disabled', cursor='hand2')
        self.btn_charger.pack(pady=(15, 0))
        
        # Section 2: Recherche
        search_frame = tk.LabelFrame(main_frame, text="🔍 Recherche de patient", 
                                     font=('Segoe UI', 11, 'bold'), bg='#f0f4f8', 
                                     fg='#2c5282', padx=15, pady=15)
        search_frame.pack(fill='x', pady=(0, 20))
        
        search_input_frame = tk.Frame(search_frame, bg='#f0f4f8')
        search_input_frame.pack(fill='x', pady=5)
        
        tk.Label(search_input_frame, text="Nom du patient:", 
                font=('Segoe UI', 10), bg='#f0f4f8').pack(side='left', padx=(0, 10))
        
        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(search_input_frame, textvariable=self.search_var, 
                                    font=('Segoe UI', 12), relief='solid', bd=1, width=40)
        self.search_entry.pack(side='left', fill='x', expand=True, ipady=8)
        self.search_entry.bind('<Return>', lambda e: self.rechercher())
        
        self.btn_search = tk.Button(search_input_frame, text="🔎 Rechercher", 
                                   command=self.rechercher, font=('Segoe UI', 10, 'bold'),
                                   bg='#4299e1', fg='white', relief='flat', 
                                   padx=25, pady=8, state='disabled', cursor='hand2')
        self.btn_search.pack(side='right', padx=(10, 0))
        
        # Section 3: Résultats
        result_frame = tk.LabelFrame(main_frame, text="📋 Résultats", 
                                     font=('Segoe UI', 11, 'bold'), bg='#f0f4f8', 
                                     fg='#2c5282', padx=15, pady=15)
        result_frame.pack(fill='both', expand=True)
        
        # Zone de texte avec scrollbar
        self.result_text = scrolledtext.ScrolledText(result_frame, font=('Consolas', 10), 
                                                     wrap='word', relief='solid', bd=1,
                                                     bg='white', height=15)
        self.result_text.pack(fill='both', expand=True)
        
        # Message initial
        self.result_text.insert(tk.END, "═" * 70 + "\n")
        self.result_text.insert(tk.END, "BIENVENUE DANS L'APPLICATION DE RECHERCHE CSH\n")
        self.result_text.insert(tk.END, "═" * 70 + "\n\n")
        self.result_text.insert(tk.END, "📌 Pour commencer:\n")
        self.result_text.insert(tk.END, "   1. Sélectionnez un dossier ou des fichiers Excel\n")
        self.result_text.insert(tk.END, "   2. Cliquez sur 'Charger les fichiers'\n")
        self.result_text.insert(tk.END, "   3. Recherchez vos patients\n\n")
        self.result_text.insert(tk.END, "✨ L'application recherche dans les bilans de 2015 à 2026\n")
        
        # Barre de statut
        self.status_var = tk.StringVar(value="Prêt - Sélectionnez des fichiers pour commencer")
        status_bar = tk.Label(self, textvariable=self.status_var, font=('Segoe UI', 9),
                             bg='#e2e8f0', anchor='w', padx=15, pady=8, relief='sunken')
        status_bar.pack(fill='x', side='bottom')
    
    def select_folder(self):
        """Ouvre le dialogue de sélection de dossier"""
        folder = filedialog.askdirectory(title="Sélectionner le dossier contenant les fichiers Excel")
        if folder:
            self.dossier = folder
            # Lister tous les fichiers Excel du dossier
            fichiers_excel = []
            for ext in ['*.xlsx', '*.xls']:
                fichiers_excel.extend(Path(folder).rglob(ext))
            
            self.fichiers_selectionnes = [str(f) for f in fichiers_excel]
            
            if self.fichiers_selectionnes:
                self.files_info_label.config(
                    text=f"✅ Dossier: {folder}\n📄 {len(self.fichiers_selectionnes)} fichier(s) Excel trouvé(s)",
                    fg='#2d3748'
                )
                self.btn_charger.config(state='normal')
                self.status_var.set(f"Dossier sélectionné - {len(self.fichiers_selectionnes)} fichiers trouvés")
            else:
                messagebox.showwarning("Attention", "Aucun fichier Excel trouvé dans ce dossier")
                self.files_info_label.config(text="⚠️ Aucun fichier Excel trouvé", fg='#e53e3e')
    
    def select_files(self):
        """Ouvre le dialogue de sélection de fichiers multiples"""
        fichiers = filedialog.askopenfilenames(
            title="Sélectionner les fichiers Excel",
            filetypes=[("Fichiers Excel", "*.xlsx *.xls"), ("Tous les fichiers", "*.*")]
        )
        if fichiers:
            self.fichiers_selectionnes = list(fichiers)
            noms_fichiers = [Path(f).name for f in fichiers[:3]]
            texte_fichiers = ", ".join(noms_fichiers)
            if len(fichiers) > 3:
                texte_fichiers += f" ... (+{len(fichiers)-3} autres)"
            
            self.files_info_label.config(
                text=f"✅ {len(fichiers)} fichier(s) sélectionné(s):\n{texte_fichiers}",
                fg='#2d3748'
            )
            self.btn_charger.config(state='normal')
            self.status_var.set(f"{len(fichiers)} fichier(s) sélectionné(s)")
    
    def charger_fichiers(self):
        """Charge les fichiers Excel"""
        if not self.fichiers_selectionnes:
            messagebox.showwarning("Attention", "Veuillez d'abord sélectionner des fichiers")
            return
        
        self.status_var.set("Chargement en cours...")
        self.update()
        
        success = self.app.charger_tous_les_fichiers(self.fichiers_selectionnes)
        
        # Afficher le log
        self.result_text.delete('1.0', tk.END)
        self.result_text.insert(tk.END, "═" * 70 + "\n")
        self.result_text.insert(tk.END, "CHARGEMENT DES FICHIERS\n")
        self.result_text.insert(tk.END, "═" * 70 + "\n\n")
        for msg in self.app.log_messages:
            self.result_text.insert(tk.END, msg + "\n")
        
        if success:
            self.btn_search.config(state='normal')
            self.search_entry.focus()
            self.status_var.set(f"✅ Base chargée: {len(self.app.donnees)} interventions - Prêt pour la recherche")
            messagebox.showinfo("Succès", 
                              f"Base de données chargée avec succès!\n\n"
                              f"• {len(self.app.donnees)} interventions\n"
                              f"• {len(self.app.fichiers_charges)} fichiers")
        else:
            self.status_var.set("❌ Erreur de chargement - Vérifiez les fichiers")
            messagebox.showerror("Erreur", "Erreur lors du chargement des fichiers.\nVérifiez les logs.")
    
    def rechercher(self):
        """Lance la recherche"""
        if self.app.donnees.empty:
            messagebox.showwarning("Attention", "Veuillez d'abord charger les fichiers")
            return
        
        nom = self.search_var.get().strip()
        if not nom:
            messagebox.showwarning("Attention", "Veuillez entrer un nom à rechercher")
            return
        
        resultats = self.app.rechercher(nom)
        
        self.result_text.delete('1.0', tk.END)
        self.result_text.insert(tk.END, "═" * 70 + "\n")
        self.result_text.insert(tk.END, f"RECHERCHE: {nom.upper()}\n")
        self.result_text.insert(tk.END, "═" * 70 + "\n\n")
        
        if len(resultats) == 0:
            self.result_text.insert(tk.END, "🆕 AUCUN RÉSULTAT TROUVÉ\n\n")
            self.result_text.insert(tk.END, f"Aucune intervention trouvée pour '{nom}'\n")
            self.result_text.insert(tk.END, "→ Il s'agit probablement d'une nouvelle demande à traiter\n\n")
            self.result_text.insert(tk.END, "💡 Vérifiez l'orthographe ou essayez avec moins de lettres\n")
            self.status_var.set(f"Aucun résultat pour '{nom}'")
        else:
            nom_patient = resultats.iloc[0]['NOM_PATIENT']
            self.result_text.insert(tk.END, f"✅ PATIENT TROUVÉ: {nom_patient}\n")
            self.result_text.insert(tk.END, f"   {len(resultats)} intervention(s) enregistrée(s)\n\n")
            self.result_text.insert(tk.END, "─" * 70 + "\n")
            
            for idx, (_, row) in enumerate(resultats.iterrows(), 1):
                self.result_text.insert(tk.END, f"\n📌 INTERVENTION #{idx}\n")
                self.result_text.insert(tk.END, f"   📅 Année: {row.get('ANNEE', 'N/A')}\n")
                self.result_text.insert(tk.END, f"   👨‍⚕️ Médecin: {row.get('MEDECIN', 'N/A')}\n")
                
                cmei = row.get('CMEI', '')
                audit = row.get('AUDIT', '')
                if pd.notna(cmei) and str(cmei) != '':
                    self.result_text.insert(tk.END, f"   🏠 CMEI: {cmei}\n")
                if pd.notna(audit) and str(audit) != '':
                    self.result_text.insert(tk.END, f"   📊 Audit: {audit}\n")
                
                code = row.get('CODE_AFFAIRE', '')
                if pd.notna(code) and str(code) != '':
                    self.result_text.insert(tk.END, f"   📝 Code affaire: {code}\n")
                
                date_audit = row.get('DATE_AUDIT', '')
                if pd.notna(date_audit) and str(date_audit) != '':
                    self.result_text.insert(tk.END, f"   📆 Date audit: {date_audit}\n")
                
                remarques = row.get('REMARQUES', '')
                if pd.notna(remarques) and str(remarques) not in ['', 'nan']:
                    self.result_text.insert(tk.END, f"   💬 Remarques: {remarques}\n")
                
                fichier = row.get('FICHIER', '')
                if fichier:
                    self.result_text.insert(tk.END, f"   📄 Fichier: {fichier}\n")
                
                if idx < len(resultats):
                    self.result_text.insert(tk.END, "\n" + "─" * 70 + "\n")
            
            self.status_var.set(f"✅ {len(resultats)} intervention(s) trouvée(s) pour '{nom}'")


def main():
    """Point d'entrée de l'application"""
    try:
        app = Application()
        app.mainloop()
    except Exception as e:
        messagebox.showerror("Erreur critique", 
                           f"Une erreur est survenue:\n{str(e)}\n\n"
                           f"Veuillez contacter le support technique.")
        sys.exit(1)


if __name__ == "__main__":
    main()
