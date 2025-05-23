import pandas as pd
import random


# Charger les données
questionnaire = pd.read_excel('Questionnaire bénévoles 2025 - 19e édition (réponses).xlsx')
schedule_raw = pd.read_csv('Tableau rempli - Feuille 1.csv', header=None)
questionnaire.to_csv("Questionnaire_bénévoles_2025.csv")

lignes_supprimees = questionnaire[questionnaire["Si oui, as tu quand même envie d'être bénévole ?"] == "Non !"]
questionnaire = questionnaire[questionnaire["Si oui, as tu quand même envie d'être bénévole ?"] != "Non !"]
print(f"{len(lignes_supprimees)} ligne(s) supprimée(s).")
print("Lignes supprimées :")
print(lignes_supprimees[["Prénom", "NOM"]])

# Détection des colonnes clés
cols = questionnaire.columns
dispo_col = next(c for c in cols if "Combien de temps" in c)
accept_col = next(c for c in cols if "Acceptes-tu de faire" in c)
sens_col = next(c for c in cols if "Tu te sens plutôt" in c)
sec_col = next(c for c in cols if "brevet de secourisme" in c)
bar_col = next(c for c in cols if c.startswith("Bar 14h-23h"))
catering_col = next(c for c in cols if c.startswith("Catering 17h-22h"))
billetterie_col = next(c for c in cols if c.startswith("Billetterie"))
ateliers_col = next(c for c in cols if c.startswith("Ateliers 14h 18h"))

# Prétraitement des bénévoles
df = questionnaire.copy()
# Nettoyer les disponibilités horaires
dispo_time_col = "Quelles sont tes disponibilités ?"

def parse_dispo_times(val):
    if pd.isna(val):
        return set()
    return set(map(str.strip, str(val).split(',')))

df['dispo_hours'] = df[dispo_time_col].apply(parse_dispo_times)
df['dispo_max'] = df[dispo_col].map({
    '1h': 1,
    '2h !': 2,
    "Pas de limite j'adore le travail !": 5
})
df['gros_biscotos'] = df[accept_col].str.contains("gros biscotos", na=False) | \
                     df[sens_col].str.contains("Gros biscotos", na=False)
df['secouriste'] = df[sec_col].str.contains("Oui", na=False)
df['accept_any'] = ~df[accept_col].str.startswith("Non", na=False)
df['service_bar'] = df[bar_col].str.contains("Service au bar", na=False)
df['vaisselle'] = df[bar_col].str.contains("Vaisselle", na=False)
df['install_range'] = df[catering_col].str.contains("Installation|rangement", na=False)
df['install_ateliers'] = df[ateliers_col].str.contains("Installation|Rangement", na=False)
df['maquillage'] = df[ateliers_col].str.contains("Maquillages enfants", na=False)
df['vente_cacoins'] = df[billetterie_col].notna()

vols = df[['Prénom','NOM','dispo_max','gros_biscotos','secouriste',
           'accept_any','service_bar','vaisselle',
           'install_range','install_ateliers','maquillage','vente_cacoins', 'dispo_hours']].copy()
vols['assigned_hours'] = 0


# Heures et quotas
hours = schedule_raw.iloc[1,1:].tolist()
n_cols = schedule_raw.shape[1]
# Lignes contenant des quotas numériques
quota_rows = schedule_raw.iloc[:,1:].applymap(lambda x: pd.to_numeric(x, errors='coerce')).notna().sum(axis=1)
quota_rows = quota_rows[quota_rows>0].index.tolist()

# Copie du planning pour remplissage
filled = schedule_raw.copy()

# historique des postes par bénévole
assignments = {idx: [] for idx in vols.index}
# Historique : heures déjà assignées
assigned_times = {idx: set() for idx in vols.index}

def select(mask, hour, n, require_sec=False, heavy=False, role_name=None):
    avail = vols[vols['assigned_hours'] < vols['dispo_max']]
    avail = avail[avail['dispo_hours'].apply(lambda hset: hour in hset)]
    avail = avail[[hour not in assigned_times[idx] for idx in avail.index]]
    if mask is not None:
        avail = avail[mask.loc[avail.index]]
    if heavy:
        avail = avail[avail['gros_biscotos']]
    chosen = []

    if require_sec:
        secs = avail[avail['secouriste']]
        if not secs.empty:
            idx = secs.sample(1).index[0]
            chosen.append(idx)
            avail = avail.drop(idx)

    rem = n - len(chosen)
    
    # Étape 1 : bénévoles déjà sur ce poste
    prior = [idx for idx in avail.index if role_name in assignments[idx]]
    
    while rem > 0 and (prior or not avail.empty):
        if prior:
            idx = prior.pop(0)
        else:
            zero = avail[avail['assigned_hours']==0]
            if not zero.empty:
                idx = zero.sample(1).index[0]
            else:
                minh = avail['assigned_hours'].min()
                idx = avail[avail['assigned_hours']==minh].sample(1).index[0]
        chosen.append(idx)
        avail = avail.drop(idx, errors='ignore')
        rem -= 1

    for idx in chosen:
        vols.at[idx, 'assigned_hours'] += 1
        assigned_times[idx].add(hour)
        if role_name:
            assignments[idx].append(role_name)

    return vols.loc[chosen]


# Déduire les heures déjà assignées manuellement dans le fichier CSV
for j in range(1, n_cols):  # Colonnes horaires (heures)
    hour = hours[j-1]       # Récupère l'heure correspondant à la colonne
    for i in range(schedule_raw.shape[0]):  # Toutes les lignes (rôles)
        cell = schedule_raw.iat[i, j]
        if pd.notna(cell):
            name = str(cell).strip()
            if name == "" or name.isdigit():
                continue
            # Tenter de trouver le prénom/nom dans la liste des volontaires
            matched = vols[(vols['Prénom'] + " " + vols['NOM']).str.strip().str.lower() == name.lower()]
            if not matched.empty:
                idx = matched.index[0]
                # Incrémente assigned_hours, sans dépasser le max
                vols.at[idx, 'assigned_hours'] = min(vols.at[idx, 'assigned_hours'] + 1, vols.at[idx, 'dispo_max'])
                # Supprime cette heure de ses disponibilités
                vols.at[idx, 'dispo_hours'].discard(hour)
                assigned_times[idx].add(hour)



# Remplissage complet sans écraser les cellules déjà remplies
for qr in quota_rows:
    role = str(schedule_raw.iat[qr, 0]).strip().lower()
    for j in range(1, n_cols):
        cell = schedule_raw.iat[qr, j]
        # Ne rien faire si quota vide ou nul
        if pd.notna(cell) and int(cell) > 0:
            hour = hours[j - 1]
            require_sec = 'vente cacoins' in role
            heavy = ('installation ateliers et rangement' in role) or ('service au bar' in role and hour == '23h-00h')

            # Définir le masque selon le rôle
            if 'service au bar' in role:
                mask = vols['service_bar'] | vols['accept_any']
            elif 'vaisselle' in role:
                mask = vols['vaisselle'] | vols['accept_any']
            elif 'installation/rangement' in role:
                mask = vols['install_range'] | vols['accept_any']
            elif 'installation ateliers et rangement' in role:
                mask = vols['install_ateliers'] | vols['accept_any']
            elif 'anaïs et ses pinceaux' in role:
                mask = vols['maquillage'] | vols['accept_any']
            elif 'vente cacoins' in role:
                mask = vols['vente_cacoins'] | vols['accept_any']
            else:
                mask = vols['accept_any']

            # Vérifier combien de personnes sont déjà assignées
            already_filled = 0
            for k in range(1, 10):  # on vérifie jusqu'à 9 lignes en dessous
                content = filled.iat[qr + k, j] if (qr + k < filled.shape[0]) else None
                if pd.isna(content) or str(content).strip() == "":
                    break
                else:
                    already_filled += 1

            n = int(cell) - already_filled
            if n <= 0:
                continue  # rien à remplir

            selected = select(mask, hour, n, require_sec=require_sec, heavy=heavy)
            line_offset = 1
            for _, vol in selected.iterrows():
                # Trouver la première ligne vide sous la ligne quota
                while (qr + line_offset < filled.shape[0]) and pd.notna(filled.iat[qr + line_offset, j]) and str(filled.iat[qr + line_offset, j]).strip() != "":
                    line_offset += 1
                filled.iat[qr + line_offset, j] = f"{vol['Prénom']} {vol['NOM']}"
                line_offset += 1

# Enregistrement et affichage final
output_path = 'Tableau_final_benevoles_complet.xlsx'
filled.to_excel(output_path, index=False)
filled.to_csv("Tableau_final_benevoles_complet.csv")
print(f"✅ Planning complet mis à jour : {output_path}")
