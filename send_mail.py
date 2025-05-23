import pandas as pd

questionnaire = pd.read_excel('Questionnaire bénévoles 2025 - 19e édition (réponses).xlsx')
schedule_raw = pd.read_excel('Tableau rempli(1).xlsx')
schedule_raw.to_csv("pour_claud.csv")


def extract_volunteer_missions(schedule_raw: pd.DataFrame, questionnaire: pd.DataFrame) -> list:
    volunteers = {}

    # Heure correspond à ligne 1 (index 1) et colonnes 1+
    hours = schedule_raw.iloc[1, 1:].tolist()
    
    # Balayer chaque ligne du planning à partir de la ligne 2 (les missions)
    for i in range(2, schedule_raw.shape[0]):
        mission = schedule_raw.iat[i, 0]  # Nom de la mission
        for j in range(1, schedule_raw.shape[1]):
            hour = hours[j-1]
            cell = schedule_raw.iat[i, j]
            if pd.isna(cell):
                continue
            names = [name.strip() for name in str(cell).split(",") if name.strip()]
            for name in names:
                if name == "" or name.isdigit():
                    continue
                name_lower = name.lower().strip()
                if name_lower not in volunteers:
                    volunteers[name_lower] = {
                        "name": name,
                        "email": None,
                        "mission": []
                    }
                volunteers[name_lower]["mission"].append((mission, hour))

    # Nettoyage des noms dans la df questionnaire
    questionnaire["full_name"] = (questionnaire["Prénom"].str.strip() + " " + questionnaire["NOM"].str.strip()).str.lower()

    # Associer les emails
    for v in volunteers.values():
        name_lower = v["name"].lower().strip()
        match = questionnaire[questionnaire["full_name"] == name_lower]
        if not match.empty:
            v["email"] = match.iloc[0]["Adresse mail"]

    return list(volunteers.values())

if __name__ == "__main__":
    print(questionnaire["NOM"].head(3))
    print(questionnaire["Prénom"].head(3))
    print(questionnaire["Adresse mail"].head(3))
    print(extract_volunteer_missions(schedule_raw, questionnaire))