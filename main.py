import requests

# Exemple de variables (à adapter)
USER_ID = "ton_user_id_ou_email"
ACCESS_TOKEN = "ton_token_d_acces_valide"

headers = {
    "Authorization": f"Bearer {ACCESS_TOKEN}",
    "Content-Type": "application/json"
}

def get_folders():
    url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/mailFolders?$top=100"
    response = requests.get(url, headers=headers)
    print(f"[DEBUG] Status code: {response.status_code}")
    print(f"[DEBUG] Response text: {response.text}")

    if response.status_code != 200:
        raise Exception(f"Erreur API: {response.status_code} - {response.text}")

    try:
        return response.json()
    except Exception as e:
        print(f"Erreur lors du parsing JSON: {e}")
        print(f"Contenu reçu: {response.text}")
        raise

def get_folder_ids(keywords):
    existing = get_folders()
    folder_ids = {}
    for folder in existing.get("value", []):
        if folder["displayName"] in keywords:
            folder_ids[folder["displayName"]] = folder["id"]
    return folder_ids

def trier_emails():
    # Exemple de mots clés pour trier
    keywords = {"Factures", "Offres", "Important"}
    folder_ids = get_folder_ids(keywords.keys())
    print(f"Folder IDs trouvés: {folder_ids}")
    # Ici tu mettra la suite de ton code de tri

if __name__ == "__main__":
    try:
        print("✅ Connexion réussie (TEST)")
        trier_emails()
    except Exception as e:
        print(f"Erreur détectée : {e}")
