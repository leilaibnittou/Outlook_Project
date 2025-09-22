import os
from msal import ConfidentialClientApplication
import requests
import re
import sys

# -------------------------------
# VARIABLE UTILISATEUR / SERVICE
# -------------------------------
USER_ID = os.environ.get("OUTLOOK_USER_ID")
if not USER_ID:
    print("❌ Erreur : variable d'environnement OUTLOOK_USER_ID manquante")
    sys.exit(1)

# -------------------------------
# ENVIRONNEMENT : TEST ou PROD
# -------------------------------
ENV = os.environ.get("APP_ENV", "TEST")

if ENV == "PROD":
    CLIENT_ID = os.environ.get("PROD_CLIENT_ID")
    CLIENT_SECRET = os.environ.get("PROD_CLIENT_SECRET")
    TENANT_ID = os.environ.get("PROD_TENANT_ID")
else:
    CLIENT_ID = os.environ.get("TEST_CLIENT_ID")
    CLIENT_SECRET = os.environ.get("TEST_CLIENT_SECRET")
    TENANT_ID = os.environ.get("TEST_TENANT_ID")

SCOPES = ["https://graph.microsoft.com/.default"]

# -------------------------------
# Vérification des variables essentielles
# -------------------------------
if not CLIENT_ID or not CLIENT_SECRET or not TENANT_ID:
    print("❌ Erreur : variables d'environnement manquantes (CLIENT_ID / SECRET / TENANT_ID)")
    sys.exit(1)

# -------------------------------
# AUTHENTIFICATION
# -------------------------------
app = ConfidentialClientApplication(
    client_id=CLIENT_ID,
    client_credential=CLIENT_SECRET,
    authority=f"https://login.microsoftonline.com/{TENANT_ID}"
)

result = app.acquire_token_for_client(scopes=SCOPES)
token = result.get("access_token")
if not token:
    print("❌ Erreur d'authentification :", result)
    sys.exit(1)

headers = {
    "Authorization": f"Bearer {token}",
    "Accept": "application/json",
    "Content-Type": "application/json"
}

print(f"✅ Connexion réussie ({ENV})")

# -------------------------------
# Mots‑clés et dossiers cibles
# -------------------------------
keywords = {
    "P1": [r"\bp1\b"],
    "P2": [r"\bp2\b", r"\bcertificate\b"],
    "P3": [r"\bp3\b"],
    "P4": [r"\bp4\b"]
}
compiled_keywords = {folder: [re.compile(pat, re.IGNORECASE) for pat in pats] 
                     for folder, pats in keywords.items()}

# -------------------------------
# FONCTIONS UTILITAIRES (avec debug)
# -------------------------------
def get_folders():
    url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/mailFolders?$top=100"
    folders = []
    while url:
        try:
            resp = requests.get(url, headers=headers)
        except Exception as e:
            print("❌ Erreur lors de la requête HTTP:", e)
            sys.exit(1)

        # Debug : status code et corps
        print(f"🔍 Requête vers: {url}")
        print(f"🔍 Status Code: {resp.status_code}")
        print(f"🔍 Corps de la réponse: {resp.text[:500]}")  # Limite pour ne pas flooder trop

        if resp.status_code != 200:
            print("❌ Erreur HTTP : impossible d'obtenir la liste des dossiers.")
            sys.exit(1)

        try:
            data = resp.json()
        except ValueError:
            print("❌ Réponse non‑valide JSON.")
            print(resp.text)
            sys.exit(1)

        folders.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return folders

def get_folder_ids(targets):
    folder_ids = {}
    existing = get_folders()
    for f in targets:
        folder = next((x for x in existing if x["displayName"].lower() == f.lower()), None)
        if folder:
            folder_ids[f] = folder["id"]
        else:
            # Création du dossier si non existant
            resp = requests.post(
                f"https://graph.microsoft.com/v1.0/users/{USER_ID}/mailFolders",
                headers=headers,
                json={"displayName": f}
            )
            print(f"🔍 Création dossier {f}, status {resp.status_code}")
            print(f"🔍 Corps réponse création dossier: {resp.text}")
            if resp.status_code in (200, 201) and "id" in resp.json():
                folder_ids[f] = resp.json()["id"]
                print(f"📁 Dossier créé : {f}")
            else:
                print(f"❌ Erreur lors de la création du dossier {f}")
    return folder_ids

def get_emails():
    url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/mailFolders/Inbox/messages?$top=200&$orderby=receivedDateTime DESC"
    resp = requests.get(url, headers=headers)
    if resp.status_code != 200:
        print(f"❌ Erreur récupération emails, code: {resp.status_code}")
        print(resp.text)
        sys.exit(1)
    try:
        data = resp.json()
    except ValueError:
        print("❌ Réponse non‑valide JSON pour emails")
        print(resp.text)
        sys.exit(1)
    return data.get("value", [])

def delete_email(mail_id):
    url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/messages/{mail_id}"
    resp = requests.delete(url, headers=headers)
    print(f"🔍 Suppression email {mail_id}, status {resp.status_code}")
    return resp.status_code == 204

def move_email(mail_id, folder_id):
    url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/messages/{mail_id}/move"
    resp = requests.post(url, headers=headers, json={"destinationId": folder_id})
    print(f"🔍 Déplacement email {mail_id} → dossier {folder_id}, status {resp.status_code}")
    return resp.status_code in (200, 201)

# -------------------------------
# FONCTION PRINCIPALE
# -------------------------------
def trier_emails():
    folder_ids = get_folder_ids(keywords.keys())
    emails = get_emails()
    print(f"📨 {len(emails)} emails récupérés")

    seen_subjects = set()
    emails_unique = []
    for mail in emails:
        subject = (mail.get("subject") or "").strip().lower()
        mail_id = mail["id"]
        if subject in seen_subjects:
            if delete_email(mail_id):
                print(f"🗑️ Doublon supprimé : '{subject}'")
            else:
                print(f"⚠️ Erreur suppression doublon : '{subject}'")
        else:
            seen_subjects.add(subject)
            emails_unique.append(mail)

    for mail in emails_unique:
        subject = (mail.get("subject") or "")
        mail_id = mail["id"]
        target_folder = None
        for folder, regex_list in compiled_keywords.items():
            if any(regex.search(subject) for regex in regex_list):
                target_folder = folder
                break
        if target_folder:
            if move_email(mail_id, folder_ids.get(target_folder, "")):
                print(f"📌 '{subject}' déplacé vers {target_folder}")
            else:
                print(f"⚠️ Erreur déplacement '{subject}'")
        else:
            print(f"✉️ '{subject}' laissé dans Inbox")

    print("✅ Tri terminé")

# -------------------------------
# EXECUTION
# -------------------------------
if __name__ == "__main__":
    trier_emails()
