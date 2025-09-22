import os
import sys
import re
import requests
from msal import ConfidentialClientApplication

# -------------------------------
# 1. CONFIG & ENV
# -------------------------------
ENV = os.environ.get("APP_ENV", "TEST").upper()
USER_ID = os.environ.get("OUTLOOK_USER_ID")

if ENV == "PROD":
    CLIENT_ID = os.environ.get("PROD_CLIENT_ID")
    CLIENT_SECRET = os.environ.get("PROD_CLIENT_SECRET")
    TENANT_ID = os.environ.get("PROD_TENANT_ID")
else:
    CLIENT_ID = os.environ.get("TEST_CLIENT_ID")
    CLIENT_SECRET = os.environ.get("TEST_CLIENT_SECRET")
    TENANT_ID = os.environ.get("TEST_TENANT_ID")

# Validation
if not all([CLIENT_ID, CLIENT_SECRET, TENANT_ID, USER_ID]):
    print("❌ Erreur : Variables d’environnement manquantes.")
    sys.exit(1)

# -------------------------------
# 2. AUTHENTIFICATION
# -------------------------------
app = ConfidentialClientApplication(
    client_id=CLIENT_ID,
    client_credential=CLIENT_SECRET,
    authority=f"https://login.microsoftonline.com/{TENANT_ID}"
)

token_response = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
access_token = token_response.get("access_token")

if not access_token:
    print("❌ Erreur d’authentification :", token_response)
    sys.exit(1)

headers = {
    "Authorization": f"Bearer {access_token}",
    "Accept": "application/json",
    "Content-Type": "application/json"
}

print(f"✅ Connexion réussie ({ENV})")

# -------------------------------
# 3. MOTS-CLÉS & REGEX
# -------------------------------
keywords = {
    "P1": [r"\bp1\b"],
    "P2": [r"\bp2\b", r"\bcertificate\b"],
    "P3": [r"\bp3\b"],
    "P4": [r"\bp4\b"]
}

compiled_keywords = {
    folder: [re.compile(pat, re.IGNORECASE) for pat in patterns]
    for folder, patterns in keywords.items()
}

# -------------------------------
# 4. FONCTIONS
# -------------------------------
def get_folders():
    url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/mailFolders?$top=100"
    folders = []

    while url:
        resp = requests.get(url, headers=headers)
        if resp.status_code != 200:
            print(f"❌ Erreur API get_folders: {resp.status_code}")
            print("Réponse brute:", resp.text)
            sys.exit(1)

        data = resp.json()
        folders.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return folders

def get_folder_ids(target_folders):
    folder_ids = {}
    existing_folders = get_folders()

    for folder in target_folders:
        match = next((f for f in existing_folders if f["displayName"].lower() == folder.lower()), None)

        if match:
            folder_ids[folder] = match["id"]
        else:
            resp = requests.post(
                f"https://graph.microsoft.com/v1.0/users/{USER_ID}/mailFolders",
                headers=headers,
                json={"displayName": folder}
            )
            if resp.status_code not in [200, 201]:
                print(f"❌ Erreur création du dossier '{folder}' : {resp.status_code}")
                print(resp.text)
                sys.exit(1)

            folder_ids[folder] = resp.json().get("id")
            print(f"📁 Dossier créé : {folder}")
    return folder_ids

def get_emails():
    url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/mailFolders/Inbox/messages?$top=200&$orderby=receivedDateTime DESC"
    resp = requests.get(url, headers=headers)

    if resp.status_code != 200:
        print(f"❌ Erreur API get_emails: {resp.status_code}")
        print("Réponse brute:", resp.text)
        sys.exit(1)

    return resp.json().get("value", [])

def delete_email(mail_id):
    url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/messages/{mail_id}"
    return requests.delete(url, headers=headers).status_code == 204

def move_email(mail_id, folder_id):
    url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/messages/{mail_id}/move"
    resp = requests.post(url, headers=headers, json={"destinationId": folder_id})
    return resp.status_code in [200, 201]

# -------------------------------
# 5. TRI DES EMAILS
# -------------------------------
def trier_emails():
    folder_ids = get_folder_ids(keywords.keys())
    emails = get_emails()
    print(f"📨 {len(emails)} emails récupérés")

    seen_subjects = set()
    unique_emails = []

    for email in emails:
        subject = (email.get("subject") or "").strip().lower()
        mail_id = email["id"]

        if subject in seen_subjects:
            if delete_email(mail_id):
                print(f"🗑️ Doublon supprimé : '{subject}'")
            else:
                print(f"⚠️ Erreur suppression : '{subject}'")
        else:
            seen_subjects.add(subject)
            unique_emails.append(email)

    for email in unique_emails:
        subject = (email.get("subject") or "")
        mail_id = email["id"]
        target_folder = None

        for folder, regex_list in compiled_keywords.items():
            if any(regex.search(subject) for regex in regex_list):
                target_folder = folder
                break

        if target_folder:
            if move_email(mail_id, folder_ids[target_folder]):
                print(f"📌 '{subject}' déplacé vers {target_folder}")
            else:
                print(f"⚠️ Erreur déplacement : '{subject}'")
        else:
            print(f"✉️ '{subject}' laissé dans Inbox")

    print("✅ Tri terminé")

# -------------------------------
# 6. MAIN
# -------------------------------
if __name__ == "__main__":
    trier_emails()
