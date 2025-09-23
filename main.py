import os
import requests
from msal import ConfidentialClientApplication
import re

# ----------------------------
# Configuration via GitHub Secrets
# ----------------------------
CLIENT_ID = os.environ.get("APP_CLIENT_ID")
CLIENT_SECRET = os.environ.get("APP_CLIENT_SECRET")
TENANT_ID = os.environ.get("APP_TENANT_ID")
SCOPES = ["https://graph.microsoft.com/.default"]

# ----------------------------
# Authentification
# ----------------------------
app = ConfidentialClientApplication(
    client_id=CLIENT_ID,
    client_credential=CLIENT_SECRET,
    authority=f"https://login.microsoftonline.com/{TENANT_ID}"
)

result = app.acquire_token_for_client(scopes=SCOPES)

if "access_token" not in result:
    print("❌ Échec de la connexion :", result.get("error_description"))
    exit()

token = result["access_token"]
headers = {
    "Authorization": f"Bearer {token}",
    "Accept": "application/json",
    "Content-Type": "application/json"
}
print("✅ Authentification réussie")

# ----------------------------
# Regex des priorités
# ----------------------------
keywords = {
    "P1": [r"\bp1\b"],
    "P2": [r"\bp2\b", r"\bcertificate\b"],
    "P3": [r"\bp3\b"],
    "P4": [r"\bp4\b"]
}
compiled_keywords = {
    folder: [re.compile(pat, re.IGNORECASE) for pat in pats]
    for folder, pats in keywords.items()
}

# ----------------------------
# Fonctions principales
# ----------------------------
def get_folders():
    url = "https://graph.microsoft.com/v1.0/me/mailFolders?$top=100"
    resp = requests.get(url, headers=headers)
    return resp.json().get("value", [])

def get_folder_ids(targets):
    folder_ids = {}
    existing = get_folders()
    for name in targets:
        match = next((f for f in existing if f["displayName"].lower() == name.lower()), None)
        if match:
            folder_ids[name] = match["id"]
        else:
            resp = requests.post(
                "https://graph.microsoft.com/v1.0/me/mailFolders",
                headers=headers,
                json={"displayName": name}
            )
            folder_ids[name] = resp.json()["id"]
    return folder_ids

def get_emails():
    url = "https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages?$top=50"
    resp = requests.get(url, headers=headers)
    return resp.json().get("value", [])

def delete_email(mail_id):
    url = f"https://graph.microsoft.com/v1.0/me/messages/{mail_id}"
    return requests.delete(url, headers=headers).status_code == 204

def move_email(mail_id, folder_id):
    url = f"https://graph.microsoft.com/v1.0/me/messages/{mail_id}/move"
    data = {"destinationId": folder_id}
    return requests.post(url, headers=headers, json=data).status_code in (200, 201)

# ----------------------------
# Exécution principale
# ----------------------------
folder_ids = get_folder_ids(keywords.keys())
emails = get_emails()
print(f"📨 {len(emails)} email(s) récupérés")

seen_subjects = set()
emails_unique = []

for mail in emails:
    subject = (mail.get("subject") or "").strip().lower()
    mail_id = mail["id"]
    if subject in seen_subjects:
        if delete_email(mail_id):
            print(f"🗑️ Doublon supprimé : '{subject}'")
    else:
        seen_subjects.add(subject)
        emails_unique.append(mail)

for mail in emails_unique:
    subject = (mail.get("subject") or "")
    mail_id = mail["id"]
    destination = None
    for folder, patterns in compiled_keywords.items():
        if any(pat.search(subject) for pat in patterns):
            destination = folder
            break
    if destination:
        if move_email(mail_id, folder_ids[destination]):
            print(f"📌 '{subject}' déplacé vers {destination}")
        else:
            print(f"⚠️ Échec déplacement : {subject}")
    else:
        print(f"✉️ Aucun mot-clé trouvé : {subject}")
