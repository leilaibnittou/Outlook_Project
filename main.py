from msal import ConfidentialClientApplication
import requests
import re
import os

# -------------------------------
# CONFIGURATION (via secrets GitHub)
# -------------------------------
APP_ENV = os.getenv("APP_ENV", "TEST")

if APP_ENV == "TEST":
    CLIENT_ID = os.getenv("TEST_CLIENT_ID")
    CLIENT_SECRET = os.getenv("TEST_CLIENT_SECRET")
    TENANT_ID = os.getenv("TEST_TENANT_ID")
else:
    CLIENT_ID = os.getenv("PROD_CLIENT_ID")
    CLIENT_SECRET = os.getenv("PROD_CLIENT_SECRET")
    TENANT_ID = os.getenv("PROD_TENANT_ID")

SCOPES = ["https://graph.microsoft.com/.default"]

# -------------------------------
# AUTHENTIFICATION
# -------------------------------
app = ConfidentialClientApplication(
    client_id=CLIENT_ID,
    client_credential=CLIENT_SECRET,
    authority=f"https://login.microsoftonline.com/{TENANT_ID}"
)

result = app.acquire_token_silent(SCOPES, account=None)

if not result:
    result = app.acquire_token_for_client(scopes=SCOPES)

if "access_token" not in result:
    print("❌ Authentification échouée :", result.get("error_description"))
    exit()

token = result["access_token"]
headers = {
    "Authorization": f"Bearer {token}",
    "Accept": "application/json",
    "Content-Type": "application/json"
}
print("✅ Authentification réussie")

# -------------------------------
# MOTS-CLÉS ET DOSSIERS
# -------------------------------
keywords = {
    "P1": [r"\bp1\b"],
    "P2": [r"\bp2\b", r"\bcertificate\b"],
    "P3": [r"\bp3\b"],
    "P4": [r"\bp4\b"]
}

compiled_keywords = {
    folder: [re.compile(pattern, re.IGNORECASE) for pattern in patterns]
    for folder, patterns in keywords.items()
}

# -------------------------------
# FONCTIONS API OUTLOOK
# -------------------------------
def get_folders():
    url = "https://graph.microsoft.com/v1.0/me/mailFolders?$top=100"
    folders = []
    while url:
        resp = requests.get(url, headers=headers).json()
        folders.extend(resp.get("value", []))
        url = resp.get("@odata.nextLink")
    return folders

def get_folder_ids(targets):
    folder_ids = {}
    existing = get_folders()
    for name in targets:
        match = next((f for f in existing if f["displayName"].lower() == name.lower()), None)
        if match:
            folder_ids[name] = match["id"]
        else:
            # Créer le dossier
            resp = requests.post(
                "https://graph.microsoft.com/v1.0/me/mailFolders",
                headers=headers,
                json={"displayName": name}
            )
            resp_json = resp.json()

            if resp.status_code >= 400:
                print(f"❌ Erreur lors de la création du dossier '{name}': {resp.status_code}")
                print(f"🔎 Réponse : {resp_json}")
                continue

            folder_ids[name] = resp_json.get("id")
    return folder_ids

def get_emails():
    url = "https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages?$top=100&$orderby=receivedDateTime DESC"
    resp = requests.get(url, headers=headers)
    return resp.json().get("value", [])

def delete_email(mail_id):
    url = f"https://graph.microsoft.com/v1.0/me/messages/{mail_id}"
    resp = requests.delete(url, headers=headers)
    return resp.status_code == 204

def move_email(mail_id, folder_id):
    url = f"https://graph.microsoft.com/v1.0/me/messages/{mail_id}/move"
    resp = requests.post(url, headers=headers, json={"destinationId": folder_id})
    return resp.status_code in (200, 201)

# -------------------------------
# EXÉCUTION DU TRI
# -------------------------------
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

    for folder, regexes in compiled_keywords.items():
        if any(regex.search(subject) for regex in regexes):
            target_folder = folder
            break

    if target_folder:
        if move_email(mail_id, folder_ids.get(target_folder)):
            print(f"📌 '{subject}' déplacé vers {target_folder}")
        else:
            print(f"⚠️ Erreur déplacement : '{subject}'")
    else:
        print(f"✉️ '{subject}' laissé dans Inbox")

print("✅ Traitement terminé.")
