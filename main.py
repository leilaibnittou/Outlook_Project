from msal import PublicClientApplication
import requests
import re

# -------------------------------
# CONFIGURATION
# -------------------------------
CLIENT_ID = "0f9136a2-194d-458f-8044-55854a40f5f0"
TENANT_ID = "common"
SCOPES = ["Mail.ReadWrite"]

# -------------------------------
# AUTHENTIFICATION
# -------------------------------
app = PublicClientApplication(
    client_id=CLIENT_ID,
    authority=f"https://login.microsoftonline.com/{TENANT_ID}"
)

accounts = app.get_accounts()
result = app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None
if not result:
    print("Connexion nécessaire, ouverture du navigateur...")
    result = app.acquire_token_interactive(scopes=SCOPES)
if "access_token" not in result:
    print("Échec de la connexion :", result.get("error_description"))
    exit()

token = result["access_token"]
headers = {
    "Authorization": f"Bearer {token}",
    "Accept": "application/json",
    "Content-Type": "application/json"
}
print("✅ Connexion réussie !")

# -------------------------------
# DOSSIERS ET MOTS-CLÉS (regex exacts)
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
# FONCTIONS UTILES
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
    for f in targets:
        folder = next((x for x in existing if x["displayName"].lower() == f.lower()), None)
        if folder:
            folder_ids[f] = folder["id"]
        else:
            resp = requests.post(
                "https://graph.microsoft.com/v1.0/me/mailFolders",
                headers=headers,
                json={"displayName": f}
            )
            folder_ids[f] = resp.json()["id"]
    return folder_ids

def get_emails():
    url = "https://graph.microsoft.com/v1.0/me/mailfolders/Inbox/messages?$top=200&$orderby=receivedDateTime DESC"
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
# CRÉER LES DOSSIERS
# -------------------------------
folder_ids = get_folder_ids(keywords.keys())

# -------------------------------
# TRI DES EMAILS AVEC SUPPRESSION DES DOUBLONS
# -------------------------------
emails = get_emails()
print(f"📨 {len(emails)} emails récupérés")

# Vérifier les doublons par subject
seen_subjects = set()
emails_unique = []

for mail in emails:
    subject = (mail.get("subject") or "").strip().lower()
    mail_id = mail["id"]

    if subject in seen_subjects:
        # doublon → supprimer
        if delete_email(mail_id):
            print(f"🗑️ Doublon supprimé : '{subject}'")
        else:
            print(f"⚠️ Erreur suppression doublon : '{subject}'")
    else:
        seen_subjects.add(subject)
        emails_unique.append(mail)

# utiliser emails_unique pour le tri
for mail in emails_unique:
    subject = (mail.get("subject") or "")
    mail_id = mail["id"]
    target_folder = None

    for folder, regex_list in compiled_keywords.items():
        if any(regex.search(subject) for regex in regex_list):
            target_folder = folder
            break

    if target_folder:
        if move_email(mail_id, folder_ids[target_folder]):
            print(f"📌 '{subject}' déplacé vers {target_folder}")
        else:
            print(f"⚠️ Erreur déplacement '{subject}'")
    else:
        print(f"✉️ '{subject}' laissé dans Inbox")

print("✅ Tri et suppression des doublons terminé. Vérifie dans Outlook.")
