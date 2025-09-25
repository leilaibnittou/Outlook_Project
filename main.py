import logging
import requests
import re
from msal import ConfidentialClientApplication
from flask import Flask, jsonify
import os

# -------------------------------
# CONFIGURATION
# -------------------------------
CLIENT_ID = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]
TENANT_ID = os.environ["TENANT_ID"]
SCOPES = ["https://graph.microsoft.com/.default"]
USER_ID = "compte_test_projet@outlook.com"  # 👉 Modifie selon le compte cible

# -------------------------------
# AUTHENTIFICATION
# -------------------------------
def get_access_token():
    app = ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}"
    )

    result = app.acquire_token_silent(SCOPES, account=None)
    if not result:
        logging.info("🔐 Récupération du token via client_credentials flow...")
        result = app.acquire_token_for_client(scopes=SCOPES)

    if "access_token" in result:
        logging.info("✅ Authentification réussie.")
        return result["access_token"]
    else:
        logging.error(f"❌ Authentification échouée : {result.get('error_description')}")
        raise Exception("Impossible d'obtenir un token d'accès")

# -------------------------------
# MOTS-CLÉS ET DOSSIERS
# -------------------------------
KEYWORDS = {
    "P1": [r"\bp1\b"],
    "P2": [r"\bp2\b", r"\bcertificate\b"],
    "P3": [r"\bp3\b"],
    "P4": [r"\bp4\b"]
}

COMPILED_KEYWORDS = {
    folder: [re.compile(pat, re.IGNORECASE) for pat in pats]
    for folder, pats in KEYWORDS.items()
}

# -------------------------------
# FONCTIONS UTILITAIRES
# -------------------------------
def get_folders(headers):
    url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/mailFolders?$top=100"
    folders = []
    while url:
        resp = requests.get(url, headers=headers)
        resp.raise_for_status()
        data = resp.json()
        folders.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return folders

def get_folder_ids(headers, targets):
    existing = get_folders(headers)
    folder_ids = {}
    for f in targets:
        folder = next((x for x in existing if x["displayName"].lower() == f.lower()), None)
        if folder:
            folder_ids[f] = folder["id"]
        else:
            resp = requests.post(
                f"https://graph.microsoft.com/v1.0/users/{USER_ID}/mailFolders",
                headers=headers,
                json={"displayName": f}
            )
            resp.raise_for_status()
            folder_ids[f] = resp.json()["id"]
    return folder_ids

def get_emails(headers):
    url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/mailFolders/Inbox/messages?$top=200&$orderby=receivedDateTime DESC"
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    return resp.json().get("value", [])

def delete_email(headers, mail_id):
    url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/messages/{mail_id}"
    resp = requests.delete(url, headers=headers)
    return resp.status_code == 204

def move_email(headers, mail_id, folder_id):
    url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/messages/{mail_id}/move"
    resp = requests.post(url, headers=headers, json={"destinationId": folder_id})
    return resp.status_code in (200, 201)

# -------------------------------
# TRAITEMENT PRINCIPAL
# -------------------------------
def handler():
    token = get_access_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/json"
    }

    folder_ids = get_folder_ids(headers, KEYWORDS.keys())
    emails = get_emails(headers)
    logging.info(f"📨 {len(emails)} emails récupérés")

    seen_subjects = set()
    emails_unique = []

    for mail in emails:
        subject = (mail.get("subject") or "").strip().lower()
        mail_id = mail["id"]
        if subject in seen_subjects:
            if delete_email(headers, mail_id):
                logging.info(f"🗑️ Doublon supprimé : '{subject}'")
            else:
                logging.warning(f"⚠️ Erreur suppression doublon : '{subject}'")
        else:
            seen_subjects.add(subject)
            emails_unique.append(mail)

    for mail in emails_unique:
        subject = mail.get("subject") or ""
        mail_id = mail["id"]
        target_folder = None
        for folder, regex_list in COMPILED_KEYWORDS.items():
            if any(regex.search(subject) for regex in regex_list):
                target_folder = folder
                break

        if target_folder:
            if move_email(headers, mail_id, folder_ids[target_folder]):
                logging.info(f"📌 '{subject}' déplacé vers {target_folder}")
            else:
                logging.warning(f"⚠️ Erreur déplacement '{subject}'")
        else:
            logging.info(f"✉️ '{subject}' laissé dans Inbox")

    logging.info("✅ Tri terminé.")
    return {"status": "success", "message": "Tri terminé"}

# -------------------------------
# SERVEUR WEB FLASK
# -------------------------------
app = Flask(__name__)

@app.route("/")
def home():
    return "API Email Sorter is running."

@app.route("/run-script")
def run_script():
    try:
        result = handler()
        return jsonify(result)
    except Exception as e:
        logging.error(f"Erreur lors de l'exécution: {str(e)}")
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == "__main__":
    # Replit attends souvent que Flask écoute sur 0.0.0.0 et le port défini par l'environnement
    port = int(os.environ.get("PORT", 3000))
    app.run(host="0.0.0.0", port=port)
