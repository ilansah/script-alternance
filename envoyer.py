"""Envoi d'emails de candidature en alternance — throttlé, anti-doublons, via Outlook SMTP."""
import csv
import smtplib
import ssl
import sys
import time
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

LOG_FILE = Path("envoyes.txt")
CONTACTS_FILE = Path("contacts.csv")
TEMPLATES_DIR = Path("templates")


def load_env():
    env = {}
    path = Path(".env")
    if not path.exists():
        return env
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        k, _, v = line.partition("=")
        env[k.strip()] = v.strip().strip('"').strip("'")
    return env


ENV = load_env()
EMAIL = ENV.get("OUTLOOK_EMAIL", "")
PASSWORD = ENV.get("OUTLOOK_PASSWORD", "")
REPLY_TO = ENV.get("REPLY_TO", "")
CV_URL = ENV.get("CV_URL", "")
NOM_COMPLET = ENV.get("NOM_COMPLET", "Ilan Sahraoui")
TEL = ENV.get("TELEPHONE", "")
LINKEDIN = ENV.get("LINKEDIN", "")
GITHUB = ENV.get("GITHUB", "")
SCRIPT_URL = ENV.get("SCRIPT_URL", "")
PAUSE_SECONDES = int(ENV.get("PAUSE_SECONDES", "90"))
MAX_PAR_JOUR = int(ENV.get("MAX_PAR_JOUR", "25"))
SMTP_HOST = ENV.get("SMTP_HOST", "smtp.office365.com")
SMTP_PORT = int(ENV.get("SMTP_PORT", "587"))


def deja_envoyes() -> set[str]:
    if not LOG_FILE.exists():
        return set()
    emails = set()
    for line in LOG_FILE.read_text(encoding="utf-8").splitlines():
        parts = line.split(";")
        if len(parts) >= 2 and parts[3] == "ok":
            emails.add(parts[1].strip().lower())
    return emails


def charger_template(domaine: str) -> str:
    path = TEMPLATES_DIR / f"{domaine}.txt"
    if not path.exists():
        raise FileNotFoundError(f"Template manquant : {path} (domaines valides : ia, cyber, web, salarie)")
    return path.read_text(encoding="utf-8")


def personnaliser(template: str, contact: dict) -> str:
    prenom_brut = contact.get("prenom", "").strip()
    salutation = f"Bonjour {prenom_brut}," if prenom_brut else "Bonjour,"
    return (
        template
        .replace("{salutation}", salutation)
        .replace("{prenom}", prenom_brut or "Madame, Monsieur")
        .replace("{entreprise}", contact.get("entreprise", "").strip())
        .replace("{cv_url}", CV_URL)
        .replace("{nom_complet}", NOM_COMPLET)
        .replace("{tel}", TEL)
        .replace("{linkedin}", LINKEDIN)
        .replace("{github}", GITHUB)
        .replace("{script_url}", SCRIPT_URL)
    )


def extraire_sujet_corps(texte: str) -> tuple[str, str]:
    lignes = texte.splitlines()
    for i, ligne in enumerate(lignes):
        low = ligne.lower()
        if low.startswith("sujet:") or low.startswith("subject:"):
            sujet = ligne.split(":", 1)[1].strip()
            corps = "\n".join(lignes[i + 1:]).lstrip("\n")
            return sujet, corps
    raise ValueError("Le template doit commencer par une ligne 'Sujet: ...'")


def log(email: str, entreprise: str, statut: str, msg: str = ""):
    ts = datetime.now().isoformat(timespec="seconds")
    msg_clean = msg.replace("\n", " ").replace(";", ",")
    with LOG_FILE.open("a", encoding="utf-8") as f:
        f.write(f"{ts};{email};{entreprise};{statut};{msg_clean}\n")


def charger_contacts() -> list[dict]:
    if not CONTACTS_FILE.exists():
        print(f"[erreur] {CONTACTS_FILE} introuvable.")
        sys.exit(1)
    with CONTACTS_FILE.open(encoding="utf-8", newline="") as f:
        return list(csv.DictReader(f, delimiter=";"))


def valider_config():
    manquants = []
    if not EMAIL:
        manquants.append("OUTLOOK_EMAIL")
    if not PASSWORD:
        manquants.append("OUTLOOK_PASSWORD")
    if not CV_URL:
        manquants.append("CV_URL")
    if manquants:
        print(f"[erreur] Variables manquantes dans .env : {', '.join(manquants)}")
        print("         Copie .env.example en .env et remplis-le.")
        sys.exit(1)


def main():
    dry_run = "--dry-run" in sys.argv

    if not dry_run:
        valider_config()

    deja = deja_envoyes()
    contacts = charger_contacts()
    a_envoyer = [
        c for c in contacts
        if c.get("email", "").strip()
        and c["email"].strip().lower() not in deja
        and c.get("domaine", "").strip() in {"ia", "cyber", "web", "salarie"}
    ]
    print(f"[info] {len(contacts)} contacts chargés, {len(deja)} déjà contactés, "
          f"{len(a_envoyer)} candidats pour ce run.")
    a_envoyer = a_envoyer[:MAX_PAR_JOUR]
    print(f"[info] {len(a_envoyer)} seront traités (plafond MAX_PAR_JOUR={MAX_PAR_JOUR}).")

    if not a_envoyer:
        print("[info] rien à envoyer, fin.")
        return

    if dry_run:
        print("\n[dry-run] aperçu du premier mail :\n" + "-" * 60)
        c = a_envoyer[0]
        sujet, corps = extraire_sujet_corps(personnaliser(charger_template(c["domaine"]), c))
        print(f"À      : {c['email']}")
        print(f"Sujet  : {sujet}")
        print("-" * 60)
        print(corps)
        print("-" * 60)
        print(f"\n[dry-run] {len(a_envoyer)} mails seraient envoyés. Relance sans --dry-run pour envoyer.")
        return

    print(f"[info] connexion SMTP à {SMTP_HOST}:{SMTP_PORT}...")
    context = ssl.create_default_context()
    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as smtp:
            smtp.ehlo()
            smtp.starttls(context=context)
            smtp.ehlo()
            smtp.login(EMAIL, PASSWORD)
            print("[ok] authentifié.")

            for i, c in enumerate(a_envoyer, 1):
                dest = c["email"].strip()
                entreprise = c.get("entreprise", "").strip()
                try:
                    tmpl = charger_template(c["domaine"])
                    sujet, corps = extraire_sujet_corps(personnaliser(tmpl, c))
                    msg = EmailMessage()
                    msg["From"] = EMAIL
                    msg["To"] = dest
                    if REPLY_TO:
                        msg["Reply-To"] = REPLY_TO
                    msg["Subject"] = sujet
                    msg.set_content(corps)
                    smtp.send_message(msg)
                    log(dest, entreprise, "ok")
                    print(f"[{i}/{len(a_envoyer)}] ok   → {dest} ({entreprise})")
                except Exception as e:
                    log(dest, entreprise, "erreur", str(e))
                    print(f"[{i}/{len(a_envoyer)}] FAIL → {dest} ({entreprise}) : {e}")

                if i < len(a_envoyer):
                    time.sleep(PAUSE_SECONDES)
    except smtplib.SMTPAuthenticationError as e:
        print(f"\n[erreur] Auth SMTP refusée : {e}")
        print("  → Vérifie OUTLOOK_EMAIL / OUTLOOK_PASSWORD dans .env")
        print("  → L'école a peut-être désactivé 'Authenticated SMTP' sur Microsoft 365.")
        print("    Contacte l'IT ou bascule sur un Gmail perso (voir README).")
        sys.exit(2)

    print(f"\n[fini] log complet dans {LOG_FILE}")


if __name__ == "__main__":
    main()
