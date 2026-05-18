# -*- coding: utf-8 -*-
import pandas as pd
from datetime import datetime, timedelta
import sys
import locale
from pathlib import Path
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import importlib.util

# =====================================================
# PATHS & IMPORTS AVEC TIRETS
# =====================================================

ROOT = Path(__file__).parent.parent.parent  # dev/

def load_module(module_name, folder_name, file_name):
    """Charge un module depuis un dossier avec tirets"""
    module_path = ROOT / "modules" / folder_name / file_name
    spec = importlib.util.spec_from_file_location(module_name, module_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module

# Charge les modules
file_manager_module = load_module("file_manager", "file-manager", "file_manager.py")
dbf_loader_module = load_module("dbf_loader", "dbf-loader", "dbf_loader.py")
mailer_module = load_module("mailer", "quincaillerie-mailer", "mailer.py")

# Récupère les fonctions
generer_chemin_fichier = file_manager_module.generer_chemin_fichier
generer_nom_fichier = file_manager_module.generer_nom_fichier
get_dbf = dbf_loader_module.get_dbf
envoyer_email = mailer_module.envoyer_email

# =====================================================
# CONFIGURATION
# =====================================================
SOC = "QC"

MAIL_MAINTENANCE = ["support@quincaillerie.nc"]
MAIL_TARGET = [
    "support@quincaillerie.nc",
    "exploitation@quincaillerie.nc",
    "n.leroux@quincaillerie.nc",
    "indices@isee.nc"
]

NARTS = [
    "710092","760043","760041","761467","760200","760049",
    "760403","760414","210255","120139","590219","833848",
    "820097","870145","870127","760251","760105","761336",
    "240153","790449","160014","160890"
]

# =====================================================
# OUTILS
# =====================================================
def log(message):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")

def generate_excel(df, report_date):
    try:
        locale.setlocale(locale.LC_TIME, "fr_FR.UTF-8")
    except locale.Error:
        locale.setlocale(locale.LC_TIME, "French_France")

    date_file = report_date.strftime("%b-%y")
    file_name = f"{SOC}_comISEE_{date_file}.xlsx"
    file_path = generer_chemin_fichier(file_name)

    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="ISEE")
        worksheet = writer.sheets["ISEE"]

        header_fill  = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font  = Font(bold=True, color="FFFFFF", size=11)
        header_align = Alignment(horizontal="center")

        for col in range(1, len(df.columns) + 1):
            cell = worksheet.cell(row=1, column=col)
            cell.fill      = header_fill
            cell.font      = header_font
            cell.alignment = header_align

        for col in worksheet.columns:
            max_length = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            worksheet.column_dimensions[col_letter].width = max_length + 4

    log(f"📁 Fichier généré : {file_path}")
    return file_path, file_name

def send_mails(file_path, file_name, report_date):
    try:
        locale.setlocale(locale.LC_TIME, "fr_FR.UTF-8")
    except locale.Error:
        locale.setlocale(locale.LC_TIME, "French_France")

    date_mail = report_date.strftime("%B %Y").upper()
    sujet     = f"[{SOC}] - Relevé ISEE prix HT - {date_mail}"

    corps_html = f"""
    <html>
    <head>
        <style>
            body     {{ font-family: Arial, sans-serif; color: #333; }}
            .header  {{ background-color: #4472C4; color: white; padding: 10px; text-align: center; font-size: 18px; font-weight: bold; }}
            .content {{ margin: 20px; font-size: 14px; }}
            .footer  {{ margin: 20px; font-size: 12px; color: #777; }}
        </style>
    </head>
    <body>
        <div class="header">Relevé ISEE prix HT - {date_mail}</div>
        <div class="content">
            <p>Bonjour,</p>
            <p>Veuillez trouver en pièce jointe le relevé des prix HT pour <strong>{date_mail}</strong>.</p>
            <p>Cordialement,<br>Service Analyse Commercial</p>
        </div>
        <div class="footer">
            <p>Ce message est envoyé automatiquement, merci de ne pas y répondre.</p>
        </div>
    </body>
    </html>
    """

    tous_les_destinataires = list(set(MAIL_MAINTENANCE + MAIL_TARGET))

    for dest in tous_les_destinataires:
        log(f"📤 Envoi à {dest}...")
        envoyer_email(
            destinataire=dest,
            sujet=sujet,
            corps=corps_html,
            chemin_piece_jointe=str(file_path)
        )
        log(f"✅ Mail envoyé à {dest}")

# =====================================================
# MAIN
# =====================================================
def main():
    log("Début du script ComISEE")
    today       = datetime.now()
    report_date = today.replace(day=1) - timedelta(days=1)
    log(f"Rapport pour {report_date.strftime('%B %Y')}")

    try:
        df = get_dbf("qc/article.dbf")

        df_filtered = df[df["NART"].astype(str).isin(NARTS)]
        df_result   = df_filtered[["NART", "DESIGN", "PVTE"]].rename(
            columns={"NART": "Réf_Mag", "DESIGN": "Produit", "PVTE": "PVTE_HT"}
        )
        log(f"{len(df_result)} articles sélectionnés")

        file_path, file_name = generate_excel(df_result, report_date)
        send_mails(file_path, file_name, report_date)

        log("✅ Script terminé avec succès")

    except Exception as e:
        log(f"❌ ERREUR : {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
