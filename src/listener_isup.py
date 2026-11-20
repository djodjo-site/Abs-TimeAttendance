import sys
import os
import json
import re
import io
import pandas as pd
from zoneinfo import ZoneInfo

# Ajoute dynamique du dossier parent dans le chemin Python
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from fastapi import FastAPI, Request, Response, Form, Cookie, UploadFile, File
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse, StreamingResponse
from datetime import datetime, date, time, timedelta
from database.db_config import get_db_connection

app = FastAPI(title="Hikvision Attendance System")

TRANSLATIONS = {
    "fr": {
        "login_title": "Connexion DRH - Hikvision",
        "login_heading": "Espace DRH",
        "login_subtitle": "Veuillez vous connecter pour accéder au suivi",
        "username": "Nom d'utilisateur",
        "password": "Mot de passe",
        "login_btn": "Se connecter",
        "invalid_creds": "Identifiants+incorrects",
        "forbidden_role": "Accès+réservé+au+compte+DRH",
        "sup_page_title": "Exportation des Rapports - Superviseur",
        "sup_header": "Espace Superviseur - Exportation",
        "sup_heading": "Exportation des Rapports de Présence",
        "sup_subtitle": "Sélectionnez une période pour exporter le rapport journalier (.xlsx)",
        "sup_export_btn": "Exporter le Rapport (.xlsx)",
        "sup_notice": "Ce compte est limité exclusivement à l'exportation du rapport.",
        "dev_by": "Application développée par",
        "nav_brand": "Suivi de Présence & Pointage Hikvision",
        "nav_dashboard": "Tableau de bord",
        "nav_reports": "Rapport Journalier",
        "nav_summary": "Synthèse Employés",
        "nav_logout": "Déconnexion",
        "kpi_employees": "Employés Enregistrés",
        "kpi_status": "Statut du Système",
        "kpi_online": "En ligne 🟢",
        "kpi_devices": "Terminaux Connectés",
        "detected": "détecté(s)",
        "import_title": "Importation de la liste des Employés (Fichier Excel / CSV)",
        "import_label": "Sélectionner un fichier (.xlsx, .xls, .csv)",
        "import_help": "Colonnes attendues dans votre fichier :",
        "import_btn": "Importer les Employés",
        "sched_add_title": "Configurer un Horaire de Travail",
        "sched_name": "Nom de l'horaire",
        "sched_name_ph": "ex: Horaire Normal, Équipe Nuit",
        "sched_start": "Heure d'arrivée",
        "sched_end": "Heure de départ",
        "sched_tol": "Tolérance de retard (minutes)",
        "sched_add_btn": "Ajouter l'horaire",
        "sched_list_title": "Horaires configurés",
        "sched_col_name": "Nom",
        "sched_col_in": "Entrée",
        "sched_col_out": "Sortie",
        "sched_col_tol": "Tolérance",
        "sched_del_confirm": "Voulez-vous supprimer cet horaire ?",
        "dev_title": "Gestion & Ajout Manuel des Terminaux de Pointage",
        "dev_key_ph": "ID Technique (ex: Terminal_Porte1 ou Serial No)",
        "dev_name_ph": "Nom personnalisé (ex: Porte Principale)",
        "add_btn": "Ajouter",
        "save_btn": "Enregistrer",
        "dev_col_key": "ID Technique / IP d'origine",
        "dev_col_name": "Nom Personnalisé (Ex: Porte Principale, Entrepôt...)",
        "dev_none": "Aucun terminal détecté pour l'instant. Effectuez un pointage.",
        "dev_del_confirm": "Voulez-vous vraiment supprimer ce terminal ?",
        "emp_add_title": "Enregistrer un Nouvel Employé",
        "emp_badge": "Matricule / ID Badge",
        "emp_firstname": "Prénom",
        "emp_lastname": "Nom",
        "emp_dept": "Département",
        "emp_sched": "Horaire de travail assigné",
        "emp_save_btn": "Enregistrer l'employé",
        "emp_list_title": "Liste des Employés",
        "emp_summary_link": "Voir la synthèse globale",
        "emp_col_badge": "Matricule / Badge",
        "emp_col_fullname": "Nom & Prénom",
        "emp_col_dept": "Département",
        "emp_col_sched": "Horaire assigné",
        "emp_col_status": "Statut",
        "emp_col_actions": "Actions",
        "emp_active": "Actif",
        "emp_inactive": "Inactif",
        "emp_none": "Aucun employé trouvé.",
        "emp_del_confirm": "Voulez-vous vraiment supprimer cet employé ?",
        "emp_edit_title": "Modifier un Employé - Hikvision",
        "emp_edit_header": "Modification d'un Employé",
        "emp_edit_sub": "Modifier les informations de",
        "emp_active_chk": "Employé actif",
        "cancel_btn": "Annuler",
        "rep_page_title": "Rapport Journalier - Hikvision",
        "rep_header": "Rapport Journalier Détaillé",
        "rep_filter_title": "Filtrer le rapport journalier",
        "start_date": "Date de début",
        "end_date": "Date de fin",
        "terminal": "Terminal",
        "all_terminals": "Tous les terminaux",
        "filter_btn": "Filtrer",
        "rep_detail_title": "Détail des pointages",
        "records": "enregistrements",
        "rep_none": "Aucun pointage trouvé pour cette période.",
        "date": "Date",
        "matricule": "Matricule",
        "employee": "Employé",
        "arrival": "Arrivée",
        "departure": "Départ",
        "hours": "Heures",
        "device": "Terminal",
        "status": "Statut de Conformité",
        "action": "Action",
        "missing": "Manquant",
        "late_missing": "Retard/Manquant",
        "absent_inc": "Absent (Incomplet)",
        "late": "En retard",
        "early_leave": "Sortie anticipée",
        "present_ok": "Présent (Conforme)",
        "non_compliant": "Non conforme",
        "punch_del_confirm": "Voulez-vous vraiment supprimer ce pointage ?",
        "export_excel": "Exporter Excel (.xlsx)",
        "print_pdf": "Imprimer / PDF",
        "sum_page_title": "Synthèse des Présences - Hikvision",
        "sum_header": "Synthèse Globale des Présences",
        "sum_filter_title": "Période de résumé du filtre",
        "sum_report_title": "Rapport de synthèse par employé",
        "collaborators": "employés",
        "sum_none": "Aucune donnée trouvée pour cette période.",
        "sum_present_days": "De nos jours",
        "sum_absent_days": "Jours Absents",
        "sum_late_days": "Jours tardifs",
        "sum_worked_hours": "Heures Travaillées",
        "export_summary_excel": "Résumé d'exportation (.xlsx)",
        "unknown": "Inconnu",
        "excel_sheet_rep": "Rapport Pointage",
        "excel_sheet_sum": "Synthèse Présences",
        "excel_col_date": "Date",
        "excel_col_badge": "Matricule",
        "excel_col_emp": "Employé",
        "excel_col_dept": "Département",
        "excel_col_in": "Arrivée",
        "excel_col_out": "Départ",
        "excel_col_hours": "Heures Travaillées",
        "excel_col_total_hours": "Total Heures Travaillées",
        "excel_col_dev": "Terminal",
        "excel_col_status": "Statut",
        "excel_col_present": "Jours Présents",
        "excel_col_absent": "Jours Absents",
        "excel_col_late": "Jours en Retard",
    },
    "en": {
        "login_title": "HR Login - Hikvision",
        "login_heading": "HR Portal",
        "login_subtitle": "Please sign in to access attendance tracking",
        "username": "Username",
        "password": "Password",
        "login_btn": "Sign in",
        "invalid_creds": "Invalid+credentials",
        "forbidden_role": "Access+restricted+to+HR+account",
        "sup_page_title": "Report Export - Supervisor",
        "sup_header": "Supervisor Portal - Export",
        "sup_heading": "Attendance Report Export",
        "sup_subtitle": "Select a date range to export the daily report (.xlsx)",
        "sup_export_btn": "Export Report (.xlsx)",
        "sup_notice": "This account is restricted strictly to exporting the report.",
        "dev_by": "Application developed by",
        "nav_brand": "Hikvision Attendance & Time Tracking",
        "nav_dashboard": "Dashboard",
        "nav_reports": "Daily Report",
        "nav_summary": "Employee Summary",
        "nav_logout": "Logout",
        "kpi_employees": "Registered Employees",
        "kpi_status": "System Status",
        "kpi_online": "Online 🟢",
        "kpi_devices": "Connected Devices",
        "detected": "detected",
        "import_title": "Import Employee List (Excel / CSV File)",
        "import_label": "Select a file (.xlsx, .xls, .csv)",
        "import_help": "Expected columns in your file:",
        "import_btn": "Import Employees",
        "sched_add_title": "Configure a Work Schedule",
        "sched_name": "Schedule name",
        "sched_name_ph": "e.g. Standard Shift, Night Shift",
        "sched_start": "Check-in time",
        "sched_end": "Check-out time",
        "sched_tol": "Late tolerance (minutes)",
        "sched_add_btn": "Add schedule",
        "sched_list_title": "Configured schedules",
        "sched_col_name": "Name",
        "sched_col_in": "Check-in",
        "sched_col_out": "Check-out",
        "sched_col_tol": "Tolerance",
        "sched_del_confirm": "Do you want to delete this schedule?",
        "dev_title": "Manage & Manually Add Attendance Devices",
        "dev_key_ph": "Technical ID (e.g. Terminal_Door1 or Serial No)",
        "dev_name_ph": "Friendly name (e.g. Main Door)",
        "add_btn": "Add",
        "save_btn": "Save",
        "dev_col_key": "Technical ID / Source IP",
        "dev_col_name": "Friendly Name (e.g. Main Door, Warehouse...)",
        "dev_none": "No device detected yet. Perform a punch.",
        "dev_del_confirm": "Are you sure you want to delete this device?",
        "emp_add_title": "Register a New Employee",
        "emp_badge": "Employee ID / Badge ID",
        "emp_firstname": "First name",
        "emp_lastname": "Last name",
        "emp_dept": "Department",
        "emp_sched": "Assigned work schedule",
        "emp_save_btn": "Save employee",
        "emp_list_title": "Employee List",
        "emp_summary_link": "View global summary",
        "emp_col_badge": "Employee ID / Badge",
        "emp_col_fullname": "Full Name",
        "emp_col_dept": "Department",
        "emp_col_sched": "Assigned schedule",
        "emp_col_status": "Status",
        "emp_col_actions": "Actions",
        "emp_active": "Active",
        "emp_inactive": "Inactive",
        "emp_none": "No employees found.",
        "emp_del_confirm": "Are you sure you want to delete this employee?",
        "emp_edit_title": "Edit Employee - Hikvision",
        "emp_edit_header": "Edit Employee",
        "emp_edit_sub": "Update information for",
        "emp_active_chk": "Active employee",
        "cancel_btn": "Cancel",
        "rep_page_title": "Daily Report - Hikvision",
        "rep_header": "Detailed Daily Report",
        "rep_filter_title": "Filter daily report",
        "start_date": "Start date",
        "end_date": "End date",
        "terminal": "Device",
        "all_terminals": "All devices",
        "filter_btn": "Filter",
        "rep_detail_title": "Attendance details",
        "records": "records",
        "rep_none": "No attendance records found for this period.",
        "date": "Date",
        "matricule": "Employee ID",
        "employee": "Employee",
        "arrival": "Check-in",
        "departure": "Check-out",
        "hours": "Hours",
        "device": "Device",
        "status": "Compliance Status",
        "action": "Action",
        "missing": "Missing",
        "late_missing": "Late/Missing",
        "absent_inc": "Absent (Incomplete)",
        "late": "Late",
        "early_leave": "Early departure",
        "present_ok": "Present (Compliant)",
        "non_compliant": "Non-compliant",
        "punch_del_confirm": "Are you sure you want to delete this attendance log?",
        "export_excel": "Export Excel (.xlsx)",
        "print_pdf": "Print / PDF",
        "sum_page_title": "Attendance Summary - Hikvision",
        "sum_header": "Global Attendance Summary",
        "sum_filter_title": "Filter summary period",
        "sum_report_title": "Summary Report by Employee",
        "collaborators": "employees",
        "sum_none": "No data found for this period.",
        "sum_present_days": "Present Days",
        "sum_absent_days": "Absent Days",
        "sum_late_days": "Late Days",
        "sum_worked_hours": "Worked Hours",
        "export_summary_excel": "Export Summary (.xlsx)",
        "unknown": "Unknown",
        "excel_sheet_rep": "Attendance Report",
        "excel_sheet_sum": "Attendance Summary",
        "excel_col_date": "Date",
        "excel_col_badge": "Employee ID",
        "excel_col_emp": "Employee",
        "excel_col_dept": "Department",
        "excel_col_in": "Check-in",
        "excel_col_out": "Check-out",
        "excel_col_hours": "Worked Hours",
        "excel_col_total_hours": "Total Worked Hours",
        "excel_col_dev": "Device",
        "excel_col_status": "Status",
        "excel_col_present": "Present Days",
        "excel_col_absent": "Absent Days",
        "excel_col_late": "Late Days",
    },
}


def get_lang(request: Request, lang_cookie: str = None) -> str:
    q_lang = (request.query_params.get("lang") or "").lower()
    if q_lang in ("fr", "en"):
        return q_lang
    c_lang = (lang_cookie or "").lower()
    if c_lang in ("fr", "en"):
        return c_lang
    return "fr"


def tr(key: str, lang: str = "fr") -> str:
    return TRANSLATIONS.get(lang, TRANSLATIONS["fr"]).get(key, TRANSLATIONS["fr"].get(key, key))


def lang_switcher_html(lang: str = "fr") -> str:
    fr_cls = "btn-light fw-bold" if lang == "fr" else "btn-outline-light"
    en_cls = "btn-light fw-bold" if lang == "en" else "btn-outline-light"
    return f"""
    <div class="btn-group btn-group-sm" role="group" aria-label="Language">
        <a href="/set-lang/fr" class="btn {fr_cls}">FR</a>
        <a href="/set-lang/en" class="btn {en_cls}">EN</a>
    </div>
    """


def login_lang_switcher_html(lang: str = "fr") -> str:
    fr_cls = "btn-primary fw-bold" if lang == "fr" else "btn-outline-secondary"
    en_cls = "btn-primary fw-bold" if lang == "en" else "btn-outline-secondary"
    return f"""
    <div class="d-flex justify-content-end mb-2">
        <div class="btn-group btn-group-sm" role="group" aria-label="Language">
            <a href="/set-lang/fr" class="btn {fr_cls}">FR</a>
            <a href="/set-lang/en" class="btn {en_cls}">EN</a>
        </div>
    </div>
    """


@app.get("/set-lang/{lang_code}")
async def set_language(lang_code: str, request: Request):
    target_lang = lang_code.lower() if lang_code.lower() in ("fr", "en") else "fr"
    referer = request.headers.get("referer") or "/"
    response = RedirectResponse(url=referer, status_code=303)
    response.set_cookie(key="lang", value=target_lang, httponly=False, max_age=60 * 60 * 24 * 365)
    return response


# Routes directe pour servir le logo ABS
@app.get("/static/logo.png")
async def get_logo():
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    logo_path = os.path.join(base_dir, "static", "logo.png")
    if os.path.exists(logo_path):
        return FileResponse(logo_path)
    return Response(status_code=404)


DRH_USERNAME = "drh"
DRH_PASSWORD = "password123"
SUPERVISOR_USERNAME = "superviseur"
SUPERVISOR_PASSWORD = "password123"
LECTEUR_USERNAME = "lecteur"
LECTEUR_PASSWORD = "password123"


def get_session_role(drh_session: str = None) -> str:
    if drh_session == "drh":
        return "drh"
    if drh_session == "supervisor":
        return "supervisor"
    if drh_session == "viewer":
        return "viewer"
    return None


def _table_exists(cursor, table_name: str) -> bool:
    cursor.execute("""
        SELECT EXISTS (
            SELECT 1
            FROM information_schema.tables
            WHERE table_schema = 'public' AND table_name = %s
        ) AS exists_flag
    """, (table_name,))
    res = cursor.fetchone()
    return bool(res and res.get("exists_flag"))


# Fonction utilitaire pour initialiser les tables et horaires par défaut
def init_database_tables():
    conn = None
    try:
        conn = get_db_connection()
        conn.autocommit = False
        with conn.cursor() as cursor:
            # Table des employés (si schema.sql ne l'a pas encore créée)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS employees (
                    badge_or_card_id VARCHAR(100) PRIMARY KEY,
                    first_name VARCHAR(100),
                    last_name VARCHAR(100),
                    department VARCHAR(100),
                    is_active BOOLEAN DEFAULT TRUE
                )
            """)

            # Table des logs bruts (si schema.sql ne l'a pas encore créée)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS attendance_raw_logs (
                    id SERIAL PRIMARY KEY,
                    badge_or_card_id VARCHAR(100),
                    punch_timestamp TIMESTAMP,
                    device_id VARCHAR(100),
                    punch_type VARCHAR(20) DEFAULT 'IN'
                )
            """)

            # Table de renommage des appareils
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS devices_mapping (
                    device_key VARCHAR(100) PRIMARY KEY,
                    friendly_name VARCHAR(100) NOT NULL
                )
            """)

            # Table des horaires de travail
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS work_schedules (
                    id SERIAL PRIMARY KEY,
                    name VARCHAR(100) NOT NULL,
                    start_time TIME NOT NULL,
                    end_time TIME NOT NULL,
                    late_tolerance_minutes INT DEFAULT 30
                )
            """)

            # Insérer un horaire par défaut s'il n'en existe aucun
            cursor.execute("SELECT COUNT(*) as cnt FROM work_schedules")
            res = cursor.fetchone()
            if res and res['cnt'] == 0:
                cursor.execute("""
                    INSERT INTO work_schedules (name, start_time, end_time, late_tolerance_minutes)
                    VALUES ('Standard (08:00 - 17:00)', '08:00:00', '17:00:00', 30)
                """)

            # Vérifier si la colonne schedule_id existe dans employees, sinon l'ajouter
            if _table_exists(cursor, "employees"):
                cursor.execute("""
                    SELECT column_name
                    FROM information_schema.columns
                    WHERE table_schema = 'public' AND table_name = 'employees' AND column_name = 'schedule_id'
                """)
                if not cursor.fetchone():
                    cursor.execute("""
                        ALTER TABLE employees ADD COLUMN schedule_id INT REFERENCES work_schedules(id)
                    """)

        conn.commit()
    except Exception as e:
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        print(f"Erreur initialisation tables BDD : {e}")
    finally:
        if conn:
            try:
                conn.close()
            except Exception:
                pass


@app.on_event("startup")
async def startup_event():
    init_database_tables()


@app.get("/login", response_class=HTMLResponse)
async def login_page(request: Request, error: str = None, lang: str = Cookie(None)):
    current_lang = get_lang(request, lang)
    error_html = f'<div class="alert alert-danger py-2 mb-3 text-center small">{error}</div>' if error else ''
    html_content = f"""
    <!DOCTYPE html>
    <html lang="{current_lang}">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{tr("login_title", current_lang)}</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
        <style>
            body {{
                background: linear-gradient(135deg, #1e293b, #0f172a);
                height: 100vh;
                display: flex;
                align-items: center;
                justify-content: center;
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            }}
            .login-card {{
                width: 100%;
                max-width: 420px;
                border: none;
                border-radius: 16px;
                box-shadow: 0 10px 25px rgba(0,0,0,0.3);
                background: #ffffff;
            }}
            .logo-container img {{
                max-height: 80px;
                width: auto;
                object-fit: contain;
            }}
            .developer-footer {{
                border-top: 1px solid #f1f5f9;
                padding-top: 15px;
                margin-top: 20px;
                font-size: 0.8rem;
                color: #64748b;
            }}
        </style>
    </head>
    <body>
        <div class="card login-card p-4">
            {login_lang_switcher_html(current_lang)}
            <div class="text-center mb-4 logo-container">
                <img src="/static/logo.png" alt="Logo ABS" class="mb-3">
                <h4 class="fw-bold text-dark">{tr("login_heading", current_lang)}</h4>
                <p class="text-muted small">{tr("login_subtitle", current_lang)}</p>
            </div>
            {error_html}
            <form action="/login" method="POST">
                <div class="mb-3">
                    <label class="form-label text-secondary small fw-bold">{tr("username", current_lang)}</label>
                    <div class="input-group">
                        <span class="input-group-text bg-light"><i class="fa-solid fa-user text-muted"></i></span>
                        <input type="text" class="form-control" name="username" placeholder="ex: drh, superviseur ou lecteur" required>
                    </div>
                </div>
                <div class="mb-4">
                    <label class="form-label text-secondary small fw-bold">{tr("password", current_lang)}</label>
                    <div class="input-group">
                        <span class="input-group-text bg-light"><i class="fa-solid fa-lock text-muted"></i></span>
                        <input type="password" class="form-control" name="password" placeholder="••••••••" required>
                    </div>
                </div>
                <button type="submit" class="btn btn-primary w-100 py-2 fw-semibold shadow-sm"><i class="fa-solid fa-right-to-bracket me-2"></i> {tr("login_btn", current_lang)}</button>
            </form>

            <div class="text-center developer-footer">
                <i class="fa-solid fa-code me-1 text-primary"></i> {tr("dev_by", current_lang)} <strong>Ir Jonathan LUKOKI</strong>
            </div>
        </div>
    </body>
    </html>
    """
    return HTMLResponse(content=html_content)


@app.post("/login")
async def login_action(request: Request, username: str = Form(...), password: str = Form(...), lang: str = Cookie(None)):
    current_lang = get_lang(request, lang)
    if username == DRH_USERNAME and password == DRH_PASSWORD:
        response = RedirectResponse(url="/", status_code=303)
        response.set_cookie(key="drh_session", value="drh", httponly=True)
        return response
    elif username == SUPERVISOR_USERNAME and password == SUPERVISOR_PASSWORD:
        response = RedirectResponse(url="/supervisor", status_code=303)
        response.set_cookie(key="drh_session", value="supervisor", httponly=True)
        return response
    elif username == LECTEUR_USERNAME and password == LECTEUR_PASSWORD:
        response = RedirectResponse(url="/viewer-export", status_code=303)
        response.set_cookie(key="drh_session", value="viewer", httponly=True)
        return response
    else:
        msg = tr("invalid_creds", current_lang)
        return RedirectResponse(url=f"/login?error={msg}", status_code=303)


@app.get("/logout")
async def logout():
    response = RedirectResponse(url="/login", status_code=303)
    response.delete_cookie(key="drh_session")
    return response


@app.get("/supervisor", response_class=HTMLResponse)
async def supervisor_export_page(request: Request, start_date: str = None, end_date: str = None, drh_session: str = Cookie(None), lang: str = Cookie(None)):
    role = get_session_role(drh_session)
    if not role:
        return RedirectResponse(url="/login", status_code=303)
    if role != "drh" and role != "supervisor":
        return RedirectResponse(url="/login", status_code=303)

    current_lang = get_lang(request, lang)
    today_str = date.today().isoformat()
    start_date = start_date or today_str
    end_date = end_date or today_str

    html_content = f"""
    <!DOCTYPE html>
    <html lang="{current_lang}">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{tr("sup_page_title", current_lang)}</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
        <style>
            body {{
                background-color: #f4f7f6;
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            }}
            .navbar {{
                background: linear-gradient(135deg, #1e293b, #0f172a) !important;
            }}
            .card {{
                border: none;
                border-radius: 12px;
                box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
            }}
        </style>
    </head>
    <body>
        <nav class="navbar navbar-dark shadow-sm py-3">
            <div class="container-fluid px-4">
                <span class="navbar-brand mb-0 h5 fw-bold"><i class="fa-solid fa-file-export me-2"></i> {tr("sup_header", current_lang)}</span>
                <div class="d-flex align-items-center gap-2">
                    {lang_switcher_html(current_lang)}
                    <a href="/logout" class="btn btn-danger btn-sm px-3 ms-1"><i class="fa-solid fa-right-from-bracket me-1"></i> {tr("nav_logout", current_lang)}</a>
                </div>
            </div>
        </nav>

        <div class="container my-5" style="max-width: 640px;">
            <div class="card shadow-sm">
                <div class="card-header bg-primary text-white py-3">
                    <h5 class="mb-0 fs-6 fw-bold"><i class="fa-solid fa-file-excel me-2"></i> {tr("sup_heading", current_lang)}</h5>
                </div>
                <div class="card-body p-4">
                    <p class="text-muted small mb-4">{tr("sup_subtitle", current_lang)}</p>
                    <form method="GET" action="/reports/export/excel" class="row g-3">
                        <input type="hidden" name="device_id" value="ALL">
                        <div class="col-md-6">
                            <label class="form-label text-secondary small fw-bold">{tr("start_date", current_lang)}</label>
                            <input type="date" class="form-control" name="start_date" value="{start_date}" required>
                        </div>
                        <div class="col-md-6">
                            <label class="form-label text-secondary small fw-bold">{tr("end_date", current_lang)}</label>
                            <input type="date" class="form-control" name="end_date" value="{end_date}" required>
                        </div>
                        <div class="col-12 pt-2">
                            <button type="submit" class="btn btn-success w-100 py-2 fw-semibold shadow-sm">
                                <i class="fa-solid fa-file-excel me-2"></i> {tr("sup_export_btn", current_lang)}
                            </button>
                        </div>
                    </form>
                    <hr class="my-4">
                    <form method="GET" action="/summary/export/excel" class="row g-3">
                        <div class="col-md-6">
                            <label class="form-label text-secondary small fw-bold">{tr("start_date", current_lang)}</label>
                            <input type="date" class="form-control" name="start_date" value="{start_date}" required>
                        </div>
                        <div class="col-md-6">
                            <label class="form-label text-secondary small fw-bold">{tr("end_date", current_lang)}</label>
                            <input type="date" class="form-control" name="end_date" value="{end_date}" required>
                        </div>
                        <div class="col-12 pt-1">
                            <button type="submit" class="btn btn-primary w-100 py-2 fw-semibold shadow-sm">
                                <i class="fa-solid fa-file-excel me-2"></i> {tr("export_summary_excel", current_lang)}
                            </button>
                        </div>
                    </form>
                    <div class="alert alert-secondary mt-4 mb-0 py-2 small text-center">
                        <i class="fa-solid fa-lock me-1"></i> {tr("sup_notice", current_lang)}
                    </div>
                </div>
            </div>
        </div>
    </body>
    </html>
    """
    return HTMLResponse(content=html_content)



@app.get("/viewer-export", response_class=HTMLResponse)
async def viewer_export_page(request: Request, start_date: str = None, end_date: str = None, drh_session: str = Cookie(None), lang: str = Cookie(None)):
    role = get_session_role(drh_session)
    if not role:
        return RedirectResponse(url="/login", status_code=303)
    if role == "supervisor":
        return RedirectResponse(url="/supervisor", status_code=303)
    if role == "drh":
        return RedirectResponse(url="/", status_code=303)

    current_lang = get_lang(request, lang)
    today_str = date.today().isoformat()
    start_date = start_date or today_str
    end_date = end_date or today_str

    html_content = f"""
    <!DOCTYPE html>
    <html lang="{current_lang}">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{tr("sup_page_title", current_lang)}</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
        <style>
            body {{
                background-color: #f4f7f6;
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            }}
            .navbar {{
                background: linear-gradient(135deg, #1e293b, #0f172a) !important;
            }}
            .card {{
                border: none;
                border-radius: 12px;
                box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
            }}
        </style>
    </head>
    <body>
        <nav class="navbar navbar-dark shadow-sm py-3">
            <div class="container-fluid px-4">
                <span class="navbar-brand mb-0 h5 fw-bold"><i class="fa-solid fa-file-export me-2"></i> {tr("sup_header", current_lang)}</span>
                <div class="d-flex align-items-center gap-2">
                    {lang_switcher_html(current_lang)}
                    <a href="/logout" class="btn btn-danger btn-sm px-3 ms-1"><i class="fa-solid fa-right-from-bracket me-1"></i> {tr("nav_logout", current_lang)}</a>
                </div>
            </div>
        </nav>

        <div class="container my-5" style="max-width: 640px;">
            <div class="card shadow-sm">
                <div class="card-header bg-primary text-white py-3">
                    <h5 class="mb-0 fs-6 fw-bold"><i class="fa-solid fa-file-excel me-2"></i> {tr("sup_heading", current_lang)}</h5>
                </div>
                <div class="card-body p-4">
                    <p class="text-muted small mb-4">{tr("sup_subtitle", current_lang)}</p>
                    <form method="GET" action="/reports/export/excel" class="row g-3">
                        <input type="hidden" name="device_id" value="ALL">
                        <div class="col-md-6">
                            <label class="form-label text-secondary small fw-bold">{tr("start_date", current_lang)}</label>
                            <input type="date" class="form-control" name="start_date" value="{start_date}" required>
                        </div>
                        <div class="col-md-6">
                            <label class="form-label text-secondary small fw-bold">{tr("end_date", current_lang)}</label>
                            <input type="date" class="form-control" name="end_date" value="{end_date}" required>
                        </div>
                        <div class="col-12 pt-2">
                            <button type="submit" class="btn btn-success w-100 py-2 fw-semibold shadow-sm">
                                <i class="fa-solid fa-file-excel me-2"></i> {tr("sup_export_btn", current_lang)}
                            </button>
                        </div>
                    </form>
                    <hr class="my-4">
                    <form method="GET" action="/summary/export/excel" class="row g-3">
                        <div class="col-md-6">
                            <label class="form-label text-secondary small fw-bold">{tr("start_date", current_lang)}</label>
                            <input type="date" class="form-control" name="start_date" value="{start_date}" required>
                        </div>
                        <div class="col-md-6">
                            <label class="form-label text-secondary small fw-bold">{tr("end_date", current_lang)}</label>
                            <input type="date" class="form-control" name="end_date" value="{end_date}" required>
                        </div>
                        <div class="col-12 pt-1">
                            <button type="submit" class="btn btn-primary w-100 py-2 fw-semibold shadow-sm">
                                <i class="fa-solid fa-file-excel me-2"></i> {tr("export_summary_excel", current_lang)}
                            </button>
                        </div>
                    </form>
                    <div class="alert alert-secondary mt-4 mb-0 py-2 small text-center">
                        <i class="fa-solid fa-lock me-1"></i> {tr("sup_notice", current_lang)}
                    </div>
                </div>
            </div>
        </div>
    </body>
    </html>
    """
    return HTMLResponse(content=html_content)

@app.head("/")
async def head_root():
    return Response(status_code=200)


@app.get("/", response_class=HTMLResponse)
async def dashboard(request: Request, error: str = None, success: str = None, drh_session: str = Cookie(None), lang: str = Cookie(None)):
    role = get_session_role(drh_session)
    if not role:
        return RedirectResponse(url="/login", status_code=303)
    if role == "supervisor":
        return RedirectResponse(url="/supervisor", status_code=303)
    if role == "viewer":
        return RedirectResponse(url="/viewer-export", status_code=303)

    is_readonly = (role == "viewer")
    current_lang = get_lang(request, lang)
    employees = []
    devices = []
    schedules = []

    conn = None
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            try:
                cursor.execute("""
                    SELECT e.*, s.name as schedule_name
                    FROM employees e
                    LEFT JOIN work_schedules s ON e.schedule_id = s.id
                    ORDER BY e.first_name ASC
                """)
                employees = cursor.fetchall()
            except Exception:
                conn.rollback()
                employees = []

            try:
                cursor.execute("SELECT * FROM work_schedules ORDER BY id ASC")
                schedules = cursor.fetchall()
            except Exception:
                conn.rollback()
                schedules = []

            cursor.execute("SELECT DISTINCT device_id FROM attendance_raw_logs WHERE device_id IS NOT NULL")
            raw_devs = [row['device_id'] for row in cursor.fetchall()]

            cursor.execute("SELECT * FROM devices_mapping")
            mappings = {row['device_key']: row['friendly_name'] for row in cursor.fetchall()}

            all_keys = set(raw_devs).union(set(mappings.keys()))

            for dev in all_keys:
                devices.append({
                    "key": dev,
                    "name": mappings.get(dev, dev)
                })

    except Exception as e:
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        print(f"Attention - Connexion BDD non établie : {e}")
    finally:
        if conn:
            try:
                conn.close()
            except Exception:
                pass

    alert_html = ""
    if error:
        alert_html = f'<div class="alert alert-danger alert-dismissible fade show" role="alert"><i class="fa-solid fa-circle-exclamation me-2"></i>{error}<button type="button" class="btn-close" data-bs-dismiss="alert"></button></div>'
    elif success:
        alert_html = f'<div class="alert alert-success alert-dismissible fade show" role="alert"><i class="fa-solid fa-circle-check me-2"></i>{success}<button type="button" class="btn-close" data-bs-dismiss="alert"></button></div>'

    emp_rows_html = ""
    if employees:
        for emp in employees:
            is_act = emp.get('is_active', True)
            status_badge = f'<span class="badge bg-success bg-opacity-10 text-success">{tr("emp_active", current_lang)}</span>' if is_act else f'<span class="badge bg-danger bg-opacity-10 text-danger">{tr("emp_inactive", current_lang)}</span>'
            emp_rows_html += f"""
                <tr>
                    <td class="ps-3"><code>{emp['badge_or_card_id']}</code></td>
                    <td class="fw-semibold">{emp['first_name']} {emp['last_name']}</td>
                    <td>{emp.get('department') or '-'}</td>
                    <td><span class="badge bg-light text-dark border">{emp.get('schedule_name') or 'Standard'}</span></td>
                    <td>{status_badge}</td>
                    <td class="pe-3 text-end">
                        <a href="/employees/edit/{emp['badge_or_card_id']}" class="btn btn-sm btn-outline-primary me-1" title="{tr("emp_edit_header", current_lang)}"><i class="fa-solid fa-pen-to-square"></i></a>
                        <a href="/employees/delete/{emp['badge_or_card_id']}" class="btn btn-sm btn-outline-danger" title="{tr("action", current_lang)}" onclick="return confirm('{tr("emp_del_confirm", current_lang)}');"><i class="fa-solid fa-trash"></i></a>
                    </td>
                </tr>
            """
    else:
        emp_rows_html = f"""
                <tr>
                    <td colspan="6" class="text-center py-4 text-muted">{tr("emp_none", current_lang)}</td>
                </tr>
        """

    sched_rows_html = ""
    sched_options_html = ""
    for s in schedules:
        sched_options_html += f'<option value="{s["id"]}">{s["name"]} ({str(s["start_time"])[:5]} - {str(s["end_time"])[:5]})</option>'
        sched_rows_html += f"""
            <tr>
                <td class="ps-3 fw-semibold">{s['name']}</td>
                <td><span class="badge bg-primary bg-opacity-10 text-primary">{str(s['start_time'])[:5]}</span></td>
                <td><span class="badge bg-secondary bg-opacity-10 text-secondary">{str(s['end_time'])[:5]}</span></td>
                <td>{s['late_tolerance_minutes']} min</td>
                <td class="text-end pe-3">
                    <a href="/schedules/delete/{s['id']}" class="btn btn-sm btn-outline-danger" title="{tr("action", current_lang)}" onclick="return confirm('{tr("sched_del_confirm", current_lang)}');"><i class="fa-solid fa-trash"></i></a>
                </td>
            </tr>
        """

    dev_rows_html = ""
    if devices:
        for d in devices:
            dev_rows_html += f"""
                <tr>
                    <td class="ps-3"><code>{d['key']}</code></td>
                    <td>
                        <form action="/devices/rename" method="POST" class="d-flex gap-2">
                            <input type="hidden" name="device_key" value="{d['key']}">
                            <input type="text" class="form-control form-control-sm" name="friendly_name" value="{d['name']}" required>
                            <button type="submit" class="btn btn-sm btn-success text-nowrap"><i class="fa-solid fa-floppy-disk"></i> {tr("save_btn", current_lang)}</button>
                        </form>
                    </td>
                    <td class="text-end pe-3">
                        <a href="/devices/delete/{d['key']}" class="btn btn-sm btn-outline-danger" title="{tr("action", current_lang)}" onclick="return confirm('{tr("dev_del_confirm", current_lang)}');"><i class="fa-solid fa-trash"></i></a>
                    </td>
                </tr>
            """
    else:
        dev_rows_html = f'<tr><td colspan="3" class="text-center py-3 text-muted">{tr("dev_none", current_lang)}</td></tr>'

    import_section_html = "" if is_readonly else f"""
            <!-- SECTION IMPORTATION MASSIVE EXCEL/CSV -->
            <div class="card shadow-sm mb-4">
                <div class="card-header bg-dark text-white py-3">
                    <h5 class="mb-0 fs-6 fw-bold"><i class="fa-solid fa-file-excel me-2 text-success"></i> {tr("import_title", current_lang)}</h5>
                </div>
                <div class="card-body p-4">
                    <form action="/employees/import" method="POST" enctype="multipart/form-data" class="row g-3 align-items-center">
                        <div class="col-md-8">
                            <label class="form-label text-secondary small fw-bold">{tr("import_label", current_lang)}</label>
                            <input type="file" class="form-control" name="file" accept=".xlsx, .xls, .csv" required>
                            <div class="form-text">
                                {tr("import_help", current_lang)} <code>badge</code>, <code>prenom</code>, <code>nom</code>, <code>departement</code>.
                            </div>
                        </div>
                        <div class="col-md-4 d-flex align-items-end">
                            <button type="submit" class="btn btn-success w-100 py-2 fw-semibold shadow-sm"><i class="fa-solid fa-file-import me-2"></i> {tr("import_btn", current_lang)}</button>
                        </div>
                    </form>
                </div>
            </div>
    """

    sched_form_html = "" if is_readonly else f"""
                <div class="col-lg-5">
                    <div class="card h-100">
                        <div class="card-header bg-success text-white py-3">
                            <h5 class="mb-0 fs-6 fw-bold"><i class="fa-solid fa-clock me-2"></i> {tr("sched_add_title", current_lang)}</h5>
                        </div>
                        <div class="card-body p-4">
                            <form action="/schedules/add" method="POST">
                                <div class="mb-3">
                                    <label class="form-label text-secondary small fw-bold">{tr("sched_name", current_lang)}</label>
                                    <input type="text" class="form-control" name="name" placeholder="{tr("sched_name_ph", current_lang)}" required>
                                </div>
                                <div class="row">
                                    <div class="col-md-6 mb-3">
                                        <label class="form-label text-secondary small fw-bold">{tr("sched_start", current_lang)}</label>
                                        <input type="time" class="form-control" name="start_time" value="08:00" required>
                                    </div>
                                    <div class="col-md-6 mb-3">
                                        <label class="form-label text-secondary small fw-bold">{tr("sched_end", current_lang)}</label>
                                        <input type="time" class="form-control" name="end_time" value="17:00" required>
                                    </div>
                                </div>
                                <div class="mb-3">
                                    <label class="form-label text-secondary small fw-bold">{tr("sched_tol", current_lang)}</label>
                                    <input type="number" class="form-control" name="late_tolerance_minutes" value="30" required>
                                </div>
                                <button type="submit" class="btn btn-success w-100 py-2 fw-semibold shadow-sm"><i class="fa-solid fa-plus me-2"></i> {tr("sched_add_btn", current_lang)}</button>
                            </form>
                        </div>
                    </div>
                </div>
    """

    dev_add_form_html = "" if is_readonly else f"""
                <div class="card-body bg-light border-bottom p-3">
                    <form action="/devices/add" method="POST" class="row g-2 align-items-center">
                        <div class="col-md-5">
                            <input type="text" class="form-control form-control-sm" name="device_key" placeholder="{tr("dev_key_ph", current_lang)}" required>
                        </div>
                        <div class="col-md-5">
                            <input type="text" class="form-control form-control-sm" name="friendly_name" placeholder="{tr("dev_name_ph", current_lang)}" required>
                        </div>
                        <div class="col-md-2">
                            <button type="submit" class="btn btn-sm btn-primary w-100 fw-semibold"><i class="fa-solid fa-plus me-1"></i> {tr("add_btn", current_lang)}</button>
                        </div>
                    </form>
                </div>
    """

    emp_add_form_html = "" if is_readonly else f"""
                <div class="col-lg-4">
                    <div class="card">
                        <div class="card-header bg-primary text-white py-3">
                            <h5 class="mb-0 fs-6 fw-bold"><i class="fa-solid fa-user-plus me-2"></i> {tr("emp_add_title", current_lang)}</h5>
                        </div>
                        <div class="card-body p-4">
                            <form action="/employees/add" method="POST">
                                <div class="mb-3">
                                    <label class="form-label text-secondary small fw-bold">{tr("emp_badge", current_lang)}</label>
                                    <input type="text" class="form-control" name="badge_or_card_id" placeholder="ex: 02132" required>
                                </div>
                                <div class="mb-3">
                                    <label class="form-label text-secondary small fw-bold">{tr("emp_firstname", current_lang)}</label>
                                    <input type="text" class="form-control" name="first_name" placeholder="ex: Jonathan" required>
                                </div>
                                <div class="mb-3">
                                    <label class="form-label text-secondary small fw-bold">{tr("emp_lastname", current_lang)}</label>
                                    <input type="text" class="form-control" name="last_name" placeholder="ex: Lukoki" required>
                                </div>
                                <div class="mb-3">
                                    <label class="form-label text-secondary small fw-bold">{tr("emp_dept", current_lang)}</label>
                                    <input type="text" class="form-control" name="department" placeholder="ex: IT">
                                </div>
                                <div class="mb-3">
                                    <label class="form-label text-secondary small fw-bold">{tr("emp_sched", current_lang)}</label>
                                    <select class="form-select" name="schedule_id">
                                        {sched_options_html}
                                    </select>
                                </div>
                                <button type="submit" class="btn btn-primary w-100 py-2 fw-semibold shadow-sm"><i class="fa-solid fa-check me-2"></i> {tr("emp_save_btn", current_lang)}</button>
                            </form>
                        </div>
                    </div>
                </div>
    """

    readonly_banner_html = "" if not is_readonly else f'<span class="badge bg-warning text-dark ms-2"><i class="fa-solid fa-eye me-1"></i> {tr("readonly_badge", current_lang)}</span>'
    sched_col_cls = "col-12" if is_readonly else "col-lg-7"
    emp_col_cls = "col-12" if is_readonly else "col-lg-8"
    sched_th_action = "" if is_readonly else f'<th class="text-end pe-3">{tr("action", current_lang)}</th>'
    dev_th_action = "" if is_readonly else f'<th class="text-end pe-3">{tr("action", current_lang)}</th>'
    emp_th_action = "" if is_readonly else f'<th class="pe-3 text-end">{tr("emp_col_actions", current_lang)}</th>'

    html_content = f"""
    <!DOCTYPE html>
    <html lang="{current_lang}">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{tr("nav_brand", current_lang)}</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
        <style>
            body {{
                background-color: #f4f7f6;
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            }}
            .navbar {{
                background: linear-gradient(135deg, #1e293b, #0f172a) !important;
            }}
            .card {{
                border: none;
                border-radius: 12px;
                box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
            }}
            .card-header {{
                background-color: #ffffff;
                border-bottom: 1px solid #f1f5f9;
                border-top-left-radius: 12px !important;
                border-top-right-radius: 12px !important;
                font-weight: 600;
            }}
            .table th {{
                font-weight: 600;
                color: #475569;
                background-color: #f8fafc !important;
            }}
            .badge {{
                padding: 0.5em 0.8em;
                font-weight: 500;
                border-radius: 6px;
            }}
        </style>
    </head>
    <body>
        <nav class="navbar navbar-dark shadow-sm py-3">
            <div class="container-fluid px-4">
                <span class="navbar-brand mb-0 h5 fw-bold"><i class="fa-solid fa-clock-rotate-left me-2"></i> {tr("nav_brand", current_lang)}</span>
                <div class="d-flex align-items-center gap-2 flex-wrap">
                    <a href="/" class="btn btn-outline-light btn-sm px-3 active"><i class="fa-solid fa-house me-1"></i> {tr("nav_dashboard", current_lang)}</a>
                    <a href="/reports" class="btn btn-outline-light btn-sm px-3"><i class="fa-solid fa-calendar-days me-1"></i> {tr("nav_reports", current_lang)}</a>
                    <a href="/summary" class="btn btn-light btn-sm px-3 fw-semibold"><i class="fa-solid fa-chart-pie me-1"></i> {tr("nav_summary", current_lang)}</a>
                    {lang_switcher_html(current_lang)}
                    <a href="/logout" class="btn btn-danger btn-sm px-3 ms-1"><i class="fa-solid fa-right-from-bracket me-1"></i> {tr("nav_logout", current_lang)}</a>
                </div>
            </div>
        </nav>

        <div class="container-fluid px-4 my-4">
            {alert_html}

            <div class="row g-4 mb-4">
                <div class="col-md-4">
                    <div class="card border-start border-primary border-4">
                        <div class="card-body d-flex justify-content-between align-items-center">
                            <div>
                                <h6 class="text-muted fw-semibold mb-1">{tr("kpi_employees", current_lang)}</h6>
                                <h3 class="fw-bold text-dark mb-0">{len(employees)}</h3>
                            </div>
                            <div class="text-primary bg-primary bg-opacity-10 p-3 rounded-circle">
                                <i class="fa-solid fa-users fa-xl"></i>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="col-md-4">
                    <div class="card border-start border-success border-4">
                        <div class="card-body d-flex justify-content-between align-items-center">
                            <div>
                                <h6 class="text-muted fw-semibold mb-1">{tr("kpi_status", current_lang)}</h6>
                                <h3 class="fw-bold text-success mb-0">{tr("kpi_online", current_lang)}</h3>
                            </div>
                            <div class="text-success bg-success bg-opacity-10 p-3 rounded-circle">
                                <i class="fa-solid fa-server fa-xl"></i>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="col-md-4">
                    <div class="card border-start border-info border-4">
                        <div class="card-body">
                            <h6 class="text-muted fw-semibold mb-1">{tr("kpi_devices", current_lang)}</h6>
                            <h3 class="fw-bold text-info mb-0">{len(devices)} {tr("detected", current_lang)}</h3>
                        </div>
                    </div>
                </div>
            </div>

            <!-- SECTION IMPORTATION MASSIVE EXCEL/CSV -->
            <div class="card shadow-sm mb-4">
                <div class="card-header bg-dark text-white py-3">
                    <h5 class="mb-0 fs-6 fw-bold"><i class="fa-solid fa-file-excel me-2 text-success"></i> {tr("import_title", current_lang)}</h5>
                </div>
                <div class="card-body p-4">
                    <form action="/employees/import" method="POST" enctype="multipart/form-data" class="row g-3 align-items-center">
                        <div class="col-md-8">
                            <label class="form-label text-secondary small fw-bold">{tr("import_label", current_lang)}</label>
                            <input type="file" class="form-control" name="file" accept=".xlsx, .xls, .csv" required>
                            <div class="form-text">
                                {tr("import_help", current_lang)} <code>badge</code>, <code>prenom</code>, <code>nom</code>, <code>departement</code>.
                            </div>
                        </div>
                        <div class="col-md-4 d-flex align-items-end">
                            <button type="submit" class="btn btn-success w-100 py-2 fw-semibold shadow-sm"><i class="fa-solid fa-file-import me-2"></i> {tr("import_btn", current_lang)}</button>
                        </div>
                    </form>
                </div>
            </div>

            <!-- SECTION GESTION DES HORAIRES DE TRAVAIL -->
            <div class="row g-4 mb-4">
                <div class="col-lg-5">
                    <div class="card h-100">
                        <div class="card-header bg-success text-white py-3">
                            <h5 class="mb-0 fs-6 fw-bold"><i class="fa-solid fa-clock me-2"></i> {tr("sched_add_title", current_lang)}</h5>
                        </div>
                        <div class="card-body p-4">
                            <form action="/schedules/add" method="POST">
                                <div class="mb-3">
                                    <label class="form-label text-secondary small fw-bold">{tr("sched_name", current_lang)}</label>
                                    <input type="text" class="form-control" name="name" placeholder="{tr("sched_name_ph", current_lang)}" required>
                                </div>
                                <div class="row">
                                    <div class="col-md-6 mb-3">
                                        <label class="form-label text-secondary small fw-bold">{tr("sched_start", current_lang)}</label>
                                        <input type="time" class="form-control" name="start_time" value="08:00" required>
                                    </div>
                                    <div class="col-md-6 mb-3">
                                        <label class="form-label text-secondary small fw-bold">{tr("sched_end", current_lang)}</label>
                                        <input type="time" class="form-control" name="end_time" value="17:00" required>
                                    </div>
                                </div>
                                <div class="mb-3">
                                    <label class="form-label text-secondary small fw-bold">{tr("sched_tol", current_lang)}</label>
                                    <input type="number" class="form-control" name="late_tolerance_minutes" value="30" required>
                                </div>
                                <button type="submit" class="btn btn-success w-100 py-2 fw-semibold shadow-sm"><i class="fa-solid fa-plus me-2"></i> {tr("sched_add_btn", current_lang)}</button>
                            </form>
                        </div>
                    </div>
                </div>
                <div class="col-lg-7">
                    <div class="card h-100">
                        <div class="card-header bg-dark text-white py-3">
                            <h5 class="mb-0 fs-6 fw-bold"><i class="fa-solid fa-business-time me-2"></i> {tr("sched_list_title", current_lang)}</h5>
                        </div>
                        <div class="card-body p-0">
                            <div class="table-responsive">
                                <table class="table table-hover align-middle mb-0">
                                    <thead>
                                        <tr>
                                            <th class="ps-3">{tr("sched_col_name", current_lang)}</th>
                                            <th>{tr("sched_col_in", current_lang)}</th>
                                            <th>{tr("sched_col_out", current_lang)}</th>
                                            <th>{tr("sched_col_tol", current_lang)}</th>
                                            <th class="text-end pe-3">{tr("action", current_lang)}</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {sched_rows_html}
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <!-- SECTION TERMINAUX -->
            <div class="card shadow-sm mb-4">
                <div class="card-header bg-secondary text-white py-3 d-flex justify-content-between align-items-center">
                    <h5 class="mb-0 fs-6 fw-bold"><i class="fa-solid fa-network-wired me-2"></i> {tr("dev_title", current_lang)}</h5>
                </div>
                <div class="card-body bg-light border-bottom p-3">
                    <form action="/devices/add" method="POST" class="row g-2 align-items-center">
                        <div class="col-md-5">
                            <input type="text" class="form-control form-control-sm" name="device_key" placeholder="{tr("dev_key_ph", current_lang)}" required>
                        </div>
                        <div class="col-md-5">
                            <input type="text" class="form-control form-control-sm" name="friendly_name" placeholder="{tr("dev_name_ph", current_lang)}" required>
                        </div>
                        <div class="col-md-2">
                            <button type="submit" class="btn btn-sm btn-primary w-100 fw-semibold"><i class="fa-solid fa-plus me-1"></i> {tr("add_btn", current_lang)}</button>
                        </div>
                    </form>
                </div>
                <div class="card-body p-0">
                    <div class="table-responsive">
                        <table class="table table-hover align-middle mb-0">
                            <thead>
                                <tr>
                                    <th class="ps-3">{tr("dev_col_key", current_lang)}</th>
                                    <th>{tr("dev_col_name", current_lang)}</th>
                                    <th class="text-end pe-3">{tr("action", current_lang)}</th>
                                </tr>
                            </thead>
                            <tbody>
                                {dev_rows_html}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- SECTION EMPLOYÉS -->
            <div class="row g-4">
                <div class="col-lg-4">
                    <div class="card">
                        <div class="card-header bg-primary text-white py-3">
                            <h5 class="mb-0 fs-6 fw-bold"><i class="fa-solid fa-user-plus me-2"></i> {tr("emp_add_title", current_lang)}</h5>
                        </div>
                        <div class="card-body p-4">
                            <form action="/employees/add" method="POST">
                                <div class="mb-3">
                                    <label class="form-label text-secondary small fw-bold">{tr("emp_badge", current_lang)}</label>
                                    <input type="text" class="form-control" name="badge_or_card_id" placeholder="ex: 02132" required>
                                </div>
                                <div class="mb-3">
                                    <label class="form-label text-secondary small fw-bold">{tr("emp_firstname", current_lang)}</label>
                                    <input type="text" class="form-control" name="first_name" placeholder="ex: Jonathan" required>
                                </div>
                                <div class="mb-3">
                                    <label class="form-label text-secondary small fw-bold">{tr("emp_lastname", current_lang)}</label>
                                    <input type="text" class="form-control" name="last_name" placeholder="ex: Lukoki" required>
                                </div>
                                <div class="mb-3">
                                    <label class="form-label text-secondary small fw-bold">{tr("emp_dept", current_lang)}</label>
                                    <input type="text" class="form-control" name="department" placeholder="ex: IT">
                                </div>
                                <div class="mb-3">
                                    <label class="form-label text-secondary small fw-bold">{tr("emp_sched", current_lang)}</label>
                                    <select class="form-select" name="schedule_id">
                                        {sched_options_html}
                                    </select>
                                </div>
                                <button type="submit" class="btn btn-primary w-100 py-2 fw-semibold shadow-sm"><i class="fa-solid fa-check me-2"></i> {tr("emp_save_btn", current_lang)}</button>
                            </form>
                        </div>
                    </div>
                </div>

                <div class="col-lg-8">
                    <div class="card mt-0">
                        <div class="card-header py-3 d-flex justify-content-between align-items-center">
                            <h5 class="mb-0 fs-6 fw-bold text-dark"><i class="fa-solid fa-users-gear me-2 text-primary"></i> {tr("emp_list_title", current_lang)}</h5>
                            <a href="/summary" class="btn btn-sm btn-outline-primary fw-semibold"><i class="fa-solid fa-chart-pie me-1"></i> {tr("emp_summary_link", current_lang)}</a>
                        </div>
                        <div class="card-body p-0">
                            <div class="table-responsive">
                                <table class="table table-hover align-middle mb-0">
                                    <thead>
                                        <tr>
                                            <th class="ps-3">{tr("emp_col_badge", current_lang)}</th>
                                            <th>{tr("emp_col_fullname", current_lang)}</th>
                                            <th>{tr("emp_col_dept", current_lang)}</th>
                                            <th>{tr("emp_col_sched", current_lang)}</th>
                                            <th>{tr("emp_col_status", current_lang)}</th>
                                            <th class="pe-3 text-end">{tr("emp_col_actions", current_lang)}</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {emp_rows_html}
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    </body>
    </html>
    """
    return HTMLResponse(content=html_content)


# --- ROUTE IMPORTATION D'EMPLOYÉS (EXCEL / CSV) ---
@app.post("/employees/import")
async def import_employees(file: UploadFile = File(...), drh_session: str = Cookie(None)):
    role = get_session_role(drh_session)
    if not role:
        return RedirectResponse(url="/login", status_code=303)
    if role == "supervisor":
        return RedirectResponse(url="/supervisor", status_code=303)
    if role == "viewer":
        return RedirectResponse(url="/viewer-export" if role == "viewer" else "/summary", status_code=303)

    conn = None
    try:
        contents = await file.read()
        filename = file.filename.lower()

        if filename.endswith('.csv'):
            df = pd.read_csv(io.BytesIO(contents))
        elif filename.endswith('.xlsx') or filename.endswith('.xls'):
            df = pd.read_excel(io.BytesIO(contents))
        else:
            return RedirectResponse(url="/?error=Format+de+fichier+non+supporté.+Veuillez+utiliser+.xlsx,+xls+ou+.csv", status_code=303)

        df.columns = [str(col).strip().lower() for col in df.columns]

        col_badge = next((c for c in df.columns if c in ['badge', 'badge_id', 'matricule', 'id', 'cardno', 'card_no', 'employee_id']), None)
        col_prenom = next((c for c in df.columns if c in ['prenom', 'prénom', 'first_name', 'firstname', 'prenoms']), None)
        col_nom = next((c for c in df.columns if c in ['nom', 'last_name', 'lastname', 'name']), None)
        col_dept = next((c for c in df.columns if c in ['departement', 'département', 'department', 'service', 'dept']), None)

        if not col_badge or not col_nom:
            return RedirectResponse(
                url="/?error=Le+fichier+doit+contenir+au+moins+les+colonnes+matricule+(ou+badge)+et+nom",
                status_code=303
            )

        imported_count = 0
        conn = get_db_connection()
        with conn.cursor() as cursor:
            for _, row in df.iterrows():
                badge_val = str(row[col_badge]).strip() if pd.notna(row[col_badge]) else ""
                if not badge_val or badge_val.lower() == 'nan':
                    continue

                prenom_val = str(row[col_prenom]).strip() if col_prenom and pd.notna(row[col_prenom]) else ""
                nom_val = str(row[col_nom]).strip() if col_nom and pd.notna(row[col_nom]) else ""
                dept_val = str(row[col_dept]).strip() if col_dept and pd.notna(row[col_dept]) else None

                cursor.execute(
                    """
                    INSERT INTO employees (badge_or_card_id, first_name, last_name, department, is_active)
                    VALUES (%s, %s, %s, %s, TRUE)
                    ON CONFLICT (badge_or_card_id) DO UPDATE
                    SET first_name = EXCLUDED.first_name,
                        last_name = EXCLUDED.last_name,
                        department = EXCLUDED.department
                    """,
                    (badge_val, prenom_val, nom_val, dept_val)
                )
                imported_count += 1
            conn.commit()
        conn.close()

        return RedirectResponse(url=f"/?success={imported_count}+employé(s)+importé(s)+avec+succès", status_code=303)

    except Exception as e:
        if conn:
            try:
                conn.rollback()
                conn.close()
            except Exception:
                pass
        print(f"Erreur lors de l'importation Excel : {e}")
        return RedirectResponse(url=f"/?error=Erreur+de+lecture+du+fichier:+{str(e)}", status_code=303)


# --- ROUTES GESTION DES HORAIRES ---
@app.post("/schedules/add")
async def add_schedule(
    name: str = Form(...),
    start_time: str = Form(...),
    end_time: str = Form(...),
    late_tolerance_minutes: int = Form(30),
    drh_session: str = Cookie(None)
):
    role = get_session_role(drh_session)
    if not role:
        return RedirectResponse(url="/login", status_code=303)
    if role == "supervisor":
        return RedirectResponse(url="/supervisor", status_code=303)
    if role == "viewer":
        return RedirectResponse(url="/viewer-export" if role == "viewer" else "/summary", status_code=303)
    conn = None
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            cursor.execute("""
                INSERT INTO work_schedules (name, start_time, end_time, late_tolerance_minutes)
                VALUES (%s, %s, %s, %s)
            """, (name.strip(), start_time, end_time, late_tolerance_minutes))
            conn.commit()
        conn.close()
    except Exception as e:
        if conn:
            try:
                conn.rollback()
                conn.close()
            except Exception:
                pass
        print(f"Erreur ajout horaire : {e}")
    return RedirectResponse(url="/", status_code=303)


@app.get("/schedules/delete/{sched_id}")
async def delete_schedule(sched_id: int, drh_session: str = Cookie(None)):
    role = get_session_role(drh_session)
    if not role:
        return RedirectResponse(url="/login", status_code=303)
    if role == "supervisor":
        return RedirectResponse(url="/supervisor", status_code=303)
    if role == "viewer":
        return RedirectResponse(url="/viewer-export" if role == "viewer" else "/summary", status_code=303)
    conn = None
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            cursor.execute("UPDATE employees SET schedule_id = NULL WHERE schedule_id = %s", (sched_id,))
            cursor.execute("DELETE FROM work_schedules WHERE id = %s", (sched_id,))
            conn.commit()
        conn.close()
    except Exception as e:
        if conn:
            try:
                conn.rollback()
                conn.close()
            except Exception:
                pass
        print(f"Erreur suppression horaire : {e}")
    return RedirectResponse(url="/", status_code=303)


@app.post("/devices/add")
async def add_device_manual(device_key: str = Form(...), friendly_name: str = Form(...), drh_session: str = Cookie(None)):
    role = get_session_role(drh_session)
    if not role:
        return RedirectResponse(url="/login", status_code=303)
    if role == "supervisor":
        return RedirectResponse(url="/supervisor", status_code=303)
    if role == "viewer":
        return RedirectResponse(url="/viewer-export" if role == "viewer" else "/summary", status_code=303)
    conn = None
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            cursor.execute("""
                INSERT INTO devices_mapping (device_key, friendly_name)
                VALUES (%s, %s)
                ON CONFLICT (device_key) DO UPDATE SET friendly_name = EXCLUDED.friendly_name
            """, (device_key.strip(), friendly_name.strip()))
            conn.commit()
        conn.close()
        return RedirectResponse(url="/?success=Terminal+ajouté+avec+succès", status_code=303)
    except Exception as e:
        if conn:
            try:
                conn.rollback()
                conn.close()
            except Exception:
                pass
        print(f"Erreur ajout manuel terminal : {e}")
        return RedirectResponse(url=f"/?error=Erreur+lors+de+l'ajout+du+terminal:+{str(e)}", status_code=303)


@app.post("/devices/rename")
async def rename_device(device_key: str = Form(...), friendly_name: str = Form(...), drh_session: str = Cookie(None)):
    role = get_session_role(drh_session)
    if not role:
        return RedirectResponse(url="/login", status_code=303)
    if role == "supervisor":
        return RedirectResponse(url="/supervisor", status_code=303)
    if role == "viewer":
        return RedirectResponse(url="/viewer-export" if role == "viewer" else "/summary", status_code=303)
    conn = None
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            cursor.execute("""
                INSERT INTO devices_mapping (device_key, friendly_name)
                VALUES (%s, %s)
                ON CONFLICT (device_key) DO UPDATE SET friendly_name = EXCLUDED.friendly_name
            """, (device_key.strip(), friendly_name.strip()))
            conn.commit()
        conn.close()
    except Exception as e:
        if conn:
            try:
                conn.rollback()
                conn.close()
            except Exception:
                pass
        print(f"Erreur lors du renommage du terminal : {e}")
    return RedirectResponse(url="/", status_code=303)


@app.get("/devices/delete/{device_key}")
async def delete_device(device_key: str, drh_session: str = Cookie(None)):
    role = get_session_role(drh_session)
    if not role:
        return RedirectResponse(url="/login", status_code=303)
    if role == "supervisor":
        return RedirectResponse(url="/supervisor", status_code=303)
    if role == "viewer":
        return RedirectResponse(url="/viewer-export" if role == "viewer" else "/summary", status_code=303)

    conn = None
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            cursor.execute("DELETE FROM devices_mapping WHERE device_key = %s", (device_key,))
            cursor.execute("DELETE FROM attendance_raw_logs WHERE device_id = %s", (device_key,))
            conn.commit()
        conn.close()
    except Exception as e:
        if conn:
            try:
                conn.rollback()
                conn.close()
            except Exception:
                pass
        print(f"Erreur lors de la suppression du terminal : {e}")

    return RedirectResponse(url="/", status_code=303)


@app.get("/reports/delete/{log_id}")
async def delete_attendance_log(log_id: int, drh_session: str = Cookie(None)):
    role = get_session_role(drh_session)
    if not role:
        return RedirectResponse(url="/login", status_code=303)
    if role == "supervisor":
        return RedirectResponse(url="/supervisor", status_code=303)
    if role == "viewer":
        return RedirectResponse(url="/reports", status_code=303)

    conn = None
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            cursor.execute("DELETE FROM attendance_raw_logs WHERE id = %s", (log_id,))
            conn.commit()
        conn.close()
    except Exception as e:
        if conn:
            try:
                conn.rollback()
                conn.close()
            except Exception:
                pass
        print(f"Erreur lors de la suppression du pointage : {e}")

    return RedirectResponse(url="/reports", status_code=303)


@app.get("/reports", response_class=HTMLResponse)
async def attendance_reports(request: Request, start_date: str = None, end_date: str = None, device_id: str = None, drh_session: str = Cookie(None), lang: str = Cookie(None)):
    role = get_session_role(drh_session)
    if not role:
        return RedirectResponse(url="/login", status_code=303)
    if role == "supervisor":
        return RedirectResponse(url="/supervisor", status_code=303)
    if role == "viewer":
        return RedirectResponse(url="/viewer-export", status_code=303)

    is_readonly = (role == "viewer")
    current_lang = get_lang(request, lang)
    unknown_label = tr("unknown", current_lang)
    reports = []
    devices = []
    mappings = {}
    today_str = date.today().isoformat()

    if not start_date:
        start_date = today_str
    if not end_date:
        end_date = today_str

    conn = None
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            cursor.execute("SELECT * FROM devices_mapping")
            mappings = {row['device_key']: row['friendly_name'] for row in cursor.fetchall()}

            cursor.execute("SELECT DISTINCT device_id FROM attendance_raw_logs WHERE device_id IS NOT NULL ORDER BY device_id")
            devices = [{"key": row['device_id'], "name": mappings.get(row['device_id'], row['device_id'])} for row in cursor.fetchall()]

            base_query = """
                SELECT
                    l.id,
                    l.badge_or_card_id,
                    COALESCE(e.first_name, %s) as first_name,
                    COALESCE(e.last_name, '') as last_name,
                    COALESCE(e.department, '-') as department,
                    s.start_time as schedule_start,
                    s.end_time as schedule_end,
                    s.late_tolerance_minutes,
                    DATE(l.punch_timestamp) as punch_date,
                    l.punch_timestamp as exact_punch,
                    l.device_id
                FROM attendance_raw_logs l
                LEFT JOIN employees e ON l.badge_or_card_id = e.badge_or_card_id
                LEFT JOIN work_schedules s ON e.schedule_id = s.id
                WHERE DATE(l.punch_timestamp) BETWEEN %s AND %s
            """

            params = [unknown_label, start_date, end_date]

            if device_id and device_id != "ALL":
                base_query += " AND l.device_id = %s"
                params.append(device_id)

            base_query += " ORDER BY l.punch_timestamp DESC"

            cursor.execute(base_query, tuple(params))
            raw_logs = cursor.fetchall()

            grouped = {}
            for r in raw_logs:
                key = (r['badge_or_card_id'], str(r['punch_date']), r['device_id'])
                if key not in grouped:
                    grouped[key] = {
                        "badge_or_card_id": r['badge_or_card_id'],
                        "first_name": r['first_name'],
                        "last_name": r['last_name'],
                        "schedule_start": r['schedule_start'] or time(8, 0),
                        "schedule_end": r['schedule_end'] or time(17, 0),
                        "late_tolerance_minutes": r['late_tolerance_minutes'] if r['late_tolerance_minutes'] is not None else 30,
                        "punch_date": r['punch_date'],
                        "device_id": r['device_id'],
                        "timestamps": []
                    }
                grouped[key]["timestamps"].append((r['id'], r['exact_punch']))

            for key, data in grouped.items():
                sorted_ts = sorted(data["timestamps"], key=lambda x: x[1])
                check_in_item = sorted_ts[0]
                check_out_item = sorted_ts[-1] if len(sorted_ts) > 1 else None

                reports.append({
                    "badge_or_card_id": data["badge_or_card_id"],
                    "first_name": data["first_name"],
                    "last_name": data["last_name"],
                    "schedule_start": data["schedule_start"],
                    "schedule_end": data["schedule_end"],
                    "late_tolerance_minutes": data["late_tolerance_minutes"],
                    "punch_date": data["punch_date"],
                    "check_in_id": check_in_item[0],
                    "check_in": check_in_item[1],
                    "check_out_id": check_out_item[0] if check_out_item else None,
                    "check_out": check_out_item[1] if check_out_item else None,
                    "device_id": data["device_id"]
                })

        conn.close()
    except Exception as e:
        if conn:
            try:
                conn.rollback()
                conn.close()
            except Exception:
                pass
        print(f"Erreur lors de la génération du rapport : {e}")

    report_rows_html = ""
    if reports:
        for row in reports:
            emp_name = f"{row['first_name']} {row['last_name']}".strip()
            check_in = row['check_in']
            check_out = row['check_out']
            raw_dev = row.get('device_id', '-')
            dev_display = mappings.get(raw_dev, raw_dev) if raw_dev in mappings else raw_dev

            check_in_str = str(check_in).split(' ')[1] if check_in else '-'
            check_out_str = str(check_out).split(' ')[1] if check_out else '-'

            del_id = row['check_out_id'] if row['check_out_id'] else row['check_in_id']

            sched_start = row['schedule_start']
            sched_end = row['schedule_end']
            tolerance = row['late_tolerance_minutes']
            p_date = row['punch_date'] if isinstance(row['punch_date'], date) else date.today()

            if not check_in or not check_out or check_in == check_out:
                report_rows_html += f"""
                            <tr>
                                <td><strong>{row['punch_date']}</strong></td>
                                <td><code>{row['badge_or_card_id']}</code></td>
                                <td class="text-start ps-4 fw-semibold">{emp_name}</td>
                                <td><span class="badge bg-danger text-white">{tr("late_missing", current_lang)} ({check_in_str})</span></td>
                                <td><span class="text-muted fst-italic">{tr("missing", current_lang)}</span></td>
                                <td>-</td>
                                <td><span class="badge bg-secondary">{dev_display}</span></td>
                                <td><span class="badge bg-danger text-white">{tr("absent_inc", current_lang)}</span></td>
                                <td class="no-print">
                                    <a href="/reports/delete/{del_id}" class="btn btn-sm btn-outline-danger" title="{tr("action", current_lang)}" onclick="return confirm('{tr("punch_del_confirm", current_lang)}');"><i class="fa-solid fa-trash"></i></a>
                                </td>
                            </tr>
                """
                continue

            check_out_t = check_out.time()

            duration = check_out - check_in
            total_seconds = int(duration.total_seconds())
            hours = total_seconds // 3600
            minutes = (total_seconds % 3600) // 60
            work_duration = f"{hours}h {minutes:02d}m"

            max_arrival_dt = datetime.combine(p_date, sched_start) + timedelta(minutes=tolerance)
            is_arrival_valid = check_in <= max_arrival_dt
            is_departure_valid = check_out_t >= sched_end

            if is_arrival_valid:
                arr_badge = f'<span class="badge bg-success text-white">{check_in_str}</span>'
            else:
                arr_badge = f'<span class="badge bg-danger text-white">{tr("late", current_lang)} ({check_in_str})</span>'

            if is_departure_valid:
                dep_badge = f'<span class="badge bg-success text-white">{check_out_str}</span>'
            else:
                dep_badge = f'<span class="badge bg-warning text-dark">{tr("early_leave", current_lang)} ({check_out_str})</span>'

            if is_arrival_valid and is_departure_valid:
                global_status = f'<span class="badge bg-success text-white">{tr("present_ok", current_lang)}</span>'
            else:
                global_status = f'<span class="badge bg-danger text-white">{tr("non_compliant", current_lang)}</span>'

            report_rows_html += f"""
                            <tr>
                                <td><strong>{row['punch_date']}</strong></td>
                                <td><code>{row['badge_or_card_id']}</code></td>
                                <td class="text-start ps-4 fw-semibold">{emp_name}</td>
                                <td>{arr_badge}</td>
                                <td>{dep_badge}</td>
                                <td><strong>{work_duration}</strong></td>
                                <td><span class="badge bg-light text-dark border">{dev_display}</span></td>
                                <td>{global_status}</td>
                                <td class="no-print">
                                    <a href="/reports/delete/{del_id}" class="btn btn-sm btn-outline-danger" title="{tr("action", current_lang)}" onclick="return confirm('{tr("punch_del_confirm", current_lang)}');"><i class="fa-solid fa-trash"></i></a>
                                </td>
                            </tr>
            """
    else:
        report_rows_html = f"""
                            <tr>
                                <td colspan="9" class="text-center py-4 text-muted">{tr("rep_none", current_lang)}</td>
                            </tr>
        """

    device_options = f'<option value="ALL">{tr("all_terminals", current_lang)}</option>'
    for d in devices:
        selected = 'selected' if device_id == d['key'] else ''
        device_options += f'<option value="{d["key"]}" {selected}>{d["name"]}</option>'

    dev_param = f"&device_id={device_id}" if device_id else ""
    export_excel_url = f"/reports/export/excel?start_date={start_date}&end_date={end_date}{dev_param}"

    html_content = f"""
    <!DOCTYPE html>
    <html lang="{current_lang}">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{tr("rep_page_title", current_lang)}</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
        <style>
            body {{
                background-color: #f4f7f6;
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            }}
            .navbar {{
                background: linear-gradient(135deg, #1e293b, #0f172a) !important;
            }}
            .card {{
                border: none;
                border-radius: 12px;
                box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
            }}
            .card-header {{
                background-color: #ffffff;
                border-bottom: 1px solid #f1f5f9;
                border-top-left-radius: 12px !important;
                border-top-right-radius: 12px !important;
                font-weight: 600;
            }}
            .table th {{
                font-weight: 600;
                color: #475569;
                background-color: #f8fafc !important;
            }}
            @media print {{
                @page {{
                    size: landscape;
                    margin: 5mm;
                }}
                .no-print {{
                    display: none !important;
                }}
                body {{
                    background-color: white !important;
                    font-size: 11pt;
                    width: 100% !important;
                    margin: 0 !important;
                }}
                .container-fluid {{
                    max-width: 100% !important;
                    width: 100% !important;
                    margin: 0 !important;
                    padding: 0 !important;
                }}
                .card {{
                    border: none !important;
                    box-shadow: none !important;
                    width: 100% !important;
                }}
                .card-body {{
                    padding: 0 !important;
                }}
                .table-responsive {{
                    overflow: visible !important;
                    width: 100% !important;
                }}
                .table {{
                    width: 100% !important;
                    max-width: 100% !important;
                }}
            }}
        </style>
    </head>
    <body>
        <nav class="navbar navbar-dark shadow-sm py-3 no-print">
            <div class="container-fluid px-4">
                <span class="navbar-brand mb-0 h5 fw-bold"><i class="fa-solid fa-calendar-days me-2"></i> {tr("rep_header", current_lang)}</span>
                <div class="d-flex align-items-center gap-2 flex-wrap">
                    <a href="/" class="btn btn-outline-light btn-sm px-3"><i class="fa-solid fa-house me-1"></i> {tr("nav_dashboard", current_lang)}</a>
                    <a href="/reports" class="btn btn-light btn-sm px-3 fw-semibold"><i class="fa-solid fa-calendar-days me-1"></i> {tr("nav_reports", current_lang)}</a>
                    <a href="/summary" class="btn btn-outline-light btn-sm px-3"><i class="fa-solid fa-chart-pie me-1"></i> {tr("nav_summary", current_lang)}</a>
                    {lang_switcher_html(current_lang)}
                    <a href="/logout" class="btn btn-danger btn-sm px-3 ms-1"><i class="fa-solid fa-right-from-bracket me-1"></i> {tr("nav_logout", current_lang)}</a>
                </div>
            </div>
        </nav>

        <div class="container-fluid px-4 my-4">
            <div class="card shadow-sm mb-4 no-print">
                <div class="card-header bg-primary text-white py-3">
                    <h5 class="mb-0 fs-6 fw-bold"><i class="fa-solid fa-filter me-2"></i> {tr("rep_filter_title", current_lang)}</h5>
                </div>
                <div class="card-body p-4">
                    <form method="GET" action="/reports" class="row g-3">
                        <div class="col-md-4">
                            <label class="form-label text-secondary small fw-bold">{tr("start_date", current_lang)}</label>
                            <input type="date" class="form-control" name="start_date" value="{start_date}" required>
                        </div>
                        <div class="col-md-4">
                            <label class="form-label text-secondary small fw-bold">{tr("end_date", current_lang)}</label>
                            <input type="date" class="form-control" name="end_date" value="{end_date}" required>
                        </div>
                        <div class="col-md-2">
                            <label class="form-label text-secondary small fw-bold">{tr("terminal", current_lang)}</label>
                            <select class="form-select" name="device_id">
                                {device_options}
                            </select>
                        </div>
                        <div class="col-md-2 d-flex align-items-end">
                            <button type="submit" class="btn btn-success w-100 py-2 fw-semibold shadow-sm"><i class="fa-solid fa-magnifying-glass me-1"></i> {tr("filter_btn", current_lang)}</button>
                        </div>
                    </form>
                </div>
            </div>

            <div class="card shadow-sm">
                <div class="card-header d-flex justify-content-between align-items-center py-3 flex-wrap gap-2">
                    <h5 class="mb-0 fs-6 fw-bold text-dark"><i class="fa-solid fa-clipboard-user me-2 text-primary"></i> {tr("rep_detail_title", current_lang)} ({len(reports)} {tr("records", current_lang)})</h5>
                    <div class="d-flex gap-2 no-print">
                        <a href="{export_excel_url}" class="btn btn-success btn-sm fw-semibold"><i class="fa-solid fa-file-excel me-1"></i> {tr("export_excel", current_lang)}</a>
                        <button onclick="window.print();" class="btn btn-outline-dark btn-sm fw-semibold"><i class="fa-solid fa-print me-1"></i> {tr("print_pdf", current_lang)}</button>
                    </div>
                </div>
                <div class="card-body p-0">
                    <div class="table-responsive">
                        <table class="table table-striped table-hover mb-0 align-middle text-center">
                            <thead>
                                <tr>
                                    <th>{tr("date", current_lang)}</th>
                                    <th>{tr("matricule", current_lang)}</th>
                                    <th class="text-start ps-4">{tr("employee", current_lang)}</th>
                                    <th>{tr("arrival", current_lang)}</th>
                                    <th>{tr("departure", current_lang)}</th>
                                    <th>{tr("hours", current_lang)}</th>
                                    <th>{tr("device", current_lang)}</th>
                                    <th>{tr("status", current_lang)}</th>
                                    <th class="no-print">{tr("action", current_lang)}</th>
                                </tr>
                            </thead>
                            <tbody>
                                {report_rows_html}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
    </body>
    </html>
    """
    return HTMLResponse(content=html_content)


# --- ROUTE D'EXPORTATION EXCEL DU RAPPORT JOURNALIER ---
@app.get("/reports/export/excel")
async def export_reports_excel(request: Request, start_date: str = None, end_date: str = None, device_id: str = None, drh_session: str = Cookie(None), lang: str = Cookie(None)):
    role = get_session_role(drh_session)
    if not role:
        return RedirectResponse(url="/login", status_code=303)

    current_lang = get_lang(request, lang)
    unknown_label = tr("unknown", current_lang)
    today_str = date.today().isoformat()
    start_date = start_date or today_str
    end_date = end_date or today_str

    conn = None
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            cursor.execute("SELECT * FROM devices_mapping")
            mappings = {row['device_key']: row['friendly_name'] for row in cursor.fetchall()}

            base_query = """
                SELECT
                    l.badge_or_card_id,
                    COALESCE(e.first_name, %s) as first_name,
                    COALESCE(e.last_name, '') as last_name,
                    COALESCE(e.department, '-') as department,
                    s.start_time as schedule_start,
                    s.end_time as schedule_end,
                    s.late_tolerance_minutes,
                    DATE(l.punch_timestamp) as punch_date,
                    l.punch_timestamp as exact_punch,
                    l.device_id
                FROM attendance_raw_logs l
                LEFT JOIN employees e ON l.badge_or_card_id = e.badge_or_card_id
                LEFT JOIN work_schedules s ON e.schedule_id = s.id
                WHERE DATE(l.punch_timestamp) BETWEEN %s AND %s
            """
            params = [unknown_label, start_date, end_date]
            if device_id and device_id != "ALL":
                base_query += " AND l.device_id = %s"
                params.append(device_id)

            base_query += " ORDER BY l.punch_timestamp DESC"
            cursor.execute(base_query, tuple(params))
            raw_logs = cursor.fetchall()
        conn.close()

        grouped = {}
        for r in raw_logs:
            key = (r['badge_or_card_id'], str(r['punch_date']), r['device_id'])
            if key not in grouped:
                grouped[key] = {"timestamps": [], "row": r}
            grouped[key]["timestamps"].append(r['exact_punch'])

        export_data = []
        for key, info in grouped.items():
            r = info["row"]
            sorted_ts = sorted(info["timestamps"])
            cin = sorted_ts[0]
            cout = sorted_ts[-1] if len(sorted_ts) > 1 else None

            cin_str = str(cin).split(' ')[1] if cin else '-'
            cout_str = str(cout).split(' ')[1] if cout else '-'
            emp_name = f"{r['first_name']} {r['last_name']}".strip()
            dev_display = mappings.get(r['device_id'], r['device_id'])

            if not cin or not cout or cin == cout:
                work_duration = "-"
                status = tr("absent_inc", current_lang)
            else:
                duration = cout - cin
                total_seconds = int(duration.total_seconds())
                hrs = total_seconds // 3600
                mins = (total_seconds % 3600) // 60
                work_duration = f"{hrs}h {mins:02d}m"
                status = tr("present_ok", current_lang)

            export_data.append({
                tr("excel_col_date", current_lang): str(r['punch_date']),
                tr("excel_col_badge", current_lang): r['badge_or_card_id'],
                tr("excel_col_emp", current_lang): emp_name,
                tr("excel_col_dept", current_lang): r['department'],
                tr("excel_col_in", current_lang): cin_str,
                tr("excel_col_out", current_lang): cout_str,
                tr("excel_col_hours", current_lang): work_duration,
                tr("excel_col_dev", current_lang): dev_display,
                tr("excel_col_status", current_lang): status
            })

        df = pd.DataFrame(export_data)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name=tr("excel_sheet_rep", current_lang)[:31])
        output.seek(0)

        filename = f"Rapport_Pointage_{start_date}_au_{end_date}.xlsx"
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )

    except Exception as e:
        if conn:
            try:
                conn.rollback()
                conn.close()
            except Exception:
                pass
        print(f"Erreur export Excel : {e}")
        return RedirectResponse(url="/reports", status_code=303)


@app.get("/summary", response_class=HTMLResponse)
async def attendance_summary(request: Request, start_date: str = None, end_date: str = None, drh_session: str = Cookie(None), lang: str = Cookie(None)):
    role = get_session_role(drh_session)
    if not role:
        return RedirectResponse(url="/login", status_code=303)
    if role == "supervisor":
        return RedirectResponse(url="/supervisor", status_code=303)
    if role == "viewer":
        return RedirectResponse(url="/viewer-export", status_code=303)

    is_readonly = (role == "viewer")
    current_lang = get_lang(request, lang)
    unknown_label = tr("unknown", current_lang)
    summary_rows = []
    today_str = date.today().isoformat()

    if not start_date:
        start_date = today_str
    if not end_date:
        end_date = today_str

    conn = None
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            query = """
                SELECT
                    COALESCE(e.first_name, %s) as first_name,
                    COALESCE(e.last_name, '') as last_name,
                    l.badge_or_card_id,
                    s.start_time as schedule_start,
                    s.end_time as schedule_end,
                    s.late_tolerance_minutes,
                    DATE(l.punch_timestamp) as p_date,
                    MIN(l.punch_timestamp) as check_in,
                    MAX(l.punch_timestamp) as check_out
                FROM attendance_raw_logs l
                LEFT JOIN employees e ON l.badge_or_card_id = e.badge_or_card_id
                LEFT JOIN work_schedules s ON e.schedule_id = s.id
                WHERE DATE(l.punch_timestamp) BETWEEN %s AND %s
                GROUP BY l.badge_or_card_id, DATE(l.punch_timestamp), e.first_name, e.last_name, s.start_time, s.end_time, s.late_tolerance_minutes
            """
            cursor.execute(query, (unknown_label, start_date, end_date))
            raw_logs = cursor.fetchall()

            emp_stats = {}
            for row in raw_logs:
                badge = row['badge_or_card_id']
                name = f"{row['first_name']} {row['last_name']}".strip()

                if badge not in emp_stats:
                    emp_stats[badge] = {
                        "name": name,
                        "present_days": 0,
                        "late_days": 0,
                        "absent_days": 0,
                        "total_seconds": 0
                    }

                cin = row['check_in']
                cout = row['check_out']
                sched_start = row['schedule_start'] or time(8, 0)
                tolerance = row['late_tolerance_minutes'] if row['late_tolerance_minutes'] is not None else 30
                p_date = row['p_date'] if isinstance(row['p_date'], date) else date.today()

                if not cin or not cout or cin == cout:
                    emp_stats[badge]["absent_days"] += 1
                else:
                    max_arrival_dt = datetime.combine(p_date, sched_start) + timedelta(minutes=tolerance)
                    is_arr_valid = cin <= max_arrival_dt

                    if is_arr_valid:
                        emp_stats[badge]["present_days"] += 1
                    else:
                        emp_stats[badge]["late_days"] += 1
                        emp_stats[badge]["present_days"] += 1

                    duration = cout - cin
                    if duration.total_seconds() > 0:
                        emp_stats[badge]["total_seconds"] += int(duration.total_seconds())

            for badge, data in emp_stats.items():
                total_secs = data["total_seconds"]
                hrs = total_secs // 3600
                mins = (total_secs % 3600) // 60
                formatted_hours = f"0 h {mins:02d} min" if hrs == 0 else f"{hrs}h {mins:02d} min"

                summary_rows.append({
                    "badge": badge,
                    "employee": data["name"],
                    "present_days": data["present_days"],
                    "late_days": data["late_days"],
                    "absent_days": data["absent_days"],
                    "worked_hours": formatted_hours
                })

        conn.close()
    except Exception as e:
        if conn:
            try:
                conn.rollback()
                conn.close()
            except Exception:
                pass
        print(f"Erreur lors de la génération du rapport de synthèse : {e}")

    summary_table_rows = ""
    if summary_rows:
        for row in summary_rows:
            summary_table_rows += f"""
                                <tr>
                                    <td class="text-start ps-4 fw-semibold">{row['employee']}</td>
                                    <td><span class="badge bg-success bg-opacity-10 text-success">{row['present_days']}</span></td>
                                    <td><span class="badge bg-danger bg-opacity-10 text-danger">{row['absent_days']}</span></td>
                                    <td><span class="badge bg-warning bg-opacity-20 text-dark">{row['late_days']}</span></td>
                                    <td class="pe-4 fw-bold text-primary">{row['worked_hours']}</td>
                                </tr>
            """
    else:
        summary_table_rows = f"""
                                <tr>
                                    <td colspan="5" class="text-center py-4 text-muted">{tr("sum_none", current_lang)}</td>
                                </tr>
        """

    export_summary_url = f"/summary/export/excel?start_date={start_date}&end_date={end_date}"

    html_content = f"""
    <!DOCTYPE html>
    <html lang="{current_lang}">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{tr("sum_page_title", current_lang)}</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
        <style>
            body {{
                background-color: #f4f7f6;
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            }}
            .navbar {{
                background: linear-gradient(135deg, #1e293b, #0f172a) !important;
            }}
            .card {{
                border: none;
                border-radius: 12px;
                box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
            }}
            .card-header {{
                background-color: #ffffff;
                border-bottom: 1px solid #f1f5f9;
                border-top-left-radius: 12px !important;
                border-top-right-radius: 12px !important;
                font-weight: 600;
            }}
            .table th {{
                font-weight: 600;
                color: #475569;
                background-color: #f8fafc !important;
            }}
            @media print {{
                @page {{
                    size: landscape;
                    margin: 5mm;
                }}
                .no-print {{
                    display: none !important;
                }}
                body {{
                    background-color: white !important;
                    font-size: 10pt;
                    width: 100% !important;
                    margin: 0 !important;
                }}
                .container-fluid {{
                    max-width: 100% !important;
                    width: 100% !important;
                    margin: 0 !important;
                    padding: 0 !important;
                }}
                .card {{
                    border: none !important;
                    box-shadow: none !important;
                    width: 100% !important;
                }}
                .card-body {{
                    padding: 0 !important;
                }}
                .table-responsive {{
                    overflow: visible !important;
                    width: 100% !important;
                }}
                .table {{
                    width: 100% !important;
                    max-width: 100% !important;
                }}
            }}
        </style>
    </head>
    <body>
        <nav class="navbar navbar-dark shadow-sm py-3 no-print">
            <div class="container-fluid px-4">
                <span class="navbar-brand mb-0 h5 fw-bold"><i class="fa-solid fa-chart-pie me-2"></i> {tr("sum_header", current_lang)}</span>
                <div class="d-flex align-items-center gap-2 flex-wrap">
                    <a href="/" class="btn btn-outline-light btn-sm px-3"><i class="fa-solid fa-house me-1"></i> {tr("nav_dashboard", current_lang)}</a>
                    <a href="/reports" class="btn btn-outline-light btn-sm px-3"><i class="fa-solid fa-calendar-days me-1"></i> {tr("nav_reports", current_lang)}</a>
                    <a href="/summary" class="btn btn-light btn-sm px-3 fw-semibold"><i class="fa-solid fa-chart-pie me-1"></i> {tr("nav_summary", current_lang)}</a>
                    {lang_switcher_html(current_lang)}
                    <a href="/logout" class="btn btn-danger btn-sm px-3 ms-1"><i class="fa-solid fa-right-from-bracket me-1"></i> {tr("nav_logout", current_lang)}</a>
                </div>
            </div>
        </nav>

        <div class="container-fluid px-4 my-4">
            <div class="card shadow-sm mb-4 no-print">
                <div class="card-header bg-primary text-white py-3">
                    <h5 class="mb-0 fs-6 fw-bold"><i class="fa-solid fa-filter me-2"></i> {tr("sum_filter_title", current_lang)}</h5>
                </div>
                <div class="card-body p-4">
                    <form method="GET" action="/summary" class="row g-3">
                        <div class="col-md-5">
                            <label class="form-label text-secondary small fw-bold">{tr("start_date", current_lang)}</label>
                            <input type="date" class="form-control" name="start_date" value="{start_date}" required>
                        </div>
                        <div class="col-md-5">
                            <label class="form-label text-secondary small fw-bold">{tr("end_date", current_lang)}</label>
                            <input type="date" class="form-control" name="end_date" value="{end_date}" required>
                        </div>
                        <div class="col-md-2 d-flex align-items-end">
                            <button type="submit" class="btn btn-success w-100 py-2 fw-semibold shadow-sm"><i class="fa-solid fa-magnifying-glass me-1"></i> {tr("filter_btn", current_lang)}</button>
                        </div>
                    </form>
                </div>
            </div>

            <div class="card shadow-sm">
                <div class="card-header d-flex justify-content-between align-items-center py-3 flex-wrap gap-2">
                    <h5 class="mb-0 fs-6 fw-bold text-dark"><i class="fa-solid fa-clipboard-user me-2 text-primary"></i> {tr("sum_report_title", current_lang)} ({len(summary_rows)} {tr("collaborators", current_lang)})</h5>
                    <div class="d-flex gap-2 no-print">
                        <a href="{export_summary_url}" class="btn btn-success btn-sm fw-semibold"><i class="fa-solid fa-file-excel me-1"></i> {tr("export_summary_excel", current_lang)}</a>
                        <button onclick="window.print();" class="btn btn-outline-dark btn-sm fw-semibold"><i class="fa-solid fa-print me-1"></i> {tr("print_pdf", current_lang)}</button>
                    </div>
                </div>
                <div class="card-body p-0">
                    <div class="table-responsive">
                        <table class="table table-striped table-hover mb-0 align-middle text-center">
                            <thead>
                                <tr>
                                    <th class="text-start ps-4">{tr("employee", current_lang)}</th>
                                    <th>{tr("sum_present_days", current_lang)}</th>
                                    <th>{tr("sum_absent_days", current_lang)}</th>
                                    <th>{tr("sum_late_days", current_lang)}</th>
                                    <th class="pe-4">{tr("sum_worked_hours", current_lang)}</th>
                                </tr>
                            </thead>
                            <tbody>
                                {summary_table_rows}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
    </body>
    </html>
    """
    return HTMLResponse(content=html_content)


# --- ROUTE D'EXPORTATION EXCEL DE LA SYNTHÈSE ---
@app.get("/summary/export/excel")
async def export_summary_excel(request: Request, start_date: str = None, end_date: str = None, drh_session: str = Cookie(None), lang: str = Cookie(None)):
    role = get_session_role(drh_session)
    if not role:
        return RedirectResponse(url="/login", status_code=303)
    if role == "supervisor":
        return RedirectResponse(url="/supervisor", status_code=303)

    current_lang = get_lang(request, lang)
    unknown_label = tr("unknown", current_lang)
    today_str = date.today().isoformat()
    start_date = start_date or today_str
    end_date = end_date or today_str

    conn = None
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            query = """
                SELECT
                    COALESCE(e.first_name, %s) as first_name,
                    COALESCE(e.last_name, '') as last_name,
                    l.badge_or_card_id,
                    s.start_time as schedule_start,
                    s.end_time as schedule_end,
                    s.late_tolerance_minutes,
                    DATE(l.punch_timestamp) as p_date,
                    MIN(l.punch_timestamp) as check_in,
                    MAX(l.punch_timestamp) as check_out
                FROM attendance_raw_logs l
                LEFT JOIN employees e ON l.badge_or_card_id = e.badge_or_card_id
                LEFT JOIN work_schedules s ON e.schedule_id = s.id
                WHERE DATE(l.punch_timestamp) BETWEEN %s AND %s
                GROUP BY l.badge_or_card_id, DATE(l.punch_timestamp), e.first_name, e.last_name, s.start_time, s.end_time, s.late_tolerance_minutes
            """
            cursor.execute(query, (unknown_label, start_date, end_date))
            raw_logs = cursor.fetchall()
        conn.close()

        emp_stats = {}
        for row in raw_logs:
            badge = row['badge_or_card_id']
            name = f"{row['first_name']} {row['last_name']}".strip()

            if badge not in emp_stats:
                emp_stats[badge] = {
                    "name": name,
                    "present_days": 0,
                    "late_days": 0,
                    "absent_days": 0,
                    "total_seconds": 0
                }

            cin = row['check_in']
            cout = row['check_out']
            sched_start = row['schedule_start'] or time(8, 0)
            tolerance = row['late_tolerance_minutes'] if row['late_tolerance_minutes'] is not None else 30
            p_date = row['p_date'] if isinstance(row['p_date'], date) else date.today()

            if not cin or not cout or cin == cout:
                emp_stats[badge]["absent_days"] += 1
            else:
                max_arr = datetime.combine(p_date, sched_start) + timedelta(minutes=tolerance)
                if cin <= max_arr:
                    emp_stats[badge]["present_days"] += 1
                else:
                    emp_stats[badge]["late_days"] += 1
                    emp_stats[badge]["present_days"] += 1

                duration = cout - cin
                if duration.total_seconds() > 0:
                    emp_stats[badge]["total_seconds"] += int(duration.total_seconds())

        export_summary = []
        for badge, data in emp_stats.items():
            hrs = data["total_seconds"] // 3600
            mins = (data["total_seconds"] % 3600) // 60
            export_summary.append({
                tr("excel_col_badge", current_lang): badge,
                tr("excel_col_emp", current_lang): data["name"],
                tr("excel_col_present", current_lang): data["present_days"],
                tr("excel_col_absent", current_lang): data["absent_days"],
                tr("excel_col_late", current_lang): data["late_days"],
                tr("excel_col_total_hours", current_lang): f"{hrs}h {mins:02d}m"
            })

        df = pd.DataFrame(export_summary)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name=tr("excel_sheet_sum", current_lang)[:31])
        output.seek(0)

        filename = f"Synthese_Presence_{start_date}_au_{end_date}.xlsx"
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )

    except Exception as e:
        if conn:
            try:
                conn.rollback()
                conn.close()
            except Exception:
                pass
        print(f"Erreur export Synthèse Excel : {e}")
        return RedirectResponse(url="/viewer-export" if role == "viewer" else "/summary", status_code=303)


@app.post("/employees/add")
async def add_employee(
    badge_or_card_id: str = Form(...),
    first_name: str = Form(...),
    last_name: str = Form(...),
    department: str = Form(None),
    schedule_id: int = Form(None),
    drh_session: str = Cookie(None)
):
    role = get_session_role(drh_session)
    if not role:
        return RedirectResponse(url="/login", status_code=303)
    if role == "supervisor":
        return RedirectResponse(url="/supervisor", status_code=303)
    if role == "viewer":
        return RedirectResponse(url="/viewer-export" if role == "viewer" else "/summary", status_code=303)

    conn = None
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO employees (badge_or_card_id, first_name, last_name, department, schedule_id, is_active)
                VALUES (%s, %s, %s, %s, %s, %s)
                ON CONFLICT (badge_or_card_id) DO UPDATE
                SET first_name = EXCLUDED.first_name,
                    last_name = EXCLUDED.last_name,
                    department = EXCLUDED.department,
                    schedule_id = EXCLUDED.schedule_id,
                    is_active = EXCLUDED.is_active
                """,
                (
                    badge_or_card_id.strip(),
                    first_name.strip(),
                    last_name.strip(),
                    department.strip() if department else None,
                    schedule_id,
                    True
                )
            )
            conn.commit()
        conn.close()
    except Exception as e:
        if conn:
            try:
                conn.rollback()
                conn.close()
            except Exception:
                pass
        print(f"Erreur lors de l'ajout de l'employé : {e}")

    return RedirectResponse(url="/", status_code=303)


@app.get("/employees/delete/{badge_id}")
async def delete_employee(badge_id: str, drh_session: str = Cookie(None)):
    role = get_session_role(drh_session)
    if not role:
        return RedirectResponse(url="/login", status_code=303)
    if role == "supervisor":
        return RedirectResponse(url="/supervisor", status_code=303)
    if role == "viewer":
        return RedirectResponse(url="/viewer-export" if role == "viewer" else "/summary", status_code=303)

    conn = None
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            cursor.execute("DELETE FROM attendance_raw_logs WHERE badge_or_card_id = %s", (badge_id,))
            cursor.execute("DELETE FROM employees WHERE badge_or_card_id = %s", (badge_id,))
            conn.commit()
        conn.close()
    except Exception as e:
        if conn:
            try:
                conn.rollback()
                conn.close()
            except Exception:
                pass
        print(f"Erreur lors de la suppression de l'employé : {e}")

    return RedirectResponse(url="/", status_code=303)


@app.get("/employees/edit/{badge_id}", response_class=HTMLResponse)
async def edit_employee_page(request: Request, badge_id: str, drh_session: str = Cookie(None), lang: str = Cookie(None)):
    role = get_session_role(drh_session)
    if not role:
        return RedirectResponse(url="/login", status_code=303)
    if role == "supervisor":
        return RedirectResponse(url="/supervisor", status_code=303)
    if role == "viewer":
        return RedirectResponse(url="/viewer-export" if role == "viewer" else "/summary", status_code=303)

    current_lang = get_lang(request, lang)
    employee = None
    schedules = []
    conn = None
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            cursor.execute("SELECT * FROM employees WHERE badge_or_card_id = %s", (badge_id,))
            employee = cursor.fetchone()

            cursor.execute("SELECT * FROM work_schedules ORDER BY id ASC")
            schedules = cursor.fetchall()
        conn.close()
    except Exception as e:
        if conn:
            try:
                conn.rollback()
                conn.close()
            except Exception:
                pass
        print(f"Erreur lors de la récupération de l'employé : {e}")

    if not employee:
        return RedirectResponse(url="/", status_code=303)

    sched_options_html = ""
    for s in schedules:
        sel = 'selected' if employee.get('schedule_id') == s['id'] else ''
        sched_options_html += f'<option value="{s["id"]}" {sel}>{s["name"]} ({str(s["start_time"])[:5]} - {str(s["end_time"])[:5]})</option>'

    html_content = f"""
    <!DOCTYPE html>
    <html lang="{current_lang}">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{tr("emp_edit_title", current_lang)}</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
        <style>
            body {{ background-color: #f4f7f6; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }}
            .navbar {{ background: linear-gradient(135deg, #1e293b, #0f172a) !important; }}
            .card {{ border: none; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); }}
        </style>
    </head>
    <body>
        <nav class="navbar navbar-dark shadow-sm py-3">
            <div class="container-fluid px-4">
                <span class="navbar-brand mb-0 h5 fw-bold"><i class="fa-solid fa-user-pen me-2"></i> {tr("emp_edit_header", current_lang)}</span>
                <div class="d-flex align-items-center gap-2">
                    {lang_switcher_html(current_lang)}
                    <a href="/" class="btn btn-outline-light btn-sm px-3"><i class="fa-solid fa-house me-1"></i> {tr("nav_dashboard", current_lang)}</a>
                </div>
            </div>
        </nav>

        <div class="container my-5" style="max-width: 600px;">
            <div class="card shadow-sm">
                <div class="card-header bg-primary text-white py-3">
                    <h5 class="mb-0 fs-6 fw-bold"><i class="fa-solid fa-user-gear me-2"></i> {tr("emp_edit_sub", current_lang)} {employee['first_name']} {employee['last_name']}</h5>
                </div>
                <div class="card-body p-4">
                    <form action="/employees/update/{badge_id}" method="POST">
                        <div class="mb-3">
                            <label class="form-label text-secondary small fw-bold">{tr("emp_badge", current_lang)}</label>
                            <input type="text" class="form-control" name="badge_or_card_id" value="{employee['badge_or_card_id']}" required>
                        </div>
                        <div class="mb-3">
                            <label class="form-label text-secondary small fw-bold">{tr("emp_firstname", current_lang)}</label>
                            <input type="text" class="form-control" name="first_name" value="{employee['first_name']}" required>
                        </div>
                        <div class="mb-3">
                            <label class="form-label text-secondary small fw-bold">{tr("emp_lastname", current_lang)}</label>
                            <input type="text" class="form-control" name="last_name" value="{employee['last_name']}" required>
                        </div>
                        <div class="mb-3">
                            <label class="form-label text-secondary small fw-bold">{tr("emp_dept", current_lang)}</label>
                            <input type="text" class="form-control" name="department" value="{employee.get('department') or ''}">
                        </div>
                        <div class="mb-3">
                            <label class="form-label text-secondary small fw-bold">{tr("emp_sched", current_lang)}</label>
                            <select class="form-select" name="schedule_id">
                                {sched_options_html}
                            </select>
                        </div>
                        <div class="mb-3 form-check">
                            <input type="checkbox" class="form-check-input" id="is_active" name="is_active" value="true" {'checked' if employee.get('is_active') else ''}>
                            <label class="form-check-label small fw-bold text-secondary" for="is_active">{tr("emp_active_chk", current_lang)}</label>
                        </div>
                        <div class="d-flex gap-2">
                            <button type="submit" class="btn btn-primary w-100 py-2 fw-semibold shadow-sm"><i class="fa-solid fa-check me-2"></i> {tr("save_btn", current_lang)}</button>
                            <a href="/" class="btn btn-secondary w-100 py-2 fw-semibold shadow-sm text-center text-decoration-none pt-2">{tr("cancel_btn", current_lang)}</a>
                        </div>
                    </form>
                </div>
            </div>
        </div>
    </body>
    </html>
    """
    return HTMLResponse(content=html_content)


@app.post("/employees/update/{old_badge_id}")
async def update_employee(
    old_badge_id: str,
    badge_or_card_id: str = Form(...),
    first_name: str = Form(...),
    last_name: str = Form(...),
    department: str = Form(None),
    schedule_id: int = Form(None),
    is_active: str = Form(None),
    drh_session: str = Cookie(None)
):
    role = get_session_role(drh_session)
    if not role:
        return RedirectResponse(url="/login", status_code=303)
    if role == "supervisor":
        return RedirectResponse(url="/supervisor", status_code=303)
    if role == "viewer":
        return RedirectResponse(url="/viewer-export" if role == "viewer" else "/summary", status_code=303)

    is_active_bool = True if is_active == "true" else False
    conn = None
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            cursor.execute(
                """
                UPDATE employees
                SET badge_or_card_id = %s, first_name = %s, last_name = %s, department = %s, schedule_id = %s, is_active = %s
                WHERE badge_or_card_id = %s
                """,
                (
                    badge_or_card_id.strip(),
                    first_name.strip(),
                    last_name.strip(),
                    department.strip() if department else None,
                    schedule_id,
                    is_active_bool,
                    old_badge_id
                )
            )
            conn.commit()
        conn.close()
    except Exception as e:
        if conn:
            try:
                conn.rollback()
                conn.close()
            except Exception:
                pass
        print(f"Erreur lors de la mise à jour de l'employé : {e}")

    return RedirectResponse(url="/", status_code=303)


@app.post("/api/v1/attendance-webhook")
async def receive_attendance(request: Request):
    conn = None
    try:
        body = await request.body()
        raw_text = body.decode("utf-8", errors="ignore")

        print("=== REQUÊTE HIKVISION REÇUE ===")
        print(raw_text)
        print("================================")

        employee_no = "UNKNOWN"
        device_id = None
        punch_type = "IN"

        query_dev = request.query_params.get("dev")
        if query_dev:
            device_id = f"Terminal_{query_dev.strip()}"

        # Extraire le bloc JSON même s'il est encapsulé dans multipart/form-data (--MIME_boundary)
        json_match = re.search(r'\{[\s\S]*\}', raw_text)
        json_str = json_match.group(0) if json_match else raw_text

        try:
            data = json.loads(json_str)

            def find_key(obj, key):
                if isinstance(obj, dict):
                    if key in obj and obj[key] is not None and str(obj[key]).strip() != "":
                        return str(obj[key]).strip()
                    for _, v in obj.items():
                        res = find_key(v, key)
                        if res:
                            return res
                elif isinstance(obj, list):
                    for item in obj:
                        res = find_key(item, key)
                        if res:
                            return res
                return None

            if not device_id:
                for key_candidate in ["serialNo", "serialNumber", "deviceID", "deviceSerialNo", "macAddress", "ipAddress", "deviceIpAddress"]:
                    val = find_key(data, key_candidate)
                    if val:
                        device_id = val
                        break

            # Priorité absolue à AccessControllerEvent.employeeNoString puis employeeNo
            acs_event = data.get("AccessControllerEvent") if isinstance(data, dict) else None
            if isinstance(acs_event, dict) and acs_event.get("employeeNoString") is not None and str(acs_event["employeeNoString"]).strip() != "":
                employee_no = str(acs_event["employeeNoString"]).strip()
            elif isinstance(acs_event, dict) and acs_event.get("employeeNo") is not None and str(acs_event["employeeNo"]).strip() != "":
                employee_no = str(acs_event["employeeNo"]).strip()
            else:
                for key_candidate in ["employeeNoString", "employeeNo"]:
                    val = find_key(data, key_candidate)
                    if val:
                        employee_no = val
                        break
        except Exception as e:
            print(f"Erreur parsing JSON Hikvision : {e}")

        if not device_id:
            for dev_pattern in [
                r'"(?:serialNo|serialNumber|deviceID|deviceSerialNo|macAddress)"\s*[:=]\s*["\']?([^"\'<>,\s]+)["\']?',
                r'<(?:serialNo|serialNumber|deviceID|macAddress)>([^<]+)</(?:serialNo|serialNumber|deviceID|macAddress)>'
            ]:
                m_dev = re.search(dev_pattern, raw_text)
                if m_dev:
                    device_id = m_dev.group(1).strip()
                    break

        if employee_no == "UNKNOWN":
            for pattern in [
                r'"(?:employeeNoString|employeeNo)"\s*[:=]\s*["\']?([^"\'<>,\s]+)["\']?',
                r'<employeeNoString>([^<]+)</employeeNoString>',
                r'<employeeNo>([^<]+)</employeeNo>'
            ]:
                match = re.search(pattern, raw_text)
                if match:
                    employee_no = match.group(1).strip()
                    break

        # Ignorer les identifiants non signés 64-bit / faux cardNo ou matricules invalides
        if (
            employee_no == "UNKNOWN"
            or not employee_no
            or len(employee_no) > 15
            or employee_no.startswith("18446744073609")
        ):
            return Response(content="<Response><statusString>IGNORED</statusString></Response>", media_type="application/xml", status_code=200)

        if not device_id:
            forwarded = request.headers.get("x-forwarded-for")
            if forwarded:
                client_host = forwarded.split(",")[0].strip()
            else:
                client_host = request.client.host if request.client else "UNKNOWN_IP"
            device_id = f"IP_{client_host}"

        print(f"-> Pointage validé | Matricule : {employee_no} | Appareil : {device_id}")

        local_punch_time = datetime.now(ZoneInfo("Africa/Kinshasa")).replace(tzinfo=None)

        conn = get_db_connection()
        with conn.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO devices_mapping (device_key, friendly_name)
                VALUES (%s, %s)
                ON CONFLICT (device_key) DO NOTHING
                """,
                (device_id, device_id)
            )
            cursor.execute(
                """
                INSERT INTO attendance_raw_logs (badge_or_card_id, punch_timestamp, device_id, punch_type)
                VALUES (%s, %s, %s, %s)
                """,
                (employee_no, local_punch_time, device_id, punch_type)
            )
            conn.commit()
        conn.close()

        return Response(content="<Response><statusString>OK</statusString></Response>", media_type="application/xml", status_code=200)

    except Exception as e:
        if conn:
            try:
                conn.rollback()
                conn.close()
            except Exception:
                pass
        print(f"Erreur critique webhook : {e}")
        return Response(content=f"<Response><statusString>ERROR: {str(e)}</statusString></Response>", media_type="application/xml", status_code=500)
