import streamlit as st
import os
import requests
import pandas as pd
from PIL import Image
from io import BytesIO
from datetime import datetime

# 1. Seiteneinstellungen (Auf Smartphones optimiert)
st.set_page_config(
    page_title="pg finance Dateneingabe", 
    page_icon="🚗", 
    layout="centered"
)

# === MICROSOFT ONEDRIVE CONFIG ===
USER_SYSTEM_NAME = "s.seyschab@pg-finance.de"
CLIENT_ID = st.secrets["microsoft"]["client_id"]
CLIENT_SECRET = st.secrets["microsoft"]["client_secret"]
TENANT_ID = st.secrets["microsoft"]["tenant_id"]

@st.cache_data(ttl=600) # RESSOURCEN-SCHUTZ: Cachet die Token-Abfrage für 10 Minuten (600 Sekunden)
def get_onedrive_token():
    """Authentifiziert die App bei Microsoft Graph API."""
    try:
        url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        data = {
            "client_id": CLIENT_ID,
            "scope": "https://graph.microsoft.com/.default",
            "client_secret": CLIENT_SECRET,
            "grant_type": "client_credentials"
        }
        response = requests.post(url, headers=headers, data=data)
        if response.status_code == 200:
            return response.json().get("access_token")
    except Exception:
        pass
    return None

def create_initial_csv_on_onedrive(token):
    """Erstellt eine leere kunden_pins.csv auf OneDrive, falls diese fehlt."""
    try:
        # Prüfen, ob die CSV bereits existiert
        check_url = f"https://graph.microsoft.com/v1.0/users/{USER_SYSTEM_NAME}/drive/root:/pg_finance_Dateneingabe/kunden_pins.csv"
        headers = {"Authorization": f"Bearer {token}"}
        check_res = requests.get(check_url, headers=headers)
        
        if check_res.status_code == 404: # Datei existiert noch nicht
            # Strukturierte CSV-Vorlage erstellen
            csv_data = "PIN,Vorname,Nachname,Status\n1234,Patrick,Grellner,Bereit\n0000,Simon,Seyschab,Bereit\n"
            
            # Hochladen in den Ordner pg_finance_Dateneingabe
            upload_url = f"https://graph.microsoft.com/v1.0/users/{USER_SYSTEM_NAME}/drive/root:/pg_finance_Dateneingabe/kunden_pins.csv:/content"
            requests.put(upload_url, headers=headers, data=csv_data.encode("utf-8"))
    except Exception:
        pass

@st.cache_data(ttl=300) # RESSOURCEN-SCHUTZ: Fragt OneDrive nur alle 5 Minuten nach PIN-Updates ab!
def read_csv_from_onedrive(token):
    """Liest die kunden_pins.csv aus Simons OneDrive-Ordner."""
    try:
        url = f"https://graph.microsoft.com/v1.0/users/{USER_SYSTEM_NAME}/drive/root:/pg_finance_Dateneingabe/kunden_pins.csv:/content"
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            # Liest die CSV-Datei ein
            df = pd.read_csv(BytesIO(response.content))
            datenbank = {}
            for _, row in df.iterrows():
                pin = str(row["PIN"]).strip()
                datenbank[pin] = {
                    "nachname": str(row["Nachname"]).strip(),
                    "vorname": str(row["Vorname"]).strip(),
                    "status": str(row.get("Status", "Bereit")).strip()
                }
            return datenbank
        elif response.status_code == 404:
            # Falls die CSV fehlt, erstellen wir sie automatisch!
            create_initial_excel_on_onedrive(token) # Lokales Fallback aufrufen
    except Exception:
        pass
    return {}

def check_file_exists_on_onedrive(token, path_structure, file_name):
    """Prüft, ob eine Datei bereits auf OneDrive existiert."""
    url = f"https://graph.microsoft.com/v1.0/users/{USER_SYSTEM_NAME}/drive/root:/pg_finance_Dateneingabe/{path_structure}/{file_name}"
    headers = {"Authorization": f"Bearer {token}"}
    res = requests.get(url, headers=headers)
    return res.status_code == 200

def upload_file_to_onedrive(token, path_structure, file_name, file_content):
    """Lädt eine Datei auf OneDrive hoch. Verhindert Überschreiben durch automatischen Zähler."""
    aktuelle_jahreszahl = datetime.now().strftime("%Y")
    name_teil, endung = os.path.splitext(file_name)
    
    # 1. Namen mit Jahreszahl versehen (z. B. Fahrzeugschein_2026.pdf)
    neuer_dateiname = f"{name_teil}_{aktuelle_jahreszahl}{endung}"
    
    # 2. Prüfen, ob die Datei mit diesem Jahr bereits existiert (Überschreibschutz!)
    zaehler = 1
    while check_file_exists_on_onedrive(token, path_structure, neuer_dateiname):
        neuer_dateiname = f"{name_teil}_{aktuelle_jahreszahl}_{zaehler}{endung}"
        zaehler += 1
        
    # 3. Datei mit dem endgültigen, sicheren Namen hochladen
    url = f"https://graph.microsoft.com/v1.0/users/{USER_SYSTEM_NAME}/drive/root:/pg_finance_Dateneingabe/{path_structure}/{neuer_dateiname}:/content"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/octet-stream"
    }
    requests.put(url, headers=headers, data=file_content)

def update_csv_status_on_onedrive(token, pin_to_update):
    """Markiert den Kunden in der kunden_pins.csv als 'Ausgefüllt'."""
    try:
        url = f"https://graph.microsoft.com/v1.0/users/{USER_SYSTEM_NAME}/drive/root:/pg_finance_Dateneingabe/kunden_pins.csv:/content"
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            df = pd.read_csv(BytesIO(response.content))
            df.loc[df["PIN"].astype(str) == str(pin_to_update), "Status"] = "Ausgefüllt"
            
            # CSV im Speicher neu schreiben
            csv_buffer = BytesIO()
            df.to_csv(csv_buffer, index=False)
            requests.put(url, headers=headers, data=csv_buffer.getvalue())
            
            # Cache leeren, damit die Änderung sofort aktiv wird
            st.cache_data.clear()
    except Exception:
        pass

# Token und Datenbank generieren
token = get_onedrive_token()
KUNDEN_DATENBANK = read_csv_from_onedrive(token) if token else {}

# Sicheres Fallback, falls die Verbindung mal komplett abreißt
if not KUNDEN_DATENBANK:
    KUNDEN_DATENBANK = {
        "0000": {"nachname": "Seyschab", "vorname": "Simon", "status": "Bereit"},
        "1234": {"nachname": "Grellner", "vorname": "Patrick", "status": "Bereit"}
    }

# CSS (Unverändert!)
st.markdown("""
    <style>
    .stApp {
        background-color: #ffffff !important;
        background-image: none !important;
        color: #2d3748 !important;
        font-family: 'Varela Round', 'Varela', sans-serif !important;
    }
    .main-title {
        font-family: 'Varela Round', sans-serif;
        color: #0b4aa0;
        font-weight: 700;
        font-size: 2.2rem;
        text-align: center;
        margin-top: 10px;
        margin-bottom: 5px;
    }
    .main-subtitle {
        text-align: center;
        color: #718096;
        font-size: 1.1rem;
        margin-bottom: 20px;
    }
    .form-card {
        background-color: #ffffff;
        padding: 24px;
        border-radius: 12px;
        border: 1px solid #edf2f7;
        margin-top: 0px !important;
        margin-bottom: 25px !important;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.02);
    }
    .card-header {
        color: #0b4aa0 !important;
        font-size: 1.35rem !important;
        font-weight: bold !important;
        font-family: 'Varela Round', sans-serif !important;
        margin-top: 0px !important;
        margin-bottom: 20px !important;
        border-bottom: 2px solid #00aeeb !important;
        padding-bottom: 8px !important;
        display: block !important;
    }
    .stTextInput input, .stNumberInput input {
        background-color: #ffffff !important;
        color: #0b4aa0 !important;
        border: 2px solid #00aeeb !important;
        border-radius: 6px !important;
        padding: 10px 14px !important;
        font-family: 'Varela Round', sans-serif !important;
        font-weight: bold !important;
        outline: none !important;
        box-shadow: none !important;
        transition: all 0.2s ease-in-out !important;
    }
    .stTextInput div[data-baseweb="input"], .stNumberInput div[data-baseweb="input"] {
        border: none !important;
        background-color: transparent !important;
        box-shadow: none !important;
    }
    .stSelectbox div[role="button"], 
    .stSelectbox div[data-baseweb="select"], 
    .stSelectbox [data-baseweb="select"] > div {
        border: none !important;
        background-color: transparent !important;
        box-shadow: none !important;
        outline: none !important;
    }
    .stSelectbox [data-baseweb="select"] {
        background-color: #ffffff !important;
        border: 2px solid #00aeeb !important;
        border-radius: 6px !important;
        font-family: 'Varela Round', sans-serif !important;
        font-weight: bold !important;
        box-shadow: none !important;
        outline: none !important;
        padding: 2px 4px !important;
    }
    .stSelectbox span, .stSelectbox div {
        color: #0b4aa0 !important;
        font-weight: bold !important;
    }
    .stTextInput input:hover, .stNumberInput input:hover, .stSelectbox [data-baseweb="select"]:hover {
        border-color: #0b4aa0 !important;
    }
    .stTextInput input:focus, .stNumberInput input:focus, .stSelectbox [data-baseweb="select"]:focus {
        border-color: #0b4aa0 !important;
        box-shadow: 0 0 0 3px rgba(11, 74, 160, 0.2) !important;
    }
    [data-testid="stFileUploaderDropzone"] {
        border: 2px dashed #00aeeb !important;
        background-color: #f7fafc !important;
        border-radius: 8px !important;
        padding: 15px !important;
        box-shadow: none !important;
        transition: all 0.3s ease !important;
    }
    .stFileUploader button {
        background-color: #00aeeb !important;
        color: #ffffff !important;
        border: none !important;
        font-weight: bold !important;
        font-family: 'Varela Round', sans-serif !important;
        font-size: 13px !important;
        padding: 6px 14px !important;
        border-radius: 6px !important;
    }
    .stAlert p, .stAlert span, .stAlert div {
        color: #1a1a1a !important;
        font-weight: bold !important;
    }
    label {
        color: #0b4aa0 !important;
        font-weight: bold !important;
        font-size: 0.95rem !important;
        margin-bottom: 6px !important;
    }
    div.stButton > button:first-child {
        background-color: #00aeeb !important;
        color: #ffffff !important;
        border: none !important;
        font-weight: 600 !important;
        font-family: 'Varela Round', sans-serif !important;
        border-radius: 6px !important;
        padding: 14px 40px !important;
        font-size: 16px !important;
        width: 100% !important;
        transition: all 0.3s ease;
        box-shadow: 0 4px 14px rgba(0, 174, 235, 0.2) !important;
        margin-bottom: 20px !important;
    }
    </style>
""", unsafe_allow_html=True)

LOGO_DATEINAME = "pg-finance_Logo.jpg"
if os.path.exists(LOGO_DATEINAME):
    logo = Image.open(LOGO_DATEINAME)
    col1, col2, col3 = st.columns(3)
    with col2:
        st.image(logo, use_container_width=True)

st.markdown("<h1 class='main-title'>Datenerfassung</h1>", unsafe_allow_html=True)
st.markdown("<p class='main-subtitle'>Schnell und sicher alle Daten für Ihre Autoversicherung einreichen</p>", unsafe_allow_html=True)

# 4. Sicherheits-Schranke: Die PIN-Eingabe (In der Karte verpackt)
st.markdown("""
        <div class="form-card">
            <span class="card-header">🔑 Zugang freischalten</span>
    """, unsafe_allow_html=True)
pin_eingabe = st.text_input("PIN-Eingabe", type="password", label_visibility="collapsed", placeholder="Bitte Ihre persönliche Kunden-PIN eingeben...")
st.markdown('</div>', unsafe_allow_html=True)

# PRÜFEN, OB DIE PIN IN DER DATENBANK EXISTIERT & BEREIT IST
if pin_eingabe in KUNDEN_DATENBANK and KUNDEN_DATENBANK[pin_eingabe].get("status", "Bereit") == "Bereit":
    kunden_info = KUNDEN_DATENBANK[pin_eingabe]
    kunden_nachname = kunden_info["nachname"]
    kunden_vorname = kunden_info["vorname"]

    st.success(f"🔓 Willkommen {kunden_vorname} {kunden_nachname}! Ihr Formular wurde freigeschaltet.")

    # === KARTE 1: PERSÖNLICHE DATEN ===
    st.markdown("""
        <div class="form-card">
            <span class="card-header">👤 Persönliche Daten</span>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        name = st.text_input("Nachname *", value=kunden_nachname, disabled=True)
        vorname = st.text_input("Vorname *", value=kunden_vorname, disabled=True)
    with col2:
        geburtsort = st.text_input("Geburtsort *")
        familienstand = st.selectbox("Familienstand", ["Ledig", "Verheiratet", "Eingetragene Lebenspartnerschaft", "Geschieden", "Verwitwet"])
    st.markdown('</div>', unsafe_allow_html=True)

    # === KARTE 2: FAHRZEUG ===
    st.markdown("""
        <div class="form-card">
            <span class="card-header">🚘 Fahrzeug & Nutzung</span>
    """, unsafe_allow_html=True)
    
    col_fz1, col_col2 = st.columns(2)
    with col_fz1:
        fahrleistung = st.number_input("Jährliche Fahrleistung (in km) *", value=10000, step=1000)
        km_stand = st.number_input("Aktueller Kilometerstand", value=0, step=5000)
    with col_col2:
        garage = st.selectbox("Abstellort des Fahrzeugs (Garage) *", ["Einzel-/Doppelgarage", "Tiefgarage", "Carport", "Privatgrundstück (befriedet)", "Straße / Laternenparker"])
    
    st.write("") 
    col_kz1, col_kz2 = st.columns([2, 1])
    with col_kz2:
        neufahrzeug = st.checkbox("Neufahrzeug", help="Aktivieren Sie dies, falls das Fahrzeug noch nicht zugelassen ist.")
    with col_kz1:
        if neufahrzeug:
            kennzeichen = st.text_input("Amtliches Kennzeichen", value="NEUFAHRZEUG", disabled=True)
        else:
            kennzeichen = st.text_input("Amtliches Kennzeichen *", placeholder="z.B. BA-PG-99")

    st.write("")
    fahrzeugschein = st.file_uploader("Fahrzeugschein hier hochladen/fotografieren *", type=["pdf", "png", "jpg", "jpeg"])
    st.markdown('</div>', unsafe_allow_html=True)

    # === KARTE 3: DOKUMENTE ===
    st.markdown("""
        <div class="form-card">
            <span class="card-header">📂 Dokumente hochladen</span>
    """, unsafe_allow_html=True)
    st.info("💡 Dokumente oder Fotos können Sie ganz einfach direkt mit Ihrer Smartphone-Kamera aufnehmen.")
    
    st.write("") 
    police = st.file_uploader("Letzte Versicherungspolice hier hochladen/fotografieren *", type=["pdf", "png", "jpg", "jpeg"])
    
    # Führerscheine (OPTIONAL)
    st.markdown("<p style='color: #0b4aa0; font-weight: bold; margin-bottom: 2px;'>Führerschein aller berechtigten Fahrer (Optional)</p>", unsafe_allow_html=True)
    col_fs1, col_fs2 = st.columns(2)
    with col_fs1:
        fs_vorderseite = st.file_uploader("Vorderseite", type=["pdf", "png", "jpg", "jpeg"], key="fs_vorn")
    with col_fs2:
        fs_rueckseite = st.file_uploader("Rückseite", type=["pdf", "png", "jpg", "jpeg"], key="fs_hinten")
        
    # Ausweise (OPTIONAL)
    st.markdown("<p style='color: #0b4aa0; font-weight: bold; margin-top: 15px; margin-bottom: 2px;'>Ausweisdokument aller berechtigten Fahrer (Optional)</p>", unsafe_allow_html=True)
    col_id1, col_id2 = st.columns(2)
    with col_id1:
        ausweis_vorderseite = st.file_uploader("Vorderseite / Reisepass-Hauptseite", type=["pdf", "png", "jpg", "jpeg"], key="ausweis_vorn")
    with col_id2:
        ausweis_rueckseite = st.file_uploader("Rückseite", type=["pdf", "png", "jpg", "jpeg"], key="ausweis_hinten")
    st.markdown('</div>', unsafe_allow_html=True)

    # === DATENSCHUTZ-EINWILLIGUNG ===
    st.markdown('<div class="form-card"><span class="card-header">⚖️ Datenschutz & Einwilligung</span>', unsafe_allow_html=True)
    datenschutz_akzeptiert = st.checkbox("Ich willige ein, dass meine Daten und Dokumente zum Zwecke der Angebotserstellung für die Kfz-Versicherung verarbeitet und gespeichert werden dürfen. Diese Einwilligung kann ich jederzeit widerrufen. *")
    st.markdown('</div>', unsafe_allow_html=True)

    # Pflichtfelder-Validierung
    pflicht_kennzeichen = neufahrzeug or (kennzeichen and kennzeichen != "")
    pflichtfelder_ausgefuellt = (name and vorname and geburtsort and fahrzeugschein and police 
                                 and pflicht_kennzeichen and datenschutz_akzeptiert)

    if pflichtfelder_ausgefuellt:
        if st.button("DATEN JETZT SICHER ÜBERTRAGEN", type="primary"):
            
            # Ordnerstruktur (Nachname_Vorname / Kennzeichen)
            safe_kz = kennzeichen.replace(" ", "-").replace("/", "-")
            kunden_ordner_name = f"{name}_{vorname}"
            vollstaendiger_upload_pfad = f"{kunden_ordner_name}/{safe_kz}"
            
            # Lokales Backup erstellen
            local_backup_path = os.path.join(kunden_ordner_name, safe_kz)
            if not os.path.exists(local_backup_path):
                os.makedirs(local_backup_path)

            infotext = f"""=== KUNDENDATEN FÜR NAFI / COMPARIT ===
Name: {name}
Vorname: {vorname}
Geburtsort: {geburtsort}
Familienstand: {familienstand}
----------------------------------------
Kennzeichen: {kennzeichen}
Fahrleistung: {fahrleistung} km/Jahr
Aktueller KM-Stand: {km_stand} km
Garage: {garage}
========================================"""

            with open(os.path.join(local_backup_path, "Kopier_Vorlage.txt"), "w", encoding="utf-8") as f:
                f.write(infotext)

            # === ONEDRIVE HOCHLADEN-PROZESS ===
            if token:
                with st.spinner("Daten werden sicher übertragen..."):
                    # 1. ERSTELLT DIE CSV-DATEI AUTOMATISCH, FALLS SIE NOCH FEHLT!
                    create_initial_csv_on_onedrive(token)
                    
                    # 2. Lädt die Kundendaten hoch (Mit automatischem Überschreibschutz!)
                    upload_file_to_onedrive(token, vollstaendiger_upload_pfad, "Kopier_Vorlage.txt", infotext.encode("utf-8"))
                    
                    if fahrzeugschein:
                        upload_file_to_onedrive(token, vollstaendiger_upload_pfad, f"Fahrzeugschein_{fahrzeugschein.name}", fahrzeugschein.getbuffer())
                    if police:
                        upload_file_to_onedrive(token, vollstaendiger_upload_pfad, f"Police_{police.name}", police.getbuffer())
                    if fs_vorderseite:
                        upload_file_to_onedrive(token, vollstaendiger_upload_pfad, f"FS_Vorderseite_{fs_vorderseite.name}", fs_vorderseite.getbuffer())
                    if fs_rueckseite:
                        upload_file_to_onedrive(token, vollstaendiger_upload_pfad, f"FS_Rueckseite_{fs_rueckseite.name}", fs_rueckseite.getbuffer())
                    if ausweis_vorderseite:
                        upload_file_to_onedrive(token, vollstaendiger_upload_pfad, f"Ausweis_Vorderseite_{ausweis_vorderseite.name}", ausweis_vorderseite.getbuffer())
                    if ausweis_rueckseite:
                        upload_file_to_onedrive(token, vollstaendiger_upload_pfad, f"Ausweis_Rueckseite_{ausweis_rueckseite.name}", ausweis_rueckseite.getbuffer())
                    
                    # 3. Status in OneDrive CSV markieren
                    update_csv_status_on_onedrive(token, pin_eingabe)
                    
                st.balloons()
                st.success("🎉 Übertragung erfolgreich! Vielen Dank für Ihre Mühe. Sie können das Browserfenster jetzt schließen.")
            else:
                st.warning("⚠️ OneDrive Verbindung offline. Daten wurden nur lokal gesichert.")

    else:
        st.warning("⚠️ Bitte füllen Sie alle mit * markierten Felder aus, laden Sie die Pflichtdokumente hoch und bestätigen Sie die Datenschutzerklärung, um die Übertragung zu starten.")

    # 6. STIMMUNGSBILD GANZ UNTEN
    THEME_BILD = "pg-finance-theme.jpg"
    if os.path.exists(THEME_BILD):
        theme_img = Image.open(THEME_BILD)
        st.image(theme_img, use_container_width=True)

else:
    if pin_eingabe != "":
        st.error("❌ Falsche PIN oder PIN bereits abgelaufen. Bitte prüfen Sie Ihre Eingabe.")
