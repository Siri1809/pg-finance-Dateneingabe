import streamlit as st
import os
import requests
import pandas as pd
from PIL import Image
from io import BytesIO
from datetime import datetime, date

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

# === 🔑 HIER VERWALTEN SIE DIE MAKLER-PINS (Direkt im Code) ===
PIN_SIMON = "Siri0921"   # Ihre PIN (Simon)
PIN_CHEF = "legen_daer"    # Die PIN Ihres Chefs (Patrick)

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
    """Erstellt eine leere, perfekt formatierte kunden_pins.csv auf OneDrive."""
    try:
        check_url = f"https://graph.microsoft.com/v1.0/users/{USER_SYSTEM_NAME}/drive/root:/pg_finance_Dateneingabe/kunden_pins.csv"
        headers = {"Authorization": f"Bearer {token}"}
        check_res = requests.get(check_url, headers=headers)
        
        if check_res.status_code == 404: # Datei fehlt
            # Erstellt die CSV mit allen Spalten, um KeyErrors von Anfang an zu verhindern!
            csv_content = "PIN,Vorname,Nachname,Geburtsort,Familienstand,Kennzeichen,Fahrleistung,Kilometerstand,Garage,Weitere_Fahrer\n"
            upload_url = f"https://graph.microsoft.com/v1.0/users/{USER_SYSTEM_NAME}/drive/root:/pg_finance_Dateneingabe/kunden_pins.csv:/content"
            requests.put(upload_url, headers=headers, data=csv_content.encode("utf-8"))
    except Exception:
        pass

@st.cache_data(ttl=60) # Fragt OneDrive alle 60 Sekunden nach neuen Einträgen ab
def read_csv_from_onedrive(token):
    """Liest die kunden_pins.csv aus Simons OneDrive-Ordner."""
    try:
        url = f"https://graph.microsoft.com/v1.0/users/{USER_SYSTEM_NAME}/drive/root:/pg_finance_Dateneingabe/kunden_pins.csv:/content"
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            df = pd.read_csv(BytesIO(response.content))
            return df
        elif response.status_code == 404:
            create_initial_csv_on_onedrive(token)
    except Exception:
        pass
    return pd.DataFrame()

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
    neuer_dateiname = f"{name_teil}_{aktuelle_jahreszahl}{endung}"
    
    zaehler = 1
    while check_file_exists_on_onedrive(token, path_structure, neuer_dateiname):
        neuer_dateiname = f"{name_teil}_{aktuelle_jahreszahl}_{zaehler}{endung}"
        zaehler += 1
        
    url = f"https://graph.microsoft.com/v1.0/users/{USER_SYSTEM_NAME}/drive/root:/pg_finance_Dateneingabe/{path_structure}/{neuer_dateiname}:/content"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/octet-stream"
    }
    requests.put(url, headers=headers, data=file_content)

def save_database_to_onedrive(token, df):
    """Speichert die aktualisierte Pandas-Datenbank als CSV auf OneDrive."""
    try:
        url = f"https://graph.microsoft.com/v1.0/users/{USER_SYSTEM_NAME}/drive/root:/pg_finance_Dateneingabe/kunden_pins.csv:/content"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/octet-stream"
        }
        csv_buffer = BytesIO()
        df.to_csv(csv_buffer, index=False)
        requests.put(url, headers=headers, data=csv_buffer.getvalue())
        st.cache_data.clear() # Leert den Cache sofort
    except Exception:
        pass

# Token und Datenbank (DataFrame) laden
token = get_onedrive_token()
df_db = read_csv_from_onedrive(token) if token else pd.DataFrame()

# KORREKTE ERGÄNZUNG: Zwingt den DataFrame, alle Spalten zu besitzen, um KeyErrors auszuschließen!
erforderliche_spalten = ["PIN", "Vorname", "Nachname", "Geburtsort", "Familienstand", "Kennzeichen", "Fahrleistung", "Kilometerstand", "Garage", "Weitere_Fahrer"]
for col in erforderliche_spalten:
    if col not in df_db.columns:
        df_db[col] = ""

# Bereinigung der PINs im DataFrame, falls die Tabelle gelesen werden konnte
df_db["PIN"] = df_db["PIN"].astype(str).str.strip().str.replace(".0", "", regex=False)
df_db["PIN"] = df_db["PIN"].apply(lambda x: x.zfill(4) if (x.isdigit() and len(x) < 4) else x)

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
    .stTextInput input, .stNumberInput input, .stDateInput input {
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
    .stTextInput div[data-baseweb="input"], .stNumberInput div[data-baseweb="input"], .stDateInput div[data-baseweb="input"] {
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
    .stTextInput input:hover, .stNumberInput input:hover, .stSelectbox [data-baseweb="select"]:hover, .stDateInput input:hover {
        border-color: #0b4aa0 !important;
    }
    .stTextInput input:focus, .stNumberInput input:focus, .stSelectbox [data-baseweb="select"]:focus, .stDateInput input:focus {
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
    .makler-label {
        font-weight: bold;
        color: #718096;
        font-size: 0.85rem;
        text-transform: uppercase;
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

# KUNDEN_DATENBANK im Speicher füllen (Zieht sich die PINs aus dem DataFrame)
KUNDEN_DATENBANK = {}
if not df_db.empty:
    for _, r in df_db.iterrows():
        p_val = str(r["PIN"]).strip()
        KUNDEN_DATENBANK[p_val] = {
            "nachname": str(r["Nachname"]).strip(),
            "vorname": str(r["Vorname"]).strip()
        }

# Falls die Excel leer ist, erzwingen wir die Master-PINs als Fallback im Speicher
if PIN_SIMON not in KUNDEN_DATENBANK:
    KUNDEN_DATENBANK[PIN_SIMON] = {"nachname": "Seyschab", "vorname": "Simon"}
if PIN_CHEF not in KUNDEN_DATENBANK:
    KUNDEN_DATENBANK[PIN_CHEF] = {"nachname": "Grellner", "vorname": "Patrick"}

# Login-Kontrollen initialisieren
is_master_pin = str(pin_eingabe).strip() in [PIN_SIMON, PIN_CHEF]
is_valid_customer = str(pin_eingabe).strip() in KUNDEN_DATENBANK
weitere_fahrer_vorhanden = False # KORREKTE INITIALISIERUNG gegen NameErrors!

if is_master_pin or is_valid_customer:
    
    # =========================================================================
    # 📋 MODUS A: MAKLER-MODUS (NUR FÜR SIMON & PATRICK!)
    # =========================================================================
    if is_master_pin:
        st.success(f"🔓 Willkommen im Makler-Verwaltungs-Modus!")
        st.markdown("<h2 style='color:#0b4aa0;'>📋 Kundendaten-Manager</h2>", unsafe_allow_html=True)
        st.write("Wählen Sie einen Kunden aus, um alle eingetragenen Daten komfortabel zu kopieren.")
        
        # Filtert alle echten Kunden aus der Tabelle heraus
        kunden_df = df_db[~df_db["PIN"].isin([PIN_SIMON, PIN_CHEF])].copy()
        
        # Spalten-Fallback gegen KeyErrors
        for col in ["Nachname", "Vorname", "Kennzeichen"]:
            if col not in kunden_df.columns:
                kunden_df[col] = "Nicht angegeben"
        
        if not kunden_df.empty:
            kunden_df["Anzeige_Name"] = kunden_df["Nachname"].astype(str) + ", " + kunden_df["Vorname"].astype(str) + " (" + kunden_df["Kennzeichen"].astype(str) + ")"
            ausgewaehlter_kunden_name = st.selectbox("Kunde auswählen:", kunden_df["Anzeige_Name"].unique())
            
            if ausgewaehlter_kunden_name:
                nachname_sel = ausgewaehlter_kunden_name.split(", ")[0]
                vorname_sel = ausgewaehlter_kunden_name.split(", ")[1].split(" (")[0]
                kunden_daten = kunden_df[(kunden_df["Nachname"] == nachname_sel) & (kunden_df["Vorname"] == vorname_sel)]
                
                if len(kunden_daten) > 1:
                    auto_auswahl = st.selectbox("Fahrzeug wählen:", kunden_daten["Kennzeichen"].unique())
                    k_row = kunden_daten[kunden_daten["Kennzeichen"] == auto_auswahl].iloc[0]
                else:
                    k_row = kunden_daten.iloc[0]
                
                st.write("")
                st.markdown(f"### Stammdaten von {k_row['Vorname']} {k_row['Nachname']}")
                
                st.markdown("<span class='makler-label'>Nachname</span>", unsafe_allow_html=True)
                st.code(str(k_row["Nachname"]), language="text")
                
                st.markdown("<span class='makler-label'>Vorname</span>", unsafe_allow_html=True)
                st.code(str(k_row["Vorname"]), language="text")
                
                st.markdown("<span class='makler-label'>Geburtsort</span>", unsafe_allow_html=True)
                st.code(str(k_row.get("Geburtsort", "Nicht angegeben")), language="text")
                
                st.markdown("<span class='makler-label'>Familienstand</span>", unsafe_allow_html=True)
                st.code(str(k_row.get("Familienstand", "Nicht angegeben")), language="text")
                
                st.write("")
                st.markdown("### Fahrzeugdaten")
                
                st.markdown("<span class='makler-label'>Amtliches Kennzeichen</span>", unsafe_allow_html=True)
                st.code(str(k_row.get("Kennzeichen", "Nicht angegeben")), language="text")
                
                st.markdown("<span class='makler-label'>Jährliche Fahrleistung</span>", unsafe_allow_html=True)
                st.code(str(k_row.get("Fahrleistung", "Nicht angegeben")), language="text")
                
                st.markdown("<span class='makler-label'>Aktueller Kilometerstand</span>", unsafe_allow_html=True)
                st.code(str(k_row.get("Kilometerstand", "Nicht angegeben")), language="text")
                
                st.markdown("<span class='makler-label'>Abstellort (Garage)</span>", unsafe_allow_html=True)
                st.code(str(k_row.get("Garage", "Nicht angegeben")), language="text")
                
                st.write("")
                st.markdown("### Weitere Fahrer")
                st.markdown("<span class='makler-label'>Verzeichnete Fahrerdaten</span>", unsafe_allow_html=True)
                st.code(str(k_row.get("Weitere_Fahrer", "Keine weiteren Fahrer verzeichnet")), language="text")
            else:
                st.info("Es sind aktuell noch keine Kundendaten auf Ihrem OneDrive verzeichnet.")
        else:
            st.info("Es sind aktuell noch keine Kundendaten auf Ihrem OneDrive verzeichnet.")

    # =========================================================================
    # 🚗 MODUS B: DIE INTELLIGENTE EINGABEMASKE FÜR IHRE KUNDEN
    # =========================================================================
    else:
        k_info = KUNDEN_DATENBANK[pin_eingabe]
        kunden_nachname = k_info["nachname"]
        kunden_vorname = k_info["vorname"]
        
        # Intelligentes Voraussortieren aus OneDrive
        db_geburtsort = k_info.get("Geburtsort", "") if pd.notna(k_info.get("Geburtsort")) else ""
        db_familienstand = k_info.get("Familienstand", "Ledig") if pd.notna(k_info.get("Familienstand")) else "Ledig"
        db_garage = k_info.get("Garage", "Einzel-/Doppelgarage") if pd.notna(k_info.get("Garage")) else "Einzel-/Doppelgarage"
        db_fahrleistung = int(k_info.get("Fahrleistung", 10000)) if (pd.notna(k_info.get("Fahrleistung")) and str(k_info.get("Fahrleistung")).isdigit()) else 10000

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
            geburtsort = st.text_input("Geburtsort *", value=db_geburtsort)
            
            liste_familienstand = ["Ledig", "Verheiratet", "Eingetragene Lebenspartnerschaft", "Geschieden", "Verwitwet"]
            default_fam_index = liste_familienstand.index(db_familienstand) if db_familienstand in liste_familienstand else 0
            familienstand = st.selectbox("Familienstand", liste_familienstand, index=default_fam_index)
        st.markdown('</div>', unsafe_allow_html=True)

        # === KARTE 2: FAHRZEUG ===
        st.markdown("""
            <div class="form-card">
                <span class="card-header">🚘 Fahrzeug & Nutzung</span>
        """, unsafe_allow_html=True)
        
        col_fz1, col_col2 = st.columns(2)
        with col_fz1:
            fahrleistung = st.number_input("Jährliche Fahrleistung (in km) *", value=db_fahrleistung, step=1000)
            km_stand = st.number_input("Aktueller Kilometerstand", value=0, step=5000)
        with col_col2:
            liste_garage = ["Einzel-/Doppelgarage", "Tiefgarage", "Carport", "Privatgrundstück (befriedet)", "Straße / Laternenparker"]
            default_gar_index = liste_garage.index(db_garage) if db_garage in liste_garage else 0
            garage = st.selectbox("Abstellort des Fahrzeugs (Garage) *", liste_garage, index=default_gar_index)
        
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

        # === KARTE 3: WEITERE FAHRER ===
        st.markdown("""
            <div class="form-card">
                <span class="card-header">👥 Weitere Fahrer</span>
        """, unsafe_allow_html=True)
        
        weitere_fahrer_vorhanden = st.checkbox("Gibt es weitere berechtigte Fahrer für dieses Fahrzeug?", value=False)
        weitere_fahrer_liste = []
        
        if weitere_fahrer_vorhanden:
            if "anzahl_fahrer" not in st.session_state:
                st.session_state.anzahl_fahrer = 1
                
            col_plus1, col_plus2 = st.columns([3, 1])
            with col_plus1:
                st.write("Hier können Sie beliebig viele zusätzliche Fahrer für dieses Auto verzeichnen:")
            with col_plus2:
                if st.button("➕ Hinzufügen"):
                    st.session_state.anzahl_fahrer += 1
                    
            for i in range(st.session_state.anzahl_fahrer):
                st.markdown(f"#### 👤 Fahrer {i+1}")
                col_d1, col_d2 = st.columns(2)
                with col_d1:
                    wf_n = st.text_input(f"Nachname (Fahrer {i+1}) *", key=f"wf_n_{i}")
                    wf_v = st.text_input(f"Vorname (Fahrer {i+1}) *", key=f"wf_v_{i}")
                with col_d2:
                    wf_geb = st.date_input(f"Geburtsdatum (Fahrer {i+1}) *", min_value=date(1930, 1, 1), max_value=date.today(), value=date(1995, 1, 1), key=f"wf_geb_{i}")
                    wf_fs = st.date_input(f"Führerschein-Erwerb (Fahrer {i+1}) *", min_value=date(1948, 1, 1), max_value=date.today(), value=date(2013, 1, 1), key=f"wf_fs_{i}")
                
                if wf_n and wf_v:
                    weitere_fahrer_liste.append(f"Fahrer {i+1}: {wf_n}, {wf_v} (Geb: {wf_geb.strftime('%d.%m.%Y')}, FS seit: {wf_fs.strftime('%d.%m.%Y')})")
                st.markdown("---")
                
            if st.session_state.anzahl_fahrer > 1:
                if st.button("❌ Alle zusätzlichen Fahrer zurücksetzen"):
                    st.session_state.anzahl_fahrer = 1
                    st.rerun()
        else:
            st.session_state.anzahl_fahrer = 0
            
        st.markdown('</div>', unsafe_allow_html=True)

        # === KARTE 4: DOKUMENTE ===
        st.markdown("""
            <div class="form-card">
                <span class="card-header">📂 Dokumente hochladen</span>
        """, unsafe_allow_html=True)
        st.info("💡 Dokumente oder Fotos können Sie ganz einfach direkt mit Ihrer Smartphone-Kamera aufnehmen.")
        
        st.write("") 
        police = st.file_uploader("Letzte Versicherungspolice hier hochladen/fotografieren *", type=["pdf", "png", "jpg", "jpeg"])
        
        col_fs1, col_fs2 = st.columns(2)
        with col_fs1:
            fs_vorderseite = st.file_uploader("Führerschein Vorderseite (Optional)", type=["pdf", "png", "jpg", "jpeg"], key="fs_vorn")
        with col_fs2:
            fs_rueckseite = st.file_uploader("Führerschein Rückseite (Optional)", type=["pdf", "png", "jpg", "jpeg"], key="fs_hinten")
            
        col_id1, col_id2 = st.columns(2)
        with col_id1:
            ausweis_vorderseite = st.file_uploader("Ausweis Vorderseite / Reisepass (Optional)", type=["pdf", "png", "jpg", "jpeg"], key="ausweis_vorn")
        with col_id2:
            ausweis_rueckseite = st.file_uploader("Ausweis Rückseite (Optional)", type=["pdf", "png", "jpg", "jpeg"], key="ausweis_hinten")
        st.markdown('</div>', unsafe_allow_html=True)

        # === DATENSCHUTZ-EINWILLIGUNG ===
        st.markdown('<div class="form-card"><span class="card-header">⚖️ Datenschutz & Einwilligung</span>', unsafe_allow_html=True)
        datenschutz_akzeptiert = st.checkbox("Ich willige ein, dass meine Daten und Dokumente zum Zwecke der Angebotserstellung für die Kfz-Versicherung verarbeitet und gespeichert werden dürfen. Diese Einwilligung kann ich jederzeit widerrufen. *")
        st.markdown('</div>', unsafe_allow_html=True)

        # Pflichtfelder-Validierung
        pflicht_kennzeichen = neufahrzeug or (kennzeichen and kennzeichen != "")
        if weitere_fahrer_vorhanden:
            pflicht_wf = len(weitere_fahrer_liste) == st.session_state.anzahl_fahrer
        else:
            pflicht_wf = True
            
        pflichtfelder_ausgefuellt = (name and vorname and geburtsort and fahrzeugschein and police 
                                     and pflicht_kennzeichen and datenschutz_akzeptiert and pflicht_wf)

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

                wf_text_block = "Nein"
                if weitere_fahrer_liste:
                    wf_text_block = " | ".join(weitere_fahrer_liste)

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
----------------------------------------
Weitere Fahrer: {wf_text_block}
========================================"""

                with open(os.path.join(local_backup_path, "Kopier_Vorlage.txt"), "w", encoding="utf-8") as f:
                    f.write(infotext)

                # === ONEDRIVE HOCHLADEN-PROZESS ===
                if token:
                    with st.spinner("Daten werden sicher übertragen..."):
                        create_initial_csv_on_onedrive(token)
                        
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
                        
                        # === LERNENDER DATENBANK-PROZESS ===
                        if not df_db.empty:
                            match_auto = df_db[(df_db["PIN"] == str(pin_eingabe)) & (df_db["Kennzeichen"] == str(kennzeichen))]
                            
                            neue_zeile = {
                                "PIN": str(pin_eingabe),
                                "Vorname": str(vorname),
                                "Nachname": str(name),
                                "Geburtsort": str(geburtsort),
                                "Familienstand": str(familienstand),
                                "Kennzeichen": str(kennzeichen),
                                "Fahrleistung": int(fahrleistung),
                                "Kilometerstand": int(km_stand),
                                "Garage": str(garage),
                                "Weitere_Fahrer": str(wf_text_block)
                            }
                            
                            if not match_auto.empty:
                                idx = match_auto.index[0]
                                for col, val in neue_zeile.items():
                                    df_db.at[idx, col] = val
                            else:
                                df_db = pd.concat([df_db, pd.DataFrame([neue_zeile])], ignore_index=True)
                                
                            save_database_to_onedrive(token, df_db)
                        
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
