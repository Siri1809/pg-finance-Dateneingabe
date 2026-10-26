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

# === 🔑 MAKLER-PINS (Sicher im Code verankert) ===
PIN_SIMON = "Siri0921"
PIN_CHEF = "legen_daer"

# =========================================================================
# 🛠️ ALLE HILFS-FUNKTIONEN (Müssen ganz oben stehen, damit Python sie kennt!)
# =========================================================================

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
            base_cols = "PIN,Vorname,Nachname,Weitere_Vornamen,Geburtsort,Familienstand,VN_Status,Wohneigentum,Kennzeichen,Fahrleistung,Kilometerstand,Garage,Finanzierung"
            for i in range(1, 6):
                base_cols += f",Fahrer_{i}_Nachname,Fahrer_{i}_Vorname,Fahrer_{i}_Geburtsdatum,Fahrer_{i}_Erwerbsdatum"
            base_cols += "\n"
            
            upload_url = f"https://graph.microsoft.com/v1.0/users/{USER_SYSTEM_NAME}/drive/root:/pg_finance_Dateneingabe/kunden_pins.csv:/content"
            requests.put(upload_url, headers=headers, data=base_cols.encode("utf-8"))
    except Exception:
        pass

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

def create_onedrive_folder(token, folder_name):
    """Erstellt einen leeren Kundenordner direkt auf OneDrive."""
    try:
        url = f"https://graph.microsoft.com/v1.0/users/{USER_SYSTEM_NAME}/drive/root:/pg_finance_Dateneingabe/{folder_name}"
        headers = {"Authorization": f"Bearer {token}"}
        res = requests.get(url, headers=headers)
        if res.status_code == 404: # Ordner existiert noch nicht
            create_url = f"https://graph.microsoft.com/v1.0/users/{USER_SYSTEM_NAME}/drive/root:/pg_finance_Dateneingabe:/children"
            create_headers = {
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json"
            }
            body = {
                "name": folder_name,
                "folder": {},
                "@microsoft.graph.conflictBehavior": "replace"
            }
            requests.post(create_url, headers=create_headers, json=body)
    except Exception:
        pass

def upload_file_to_onedrive(token, path_structure, file_name, file_content):
    """Lädt eine Datei auf OneDrive hoch. Verhindert Überschreiben durch automatischen Zähler."""
    aktuelle_jahreszahl = datetime.now().strftime("%Y")
    name_teil, endung = os.path.splitext(file_name)
    neuer_dateiname = f"{name_teil}_{aktuelle_jahreszahl}{endung}"
    
    zaehler = 1
    while check_file_exists_on_onedrive(token, path_structure, neuer_dateiname):
        neuer_dateiname = f"{name_teil}_{aktuelle_jahreszahl}_{zaehler}{endung}"
        zaehler += 1
        
    url = f"https://graph.microsoft.com/v1.0/users/{USER_SYSTEM_NAME}/drive/root:/pg_finance_Dateneingabe/{path_structure}/{neuer_dateiname}:/content?@microsoft.graph.conflictBehavior=replace"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/octet-stream"
    }
    requests.put(url, headers=headers, data=file_content)

def save_database_to_onedrive(token, df):
    """Speichert die aktualisierte Pandas-Datenbank als CSV auf OneDrive."""
    try:
        url = f"https://graph.microsoft.com/v1.0/users/{USER_SYSTEM_NAME}/drive/root:/pg_finance_Dateneingabe/kunden_pins.csv:/content?@microsoft.graph.conflictBehavior=replace"
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

# =========================================================================
# ⚙️ SYSTEM-ABFRAGEN UND START DER APPLIKATION
# =========================================================================

# Token und Datenbank (DataFrame) laden
token = get_onedrive_token()
df_db = read_csv_from_onedrive(token) if token else pd.DataFrame()

# Zwingt den DataFrame, alle Spalten im Speicher zu besitzen
erforderliche_spalten = ["PIN", "Vorname", "Nachname", "Weitere_Vornamen", "Geburtsort", "Familienstand", "VN_Status", "Wohneigentum", "Kennzeichen", "Fahrleistung", "Kilometerstand", "Garage", "Finanzierung"]
for i in range(1, 6):
    erforderliche_spalten += [f"Fahrer_{i}_Nachname", f"Fahrer_{i}_Vorname", f"Fahrer_{i}_Geburtsdatum", f"Fahrer_{i}_Erwerbsdatum"]

for col in erforderliche_spalten:
    if col not in df_db.columns:
        df_db[col] = ""

# Bereinigung der PIN-Formate (Löst Geisterzeilen & Zahlen-Fehler)
if not df_db.empty:
    df_db = df_db.dropna(subset=["PIN"])
    df_db = df_db[df_db["PIN"].astype(str).str.strip() != ""]
    
    def clean_pin_val(x):
        val = str(x).strip()
        if val.endswith(".0"):
            val = val[:-2]
        if val.isdigit() and len(val) < 4:
            val = val.zfill(4)
        return val

    df_db["PIN"] = df_db["PIN"].apply(clean_pin_val)

# KUNDEN_DATENBANK im Speicher befüllen
KUNDEN_DATENBANK = {}
if not df_db.empty:
    for _, r in df_db.iterrows():
        p_val = str(r["PIN"]).strip()
        KUNDEN_DATENBANK[p_val] = {
            "nachname": str(r["Nachname"]).strip() if pd.notna(r["Nachname"]) else "",
            "vorname": str(r["Vorname"]).strip() if pd.notna(r["Vorname"]) else "",
            "weitere_vornamen": r.get("Weitere_Vornamen", "") if pd.notna(r.get("Weitere_Vornamen")) else "",
            "Geburtsort": r.get("Geburtsort", "") if pd.notna(r.get("Geburtsort")) else "",
            "Familienstand": r.get("Familienstand", "Ledig") if pd.notna(r.get("Familienstand")) else "Ledig",
            "VN_Status": r.get("VN_Status", "Angestellter") if pd.notna(r.get("VN_Status")) else "Angestellter",
            "Wohneigentum": r.get("Wohneigentum", "Nein, zur Miete") if pd.notna(r.get("Wohneigentum")) else "Nein, zur Miete",
            "Garage": r.get("Garage", "Einzel-/Doppelgarage") if pd.notna(r.get("Garage")) else "Einzel-/Doppelgarage",
            "Fahrleistung": r.get("Fahrleistung", 10000) if pd.notna(r.get("Fahrleistung")) else 10000,
            "Kennzeichen": r.get("Kennzeichen", "") if pd.notna(r.get("Kennzeichen")) else ""
        }

# Ergänzt die Master-PINs im Speicher-Fallback
if PIN_SIMON not in KUNDEN_DATENBANK:
    KUNDEN_DATENBANK[PIN_SIMON] = {"nachname": "Seyschab", "vorname": "Simon"}
if PIN_CHEF not in KUNDEN_DATENBANK:
    KUNDEN_DATENBANK[PIN_CHEF] = {"nachname": "Grellner", "vorname": "Patrick"}

# =========================================================================
# 🏠 STREAMLIT OBERFLÄCHE (VISUELL)
# =========================================================================

# Logo & Überschrift laden (Immer sichtbar)
LOGO_DATEINAME = "pg-finance_Logo.jpg"
if os.path.exists(LOGO_DATEINAME):
    logo = Image.open(LOGO_DATEINAME)
    col1, col2, col3 = st.columns(3)
    with col2:
        st.image(logo, use_container_width=True)

st.markdown("<h1 class='main-title'>Datenerfassung</h1>", unsafe_allow_html=True)
st.markdown("<p class='main-subtitle'>Schnell und sicher alle Daten für Ihre Autoversicherung einreichen</p>", unsafe_allow_html=True)

# PIN-Eingabemaske
st.markdown("""
        <div class="form-card">
            <span class="card-header">🔑 Zugang freischalten</span>
    """, unsafe_allow_html=True)
pin_eingabe = st.text_input("PIN-Eingabe", type="password", key="pin_eingabe_field", label_visibility="collapsed", placeholder="Bitte Ihre persönliche Kunden-PIN eingeben...")
st.markdown('</div>', unsafe_allow_html=True)

# Bereinigung der Benutzereingabe
tipp_pin = str(pin_eingabe).strip()
if tipp_pin.isdigit() and len(tipp_pin) < 4:
    tipp_pin = tipp_pin.zfill(4)

is_master_pin = tipp_pin in [PIN_SIMON, PIN_CHEF]
is_valid_customer = tipp_pin in KUNDEN_DATENBANK
weitere_fahrer_vorhanden = False # NameError Schutz

# PRÜFUNG DES LOGINS
if is_master_pin or is_valid_customer:
    
    # =========================================================================
    # 📋 MODUS A: MAKLER-MODUS (NUR FÜR SIMON & PATRICK!)
    # =========================================================================
    if is_master_pin:
        st.success(f"🔓 Willkommen im Makler-Verwaltungs-Modus!")
        makler_reiter = st.radio("Aktion wählen:", ["➕ Neuen Kunden-Zugang anlegen", "📋 Kundendaten kopieren"])
        
        # --- REITER 1: KUNDEN MOBIL ANLEGEN ---
        if makler_reiter == "➕ Neuen Kunden-Zugang anlegen":
            st.markdown("<h3 style='color:#0b4aa0;'>Neuen Kunden-Zugang anlegen</h3>", unsafe_allow_html=True)
            st.write("Erstellen Sie hier blitzschnell einen Zugang für einen neuen Kunden:")
            
            with st.form("kunden_anlegen_form", clear_on_submit=True):
                new_n = st.text_input("Nachname des Kunden *")
                new_v = st.text_input("Vorname des Kunden *")
                new_pin = st.text_input("Gewünschte PIN (z.B. PG-123) *")
                
                if st.form_submit_button("🚀 Kunden-Zugang erstellen"):
                    if new_n and new_v and new_pin:
                        safe_new_pin = str(new_pin).strip()
                        
                        neuer_k_row = {col: "" for col in df_db.columns}
                        neuer_k_row["PIN"] = safe_new_pin
                        neuer_k_row["Nachname"] = str(new_n).strip()
                        neuer_k_row["Vorname"] = str(new_v).strip()
                        
                        df_db = pd.concat([df_db, pd.DataFrame([neuer_k_row])], ignore_index=True)
                        save_database_to_onedrive(token, df_db)
                        
                        k_ordner = f"{str(new_n).strip()}_{str(new_v).strip()}"
                        create_onedrive_folder(token, k_ordner)
                        
                        st.success(f"🎉 Zugang für {new_v} {new_n} erfolgreich auf OneDrive erstellt!")
                        st.balloons()
                        
                        st.write("---")
                        st.write("📋 **WhatsApp-Nachricht zum Kopieren:**")
                        wa_text = f"Hallo {new_v},\n\nhier ist dein persönlicher Link zur schnellen Datenerfassung für deine Autoversicherung:\n\n🔗 https://kfz-tarif-helfer.streamlit.app\n\n🔑 Deine persönliche PIN lautet: {safe_new_pin}"
                        st.code(wa_text, language="text")
                    else:
                        st.error("Bitte alle Felder ausfüllen!")
                        
        # --- REITER 2: KUNDENDATEN KOPIEREN ---
        else:
            st.markdown("<h3 style='color:#0b4aa0;'>📋 Kundendaten-Manager</h3>", unsafe_allow_html=True)
            st.write("Wählen Sie einen Kunden aus, um alle eingetragenen Daten komfortabel zu kopieren.")
            
            kunden_df = df_db[~df_db["PIN"].isin([PIN_SIMON, PIN_CHEF])].copy()
            
            if not kunden_df.empty:
                kunden_df["Nachname"] = kunden_df["Nachname"].fillna("Unbekannt").astype(str).str.strip()
                kunden_df["Vorname"] = kunden_df["Vorname"].fillna("Unbekannt").astype(str).str.strip()
                kunden_df["Kennzeichen"] = kunden_df["Kennzeichen"].fillna("").astype(str).str.strip()
                
                kunden_df["KZ_Anzeige"] = kunden_df["Kennzeichen"].apply(lambda x: x if x != "" else "Noch kein Auto eingereicht")
                kunden_df["Anzeige_Name"] = kunden_df["Nachname"] + ", " + kunden_df["Vorname"] + " (" + kunden_df["KZ_Anzeige"] + ")"
                
                dropdown_optionen = ["-- Bitte wählen Sie einen Kunden aus --"] + list(kunden_df["Anzeige_Name"].unique())
                ausgewaehlter_kunden_name = st.selectbox("Kunde auswählen:", dropdown_optionen, index=0)
                
                if ausgewaehlter_kunden_name and ausgewaehlter_kunden_name != "-- Bitte wählen Sie einen Kunden aus --":
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
                    
                    st.markdown("<span class='makler-label'>Weitere Vornamen</span>", unsafe_allow_html=True)
                    st.code(str(k_row.get("Weitere_Vornamen", "Keine weiteren Vornamen verzeichnet")), language="text")
                    
                    st.markdown("<span class='makler-label'>Geburtsort</span>", unsafe_allow_html=True)
                    st.code(str(k_row.get("Geburtsort", "Nicht angegeben")), language="text")
                    
                    st.markdown("<span class='makler-label'>Familienstand</span>", unsafe_allow_html=True)
                    st.code(str(k_row.get("Familienstand", "Nicht angegeben")), language="text")
                    
                    st.markdown("<span class='makler-label'>Status des Versicherungsnehmers (Beruf)</span>", unsafe_allow_html=True)
                    st.code(str(k_row.get("VN_Status", "Nicht angegeben")), language="text")
                    
                    st.markdown("<span class='makler-label'>Wohneigentum vorhanden</span>", unsafe_allow_html=True)
                    st.code(str(k_row.get("Wohneigentum", "Nicht angegeben")), language="text")
                    
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
                    
                    st.markdown("<span class='makler-label'>Finanzierungsart des Fahrzeugs</span>", unsafe_allow_html=True)
                    st.code(str(k_row.get("Finanzierung", "Nicht angegeben")), language="text")
                    
                    st.write("")
                    st.markdown("### Weitere Fahrer")
                    for j in range(1, 6):
                        f_n = k_row.get(f"Fahrer_{j}_Nachname")
                        f_v = k_row.get(f"Fahrer_{j}_Vorname")
                        if pd.notna(f_n) and str(f_n).strip() != "":
                            st.markdown(f"<span class='makler-label'>Fahrer {j} (Name / Geburtsdatum / Führerschein)</span>", unsafe_allow_html=True)
                            f_geb_raw = k_row.get(f"Fahrer_{j}_Geburtsdatum", "")
                            f_fs_raw = k_row.get(f"Fahrer_{j}_Erwerbsdatum", "")
                            st.code(f"{f_v} {f_n} (Geb: {f_geb_raw}, FS: {f_fs_raw})", language="text")
                else:
                    st.info("Bitte wählen Sie einen Kunden aus der Liste aus, um seine Daten einzusehen.")
            else:
                st.info("Es sind aktuell noch keine Kundendaten auf Ihrem OneDrive verzeichnet.")

    # =========================================================================
    # 🚗 MODUS B: DIE INTELLIGENTE EINGABEMASKE FÜR IHRE KUNDEN
    # =========================================================================
    else:
        k_info = KUNDEN_DATENBANK[tipp_pin]
        kunden_nachname = k_info["nachname"]
        kunden_vorname = k_info["vorname"]
        
        db_vornamen = k_info.get("weitere_vornamen", "") if pd.notna(k_info.get("weitere_vornamen")) else ""
        db_geburtsort = k_info.get("Geburtsort", "") if pd.notna(k_info.get("Geburtsort")) else ""
        db_familienstand = k_info.get("Familienstand", "Ledig") if pd.notna(k_info.get("Familienstand")) else "Ledig"
        db_vn_status = k_info.get("VN_Status", "Angestellter") if pd.notna(k_info.get("VN_Status")) else "Angestellter"
        db_wohneigentum = k_info.get("Wohneigentum", "Nein, zur Miete") if pd.notna(k_info.get("Wohneigentum")) else "Nein, zur Miete"
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
            weitere_vornamen = st.text_input("Weitere Vornamen (falls im Führerschein verzeichnet)", value=db_vornamen)
        with col2:
            geburtsort = st.text_input("Geburtsort *", value=db_geburtsort)
            
            liste_familienstand = ["Ledig", "Verheiratet", "Eingetragene Lebenspartnerschaft", "Geschieden", "Verwitwet"]
            default_fam_index = liste_familienstand.index(db_familienstand) if db_familienstand in liste_familienstand else 0
            familienstand = st.selectbox("Familienstand", liste_familienstand, index=default_fam_index)
            
            liste_vn_status = ["Angestellter", "Beamter", "Selbstständiger", "Freiberufler", "Hausfrau/Hausmann", "Rentner/Pensionär", "Arbeitslos", "Student/Schüler"]
            default_vn_index = liste_vn_status.index(db_vn_status) if db_vn_status in liste_vn_status else 0
            vn_status = st.selectbox("Status des Versicherungsnehmers (Beruf) *", liste_vn_status, index=default_vn_index)
            
            liste_wohneigentum = ["Ja, Eigenheim (Einfamilienhaus)", "Ja, Eigentumswohnung", "Nein, zur Miete"]
            default_wohn_index = liste_wohneigentum.index(db_wohneigentum) if db_wohneigentum in liste_wohneigentum else 2
            wohneigentum = st.selectbox("Wohneigentum vorhanden *", liste_wohneigentum, index=default_wohn_index)
            
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
            
            finanzierungsart = st.selectbox("Finanzierungsart des Fahrzeugs *", ["Eigenfinanziert (Bar / Kauf)", "Geleast", "Kreditfinanziert (Finanzierung)"])
        
        st.write("") 
        col_kz1, col_kz2 = st.columns([2, 1])
        with col_kz2:
            neufahrzeug = st.checkbox("Neufahrzeug", help="Aktivieren Sie dies, falls das Fahrzeug noch nicht zugelassen ist.")
        with col_kz1:
            if neufahrzeug:
                kennzeichen = st.text_input("Amtliches Kennzeichen", value="NEUFAHRZEUG", disabled=True)
            else:
                kennzeichen = st.text_input("Amtliches Kennzeichen *", placeholder="z.B. BA-PG-99")

        fahrzeugschein = st.file_uploader("Fahrzeugschein hier hochladen/fotografieren *", type=["pdf", "png", "jpg", "jpeg"])
        st.markdown('</div>', unsafe_allow_html=True)

        # === KARTE 3: WEITERE FAHRER ===
        st.markdown("""
            <div class="form-card">
                <span class="card-header">👥 Weitere Fahrer</span>
        """, unsafe_allow_html=True)
        
        weitere_fahrer_vorhanden = st.checkbox("Gibt es weitere berechtigte Fahrer für dieses Fahrzeug?", value=False)
        weitere_fahrer_daten_liste = []
        
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
                    weitere_fahrer_daten_liste.append({
                        "id": i + 1,
                        "nachname": str(wf_n).strip(),
                        "vorname": str(wf_v).strip(),
                        "geburtsdatum": wf_geb.strftime('%d.%m.%Y'),
                        "erwerbsdatum": wf_fs.strftime('%d.%m.%Y')
                    })
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
            pflicht_wf = len(weitere_fahrer_daten_liste) == st.session_state.anzahl_fahrer
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

                wf_text_for_doc = "Nein"
                if weitere_fahrer_daten_liste:
                    wf_text_for_doc = " | ".join([f"Fahrer {f['id']}: {f['vorname']} {f['nachname']} (Geb: {f['geburtsdatum']}, Erwerb: {f['erwerbsdatum']})" for f in weitere_fahrer_daten_liste])

                infotext = f"""=== KUNDENDATEN FÜR NAFI / COMPARIT ===
Name: {name}
Vorname: {vorname}
Weitere Vornamen: {weitere_vornamen}
Geburtsort: {geburtsort}
Familienstand: {familienstand}
Status VN: {vn_status}
Wohneigentum: {wohneigentum}
----------------------------------------
Kennzeichen: {kennzeichen}
Fahrleistung: {fahrleistung} km/Jahr
Aktueller KM-Stand: {km_stand} km
Garage: {garage}
Finanzierung: {finanzierungsart}
----------------------------------------
Weitere Fahrer: {wf_text_for_doc}
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
                            match_empty_stamm = df_db[(df_db["PIN"] == str(pin_eingabe)) & (df_db["Kennzeichen"].isna() | (df_db["Kennzeichen"] == ""))]
                            
                            neue_zeile = {
                                "PIN": str(pin_eingabe),
                                "Vorname": str(vorname),
                                "Nachname": str(name),
                                "Weitere_Vornamen": str(weitere_vornamen),
                                "Geburtsort": str(geburtsort),
                                "Familienstand": str(familienstand),
                                "VN_Status": str(vn_status),
                                "Wohneigentum": str(wohneigentum),
                                "Kennzeichen": str(kennzeichen),
                                "Fahrleistung": int(fahrleistung),
                                "Kilometerstand": int(km_stand),
                                "Garage": str(garage),
                                "Finanzierung": str(finanzierungsart)
                            }
                            
                            for j in range(1, 6):
                                if j <= len(weitere_fahrer_daten_liste):
                                    f_data = weitere_fahrer_daten_liste[j-1]
                                    neue_zeile[f"Fahrer_{j}_Nachname"] = f_data["nachname"]
                                    neue_zeile[f"Fahrer_{j}_Vorname"] = f_data["vorname"]
                                    neue_zeile[f"Fahrer_{j}_Geburtsdatum"] = f_data["geburtsdatum"]
                                    neue_zeile[f"Fahrer_{j}_Erwerbsdatum"] = f_data["erwerbsdatum"]
                                else:
                                    neue_zeile[f"Fahrer_{j}_Nachname"] = ""
                                    neue_zeile[f"Fahrer_{j}_Vorname"] = ""
                                    neue_zeile[f"Fahrer_{j}_Geburtsdatum"] = ""
                                    neue_zeile[f"Fahrer_{j}_Erwerbsdatum"] = ""
                            
                            if not match_auto.empty:
                                idx = match_auto.index[0]
                                for col, val in neue_zeile.items():
                                    df_db.at[idx, col] = val
                            elif not match_empty_stamm.empty:
                                idx = match_empty_stamm.index[0]
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
