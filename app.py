import streamlit as st
import pandas as pd
import datetime
import os
import json

# --- SAYFA AYARLARI ---
st.set_page_config(page_title="ALASAR GRUP - Kalite Yönetim Sistemi", page_icon="🛡️", layout="wide")

# --- DOSYA VE KLASÖR YOLLARI ---
EXCEL_FILE = "alasar_kalite_vt.xlsx"
USERS_FILE = "users.json"
UPLOAD_DIR = "yuklenen_belgeler"

if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)

# --- VARSAYILAN KULLANICI LİSTESİ VE ŞİFRELER ---
DEFAULT_USERS = {
    "Mehmet Alaşar": {"password": "malsr3434.", "role": "Yönetici", "can_edit": False},
    "Dilber Alaşar": {"password": "dalsr4141.", "role": "Yönetici", "can_edit": True},
    "Nilay Kiraz": {"password": "nkrz5151.", "role": "İK", "can_edit": False},
    "Ömer OCAK": {"password": "oock6161.", "role": "Kalite Sistem Mühendisi", "can_edit": True}
}

# --- KULLANICI VERİLERİNİ YÜKLEME / KAYDETME ---
def load_users():
    if os.path.exists(USERS_FILE):
        try:
            with open(USERS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return DEFAULT_USERS
    else:
        save_users(DEFAULT_USERS)
        return DEFAULT_USERS

def save_users(users_dict):
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users_dict, f, ensure_ascii=False, indent=4)

USERS = load_users()

# --- VERİ TABANI HAZIRLAMA ---
def load_data(sheet_name):
    if not os.path.exists(EXCEL_FILE):
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
            pd.DataFrame(columns=["Tarih / Saat", "Departman", "Doküman Adı", "Açıklama / Not", "Dosya Adı", "Ekleyen"]).to_excel(writer, sheet_name="Departman_Dokumanlari", index=False)
            pd.DataFrame(columns=["Tarih / Saat", "İşlemi Yapan", "Departman / Modül", "Detay / Doküman"]).to_excel(writer, sheet_name="Bildirimler", index=False)
    try:
        return pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()

def save_data(df_new, sheet_name):
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_new.to_excel(writer, sheet_name=sheet_name, index=False)

def save_uploaded_file(uploaded_file):
    if uploaded_file is not None:
        file_path = os.path.join(UPLOAD_DIR, uploaded_file.name)
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        return uploaded_file.name
    return "Yok"

def add_notification(user, modul, detail):
    df_notif = load_data("Bildirimler")
    now_str = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
    new_notif = {
        "Tarih / Saat": now_str,
        "İşlemi Yapan": user,
        "Departman / Modül": modul,
        "Detay / Doküman": detail
    }
    df_notif = pd.concat([pd.DataFrame([new_notif]), df_notif], ignore_index=True)
    save_data(df_notif, "Bildirimler")

# --- OTURUM DURUMU ---
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
if "username" not in st.session_state:
    st.session_state["username"] = ""
if "role" not in st.session_state:
    st.session_state["role"] = ""
if "can_edit" not in st.session_state:
    st.session_state["can_edit"] = False

# --- GİRİŞ EKRANI (LOGIN) ---
if not st.session_state["logged_in"]:
    st.title("🏢 ALASAR GRUP")
    st.subheader("Kalite & Departman Yönetim Sistemi - Giriş Paneli")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.form("login_form"):
            selected_user = st.selectbox("Kullanıcı Seçiniz", list(USERS.keys()))
            input_password = st.text_input("Şifre", type="password")
            submit_login = st.form_submit_button("🔑 Giriş Yap")
            
            if submit_login:
                if input_password == USERS[selected_user]["password"]:
                    st.session_state["logged_in"] = True
                    st.session_state["username"] = selected_user
                    st.session_state["role"] = USERS[selected_user]["role"]
                    st.session_state["can_edit"] = selected_user in ["Ömer OCAK", "Dilber Alaşar"]
                    st.success(f"Hoş geldiniz, {selected_user}!")
                    st.rerun()
                else:
                    st.error("Hatalı şifre! Lütfen tekrar deneyiniz.")
    st.stop()

# --- ANA SİSTEM SOL MENÜ ---
st.sidebar.title("🏢 ALASAR GRUP")
st.sidebar.write(f"👤 **{st.session_state['username']}**")
st.sidebar.caption(f"Rol: {st.session_state['role']}")

if st.session_state["can_edit"]:
    st.sidebar.success("✏️ Doküman Yükleme Yetkisi Var")
else:
    st.sidebar.info("👁️ Sadece Okuma / İndirme Yetkisi Var")

if st.sidebar.button("🚪 Çıkış Yap"):
    st.session_state["logged_in"] = False
    st.session_state["username"] = ""
    st.session_state["role"] = ""
    st.session_state["can_edit"] = False
    st.rerun()

st.sidebar.markdown("---")

if os.path.exists(EXCEL_FILE):
    with open(EXCEL_FILE, "rb") as file:
        st.sidebar.download_button(
            label="📥 Excel Veri Tabanını İndir",
            data=file,
            file_name=f"Alasar_Kalite_VT_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.sidebar.markdown("---")
modul = st.sidebar.radio(
    "DEPARTMANLAR VE MENÜ:",
    [
        "👥 İNSAN KAYNAKLARI DEPARTMANI",
        "⚙️ ÜRETİM DEPARTMANI",
        "👔 YÖNETİM DEPARTMANI",
        "🌐 ENTEGRE YÖNETİM SİSTEMİ DEPARTMANI",
        "🛡️ KALİTE DEPARTMANI",
        "📦 DEPO-SEVKİYAT DEPARTMANI",
        "🔔 Bildirim Geçmişi"
    ]
)

st.sidebar.markdown("---")

# --- KULLANICI ŞİFRE DEĞİŞTİRME ALANI (SOL MENÜ ALT KISIM) ---
with st.sidebar.expander("🔑 Şifremi Değiştir"):
    with st.form("change_password_form", clear_on_submit=True):
        old_pass = st.text_input("Mevcut Şifre", type="password")
        new_pass = st.text_input("Yeni Şifre", type="password")
        new_pass_confirm = st.text_input("Yeni Şifre (Tekrar)", type="password")
        btn_pass = st.form_submit_button("Güncelle")
        
        if btn_pass:
            current_user = st.session_state["username"]
            if old_pass != USERS[current_user]["password"]:
                st.error("Mevcut şifreniz hatalı!")
            elif new_pass != new_pass_confirm:
                st.error("Yeni şifreler eşleşmiyor!")
            elif len(new_pass) < 4:
                st.error("Şifre en az 4 karakter olmalıdır!")
            else:
                USERS[current_user]["password"] = new_pass
                save_users(USERS)
                st.success("Şifreniz başarıyla değiştirildi!")

# --- CANLI BİLDİRİM PANELİ (GÜVENLİ OKUMA KONTROLÜ İLE) ---
df_notif_top = load_data("Bildirimler")
if not df_notif_top.empty:
    latest = df_notif_top.iloc[0]
    
    # Kolon adının eski/yeni Excel dosyalarında uyuşmazlık çıkarmasını önleyen kontrol
    dept_val = latest.get('Departman / Modül', latest.get('Modül', 'Genel'))
    time_val = latest.get('Tarih / Saat', '-')
    user_val = latest.get('İşlemi Yapan', '-')
    detail_val = latest.get('Detay / Doküman', '-')
    
    st.info(f"🔔 **Son Güncelleme / Revizyon Bildirimi:** [{time_val}] **{user_val}** tarafından **{dept_val}** alanına doküman/revizyon eklendi: *{detail_val}*")

# --- BİLDİRİM GEÇMİŞİ MODÜLÜ ---
if modul == "🔔 Bildirim Geçmişi":
    st.title("🔔 Tüm Güncelleme & Revizyon Geçmişi")
    st.write("Sistem üzerinde yapılan tüm departman doküman yükleme ve güncelleme işlemlerinin dökümü:")
    df_notif_all = load_data("Bildirimler")
    st.dataframe(df_notif_all, use_container_width=True)

# --- DEPARTMAN MODÜLLERİ ---
else:
    dept_name = modul.replace("👥 ", "").replace("⚙️ ", "").replace("👔 ", "").replace("🌐 ", "").replace("🛡️ ", "").replace("📦 ", "")
    st.title(f"📂 {dept_name}")
    st.caption(f"{dept_name} bünyesine ait tüm prosedur, talimat, form ve revize dokümanlar bu alanda yönetilir.")
    
    df_docs = load_data("Departman_Dokumanlari")
    
    # SADECE ÖMER OCAK VE DİLBER ALAŞAR DOKÜMAN YÜKLEYEBİLİR
    if st.session_state["can_edit"]:
        with st.form(f"form_{dept_name}", clear_on_submit=True):
            st.subheader(f"📤 {dept_name} İçin Yeni Doküman / Revizyon Yükle")
            col1, col2 = st.columns(2)
            doc_title = col1.text_input("Doküman / Belge Adı (Örn: İK Prosedürü, Kalibrasyon Planı)")
            doc_note = col2.text_input("Açıklama / Revizyon Notu (Örn: Rev.01 Güncellendi)")
            uploaded_file = st.file_uploader("Dosya Seçiniz (PDF, Word, Excel, Görsel vb.)", type=["pdf", "png", "jpg", "jpeg", "xlsx", "docx", "zip", "rar"])
            
            if st.form_submit_button("🚀 Dokümanı Departmana Kaydet ve Bildir"):
                if uploaded_file is not None and doc_title != "":
                    file_name = save_uploaded_file(uploaded_file)
                    now_str = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
                    
                    new_doc_record = {
                        "Tarih / Saat": now_str,
                        "Departman": dept_name,
                        "Doküman Adı": doc_title,
                        "Açıklama / Not": doc_note,
                        "Dosya Adı": file_name,
                        "Ekleyen": st.session_state["username"]
                    }
                    
                    df_docs = pd.concat([pd.DataFrame([new_doc_record]), df_docs], ignore_index=True)
                    save_data(df_docs, "Departman_Dokumanlari")
                    
                    # Bildirim Oluştur
                    add_notification(st.session_state["username"], dept_name, f"{doc_title} ({file_name})")
                    
                    st.success(f"'{doc_title}' dokümanı {dept_name} alanına eklendi ve bildirim gönderildi!")
                    st.rerun()
                else:
                    st.error("Lütfen bir doküman adı yazınız ve yüklenecek dosyayı seçiniz.")
        st.markdown("---")
    else:
        st.info(f"Aşağıda **{dept_name}** için yüklenmiş aktif dokümanları inceleyebilir ve indirebilirsiniz.")

    # AİT OLDUĞU DEPARTMANIN DOSYALARINI FİLTRELEME VE LİSTELEME
    if not df_docs.empty and "Departman" in df_docs.columns:
        filtered_docs = df_docs[df_docs["Departman"] == dept_name]
    else:
        filtered_docs = pd.DataFrame()

    if filtered_docs.empty:
        st.warning(f"Henüz {dept_name} için yüklenmiş bir doküman bulunmuyor.")
    else:
        st.subheader("📄 Departmana Ait Belgeler ve Revizyonlar")
        st.dataframe(filtered_docs[["Tarih / Saat", "Doküman Adı", "Açıklama / Not", "Dosya Adı", "Ekleyen"]], use_container_width=True)
        
        st.markdown("#### 📥 Dosya İndirme Alanı")
        for idx, row in filtered_docs.iterrows():
            f_name = row["Dosya Adı"]
            if f_name and f_name != "Yok":
                f_path = os.path.join(UPLOAD_DIR, f_name)
                if os.path.exists(f_path):
                    col1, col2 = st.columns([3, 1])
                    col1.write(f"📄 **{row['Doküman Adı']}** (*{f_name}*) - {row['Açıklama / Not']}")
                    with open(f_path, "rb") as f:
                        col2.download_button(
                            label="📥 İndir",
                            data=f,
                            file_name=f_name,
                            key=f"{dept_name}_{idx}_{f_name}"
                        )
