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
            pd.DataFrame(columns=["Sertifika Adı", "Kurum", "Alınış Tarihi", "Geçerlilik Tarihi", "Dosya Adı", "Ekleyen", "Durum"]).to_excel(writer, sheet_name="Sertifikalar", index=False)
            pd.DataFrame(columns=["Cihaz / Ekipman", "Seri No", "Son Kalibrasyon", "Gelecek Kalibrasyon", "Firma", "Dosya Adı", "Ekleyen", "Durum"]).to_excel(writer, sheet_name="Kalibrasyon", index=False)
            pd.DataFrame(columns=["Rapor No", "Ürün/Parça No", "Kontrol Tipi", "Miktar", "Uygunluk", "Dosya Adı", "Kontrolör"]).to_excel(writer, sheet_name="Kalite_Kontrol", index=False)
            pd.DataFrame(columns=["Standart", "Madde No", "Uygunsuzluk / Bulgu", "Sorumlu", "Hedef Tarih", "Dosya Adı", "Ekleyen", "Durum"]).to_excel(writer, sheet_name="Entegre_Yonetim", index=False)
            pd.DataFrame(columns=["Tarih / Saat", "İşlemi Yapan", "Modül", "Detay / Doküman"]).to_excel(writer, sheet_name="Bildirimler", index=False)
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
        "Modül": modul,
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
    st.subheader("Kalite Yönetim Sistemi - Giriş Paneli")
    
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
                    # Yükleme Yetkisi Sadece Ömer OCAK ve Dilber Alaşar için Aktif
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
    "Modül Seçiniz:",
    [
        "📜 Sertifikalar",
        "🔧 Kalibrasyon Takibi",
        "🔍 Kalite Kontrol",
        "🌐 Entegre Yönetim Sistemleri (ISO)",
        "📁 Yüklenen Dosyalar Kütüphanesi",
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

# --- CANLI BİLDİRİM PANELİ (SİTE EN ÜST KISMI) ---
df_notif_top = load_data("Bildirimler")
if not df_notif_top.empty:
    latest = df_notif_top.iloc[0]
    st.info(f"🔔 **Son Güncelleme / Revizyon Bildirimi:** [{latest['Tarih / Saat']}] **{latest['İşlemi Yapan']}** tarafından **{latest['Modül']}** modülüne yeni kayıt/doküman eklendi: *{latest['Detay / Doküman']}*")

# --- MODÜL 1: SERTİFİKALAR ---
if modul == "📜 Sertifikalar":
    st.title("📜 Sertifika Yönetimi")
    df = load_data("Sertifikalar")
    
    if st.session_state["can_edit"]:
        with st.form("sertifika_form", clear_on_submit=True):
            st.subheader("➕ Yeni Sertifika / Belge Ekle")
            col1, col2 = st.columns(2)
            sertifika_adi = col1.text_input("Sertifika / Belge Adı")
            kurum = col2.text_input("Belgelendirme Kuruluşu")
            alinis = col1.date_input("Alınış Tarihi")
            gecerlilik = col2.date_input("Geçerlilik Bitiş Tarihi")
            uploaded_doc = st.file_uploader("📄 Sertifika Dokümanı Yükle (PDF / Görsel)", type=["pdf", "png", "jpg", "docx"])
            
            if st.form_submit_button("💾 Sertifikayı Ekle"):
                file_name = save_uploaded_file(uploaded_doc)
                yeni = {
                    "Sertifika Adı": sertifika_adi,
                    "Kurum": kurum,
                    "Alınış Tarihi": str(alinis),
                    "Geçerlilik Tarihi": str(gecerlilik),
                    "Dosya Adı": file_name,
                    "Ekleyen": st.session_state["username"],
                    "Durum": "Aktif"
                }
                df = pd.concat([df, pd.DataFrame([yeni])], ignore_index=True)
                save_data(df, "Sertifikalar")
                
                detay_metni = f"{sertifika_adi} ({file_name if file_name != 'Yok' else 'Dokümansız'})"
                add_notification(st.session_state["username"], "Sertifikalar", detay_metni)
                
                st.success("Sertifika kaydedildi ve bildirim oluşturuldu!")
                st.rerun()
    else:
        st.warning("⚠️ Yeni belge ekleme / düzenleme yetkiniz bulunmamaktadır. Aşağıdaki listeden mevcut belgeleri inceleyip indirebilirsiniz.")
            
    st.dataframe(df, use_container_width=True)

# --- MODÜL 2: KALİBRASYON ---
elif modul == "🔧 Kalibrasyon Takibi":
    st.title("🔧 Cihaz & Ekipman Kalibrasyon Takibi")
    df = load_data("Kalibrasyon")
    
    if st.session_state["can_edit"]:
        with st.form("kalibrasyon_form", clear_on_submit=True):
            st.subheader("➕ Yeni Kalibrasyon Kaydı Ekle")
            col1, col2 = st.columns(2)
            cihaz = col1.text_input("Cihaz / Ölçüm Aleti Adı")
            seri_no = col2.text_input("Seri No / Kod")
            son_kal = col1.date_input("Son Kalibrasyon Tarihi")
            gelecek_kal = col2.date_input("Gelecek Kalibrasyon Tarihi")
            firma = col1.text_input("Kalibrasyon Yapan Firma")
            uploaded_doc = st.file_uploader("📄 Kalibrasyon Raporu / Sertifikası Yükle", type=["pdf", "png", "jpg", "xlsx", "docx"])
            
            if st.form_submit_button("💾 Kalibrasyon Kaydını Ekle"):
                file_name = save_uploaded_file(uploaded_doc)
                yeni = {
                    "Cihaz / Ekipman": cihaz,
                    "Seri No": seri_no,
                    "Son Kalibrasyon": str(son_kal),
                    "Gelecek Kalibrasyon": str(gelecek_kal),
                    "Firma": firma,
                    "Dosya Adı": file_name,
                    "Ekleyen": st.session_state["username"],
                    "Durum": "Geçerli"
                }
                df = pd.concat([df, pd.DataFrame([yeni])], ignore_index=True)
                save_data(df, "Kalibrasyon")
                
                detay_metni = f"{cihaz} (Seri No: {seri_no}) - {file_name}"
                add_notification(st.session_state["username"], "Kalibrasyon", detay_metni)
                
                st.success("Kalibrasyon kaydı eklendi ve bildirim oluşturuldu!")
                st.rerun()
    else:
        st.warning("⚠️ Yeni belge ekleme / düzenleme yetkiniz bulunmamaktadır. Aşağıdaki listeden kayıtları inceleyebilirsiniz.")
            
    st.dataframe(df, use_container_width=True)

# --- MODÜL 3: KALİTE KONTROL ---
elif modul == "🔍 Kalite Kontrol":
    st.title("🔍 Giriş & Proses Kalite Kontrol")
    df = load_data("Kalite_Kontrol")
    
    if st.session_state["can_edit"]:
        with st.form("kk_form", clear_on_submit=True):
            st.subheader("➕ Yeni Kalite Kontrol Raporu Ekle")
            col1, col2 = st.columns(2)
            rapor_no = col1.text_input("Rapor No", f"KK-{datetime.date.today().year}-001")
            parca = col2.text_input("Ürün / Parça Kodu")
            ktipi = col1.selectbox("Kontrol Tipi", ["Giriş Kalite", "Proses Kontrol", "Final Kontrol"])
            miktar = col2.number_input("Kontrol Edilen Miktar", min_value=1, value=100)
            karar = col1.selectbox("Uygunluk Kararı", ["Kabul", "Şartlı Kabul", "Red / Karantina"])
            uploaded_doc = st.file_uploader("📄 Kalite Kontrol Raporu / Ölçüm Raporu Yükle", type=["pdf", "png", "jpg", "xlsx", "docx"])
            
            if st.form_submit_button("💾 Raporu Kaydet"):
                file_name = save_uploaded_file(uploaded_doc)
                yeni = {
                    "Rapor No": rapor_no,
                    "Ürün/Parça No": parca,
                    "Kontrol Tipi": ktipi,
                    "Miktar": miktar,
                    "Uygunluk": karar,
                    "Dosya Adı": file_name,
                    "Kontrolör": st.session_state["username"]
                }
                df = pd.concat([df, pd.DataFrame([yeni])], ignore_index=True)
                save_data(df, "Kalite_Kontrol")
                
                detay_metni = f"Rapor No: {rapor_no} - Parça: {parca} ({karar})"
                add_notification(st.session_state["username"], "Kalite Kontrol", detay_metni)
                
                st.success("Kalite kontrol kaydı oluşturuldu ve bildirim gönderildi!")
                st.rerun()
    else:
        st.warning("⚠️ Yeni belge ekleme / düzenleme yetkiniz bulunmamaktadır. Aşağıdaki listeden kayıtları inceleyebilirsiniz.")
            
    st.dataframe(df, use_container_width=True)

# --- MODÜL 4: EYS (ISO) ---
elif modul == "🌐 Entegre Yönetim Sistemleri (ISO)":
    st.title("🌐 Entegre Yönetim Sistemleri (EYS)")
    df = load_data("Entegre_Yonetim")
    
    if st.session_state["can_edit"]:
        with st.form("eys_form", clear_on_submit=True):
            st.subheader("➕ Yeni EYS Aksiyonu / Dokümanı Ekle")
            col1, col2 = st.columns(2)
            iso_std = col1.selectbox("ISO Standardı", ["ISO 9001 (Kalite)", "ISO 14001 (Çevre)", "ISO 45001 (İSG)"])
            madde = col2.text_input("Madde / Süreç No (Örn: 8.5.1)")
            bulgu = st.text_area("Tetkik Bulgusu / Uygunsuzluk Tanımı")
            sorumlu = col1.text_input("Aksiyon Sorumlusu")
            hedef = col2.date_input("Hedef Kapanış Tarihi")
            uploaded_doc = st.file_uploader("📄 Prosedür / DÖF Formu / Doküman Yükle", type=["pdf", "png", "jpg", "xlsx", "docx"])
            
            if st.form_submit_button("💾 EYS Kaydını Ekle"):
                file_name = save_uploaded_file(uploaded_doc)
                yeni = {
                    "Standart": iso_std,
                    "Madde No": madde,
                    "Uygunsuzluk / Bulgu": bulgu,
                    "Sorumlu": sorumlu,
                    "Hedef Tarih": str(hedef),
                    "Dosya Adı": file_name,
                    "Ekleyen": st.session_state["username"],
                    "Durum": "Açık"
                }
                df = pd.concat([df, pd.DataFrame([yeni])], ignore_index=True)
                save_data(df, "Entegre_Yonetim")
                
                detay_metni = f"{iso_std} - Madde: {madde} ({file_name})"
                add_notification(st.session_state["username"], "Entegre Yönetim Sistemleri", detay_metni)
                
                st.success("EYS dokümanı eklendi ve bildirim oluşturuldu!")
                st.rerun()
    else:
        st.warning("⚠️ Yeni belge ekleme / düzenleme yetkiniz bulunmamaktadır. Aşağıdaki listeden kayıtları inceleyebilirsiniz.")
            
    st.dataframe(df, use_container_width=True)

# --- MODÜL 5: DOSYA KÜTÜPHANESİ ---
elif modul == "📁 Yüklenen Dosyalar Kütüphanesi":
    st.title("📁 Yüklenen Belgeler & Doküman Kütüphanesi")
    
    # SADECE ÖMER OCAK VE DİLBER ALAŞAR DOKÜMAN YÜKLEYEBİLİR
    if st.session_state["can_edit"]:
        with st.form("genel_dosya_form", clear_on_submit=True):
            st.subheader("📤 Yeni Doküman / Belge Yükle")
            uploaded_genel_doc = st.file_uploader("Sisteme Eklemek İstediğiniz Belgeyi Seçiniz (PDF, Word, Excel, Görsel)", type=["pdf", "png", "jpg", "jpeg", "xlsx", "docx", "zip", "rar"])
            
            if st.form_submit_button("🚀 Belgeyi Kütüphaneye Yükle"):
                if uploaded_genel_doc is not None:
                    file_name = save_uploaded_file(uploaded_genel_doc)
                    add_notification(st.session_state["username"], "Dosya Kütüphanesi", f"Yeni Genel Belge Yüklendi: {file_name}")
                    st.success(f"'{file_name}' dosyası kütüphaneye başarıyla yüklendi!")
                    st.rerun()
                else:
                    st.error("Lütfen yüklemek için bir dosya seçiniz.")
        st.markdown("---")
    else:
        st.info("Sistemdeki tüm dokümanları ve yüklenen dosyaları buradan görüntüleyebilir ve indirebilirsiniz.")
    
    # YÜKLENEN DOSYALARI LİSTELEME VE İNDİRME (TÜM KULLANICILAR İÇİN AÇIK)
    files = os.listdir(UPLOAD_DIR)
    if not files:
        st.warning("Henüz sisteme yüklenmiş bir dosya bulunmuyor.")
    else:
        st.subheader("📄 Mevcut Belgeler Listesi")
        for file in files:
            file_path = os.path.join(UPLOAD_DIR, file)
            col1, col2 = st.columns([3, 1])
            col1.write(f"📁 **{file}**")
            with open(file_path, "rb") as f:
                col2.download_button(
                    label="📥 İndir",
                    data=f,
                    file_name=file,
                    key=file
                )

# --- MODÜL 6: BİLDİRİM GEÇMİŞİ ---
elif modul == "🔔 Bildirim Geçmişi":
    st.title("🔔 Tüm Güncelleme & Revizyon Geçmişi")
    st.write("Sistem üzerinde yapılan tüm veri ve belge yükleme/güncelleme işlemlerinin tarihsel dökümü:")
    df_notif_all = load_data("Bildirimler")
    st.dataframe(df_notif_all, use_container_width=True)
