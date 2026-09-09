import streamlit as st
import pandas as pd
import datetime
import os
import json
import shutil

# --- SAYFA AYARLARI ---
st.set_page_config(page_title="ALASAR GRUP - Kalite Yönetim Sistemi", page_icon="🛡️", layout="wide")

# --- DOSYA VE KLASÖR YOLLARI ---
EXCEL_FILE = "alasar_kalite_vt.xlsx"
USERS_FILE = "users.json"
UPLOAD_DIR = "yuklenen_belgeler"
ARCHIVE_DIR = "arsivlenenler"

for d in [UPLOAD_DIR, ARCHIVE_DIR]:
    if not os.path.exists(d):
        os.makedirs(d)

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
            pd.DataFrame(columns=["Tarih / Saat", "Departman", "Doküman No", "Doküman Adı", "Açıklama / Not", "Dosya Adı", "Ekleyen", "Revizyon Mu"]).to_excel(writer, sheet_name="Departman_Dokumanlari", index=False)
            pd.DataFrame(columns=["Tarih / Saat", "Departman", "Doküman No", "Doküman Adı", "Açıklama / Not", "Dosya Adı", "Ekleyen", "Arşivlenme Tarihi"]).to_excel(writer, sheet_name="Arsiv_Dokumanlari", index=False)
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
    st.sidebar.success("✏️ Doküman Yükleme / Revize Yetkisi Var")
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

# --- KULLANICI ŞİFRE DEĞİŞTİRME ALANI ---
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

# --- CANLI BİLDİRİM PANELİ ---
df_notif_top = load_data("Bildirimler")
if not df_notif_top.empty:
    latest = df_notif_top.iloc[0]
    dept_val = latest.get('Departman / Modül', latest.get('Modül', 'Genel'))
    time_val = latest.get('Tarih / Saat', '-')
    user_val = latest.get('İşlemi Yapan', '-')
    detail_val = latest.get('Detay / Doküman', '-')
    st.info(f"🔔 **Son Güncelleme / Revizyon Bildirimi:** [{time_val}] **{user_val}** tarafından **{dept_val}** alanında işlem yapıldı: *{detail_val}*")

# --- BİLDİRİM GEÇMİŞİ MODÜLÜ ---
if modul == "🔔 Bildirim Geçmişi":
    st.title("🔔 Tüm Güncelleme & Revizyon Geçmişi")
    st.write("Sistem üzerinde yapılan tüm departman doküman yükleme, revizyon ve arşivleme işlemlerinin dökümü:")
    df_notif_all = load_data("Bildirimler")
    st.dataframe(df_notif_all, use_container_width=True)

# --- DEPARTMAN MODÜLLERİ ---
else:
    dept_name = modul.replace("👥 ", "").replace("⚙️ ", "").replace("👔 ", "").replace("🌐 ", "").replace("🛡️ ", "").replace("📦 ", "")
    st.title(f"📂 {dept_name}")
    st.caption(f"{dept_name} bünyesine ait tüm prosedür, talimat, form ve revize dokümanlar bu alanda yönetilir.")
    
    df_docs = load_data("Departman_Dokumanlari")
    df_archive = load_data("Arsiv_Dokumanlari")
    
    # DOKÜMAN YÜKLEME VE REVİZYON KONTROL ALANI
    if st.session_state["can_edit"]:
        st.subheader(f"📤 {dept_name} İçin Doküman Yükleme / Revize Etme")
        
        with st.form(f"form_{dept_name}"):
            col1, col2, col3 = st.columns([1, 2, 2])
            doc_no = col1.text_input("Doküman No / Kodu (Örn: PR-01, FR-05)").strip().upper()
            doc_title = col2.text_input("Doküman Adı (Örn: İK Prosedürü, İzin Formu)")
            doc_note = col3.text_input("Açıklama / Revizyon Notu (Örn: Maddeler Güncellendi)")
            uploaded_file = st.file_uploader("Dosya Seçiniz (PDF, Word, Excel vb.)", type=["pdf", "png", "jpg", "jpeg", "xlsx", "docx", "zip", "rar"])
            
            submit_doc = st.form_submit_button("🔍 Dokümanı İncele ve İlerle")
        
        if submit_doc:
            if not doc_no or not doc_title or uploaded_file is None:
                st.error("Lütfen Doküman Numarası, Doküman Adı giriniz ve bir dosya seçiniz.")
            else:
                # DOKÜMAN NO KARŞILAŞTIRMA KONTROLÜ
                existing = df_docs[(df_docs["Departman"] == dept_name) & (df_docs["Doküman No"] == doc_no)]
                
                if not existing.empty:
                    old_row = existing.iloc[0]
                    st.warning(f"⚠️ **DİKKAT:** `{doc_no}` numaralı **'{old_row['Doküman Adı']}'** isimli doküman bu departmanda zaten mevcut!")
                    st.info(f"📌 **Mevcut Dosya:** {old_row['Dosya Adı']} | **Yükleyen:** {old_row['Ekleyen']} | **Tarih:** {old_row['Tarih / Saat']}")
                    
                    st.session_state["pending_rev"] = {
                        "dept": dept_name,
                        "doc_no": doc_no,
                        "doc_title": doc_title,
                        "doc_note": doc_note,
                        "uploaded_file": uploaded_file,
                        "old_row": old_row.to_dict()
                    }
                else:
                    # Yeni doküman kaydı
                    file_name = save_uploaded_file(uploaded_file)
                    now_str = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
                    
                    new_rec = {
                        "Tarih / Saat": now_str,
                        "Departman": dept_name,
                        "Doküman No": doc_no,
                        "Doküman Adı": doc_title,
                        "Açıklama / Not": doc_note,
                        "Dosya Adı": file_name,
                        "Ekleyen": st.session_state["username"],
                        "Revizyon Mu": "Hayır"
                    }
                    df_docs = pd.concat([pd.DataFrame([new_rec]), df_docs], ignore_index=True)
                    save_data(df_docs, "Departman_Dokumanlari")
                    add_notification(st.session_state["username"], dept_name, f"Yeni Doküman Eklendi: {doc_no} - {doc_title}")
                    st.success(f"✅ `{doc_no}` numaralı yeni doküman başarıyla yüklendi!")
                    st.rerun()

        # EĞER REVİZYON ÇAKIŞMASI VARSA ONAY BUTONLARI
        if "pending_rev" in st.session_state and st.session_state["pending_rev"]["dept"] == dept_name:
            p = st.session_state["pending_rev"]
            st.error("Bu işlem bir **REVİZYON** güncellemesi mi?")
            col_rev1, col_rev2 = st.columns(2)
            
            if col_rev1.button("🔄 EVET, Bu Bir Revizyondur (Eski Dosyayı Arşive Kaldır)"):
                now_str = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
                old_info = p["old_row"]
                
                # 1. Eski Dosyayı 'arsivlenenler' Klasörüne Taşı
                old_file_name = old_info["Dosya Adı"]
                if old_file_name != "Yok":
                    src_p = os.path.join(UPLOAD_DIR, old_file_name)
                    dst_p = os.path.join(ARCHIVE_DIR, old_file_name)
                    if os.path.exists(src_p):
                        shutil.move(src_p, dst_p)
                
                # 2. Eski Doküman Verisini Excel 'Arsiv_Dokumanlari' Sayfasına Ekle
                old_archive_rec = {
                    "Tarih / Saat": old_info["Tarih / Saat"],
                    "Departman": dept_name,
                    "Doküman No": old_info["Doküman No"],
                    "Doküman Adı": old_info["Doküman Adı"],
                    "Açıklama / Not": old_info["Açıklama / Not"],
                    "Dosya Adı": old_file_name,
                    "Ekleyen": old_info["Ekleyen"],
                    "Arşivlenme Tarihi": now_str
                }
                df_archive = pd.concat([pd.DataFrame([old_archive_rec]), df_archive], ignore_index=True)
                save_data(df_archive, "Arsiv_Dokumanlari")
                
                # 3. Eski Dokümanı Aktif Listeden Çıkar
                df_docs = df_docs[~((df_docs["Departman"] == dept_name) & (df_docs["Doküman No"] == p["doc_no"]))]
                
                # 4. Yeni Revize Dosyayı Kaydet ve Aktif Listeye Ekle
                new_file_name = save_uploaded_file(p["uploaded_file"])
                new_rec = {
                    "Tarih / Saat": now_str,
                    "Departman": dept_name,
                    "Doküman No": p["doc_no"],
                    "Doküman Adı": p["doc_title"],
                    "Açıklama / Not": p["doc_note"],
                    "Dosya Adı": new_file_name,
                    "Ekleyen": st.session_state["username"],
                    "Revizyon Mu": "Evet"
                }
                df_docs = pd.concat([pd.DataFrame([new_rec]), df_docs], ignore_index=True)
                save_data(df_docs, "Departman_Dokumanlari")
                
                add_notification(st.session_state["username"], dept_name, f"REVİZYON YAPILDI: {p['doc_no']} - {p['doc_title']} (Eski versiyon arşive kaldırıldı)")
                del st.session_state["pending_rev"]
                st.success("✅ Revizyon başarıyla işlendi! Eski versiyon 'arsivlenenler' klasörüne kaldırıldı.")
                st.rerun()

            if col_rev2.button("❌ HAYIR, Farklı Doküman (İptal Et)"):
                del st.session_state["pending_rev"]
                st.info("İşlem iptal edildi.")
                st.rerun()
                
        st.markdown("---")

    # AİT OLDUĞU DEPARTMANIN DOKÜMANLARINI SEKMELERLE LİSTELEME
    tab1, tab2, tab3 = st.tabs(["📄 Aktif Dokümanlar", "🔄 Son Revizeler / Değişiklikler", "📁 Arşivlenen Eski Versiyonlar"])

    if not df_docs.empty and "Departman" in df_docs.columns:
        dept_active_docs = df_docs[df_docs["Departman"] == dept_name]
    else:
        dept_active_docs = pd.DataFrame()

    if not df_archive.empty and "Departman" in df_archive.columns:
        dept_archive_docs = df_archive[df_archive["Departman"] == dept_name]
    else:
        dept_archive_docs = pd.DataFrame()

    # TAB 1: AKTİF DOKÜMANLAR
    with tab1:
        if dept_active_docs.empty:
            st.warning("Henüz bu departmana ait aktif doküman bulunmuyor.")
        else:
            st.dataframe(dept_active_docs[["Tarih / Saat", "Doküman No", "Doküman Adı", "Açıklama / Not", "Dosya Adı", "Ekleyen", "Revizyon Mu"]], use_container_width=True)
            st.markdown("#### 📥 Güncel Dosyaları İndir")
            for idx, row in dept_active_docs.iterrows():
                f_name = row["Dosya Adı"]
                if f_name and f_name != "Yok":
                    f_path = os.path.join(UPLOAD_DIR, f_name)
                    if os.path.exists(f_path):
                        c1, c2 = st.columns([3, 1])
                        rev_badge = " 🔄 **[REVİZE DOKÜMAN]**" if row.get("Revizyon Mu") == "Evet" else ""
                        c1.write(f"📄 **[{row['Doküman No']}]** {row['Doküman Adı']}{rev_badge} - *{row['Açıklama / Not']}*")
                        with open(f_path, "rb") as f:
                            c2.download_button(label="📥 İndir", data=f, file_name=f_name, key=f"active_{idx}_{f_name}")

    # TAB 2: SON REVİZELER / DEĞİŞİKLİKLER
    with tab2:
        revised_docs = dept_active_docs[dept_active_docs["Revizyon Mu"] == "Evet"]
        if revised_docs.empty:
            st.info("Bu departmanda henüz revize edilmiş bir doküman bulunmuyor.")
        else:
            st.subheader("🔄 Son Revize Edilen Güncel Dokümanlar")
            st.dataframe(revised_docs[["Tarih / Saat", "Doküman No", "Doküman Adı", "Açıklama / Not", "Dosya Adı", "Ekleyen"]], use_container_width=True)
            for idx, row in revised_docs.iterrows():
                f_name = row["Dosya Adı"]
                if f_name and f_name != "Yok":
                    f_path = os.path.join(UPLOAD_DIR, f_name)
                    if os.path.exists(f_path):
                        c1, c2 = st.columns([3, 1])
                        c1.write(f"🔄 **[{row['Doküman No']}]** {row['Doküman Adı']} - *Son Revizyon Notu: {row['Açıklama / Not']}*")
                        with open(f_path, "rb") as f:
                            c2.download_button(label="📥 Son Revizyonu İndir", data=f, file_name=f_name, key=f"rev_{idx}_{f_name}")

    # TAB 3: ARŞİVLENEN ESKİ VERSİYONLAR
    with tab3:
        if dept_archive_docs.empty:
            st.info("Bu departman için arşivlenmiş eski bir doküman versiyonu bulunmuyor.")
        else:
            st.subheader("📁 Arşive Kaldırılan Eski Versiyon Dokümanlar")
            st.dataframe(dept_archive_docs[["Arşivlenme Tarihi", "Doküman No", "Doküman Adı", "Açıklama / Not", "Dosya Adı", "Ekleyen"]], use_container_width=True)
            for idx, row in dept_archive_docs.iterrows():
                f_name = row["Dosya Adı"]
                if f_name and f_name != "Yok":
                    f_path = os.path.join(ARCHIVE_DIR, f_name)
                    if os.path.exists(f_path):
                        c1, c2 = st.columns([3, 1])
                        c1.write(f"📁 **[{row['Doküman No']}]** {row['Doküman Adı']} *(Eski Versiyon)* - Arşivlenme: {row['Arşivlenme Tarihi']}")
                        with open(f_path, "rb") as f:
                            c2.download_button(label="📥 Eski Versiyonu İndir", data=f, file_name=f_name, key=f"arch_{idx}_{f_name}")
