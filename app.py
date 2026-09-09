import streamlit as st
import pandas as pd
import datetime
import os

EXCEL_FILE = "alasar_kalite_vt.xlsx"

def load_data(sheet_name):
    if not os.path.exists(EXCEL_FILE):
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
            pd.DataFrame(columns=["Sertifika Adı", "Kurum", "Alınış Tarihi", "Geçerlilik Tarihi", "Durum"]).to_excel(writer, sheet_name="Sertifikalar", index=False)
            pd.DataFrame(columns=["Cihaz / Ekipman", "Seri No", "Son Kalibrasyon", "Gelecek Kalibrasyon", "Firma", "Durum"]).to_excel(writer, sheet_name="Kalibrasyon", index=False)
            pd.DataFrame(columns=["Rapor No", "Ürün/Parça No", "Kontrol Tipi", "Miktar", "Uygunluk", "Kontrolör"]).to_excel(writer, sheet_name="Kalite_Kontrol", index=False)
            pd.DataFrame(columns=["Standart", "Madde No", "Uygunsuzluk / Bulgu", "Sorumlu", "Hedef Tarih", "Durum"]).to_excel(writer, sheet_name="Entegre_Yonetim", index=False)
    
    try:
        return pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()

def save_data(df_new, sheet_name):
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_new.to_excel(writer, sheet_name=sheet_name, index=False)

st.set_page_config(page_title="ALASAR GRUP - Kalite Yönetim Sistemi", page_icon="🛡️", layout="wide")

st.sidebar.title("🏢 ALASAR GRUP")
st.sidebar.subheader("Kalite Yönetim Sistemi")

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
        "🌐 Entegre Yönetim Sistemleri (ISO)"
    ]
)

if modul == "📜 Sertifikalar":
    st.title("📜 Sertifika Yönetimi")
    df = load_data("Sertifikalar")
    
    with st.form("sertifika_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        sertifika_adi = col1.text_input("Sertifika / Belge Adı")
        kurum = col2.text_input("Belgelendirme Kuruluşu")
        alinis = col1.date_input("Alınış Tarihi")
        gecerlilik = col2.date_input("Geçerlilik Bitiş Tarihi")
        
        if st.form_submit_button("💾 Sertifikayı Excel'e Kaydet"):
            yeni = {"Sertifika Adı": sertifika_adi, "Kurum": kurum, "Alınış Tarihi": str(alinis), "Geçerlilik Tarihi": str(gecerlilik), "Durum": "Aktif"}
            df = pd.concat([df, pd.DataFrame([yeni])], ignore_index=True)
            save_data(df, "Sertifikalar")
            st.success("Sertifika kaydedildi!")
            st.rerun()
            
    st.dataframe(df, use_container_width=True)

elif modul == "🔧 Kalibrasyon Takibi":
    st.title("🔧 Cihaz & Ekipman Kalibrasyon Takibi")
    df = load_data("Kalibrasyon")
    
    with st.form("kalibrasyon_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        cihaz = col1.text_input("Cihaz / Ölçüm Aleti Adı")
        seri_no = col2.text_input("Seri No / Kod")
        son_kal = col1.date_input("Son Kalibrasyon Tarihi")
        gelecek_kal = col2.date_input("Gelecek Kalibrasyon Tarihi")
        firma = col1.text_input("Kalibrasyon Yapan Firma")
        
        if st.form_submit_button("💾 Kalibrasyon Kaydını Ekle"):
            yeni = {"Cihaz / Ekipman": cihaz, "Seri No": seri_no, "Son Kalibrasyon": str(son_kal), "Gelecek Kalibrasyon": str(gelecek_kal), "Firma": firma, "Durum": "Geçerli"}
            df = pd.concat([df, pd.DataFrame([yeni])], ignore_index=True)
            save_data(df, "Kalibrasyon")
            st.success("Kalibrasyon kaydı Excel'e eklendi!")
            st.rerun()
            
    st.dataframe(df, use_container_width=True)

elif modul == "🔍 Kalite Kontrol":
    st.title("🔍 Giriş & Proses Kalite Kontrol")
    df = load_data("Kalite_Kontrol")
    
    with st.form("kk_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        rapor_no = col1.text_input("Rapor No", f"KK-{datetime.date.today().year}-001")
        parca = col2.text_input("Ürün / Parça Kodu")
        ktipi = col1.selectbox("Kontrol Tipi", ["Giriş Kalite", "Proses Kontrol", "Final Kontrol"])
        miktar = col2.number_input("Kontrol Edilen Miktar", min_value=1, value=100)
        karar = col1.selectbox("Uygunluk Kararı", ["Kabul", "Şartlı Kabul", "Red / Karantina"])
        kontrolor = col2.text_input("Kontrol Eden")
        
        if st.form_submit_button("💾 Kalite Raporunu Kaydet"):
            yeni = {"Rapor No": rapor_no, "Ürün/Parça No": parca, "Kontrol Tipi": ktipi, "Miktar": miktar, "Uygunluk": karar, "Kontrolör": kontrolor}
            df = pd.concat([df, pd.DataFrame([yeni])], ignore_index=True)
            save_data(df, "Kalite_Kontrol")
            st.success("Kalite kontrol kaydı oluşturuldu!")
            st.rerun()
            
    st.dataframe(df, use_container_width=True)

elif modul == "🌐 Entegre Yönetim Sistemleri (ISO)":
    st.title("🌐 Entegre Yönetim Sistemleri (EYS)")
    df = load_data("Entegre_Yonetim")
    
    with st.form("eys_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        iso_std = col1.selectbox("ISO Standardı", ["ISO 9001 (Kalite)", "ISO 14001 (Çevre)", "ISO 45001 (İSG)"])
        madde = col2.text_input("Madde / Süreç No (Örn: 8.5.1)")
        bulgu = st.text_area("Tetkik Bulgusu / Uygunsuzluk Tanımı")
        sorumlu = col1.text_input("Aksiyon Sorumlusu")
        hedef = col2.date_input("Hedef Kapanış Tarihi")
        
        if st.form_submit_button("💾 EYS Aksiyonunu Kaydet"):
            yeni = {"Standart": iso_std, "Madde No": madde, "Uygunsuzluk / Bulgu": bulgu, "Sorumlu": sorumlu, "Hedef Tarih": str(hedef), "Durum": "Açık"}
            df = pd.concat([df, pd.DataFrame([yeni])], ignore_index=True)
            save_data(df, "Entegre_Yonetim")
            st.success("EYS kaydı Excel'e yazıldı!")
            st.rerun()
            
    st.dataframe(df, use_container_width=True)
