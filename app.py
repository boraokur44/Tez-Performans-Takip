import streamlit as st
import pandas as pd
import plotly.express as px
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time
from xhtml2pdf import pisa
from io import BytesIO
import os

# 1. Sayfa Ayarları
st.set_page_config(page_title="Performans Takip Sistemi", layout="wide", initial_sidebar_state="expanded")

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

# --- ŞİFRELİ GİRİŞ EKRANI ---
if not st.session_state["logged_in"]:
    st.markdown("<h2 style='text-align: center; color: #2C3E50; margin-top:50px;'>Nevşehir Hacı Bektaş Veli Üniversitesi</h2>", unsafe_allow_html=True)
    st.markdown("<h4 style='text-align: center; color: #7F8C8D;'>Turizm Araştırmaları Enstitüsü Yönetim Paneli</h4>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        st.write("")
        with st.form("login_form"):
            st.write("🔒 **Lütfen Giriş Yapınız**")
            kullanici = st.text_input("Kullanıcı Adı")
            sifre = st.text_input("Şifre", type="password")
            submit = st.form_submit_submit_button("Giriş Yap", use_container_width=True)
            
            if submit:
                try:
                    if kullanici == st.secrets["kimlik"]["kullanici_adi"] and sifre == st.secrets["kimlik"]["sifre"]:
                        st.session_state["logged_in"] = True
                        st.rerun()
                    else:
                        st.error("Kullanıcı adı veya şifre hatalı!")
                except Exception as e:
                    st.error("Secrets (Şifre) ayarları bulunamadı. Lütfen Streamlit Cloud 'Advanced Settings' menüsünü kontrol edin.")
else:
    # --- ANA UYGULAMA (GİRİŞ YAPILDIKTAN SONRA) ---
    try:
        # Şablon Oluşturucu (Bulut için)
        if not os.path.exists("veriler.xlsx"):
            with pd.ExcelWriter("veriler.xlsx", engine="openpyxl") as writer:
                pd.DataFrame(columns=["Yıl", "Öğrenci Adı Soyadı", "Kayıt olunan Anabilim dalı", "Program", "Danışman Adı ve Soyadı", "Tez Adı", "Eserin Yayın Yılı", "Eserdeki Yazar Sayısı", "Eser Türü", "Eserin Indeksi", "Eserin Linki"]).to_excel(writer, sheet_name="Sheet1", index=False)
                pd.DataFrame(columns=["Danışman Adı ve Soyadı", "E-Posta Adresi"]).to_excel(writer, sheet_name="Mailler", index=False)
                pd.DataFrame(columns=["Danışman Adı ve Soyadı", "Eser Türü", "Eserin Indeksi", "Puan"]).to_excel(writer, sheet_name="Kişisel Performans", index=False)

        # Verileri Okuma
        df = pd.read_excel("veriler.xlsx", sheet_name=0)
        try: mailler_df = pd.read_excel("veriler.xlsx", sheet_name="Mailler")
        except: mailler_df = pd.DataFrame(columns=["Danışman Adı ve Soyadı", "E-Posta Adresi"])
        try: kisisel_df = pd.read_excel("veriler.xlsx", sheet_name="Kişisel Performans")
        except: kisisel_df = pd.DataFrame(columns=["Danışman Adı ve Soyadı", "Eser Türü", "Eserin Indeksi", "Puan"])

        # YAN MENÜ TASARIMI
        st.sidebar.image("https://nevsehir.edu.tr/assets/images/logo.png", width=150)
        menu = st.sidebar.radio("📌 GÖREV MENÜSÜ", ["📊 Performans Paneli", "💾 Veri Girişi ve Düzenleme"])
        
        st.sidebar.divider()
        if st.sidebar.button("🚪 Çıkış Yap"):
            st.session_state["logged_in"] = False
            st.rerun()

        # =========================================================================
        # SAYFA 1: VERİ YÖNETİMİ
        # =========================================================================
        if menu == "💾 Veri Girişi ve Düzenleme":
            st.title("💾 Sistem Veri Yönetimi")
            st.info("💡 **ÖNEMLİ BİLGİ:** Bulut sistemlerde veriler geçici sunucularda tutulur. Veri girişinizi tamamladıktan sonra mutlaka **'Güncel Verileri İndir'** butonuna basarak kopyanızı bilgisayarınıza kaydedin.")
            
            col_up, col_down = st.columns(2)
            with col_up:
                st.subheader("📥 Veri Yükle (Excel'den)")
                uploaded_file = st.file_uploader("Elinizdeki güncel 'veriler.xlsx' dosyasını sisteme yükleyin:", type=["xlsx"])
                if uploaded_file:
                    with open("veriler.xlsx", "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    st.success("✅ Yeni Excel başarıyla sisteme yüklendi!")
                    time.sleep(1)
                    st.rerun()
                    
            with col_down:
                st.subheader("📤 Veri İndir (Yedekle)")
                st.write("Sistemde düzenlenen en güncel verileri bilgisayarınıza indirin.")
                if os.path.exists("veriler.xlsx"):
                    with open("veriler.xlsx", "rb") as f:
                        st.download_button(label="📥 Güncel Verileri İndir (Excel)", data=f, file_name="guncel_veriler.xlsx", mime="application/vnd.ms-excel", use_container_width=True)

            st.divider()
            st.subheader("📝 Sistem Üzerinden Doğrudan Veri Girişi")
            
            tab1, tab2, tab3 = st.tabs(["1️⃣ Lisansüstü Tezler", "2️⃣ Kişisel Performanslar", "3️⃣ E-Posta Adresleri"])
            with tab1: yeni_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)
            with tab2: yeni_kisisel = st.data_editor(kisisel_df, num_rows="dynamic", use_container_width=True)
            with tab3: yeni_mailler = st.data_editor(mailler_df, num_rows="dynamic", use_container_width=True)

            if st.button("💾 Tablodaki Değişiklikleri Sisteme Kaydet", type="primary", use_container_width=True):
                with pd.ExcelWriter("veriler.xlsx", engine="openpyxl") as writer:
                    yeni_df.to_excel(writer, sheet_name="Sheet1", index=False)
                    yeni_mailler.to_excel(writer, sheet_name="Mailler", index=False)
                    yeni_kisisel.to_excel(writer, sheet_name="Kişisel Performans", index=False)
                st.success("✅ Veriler sisteme başarıyla kaydedildi. 'Performans Paneli'ne geçebilirsiniz.")

        # =========================================================================
        # SAYFA 2: PERFORMANS PANELİ (DASHBOARD)
        # =========================================================================
        elif menu == "📊 Performans Paneli":
            st.title("Nevşehir Hacı Bektaş Veli Üniversitesi")
            st.subheader("Turizm Araştırmaları Enstitüsü - Bütünleşik Akademik Performans Sistemi")

            # Veri Temizliği
            df["Danışman Adı ve Soyadı"] = df["Danışman Adı ve Soyadı"].astype(str).str.strip()
            kisisel_df["Danışman Adı ve Soyadı"] = kisisel_df["Danışman Adı ve Soyadı"].astype(str).str.strip()
            mailler_df["Danışman Adı ve Soyadı"] = mailler_df["Danışman Adı ve Soyadı"].astype(str).str.strip()
            
            if "Kayıt olunan Anabilim dalı" not in df.columns:
                df["Kayıt olunan Anabilim dalı"] = "Belirtilmemiş"

            df["Eser Türü"] = df["Eser Türü"].fillna("Yayın Yok")
            df["Eserin Indeksi"] = df["Eserin Indeksi"].fillna("İndeks Yok")
            kisisel_df["Eser Türü"] = kisisel_df.get("Eser Türü", pd.Series(dtype=str)).fillna("Yayın Yok")
            kisisel_df["Eserin Indeksi"] = kisisel_df.get("Eserin Indeksi", pd.Series(dtype=str)).fillna("İndeks Yok")

            hoca_abd_map = df.drop_duplicates(subset=["Danışman Adı ve Soyadı"])[["Danışman Adı ve Soyadı", "Kayıt olunan Anabilim dalı"]]

            # PUANLAMA MOTORLARI
            def puan_hesapla_lisansustu(row):
                try:
                    t = str(row.get("Eser Türü", "")).upper()
                    i = str(row.get("Eserin Indeksi", "")).upper().replace(" ", "")
                    if t == "YAYIN YOK" or t == "NAN" or t == "": return 0.0
                    puan = 0.0
                    if "MAKALE" in t:
                        if "SSCI" in i or "SCI" in i: puan = 1.00
                        elif "ESCI" in i or "SCOPUS" in i or "AHCI" in i: puan = 0.70
                        elif "TRDIZIN" in i or "TRDİZİN" in i: puan = 0.20
                        else: puan = 0.10
                    elif "KİTAP" in t or "KITAP" in t:
                        if "BÖLÜM" in t or "BOLUM" in t:
                            if "BKCI" in i: puan = 0.70
                            elif "ULUSAL" in i: puan = 0.20
                            else: puan = 0.20
                        else:
                            if "BKCI" in i: puan = 1.00
                            elif "ULUSAL" in i: puan = 0.20
                            else: puan = 0.30
                    elif "KONGRE" in t or "BİLDİRİ" in t or "BILDIRI" in t or "SEMPOZYUM" in t:
                        puan = 0.05
                    elif "PROJE" in t:
                        if "1001" in i: puan = 1.00
                        elif "1002" in i: puan = 0.50
                        elif "2209" in i: puan = 0.25
                        elif "AB" in i: puan = 1.00
                        elif "İHTİSAS" in i or "IHTISAS" in i: puan = 0.50
                        elif "BAP" in i: puan = 0.15
                        else: puan = 0.05
                    return puan
                except: return 0.0

            def puan_hesapla_kisisel(row):
                if "Puan" in row and pd.notna(row["Puan"]):
                    try: return float(row["Puan"])
                    except: pass
                try:
                    t = str(row.get("Eser Türü", "")).upper()
                    i = str(row.get("Eserin Indeksi
