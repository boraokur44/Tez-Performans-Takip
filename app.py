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

# 1. Sayfa Ayarları ve Session State
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
                    if kullanici == st.secrets["auth"]["username"] and sifre == st.secrets["auth"]["password"]:
                        st.session_state["logged_in"] = True
                        st.rerun()
                    else:
                        st.error("Kullanıcı adı veya şifre hatalı!")
                except Exception as e:
                    st.error("Secrets (Şifre) ayarları bulunamadı. Lütfen .streamlit/secrets.toml dosyasını kontrol edin.")
else:
    # --- ANA UYGULAMA (GİRİŞ YAPILDIKTAN SONRA) ---
    
    # Eğer veriler.xlsx yoksa, boş bir şablon oluştur (Bulut için hayat kurtarıcı)
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
    # SAYFA 1: VERİ YÖNETİMİ (EXCEL İŞLEMLERİ)
    # =========================================================================
    if menu == "💾 Veri Girişi ve Düzenleme":
        st.title("💾 Sistem Veri Yönetimi")
        st.info("💡 **ÖNEMLİ BİLGİ:** Bulut sistemlerde veriler geçici sunucularda tutulur. Veri girişinizi tamamladıktan sonra mutlaka **'Güncel Verileri İndir'** butonuna basarak kopyanızı bilgisayarınıza kaydedin.")
        
        # İndirme ve Yükleme Butonları
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
            with open("veriler.xlsx", "rb") as f:
                st.download_button(label="📥 Güncel Verileri İndir (Excel)", data=f, file_name="guncel_veriler.xlsx", mime="application/vnd.ms-excel", use_container_width=True)

        st.divider()
        
        # Veri Düzenleme Tabloları (Arayüzde canlı düzenleme)
        st.subheader("📝 Sistem Üzerinden Doğrudan Veri Girişi")
        st.write("Aşağıdaki tablolara çift tıklayarak verileri değiştirebilir, en alta inerek yeni satır ekleyebilirsiniz.")
        
        tab1, tab2, tab3 = st.tabs(["1️⃣ Lisansüstü Tezler", "2️⃣ Kişisel Performanslar", "3️⃣ E-Posta Adresleri"])
        
        with tab1:
            yeni_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)
        with tab2:
            yeni_kisisel = st.data_editor(kisisel_df, num_rows="dynamic", use_container_width=True)
        with tab3:
            yeni_mailler = st.data_editor(mailler_df, num_rows="dynamic", use_container_width=True)

        if st.button("💾 Tablodaki Değişiklikleri Sisteme Kaydet", type="primary", use_container_width=True):
            with pd.ExcelWriter("veriler.xlsx", engine="openpyxl") as writer:
                yeni_df.to_excel(writer, sheet_name="Sheet1", index=False)
                yeni_mailler.to_excel(writer, sheet_name="Mailler", index=False)
                yeni_kisisel.to_excel(writer, sheet_name="Kişisel Performans", index=False)
            st.success("✅ Veriler sisteme başarıyla kaydedildi. 'Performans Paneli'ne geçebilirsiniz.")

    # =========================================================================
    # SAYFA 2: PERFORMANS PANELİ (BÜTÜNLEŞİK SİSTEM)
    # =========================================================================
    elif menu == "📊 Performans Paneli":
        st.title("Nevşehir Hacı Bektaş Veli Üniversitesi")
        st.subheader("Turizm Araştırmaları Enstitüsü - Bütünleşik Akademik Performans Sistemi")

        # Veri Temizleme İşlemleri
        df["Danışman Adı ve Soyadı"] = df["Danışman Adı ve Soyadı"].astype(str).str.strip()
        if "Danışmanın Bağlı Olduğu ABD" not in df.columns:
            df["Danışmanın Bağlı Olduğu ABD"] = "Belirtilmemiş"
        df["Danışmanın Bağlı Olduğu ABD"] = df["Danışmanın Bağlı Olduğu ABD"].fillna("Belirtilmemiş")
        mailler_df["Danışman Adı ve Soyadı"] = mailler_df["Danışman Adı ve Soyadı"].astype(str).str.strip()
        kisisel_df["Danışman Adı ve Soyadı"] = kisisel_df["Danışman Adı ve Soyadı"].astype(str).str.strip()
        
        df["Eser Türü"] = df["Eser Türü"].fillna("Yayın Yok")
        df["Eserin Indeksi"] = df["Eserin Indeksi"].fillna("İndeks Yok")
        kisisel_df["Eser Türü"] = kisisel_df.get("Eser Türü", pd.Series(dtype=str)).fillna("Yayın Yok")
        kisisel_df["Eserin Indeksi"] = kisisel_df.get("Eserin Indeksi", pd.Series(dtype=str)).fillna("İndeks Yok")
        
        hoca_abd_map = df.drop_duplicates(subset=["Danışman Adı ve Soyadı"])[["Danışman Adı ve Soyadı", "Kayıt olunan Anabilim dalı"]]

        # --- 1. LİSANSÜSTÜ PUANLAMA MOTORU ---
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

        # --- 2. KİŞİSEL PUANLAMA MOTORU (NEVÜ EK-1 YÖNERGESİ) ---
        def puan_hesapla_kisisel(row):
            if "Puan" in row and pd.notna(row["Puan"]):
                try: return float(row["Puan"])
                except: pass
            try:
                t = str(row.get("Eser Türü", "")).upper()
                i = str(row.get("Eserin Indeksi", "")).upper().replace(" ", "")
                if t == "YAYIN YOK" or t == "NAN" or t == "": return 0.0
                puan = 0.0
                if "MAKALE" in t:
                    if "Q1" in i: puan = 30.0
                    elif "Q2" in i: puan = 20.0
                    elif "Q3" in i: puan = 15.0
                    elif "Q4" in i or "TRDIZIN" in i or "TRDİZİN" in i: puan = 10.0
                    elif "AHCI" in i: puan = 20.0
                    elif "ESCI" in i or "SCOPUS" in i: puan = 10.0
                    else: puan = 5.0
                elif "KİTAP" in t or "KITAP" in t:
                    if "BÖLÜM" in t or "BOLUM" in t:
                        puan = 10.0 if "BKCI" in i else 5.0
                    elif "ÇEVİRİ" in t or "CEVIRI" in t: puan = 5.0
                    else:
                        puan = 20.0 if "BKCI" in i else 5.0
                elif "BİLDİRİ" in t or "BILDIRI" in t or "TOPLANTI" in t or "KONGRE" in t:
                    puan = 5.0 if "CPCI" in i else 3.0
                elif "ATIF" in t:
                    if "SCI" in i or "SSCI" in i: puan = 3.0
                    elif "BKCI" in i or "TR" in i: puan = 2.0
                    else: puan = 1.0
                elif "HAKEMLİK" in t or "HAKEMLIK" in t or "EDİTÖR" in t or "EDITOR" in t:
                    puan = 2.0
                elif "PATENT" in t:
                    puan = 20.0 if "ULUSLARARASI" in i else 10.0
                elif "ÖDÜL" in t or "ODUL" in t:
                    puan = 25.0
                elif "PROJE" in t:
                    puan = 15.0 if "1001" in i or "AB" in i else 10.0
                return puan
            except: return 0.0

        df["Eser Puanı"] = df.apply(puan_hesapla_lisansustu, axis=1)

        if not kisisel_df.empty:
            kisisel_df["Hesaplanan_Puan"] = kisisel_df.apply(puan_hesapla_kisisel, axis=1)
            kisisel_puan_toplam = kisisel_df.groupby("Danışman Adı ve Soyadı")["Hesaplanan_Puan"].sum().reset_index()
            kisisel_puan_toplam["Kişisel İndeks Puanı"] = kisisel_puan_toplam["Hesaplanan_Puan"] / 100.0
            kisisel_puan_toplam = kisisel_puan_toplam[["Danışman Adı ve Soyadı", "Kişisel İndeks Puanı"]]
        else:
            kisisel_puan_toplam = pd.DataFrame(columns=["Danışman Adı ve Soyadı", "Kişisel İndeks Puanı"])

        # FİLTRELEME
        st.sidebar.markdown("---")
        secilen_anabilim = st.sidebar.multiselect("Anabilim Dalı:", options=df["Kayıt olunan Anabilim dalı"].dropna().unique(), default=df["Kayıt olunan Anabilim dalı"].dropna().unique())
        secilen_program = st.sidebar.multiselect("Program Seçiniz:", options=df["Program"].dropna().unique(), default=df["Program"].dropna().unique())
        secilen_turler = st.sidebar.multiselect("Eser Türü:", options=df["Eser Türü"].unique(), default=df["Eser Türü"].unique())
        secilen_indeksler = st.sidebar.multiselect("İndeks:", options=df["Eserin Indeksi"].unique(), default=df["Eserin Indeksi"].unique())
        
        filtrelenmis_df = df[
            (df["Kayıt olunan Anabilim dalı"].isin(secilen_anabilim)) &
            (df["Program"].isin(secilen_program)) &
            (df["Eser Türü"].isin(secilen_turler)) & 
            (df["Eserin Indeksi"].isin(secilen_indeksler))
        ]
        
        # --- ARKA PLAN HESAPLAMALARI ---
        anabilim_data = filtrelenmis_df.groupby("Kayıt olunan Anabilim dalı")["Tez Adı"].nunique().reset_index(name="Sayı")
        anabilim_data.sort_values(by="Sayı", ascending=False, inplace=True)
        anabilim_data.columns = ["Anabilim Dalı", "Sayı"]
        
        program_data = filtrelenmis_df.groupby("Program")["Tez Adı"].nunique().reset_index(name="Sayı")
        program_data.sort_values(by="Sayı", ascending=False, inplace=True)
        program_data.columns = ["Program", "Sayı"]

        gercek_eserler_df = filtrelenmis_df[filtrelenmis_df["Eser Türü"] != "Yayın Yok"]

        ssci_hocalar_global = filtrelenmis_df[filtrelenmis_df["Eserin Indeksi"].str.contains("SSCI|SCI", case=False, na=False)]["Danışman Adı ve Soyadı"].unique()

        def ssci_renklendir(row):
            if row["Danışman Adı ve Soyadı"] in ssci_hocalar_global:
                return ['background-color: rgba(46, 204, 113, 0.25)'] * len(row)
            return [''] * len(row)

        # BÜTÜNLEŞİK PUAN HESAPLAMASI
        yl_df = filtrelenmis_df[filtrelenmis_df["Program"].str.contains("yüksek|lisans|yl", case=False, na=False)]
        dr_df = filtrelenmis_df[filtrelenmis_df["Program"].str.contains("doktora|dr|phd", case=False, na=False)]

        def alt_puan_hesapla(temp_df):
            t_puan = temp_df.groupby("Danışman Adı ve Soyadı")["Tez Adı"].nunique() * 0.05
            e_puan = temp_df.groupby("Danışman Adı ve Soyadı")["Eser Puanı"].sum()
            res = pd.DataFrame({"Tez İndeksi": t_puan, "Eser İndeksi": e_puan}).fillna(0)
            res["Toplam İndeks"] = res["Tez İndeksi"] + res["Eser İndeksi"]
            return res

        yl_puan = alt_puan_hesapla(yl_df).rename(columns={"Tez İndeksi": "YL Tez İndeksi", "Eser İndeksi": "YL Eser İndeksi", "Toplam İndeks": "YL Toplam İndeks"})
        dr_puan = alt_puan_hesapla(dr_df).rename(columns={"Tez İndeksi": "DR Tez İndeksi", "Eser İndeksi": "DR Eser İndeksi", "Toplam İndeks": "DR Toplam İndeks"})

        puan_tablosu = pd.DataFrame(index=filtrelenmis_df["Danışman Adı ve Soyadı"].unique())
        puan_tablosu = puan_tablosu.join(yl_puan).join(dr_puan).fillna(0)
        puan_tablosu["Lisansüstü İndeks Puanı"] = puan_tablosu["YL Toplam İndeks"] + puan_tablosu["DR Toplam İndeks"]
        puan_tablosu.reset_index(inplace=True)
        puan_tablosu.rename(columns={"index": "Danışman Adı ve Soyadı"}, inplace=True)
        
        puan_tablosu = pd.merge(puan_tablosu, kisisel_puan_toplam, on="Danışman Adı ve Soyadı", how="left")
        puan_tablosu["Kişisel İndeks Puanı"] = puan_tablosu["Kişisel İndeks Puanı"].fillna(0.0)
        puan_tablosu["Akademisyen İndeks Puanı"] = puan_tablosu["Lisansüstü İndeks Puanı"] + puan_tablosu["Kişisel İndeks Puanı"]
        puan_tablosu = pd.merge(puan_tablosu, hoca_abd_map, on="Danışman Adı ve Soyadı", how="left")
        puan_tablosu.sort_values("Akademisyen İndeks Puanı", ascending=False, inplace=True)

        st.divider()

        # --- GÖRSEL ARAYÜZ (DASHBOARD) ---
        with st.expander("📊 1️⃣ Genel Dağılım Grafikleri", expanded=True):
            col1, col2 = st.columns(2)
            with col1:
                st.write("📌 **Anabilim Dallarına Göre Bitirilen Tez Sayısı**")
                fig1 = px.bar(anabilim_data, x="Anabilim Dalı", y="Sayı", text="Sayı", color="Anabilim Dalı")
                fig1.update_traces(textposition='inside', insidetextanchor='middle', textfont_size=22, texttemplate='<b>%{text}</b>', textfont_color="white")
                st.plotly_chart(fig1, use_container_width=True)
            with col2:
                st.write("📌 **Programa Göre Bitirilen Tez Sayısı**")
                fig2 = px.bar(program_data, x="Program", y="Sayı", text="Sayı", color="Program")
                fig2.update_traces(textposition='inside', insidetextanchor='middle', textfont_size=22, texttemplate='<b>%{text}</b>', textfont_color="white")
                st.plotly_chart(fig2, use_container_width=True)

        with st.expander("🎯 2️⃣ Bütünleşik Akademisyen İndeks Tablosu (Enstitü Geneli)", expanded=True):
            st.info("💡 **Bilgi:** Öğretim üyelerinin 'Kişisel İndeks' puanları **NEVÜ EK-1 Yönergesindeki resmi katsayılara** göre analiz edilip hesaplanmıştır.")
            gosterim_puan = puan_tablosu[["Danışman Adı ve Soyadı", "Kayıt olunan Anabilim dalı", "Lisansüstü İndeks Puanı", "Kişisel İndeks Puanı", "Akademisyen İndeks Puanı"]]
            st.dataframe(gosterim_puan.style.format({"Lisansüstü İndeks Puanı": "{:.2f}", "Kişisel İndeks Puanı": "{:.2f}", "Akademisyen İndeks Puanı": "{:.2f}"}), use_container_width=True)

        with st.expander("🎯 3️⃣ Anabilim Dalı Özelinde İndeks Sıralamaları (Bölüm İçi)", expanded=True):
            abd_listesi = sorted(puan_tablosu["Kayıt olunan Anabilim dalı"].dropna().unique())
            for abd in abd_listesi:
                st.write(f"📌 **{abd} Anabilim Dalı Atama Sıralaması**")
                temp_abd = puan_tablosu[puan_tablosu["Kayıt olunan Anabilim dalı"] == abd].copy()
                temp_abd.sort_values("Akademisyen İndeks Puanı", ascending=False, inplace=True)
                st.dataframe(temp_abd[["Danışman Adı ve Soyadı", "Lisansüstü İndeks Puanı", "Kişisel İndeks Puanı", "Akademisyen İndeks Puanı"]].style.format({"Lisansüstü İndeks Puanı": "{:.2f}", "Kişisel İndeks Puanı": "{:.2f}", "Akademisyen İndeks Puanı": "{:.2f}"}), use_container_width=True)

        with st.expander("📩 4️⃣ Otomatik E-Posta Gönderim Paneli", expanded=False):
            st.markdown("Şifreler artık **Streamlit Secrets** üzerinden güvenle alınmaktadır.")
            if st.button("🚀 HTML Performans Bültenlerini Gönder", use_container_width=True):
                try:
                    gonderici_mail = st.secrets["email"]["gonderici_mail"]
                    gonderici_sifre = st.secrets["email"]["gonderici_sifre"]
                    
                    hedef_kitle = puan_tablosu[puan_tablosu["Akademisyen İndeks Puanı"] > 0]
                    if len(hedef_kitle) == 0:
                        st.warning("Gönderilecek hoca bulunamadı.")
                    else:
                        my_bar = st.progress(0, text="Gönderiliyor...")
                        basarili = 0
                        server = smtplib.SMTP('smtp.gmail.com', 587)
                        server.starttls()
                        server.login(gonderici_mail, gonderici_sifre)
                        
                        toplam = len(hedef_kitle)
                        for i, (index, row) in enumerate(hedef_kitle.iterrows()):
                            hoca_adi = row["Danışman Adı ve Soyadı"]
                            alici_mail = row["E-Posta Adresi"] if "E-Posta Adresi" in row and pd.notna(row["E-Posta Adresi"]) else ""
                            if alici_mail.strip() == "": continue
                                
                            perf = row["Akademisyen İndeks Puanı"]
                            html_mesaj = f"""
                            <html><body>
                            <h2>Turizm Araştırmaları Enstitüsü</h2>
                            <p>Sayın {hoca_adi}, Akademisyen İndeks Puanınız <b>{perf:.2f}</b> olarak hesaplanmıştır.</p>
                            </body></html>
                            """
                            msg = MIMEMultipart('alternative')
                            msg['From'] = gonderici_mail
                            msg['To'] = alici_mail
                            msg['Subject'] = "📊 Akademik Performans Bülteni"
                            msg.attach(MIMEText(html_mesaj, 'html', 'utf-8'))
                            server.send_message(msg)
                            basarili += 1
                            time.sleep(0.5)
                            my_bar.progress((i + 1) / toplam, text=f"Gönderiliyor... {hoca_adi}")
                        server.quit()
                        my_bar.empty()
                        st.success(f"✅ {basarili} kişiye mail gönderildi!")
                except Exception as e:
                    st.error(f"❌ Mail gönderilemedi. Hata veya Secrets eksik: {e}")

except Exception as e:
    st.error(f"⚠️ Sistemsel bir hata oluştu: {e}")
