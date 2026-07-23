import streamlit as st
import pandas as pd
import plotly.express as px
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time
from xhtml2pdf import pisa
from io import BytesIO

# 1. Sayfa Ayarları
st.set_page_config(page_title="Performans Takip Sistemi", layout="wide")
st.title("Nevşehir Hacı Bektaş Veli Üniversitesi")
st.subheader("Turizm Araştırmaları Enstitüsü - Bütünleşik Akademik Performans Sistemi")

try:
    # 2. Verileri Yükleme
    df = pd.read_excel("veriler.xlsx", sheet_name=0) 
    df["Danışman Adı ve Soyadı"] = df["Danışman Adı ve Soyadı"].astype(str).str.strip()
    
    if "Danışmanın Bağlı Olduğu ABD" not in df.columns:
        df["Danışmanın Bağlı Olduğu ABD"] = "Belirtilmemiş"
    df["Danışmanın Bağlı Olduğu ABD"] = df["Danışmanın Bağlı Olduğu ABD"].fillna("Belirtilmemiş")
    
    try:
        mailler_df = pd.read_excel("veriler.xlsx", sheet_name="Mailler")
        if len(mailler_df.columns) >= 2:
            mailler_df = mailler_df.iloc[:, 0:2]
            mailler_df.columns = ["Danışman Adı ve Soyadı", "E-Posta Adresi"]
            mailler_df["Danışman Adı ve Soyadı"] = mailler_df["Danışman Adı ve Soyadı"].astype(str).str.strip()
        else:
            mailler_df = pd.DataFrame(columns=["Danışman Adı ve Soyadı", "E-Posta Adresi"])
    except:
        mailler_df = pd.DataFrame(columns=["Danışman Adı ve Soyadı", "E-Posta Adresi"])

    try:
        kisisel_df = pd.read_excel("veriler.xlsx", sheet_name="Kişisel Performans")
        kisisel_df["Danışman Adı ve Soyadı"] = kisisel_df["Danışman Adı ve Soyadı"].astype(str).str.strip()
    except:
        kisisel_df = pd.DataFrame(columns=["Danışman Adı ve Soyadı", "Eser Türü", "Eserin Indeksi"])

    df["Eser Türü"] = df["Eser Türü"].fillna("Yayın Yok")
    df["Eserin Indeksi"] = df["Eserin Indeksi"].fillna("İndeks Yok")
    kisisel_df["Eser Türü"] = kisisel_df.get("Eser Türü", pd.Series(dtype=str)).fillna("Yayın Yok")
    kisisel_df["Eserin Indeksi"] = kisisel_df.get("Eserin Indeksi", pd.Series(dtype=str)).fillna("İndeks Yok")
    
    hoca_abd_map = df.drop_duplicates(subset=["Danışman Adı ve Soyadı"])[["Danışman Adı ve Soyadı", "Kayıt olunan Anabilim dalı"]]
    
    # --- 1. LİSANSÜSTÜ PUANLAMA MOTORU (ESKİ SİSTEM MANTIGI) ---
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
        # Eğer direkt Excel'e 'Puan' isimli bir sütun açılıp elle puan girildiyse direkt onu al
        if "Puan" in row and pd.notna(row["Puan"]):
            try: return float(row["Puan"])
            except: pass
            
        try:
            t = str(row.get("Eser Türü", "")).upper()
            i = str(row.get("Eserin Indeksi", "")).upper().replace(" ", "")
            if t == "YAYIN YOK" or t == "NAN" or t == "": return 0.0
            
            puan = 0.0
            # A. YAYINLAR
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
            # B. DERGİ HAKEMLİĞİ VE EDİTÖRLÜK
            elif "HAKEMLİK" in t or "HAKEMLIK" in t or "EDİTÖR" in t or "EDITOR" in t:
                puan = 2.0
            # C. DİĞER (ÖDÜL, PATENT, PROJE)
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

    # --- SOL YAN MENÜ FİLTRELERİ ---
    st.sidebar.image("https://nevsehir.edu.tr/assets/images/logo.png", width=150)
    st.sidebar.header("🔍 Filtreleme Menüsü")
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
    makale_df = gercek_eserler_df[gercek_eserler_df["Eser Türü"].str.contains("Makale", case=False, na=False)]

    tez_sayilari = filtrelenmis_df.groupby("Danışman Adı ve Soyadı")["Tez Adı"].nunique().reset_index(name="Bitirilen Tez Sayısı")
    yayin_sayilari = gercek_eserler_df.groupby("Danışman Adı ve Soyadı").size().reset_index(name="Toplam Üretilen Eser")
    makale_sayilari = makale_df.groupby("Danışman Adı ve Soyadı").size().reset_index(name="Makale Sayısı")
    
    verimlilik_df = pd.merge(tez_sayilari, yayin_sayilari, on="Danışman Adı ve Soyadı", how="left")
    verimlilik_df = pd.merge(verimlilik_df, makale_sayilari, on="Danışman Adı ve Soyadı", how="left").fillna(0)
    verimlilik_df["Makale Verim Oranı (%)"] = (verimlilik_df["Makale Sayısı"] / verimlilik_df["Bitirilen Tez Sayısı"]) * 100
    verimlilik_df.fillna(0, inplace=True)
    verimlilik_df = pd.merge(verimlilik_df, mailler_df, on="Danışman Adı ve Soyadı", how="left")
    verimlilik_df.sort_values(by="Makale Verim Oranı (%)", ascending=False, inplace=True)
    verimlilik_df.reset_index(drop=True, inplace=True)

    ssci_hocalar_global = filtrelenmis_df[filtrelenmis_df["Eserin Indeksi"].str.contains("SSCI|SCI", case=False, na=False)]["Danışman Adı ve Soyadı"].unique()

    def ssci_renklendir(row):
        if row["Danışman Adı ve Soyadı"] in ssci_hocalar_global:
            return ['background-color: rgba(46, 204, 113, 0.25)'] * len(row)
        return [''] * len(row)
    
    basarili_hocalar_df = verimlilik_df[verimlilik_df["Makale Verim Oranı (%)"] > 0]
    makalesiz_hocalar = verimlilik_df[(verimlilik_df["Bitirilen Tez Sayısı"] > 0) & (verimlilik_df["Makale Sayısı"] == 0)].copy()
    makalesiz_hocalar.sort_values(by="Bitirilen Tez Sayısı", ascending=False, inplace=True)
    makalesiz_hocalar.reset_index(drop=True, inplace=True)

    ad_tez = filtrelenmis_df.groupby("Kayıt olunan Anabilim dalı")["Tez Adı"].nunique().reset_index(name="Bitirilen Tez Sayısı")
    ad_yayin = gercek_eserler_df.groupby("Kayıt olunan Anabilim dalı").size().reset_index(name="Toplam Üretilen Eser")
    ad_makale = makale_df.groupby("Kayıt olunan Anabilim dalı").size().reset_index(name="Makale Sayısı")
    
    ad_verimlilik = pd.merge(ad_tez, ad_yayin, on="Kayıt olunan Anabilim dalı", how="left")
    ad_verimlilik = pd.merge(ad_verimlilik, ad_makale, on="Kayıt olunan Anabilim dalı", how="left").fillna(0)
    ad_verimlilik["Makale Verim Oranı (%)"] = (ad_verimlilik["Makale Sayısı"] / ad_verimlilik["Bitirilen Tez Sayısı"]) * 100
    ad_verimlilik.fillna(0, inplace=True)
    ad_verimlilik.sort_values(by="Makale Verim Oranı (%)", ascending=False, inplace=True)
    ad_verimlilik.reset_index(drop=True, inplace=True)

    # --- BÜTÜNLEŞİK PUAN HESAPLAMASI (LİSANSÜSTÜ + KİŞİSEL) ---
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
    
    for col in ["YL Tez İndeksi", "YL Eser İndeksi", "YL Toplam İndeks", "DR Tez İndeksi", "DR Eser İndeksi", "DR Toplam İndeks", "Lisansüstü İndeks Puanı", "Kişisel İndeks Puanı", "Akademisyen İndeks Puanı"]:
        puan_tablosu[col] = puan_tablosu[col].astype(float)

    puan_tablosu.sort_values("Akademisyen İndeks Puanı", ascending=False, inplace=True)

    st.divider()

    # --- BÖLÜM 1: GENEL DAĞILIM GRAFİKLERİ ---
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

    st.write("")

    # --- BÖLÜM 2: ANABİLİM DALI PERFORMANS ---
    with st.expander("🏛️ 2️⃣ Anabilim Dalı Performans Analizi", expanded=True):
        st.dataframe(ad_verimlilik.style.format({"Bitirilen Tez Sayısı": "{:.0f}", "Toplam Üretilen Eser": "{:.0f}", "Makale Sayısı": "{:.0f}", "Makale Verim Oranı (%)": "% {:.0f}"}), use_container_width=True)

    st.write("")

    # --- BÖLÜM 3: DANIŞMAN ORANLARI ---
    with st.expander("🏆 3️⃣ Danışman Makale Verim Oranları", expanded=True):
        gosterilecek_basarili = basarili_hocalar_df.drop(columns=["E-Posta Adresi"], errors='ignore')
        st.dataframe(gosterilecek_basarili.style.apply(ssci_renklendir, axis=1).format({"Bitirilen Tez Sayısı": "{:.0f}", "Toplam Üretilen Eser": "{:.0f}", "Makale Sayısı": "{:.0f}", "Makale Verim Oranı (%)": "% {:.0f}"}), use_container_width=True)

    st.write("")

    # --- BÖLÜM 4: BÜTÜNLEŞİK AKADEMİSYEN İNDEKS TABLOSU ---
    with st.expander("🎯 4️⃣ Bütünleşik Akademisyen İndeks Tablosu (Enstitü Geneli)", expanded=True):
        st.info("💡 **Bilgi:** Öğretim üyelerinin 'Kişisel İndeks' puanları **NEVÜ EK-1 Yönergesindeki resmi katsayılara** göre analiz edilip hesaplanmıştır. Excell'de Puan sütunu açarsanız o değerler alınır.")
        
        gosterim_puan = puan_tablosu[["Danışman Adı ve Soyadı", "Kayıt olunan Anabilim dalı", "Lisansüstü İndeks Puanı", "Kişisel İndeks Puanı", "Akademisyen İndeks Puanı"]]
        st.dataframe(gosterim_puan.style.format({"Lisansüstü İndeks Puanı": "{:.2f}", "Kişisel İndeks Puanı": "{:.2f}", "Akademisyen İndeks Puanı": "{:.2f}"}), use_container_width=True)

    st.write("")

    # --- BÖLÜM 5: ANABİLİM DALI ÖZELİNDE İNDEKS SIRALAMALARI ---
    with st.expander("🎯 5️⃣ Anabilim Dalı Özelinde Akademisyen İndeks Sıralaması (Bölüm İçi)", expanded=True):
        abd_listesi = sorted(puan_tablosu["Kayıt olunan Anabilim dalı"].dropna().unique())
        for abd in abd_listesi:
            st.write(f"📌 **{abd} Anabilim Dalı Atama Sıralaması**")
            temp_abd = puan_tablosu[puan_tablosu["Kayıt olunan Anabilim dalı"] == abd].copy()
            temp_abd.sort_values("Akademisyen İndeks Puanı", ascending=False, inplace=True)
            st.dataframe(temp_abd[["Danışman Adı ve Soyadı", "Lisansüstü İndeks Puanı", "Kişisel İndeks Puanı", "Akademisyen İndeks Puanı"]].style.format({"Lisansüstü İndeks Puanı": "{:.2f}", "Kişisel İndeks Puanı": "{:.2f}", "Akademisyen İndeks Puanı": "{:.2f}"}), use_container_width=True)

    st.write("")

    # --- BÖLÜM 6: SIFIR ÇEKENLER ---
    with st.expander("⚠️ 6️⃣ Tez Bitirip Makale Ürettirmeyen Danışmanlar", expanded=False):
        if len(makalesiz_hocalar) > 0:
            st.warning(f"Sistemde tez bitirmiş olmasına rağmen hiç 'Makale' çıkartmamış {len(makalesiz_hocalar)} adet danışman tespit edilmiştir.")
            st.dataframe(makalesiz_hocalar[["Danışman Adı ve Soyadı", "Bitirilen Tez Sayısı", "Toplam Üretilen Eser", "Makale Verim Oranı (%)"]].style.format({"Bitirilen Tez Sayısı": "{:.0f}", "Toplam Üretilen Eser": "{:.0f}", "Makale Verim Oranı (%)": "% {:.0f}"}), use_container_width=True)
        else:
            st.success("✅ Harika! Tüm danışmanlar makale üretmiştir.")

    st.write("")

    # --- BÖLÜM 7: OTOMATİK E-POSTA GÖNDERİMİ ---
    with st.expander("📩 7️⃣ Otomatik E-Posta Gönderim Paneli", expanded=False):
        col_mail1, col_mail2 = st.columns(2)
        with col_mail1:
            gonderici_mail = st.text_input("Enstitü Gönderici E-Posta (Gmail)", placeholder="ornek@gmail.com")
        with col_mail2:
            gonderici_sifre = st.text_input("Uygulama Şifresi (16 Haneli)", type="password", placeholder="abcd efgh ijkl mnop")
        if st.button("🚀 HTML Performans Bültenlerini Gönder", use_container_width=True):
            st.success("Mail sistemi hazırdır.")

    st.write("")

    # --- BÖLÜM 8: PDF RAPOR OLUŞTURUCU ---
    with st.expander("📑 8️⃣ PDF / Detaylı Rapor İndirici", expanded=True):
        st.markdown("Yönetime sunulmak üzere resmi rapor hazırlayabilirsiniz.")

        def alt_rapor_uret(temp_df):
            if temp_df.empty: return "<p>Veri bulunmamaktadır.</p>", "<p>Veri yok.</p>"
            local_ssci = temp_df[temp_df["Eserin Indeksi"].str.contains("SSCI|SCI", case=False, na=False)]["Danışman Adı ve Soyadı"].unique()
            t_tez = temp_df.groupby("Kayıt olunan Anabilim dalı")["Tez Adı"].nunique().reset_index(name="Biten Tez")
            t_eser = temp_df[temp_df["Eser Türü"] != "Yayın Yok"].groupby("Kayıt olunan Anabilim dalı").size().reset_index(name="Eser Sayısı")
            r_ad = pd.merge(t_tez, t_eser, on="Kayıt olunan Anabilim dalı", how="left").fillna(0)
            r_ad["Makale Verim Oranı (%)"] = (r_ad["Eser Sayısı"] / r_ad["Biten Tez"]) * 100
            r_ad.fillna(0, inplace=True)
            r_ad.sort_values("Makale Verim Oranı (%)", ascending=False, inplace=True)
            r_ad["Makale Verim Oranı (%)"] = r_ad["Makale Verim Oranı (%)"].apply(lambda x: f"% {x:.0f}")
            r_ad["Biten Tez"] = r_ad["Biten Tez"].astype(int)
            r_ad["Eser Sayısı"] = r_ad["Eser Sayısı"].astype(int)
            
            html_ad = "<table class='mystyle' style='width:100%; border-collapse: collapse; font-size:11px; margin-top:5px;'>"
            html_ad += "<tr><th style='background-color: #34495E; color: white; padding: 6px; text-align:left;'>Anabilim Dalı</th><th style='background-color: #34495E; color: white; padding: 6px; text-align:center;'>Biten Tez</th><th style='background-color: #34495E; color: white; padding: 6px; text-align:center;'>Eser Sayısı</th><th style='background-color: #34495E; color: white; padding: 6px; text-align:center;'>Makale Verim Oranı (%)</th></tr>"
            row_count = 0
            for _, row in r_ad.iterrows():
                bg = "#F8F9F9" if row_count % 2 == 0 else "#FFFFFF"
                html_ad += f"<tr style='background-color: {bg};'><td style='border: 1px solid #ddd; padding: 6px;'>{row['Kayıt olunan Anabilim dalı']}</td><td style='border: 1px solid #ddd; padding: 6px; text-align:center;'>{row['Biten Tez']}</td><td style='border: 1px solid #ddd; padding: 6px; text-align:center;'>{row['Eser Sayısı']}</td><td style='border: 1px solid #ddd; padding: 6px; text-align:center;'><b>{row['Makale Verim Oranı (%)']}</b></td></tr>"
                row_count += 1
            html_ad += "</table>"

            h_tez = temp_df.groupby("Danışman Adı ve Soyadı")["Tez Adı"].nunique().reset_index(name="Biten Tez")
            h_eser = temp_df[temp_df["Eser Türü"] != "Yayın Yok"].groupby("Danışman Adı ve Soyadı").size().reset_index(name="Eser Sayısı")
            r_hoca = pd.merge(h_tez, h_eser, on="Danışman Adı ve Soyadı", how="left").fillna(0)
            r_hoca["Makale Verim Oranı (%)"] = (r_hoca["Eser Sayısı"] / r_hoca["Biten Tez"]) * 100
            r_hoca.fillna(0, inplace=True)
            r_hoca = r_hoca[r_hoca["Makale Verim Oranı (%)"] > 0] 
            r_hoca.sort_values("Makale Verim Oranı (%)", ascending=False, inplace=True)
            r_hoca["Makale Verim Oranı (%)"] = r_hoca["Makale Verim Oranı (%)"].apply(lambda x: f"% {x:.0f}")
            r_hoca["Biten Tez"] = r_hoca["Biten Tez"].astype(int)
            r_hoca["Eser Sayısı"] = r_hoca["Eser Sayısı"].astype(int)
            
            if r_hoca.empty:
                html_hoca = "<p>Bu program türünde henüz eser üreten danışman bulunmamaktadır.</p>"
            else:
                html_hoca = "<table class='mystyle' style='width:100%; border-collapse: collapse; font-size:11px; margin-top:5px;'>"
                html_hoca += "<tr><th style='background-color: #34495E; color: white; padding: 6px; text-align:left;'>Danışman Adı ve Soyadı</th><th style='background-color: #34495E; color: white; padding: 6px; text-align:center;'>Biten Tez</th><th style='background-color: #34495E; color: white; padding: 6px; text-align:center;'>Eser Sayısı</th><th style='background-color: #34495E; color: white; padding: 6px; text-align:center;'>Makale Verim Oranı (%)</th></tr>"
                row_count = 0
                for _, row in r_hoca.head(10).iterrows():
                    bg_color = "#D5F5E3" if row["Danışman Adı ve Soyadı"] in local_ssci else ("#F8F9F9" if row_count % 2 == 0 else "#FFFFFF")
                    html_hoca += f"<tr style='background-color: {bg_color};'><td style='border: 1px solid #ddd; padding: 6px;'>{row['Danışman Adı ve Soyadı']}</td><td style='border: 1px solid #ddd; padding: 6px; text-align:center;'>{row['Biten Tez']}</td><td style='border: 1px solid #ddd; padding: 6px; text-align:center;'>{row['Eser Sayısı']}</td><td style='border: 1px solid #ddd; padding: 6px; text-align:center;'><b>{row['Makale Verim Oranı (%)']}</b></td></tr>"
                    row_count += 1
                html_hoca += "</table>"
                if len(local_ssci) > 0:
                    ssci_isimler = ", ".join(sorted(local_ssci))
                    html_hoca += f"<div style='margin-top:10px; font-size: 10px; color: #27AE60; background-color: #EAFAF1; padding: 8px; border-left: 3px solid #2ECC71;'><b>🌟 Kalite Vurgusu:</b> Açık yeşil renkle vurgulanan danışmanlar, bu programa ait tezlerden <b>SSCI veya SCI/SCI-E</b> indekslerinde eser ürettirmiş üstün başarılı öğretim üyeleridir:<br><i>{ssci_isimler}</i></div>"
            return html_ad, html_hoca

        def tam_liste_uret(temp_df, tez_sutun_adi):
            if temp_df.empty: return "<p>Veri bulunmamaktadır.</p>"
            local_ssci = temp_df[temp_df["Eserin Indeksi"].str.contains("SSCI|SCI", case=False, na=False)]["Danışman Adı ve Soyadı"].unique()
            h_tez = temp_df.groupby("Danışman Adı ve Soyadı")["Tez Adı"].nunique().reset_index(name=tez_sutun_adi)
            h_eser = temp_df[temp_df["Eser Türü"] != "Yayın Yok"].groupby("Danışman Adı ve Soyadı").size().reset_index(name="Üretilen Eser")
            
            res = pd.merge(h_tez, h_eser, on="Danışman Adı ve Soyadı", how="left").fillna(0)
            res["Makale Verim Oranı (%)"] = (res["Üretilen Eser"] / res[tez_sutun_adi]) * 100
            res.fillna(0, inplace=True)
            res.sort_values(by=["Makale Verim Oranı (%)", "Danışman Adı ve Soyadı"], ascending=[False, True], inplace=True)
            
            res[tez_sutun_adi] = res[tez_sutun_adi].astype(int)
            res["Üretilen Eser"] = res["Üretilen Eser"].astype(int)
            res["Makale Verim Oranı (%)"] = res["Makale Verim Oranı (%)"].apply(lambda x: f"% {x:.0f}")
            
            html_table = "<table class='mystyle' style='width:100%; border-collapse: collapse; font-size:11px; margin-top:5px;'>"
            html_table += f"<tr><th style='background-color: #34495E; color: white; padding: 6px; text-align:left;'>Danışman Adı ve Soyadı</th><th style='background-color: #34495E; color: white; padding: 6px; text-align:center;'>{tez_sutun_adi}</th><th style='background-color: #34495E; color: white; padding: 6px; text-align:center;'>Üretilen Eser</th><th style='background-color: #34495E; color: white; padding: 6px; text-align:center;'>Makale Verim Oranı (%)</th></tr>"
            row_count = 0
            for _, row in res.iterrows():
                bg_color = "#D5F5E3" if row["Danışman Adı ve Soyadı"] in local_ssci else ("#F8F9F9" if row_count % 2 == 0 else "#FFFFFF")
                html_table += f"<tr style='background-color: {bg_color};'><td style='border: 1px solid #ddd; padding: 6px;'>{row['Danışman Adı ve Soyadı']}</td><td style='border: 1px solid #ddd; padding: 6px; text-align:center;'>{row[tez_sutun_adi]}</td><td style='border: 1px solid #ddd; padding: 6px; text-align:center;'>{row['Üretilen Eser']}</td><td style='border: 1px solid #ddd; padding: 6px; text-align:center;'><b>{row['Makale Verim Oranı (%)']}</b></td></tr>"
                row_count += 1
            html_table += "</table>"
            if len(local_ssci) > 0:
                ssci_isimler = ", ".join(sorted(local_ssci))
                html_table += f"<div style='margin-top:10px; font-size: 10px; color: #27AE60; background-color: #EAFAF1; padding: 8px; border-left: 3px solid #2ECC71;'><b>🌟 Kalite Vurgusu:</b> Açık yeşil renkle vurgulanan danışmanlar, bu programa ait tezlerden <b>SSCI veya SCI/SCI-E</b> indekslerinde eser ürettirmiş üstün başarılı öğretim üyeleridir:<br><i>{ssci_isimler}</i></div>"
            return html_table

        yl_df = filtrelenmis_df[filtrelenmis_df["Program"].str.contains("yüksek|lisans|yl", case=False, na=False)]
        dr_df = filtrelenmis_df[filtrelenmis_df["Program"].str.contains("doktora|dr|phd", case=False, na=False)]

        yl_tez_say = yl_df.groupby("Kayıt olunan Anabilim dalı")["Tez Adı"].nunique().reset_index(name="YL Tezi")
        dr_tez_say = dr_df.groupby("Kayıt olunan Anabilim dalı")["Tez Adı"].nunique().reset_index(name="DR Tezi")
        yl_eser_say = yl_df[yl_df["Eser Türü"] != "Yayın Yok"].groupby("Kayıt olunan Anabilim dalı").size().reset_index(name="YL Toplam Eser")
        dr_eser_say = dr_df[dr_df["Eser Türü"] != "Yayın Yok"].groupby("Kayıt olunan Anabilim dalı").size().reset_index(name="DR Toplam Eser")
        tum_tez_say = filtrelenmis_df.groupby("Kayıt olunan Anabilim dalı")["Tez Adı"].nunique().reset_index(name="Toplam Tez")
        tum_eser_say = filtrelenmis_df[filtrelenmis_df["Eser Türü"] != "Yayın Yok"].groupby("Kayıt olunan Anabilim dalı").size().reset_index(name="Toplam Eser")
        
        genel_ozet = pd.merge(tum_tez_say, tum_eser_say, on="Kayıt olunan Anabilim dalı", how="left")
        genel_ozet = pd.merge(genel_ozet, yl_tez_say, on="Kayıt olunan Anabilim dalı", how="left")
        genel_ozet = pd.merge(genel_ozet, yl_eser_say, on="Kayıt olunan Anabilim dalı", how="left")
        genel_ozet = pd.merge(genel_ozet, dr_tez_say, on="Kayıt olunan Anabilim dalı", how="left")
        genel_ozet = pd.merge(genel_ozet, dr_eser_say, on="Kayıt olunan Anabilim dalı", how="left").fillna(0)
        
        genel_ozet["Makale Verim Oranı (%)"] = (genel_ozet["Toplam Eser"] / genel_ozet["Toplam Tez"]) * 100
        genel_ozet.fillna(0, inplace=True)
        genel_ozet.sort_values("Makale Verim Oranı (%)", ascending=False, inplace=True)
        
        for col in ["Toplam Tez", "Toplam Eser", "YL Tezi", "YL Toplam Eser", "DR Tezi", "DR Toplam Eser"]:
            genel_ozet[col] = genel_ozet[col].astype(int)
        genel_ozet["Makale Verim Oranı (%)"] = genel_ozet["Makale Verim Oranı (%)"].apply(lambda x: f"% {x:.0f}")
        
        html_genel_ozet = "<table class='mystyle' style='width:100%; border-collapse: collapse; font-size:11px; margin-top:5px;'>"
        html_genel_ozet += "<tr><th style='background-color: #34495E; color: white; padding: 6px; text-align:left;'>Anabilim Dalı</th>"
        for col in ["Toplam Tez", "Toplam Eser", "YL Tezi", "YL Toplam Eser", "DR Tezi", "DR Toplam Eser", "Makale Verim Oranı (%)"]:
            html_genel_ozet += f"<th style='background-color: #34495E; color: white; padding: 6px; text-align:center;'>{col}</th>"
        html_genel_ozet += "</tr>"
        
        row_count = 0
        for _, row in genel_ozet.iterrows():
            bg_default = "#F8F9F9" if row_count % 2 == 0 else "#FFFFFF"
            bg_toplam = "#EBF5FB"  
            bg_yl = "#FEF5E7"      
            bg_dr = "#FDEDEC"      
            
            html_genel_ozet += "<tr>"
            html_genel_ozet += f"<td style='background-color: {bg_default}; border: 1px solid #ddd; padding: 6px;'>{row['Kayıt olunan Anabilim dalı']}</td>"
            html_genel_ozet += f"<td style='background-color: {bg_toplam}; border: 1px solid #ddd; padding: 6px; text-align:center;'>{row['Toplam Tez']}</td>"
            html_genel_ozet += f"<td style='background-color: {bg_toplam}; border: 1px solid #ddd; padding: 6px; text-align:center;'>{row['Toplam Eser']}</td>"
            html_genel_ozet += f"<td style='background-color: {bg_yl}; border: 1px solid #ddd; padding: 6px; text-align:center;'>{row['YL Tezi']}</td>"
            html_genel_ozet += f"<td style='background-color: {bg_yl}; border: 1px solid #ddd; padding: 6px; text-align:center;'>{row['YL Toplam Eser']}</td>"
            html_genel_ozet += f"<td style='background-color: {bg_dr}; border: 1px solid #ddd; padding: 6px; text-align:center;'>{row['DR Tezi']}</td>"
            html_genel_ozet += f"<td style='background-color: {bg_dr}; border: 1px solid #ddd; padding: 6px; text-align:center;'>{row['DR Toplam Eser']}</td>"
            html_genel_ozet += f"<td style='background-color: {bg_default}; border: 1px solid #ddd; padding: 6px; text-align:center;'><b>{row['Makale Verim Oranı (%)']}</b></td>"
            html_genel_ozet += "</tr>"
            row_count += 1
        html_genel_ozet += "</table>"

        # BÜTÜNLEŞİK AKADEMİSYEN İNDEKS TABLOSU (PDF)
        html_puan = "<table class='mystyle' style='width:100%; border-collapse: collapse; font-size:11px; margin-top:5px;'>"
        html_puan += "<tr><th style='background-color: #34495E; color: white; padding: 6px; text-align:left;'>Danışman Adı ve Soyadı</th>"
        for col in ["Kayıt Olunan ABD", "Lisansüstü İndeks", "Kişisel İndeks", "Akademisyen İndeksi"]:
            html_puan += f"<th style='background-color: #34495E; color: white; padding: 6px; text-align:center;'>{col}</th>"
        html_puan += "</tr>"
        
        row_count = 0
        for _, row in puan_tablosu.iterrows():
            bg_default = "#F8F9F9" if row_count % 2 == 0 else "#FFFFFF"
            bg_genel = "#E8F8F5" 
            html_puan += "<tr>"
            html_puan += f"<td style='background-color: {bg_default}; border: 1px solid #ddd; padding: 6px;'>{row['Danışman Adı ve Soyadı']}</td>"
            html_puan += f"<td style='background-color: {bg_default}; border: 1px solid #ddd; padding: 6px; text-align:center;'>{row['Kayıt olunan Anabilim dalı']}</td>"
            html_puan += f"<td style='background-color: {bg_default}; border: 1px solid #ddd; padding: 6px; text-align:center;'>{row['Lisansüstü İndeks Puanı']:.2f}</td>"
            html_puan += f"<td style='background-color: {bg_default}; border: 1px solid #ddd; padding: 6px; text-align:center;'>{row['Kişisel İndeks Puanı']:.2f}</td>"
            html_puan += f"<td style='background-color: {bg_genel}; border: 1px solid #ddd; padding: 6px; text-align:center; font-size:12px; color:#145A32;'><b>{row['Akademisyen İndeks Puanı']:.2f}</b></td>"
            html_puan += "</tr>"
            row_count += 1
        html_puan += "</table>"

        # ANABİLİM DALI ÖZELİNDE İNDEKS (PDF)
        html_abd_puanlari = ""
        abd_listesi = sorted(puan_tablosu["Kayıt olunan Anabilim dalı"].dropna().unique())
        for abd in abd_listesi:
            temp_abd = puan_tablosu[puan_tablosu["Kayıt olunan Anabilim dalı"] == abd].copy()
            temp_abd.sort_values("Akademisyen İndeks Puanı", ascending=False, inplace=True)
            
            html_abd_puanlari += f"<h3>📌 {abd} Anabilim Dalı Atama Sıralaması</h3>"
            html_abd_puanlari += "<table class='mystyle' style='width:100%; border-collapse: collapse; font-size:11px; margin-top:5px; margin-bottom:15px;'>"
            html_abd_puanlari += "<tr><th style='background-color: #34495E; color: white; padding: 6px; text-align:left;'>Danışman Adı ve Soyadı</th><th style='background-color: #34495E; color: white; padding: 6px; text-align:center;'>Lisansüstü İndeks Puanı</th><th style='background-color: #34495E; color: white; padding: 6px; text-align:center;'>Kişisel İndeks Puanı</th><th style='background-color: #34495E; color: white; padding: 6px; text-align:center;'>Akademisyen İndeks Puanı</th></tr>"
            
            r_count = 0
            for _, r in temp_abd.iterrows():
                bg = "#F8F9F9" if r_count % 2 == 0 else "#FFFFFF"
                html_abd_puanlari += f"<tr style='background-color: {bg};'>"
                html_abd_puanlari += f"<td style='border: 1px solid #ddd; padding: 6px;'>{r['Danışman Adı ve Soyadı']}</td>"
                html_abd_puanlari += f"<td style='border: 1px solid #ddd; padding: 6px; text-align:center;'>{r['Lisansüstü İndeks Puanı']:.2f}</td>"
                html_abd_puanlari += f"<td style='border: 1px solid #ddd; padding: 6px; text-align:center;'>{r['Kişisel İndeks Puanı']:.2f}</td>"
                html_abd_puanlari += f"<td style='background-color: #E8F8F5; border: 1px solid #ddd; padding: 6px; text-align:center; color:#145A32;'><b>{r['Akademisyen İndeks Puanı']:.2f}</b></td>"
                html_abd_puanlari += "</tr>"
                r_count += 1
            html_abd_puanlari += "</table>"

        yl_ad_html, yl_hoca_html = alt_rapor_uret(yl_df)
        dr_ad_html, dr_hoca_html = alt_rapor_uret(dr_df)
        html_yl_tam_liste = tam_liste_uret(yl_df, "Biten YL Tezi")
        html_dr_tam_liste = tam_liste_uret(dr_df, "Biten DR Tezi")

        rapor_html = f"""
        <!DOCTYPE html>
        <html>
        <head>
        <meta charset="UTF-8">
        <style>
            body {{ font-family: Helvetica, sans-serif; color: #333; font-size: 11px; line-height: 1.4; }}
            h1 {{ text-align: center; color: #2C3E50; font-size: 16px; margin-bottom: 2px; }}
            h4 {{ text-align: center; color: #7F8C8D; font-size: 12px; margin-top: 0px; margin-bottom: 15px; font-weight: normal; }}
            h2 {{ color: #2980B9; border-bottom: 1px solid #2980B9; font-size: 14px; padding-bottom: 3px; margin-top: 20px; }}
            h3 {{ color: #34495E; font-size: 12px; margin-top: 10px; margin-bottom: 5px; }}
            .mystyle {{ border-collapse: collapse; width: 100%; margin-top: 5px; font-size: 11px; }}
            .mystyle th {{ background-color: #34495E; color: white; padding: 6px; border: 1px solid #34495E; }}
            .mystyle td {{ border: 1px solid #ddd; padding: 6px; }}
        </style>
        </head>
        <body>
            <h1>TURİZM ARAŞTIRMALARI ENSTİTÜSÜ</h1>
            <h4>Bütünleşik Akademik Performans Detaylı Raporu</h4>

            <h2>BÖLÜM 1: GENEL DURUM (TÜM PROGRAMLAR)</h2>
            {html_genel_ozet}

            <pdf:nextpage />
            <h2>BÖLÜM 2: YÜKSEK LİSANS ÖZEL RAPORU</h2>
            <h3>📌 Anabilim Dalı Yüksek Lisans Başarı Dağılımı</h3>
            {yl_ad_html}
            <h3>🏆 Yüksek Lisans Programında En Başarılı Danışmanlar (İlk 10)</h3>
            {yl_hoca_html}

            <pdf:nextpage />
            <h2>BÖLÜM 3: DOKTORA ÖZEL RAPORU</h2>
            <h3>📌 Anabilim Dalı Doktora Başarı Dağılımı</h3>
            {dr_ad_html}
            <h3>🏆 Doktora Programında En Başarılı Danışmanlar (İlk 10)</h3>
            {dr_hoca_html}
            
            <pdf:nextpage />
            <h2>BÖLÜM 4: TÜM DANIŞMANLARIN DETAYLI LİSTESİ</h2>
            <h3>📌 Yüksek Lisans Danışmanlık ve Eser Çıktıları</h3>
            {html_yl_tam_liste}
            
            <h3>📌 Doktora Danışmanlık ve Eser Çıktıları</h3>
            {html_dr_tam_liste}
            
            <pdf:nextpage />
            <h2>BÖLÜM 5: BÜTÜNLEŞİK AKADEMİSYEN İNDEKS TABLOSU (ENSTİTÜ GENELİ)</h2>
            <p style="font-size:10px; color:#555; margin-bottom:5px;"><i>* <b>Lisansüstü İndeks:</b> Danışmanın yürüttüğü tezlerden çıkarttığı yayınların puanıdır.<br>
            * <b>Kişisel İndeks:</b> Danışmanın tezlerden bağımsız kendi akademik yayınlarının (NEVÜ EK-1 katsayılarıyla) puanıdır.<br>
            * Her iki indeks puanı birleştirilerek nihai <b>Akademisyen İndeks Puanı</b> oluşturulmuştur. Atama sıralamalarında bu nihai puan referans alınır.</i></p>
            {html_puan}

            <pdf:nextpage />
            <h2>BÖLÜM 6: ANABİLİM DALI ÖZELİNDE İNDEKS SIRALAMALARI (BÖLÜM İÇİ REKABET)</h2>
            <p style="font-size:10px; color:#555; margin-bottom:5px;"><i>* <b>Bölüm İçi İndeks:</b> Bu bölümde her danışmanın performansı, sadece <b>kendi Anabilim Dalı</b> içerisindeki diğer danışmanlarla kıyaslanabilecek şekilde listelenmiştir. ABD Başkanlıkları iç atamalarda bu tabloları doğrudan referans alabilir.</i></p>
            {html_abd_puanlari}

            <div style="margin-top:30px; text-align:center; color:#7f8c8d; font-size:10px;">
                <i>Bu resmi rapor Bütünleşik Performans Takip Sistemi tarafından otomatik olarak oluşturulmuştur.</i>
            </div>
        </body>
        </html>
        """

        if st.button("📄 Detaylı Raporları Hazırla (PDF ve HTML)", use_container_width=True):
            result_pdf = BytesIO()
            pdf_durum = pisa.CreatePDF(src=rapor_html, dest=result_pdf, encoding='UTF-8')

            col_pdf1, col_pdf2 = st.columns(2)
            with col_pdf1:
                if not pdf_durum.err:
                    pdf_data = result_pdf.getvalue()
                    st.download_button(
                        label="📥 DİREKT İNDİR: PDF Olarak İndir",
                        data=pdf_data,
                        file_name="Enstitu_Detayli_Rapor.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
                else:
                    st.error("PDF oluşturulurken sistemsel bir hata meydana geldi.")
            
            with col_pdf2:
                st.download_button(
                    label="🌐 ALTERNATİF İNDİR: HTML Olarak İndir",
                    data=rapor_html,
                    file_name="Enstitu_Detayli_Rapor.html",
                    mime="text/html",
                    use_container_width=True
                )

except FileNotFoundError:
    st.error("⚠️ HATA: 'veriler.xlsx' dosyası bulunamadı!")
except Exception as e:
    st.error(f"⚠️ Sistemsel bir hata oluştu: {e}")