import streamlit as st
import pandas as pd
import requests
from io import BytesIO
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import datetime

# --- ARAYÜZ AYARLARI ---
st.set_page_config(
    page_title="Açık Kaynak Tarama & Manipülasyon Radarı",
    page_icon="🛡️",
    layout="wide"
)

# --- SÖZLÜK YÜKLEME ---
@st.cache_data
def load_lexicon():
    try:
        df = pd.read_csv("tr_sentiment_lexicon.csv")
        return dict(zip(df['kelime'].str.lower(), df['skor']))
    except Exception:
        # Yedek temel sözlük
        return {
            "başarı": 0.8, "büyüme": 0.7, "rekor": 0.8, "yerli": 0.5, "milli": 0.5,
            "kriz": -0.8, "çöküş": -0.9, "iflas": -1.0, "durdu": -0.7, "fiyasko": -0.9,
            "skandal": -0.9, "gizlenen": -0.7, "iddia": -0.4, "şüphe": -0.5, "gecikti": -0.6,
            "iptal": -0.8, "ambargo": -0.7, "yetersiz": -0.6, "zarar": -0.7, "facia": -0.9
        }

lexicon = load_lexicon()

# --- MANİPÜLASYON VE SÖYLEM KELİME LİSTELERİ ---
MANIPULATION_KEYWORDS = [
    "fiyasko", "skandal", "gizlenen", "gerçekler", "şok", "iddia edildi", 
    "facia", "çöktü", "hayal", "durdu", "balon", "vurgun", "sir"
]

STRATEGIC_CATEGORIES = {
    "Savunma Sanayii": ["iha", "siha", "kaan", "aselsan", "tusaş", "baykar", "mühimmat", "savunma", "roket", "radar"],
    "Otomotiv & Mobilite": ["togg", "elektrikli araç", "batarya", "otomobil", "üretim hattı"],
    "Teknoloji & Ar-Ge": ["çip", "yazılım", "yapay zeka", "tübitak", "uzay", "uydu", "tekno"],
    "Ekonomi & Yatırım": ["yatırım", "teşvik", "sanayi üretimi", "ihracat", "fabrika", "istihdam"]
}

# --- METİN ANALİZ MOTORU ---
def analyze_article(title, description):
    full_text = f"{title} {description}".lower()
    words = full_text.replace(".", " ").replace(",", " ").replace("?", " ").replace("!", " ").split()
    
    # 1. Duygu Skoru Hesaplama
    total_score = 0.0
    matched_words = []
    for w in words:
        if w in lexicon:
            total_score += lexicon[w]
            matched_words.append(w)
            
    score = total_score / len(matched_words) if matched_words else 0.0
    
    if score > 0.15:
        sentiment = "Pozitif"
    elif score < -0.15:
        sentiment = "Negatif"
    else:
        sentiment = "Nötr"
        
    # 2. Manipülatif Dil / Risk Tespiti
    found_manipulative = [kw for kw in MANIPULATION_KEYWORDS if kw in full_text]
    risk_level = "Yüksek Risk" if len(found_manipulative) > 0 or score < -0.4 else "Normal"
    
    # 3. Kategori Tespiti
    detected_category = "Genel Sanayi/Teknoloji"
    for cat, keywords in STRATEGIC_CATEGORIES.items():
        if any(kw in full_text for kw in keywords):
            detected_category = cat
            break
            
    return round(score, 2), sentiment, risk_level, found_manipulative, detected_category

# --- NEWSAPI HABER ÇEKME ---
def fetch_news(api_key, query, from_date, to_date, max_results):
    url = "https://newsapi.org/v2/everything"
    params = {
        'q': query,
        'from': from_date,
        'to': to_date,
        'pageSize': min(max_results, 100),
        'sortBy': 'publishedAt',
        'apiKey': api_key,
        'language': 'tr'
    }
    res = requests.get(url, params=params)
    if res.status_code == 200:
        return res.json().get('articles', [])
    return []

# --- BİLGİ NOTU / RAPOR ÜRETİCİ (.DOCX) ---
def generate_osint_docx(query, df_all, stats):
    doc = Document()
    
    # Başlık
    h = doc.add_heading("T.C. AÇIK KAYNAK MEYDA TARAMA VE İSTİHBARAT RAPORU", level=0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    p_meta = doc.add_paragraph()
    p_meta.add_run("TARAMA ODAĞI / KAPSAM: ").bold = True
    p_meta.add_run(f"{query}\n")
    p_meta.add_run("RAPOR TARİHİ: ").bold = True
    p_meta.add_run(f"{datetime.datetime.now().strftime('%d.%m.%Y %H:%M')}\n")
    
    # 1. YÖNETİCİ ÖZETİ
    doc.add_heading("1. Yönetici Özeti ve Risk Değerlendirmesi", level=1)
    p_sum = doc.add_paragraph()
    p_sum.add_run(
        f"Belirtilen tarih aralığında açık kaynaklardan toplanan toplam {stats['total']} adet haber ve içerik "
        f"sistem tarafından incelenmiştir. Yapılan analiz sonucunda haberlerin "
        f"%{stats['neg_ratio']:.1f}'inin ({stats['neg']} adet) olumsuz/negatif tonda olduğu, "
        f"{stats['risk_count']} adet haberde ise kamuoyunu yönlendirmeye veya algı oluşturmaya dönük "
        f"manipülatif/sansasyonel söylem kalıplarının kullanıldığı tespit edilmiştir."
    )
    
    # 2. ÖNE ÇIKAN RİSKLİ VE NEGATİF HABERLER
    doc.add_heading("2. Kritik / Manipülatif Söylem Barındıran Haberler", level=1)
    risk_df = df_all[df_all['Risk_Durumu'] == 'Yüksek Risk']
    
    if not risk_df.empty:
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        hdr = table.rows[0].cells
        hdr[0].text = 'Tarih / Kaynak'
        hdr[1].text = 'Kategori'
        hdr[2].text = 'Haber Başlığı'
        hdr[3].text = 'Tespit Edilen Söylem/Kelimeler'
        
        for _, r in risk_df.iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = f"{r['Tarih']}\n{r['Kaynak']}"
            row_cells[1].text = r['Kategori']
            row_cells[2].text = r['Başlık']
            row_cells[3].text = ", ".join(r['Manipülasyon_Kelimeleri']) if r['Manipülasyon_Kelimeleri'] else "Yüksek Negatif Ton"
    else:
        doc.add_paragraph("Kritik düzeyde manipülatif söylem barındıran haber tespit edilmemiştir.")
        
    # 3. TÜM HABER AKIŞI TABLOSU
    doc.add_heading("3. Genel Haber Akışı ve Duygu Dağılımı", level=1)
    table2 = doc.add_table(rows=1, cols=5)
    table2.style = 'Table Grid'
    hdr2 = table2.rows[0].cells
    hdr2[0].text = 'Tarih'
    hdr2[1].text = 'Kaynak'
    hdr2[2].text = 'Kategori'
    hdr2[3].text = 'Başlık'
    hdr2[4].text = 'Duygu / Skor'
    
    for _, r in df_all.iterrows():
        rc = table2.add_row().cells
        rc[0].text = r['Tarih']
        rc[1].text = r['Kaynak']
        rc[2].text = r['Kategori']
        rc[3].text = r['Başlık']
        rc[4].text = f"{r['Duygu']} ({r['Skor']})"
        
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# --- ARAYÜZ (STREAMLIT) ---
st.title("🛡️ Sanayi & Teknoloji Açık Kaynak Tarama Radarı")
st.caption("Dezenformasyon, Manipülatif Söylem ve Negatif Haber Tespiti Platformu")

with st.sidebar:
    st.header("⚙️ Tarama Parametreleri")
    api_key = st.text_input("NewsAPI Key:", value="db32a44046bb4e6ab4c629b6269d2336" type="password")
    
    # Varsayılan Geniş Sanayi & Teknoloji Sorgusu
    default_query = "(sanayi OR teknoloji OR togg OR iha OR siha OR kaan OR aselsan OR tübitak OR " \
                    "yatırım OR ihracat OR çip) AND (kriz OR durdu OR iflas OR skandal OR iptal OR " \
                    "ambargo OR maliyet OR zarar OR iddia OR fiyasko OR tehlike OR fuj)"
                    
    query = st.text_area("Arama Sorgusu (Boolean):", value=default_query, height=120)
    
    c1, c2 = st.columns(2)
    with c1:
        s_date = st.date_input("Başlangıç", datetime.date.today() - datetime.timedelta(days=7))
    with c2:
        e_date = st.date_input("Bitiş", datetime.date.today())
        
    max_news = st.slider("Maksimum Haber Sayısı:", 10, 100, 30)
    only_negative = st.checkbox("Sadece Negatif/Riskli Haberleri Süz", value=False)
    
    btn_run = st.button("🔍 Açık Kaynak Taramasını Başlat", type="primary", use_container_width=True)

if btn_run:
    if not api_key:
        st.error("Lütfen sol panelden NewsAPI Key bilginizi giriniz.")
    else:
        with st.spinner("Açık kaynaklar taranıyor, dezenformasyon ve negatif söylemler analiz ediliyor..."):
            articles = fetch_news(api_key, query, s_date.strftime('%Y-%m-%d'), e_date.strftime('%Y-%m-%d'), max_news)
            
            if articles:
                parsed_data = []
                for a in articles:
                    title = a.get('title', '') or ''
                    desc = a.get('description', '') or ''
                    score, sentiment, risk, manip_words, category = analyze_article(title, desc)
                    
                    parsed_data.append({
                        "Tarih": a.get('publishedAt', '')[:10],
                        "Kaynak": a.get('source', {}).get('name', 'Bilinmiyor'),
                        "Kategori": category,
                        "Başlık": title,
                        "Özet": desc,
                        "Duygu": sentiment,
                        "Skor": score,
                        "Risk_Durumu": risk,
                        "Manipülasyon_Kelimeleri": manip_words,
                        "URL": a.get('url', '')
                    })
                
                df = pd.DataFrame(parsed_data)
                
                if only_negative:
                    df = df[(df['Duygu'] == 'Negatif') | (df['Risk_Durumu'] == 'Yüksek Risk')]
                
                # İstatistikler
                tot = len(df)
                neg_cnt = sum(df['Duygu'] == 'Negatif')
                risk_cnt = sum(df['Risk_Durumu'] == 'Yüksek Risk')
                neg_ratio = (neg_cnt / tot * 100) if tot > 0 else 0
                
                # METRİK KARTLARI
                k1, k2, k3, k4 = st.columns(4)
                k1.metric("İncelenen Haber", tot)
                k2.metric("Negatif Haberler", neg_cnt)
                k3.metric("Manipülasyon Riskli", risk_cnt, delta="Kritik Dil" if risk_cnt > 0 else "Normal", delta_color="inverse")
                k4.metric("Negatif Haber Oranı", f"%{neg_ratio:.1f}")
                
                st.markdown("---")
                st.subheader("🚨 Kritik / Riskli Söylem Barındıran Haberler")
                
                # Riskli Haberleri Öne Çıkarma
                risk_df_display = df[df['Risk_Durumu'] == 'Yüksek Risk']
                if not risk_df_display.empty:
                    st.warning(f"Toplam {len(risk_df_display)} haberde manipülatif dil/yüksek negatiflik tespit edildi!")
                    st.dataframe(
                        risk_df_display[['Tarih', 'Kaynak', 'Kategori', 'Başlık', 'Risk_Durumu', 'URL']],
                        use_container_width=True
                    )
                else:
                    st.success("Taramada kritik düzeyde manipülatif dil barındıran haber bulunamadı.")
                
                st.subheader("📋 Tüm Haber Akışı ve Analiz Tablosu")
                st.dataframe(
                    df[['Tarih', 'Kaynak', 'Kategori', 'Başlık', 'Duygu', 'Skor', 'Risk_Durumu', 'URL']],
                    use_container_width=True
                )
                
                # DOCX İNDİRME
                stats_dict = {'total': tot, 'neg': neg_cnt, 'risk_count': risk_cnt, 'neg_ratio': neg_ratio}
                docx_b = generate_osint_docx(query, df, stats_dict)
                
                st.download_button(
                    label="📄 AÇIK KAYNAK TARAMA RAPORUNU İNDİR (.DOCX)",
                    data=docx_b,
                    file_name=f"Acik_Kaynak_Tarama_Raporu_{datetime.date.today()}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary"
                )
            else:
                st.info("Kriterlere uygun haber verisi bulunamadı.")
