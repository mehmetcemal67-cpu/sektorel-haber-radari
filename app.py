import streamlit as st
import feedparser
import urllib.parse
import pandas as pd
import requests
import datetime
from datetime import date, datetime, timedelta
from email.utils import parsedate_to_datetime
from io import BytesIO
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

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
            # Olumlu / Stratejik Kazanımlar
            "yerlileşme": 0.8, "milli": 0.6, "rekor ihracat": 0.8, "seri üretim": 0.8,
            "patent": 0.7, "tse onaylı": 0.6, "model fabrika": 0.6, "yeşil dönüşüm": 0.6,
            "teslimat": 0.7, "başarılı entegrasyon": 0.8, "yatırım teşviki": 0.6,
            
            # Sanayi ve Teknolojiye Özel Ağır Risk / Manipülasyon
            "fiyasko": -0.9, "skandal": -0.9, "fason": -0.8, "montaj": -0.7, "illüzyon": -0.8,
            "israf": -0.8, "üretim durdu": -0.9, "şalter indirildi": -0.9, "ambargo": -0.8,
            "testi geçemedi": -0.8, "gizli ambargo": -0.8, "batık proje": -0.9, "atıl": -0.7,
            
            # Makro / İstatistiksel Riskler
            "daralma": -0.6, "sert düşüş": -0.7, "kapasite kaybı": -0.6, "gecikme": -0.5,
            "iptal": -0.8, "askıya alındı": -0.8, "karbon engeli": -0.6, "çip krizi": -0.7
        }

lexicon = load_lexicon()

# --- MANİPÜLASYON VE SÖYLEM KELİME LİSTELERİ ---
MANIPULATION_KEYWORDS = [
    "fiyasko", "skandal", "gizlenen", "gerçekler", "fason", "montaj", "yerli değil",
    "illüzyon", "şişirme", "kandırıldık", "üretim durdu", "sümen altı", "israf",
    "bağımlı", "teşvik vurgunu", "hayal kırıklığı", "yılan hikayesi", "rafa kaldırıldı", "kriz"
]

STRATEGIC_CATEGORIES = {
    "Savunma & Havacılık": [
        "aselsan", "baykar", "tusaş", "iha", "siha", "bayraktar", "kaan", 
        "çelikkubbe", "hisar", "siper", "tübitak sage", "roketsan", "havelsan", "kamikaze"
    ],
    "Otomotiv & Mobilite": [
        "togg", "elektrikli otomobil", "byd", "odmd", "şarj istasyonu", 
        "şarj soketi", "batarya teknolojileri"
    ],
    "Stratejik Hamleler & Dönüşüm": [
        "hamle programı", "dijital dönüşüm", "yeşil dönüşüm", "sınırda karbon", 
        "milli teknoloji hamlesi", "yüksek teknoloji", "model fabrika"
    ],
    "Sanayi & Kurumsal Ekosistem": [
        "sanayi ve teknoloji bakanlığı", "mehmet fatih kacır", "tübitak", "kosgeb", 
        "tüba", "organize sanayi bölgesi", "osb", "osbük", "endüstri bölgeleri", "yatırım teşvik"
    ],
    "Ekonomik Göstergeler & İstatistikler": [
        "imalat sanayii", "pmi endeksi", "reel kesim güven endeksi", "verimlilik", 
        "kapasite kullanım", "sanayi ciro", "sanayi üretim endeksi"
    ],
    "Uzay, İleri Teknoloji & Kalite": [
        "alper gezeravcı", "tua", "türkiye uzay ajansı", "tse", "türkpatent", 
        "sınai mülkiyet", "ufuk avrupa", "kalkınma ajansları", "teknofest", "çip", "yapay zeka"
    ]
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
    risk_level = "Yüksek Risk" if len(found_manipulative) > 0 or score < -0.3 else "Normal"
    
    # 3. Kategori Tespiti
    detected_category = "Genel Sanayi/Teknoloji"
    for cat, keywords in STRATEGIC_CATEGORIES.items():
        if any(kw in full_text for kw in keywords):
            detected_category = cat
            break
            
    return round(score, 2), sentiment, risk_level, found_manipulative, detected_category

# --- CANLI CANLI GOOGLE NEWS RSS HABER ÇEKME ---
def fetch_news_rss(query, time_range="1d", max_results=50):
    # 'when:1d' ekleyerek Google'ı SADECE son 24 saatin canlı haberlerini getirmeye zorluyoruz
    search_query = f"{query} when:{time_range}"
    encoded_query = urllib.parse.quote(search_query)
    rss_url = f"https://news.google.com/rss/search?q={encoded_query}&hl=tr&gl=TR&ceid=TR:tr"
    
    feed = feedparser.parse(rss_url)
    articles = []
    
    for entry in feed.entries:
        pub_date_str = entry.get('published', '')
        try:
            pub_dt = parsedate_to_datetime(pub_date_str)
            formatted_date = pub_dt.strftime('%d %b %Y %H:%M')
        except:
            formatted_date = datetime.now().strftime('%d %b %Y %H:%M')
            
        articles.append({
            'title': entry.get('title', ''),
            'description': entry.get('summary', ''),
            'url': entry.get('link', ''),
            'publishedAt': formatted_date,
            'source': {'name': entry.get('source', {}).get('title', 'Google News')}
        })
            
        if len(articles) >= max_results:
            break
            
    return articles

# --- BİLGİ NOTU / RAPOR ÜRETİCİ (.DOCX) ---
def generate_osint_docx(query, df_all, stats):
    doc = Document()
    
    # Başlık
    h = doc.add_heading("T.C. AÇIK KAYNAK MEDYA TARAMA VE İSTİHBARAT RAPORU", level=0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    p_meta = doc.add_paragraph()
    p_meta.add_run("TARAMA ODAĞI / KAPSAM: ").bold = True
    p_meta.add_run(f"{query}\n")
    p_meta.add_run("RAPOR TARİHİ: ").bold = True
    p_meta.add_run(f"{datetime.now().strftime('%d.%m.%Y %H:%M')}\n")
    
    # 1. YÖNETİCİ ÖZETİ
    doc.add_heading("1. Yönetici Özeti ve Risk Değerlendirmesi", level=1)
    p_sum = doc.add_paragraph()
    p_sum.add_run(
        f"Son 24 saatlik açık kaynak taramasında toplanan toplam {stats['total']} adet haber ve içerik "
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
st.caption("Dezenformasyon, Manipülatif Söylem ve Anlık Negatif Haber Tespiti Platformu")

with st.sidebar:
    st.header("⚙️ Tarama Parametreleri")
    
    # Anlık Takip İçin Esnetilmiş Arama Sorgusu
    default_query = (
        'sanayi OR teknoloji OR TOGG OR KAAN OR ASELSAN OR BAYKAR OR TUSAŞ OR "Çelik Kubbe" OR '
        'SİHA OR İHA OR TÜBİTAK OR KOSGEB OR OSB OR TUA OR "Milli Teknoloji" OR çip OR "Yapay Zeka"'
    )

    query = st.text_area(
        "Arama Sorgusu (Ana Terimler):",
        value=default_query,
        height=120,
        help="Geniş arama sayesinde bugün yayınlanan tüm haberler çekilir ve içindeki riskliler sistemce süzülür."
    )
    
    time_filter = st.selectbox(
        "Zaman Dilimi (Canlı Akış):",
        options=["1d", "7d", "1m"],
        format_func=lambda x: {"1d": "Son 24 Saat (Canlı/Bugün)", "7d": "Son 1 Hafta", "1m": "Son 1 Ay"}[x]
    )

    max_news = st.slider("Maksimum Haber Sayısı:", 10, 100, 40)
    only_negative = st.checkbox("Sadece Negatif/Riskli Haberleri Ekrana Getir", value=False)
    
    btn_run = st.button("🔍 Anlık Taramayı Başlat", type="primary", use_container_width=True)

if btn_run:
    with st.spinner("Anlık haber kaynakları taranıyor, canlı veriler çekiliyor..."):
        articles = fetch_news_rss(query, time_range=time_filter, max_results=max_news)
        
        if articles:
            parsed_data = []
            for a in articles:
                title = a.get('title', '') or ''
                desc = a.get('description', '') or ''
                score, sentiment, risk, manip_words, category = analyze_article(title, desc)
                
                parsed_data.append({
                    "Tarih": a.get('publishedAt', ''),
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
            st.subheader("🚨 Kritik / Riskli Söylem Barındıran Canlı Haberler")
            
            # Riskli Haberleri Öne Çıkarma
            risk_df_display = df[df['Risk_Durumu'] == 'Yüksek Risk']
            if not risk_df_display.empty:
                st.warning(f"Toplam {len(risk_df_display)} haberde manipülatif dil/yüksek negatiflik tespit edildi!")
                st.dataframe(
                    risk_df_display[['Tarih', 'Kaynak', 'Kategori', 'Başlık', 'Risk_Durumu', 'URL']],
                    use_container_width=True
                )
            else:
                st.success("Son 24 saatte kritik düzeyde manipülatif dil barındıran haber bulunamadı.")
            
            st.subheader("📋 Canlı Haber Akışı ve Analiz Tablosu")
            st.dataframe(
                df[['Tarih', 'Kaynak', 'Kategori', 'Başlık', 'Duygu', 'Skor', 'Risk_Durumu', 'URL']],
                use_container_width=True
            )
            
            # DOCX İNDİRME
            stats_dict = {'total': tot, 'neg': neg_cnt, 'risk_count': risk_cnt, 'neg_ratio': neg_ratio}
            docx_b = generate_osint_docx(query, df, stats_dict)
            
            st.download_button(
                label="📄 ANLIK AÇIK KAYNAK TARAMA RAPORUNU İNDİR (.DOCX)",
                data=docx_b,
                file_name=f"Anlik_Tarama_Raporu_{date.today()}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary"
            )
        else:
            st.info("Son 24 saat içinde kriterlere uygun canlı haber düşmedi.")