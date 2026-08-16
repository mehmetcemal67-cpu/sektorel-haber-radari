import streamlit as st
import pandas as pd
import requests
import datetime
from datetime import date, datetime
from io import BytesIO
import re
import html
import feedparser
import urllib.parse
from email.utils import parsedate_to_datetime
from duckduckgo_search import DDGS

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml

# --- ARAYÜZ AYARLARI ---
st.set_page_config(
    page_title="Sanayi & Teknoloji Açık Kaynak Radarı",
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
        return {
            "yerlileşme": 0.8, "milli": 0.6, "rekor ihracat": 0.8, "seri üretim": 0.8,
            "patent": 0.7, "tse onaylı": 0.6, "model fabrika": 0.6, "yeşil dönüşüm": 0.6,
            "fiyasko": -0.9, "skandal": -0.9, "fason": -0.8, "montaj": -0.7, "illüzyon": -0.8,
            "israf": -0.8, "üretim durdu": -0.9, "şalter indirildi": -0.9, "ambargo": -0.8,
            "daralma": -0.6, "sert düşüş": -0.7, "kapasite kaybı": -0.6, "kriz": -0.8
        }

lexicon = load_lexicon()

MANIPULATION_KEYWORDS = [
    "fiyasko", "skandal", "gizlenen", "gerçekler", "fason", "montaj", "yerli değil",
    "illüzyon", "şişirme", "kandırıldık", "sümen altı", "israf", "teşvik vurgunu",
    "hayal kırıklığı", "yılan hikayesi", "rafa kaldırıldı", "yalan", "sansür", "şüphe",
    "algı operasyonu", "makyajlı", "hayali", "balon", "vurgun", "pes dedirtti",
    "üretim durdu", "şalter indirildi", "batık proje", "atıl", "gecikme", "iptal",
    "askıya alındı", "testi geçemedi", "arıza", "çöküş", "teslim edilemedi", "patladı",
    "kapasite kaybı", "sözleşme feshi", "hazır alım", "dışa bağımlı",
    "kriz", "zarar", "iflas", "konkordato", "borç batağı", "maliyet artışı",
    "bütçe açığı", "daralma", "sert düşüş", "kaynak tükendi", "pazar kaybı",
    "ambargo", "gizli ambargo", "yaptırım", "çip krizi", "tedarik engeli",
    "karbon engeli", "lisans reddi", "kırmızı çizgi", "blokaj", "provokasyon"
]

STRATEGIC_CATEGORIES = {
    "Savunma & Havacılık": [
        "aselsan", "baykar", "tusaş", "iha", "siha", "bayraktar", "kaan", 
        "çelikkubbe", "hisar", "siper", "tübitak sage", "roketsan", "havelsan", "kamikaze"
    ],
    "Otomotiv & Mobilite": [
        "togg", "elektrikli otomobil", "byd", "odmd", "şarj istasyonu", "batarya"
    ],
    "Bölgesel Güvenlik & Jeopolitik": [
        "yunanistan", "yunan basını", "atina", "kathimerini", "ta nea", "protothema",
        "ege", "doğu akdeniz", "rafale", "f-16", "fir hattı"
    ],
    "Sanayi & Kurumsal Ekosistem": [
        "sanayi ve teknoloji bakanlığı", "mehmet fatih kacır", "tübitak", "kosgeb", 
        "organize sanayi bölgesi", "osb", "yatırım teşvik"
    ],
    "Uzay, İleri Teknoloji & Kalite": [
        "alper gezeravcı", "tua", "türkiye uzay ajansı", "tse", "türkpatent", 
        "teknofest", "çip", "yapay zeka"
    ]
}

def analyze_article(title, description):
    full_text = f"{title} {description}".lower()
    words = re.sub(r'[^\w\s]', ' ', full_text).split()
    
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
        
    found_manipulative = [kw for kw in MANIPULATION_KEYWORDS if kw in full_text]
    risk_level = "Yüksek Risk" if len(found_manipulative) > 0 or score < -0.3 else "Normal"
    
    detected_category = "Genel Sanayi/Teknoloji"
    for cat, keywords in STRATEGIC_CATEGORIES.items():
        if any(kw in full_text for kw in keywords):
            detected_category = cat
            break
            
    return round(score, 2), sentiment, risk_level, found_manipulative, detected_category

# --- KORUMALI & HİBRİT HABER ÇEKME MOTORU ---
def fetch_robust_news(query_text, time_range="1d", max_results=50):
    articles = []
    seen_urls = set()
    today = date.today()
    days_limit = {"1d": 1, "7d": 7, "14d": 14}.get(time_range, 1)

    # 1. YÖNTEM: DuckDuckGo (Temizlenmiş Alt Arama Grupları ile Ratelimit Engelini Aşma)
    raw_keywords = [k.strip().replace('"', '') for k in query_text.split('OR')]
    # En önemli ilk 5 temel arama kümesine bölme
    sub_queries = [
        " ".join(raw_keywords[:4]),
        " ".join(raw_keywords[4:8]),
        " ".join(raw_keywords[8:12])
    ]
    
    time_ddg = {"1d": "d", "7d": "w", "14d": "w"}.get(time_range, "d")

    try:
        with DDGS() as ddgs:
            for sq in sub_queries:
                if len(articles) >= max_results:
                    break
                if not sq.strip():
                    continue
                results = ddgs.news(keywords=sq, region="tr-tr", timelimit=time_ddg, max_results=20)
                for r in results:
                    url = r.get('url', '')
                    if url and url not in seen_urls:
                        seen_urls.add(url)
                        raw_date = r.get('date', '')
                        try:
                            dt = datetime.fromisoformat(raw_date.replace('Z', '+00:00'))
                            if (today - dt.date()).days > days_limit:
                                continue
                            formatted_date = dt.strftime('%d %b %Y %H:%M')
                        except:
                            formatted_date = datetime.now().strftime('%d %b %Y %H:%M')

                        articles.append({
                            'title': r.get('title', ''),
                            'description': r.get('body', ''),
                            'url': url,
                            'publishedAt': formatted_date,
                            'source': {'name': r.get('source', 'Web')}
                        })
    except Exception:
        pass  # DuckDuckGo engel yerse sessizce Google RSS yedeğine geç

    # 2. YÖNTEM (YEDEK): Google RSS (Katı Tarih Filtreli Fallback)
    if len(articles) < 5:
        clean_q = " OR ".join(raw_keywords[:6])
        search_query = f"{clean_q} when:{time_range}"
        encoded_query = urllib.parse.quote(search_query)
        rss_url = f"https://news.google.com/rss/search?q={encoded_query}&hl=tr&gl=TR&ceid=TR:tr"
        
        try:
            feed = feedparser.parse(rss_url)
            for entry in feed.entries:
                url = entry.get('link', '')
                if url and url not in seen_urls:
                    pub_date_str = entry.get('published', '')
                    try:
                        pub_dt = parsedate_to_datetime(pub_date_str)
                        if (today - pub_dt.date()).days > days_limit:
                            continue
                        formatted_date = pub_dt.strftime('%d %b %Y %H:%M')
                    except:
                        formatted_date = datetime.now().strftime('%d %b %Y %H:%M')

                    seen_urls.add(url)
                    articles.append({
                        'title': entry.get('title', ''),
                        'description': entry.get('summary', ''),
                        'url': url,
                        'publishedAt': formatted_date,
                        'source': {'name': entry.get('source', {}).get('title', 'Google News')}
                    })
                    if len(articles) >= max_results:
                        break
        except Exception:
            pass

    return articles[:max_results]

# --- GÜVENLİ WORD KÖPRÜ LİNK FONKSİYONU ---
def add_hyperlink(paragraph, url, text):
    try:
        part = paragraph.part
        safe_url = html.escape(url)
        safe_text = html.escape(text)
        
        r_id = part.relate_to(safe_url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
        hyperlink = parse_xml(f'<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" r:id="{r_id}"/>')
        new_run = parse_xml(f'<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:rStyle w:val="Hyperlink"/><w:color w:val="0000FF"/><w:u w:val="single"/></w:rPr><w:t>{safe_text}</w:t></w:r>')
        hyperlink.append(new_run)
        paragraph._p.append(hyperlink)
    except Exception:
        paragraph.add_run(f"{text} ({url})")

def generate_osint_docx(query, df_all, stats):
    doc = Document()
    
    h = doc.add_heading("T.C. AÇIK KAYNAK MEDYA TARAMA VE İSTİHBARAT RAPORU", level=0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    p_meta = doc.add_paragraph()
    p_meta.add_run("TARAMA ODAĞI / KAPSAM: ").bold = True
    p_meta.add_run(f"{query}\n")
    p_meta.add_run("RAPOR TARİHİ: ").bold = True
    p_meta.add_run(f"{datetime.now().strftime('%d.%m.%Y %H:%M')}\n")
    
    doc.add_heading("1. Yönetici Özeti ve Risk Değerlendirmesi", level=1)
    p_sum = doc.add_paragraph()
    p_sum.add_run(
        f"Seçilen zaman diliminde açık kaynak taramasında toplanan toplam {stats['total']} adet haber ve içerik "
        f"sistem tarafından incelenmiştir. Yapılan analiz sonucunda haberlerin "
        f"%{stats['neg_ratio']:.1f}'inin ({stats['neg']} adet) olumsuz/negatif tonda olduğu, "
        f"{stats['risk_count']} adet haberde ise kamuoyunu yönlendirmeye veya algı oluşturmaya dönük "
        f"manipülatif/sansasyonel söylem kalıplarının kullanıldığı tespit edilmiştir."
    )
    
    doc.add_heading("2. Kritik / Manipülatif Söylem Barındıran Haberler", level=1)
    risk_df = df_all[df_all['Risk_Durumu'] == 'Yüksek Risk']
    
    if not risk_df.empty:
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        hdr = table.rows[0].cells
        hdr[0].text = 'Tarih / Kaynak'
        hdr[1].text = 'Kategori'
        hdr[2].text = 'Haber Başlığı (Bağlantılı)'
        hdr[3].text = 'Tespit Edilen Söylem'
        
        for _, r in risk_df.iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = f"{r['Tarih']}\n{r['Kaynak']}"
            row_cells[1].text = r['Kategori']
            
            p_link = row_cells[2].paragraphs[0]
            add_hyperlink(p_link, r['URL'], r['Başlık'])
            
            row_cells[3].text = ", ".join(r['Manipülasyon_Kelimeleri']) if r['Manipülasyon_Kelimeleri'] else "Yüksek Negatif Ton"
    else:
        doc.add_paragraph("Kritik düzeyde manipülatif söylem barındıran haber tespit edilmemiştir.")
        
    doc.add_heading("3. Genel Haber Akışı ve Duygu Dağılımı", level=1)
    table2 = doc.add_table(rows=1, cols=5)
    table2.style = 'Table Grid'
    hdr2 = table2.rows[0].cells
    hdr2[0].text = 'Tarih'
    hdr2[1].text = 'Kaynak'
    hdr2[2].text = 'Kategori'
    hdr2[3].text = 'Başlık (Bağlantılı)'
    hdr2[4].text = 'Duygu / Skor'
    
    for _, r in df_all.iterrows():
        rc = table2.add_row().cells
        rc[0].text = r['Tarih']
        rc[1].text = r['Kaynak']
        rc[2].text = r['Kategori']
        
        p_link2 = rc[3].paragraphs[0]
        add_hyperlink(p_link2, r['URL'], r['Başlık'])
        
        rc[4].text = f"{r['Duygu']} ({r['Skor']})"
        
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# --- ARAYÜZ (STREAMLIT) ---
st.title("🛡️ Sanayi, Teknoloji & Güvenlik Açık Kaynak Tarama Radarı")
st.caption("Dezenformasyon, Manipülatif Söylem, Yunanistan Basını ve Anlık Negatif Haber Tespiti Platformu")

with st.sidebar:
    st.header("⚙️ Tarama Parametreleri")
    
    default_query = (
        'sanayi OR teknoloji OR TOGG OR KAAN OR ASELSAN OR BAYKAR OR TUSAŞ OR ROKETSAN OR HAVELSAN OR '
        'TÜBİTAK OR KOSGEB OR Çelik Kubbe OR SİHA OR İHA OR Milli Teknoloji OR çip OR Yapay Zeka OR '
        'Yunanistan OR Yunan basını OR Atina OR Kathimerini OR Doğu Akdeniz OR F-16 OR Rafale'
    )

    query = st.text_area(
        "Arama Sorgusu (Ana Terimler):",
        value=default_query,
        height=150
    )
    
    time_filter = st.selectbox(
        "Zaman Dilimi (Canlı Akış):",
        options=["1d", "7d", "14d"],
        format_func=lambda x: {
            "1d": "Son 24 Saat (Canlı / Bugün)", 
            "7d": "Son 1 Hafta", 
            "14d": "Son 2 Hafta"
        }[x]
    )

    max_news = st.slider("Maksimum Haber Sayısı:", 10, 100, 40)
    only_negative = st.checkbox("Sadece Negatif/Riskli Haberleri Ekrana Getir", value=False)
    
    btn_run = st.button("🔍 Açık Kaynak Taramasını Başlat", type="primary", use_container_width=True)

if btn_run:
    with st.spinner("Anlık canlı haber kaynakları taranıyor..."):
        articles = fetch_robust_news(query, time_range=time_filter, max_results=max_news)
        
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
            
            tot = len(df)
            neg_cnt = sum(df['Duygu'] == 'Negatif')
            risk_cnt = sum(df['Risk_Durumu'] == 'Yüksek Risk')
            neg_ratio = (neg_cnt / tot * 100) if tot > 0 else 0
            
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("İncelenen Haber", tot)
            k2.metric("Negatif Haberler", neg_cnt)
            k3.metric("Manipülasyon Riskli", risk_cnt, delta="Kritik Dil" if risk_cnt > 0 else "Normal", delta_color="inverse")
            k4.metric("Negatif Haber Oranı", f"%{neg_ratio:.1f}")
            
            st.markdown("---")
            st.subheader("🚨 Kritik / Riskli Söylem Barındıran Canlı Haberler")
            
            risk_df_display = df[df['Risk_Durumu'] == 'Yüksek Risk']
            if not risk_df_display.empty:
                st.warning(f"Toplam {len(risk_df_display)} haberde manipülatif dil/yüksek negatiflik tespit edildi!")
                st.dataframe(
                    risk_df_display[['Tarih', 'Kaynak', 'Kategori', 'Başlık', 'Risk_Durumu', 'URL']],
                    use_container_width=True
                )
            else:
                st.success("Seçilen zaman aralığında kritik düzeyde manipülatif dil barındıran haber bulunamadı.")
            
            st.subheader("📋 Canlı Haber Akışı ve Analiz Tablosu")
            st.dataframe(
                df[['Tarih', 'Kaynak', 'Kategori', 'Başlık', 'Duygu', 'Skor', 'Risk_Durumu', 'URL']],
                use_container_width=True
            )
            
            stats_dict = {'total': tot, 'neg': neg_cnt, 'risk_count': risk_cnt, 'neg_ratio': neg_ratio}
            docx_b = generate_osint_docx(query, df, stats_dict)
            
            st.download_button(
                label="📄 ANLIK AÇIK KAYNAK TARAMA RAPORUNU İNDİR (.DOCX)",
                data=docx_b,
                file_name=f"Acik_Kaynak_Raporu_{date.today()}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary"
            )
        else:
            st.info("Seçilen zaman diliminde kriterlere uygun haber bulunamadı.")