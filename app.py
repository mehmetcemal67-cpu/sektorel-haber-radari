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
from bs4 import BeautifulSoup
from googlenewsdecoder import gnd

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

# --- ARAYÜZ AYARLARI ---
st.set_page_config(
    page_title="Sanayi, Teknoloji & Güvenlik Açık Kaynak Radarı",
    page_icon="🛡️",
    layout="wide"
)

# --- SÖZLÜK VE KATEGORİ TANIMLARI ---
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

# --- GOOGLE LINK ÇÖZÜCÜ VE SAYFA KAZIYICI ---
def unwrap_and_scrape(url):
    """Google News linkini gerçek siteye çözümler, görsel ve tam paragrafları kazır."""
    final_url = url
    if "news.google.com" in url:
        try:
            decoded = gnd(url)
            if decoded and decoded.get("status") and decoded.get("decoded_url"):
                final_url = decoded["decoded_url"]
        except Exception:
            final_url = url

    img_url = ""
    paragraphs = []
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36'
    }
    
    try:
        resp = requests.get(final_url, headers=headers, timeout=4, allow_redirects=True)
        if resp.status_code == 200:
            soup = BeautifulSoup(resp.text, 'html.parser')
            
            # Görsel tespiti (og:image, twitter:image)
            for meta_prop in ['og:image', 'twitter:image', 'og:image:secure_url']:
                tag = soup.find('meta', property=meta_prop) or soup.find('meta', attrs={'name': meta_prop})
                if tag and tag.get('content'):
                    candidate = tag['content'].strip()
                    if candidate.startswith('http') and "gstatic.com" not in candidate and "google" not in candidate.lower():
                        img_url = candidate
                        break
                        
            # Paragraf tespiti
            for p in soup.find_all('p'):
                txt = p.get_text().strip()
                if len(txt) > 35 and not any(w in txt.lower() for w in ['çerez', 'cookie', 'abone', 'tıklayın', 'copyright', 'gizlilik']):
                    paragraphs.append(txt)
    except Exception:
        pass
        
    full_text = " ".join(paragraphs)
    return final_url, img_url, full_text

# --- DUYGU VE RİSK ANALİZİ ---
def analyze_article(title, text):
    full_text = f"{title} {text}".lower()
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

# --- AKICI DÜZYAZI VE DERİN ÖZET ÜRETİCİ ---
def build_prose_analysis(title, scraped_text, category, sentiment, risk_level, manip_words):
    """
    Maddeli/şablon yapıları tamamen kaldırır. 
    Haberin akışını ve stratejik önemini anlatan akıcı bir DÜZYAZI PARAGRAFI oluşturur.
    """
    clean_text = scraped_text.strip() if scraped_text else ""
    sentences = [s.strip() for s in re.split(r'(?<=[.!?])\s+', clean_text) if len(s.strip()) > 25]
    
    if len(sentences) >= 2:
        story = " ".join(sentences[:5])
    elif sentences:
        story = sentences[0]
    else:
        story = f"{title} konusuyla ilgili gelişmeler açık kaynaklar üzerinden takip edilmektedir."

    # Akıcı düzyazı şeklinde stratejik değerlendirme eklemesi
    prose_eval = f" İlgili gelişme {category.lower()} alanı açısından kritik önem taşımaktadır."
    
    if sentiment == "Pozitif":
        prose_eval += " Haber içeriğinde öne çıkan detaylar, yerli üretim kapasitesi, sektörel büyüme ve teknolojik yetkinliklerin güçlenmesi yönünde olumlu bir tablo çizmektedir."
    elif sentiment == "Negatif":
        prose_eval += " Haber içeriği incelendiğinde; sektörel daralma, maliyet yükü veya operasyonel risklerin öne çıktığı görülmektedir."
    else:
        prose_eval += " Makale genel itibarıyla teknik bilgilendirme ve kurumsal süreç aktarımı niteliğindedir."

    if risk_level == "Yüksek Risk":
        prose_eval += f" Ayrıca metin içerisinde kamuoyunu yönlendirme veya algı oluşturma riski barındıran söylemler ({', '.join(manip_words)}) tespit edilmiş olup, konunun kurumsal takibi önerilmektedir."

    return f"{story}{prose_eval}"

# --- HABER TOPLAMA MOTORU ---
def fetch_robust_news(query_text, time_range="1d", max_results=25):
    articles = []
    seen_urls = set()
    today = date.today()

    raw_keywords = [k.strip().replace('"', '') for k in query_text.split('OR')]
    clean_q = " OR ".join(raw_keywords[:5])
    search_query = f"{clean_q} when:{time_range}"
    encoded_query = urllib.parse.quote(search_query)
    rss_url = f"https://news.google.com/rss/search?q={encoded_query}&hl=tr&gl=TR&ceid=TR:tr"
    
    try:
        feed = feedparser.parse(rss_url)
        for entry in feed.entries:
            if len(articles) >= max_results:
                break
            raw_link = entry.get('link', '')
            if not raw_link or raw_link in seen_urls:
                continue
            
            # Google linkini çöz ve siteyi kazı
            real_url, fetched_img, full_text = unwrap_and_scrape(raw_link)
            
            if real_url in seen_urls:
                continue
            seen_urls.add(real_url)
            seen_urls.add(raw_link)

            pub_date_str = entry.get('published', '')
            try:
                pub_dt = parsedate_to_datetime(pub_date_str)
                formatted_date = pub_dt.strftime('%d %b %Y %H:%M')
            except Exception:
                formatted_date = datetime.now().strftime('%d %b %Y %H:%M')

            articles.append({
                'title': entry.get('title', ''),
                'full_text': full_text,
                'url': real_url,
                'image_url': fetched_img,
                'publishedAt': formatted_date,
                'source': {'name': entry.get('source', {}).get('title', 'Açık Basın')}
            })
    except Exception:
        pass

    return articles[:max_results]

# --- WORD DOKÜMANI VE LINK YARDIMCILARI ---
def add_safe_hyperlink(paragraph, url, text):
    """Word içerisine tıklanabilir Mavi Link ekler."""
    try:
        part = paragraph.part
        r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
        safe_text = html.escape(text)
        
        xml_str = (
            f'<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" r:id="{r_id}">'
            f'<w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:color w:val="0000FF"/><w:u w:val="single"/></w:rPr>'
            f'<w:t>{safe_text}</w:t></w:r></w:hyperlink>'
        )
        hyperlink = parse_xml(xml_str)
        paragraph._p.append(hyperlink)
    except Exception:
        paragraph.add_run(f"{text} (Link: {url})")

def download_image_stream(img_url):
    """Resmi indirip Word dosyasına eklenmeye hazır hale getirir."""
    if not img_url or "gstatic.com" in img_url:
        return None
    try:
        headers = {'User-Agent': 'Mozilla/5.0'}
        resp = requests.get(img_url, headers=headers, timeout=4)
        if resp.status_code == 200 and len(resp.content) > 2000:
            return BytesIO(resp.content)
    except Exception:
        pass
    return None

def style_table_cell(cell, bg_hex=None, bold=False, font_size=8.5, color_rgb=(0,0,0)):
    if bg_hex:
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_hex}"/>')
        cell._tc.get_or_add_tcPr().append(shd)
    
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m in ['top', 'bottom', 'left', 'right']:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), '120')
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

    for p in cell.paragraphs:
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        for r in p.runs:
            r.font.name = 'Arial'
            r.font.size = Pt(font_size)
            r.font.bold = bold
            r.font.color.rgb = RGBColor(*color_rgb)

# --- WORD RAPOR OLUŞTURUCU ---
def generate_osint_docx(query, df_all, stats):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    h = doc.add_heading("T.C. AÇIK KAYNAK MEDYA TARAMA VE İSTİHBARAT RAPORU", level=0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in h.runs:
        r.font.name = 'Arial'
        r.font.size = Pt(13)
        r.font.bold = True
        r.font.color.rgb = RGBColor(27, 54, 93)

    p_meta = doc.add_paragraph()
    p_meta.add_run("TARAMA ODAĞI / KAPSAM: ").bold = True
    p_meta.add_run(f"{query}\n")
    p_meta.add_run("RAPOR TARİHİ: ").bold = True
    p_meta.add_run(f"{datetime.now().strftime('%d.%m.%Y %H:%M')}\n")
    for r in p_meta.runs:
        r.font.name = 'Arial'
        r.font.size = Pt(9)

    doc.add_heading("1. Yönetici Özeti ve Risk Değerlendirmesi", level=1)
    p_sum = doc.add_paragraph()
    p_sum.add_run(
        f"Seçilen zaman diliminde açık kaynaklardan derlenen toplam {stats['total']} adet haber ve içerik "
        f"sistem tarafından incelenmiştir. Yapılan detaylı metin analizleri sonucunda haberlerin "
        f"%{stats['neg_ratio']:.1f}'inin ({stats['neg']} adet) olumsuz/negatif ton taşıdığı tespit edilmiştir. "
        f"Ayrıca {stats['risk_count']} adet içerikte kamuoyunu yönlendirmeye dönük "
        f"manipülatif söylem kalıplarının kullanıldığı belirlenmiştir."
    )
    for r in p_sum.runs:
        r.font.name = 'Arial'
        r.font.size = Pt(9.5)

    doc.add_heading("2. Kritik / Manipülatif Söylem Barındıran Haberler", level=1)
    risk_df = df_all[df_all['Risk_Durumu'] == 'Yüksek Risk']
    
    if not risk_df.empty:
        table = doc.add_table(rows=1, cols=6)
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        headers = ['Görsel', 'Tarih / Kaynak', 'Kategori', 'Haber Başlığı (Tıklanabilir Link)', 'Düzyazı Haber Özeti ve Analizi', 'Tespit Edilen Söylem']
        for i, title in enumerate(headers):
            hdr_cells[i].text = title
            style_table_cell(hdr_cells[i], bg_hex="1B365D", bold=True, font_size=9, color_rgb=(255, 255, 255))

        for _, r in risk_df.iterrows():
            row_cells = table.add_row().cells
            
            img_stream = download_image_stream(r.get('Görsel_URL', ''))
            if img_stream:
                try:
                    p_img = row_cells[0].paragraphs[0]
                    p_img.add_run().add_picture(img_stream, width=Inches(1.2))
                except Exception:
                    row_cells[0].text = "Görsel Yüklenemedi"
            else:
                row_cells[0].text = "Görsel Yok"

            row_cells[1].text = f"{r['Tarih']}\n{r['Kaynak']}"
            row_cells[2].text = r['Kategori']
            
            p_link = row_cells[3].paragraphs[0]
            add_safe_hyperlink(p_link, r['URL'], r['Başlık'])
            
            row_cells[4].text = r['Özet']
            row_cells[5].text = ", ".join(r['Manipülasyon_Kelimeleri']) if r['Manipülasyon_Kelimeleri'] else "Yüksek Negatif Ton"
            
            for c in row_cells:
                style_table_cell(c, font_size=8)
    else:
        doc.add_paragraph("Kritik düzeyde manipülatif söylem barındıran haber tespit edilmemiştir.")
        
    doc.add_heading("3. Genel Haber Akışı ve Duygu Dağılımı", level=1)
    table2 = doc.add_table(rows=1, cols=6)
    table2.style = 'Table Grid'
    hdr_cells2 = table2.rows[0].cells
    headers2 = ['Görsel', 'Tarih / Kaynak', 'Kategori', 'Haber Başlığı (Tıklanabilir Link)', 'Düzyazı Haber Özeti ve Analizi', 'Duygu / Skor']
    for i, title in enumerate(headers2):
        hdr_cells2[i].text = title
        style_table_cell(hdr_cells2[i], bg_hex="1B365D", bold=True, font_size=9, color_rgb=(255, 255, 255))

    for _, r in df_all.iterrows():
        rc = table2.add_row().cells
        
        img_stream = download_image_stream(r.get('Görsel_URL', ''))
        if img_stream:
            try:
                p_img = rc[0].paragraphs[0]
                p_img.add_run().add_picture(img_stream, width=Inches(1.2))
            except Exception:
                rc[0].text = "Görsel Yüklenemedi"
        else:
            rc[0].text = "Görsel Yok"

        rc[1].text = f"{r['Tarih']}\n{r['Kaynak']}"
        rc[2].text = r['Kategori']
        
        p_link2 = rc[3].paragraphs[0]
        add_safe_hyperlink(p_link2, r['URL'], r['Başlık'])
        
        rc[4].text = r['Özet']
        rc[5].text = f"{r['Duygu']} ({r['Skor']})"
        
        for c in rc:
            style_table_cell(c, font_size=8)
        
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# --- ARAYÜZ (STREAMLIT) ---
st.title("🛡️ Sanayi, Teknoloji & Güvenlik Açık Kaynak Radarı")
st.caption("Google Link Çözümleme, Orijinal Görsel Yakalama ve Düzyazı Analiz Platformu")

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

    max_news = st.slider("Maksimum Haber Sayısı:", 10, 40, 20)
    only_negative = st.checkbox("Sadece Negatif/Riskli Haberleri Ekrana Getir", value=False)
    
    btn_run = st.button("🔍 Açık Kaynak Taramasını Başlat", type="primary", use_container_width=True)

if btn_run:
    with st.spinner("Google yönlendirme linkleri çözülüyor, orijinal siteden görseller indiriliyor ve düzyazı analizler hazırlanıyor..."):
        articles = fetch_robust_news(query, time_range=time_filter, max_results=max_news)
        
        if articles:
            parsed_data = []
            for a in articles:
                title = a.get('title', '') or ''
                full_text = a.get('full_text', '') or ''
                
                score, sentiment, risk, manip_words, category = analyze_article(title, full_text)
                
                prose_summary = build_prose_analysis(
                    title, full_text, category, sentiment, risk, manip_words
                )
                
                parsed_data.append({
                    "Tarih": a.get('publishedAt', ''),
                    "Kaynak": a.get('source', {}).get('name', 'Bilinmiyor'),
                    "Kategori": category,
                    "Başlık": title,
                    "Özet": prose_summary,
                    "Duygu": sentiment,
                    "Skor": score,
                    "Risk_Durumu": risk,
                    "Manipülasyon_Kelimeleri": manip_words,
                    "URL": a.get('url', ''),
                    "Görsel_URL": a.get('image_url', '')
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
                st.warning(f"Toplam {len(risk_df_display)} haberde manipülatif dil/yüksek negatiflik tespit edilmiştir!")
                st.dataframe(
                    risk_df_display[['Tarih', 'Kaynak', 'Kategori', 'Başlık', 'Risk_Durumu', 'URL']],
                    column_config={"URL": st.column_config.LinkColumn("Orijinal Haber Linki")},
                    use_container_width=True
                )
            else:
                st.success("Seçilen zaman aralığında kritik düzeyde manipülatif dil barındıran haber bulunamamıştır.")
            
            st.subheader("📋 Canlı Haber Akışı ve Düzyazı Analiz Tablosu")
            st.dataframe(
                df[['Tarih', 'Kaynak', 'Kategori', 'Başlık', 'Özet', 'Duygu', 'Skor', 'Risk_Durumu', 'URL']],
                column_config={"URL": st.column_config.LinkColumn("Orijinal Haber Linki")},
                use_container_width=True
            )
            
            stats_dict = {'total': tot, 'neg': neg_cnt, 'risk_count': risk_cnt, 'neg_ratio': neg_ratio}
            docx_b = generate_osint_docx(query, df, stats_dict)
            
            st.download_button(
                label="📄 DÜZYAZI ANALİZLİ AÇIK KAYNAK RAPORUNU İNDİR (.DOCX)",
                data=docx_b,
                file_name=f"Acik_Kaynak_Raporu_{date.today()}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary"
            )
        else:
            st.info("Seçilen zaman diliminde kriterlere uygun haber bulunamamıştır.")