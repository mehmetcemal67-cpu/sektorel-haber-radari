import streamlit as st
import pandas as pd
import requests
import time
import datetime
from datetime import date, datetime
from io import BytesIO
import re
import html
import concurrent.futures

from bs4 import BeautifulSoup
from PIL import Image

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
    "karbon engeli", "lisans reddi", "kırmızı çizgi", "blokaj", "provokasyon",
    "soruşturma", "denetim", "para cezası", "greve gitti", "işten çıkarma",
    "geri çekildi", "ertelendi", "davası açıldı"
]

# Hedefli negatif/riskli haber yakalama için ana terimlerle birleştirilen kelime havuzu
NEGATIVE_BOOST_KEYWORDS = [
    "skandal", "kriz", "iptal", "arıza", "zarar", "ambargo", "fiyasko",
    "üretim durdu", "gecikme", "soruşturma", "yaptırım", "iflas"
]

# "Sanayi ve teknolojiyle ilgili her şey" isteğini karşılamak için geniş terim bankası.
# Kullanıcının yazdığı sorguya otomatik olarak eklenir (Kapsamlı Tarama Modu açıkken).
BROAD_TERM_BANK = [
    "sanayi", "teknoloji", "yerli üretim", "millileştirme", "Ar-Ge", "inovasyon",
    "imalat sanayi", "üretim hattı", "dijitalleşme", "endüstri 4.0", "otomasyon",
    "robotik", "yapay zeka", "makine öğrenmesi", "siber güvenlik", "yarı iletken",
    "çip üretimi", "batarya teknolojisi", "elektrikli araç", "şarj altyapısı",
    "yenilenebilir enerji", "güneş enerjisi", "rüzgar enerjisi", "hidrojen enerjisi",
    "savunma sanayii", "havacılık ve uzay", "roket teknolojisi", "insansız hava aracı",
    "telekomünikasyon", "5G", "veri merkezi", "bulut bilişim", "kuantum teknoloji",
    "nanoteknoloji", "biyoteknoloji", "3 boyutlu yazıcı", "akıllı fabrika",
    "ihracat rakamları", "start-up", "girişim sermayesi", "teknopark", "patent başvurusu"
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

# --- METİN VE DUYGU ANALİZİ ---
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

# --- DÜZYAZI ANALİZ ÜRETİCİ (TAM METİN ÖZETİ) ---
def build_prose_analysis(title, body_text, category, sentiment, risk_level, manip_words):
    """
    Artık sadece ilk 4 cümleyle sınırlı değil: elde edilen metnin (mümkünse tam
    haber metninin) anlamlı cümlelerini, makul bir uzunluğa (yaklaşık 900 karakter /
    en fazla 12 cümle) ulaşana kadar art arda ekleyerek haberin bütününü özetler.
    """
    clean_text = (body_text or "").strip()
    sentences = [s.strip() for s in re.split(r'(?<=[.!?])\s+', clean_text) if len(s.strip()) > 20]

    story_parts = []
    total_len = 0
    for s in sentences:
        if total_len > 900 or len(story_parts) >= 12:
            break
        story_parts.append(s)
        total_len += len(s)

    if story_parts:
        story = " ".join(story_parts)
    else:
        story = f"{title} hususunda basına yansıyan detaylar açık kaynak akışı üzerinden takip edilmektedir."

    prose_eval = f" Yapılan değerlendirmeye göre gelişme {category.lower()} ekosistemini doğrudan ilgilendirmektedir."

    if sentiment == "Pozitif":
        prose_eval += " İçerikte sunulan veriler; yerli üretim yetkinliklerinin gelişimi, sektörel büyüme ve kurumsal altyapının güçlenmesi yönünde olumlu mesajlar vermektedir."
    elif sentiment == "Negatif":
        prose_eval += " Detaylar incelendiğinde; pazardaki daralma, maliyet artışları veya operasyonel risklerin ön plana çıktığı görülmektedir."
    else:
        prose_eval += " Haber içeriği genel itibarıyla teknik bilgilendirme ve kurumsal süreç aktarımı niteliğindedir."

    if risk_level == "Yüksek Risk":
        prose_eval += f" Ek olarak metin içerisinde algı oluşturma veya kamuoyunu yönlendirme riski barındıran ifadelere ({', '.join(manip_words)}) rastlanmış olup konunun kurumsal takibi önerilmektedir."

    return f"{story}{prose_eval}"

# --- SORGU AYRIŞTIRMA VE GENİŞLETME ---
def build_search_terms(query_text, max_terms=20):
    raw_terms = re.split(r'\bOR\b|,|\n', query_text, flags=re.IGNORECASE)
    cleaned = []
    seen = set()
    for t in raw_terms:
        t = t.strip().strip('"').strip("'")
        if len(t) < 2:
            continue
        low = t.lower()
        if low in seen:
            continue
        seen.add(low)
        cleaned.append(t)
        if len(cleaned) >= max_terms:
            break
    return cleaned if cleaned else ["sanayi teknoloji"]

def merge_with_bank(base_terms, bank, cap=35):
    combined = list(base_terms)
    seen = set(t.lower() for t in combined)
    for b in bank:
        if len(combined) >= cap:
            break
        if b.lower() not in seen:
            combined.append(b)
            seen.add(b.lower())
    return combined

def normalize_title(t):
    t = (t or "").lower()
    t = re.sub(r'[^\w\s]', ' ', t)
    t = re.sub(r'\s+', ' ', t).strip()
    return t

# --- HABER TOPLAMA MOTORU ---
def fetch_robust_news(query_text, time_range="1d", max_results=50,
                       negative_boost=True, extra_regions=True, broad_mode=True):
    """
    - query_text gerçekten kullanılıyor (OR / virgülle ayrılmış terimlere bölünür).
    - broad_mode=True: kullanıcı sorgusu, geniş sanayi/teknoloji terim bankasıyla
      birleştirilir -> "her şey girsin" isteği için kapsamı otomatik genişletir.
    - negative_boost=True: her ana terim, 2 farklı kritik/negatif kelimeyle
      birleştirilip ayrıca aranır -> negatif/riskli haberi yakalama olasılığı artar.
    - extra_regions=True: bazı terimler 'wt-wt' (dünya geneli) bölgesinde de aranır
      -> yabancı basın (ör. Yunan basını) da yakalanabilir.
    """
    articles = []
    seen_urls = set()
    seen_titles = set()
    time_ddg = {"1d": "d", "7d": "w", "14d": "w"}.get(time_range, "d")

    base_terms = build_search_terms(query_text, max_terms=20)
    if broad_mode:
        base_terms = merge_with_bank(base_terms, BROAD_TERM_BANK, cap=35)

    search_jobs = [(t, "tr-tr") for t in base_terms]

    if extra_regions:
        for t in base_terms[:6]:
            search_jobs.append((t, "wt-wt"))

    if negative_boost:
        half = max(1, len(NEGATIVE_BOOST_KEYWORDS) // 2)
        for i, t in enumerate(base_terms[:20]):
            kw1 = NEGATIVE_BOOST_KEYWORDS[i % len(NEGATIVE_BOOST_KEYWORDS)]
            kw2 = NEGATIVE_BOOST_KEYWORDS[(i + half) % len(NEGATIVE_BOOST_KEYWORDS)]
            search_jobs.append((f"{t} {kw1}", "tr-tr"))
            search_jobs.append((f"{t} {kw2}", "tr-tr"))

    search_jobs = search_jobs[:130]  # güvenlik / rate-limit sınırı

    try:
        from ddgs import DDGS
    except ImportError:
        from duckduckgo_search import DDGS

    try:
        with DDGS() as ddgs:
            for term, region in search_jobs:
                if len(articles) >= max_results:
                    break
                try:
                    results = ddgs.news(keywords=term, region=region, timelimit=time_ddg, max_results=10)
                except Exception:
                    time.sleep(1.0)
                    continue

                for r in results or []:
                    url = r.get('url', '')
                    title = r.get('title', '') or ''
                    norm_title = normalize_title(title)
                    title_key = " ".join(norm_title.split()[:8])

                    if not url or url in seen_urls:
                        continue
                    if title_key and title_key in seen_titles:
                        continue

                    seen_urls.add(url)
                    if title_key:
                        seen_titles.add(title_key)

                    raw_date = r.get('date', '')
                    try:
                        dt = datetime.fromisoformat(raw_date.replace('Z', '+00:00'))
                        formatted_date = dt.strftime('%d %b %Y %H:%M')
                    except Exception:
                        formatted_date = datetime.now().strftime('%d %b %Y %H:%M')

                    articles.append({
                        'title': title,
                        'body': r.get('body', ''),
                        'url': url,
                        'image_url': r.get('image', ''),
                        'publishedAt': formatted_date,
                        'source': {'name': r.get('source', 'Açık Basın')}
                    })

                    if len(articles) >= max_results:
                        break

                time.sleep(0.3)
    except Exception:
        pass

    return articles[:max_results]

# --- TAM HABER METNİ ÇEKME (haberin tamamının özeti için) ---
def fetch_article_fulltext(url, timeout=5):
    if not url:
        return ""
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        resp = requests.get(url, headers=headers, timeout=timeout)
        if resp.status_code != 200 or not resp.text:
            return ""
        soup = BeautifulSoup(resp.text, "html.parser")
        for tag in soup(["script", "style", "nav", "footer", "header", "aside", "form", "iframe"]):
            tag.decompose()
        paragraphs = soup.find_all("p")
        parts = [p.get_text(" ", strip=True) for p in paragraphs]
        parts = [t for t in parts if len(t) > 40]
        full_text = " ".join(parts)
        return full_text[:6000]
    except Exception:
        return ""

def fetch_fulltexts_parallel(urls, max_workers=8, per_call_timeout=5):
    """Haber tam metinlerini paralel olarak çeker (çok sayıda haberde makul sürede tamamlanması için)."""
    results = {}
    unique_urls = [u for u in dict.fromkeys(urls) if u]
    if not unique_urls:
        return results
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as ex:
        future_map = {ex.submit(fetch_article_fulltext, u, per_call_timeout): u for u in unique_urls}
        for fut in concurrent.futures.as_completed(future_map):
            u = future_map[fut]
            try:
                results[u] = fut.result()
            except Exception:
                results[u] = ""
    return results

# --- WORD DOKÜMANI YARDIMCILARI ---
def add_safe_hyperlink(paragraph, url, text):
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
        paragraph.add_run(f"{text} ({url})")

def download_image_stream(img_url):
    if not img_url or "gstatic.com" in img_url:
        return None
    try:
        headers = {'User-Agent': 'Mozilla/5.0'}
        resp = requests.get(img_url, headers=headers, timeout=6)
        content_type = resp.headers.get('Content-Type', '')
        if resp.status_code == 200 and len(resp.content) > 1500 and content_type.startswith('image'):
            img = Image.open(BytesIO(resp.content))
            if img.mode in ("RGBA", "P", "LA"):
                img = img.convert("RGB")
            out = BytesIO()
            img.save(out, format="JPEG", quality=85)
            out.seek(0)
            return out
    except Exception:
        pass
    return None

def style_table_cell(cell, bg_hex=None, bold=False, font_size=8.5, color_rgb=(0, 0, 0)):
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

# --- WORD RAPOR OLUŞTURUCU (TAM TARAMA RAPORU) ---
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
        headers = ['Görsel', 'Tarih / Kaynak', 'Kategori', 'Haber Başlığı (Linkli)', 'Haberin Tam Özeti & Analiz', 'Tespit Edilen Söylem']
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
    headers2 = ['Görsel', 'Tarih / Kaynak', 'Kategori', 'Haber Başlığı (Linkli)', 'Haberin Tam Özeti & Analiz', 'Duygu / Skor']
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

# --- GÜNLÜK BİLGİ NOTU OLUŞTURUCU (seçili haberlerden) ---
def generate_briefing_note_docx(selected_df, query):
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.7)
        section.right_margin = Inches(0.7)

    h = doc.add_heading("GÜNLÜK BİLGİ NOTU", level=0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in h.runs:
        r.font.name = 'Arial'
        r.font.size = Pt(14)
        r.font.bold = True
        r.font.color.rgb = RGBColor(27, 54, 93)

    p_meta = doc.add_paragraph()
    p_meta.add_run("KONU: ").bold = True
    p_meta.add_run("Sanayi ve Teknoloji Alanında Açık Kaynak Değerlendirmesi\n")
    p_meta.add_run("TARİH: ").bold = True
    p_meta.add_run(f"{datetime.now().strftime('%d.%m.%Y %H:%M')}\n")
    p_meta.add_run("TARAMA KAPSAMI: ").bold = True
    p_meta.add_run(f"{query}\n")
    p_meta.add_run("SEÇİLEN HABER SAYISI: ").bold = True
    p_meta.add_run(f"{len(selected_df)}\n")
    for r in p_meta.runs:
        r.font.name = 'Arial'
        r.font.size = Pt(9.5)

    doc.add_paragraph()

    for idx, (_, row) in enumerate(selected_df.iterrows(), start=1):
        p_title = doc.add_paragraph()
        run_t = p_title.add_run(f"{idx}. {row['Başlık']}")
        run_t.bold = True
        run_t.font.name = 'Arial'
        run_t.font.size = Pt(11)
        run_t.font.color.rgb = RGBColor(27, 54, 93)

        p_meta2 = doc.add_paragraph()
        p_meta2.add_run(
            f"{row['Tarih']}  |  {row['Kaynak']}  |  {row['Kategori']}  |  "
            f"Duygu: {row['Duygu']} ({row['Skor']})  |  {row['Risk_Durumu']}"
        )
        for r in p_meta2.runs:
            r.font.name = 'Arial'
            r.font.size = Pt(8.5)
            r.italic = True

        img_stream = download_image_stream(row.get('Görsel_URL', ''))
        if img_stream:
            try:
                p_img = doc.add_paragraph()
                p_img.add_run().add_picture(img_stream, width=Inches(2.3))
            except Exception:
                pass

        p_sum = doc.add_paragraph()
        p_sum.add_run(row['Özet'])
        for r in p_sum.runs:
            r.font.name = 'Arial'
            r.font.size = Pt(9.5)

        p_link = doc.add_paragraph()
        p_link.add_run("Kaynak: ").bold = True
        add_safe_hyperlink(p_link, row['URL'], row['URL'])
        for r in p_link.runs:
            r.font.name = 'Arial'
            r.font.size = Pt(8.5)

        doc.add_paragraph("─" * 70)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# --- ARAYÜZ (STREAMLIT) ---
st.title("🛡️ Sanayi, Teknoloji & Güvenlik Açık Kaynak Radarı")
st.caption("Geniş Kapsamlı Anlık Tarama · Negatif/Riskli Ayrıştırma · Tam Metin Özeti · Günlük Bilgi Notu")

with st.sidebar:
    st.header("⚙️ Tarama Parametreleri")

    default_query = (
        'sanayi OR teknoloji OR TOGG OR KAAN OR ASELSAN OR BAYKAR OR TUSAŞ OR ROKETSAN OR HAVELSAN OR '
        'TÜBİTAK OR KOSGEB OR Çelik Kubbe OR SİHA OR İHA OR Milli Teknoloji OR çip OR Yapay Zeka OR '
        'Hisar OR Siper OR Türkiye Uzay Ajansı OR Teknofest OR şarj istasyonu OR batarya OR '
        'organize sanayi bölgesi OR yatırım teşvik OR Ege OR Doğu Akdeniz OR Yunanistan OR '
        'imalat sanayi OR Ar-Ge OR ihracat OR enerji OR siber güvenlik'
    )

    query = st.text_area("Arama Sorgusu (OR / virgül ile ayırın):", value=default_query, height=140)

    broad_mode = st.checkbox(
        "Kapsamlı Tarama Modu (Sanayi/Teknolojiyle İlgili Her Şeyi Tara)",
        value=True,
        help="Yazdığınız sorguya ek olarak, geniş bir sanayi/teknoloji terim bankasını otomatik olarak tarar."
    )

    time_filter = st.selectbox(
        "Zaman Dilimi (Canlı/Anlık Akış):",
        options=["1d", "7d", "14d"],
        format_func=lambda x: {"1d": "Son 24 Saat (Anlık)", "7d": "Son 1 Hafta", "14d": "Son 2 Hafta"}[x]
    )

    max_news = st.slider("Maksimum Haber Sayısı:", 20, 150, 60)

    negative_boost = st.checkbox(
        "Negatif/Riskli Haber Yakalama Modunu Güçlendir",
        value=True,
        help="Ana terimleri 'skandal, kriz, iptal, arıza...' gibi kritik kelimelerle birleştirip ek aramalar yapar."
    )
    extra_regions = st.checkbox(
        "Yabancı Basını da Tara (Bölgesel Güvenlik için önerilir)",
        value=True
    )
    full_text_mode = st.checkbox(
        "Tam Metin ile Zenginleştir (Haberin Tamamının Özeti — Daha Yavaş)",
        value=True,
        help="Her haberin orijinal sayfasına gidip tam metnini çeker; özet bu metinden üretilir. Kapalıyken sadece kısa haber özeti kullanılır (daha hızlı)."
    )

    btn_run = st.button("🔍 Taramayı Başlat", type="primary", use_container_width=True)

if 'df_full' not in st.session_state:
    st.session_state['df_full'] = None
    st.session_state['query_used'] = ""

if btn_run:
    with st.spinner("Haberler geniş kapsamda taranıyor..."):
        articles = fetch_robust_news(
            query,
            time_range=time_filter,
            max_results=max_news,
            negative_boost=negative_boost,
            extra_regions=extra_regions,
            broad_mode=broad_mode
        )

    if articles:
        fulltext_map = {}
        if full_text_mode:
            with st.spinner(f"{len(articles)} haberin tam metni çekiliyor (haberin tamamının özeti için)..."):
                fulltext_map = fetch_fulltexts_parallel([a.get('url', '') for a in articles])

        with st.spinner("Duygu/risk analizi ve düzyazı özetler oluşturuluyor..."):
            parsed_data = []
            for a in articles:
                title = a.get('title', '') or ''
                snippet = a.get('body', '') or ''
                url = a.get('url', '')
                fulltext = fulltext_map.get(url, '')
                body_for_analysis = fulltext if (len(fulltext) > len(snippet) and len(fulltext) > 150) else snippet

                score, sentiment, risk, manip_words, category = analyze_article(title, body_for_analysis)
                prose_summary = build_prose_analysis(title, body_for_analysis, category, sentiment, risk, manip_words)

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
                    "URL": url,
                    "Görsel_URL": a.get('image_url', '')
                })

            df = pd.DataFrame(parsed_data)
            df['Seç'] = df['Risk_Durumu'].eq('Yüksek Risk') | df['Duygu'].eq('Negatif')

        st.session_state['df_full'] = df
        st.session_state['query_used'] = query
    else:
        st.session_state['df_full'] = pd.DataFrame()
        st.info("Seçilen zaman diliminde kriterlere uygun haber bulunamamıştır. Sorguyu genişletmeyi veya zaman dilimini büyütmeyi deneyin.")

df = st.session_state.get('df_full')

if df is not None and not df.empty:
    tot = len(df)
    neg_cnt = int((df['Duygu'] == 'Negatif').sum())
    risk_cnt = int((df['Risk_Durumu'] == 'Yüksek Risk').sum())
    neg_ratio = (neg_cnt / tot * 100) if tot > 0 else 0

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("İncelenen Haber", tot)
    k2.metric("Negatif Haberler", neg_cnt)
    k3.metric("Manipülasyon Riskli", risk_cnt, delta="Kritik Dil" if risk_cnt > 0 else "Normal", delta_color="inverse")
    k4.metric("Negatif Haber Oranı", f"%{neg_ratio:.1f}")

    st.markdown("---")

    tab_all, tab_neg, tab_risk = st.tabs([
        f"📋 Tüm Haberler ({tot})",
        f"⚠️ Negatif Haberler ({neg_cnt})",
        f"🚨 Yüksek Riskli Haberler ({risk_cnt})"
    ])

    display_cols = ['Seç', 'Tarih', 'Kaynak', 'Kategori', 'Başlık', 'Özet', 'Duygu', 'Skor', 'Risk_Durumu', 'URL']

    with tab_all:
        st.caption("Bilgi notuna eklemek istediğiniz haberleri 'Seç' kutucuğuyla işaretleyin (riskli/negatif haberler otomatik işaretlenmiştir).")
        edited_df = st.data_editor(
            df[display_cols],
            column_config={
                "Seç": st.column_config.CheckboxColumn("Bilgi Notuna Ekle"),
                "URL": st.column_config.LinkColumn("Haber Linki"),
            },
            disabled=[c for c in display_cols if c != 'Seç'],
            hide_index=True,
            use_container_width=True,
            key="all_news_editor"
        )
        df.loc[edited_df.index, 'Seç'] = edited_df['Seç']

    with tab_neg:
        neg_df = df[df['Duygu'] == 'Negatif']
        if neg_df.empty:
            st.info("Negatif haber tespit edilmedi.")
        else:
            st.dataframe(
                neg_df[['Tarih', 'Kaynak', 'Kategori', 'Başlık', 'Özet', 'Skor', 'URL']],
                column_config={"URL": st.column_config.LinkColumn("Haber Linki")},
                hide_index=True,
                use_container_width=True
            )

    with tab_risk:
        risk_df_view = df[df['Risk_Durumu'] == 'Yüksek Risk']
        if risk_df_view.empty:
            st.info("Yüksek riskli/manipülatif söylem içeren haber tespit edilmedi.")
        else:
            st.dataframe(
                risk_df_view[['Tarih', 'Kaynak', 'Kategori', 'Başlık', 'Özet', 'Manipülasyon_Kelimeleri', 'URL']],
                column_config={"URL": st.column_config.LinkColumn("Haber Linki")},
                hide_index=True,
                use_container_width=True
            )

    st.markdown("---")
    st.subheader("📥 Çıktılar")

    c1, c2 = st.columns(2)

    with c1:
        stats_dict = {'total': tot, 'neg': neg_cnt, 'risk_count': risk_cnt, 'neg_ratio': neg_ratio}
        docx_full = generate_osint_docx(st.session_state['query_used'], df, stats_dict)
        st.download_button(
            label="📄 TAM TARAMA RAPORUNU İNDİR (.DOCX)",
            data=docx_full,
            file_name=f"Acik_Kaynak_Raporu_{date.today()}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )

    with c2:
        selected_df = df[df['Seç'] == True]
        st.write(f"Bilgi notuna eklenmek üzere **{len(selected_df)}** haber seçili.")
        if len(selected_df) > 0:
            briefing_docx = generate_briefing_note_docx(selected_df, st.session_state['query_used'])
            st.download_button(
                label="📝 SEÇİLİ HABERLERDEN GÜNLÜK BİLGİ NOTU İNDİR (.DOCX)",
                data=briefing_docx,
                file_name=f"Gunluk_Bilgi_Notu_{date.today()}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        else:
            st.button("📝 SEÇİLİ HABERLERDEN GÜNLÜK BİLGİ NOTU İNDİR (.DOCX)", disabled=True, use_container_width=True)

elif df is not None and df.empty:
    pass
else:
    st.info("Taramayı başlatmak için soldaki panelden parametreleri ayarlayıp '🔍 Taramayı Başlat' butonuna basın.")