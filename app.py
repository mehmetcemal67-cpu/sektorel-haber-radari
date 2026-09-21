import streamlit as st
import pandas as pd
import requests
import concurrent.futures
import time
import xml.etree.ElementTree as ET
import re, html, json
import sqlite3, hashlib
from pathlib import Path
from datetime import datetime, timedelta, timezone, date
from urllib.parse import urlparse
from email.utils import parsedate_to_datetime
from io import BytesIO
from bs4 import BeautifulSoup
from PIL import Image
from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import qn

# ============================================================
# STREAMLIT SAYFA YAPILANDIRMASI VE ŞİFRE KORUMASI
# ============================================================
st.set_page_config(
    page_title="Sanayi & Teknoloji OSINT Radarı",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

def _v55_password_gate():
    try:
        expected = str(st.secrets["APP_PASSWORD"])
    except Exception:
        # Secrets yapılandırılmamışsa sunumun kesilmemesi için erişime izin ver
        return

    if st.session_state.get("_v55_authenticated", False):
        return

    st.title("🔐 Sanayi ve Teknoloji OSINT Radar")
    st.caption("Devam etmek için lütfen uygulama şifresini girin.")

    with st.form("_v55_login_form", clear_on_submit=False):
        entered = st.text_input("Şifre", type="password")
        submitted = st.form_submit_button("Giriş Yap", use_container_width=True)

    if submitted:
        import hmac
        if hmac.compare_digest(str(entered), expected):
            st.session_state["_v55_authenticated"] = True
            st.rerun()
        else:
            st.error("Şifre hatalı. Lütfen tekrar deneyin.")

    st.stop()

_v55_password_gate()

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}

# ============================================================
# SİSTEM VE METİN TEMİZLEME YARDIMCILARI
# ============================================================
def norm(s):
    return re.sub(r'\s+', ' ', str(s or '').lower()).strip()

def title_key(s):
    return re.sub(r'[^\w\s]', ' ', norm(s)).strip()[:180]

def domain(url):
    try:
        return urlparse(url).netloc.lower().replace('www.', '')
    except Exception:
        return ''

def parse_dt(v):
    if not v: return None
    s = str(v).strip()
    for x in (s.replace('Z', '+00:00'), s):
        try:
            d = datetime.fromisoformat(x)
            if d.tzinfo is None: d = d.replace(tzinfo=timezone.utc)
            return d.astimezone(timezone.utc)
        except Exception: pass
    try:
        d = parsedate_to_datetime(s)
        if d.tzinfo is None: d = d.replace(tzinfo=timezone.utc)
        return d.astimezone(timezone.utc)
    except Exception: return None

def _to_utc_datetime(value):
    if value is None: return None
    try:
        if pd.isna(value): return None
    except Exception: pass

    if isinstance(value, pd.Timestamp):
        ts = value
        return ts.tz_localize('UTC').to_pydatetime() if ts.tzinfo is None else ts.tz_convert('UTC').to_pydatetime()

    if isinstance(value, datetime):
        return value.replace(tzinfo=timezone.utc) if value.tzinfo is None else value.astimezone(timezone.utc)

    try:
        ts = pd.to_datetime(value, utc=True, errors='coerce')
        return None if pd.isna(ts) else ts.to_pydatetime()
    except Exception:
        return None

def fmt_dt(d):
    d = _to_utc_datetime(d)
    return d.astimezone().strftime('%d.%m.%Y %H:%M:%S') if d else 'Tarih/saat bilinmiyor'

def _repair_mojibake_utf8(text):
    s = str(text or '')
    if not s: return s
    suspicious = ('Ã', 'Ä', 'Å', 'Â', 'â€', 'â€™', 'â€œ', 'â€ ', 'â€“', 'â€”')
    if not any(x in s for x in suspicious): return s
    fixes = {
        'TÃ¼rkiye': 'Türkiye', 'TÃ¼rk': 'Türk', 'genÃ§': 'genç', 'dÃ¼nya': 'dünya',
        'Ã¼lke': 'ülke', 'Ã§': 'ç', 'ÄŸ': 'ğ', 'Ä±': 'ı', 'Ã¶': 'ö', 'Ã¼': 'ü', 'ÅŸ': 'ş',
        'Â': '', 'â€™': '’', 'â€œ': '“', 'â€ ': '”', 'â€“': '–', 'â€”': '—'
    }
    for bad, good in fixes.items():
        s = s.replace(bad, good)
    return s

def _clean_note_text(value):
    text = BeautifulSoup(str(value or ''), 'html.parser').get_text(' ', strip=True)
    text = html.unescape(text)
    text = _repair_mojibake_utf8(text)
    text = ''.join(ch for ch in text if ch in ('\t', '\n', '\r') or ord(ch) >= 32)
    return re.sub(r'\s+', ' ', text).strip()

def _sentence_chunks(text):
    txt = _clean_note_text(text)
    if not txt: return []
    return [x.strip() for x in re.split(r'(?<=[.!?;:])\s+', txt) if x.strip()]

def _v66_formalize_sentence_endings(text):
    t = re.sub(r'\s+', ' ', str(text or '')).strip()
    if not t: return t
    pairs = [
        ('açıkladı', 'açıklamıştır'), ('belirtti', 'belirtmiştir'), ('bildirdi', 'bildirmiştir'),
        ('duyurdu', 'duyurmuştur'), ('ifade etti', 'ifade etmiştir'), ('kaydetti', 'kaydetmiştir'),
        ('başladı', 'başlamıştır'), ('tamamlandı', 'tamamlanmıştır'), ('gerçekleşti', 'gerçekleşmiştir'),
        ('yükseldi', 'yükselmiştir'), ('geriledi', 'gerilemiştir'), ('arttı', 'artmıştır'),
        ('azaldı', 'azalmıştır'), ('devam ediyor', 'devam etmektedir'), ('sağlıyor', 'sağlamaktadır')
    ]
    parts = _sentence_chunks(t)
    out = []
    for s in parts:
        s = s.strip()
        if not s: continue
        punct = s[-1] if s[-1] in '.!?' else '.'
        core = s[:-1].rstrip() if s[-1] in '.!?' else s
        low = core.lower()
        for old, newv in pairs:
            if low.endswith(old):
                core = core[:-len(old)] + newv
                break
        out.append(core.rstrip(' .;:') + punct)
    return ' '.join(out)

# ============================================================
# DOKÜMAN VE RAPOR ÜRETİM MODÜLLERİ (ÖGN, AKT, BİLGİ NOTU)
# ============================================================
def _word_hyperlink(paragraph, url, label):
    if not url:
        paragraph.add_run(label)
        return
    try:
        rid = paragraph.part.relate_to(
            url,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            is_external=True
        )
        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set(qn("r:id"), rid)
        run = OxmlElement("w:r")
        rpr = OxmlElement("w:rPr")
        rstyle = OxmlElement("w:rStyle")
        rstyle.set(qn("w:val"), "Hyperlink")
        rpr.append(rstyle)
        run.append(rpr)
        text = OxmlElement("w:t")
        text.text = label
        run.append(text)
        hyperlink.append(run)
        paragraph._p.append(hyperlink)
    except Exception:
        paragraph.add_run(label)

def _download_report_image(url):
    if not url: return None
    try:
        rr = requests.get(url, headers=HEADERS, timeout=4)
        if rr.status_code != 200 or len(rr.content) < 1000: return None
        im = Image.open(BytesIO(rr.content))
        if im.mode not in ("RGB", "L"): im = im.convert("RGB")
        im.thumbnail((1200, 900), Image.LANCZOS)
        bio = BytesIO()
        im.save(bio, "JPEG", quality=85)
        bio.seek(0)
        return bio
    except Exception:
        return None

def article_detail(row):
    url = str(row.get("URL") or row.get("url") or "").strip()
    title = str(row.get("Başlık") or row.get("title") or "").strip()
    snippet = str(row.get("İçerik_Özeti") or row.get("summary") or "").strip()
    source = str(row.get("Kaynak") or row.get("source") or "Açık Kaynak").strip()

    out = {
        "title": title,
        "canonical": url,
        "published": str(row.get("Tarih") or ""),
        "text": snippet,
        "images": [],
        "source": source
    }
    if not url or url.startswith("javascript"): return out

    try:
        rr = requests.get(url, headers=HEADERS, timeout=4)
        if rr.status_code == 200 and rr.text:
            soup = BeautifulSoup(rr.text, "html.parser")
            paragraphs = [p.get_text(" ", strip=True) for p in soup.find_all("p") if len(p.get_text()) > 40]
            if paragraphs:
                out["text"] = " ".join(paragraphs[:8])
            for img in soup.find_all("img"):
                src = img.get("src") or img.get("data-src")
                if src and src.startswith("http") and not any(x in src.lower() for x in ["icon", "logo", "avatar"]):
                    out["images"].append(src)
    except Exception:
        pass
    return out

# --- 1. ÖNEMLİ GELİŞMELER NOTU (ÖGN) ---
def make_important_basket_docx(basket_df):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(2); sec.bottom_margin = Cm(2)
    sec.left_margin = Cm(2.5); sec.right_margin = Cm(2.5)

    normal = doc.styles['Normal']
    normal.font.name = 'Times New Roman'
    normal.font.size = Pt(12)

    now = datetime.now().astimezone()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.add_run(now.strftime('%d/%m/%Y'))

    p = doc.add_paragraph()
    p.add_run('Konu: ').bold = True
    p.add_run('STB Temsilciliği Önemli Gelişmeler Notu')

    rows = [] if basket_df is None or basket_df.empty else basket_df.to_dict('records')

    for r in rows:
        title = _clean_note_text(r.get('title') or r.get('Başlık') or '')
        summary = _clean_note_text(r.get('summary') or r.get('İçerik_Özeti') or title)
        
        # Resmî kurum formatında 2-3 cümlelik akıcı paragraf kurgulanır
        sents = _sentence_chunks(summary)
        chosen = sents[:3] if sents else [title]
        text_body = ' '.join(chosen)
        formal_text = _v66_formalize_sentence_endings(text_body)

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_after = Pt(7)
        # Haber başlığı kaldırılmış, doğrudan kurumsal metin eklenmiştir
        p.add_run(formal_text.rstrip(' .;') + ' (STB).')

    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(12)
    p_end.add_run('Arz olunur.')

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()

# --- 2. AÇIK KAYNAK TARAMA RAPORU (AKT) ---
def make_docx(rows):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(2.0); sec.bottom_margin = Cm(2.0)
    sec.left_margin = Cm(2.5); sec.right_margin = Cm(2.5)

    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run("AÇIK KAYNAK TARAMA ÇALIŞMASI")
    r.bold = True
    r.font.size = Pt(14)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    p.add_run("Tarama Yapılan Görev Alanı: ").bold = True
    p.add_run("Sanayi ve Teknoloji")

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(8)
    p.add_run("Tarih: ").bold = True
    p.add_run(datetime.now().astimezone().strftime("%d.%m.%Y"))

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    p.add_run("Bulgular: ").bold = True
    p.add_run("Sanayi ve Teknoloji alanlarında yapılan açık kaynak taraması neticesinde tespit edilen haber detayları aşağıda sunulmuştur.")

    for i, row in enumerate(rows, 1):
        detail = article_detail(row)
        real_url = detail.get("canonical") or row.get("URL") or ""
        title = _clean_note_text(detail.get("title") or row.get("Başlık") or "")
        source = _clean_note_text(detail.get("source") or row.get("Kaynak") or "Açık Kaynak")
        body = detail.get("text") or title
        summary = _v66_formalize_sentence_endings(' '.join(_sentence_chunks(body)[:4]))

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Cm(0.75)
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(6)

        p.add_run(f"{i}. ").bold = True
        p.add_run(f'“{source}”').bold = True
        p.add_run(' isimli internet sitesinde, ')
        p.add_run(f'“{title}”').bold = True
        p.add_run(' başlığıyla bir haber yayımlanmıştır. (')
        _word_hyperlink(p, real_url, "Haber Linki")
        p.add_run(') Söz konusu haber içeriğinde, ')
        p.add_run(summary)
        p.add_run(' hususları ifade edilmiştir.')

        # Görsel İndirme
        img_stream = None
        for candidate in detail.get("images", []):
            img_stream = _download_report_image(candidate)
            if img_stream: break

        if img_stream:
            cap = doc.add_paragraph()
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cap.paragraph_format.space_before = Pt(4)
            cap.paragraph_format.space_after = Pt(4)
            cr = cap.add_run(f'Görsel {i}: “{source}” Sitesinde Yer Alan Görsel')
            cr.bold = True
            cr.font.size = Pt(10)

            ip = doc.add_paragraph()
            ip.alignment = WD_ALIGN_PARAGRAPH.CENTER
            ip.paragraph_format.space_after = Pt(10)
            ip.add_run().add_picture(img_stream, width=Cm(14))

    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(12)
    p_end.add_run("Arz olunur.")

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()

# --- 3. BİLGİ NOTU ---
def make_analyst_docx(df, title='BİLGİ NOTU'):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(2); sec.bottom_margin = Cm(2)
    sec.left_margin = Cm(2.5); sec.right_margin = Cm(2.5)

    styles = doc.styles
    styles['Normal'].font.name = 'Times New Roman'
    styles['Normal'].font.size = Pt(12)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(_clean_note_text(title))
    r.bold = True
    r.font.size = Pt(14)

    p = doc.add_paragraph()
    p.add_run('Tarih: ').bold = True
    p.add_run(datetime.now().astimezone().strftime('%d.%m.%Y'))

    x = df.copy() if df is not None else pd.DataFrame()
    rows = x.to_dict('records')

    def add_body(text):
        bp = doc.add_paragraph()
        bp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        bp.paragraph_format.first_line_indent = Cm(1.25)
        bp.paragraph_format.line_spacing = 1.15
        bp.paragraph_format.space_after = Pt(8)
        bp.add_run(_v66_formalize_sentence_endings(text))

    if rows:
        all_text = " ".join([_clean_note_text(r.get('İçerik_Özeti') or r.get('Başlık') or '') for r in rows])
        sents = _sentence_chunks(all_text)
        
        # 1. Paragraf: Giriş ve Genel Durum
        intro = ' '.join(sents[:2]) if len(sents) >= 2 else all_text
        add_body(f"Açık kaynak tespiti kapsamında; {intro}")

        # 2. Paragraf: Teknik Detay ve Veriler
        if len(sents) > 2:
            body_p = ' '.join(sents[2:8])
            add_body(body_p)

        # 3. Paragraf: Değerlendirme ve Sonuç
        add_body("Mevcut veriler ışığında, konunun ilgili kurumlar nezdinde takip edilmesinin ve olası etkilerinin izlenmesinin faydalı olacağı değerlendirilmektedir.")
    else:
        add_body("Seçilen konuya ilişkin açık kaynak verileri incelenmiştir.")

    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(12)
    p_end.add_run('Arz olunur.')

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()

# ============================================================
# TARAMA VE VERİ MOTORU (PERFORMANS İYİLEŞTİRMELİ)
# ============================================================
@st.cache_data(ttl=900, show_spinner=False)
def fetch_rss_cached(query):
    try:
        r = requests.get(
            'https://news.google.com/rss/search',
            params={'q': query, 'hl': 'tr', 'gl': 'TR', 'ceid': 'TR:tr'},
            headers=HEADERS, timeout=6
        )
        if r.status_code != 200: return []
        root = ET.fromstring(r.content)
        out = []
        for it in root.findall('.//item'):
            src = it.find('source')
            out.append({
                'title': html.unescape(it.findtext('title') or ''),
                'url': it.findtext('link') or '',
                'date': it.findtext('pubDate') or '',
                'snippet': BeautifulSoup(it.findtext('description') or '', 'html.parser').get_text(' ', strip=True),
                'source': src.text if src is not None else '',
                'source_url': src.get('url', '') if src is not None else ''
            })
        return out
    except Exception:
        return []

def run_osint_scan(period_hours=24):
    queries = [
        f'Türkiye (sanayi OR teknoloji OR üretim OR fabrika OR OSB) when:{period_hours}h',
        f'Türkiye (savunma OR ASELSAN OR TUSAŞ OR ROKETSAN OR Baykar) when:{period_hours}h',
        f'Türkiye (siber OR "yapay zeka" OR "yarı iletken" OR çip) when:{period_hours}h'
    ]
    raw_results = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        futures = [executor.submit(fetch_rss_cached, q) for q in queries]
        for f in concurrent.futures.as_completed(futures):
            try: raw_results.extend(f.result())
            except Exception: pass

    seen = set()
    cleaned = []
    for r in raw_results:
        k = title_key(r.get('title'))
        if k and k not in seen:
            seen.add(k)
            dt = parse_dt(r.get('date'))
            cleaned.append({
                'Tarih_dt': dt,
                'Tarih': fmt_dt(dt),
                'Başlık': r.get('title'),
                'İçerik_Özeti': r.get('snippet') or r.get('title'),
                'URL': r.get('url'),
                'Kaynak': r.get('source') or 'Açık Kaynak',
                'Domain': domain(r.get('url')),
                'Kategori': 'Sanayi & Teknoloji',
                'Risk_Skoru': 15,
                'Risk_Durumu': 'Normal',
                'Duygu': 'Nötr',
                'Doğrulama': 'Açık Kaynak'
            })
    return pd.DataFrame(cleaned)

# ============================================================
# STREAMLIT KULLANICI ARAYÜZÜ (UI)
# ============================================================
st.title("🛡️ Sanayi ve Teknoloji OSINT Radarı")

st.sidebar.header("⚙️ Tarama Seçenekleri")
period = st.sidebar.select_slider("Tarama Periyodu (Saat)", options=[3, 6, 12, 24, 48, 72], value=24)

if st.sidebar.button("🚀 Taramayı Başlat", use_container_width=True):
    with st.spinner("Açık kaynaklar taranıyor, lütfen bekleyin..."):
        df_results = run_osint_scan(period)
        st.session_state['scan_data'] = df_results
        st.success(f"Tarama tamamlandı! {len(df_results)} haber bulundu.")

df = st.session_state.get('scan_data', pd.DataFrame())

tab1, tab2, tab3, tab4 = st.tabs(["📊 Radar & Akış", "📌 Önemli Gelişmeler (ÖGN)", "📁 AKT Sepeti", "📝 Bilgi Notu Üret"])

with tab1:
    st.subheader("Güncel Haber Akışı")
    if not df.empty:
        st.dataframe(df[['Tarih', 'Kaynak', 'Başlık', 'Kategori', 'URL']], use_container_width=True)
    else:
        st.info("Henüz tarama yapılmadı. Sol menüden 'Taramayı Başlat' butonuna basın.")

with tab2:
    st.subheader("📌 Önemli Gelişmeler Notu (ÖGN) Hazırlama")
    st.caption("Seçilen haberler resmî Temsilcilik ÖGN formatına uygun Word belgesine dönüştürülür.")
    if not df.empty:
        selected_ogn = st.multiselect("ÖGN Raporuna Eklenecek Haberleri Seçin:", options=df['Başlık'].tolist(), key="ogn_select")
        if st.button("📄 ÖGN Word Belgesi Oluştur"):
            sub_df = df[df['Başlık'].isin(selected_ogn)]
            docx_bytes = make_important_basket_docx(sub_df)
            st.download_button(
                "⬇️ ÖGN Belgesini İndir",
                data=docx_bytes,
                file_name=f"OGN_Raporu_{date.today()}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

with tab3:
    st.subheader("📁 Açık Kaynak Tarama Raporu (AKT) Hazırlama")
    st.caption("Seçilen haberler resmî AKT şablonunda görsel altlıkları ve bağlantıları ile üretilir.")
    if not df.empty:
        selected_akt = st.multiselect("AKT Raporuna Eklenecek Haberleri Seçin:", options=df['Başlık'].tolist(), key="akt_select")
        if st.button("📄 AKT Word Belgesi Oluştur"):
            sub_df = df[df['Başlık'].isin(selected_akt)]
            docx_bytes = make_docx(sub_df.to_dict('records'))
            st.download_button(
                "⬇️ AKT Belgesini İndir",
                data=docx_bytes,
                file_name=f"AKT_Raporu_{date.today()}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

with tab4:
    st.subheader("📝 Bilgi Notu Oluşturma")
    st.caption("Seçilen konulardan 3 aşamalı analitik bilgi notu oluşturur.")
    if not df.empty:
        selected_note = st.multiselect("Bilgi Notuna Esas Haberleri Seçin:", options=df['Başlık'].tolist(), key="note_select")
        if st.button("📄 Bilgi Notu Word Belgesi Oluştur"):
            sub_df = df[df['Başlık'].isin(selected_note)]
            docx_bytes = make_analyst_docx(sub_df)
            st.download_button(
                "⬇️ Bilgi Notunu İndir",
                data=docx_bytes,
                file_name=f"Bilgi_Notu_{date.today()}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )