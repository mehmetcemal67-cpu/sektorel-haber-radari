import streamlit as st
import pandas as pd
import requests
import re
import html
import concurrent.futures
import xml.etree.ElementTree as ET
from datetime import datetime, date, timedelta, timezone
from io import BytesIO
from urllib.parse import quote_plus, urlparse
from bs4 import BeautifulSoup
from PIL import Image
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

# ============================================================
# SANAYİ & TEKNOLOJİ AÇIK KAYNAK / NEGATİF HABER RADARI
# ============================================================
# Kaynak mimarisi:
#   1) DDGS: Türkçe ve genel haber araması
#   2) GDELT DOC 2.0: küresel açık haber evreni + yüksek hacim
#   3) Doğrudan haber sayfasından metin/görsel çekme
#
# Not:
# "Tüm internet haberleri" teknik olarak garanti edilemez. Bu uygulama
# birden fazla açık kaynaktan mümkün olan en geniş akışı toplar.
# Paywall/robots/anti-bot olan sayfalarda tam metin alınamayabilir.
# ============================================================

st.set_page_config(
    page_title="Sanayi & Teknoloji OSINT Radarı",
    page_icon="🛡️",
    layout="wide",
)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/151 Safari/537.36"
    )
}

# -------------------------
# GENİŞ KAPSAMLI TERİMLER
# -------------------------
BROAD_TERM_BANK = [
    "sanayi", "sanayi üretimi", "imalat sanayi", "üretim", "fabrika",
    "organize sanayi bölgesi", "OSB", "yatırım", "yatırım teşvik",
    "teknoloji", "milli teknoloji", "Ar-Ge", "inovasyon", "patent",
    "dijital dönüşüm", "endüstri 4.0", "otomasyon", "robotik",
    "yapay zeka", "makine öğrenmesi", "siber güvenlik", "siber saldırı",
    "yarı iletken", "çip", "mikroçip", "işlemci", "elektronik",
    "batarya", "elektrikli araç", "şarj", "otomotiv", "TOGG",
    "havacılık", "uzay", "uydu", "roket", "İHA", "SİHA",
    "savunma sanayii", "ASELSAN", "TUSAŞ", "ROKETSAN", "HAVELSAN",
    "Baykar", "KAAN", "Çelik Kubbe", "HİSAR", "SİPER",
    "telekomünikasyon", "5G", "6G", "veri merkezi", "bulut",
    "kuantum", "biyoteknoloji", "medikal cihaz", "sağlık teknolojisi",
    "nanoteknoloji", "malzeme", "kompozit", "3D yazıcı",
    "yenilenebilir enerji", "güneş", "rüzgar", "hidrojen", "enerji",
    "petrokimya", "kimya", "demir çelik", "metal", "maden",
    "tekstil", "gıda teknolojisi", "tarım teknolojisi",
    "lojistik", "tersane", "gemi inşa", "denizcilik",
    "ihracat", "ithalat", "dış ticaret", "tedarik zinciri",
    "yerlileştirme", "millileştirme", "teknoloji transferi", "lisans",
    "girişim", "start-up", "teknopark", "girişim sermayesi",
    "TÜBİTAK", "KOSGEB", "Sanayi ve Teknoloji Bakanlığı", "TSE",
    "TürkPatent", "TEKNOFEST", "Türkiye Uzay Ajansı",
]

# Negatif olay sözlüğü: yüksek yakalama için bilinçli olarak geniş tutuldu.
NEGATIVE_KEYWORDS = [
    "kriz", "skandal", "fiyasko", "iflas", "konkordato", "zarar",
    "borç", "nakit sıkıntısı", "üretim durdu", "üretim durduruldu",
    "fabrika kapandı", "fabrika kapanıyor", "işten çıkarma",
    "işçi çıkarma", "grev", "lokavt", "protesto", "eylem",
    "soruşturma", "dava", "ceza", "para cezası", "denetim",
    "geri çağırma", "recall", "arıza", "kaza", "patlama", "yangın",
    "siber saldırı", "veri sızıntısı", "hacklendi", "fidye yazılımı",
    "ambargo", "yaptırım", "ihracat yasağı", "ithalat yasağı",
    "lisans reddi", "ruhsat iptali", "izin verilmedi",
    "sözleşme feshi", "ihale iptal", "iptal edildi", "askıya alındı",
    "ertelendi", "gecikme", "teslim edilemedi", "testi geçemedi",
    "kapasite kaybı", "daralma", "sert düşüş", "pazar kaybı",
    "maliyet artışı", "fiyat artışı", "enflasyon", "tedarik sorunu",
    "tedarik krizi", "çip krizi", "kıtlık", "blokaj", "sıkıntı",
    "bağımlılık", "dışa bağımlı", "teknoloji açığı", "risk",
    "tehdit", "güvenlik açığı", "zafiyet", "ifşa", "usulsüzlük",
    "yolsuzluk", "vurgun", "israf", "gizlendi", "sümen altı",
    "şüphe", "iddia", "tartışma", "tepki", "eleştiri", "kriminal",
]

HIGH_RISK_KEYWORDS = [
    "iflas", "konkordato", "üretim durdu", "fabrika kapandı",
    "fabrika kapanıyor", "siber saldırı", "veri sızıntısı",
    "fidye yazılımı", "ambargo", "yaptırım", "ihracat yasağı",
    "lisans reddi", "ruhsat iptali", "sözleşme feshi",
    "ihale iptal", "patlama", "yangın", "ölüm", "can kaybı",
    "soruşturma", "yolsuzluk", "usulsüzlük", "vurgun",
    "casusluk", "güvenlik açığı", "kritik zafiyet", "stratejik risk",
    "tedarik krizi", "çip krizi", "dışa bağımlı",
]

POSITIVE_KEYWORDS = [
    "rekor", "artış", "büyüme", "ihracat arttı", "yatırım",
    "yatırım anlaşması", "seri üretim", "üretime başladı",
    "fabrika açıldı", "kapasite artışı", "yerlileştirme",
    "millileştirme", "patent", "teknoloji transferi", "başarı",
    "ödül", "yeni tesis", "yeni fabrika", "istihdam artışı",
]

CATEGORIES = {
    "Savunma & Havacılık": [
        "aselsan", "tusaş", "roketsan", "havelsan", "baykar", "iha",
        "siha", "kaan", "siper", "hisar", "çelik kubbe", "savunma",
        "havacılık", "roket", "füze", "insansız",
    ],
    "Otomotiv & Mobilite": [
        "togg", "otomotiv", "elektrikli araç", "elektrikli otomobil",
        "batarya", "şarj", "byd", "araç",
    ],
    "Yarı İletken & Elektronik": [
        "çip", "mikroçip", "yarı iletken", "işlemci", "pcb",
        "elektronik", "transistör", "wafer",
    ],
    "Enerji & İklim Teknolojileri": [
        "yenilenebilir", "güneş", "rüzgar", "hidrojen", "batarya",
        "enerji depolama", "nükleer", "enerji",
    ],
    "Dijital & Yapay Zeka": [
        "yapay zeka", "makine öğrenmesi", "siber", "5g", "6g",
        "veri merkezi", "bulut", "kuantum", "yazılım",
    ],
    "Sanayi & Üretim": [
        "fabrika", "imalat", "üretim", "osb", "sanayi", "makine",
        "robotik", "otomasyon", "demir çelik", "kimya", "petrokimya",
    ],
    "Uzay & İleri Teknoloji": [
        "uydu", "uzay", "roket", "tua", "türkiye uzay ajansı",
        "nanoteknoloji", "biyoteknoloji", "kuantum",
    ],
    "Kurumsal Ekosistem": [
        "tübitak", "kosgeb", "sanayi ve teknoloji bakanlığı", "tse",
        "türkpatent", "teknopark", "teknofest", "yatırım teşvik",
    ],
}

SOURCE_TRUST_HINTS = {
    "reuters": 0.95, "ap": 0.95, "aa.com.tr": 0.85, "aa": 0.85,
    "trthaber": 0.80, "bbc": 0.90, "dw": 0.85, "bloomberg": 0.90,
    "ft.com": 0.90, "nytimes": 0.90, "wsj": 0.90,
}

# -------------------------
# YARDIMCILAR
# -------------------------
def normalize_text(text):
    text = (text or "").lower()
    text = re.sub(r"\s+", " ", text)
    return text.strip()

def normalize_title(title):
    t = normalize_text(title)
    t = re.sub(r"[^\w\s]", " ", t)
    return re.sub(r"\s+", " ", t).strip()

def parse_dt(value):
    if not value:
        return None
    value = str(value).strip()
    candidates = [
        value.replace("Z", "+00:00"),
        value.replace(" GMT", " +00:00"),
    ]
    for item in candidates:
        try:
            dt = datetime.fromisoformat(item)
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            return dt.astimezone(timezone.utc)
        except Exception:
            pass
    # RSS için RFC-822
    try:
        from email.utils import parsedate_to_datetime
        dt = parsedate_to_datetime(value)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(timezone.utc)
    except Exception:
        return None

def fmt_dt(dt):
    if not dt:
        return "Tarih Belirsiz"
    return dt.astimezone().strftime("%d.%m.%Y %H:%M:%S")

def domain_of(url):
    try:
        return urlparse(url).netloc.lower().replace("www.", "")
    except Exception:
        return ""

def source_score(url):
    d = domain_of(url)
    for k, v in SOURCE_TRUST_HINTS.items():
        if k in d:
            return v
    return 0.60

def detect_category(text):
    text = normalize_text(text)
    for cat, kws in CATEGORIES.items():
        if any(k in text for k in kws):
            return cat
    return "Genel Sanayi / Teknoloji"

# -------------------------
# DUYGU + RİSK
# -------------------------
def analyze_article(title, body):
    text = normalize_text(f"{title} {body}")

    neg_hits = [k for k in NEGATIVE_KEYWORDS if k in text]
    risk_hits = [k for k in HIGH_RISK_KEYWORDS if k in text]
    pos_hits = [k for k in POSITIVE_KEYWORDS if k in text]

    # Ağırlıklı skor: negatif olay kelimeleri daha güçlü.
    score = 0.0
    score += len(pos_hits) * 0.18
    score -= len(neg_hits) * 0.22
    score -= len(risk_hits) * 0.30
    score = max(-1.0, min(1.0, score))

    if score <= -0.20:
        sentiment = "Negatif"
    elif score >= 0.20:
        sentiment = "Pozitif"
    else:
        sentiment = "Nötr"

    # Yüksek risk: yalnızca "manipülasyon kelimesi" değil,
    # gerçek olay/risk göstergeleri + negatif yoğunluğu.
    if len(risk_hits) >= 1 or len(neg_hits) >= 3 or score <= -0.55:
        risk = "Yüksek Risk"
    elif len(neg_hits) >= 1 or score < -0.20:
        risk = "İzleme"
    else:
        risk = "Normal"

    return {
        "sentiment": sentiment,
        "score": round(score, 2),
        "risk": risk,
        "negative_hits": neg_hits,
        "risk_hits": risk_hits,
        "positive_hits": pos_hits,
        "category": detect_category(text),
    }

# -------------------------
# SORGU GENİŞLETME
# -------------------------
def split_query(query):
    raw = re.split(r"\bOR\b|,|\n", query or "", flags=re.I)
    out, seen = [], set()
    for x in raw:
        x = x.strip().strip('"').strip("'")
        if len(x) >= 2 and x.lower() not in seen:
            out.append(x)
            seen.add(x.lower())
    return out or ["sanayi", "teknoloji"]

def build_queries(user_query, broad=True, negative_boost=True, cap=100):
    base = split_query(user_query)
    if broad:
        for term in BROAD_TERM_BANK:
            if len(base) >= cap:
                break
            if term.lower() not in {x.lower() for x in base}:
                base.append(term)

    queries = list(base)
    if negative_boost:
        # Negatif arama evrenini ayrı tutuyoruz; böylece normal haberler
        # negatif haberlerin içinde kaybolmuyor.
        for term in base[:70]:
            for neg in NEGATIVE_KEYWORDS[:18]:
                if len(queries) >= cap:
                    break
                queries.append(f'"{term}" {neg}')
            if len(queries) >= cap:
                break
    return queries[:cap]

# -------------------------
# DDGS
# -------------------------
def ddgs_news(term, region="tr-tr", max_results=20):
    try:
        from ddgs import DDGS
    except ImportError:
        try:
            from duckduckgo_search import DDGS
        except ImportError:
            return []

    try:
        with DDGS() as ddgs:
            return list(ddgs.news(term, region=region, timelimit="d", max_results=max_results))
    except Exception:
        return []

# -------------------------
# GDELT DOC 2.0
# -------------------------
def gdelt_news(query, max_records=50):
    """
    GDELT DOC 2.0 ArticleList JSON.
    GDELT küresel haber kapsaması sağlar; DDGS Türkçe kapsamasını destekler.
    """
    url = "https://api.gdeltproject.org/api/v2/doc/doc"
    params = {
        "query": query,
        "mode": "artlist",
        "maxrecords": min(max_records, 250),
        "format": "json",
        "sort": "HybridRel",
        "timespan": "24h",
    }
    try:
        r = requests.get(url, params=params, headers=HEADERS, timeout=15)
        r.raise_for_status()
        data = r.json()
        return data.get("articles", []) or []
    except Exception:
        return []

# -------------------------
# HABER METNİ / GÖRSEL
# -------------------------
def extract_article(url, timeout=8):
    result = {"text": "", "image_url": ""}
    if not url:
        return result

    try:
        r = requests.get(url, headers=HEADERS, timeout=timeout)
        if r.status_code != 200 or not r.text:
            return result

        soup = BeautifulSoup(r.text, "html.parser")

        # OG image
        og = soup.find("meta", attrs={"property": "og:image"})
        if og and og.get("content"):
            result["image_url"] = og["content"]

        for tag in soup(["script", "style", "nav", "footer", "header", "aside", "form", "iframe", "noscript"]):
            tag.decompose()

        paragraphs = [
            p.get_text(" ", strip=True)
            for p in soup.find_all("p")
        ]
        paragraphs = [p for p in paragraphs if len(p) >= 45]

        # Tekrarlı boilerplate temizliği
        seen = set()
        clean = []
        for p in paragraphs:
            key = normalize_text(p)
            if key not in seen:
                seen.add(key)
                clean.append(p)

        result["text"] = " ".join(clean)[:12000]
        return result
    except Exception:
        return result

def extract_many(urls, workers=16):
    urls = list(dict.fromkeys([u for u in urls if u]))
    out = {}
    if not urls:
        return out
    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as ex:
        fmap = {ex.submit(extract_article, u): u for u in urls}
        for f in concurrent.futures.as_completed(fmap):
            u = fmap[f]
            try:
                out[u] = f.result()
            except Exception:
                out[u] = {"text": "", "image_url": ""}
    return out

# -------------------------
# ÖZET
# -------------------------
def summarize_text(title, body, max_chars=1800):
    body = (body or "").strip()
    if not body:
        return f"{title}. Haber metnine erişim sağlanamadığı için açık kaynak başlık/özet bilgisi esas alınmıştır."

    sentences = [
        s.strip()
        for s in re.split(r"(?<=[.!?])\s+", body)
        if len(s.strip()) >= 35
    ]

    # İlk cümleleri körlemesine almak yerine; sayı, risk ve olay cümlelerini öne çek.
    priority = []
    normal = []
    risk_terms = NEGATIVE_KEYWORDS + HIGH_RISK_KEYWORDS

    for s in sentences:
        low = normalize_text(s)
        if any(k in low for k in risk_terms) or re.search(r"\d", s):
            priority.append(s)
        else:
            normal.append(s)

    selected = []
    for s in priority + normal:
        if s not in selected:
            selected.append(s)
        if len(" ".join(selected)) >= max_chars:
            break

    text = " ".join(selected)
    return text[:max_chars].rstrip() + ("…" if len(text) > max_chars else "")

# -------------------------
# HABER TOPLAMA
# -------------------------
def fetch_news(user_query, time_hours=24, max_results=150,
               broad=True, negative_boost=True, include_global=True,
               workers=12):
    queries = build_queries(user_query, broad, negative_boost)

    # İstemci tarafı gerçek zaman filtresi.
    cutoff = datetime.now(timezone.utc) - timedelta(hours=time_hours)

    jobs = []
    for q in queries:
        jobs.append(("ddgs", q, "tr-tr"))
    if include_global:
        for q in queries[:35]:
            jobs.append(("ddgs", q, "wt-wt"))

    # GDELT için daha az ama geniş sorgular; çok sayıda istekle API'yi boğmamak için.
    gdelt_queries = [
        '(industry OR manufacturing OR technology OR semiconductor OR artificial intelligence)',
        '(defense OR aerospace OR automotive OR battery OR energy)',
        '(cybersecurity OR chip OR factory OR supply chain OR sanctions)',
        '(Turkey OR Türkiye) (industry OR technology OR manufacturing)',
    ]
    if negative_boost:
        gdelt_queries += [
            '(Turkey OR Türkiye) (factory OR manufacturing OR technology) (crisis OR shutdown OR sanction OR lawsuit OR cyberattack)',
            '(Turkey OR Türkiye) (defense OR automotive OR semiconductor) (risk OR delay OR cancellation)',
        ]

    records = []
    lock = None

    def ddgs_job(job):
        _, q, region = job
        return ddgs_news(q, region=region, max_results=20)

    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as ex:
        futures = [ex.submit(ddgs_job, j) for j in jobs]
        for f in concurrent.futures.as_completed(futures):
            try:
                records.extend(f.result())
            except Exception:
                pass

    if include_global:
        with concurrent.futures.ThreadPoolExecutor(max_workers=6) as ex:
            futures = [ex.submit(gdelt_news, q, 50) for q in gdelt_queries]
            for f in concurrent.futures.as_completed(futures):
                try:
                    records.extend(f.result())
                except Exception:
                    pass

    normalized = []
    seen_urls = set()
    seen_titles = set()

    for r in records:
        url = r.get("url") or r.get("url_mobile") or ""
        title = r.get("title") or ""
        if not url or not title:
            continue

        dt = parse_dt(r.get("date") or r.get("seendate") or r.get("publishedAt"))
        if dt and dt < cutoff:
            continue

        nt = normalize_title(title)
        # Başlığın ilk 10 kelimesiyle kaba duplicate temizliği.
        title_key = " ".join(nt.split()[:10])
        if url in seen_urls or title_key in seen_titles:
            continue

        source = (
            r.get("source")
            or r.get("domain")
            or domain_of(url)
            or "Açık Kaynak"
        )
        snippet = r.get("body") or r.get("snippet") or ""
        image = r.get("image") or r.get("socialimage") or ""

        normalized.append({
            "title": html.unescape(title),
            "url": url,
            "published_dt": dt,
            "published": fmt_dt(dt),
            "source": str(source),
            "domain": domain_of(url),
            "snippet": html.unescape(snippet),
            "image_url": image,
        })
        seen_urls.add(url)
        seen_titles.add(title_key)

    # En yeni önce.
    normalized.sort(
        key=lambda x: x["published_dt"] or datetime.min.replace(tzinfo=timezone.utc),
        reverse=True
    )

    return normalized[:max_results]

# -------------------------
# DOCX
# -------------------------
def add_hyperlink(paragraph, url, text):
    try:
        part = paragraph.part
        rid = part.relate_to(
            url,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            is_external=True,
        )
        safe_text = html.escape(text)
        xml = (
            f'<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" r:id="{rid}">'
            f'<w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:color w:val="0000FF"/>'
            f'<w:u w:val="single"/></w:rPr><w:t>{safe_text}</w:t></w:r></w:hyperlink>'
        )
        paragraph._p.append(parse_xml(xml))
    except Exception:
        paragraph.add_run(f"{text} ({url})")

def image_bytes(url):
    if not url:
        return None
    try:
        r = requests.get(url, headers=HEADERS, timeout=7)
        if r.status_code != 200 or len(r.content) < 1500:
            return None
        ct = r.headers.get("Content-Type", "")
        if not ct.startswith("image"):
            return None
        img = Image.open(BytesIO(r.content))
        if img.mode in ("RGBA", "P", "LA"):
            img = img.convert("RGB")
        out = BytesIO()
        img.save(out, format="JPEG", quality=88)
        out.seek(0)
        return out
    except Exception:
        return None

def style_cell(cell, bold=False, size=8.5):
    for p in cell.paragraphs:
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        for run in p.runs:
            run.font.name = "Arial"
            run.font.size = Pt(size)
            run.font.bold = bold

def make_docx(df, query, title="SANAYİ VE TEKNOLOJİ AÇIK KAYNAK TARAMA RAPORU"):
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(0.5)
        s.bottom_margin = Inches(0.5)
        s.left_margin = Inches(0.55)
        s.right_margin = Inches(0.55)

    h = doc.add_heading(title, 0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in h.runs:
        r.font.name = "Arial"
        r.font.size = Pt(14)
        r.font.bold = True
        r.font.color.rgb = RGBColor(27, 54, 93)

    p = doc.add_paragraph()
    p.add_run("Sorgu: ").bold = True
    p.add_run(query)
    p.add_run("\nRapor zamanı: ").bold = True
    p.add_run(datetime.now().astimezone().strftime("%d.%m.%Y %H:%M:%S"))
    p.add_run("\nHaber sayısı: ").bold = True
    p.add_run(str(len(df)))

    if df.empty:
        doc.add_paragraph("Kriterlere uygun haber bulunamadı.")
        buf = BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf

    # Kritik haberler önce.
    risk = df[df["Risk_Durumu"] == "Yüksek Risk"]
    negative = df[(df["Duygu"] == "Negatif") & (df["Risk_Durumu"] != "Yüksek Risk")]
    normal = df[~df.index.isin(risk.index) & ~df.index.isin(negative.index)]

    ordered = pd.concat([risk, negative, normal])

    for idx, (_, row) in enumerate(ordered.iterrows(), 1):
        ptitle = doc.add_paragraph()
        rr = ptitle.add_run(f"{idx}. {row['Başlık']}")
        rr.bold = True
        rr.font.name = "Arial"
        rr.font.size = Pt(11)

        meta = doc.add_paragraph(
            f"{row['Tarih']} | {row['Kaynak']} | {row['Kategori']} | "
            f"Duygu: {row['Duygu']} ({row['Skor']}) | Risk: {row['Risk_Durumu']}"
        )
        for r in meta.runs:
            r.font.name = "Arial"
            r.font.size = Pt(8.5)
            r.italic = True

        img = image_bytes(row.get("Görsel_URL", ""))
        if img:
            try:
                doc.add_paragraph().add_run().add_picture(img, width=Inches(3.0))
            except Exception:
                pass

        ps = doc.add_paragraph()
        ps.add_run("ÖZET: ").bold = True
        ps.add_run(row["Özet"])

        if row.get("Risk_Sinyalleri"):
            pr = doc.add_paragraph()
            pr.add_run("RİSK SİNYALLERİ: ").bold = True
            pr.add_run(", ".join(row["Risk_Sinyalleri"]))

        if row.get("Negatif_Sinyaller"):
            pn = doc.add_paragraph()
            pn.add_run("NEGATİF SİNYALLER: ").bold = True
            pn.add_run(", ".join(row["Negatif_Sinyaller"]))

        pl = doc.add_paragraph()
        pl.add_run("HABER LİNKİ: ").bold = True
        add_hyperlink(pl, row["URL"], row["URL"])

        doc.add_paragraph("─" * 85)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# -------------------------
# STREAMLIT
# -------------------------
st.title("🛡️ Sanayi & Teknoloji Açık Kaynak / Negatif Haber Radarı")
st.caption(
    "Anlık haber akışı · negatif/riskli ayrıştırma · kronolojik sıralama · "
    "tam metin zenginleştirme · DOCX bilgi notu · açık kaynak çalışma ekranı"
)

with st.sidebar:
    st.header("⚙️ Tarama Ayarları")

    default_query = (
        "sanayi OR teknoloji OR üretim OR fabrika OR yatırım OR "
        "savunma sanayii OR havacılık OR otomotiv OR TOGG OR "
        "ASELSAN OR TUSAŞ OR ROKETSAN OR HAVELSAN OR Baykar OR "
        "çip OR yarı iletken OR yapay zeka OR siber güvenlik OR "
        "enerji OR batarya OR TÜBİTAK OR KOSGEB OR OSB"
    )

    query = st.text_area(
        "Esnek sorgu (OR / virgül / satır sonu):",
        value=default_query,
        height=150,
    )

    broad = st.checkbox(
        "Kapsamlı Sanayi & Teknoloji Evrenini Otomatik Genişlet",
        value=True,
    )
    negative_boost = st.checkbox(
        "Negatif / yüksek risk aramalarını güçlendir",
        value=True,
    )
    include_global = st.checkbox(
        "Küresel / yabancı basını da tara",
        value=True,
    )
    enrich = st.checkbox(
        "Haber sayfasından tam metin + görsel çek",
        value=True,
    )

    hours = st.selectbox(
        "Zaman penceresi",
        [6, 12, 24, 48, 72],
        index=2,
        format_func=lambda x: f"Son {x} saat",
    )

    max_news = st.slider("Maksimum haber", 30, 250, 120, 10)

    auto_refresh = st.checkbox("Otomatik yenileme", value=False)
    refresh_min = st.selectbox("Yenileme aralığı", [2, 5, 10, 15, 30], index=1)

    run = st.button("🔍 TARAMAYI BAŞLAT / YENİLE", type="primary", use_container_width=True)

# Otomatik yenileme: harici paket zorunluluğu olmadan Streamlit fragment varsa kullan.
if auto_refresh:
    try:
        st.caption(f"🔄 Otomatik yenileme: {refresh_min} dakika")
        # Streamlit'in güncel sürümlerinde fragment run_every desteklenir.
        # Uygulamanın tamamını burada zorla yeniden çalıştırmak yerine kullanıcı
        # butonunu da bırakıyoruz.
        st.info("Otomatik yenileme için Streamlit sürümünüz fragment(run_every=...) desteklemelidir.")
    except Exception:
        pass

if "df" not in st.session_state:
    st.session_state.df = None
if "last_query" not in st.session_state:
    st.session_state.last_query = ""

if run:
    with st.spinner("Çok kaynaklı haber evreni taranıyor..."):
        raw = fetch_news(
            query,
            time_hours=hours,
            max_results=max_news,
            broad=broad,
            negative_boost=negative_boost,
            include_global=include_global,
        )

    if raw:
        if enrich:
            urls = [x["url"] for x in raw]
            with st.spinner(f"{len(urls)} haber sayfası zenginleştiriliyor..."):
                enriched = extract_many(urls)
        else:
            enriched = {}

        rows = []
        for item in raw:
            extra = enriched.get(item["url"], {})
            body = extra.get("text") or item["snippet"]
            image = item["image_url"] or extra.get("image_url", "")

            a = analyze_article(item["title"], body)
            rows.append({
                "Tarih_dt": item["published_dt"],
                "Tarih": item["published"],
                "Kaynak": item["source"],
                "Domain": item["domain"],
                "Kaynak_Güven_Hint": source_score(item["url"]),
                "Kategori": a["category"],
                "Başlık": item["title"],
                "Özet": summarize_text(item["title"], body),
                "Duygu": a["sentiment"],
                "Skor": a["score"],
                "Risk_Durumu": a["risk"],
                "Negatif_Sinyaller": a["negative_hits"],
                "Risk_Sinyalleri": a["risk_hits"],
                "Pozitif_Sinyaller": a["positive_hits"],
                "URL": item["url"],
                "Görsel_URL": image,
                "Tam_Metin_Uzunluğu": len(body or ""),
                "Seç": a["risk"] in ("Yüksek Risk", "İzleme") or a["sentiment"] == "Negatif",
            })

        df = pd.DataFrame(rows)
        if not df.empty:
            df = df.sort_values(
                "Tarih_dt",
                ascending=False,
                na_position="last"
            ).reset_index(drop=True)

        st.session_state.df = df
        st.session_state.last_query = query

df = st.session_state.df

if df is None:
    st.info("Soldaki ayarlardan sorguyu belirleyip **TARAMAYI BAŞLAT / YENİLE** düğmesine basın.")
elif df.empty:
    st.warning("Seçilen zaman penceresinde haber bulunamadı.")
else:
    total = len(df)
    neg = int((df["Duygu"] == "Negatif").sum())
    high = int((df["Risk_Durumu"] == "Yüksek Risk").sum())
    watch = int((df["Risk_Durumu"] == "İzleme").sum())

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Toplam Haber", total)
    k2.metric("Negatif", neg)
    k3.metric("Yüksek Risk", high)
    k4.metric("İzleme", watch)

    st.markdown("---")

    tab_all, tab_neg, tab_risk, tab_osint = st.tabs([
        f"📰 Kronolojik Akış ({total})",
        f"⚠️ Negatif ({neg})",
        f"🚨 Yüksek Risk ({high})",
        "🔎 Açık Kaynak Çalışma Masası",
    ])

    display_cols = [
        "Seç", "Tarih", "Kaynak", "Kategori", "Başlık",
        "Duygu", "Skor", "Risk_Durumu", "URL"
    ]

    with tab_all:
        edited = st.data_editor(
            df[display_cols],
            column_config={
                "Seç": st.column_config.CheckboxColumn("Bilgi Notuna Ekle"),
                "URL": st.column_config.LinkColumn("Haber Linki"),
            },
            disabled=[c for c in display_cols if c != "Seç"],
            hide_index=True,
            use_container_width=True,
            height=620,
            key="news_editor",
        )
        df.loc[edited.index, "Seç"] = edited["Seç"]

    with tab_neg:
        neg_df = df[df["Duygu"] == "Negatif"]
        st.dataframe(
            neg_df[["Tarih", "Kaynak", "Kategori", "Başlık", "Özet", "Skor", "URL"]],
            column_config={"URL": st.column_config.LinkColumn("Haber Linki")},
            hide_index=True,
            use_container_width=True,
            height=620,
        )

    with tab_risk:
        risk_df = df[df["Risk_Durumu"] == "Yüksek Risk"]
        st.dataframe(
            risk_df[
                ["Tarih", "Kaynak", "Kategori", "Başlık",
                 "Özet", "Risk_Sinyalleri", "URL"]
            ],
            column_config={"URL": st.column_config.LinkColumn("Haber Linki")},
            hide_index=True,
            use_container_width=True,
            height=620,
        )

    with tab_osint:
        st.subheader("🔎 Seçili Haber Üzerinden Açık Kaynak Çalışması")
        selected = df[df["Seç"] == True]

        if selected.empty:
            st.info("Çalışma yapmak için önce haber listesinden bir veya daha fazla haber seçin.")
        else:
            choice = st.selectbox(
                "İncelenecek haber:",
                list(selected.index),
                format_func=lambda i: selected.loc[i, "Başlık"],
            )
            row = selected.loc[choice]

            c1, c2 = st.columns([2, 1])
            with c1:
                st.markdown(f"### {row['Başlık']}")
                st.write(row["Özet"])
                st.markdown(f"**Kaynak:** {row['Kaynak']}  \n**Tarih:** {row['Tarih']}")
                st.markdown(f"**Kategori:** {row['Kategori']}")
                st.markdown(f"**Duygu:** {row['Duygu']} / **Risk:** {row['Risk_Durumu']}")
                st.link_button("🌐 Haberi Aç", row["URL"], use_container_width=True)

            with c2:
                st.metric("Risk Skoru", row["Skor"])
                st.metric("Kaynak Güven Hinti", f"{row['Kaynak_Güven_Hint']:.2f}")
                st.metric("Metin Uzunluğu", row["Tam_Metin_Uzunluğu"])

            st.markdown("#### Risk / negatif sinyalleri")
            st.write(", ".join(row["Risk_Sinyalleri"]) or "Yok")
            st.write(", ".join(row["Negatif_Sinyaller"]) or "Yok")

            st.markdown("#### Analist notu")
            note = st.text_area(
                "Bu haber için açık kaynak değerlendirmesi, teyit ihtiyacı, kurum/şirket isimleri, bağlantılı olaylar vb.",
                height=180,
                key=f"note_{choice}",
            )

            st.markdown("#### Hızlı kontrol listesi")
            st.checkbox("İkinci bağımsız kaynakla teyit edildi", key=f"c1_{choice}")
            st.checkbox("Birincil kaynak / resmi açıklama kontrol edildi", key=f"c2_{choice}")
            st.checkbox("Tarih ve olay zamanı doğrulandı", key=f"c3_{choice}")
            st.checkbox("Görselin kaynak/bağlamı kontrol edildi", key=f"c4_{choice}")

    st.markdown("---")
    st.subheader("📥 Çıktılar")

    selected_df = df[df["Seç"] == True]
    st.write(f"**{len(selected_df)} haber** bilgi notuna seçildi.")

    full_doc = make_docx(
        df,
        st.session_state.last_query,
        title="SANAYİ VE TEKNOLOJİ AÇIK KAYNAK TARAMA RAPORU"
    )
    st.download_button(
        "📄 TAM TARAMA RAPORUNU İNDİR (.DOCX)",
        data=full_doc,
        file_name=f"Sanayi_Teknoloji_OSINT_{date.today()}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
        use_container_width=True,
    )

    if not selected_df.empty:
        briefing_doc = make_docx(
            selected_df,
            st.session_state.last_query,
            title="GÜNLÜK BİLGİ NOTU — SANAYİ & TEKNOLOJİ"
        )
        st.download_button(
            "📝 SEÇİLİ HABERLERDEN BİLGİ NOTU İNDİR (.DOCX)",
            data=briefing_doc,
            file_name=f"Gunluk_Bilgi_Notu_{date.today()}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

    # CSV de pratik bir OSINT çalışma çıktısıdır.
    csv = df.drop(columns=["Tarih_dt"], errors="ignore").to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "📊 HAM OSINT VERİSİNİ İNDİR (.CSV)",
        data=csv,
        file_name=f"Sanayi_Teknoloji_OSINT_{date.today()}.csv",
        mime="text/csv",
        use_container_width=True,
    )

    st.caption(
        "Not: Negatif/risk sınıflandırması otomatik ön elemedir; 'yüksek risk' "
        "tek başına olayın doğrulandığı anlamına gelmez. Kritik bulgular bağımsız "
        "ve birincil kaynaklarla teyit edilmelidir."
    )