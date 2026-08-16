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
# SANAYİ & TEKNOLOJİ AÇIK KAYNAK / NEGATİF HABER RADARI — V7
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

# ============================================================
# KAYNAK ÖNCELİKLERİ
# Kullanıcının ihtiyacı: Türk basını / Türk teknoloji-savunma medyası
# birinci sırada; Yunanistan medyası yalnızca Türk savunma sanayii
# bağlamında özel olarak tutulur; genel yabancı basın varsayılan olarak
# kapalıdır.
# ============================================================
TR_MAIN_DOMAINS = [
    "aa.com.tr", "trthaber.com", "ntv.com.tr", "cnnturk.com",
    "haberturk.com", "hurriyet.com.tr", "milliyet.com.tr", "sabah.com.tr",
    "sozcu.com.tr", "cumhuriyet.com.tr", "karar.com", "yenisafak.com",
    "star.com.tr", "aksam.com.tr", "turkiyegazetesi.com.tr",
    "t24.com.tr", "haber7.com", "haberler.com", "ensonhaber.com",
    "gazeteduvar.com.tr", "odatv.com", "medyascope.tv", "tv100.com",
    "tgrthaber.com.tr", "mynet.com", "indyturk.com", "diken.com.tr",
    "dunya.com", "ekonomim.com", "bloomberght.com", "paraanaliz.com",
    "bigpara.com", "fortuneturkey.com", "doviz.com",
]

TR_TECH_DEFENSE_DOMAINS = [
    "webrazzi.com", "shiftdelete.net", "donanimhaber.com", "chip.com.tr",
    "log.com.tr", "technopat.net", "hardwareplus.com.tr", "turk-internet.com",
    "savunmasanayist.com", "savunmatr.com", "defenceturk.net", "defencehere.com",
    "c4defence.com", "savunmahaber.com", "savunmasanayi.org", "defensemirror.com",
    "gdh.digital", "stratejikortak.com", "m5dergi.com", "mavivatan.net",
]

TR_OFFICIAL_DOMAINS = [
    "sanayi.gov.tr", "tubitak.gov.tr", "kosgeb.gov.tr", "tse.org.tr",
    "turkpatent.gov.tr", "tua.gov.tr", "ticaret.gov.tr", "uab.gov.tr",
    "aselsan.com", "tusas.com", "roketsan.com.tr", "havelsan.com.tr",
    "baykartech.com", "togg.com.tr", "tei.com.tr", "tai.com.tr",
]

GR_DOMAINS = [
    "kathimerini.gr", "protothema.gr", "news247.gr", "tovima.gr",
    "enikos.gr", "naftemporiki.gr", "skai.gr", "capital.gr",
    "defence-point.gr", "defencereview.gr", "militaire.gr", "pronews.gr",
    "newsbreak.gr", "pentapostagma.gr", "hellasjournal.com",
]

TR_DEFENSE_TERMS = [
    "türkiye", "turkey", "türk", "turkish", "türk savunma", "savunma sanayii",
    "savunma sanayi", "aselsan", "tusas", "tusaş", "roketsan", "havelsan",
    "baykar", "bayraktar", "kaan", "kızılelma", "siper", "hisar", "iha", "siha",
    "insansız savaş uçağı", "milli muharip uçak", "altay", "milden", "milli denizaltı",
    "tf-2000", "milgem", "çelik kubbe", "turkish defence", "turkish defense",
    "defence industry", "defense industry", "turkey arms", "turkish arms",
    "τουρκία", "τουρκικό", "τουρκική", "τουρκικός", "τουρκικά",
    "τουρκική αμυντική βιομηχανία", "τουρκική άμυνα", "τουρκικά όπλα",
    "τουρκικά μαχητικά", "τουρκικά drones", "τουρκικά μη επανδρωμένα",
]

SOCIAL_SEARCH_DOMAINS = [
    "x.com", "twitter.com", "youtube.com", "linkedin.com", "facebook.com", "instagram.com"
]



# V5: Haber gerçekten sanayi/teknoloji bağlamında mı?
# "yatırım", "araç", "enerji", "eylem" gibi tek başına çok geniş
# kelimeler spor/gündem/hava durumu vb. sonuçları içeri sokabiliyor.
CORE_TOPIC_TERMS = [
    "sanayi", "imalat", "üretim", "fabrika", "osb", "organize sanayi",
    "teknoloji", "teknolojik", "ar-ge", "arge", "inovasyon", "patent",
    "dijital dönüşüm", "endüstri 4.0", "otomasyon", "robotik",
    "yapay zeka", "yapay zekâ", "makine öğrenmesi", "siber güvenlik",
    "siber saldırı", "veri merkezi", "bulut bilişim", "yazılım",
    "çip", "mikroçip", "yarı iletken", "işlemci", "elektronik",
    "telekom", "5g", "6g", "kuantum", "biyoteknoloji", "nanoteknoloji",
    "savunma sanay", "savunma sanayii", "savunma teknoloj", "askeri teknoloji",
    "aselsan", "tusaş", "tusas", "roketsan", "havelsan", "baykar",
    "bayraktar", "iha", "siha", "drone", "insansız", "kaan", "kızılelma",
    "siper", "hisar", "füze", "roket", "uydu", "uzay", "havacılık",
    "otomotiv", "elektrikli araç", "elektrikli otomobil", "batarya",
    "şarj istasyonu", "enerji depolama", "güneş enerjisi", "rüzgar enerjisi",
    "hidrojen", "nükleer enerji", "enerji teknoloj", "petrokimya", "kimya",
    "demir çelik", "çelik", "metalurji", "maden", "madencilik",
    "tekstil", "gıda teknoloj", "tarım teknoloj", "lojistik teknoloj",
    "gemi inşa", "tersane", "yerlileştirme", "millileştirme", "tedarik zinciri",
    "tübitak", "kosgeb", "sanayi ve teknoloji bakanlığı", "teknopark",
    "teknofest", "türk patent", "türkpatent", "yatırım teşvik",
]

# Negatif sorgularda konu bağını daha da güçlendiriyoruz.
NEGATIVE_EVENT_TERMS = [
    "iflas", "konkordato", "üretim durdu", "üretim durduruldu", "fabrika kapandı",
    "fabrika kapanıyor", "işten çıkarma", "işçi çıkarma", "grev", "lokavt",
    "soruşturma", "dava", "ceza", "geri çağırma", "siber saldırı",
    "veri sızıntısı", "fidye yazılımı", "ambargo", "yaptırım", "ihracat yasağı",
    "ithalat yasağı", "sözleşme feshi", "ihale iptal", "askıya alındı",
    "ertelendi", "gecikme", "tedarik krizi", "çip krizi", "daralma",
    "kapasite kaybı", "güvenlik açığı", "zafiyet", "usulsüzlük", "yolsuzluk",
]


def is_relevant_industry_tech(title, snippet="", user_terms=None):
    text = normalize_text(f"{title} {snippet}")
    # Haber, sanayi/teknoloji ekseninde en az bir güçlü çekirdek terim
    # veya kullanıcının açıkça verdiği bir terim içermeli.
    core_hits = [k for k in CORE_TOPIC_TERMS if k in text]
    user_hits = []
    for k in (user_terms or []):
        k = normalize_text(k)
        if k and k in text:
            user_hits.append(k)
    # Çok genel kullanıcı kelimeleri tek başına konu filtresi olmasın.
    generic = {"yatırım", "araç", "enerji", "risk", "haber", "türkiye", "türk"}
    strong_user_hits = [k for k in user_hits if k not in generic]
    return bool(core_hits or strong_user_hits)

def is_turkish_priority_source(url):
    d = domain_of(url)
    return (
        any(x in d for x in TR_MAIN_DOMAINS)
        or any(x in d for x in TR_TECH_DEFENSE_DOMAINS)
        or any(x in d for x in TR_OFFICIAL_DOMAINS)
        or any(x in d for x in SOCIAL_SEARCH_DOMAINS)
        or any(x in d for x in GR_DOMAINS)
    )

def source_group(url):
    d = domain_of(url)
    if any(x in d for x in TR_OFFICIAL_DOMAINS):
        return "🇹🇷 Resmi / Kurumsal"
    if any(x in d for x in TR_TECH_DEFENSE_DOMAINS):
        return "🇹🇷 Türk Teknoloji / Savunma Medyası"
    if any(x in d for x in TR_MAIN_DOMAINS):
        return "🇹🇷 Türk Ana Akım / Ekonomi Medyası"
    if any(x in d for x in GR_DOMAINS):
        return "🇬🇷 Yunan Medyası — Türk Savunma"
    if any(x in d for x in SOCIAL_SEARCH_DOMAINS):
        return "📱 Sosyal Medya"
    return "🌍 Diğer Yabancı"


def is_greek_turkish_defense(item):
    d = item.get("domain", "")
    if not any(x in d for x in GR_DOMAINS):
        return False
    text = normalize_text(f"{item.get('title','')} {item.get('snippet','')}")
    return any(k in text for k in TR_DEFENSE_TERMS)


def source_priority(url):
    d = domain_of(url)
    if any(x in d for x in TR_MAIN_DOMAINS):
        return 100
    if any(x in d for x in TR_TECH_DEFENSE_DOMAINS):
        return 110
    if any(x in d for x in TR_OFFICIAL_DOMAINS):
        return 120
    if any(x in d for x in GR_DOMAINS):
        return 80
    if any(x in d for x in SOCIAL_SEARCH_DOMAINS):
        return 70
    return 20

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

def build_queries(user_query, broad=True, negative_boost=True, cap=80):
    base = split_query(user_query)
    if broad:
        existing = {x.lower() for x in base}
        for term in BROAD_TERM_BANK:
            if len(base) >= cap:
                break
            if term.lower() not in existing:
                base.append(term)
                existing.add(term.lower())

    # Negatif taramayı ayrı sorgu ailesi olarak tutuyoruz.
    queries = list(base)
    if negative_boost:
        for term in base[:45]:
            queries.append(f'"{term}" (kriz OR iflas OR konkordato OR soruşturma OR yaptırım OR')
            queries[-1] += ' "üretim durdu" OR "fabrika kapandı" OR "siber saldırı" OR risk)'
            if len(queries) >= cap:
                break
    return queries[:cap]


def make_source_query(term, domains=None, when="1d"):
    parts = [f'({term})', f'when:{when}']
    if domains:
        parts.insert(1, "(" + " OR ".join(f"site:{d}" for d in domains) + ")")
    return " ".join(parts)

# -------------------------
# DDGS / GOOGLE NEWS RSS / GDELT
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


def resolve_google_news_url(url, source_url=""):
    """Google News RSS bağlantısını mümkünse gerçek yayıncı URL'sine çözer."""
    if source_url and url.startswith("https://news.google.com"):
        # Source elementi doğrudan yayıncı domainini verdiğinde en azından
        # doğru kaynak grubunu koruyoruz. İçeriğin tam URL'sini çözmek için
        # redirect isteği yapılır; başarısız olursa Google URL'si korunur.
        try:
            r = requests.get(url, headers=HEADERS, timeout=8, allow_redirects=True)
            if r.url and "news.google.com" not in urlparse(r.url).netloc:
                return r.url
        except Exception:
            pass
    return url


def google_news_rss(query, max_results=100):
    """Google News RSS üzerinden özellikle Türk yayıncıları toplar."""
    url = "https://news.google.com/rss/search"
    params = {"q": query, "hl": "tr", "gl": "TR", "ceid": "TR:tr"}
    try:
        r = requests.get(url, params=params, headers=HEADERS, timeout=15)
        r.raise_for_status()
        root = ET.fromstring(r.content)
        rows = []
        for item in root.findall(".//item")[:max_results]:
            title = item.findtext("title") or ""
            link = item.findtext("link") or ""
            desc = item.findtext("description") or ""
            pub = item.findtext("pubDate") or ""
            source_el = item.find("source")
            source = source_el.text if source_el is not None else "Google News"
            source_url = source_el.get("url", "") if source_el is not None else ""
            rows.append({
                "title": html.unescape(title),
                "url": link,
                "date": pub,
                "source": source,
                "source_url": source_url,
                "body": BeautifulSoup(desc, "html.parser").get_text(" ", strip=True),
                "snippet": BeautifulSoup(desc, "html.parser").get_text(" ", strip=True),
                "image": "",
            })
        return rows
    except Exception:
        return []


def ddgs_text(term, region="tr-tr", max_results=20):
    """Genel web / sosyal platform indeks sonuçları için DDGS text araması."""
    try:
        from ddgs import DDGS
    except ImportError:
        try:
            from duckduckgo_search import DDGS
        except ImportError:
            return []
    try:
        with DDGS() as ddgs:
            return list(ddgs.text(term, region=region, timelimit="d", max_results=max_results))
    except Exception:
        return []


def gdelt_news(query, max_records=50, timespan="1d"):
    url = "https://api.gdeltproject.org/api/v2/doc/doc"
    params = {
        "query": query,
        "mode": "artlist",
        "maxrecords": min(max_records, 250),
        "format": "json",
        "sort": "HybridRel",
        "timespan": timespan,
    }
    try:
        r = requests.get(url, params=params, headers=HEADERS, timeout=15)
        r.raise_for_status()
        data = r.json()
        return data.get("articles", []) or []
    except Exception:
        return []


def fetch_targeted_turkish_sources(user_terms, when="1d", max_per_query=80):
    """Türk kaynaklarını gerçekten domain-hedefli sorgularla toplar.
    Tek dev OR sorgusu yerine birkaç paralel kaynak/konu sepeti kullanır.
    """
    main_1 = TR_MAIN_DOMAINS[:18]
    main_2 = TR_MAIN_DOMAINS[18:]
    tech = TR_TECH_DEFENSE_DOMAINS
    official = TR_OFFICIAL_DOMAINS
    all_tr = list(dict.fromkeys(TR_MAIN_DOMAINS + TR_TECH_DEFENSE_DOMAINS + TR_OFFICIAL_DOMAINS))

    topic_queries = [
        '"sanayi" OR "üretim" OR "fabrika" OR "imalat" OR "OSB"',
        '"teknoloji" OR "yapay zeka" OR "siber güvenlik" OR "yazılım" OR "çip" OR "Ar-Ge"',
        '"savunma sanayii" OR "ASELSAN" OR "TUSAŞ" OR "ROKETSAN" OR "HAVELSAN" OR "Baykar" OR "İHA" OR "KAAN"',
        '"otomotiv" OR "TOGG" OR "batarya" OR "enerji" OR "havacılık" OR "uzay" OR "uydu"',
    ]
    user_q = " OR ".join(f'"{q}"' for q in user_terms[:20] if len(q) > 2)
    negative_q = '("iflas" OR "konkordato" OR "üretim durdu" OR "fabrika kapandı" OR "işten çıkarma" OR "soruşturma" OR "yaptırım" OR "siber saldırı" OR "veri sızıntısı" OR "ihale iptal" OR "gecikme" OR "tedarik krizi") AND (' + ' OR '.join(f'"{x}"' for x in CORE_TOPIC_TERMS[:55]) + ')'

    jobs = [
        (f'({topic_queries[0]})', main_1),
        (f'({topic_queries[1]})', main_2),
        (f'({topic_queries[2]})', tech),
        (f'({topic_queries[3]})', all_tr),
        (negative_q, all_tr),
    ]
    if user_q:
        jobs.append((f'({user_q})', all_tr))

    rows = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=min(6, len(jobs))) as ex:
        futures = [ex.submit(google_news_rss, make_source_query(q, domains, when), max_per_query) for q, domains in jobs]
        for f in concurrent.futures.as_completed(futures):
            try:
                rows.extend(f.result())
            except Exception:
                pass
    return rows

def fetch_greek_defense_sources(when="1d", max_per_query=60):
    """Yunan medyasını sadece Türkiye/Türk savunma bağlamında getirir."""
    queries = [
        "(Turkey OR Türkiye OR Turkish OR Τουρκία OR τουρκική) (defense OR defence OR savunma OR άμυνα OR arms)",
        "(Turkish OR Turkey OR Τουρκία OR τουρκική) (Baykar OR Bayraktar OR ASELSAN OR TUSAŞ OR Roketsan OR KAAN)",
        "(Turkey OR Turkish) (F-16 OR F-35 OR fighter OR missile OR drone OR UAV)",
        "(Turkey OR Turkish) (navy OR frigate OR submarine OR corvette)",
    ]
    rows = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=6) as ex:
        futures = [
            ex.submit(
                google_news_rss,
                make_source_query(q, GR_DOMAINS, when),
                max_per_query,
            )
            for q in queries
        ]
        for f in concurrent.futures.as_completed(futures):
            try:
                rows.extend(f.result())
            except Exception:
                pass
    return rows


def fetch_social_leads(queries, when="1d", max_results=30):
    """İndekslenmiş sosyal sonuçları az sayıda paralel sorguyla toplar."""
    social_domains = " OR ".join(f"site:{d}" for d in SOCIAL_SEARCH_DOMAINS)
    qs = [
        f"({queries[0] if queries else 'sanayi OR teknoloji'}) ({social_domains})",
        f"(ASELSAN OR TUSAŞ OR ROKETSAN OR BAYKAR OR TOGG OR TÜBİTAK) ({social_domains})",
        f"(kriz OR iflas OR soruşturma OR siber saldırı OR üretim durdu OR yaptırım) (sanayi OR teknoloji OR savunma) ({social_domains})",
    ]
    rows = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=3) as ex:
        futures = [ex.submit(ddgs_text, q, "tr-tr", max_results) for q in qs]
        for f in concurrent.futures.as_completed(futures):
            try:
                rows.extend(f.result())
            except Exception:
                pass
    return rows


# -------------------------
# HABER METNİ / GÖRSEL
# -------------------------
@st.cache_data(ttl=3600, show_spinner=False)
def extract_article(url, timeout=5):
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

def extract_many(urls, workers=10):
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
@st.cache_data(ttl=90, show_spinner=False)
def fetch_news(user_query, time_hours=24, max_results=150,
               broad=True, negative_boost=True, include_greek=True,
               include_global=False, include_social=True):
    """V6: hızlı, Türk medya öncelikli, teşhis edilebilir tarama."""
    user_terms = split_query(user_query)
    if broad:
        for term in BROAD_TERM_BANK:
            if term.lower() not in {x.lower() for x in user_terms} and len(user_terms) < 35:
                user_terms.append(term)

    cutoff = datetime.now(timezone.utc) - timedelta(hours=time_hours)
    # 3 saatlik pencerede 6 saatlik kaynak tarayıp Python tarafında 3 saate kesiyoruz;
    # aksi halde dar RSS penceresi bazı yayıncıları kaçırabiliyor.
    if time_hours <= 3:
        when = "6h"
    elif time_hours <= 24:
        when = "1d"
    elif time_hours <= 24 * 7:
        when = "7d"
    else:
        when = "30d"

    records = fetch_targeted_turkish_sources(user_terms, when=when, max_per_query=80)

    if include_social:
        social_queries = [
            'Türkiye sanayi teknoloji üretim savunma',
            'ASELSAN TUSAŞ ROKETSAN HAVELSAN Baykar TOGG',
            'Türkiye teknoloji siber çip yapay zeka',
        ]
        if negative_boost:
            social_queries.append('Türkiye sanayi teknoloji kriz iflas soruşturma siber saldırı')
        with concurrent.futures.ThreadPoolExecutor(max_workers=4) as ex:
            futures = [ex.submit(ddgs_text, q, "tr-tr", 20) for q in social_queries]
            for f in concurrent.futures.as_completed(futures):
                try: records.extend(f.result())
                except Exception: pass

    if include_greek:
        greek = fetch_greek_defense_sources(when=when, max_per_query=60)
        records.extend([r for r in greek if is_greek_turkish_defense({
            "domain": domain_of(r.get("url", "")),
            "title": r.get("title", ""),
            "snippet": r.get("snippet", ""),
        })])

    if include_global:
        gdelt_queries = [
            '(industry OR manufacturing OR technology) (Turkey OR Türkiye)',
            '(defense OR aerospace OR automotive OR semiconductor) (Turkey OR Türkiye)',
        ]
        with concurrent.futures.ThreadPoolExecutor(max_workers=2) as ex:
            futures = [ex.submit(gdelt_news, q, 80, when) for q in gdelt_queries]
            for f in concurrent.futures.as_completed(futures):
                try: records.extend(f.result())
                except Exception: pass

    normalized, seen_urls, seen_titles = [], set(), set()
    stats = {"ham": len(records), "zaman": 0, "konu": 0, "kaynak": 0, "yunan": 0}
    for r in records:
        url = r.get("url") or r.get("url_mobile") or ""
        title = r.get("title") or ""
        if not url or not title: continue
        dt = parse_dt(r.get("date") or r.get("seendate") or r.get("publishedAt"))
        if dt and dt < cutoff:
            stats["zaman"] += 1; continue
        if time_hours <= 3 and not dt:
            stats["zaman"] += 1; continue

        source_url = r.get("source_url") or ""
        candidate_domain = domain_of(source_url) or domain_of(url)
        snippet = r.get("body") or r.get("snippet") or ""
        if not is_relevant_industry_tech(title, snippet, user_terms):
            stats["konu"] += 1; continue
        if not is_turkish_priority_source(candidate_domain):
            stats["kaynak"] += 1; continue
        if any(x in candidate_domain for x in GR_DOMAINS) and not is_greek_turkish_defense({"domain": candidate_domain, "title": title, "snippet": snippet}):
            stats["yunan"] += 1; continue

        nt = normalize_title(title)
        title_key = " ".join(nt.split()[:10])
        if url in seen_urls or title_key in seen_titles: continue
        item = {
            "title": html.unescape(title), "url": url, "published_dt": dt,
            "published": fmt_dt(dt), "source": str(r.get("source") or candidate_domain or "Açık Kaynak"),
            "domain": candidate_domain, "snippet": html.unescape(snippet),
            "image_url": r.get("image") or r.get("socialimage") or "",
            "Kaynak_Grubu": source_group(candidate_domain),
        }
        normalized.append(item); seen_urls.add(url); seen_titles.add(title_key)

    normalized.sort(key=lambda x: (x["published_dt"] or datetime.min.replace(tzinfo=timezone.utc), source_priority(x["url"])), reverse=True)
    return normalized[:max_results], stats

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

    # Ekrandaki kronolojik sıralamayı aynen koru: en yeni → en eski.
    ordered = df.copy()
    if "Tarih_dt" in ordered.columns:
        ordered = ordered.sort_values("Tarih_dt", ascending=False, na_position="last")

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
# HABER TOPLAMA — V7
# -------------------------
# Hızlı ilk bakış için RSS tabanlı toplama; tam metin/görsel yalnızca DOCX
# hazırlanırken veya kullanıcı isterse alınır. Her TARAMAYI BAŞLAT tıklaması
# yeni bir ağ taraması başlatır; haber arama sonuçları bu nedenle cache'lenmez.

def _period_to_when(hours):
    if hours <= 3:
        return "6h"       # daha geniş çek, sonra 3 saate kes
    if hours <= 24:
        return "1d"
    if hours <= 48:
        return "2d"
    if hours <= 24 * 7:
        return "7d"
    return "30d"


def _unique_preserve(rows):
    out=[]; seen_url=set(); seen_title=set()
    for r in rows:
        u=(r.get("url") or r.get("link") or "").strip()
        t=normalize_title(r.get("title") or "")
        key=" ".join(t.split()[:14])
        if not u or not t or u in seen_url or key in seen_title:
            continue
        seen_url.add(u); seen_title.add(key); out.append(r)
    return out


def _fetch_rss_jobs(jobs):
    rows=[]
    with concurrent.futures.ThreadPoolExecutor(max_workers=min(10, len(jobs) or 1)) as ex:
        futures=[ex.submit(google_news_rss, q, 100) for q in jobs]
        for f in concurrent.futures.as_completed(futures):
            try: rows.extend(f.result())
            except Exception: pass
    return rows


def fetch_targeted_turkish_sources_v7(user_query, when):
    # Birkaç geniş konu sepeti: her biri tüm Türk yayıncı havuzunu hedefler.
    all_tr=list(dict.fromkeys(TR_MAIN_DOMAINS+TR_TECH_DEFENSE_DOMAINS+TR_OFFICIAL_DOMAINS))
    site_clause="("+" OR ".join(f"site:{d}" for d in all_tr)+")"
    qparts=split_query(user_query)
    user_clause=" OR ".join(f'"{x}"' for x in qparts[:25] if len(x)>2)
    baskets=[
        '(sanayi OR imalat OR üretim OR fabrika OR OSB OR "organize sanayi" OR makine OR otomasyon)',
        '(teknoloji OR "yapay zeka" OR "yapay zekâ" OR yazılım OR "siber güvenlik" OR çip OR "yarı iletken" OR elektronik OR robotik OR "Ar-Ge" OR patent)',
        '("savunma sanayii" OR "savunma sanayi" OR ASELSAN OR TUSAŞ OR ROKETSAN OR HAVELSAN OR Baykar OR Bayraktar OR İHA OR SİHA OR KAAN OR Kızılelma OR füze OR roket)',
        '(otomotiv OR TOGG OR batarya OR "elektrikli araç" OR enerji OR hidrojen OR "güneş enerjisi" OR "rüzgar enerjisi" OR havacılık OR uzay OR uydu OR gemi OR tersane)',
        '("tedarik zinciri" OR yerlileştirme OR millileştirme OR yatırım teşvik OR teknopark OR TÜBİTAK OR KOSGEB OR "Sanayi ve Teknoloji Bakanlığı")',
    ]
    if user_clause:
        baskets.append(f'({user_clause})')
    jobs=[f'{b} {site_clause} when:{when}' for b in baskets]
    return _fetch_rss_jobs(jobs)


def fetch_negative_turkish_sources_v7(when):
    all_tr=list(dict.fromkeys(TR_MAIN_DOMAINS+TR_TECH_DEFENSE_DOMAINS+TR_OFFICIAL_DOMAINS))
    site_clause="("+" OR ".join(f"site:{d}" for d in all_tr)+")"
    negative='("iflas" OR "konkordato" OR "üretim durdu" OR "üretim durduruldu" OR "fabrika kapandı" OR "işten çıkarma" OR "işçi çıkarma" OR grev OR lokavt OR soruşturma OR dava OR ceza OR "geri çağırma" OR "siber saldırı" OR "veri sızıntısı" OR "fidye yazılımı" OR ambargo OR yaptırım OR "ihracat yasağı" OR "sözleşme feshi" OR "ihale iptal" OR "askıya alındı" OR ertelendi OR gecikme OR "tedarik krizi" OR "çip krizi" OR daralma OR "kapasite kaybı" OR "güvenlik açığı" OR zafiyet OR usulsüzlük OR yolsuzluk)'
    topic='(sanayi OR üretim OR fabrika OR teknoloji OR savunma OR havacılık OR otomotiv OR enerji OR çip OR yazılım OR siber OR şirket OR tesis OR proje)'
    return _fetch_rss_jobs([f'{negative} {topic} {site_clause} when:{when}'])


def fetch_greek_v7(when):
    site_clause="("+" OR ".join(f"site:{d}" for d in GR_DOMAINS)+")"
    q=[
        '(Turkey OR Türkiye OR Turkish OR Τουρκία OR τουρκική) (defense OR defence OR savunma OR άμυνα)',
        '(Baykar OR Bayraktar OR ASELSAN OR TUSAŞ OR Roketsan OR HAVELSAN OR KAAN OR Kızılelma OR SİPER OR HİSAR)',
        '(Turkish OR Turkey OR Τουρκία) (drone OR UAV OR missile OR fighter OR navy OR frigate OR submarine OR defense industry)',
    ]
    rows=_fetch_rss_jobs([f'{x} {site_clause} when:{when}' for x in q])
    return [r for r in rows if is_greek_turkish_defense({"domain":domain_of(r.get("source_url") or r.get("url","")),"title":r.get("title",""),"snippet":r.get("snippet") or r.get("body","")})]


def fetch_social_v7(when):
    # Sosyal ağlar indekslenmiş açık sonuç olarak ayrı bir katmandır.
    qs=[
        'Türkiye (sanayi OR teknoloji OR üretim OR savunma) (site:x.com OR site:twitter.com OR site:youtube.com OR site:linkedin.com)',
        '(ASELSAN OR TUSAŞ OR ROKETSAN OR HAVELSAN OR Baykar OR TOGG) (site:x.com OR site:twitter.com OR site:youtube.com OR site:linkedin.com)',
    ]
    rows=[]
    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as ex:
        fs=[ex.submit(ddgs_text,q,"tr-tr",40) for q in qs]
        for f in concurrent.futures.as_completed(fs):
            try: rows.extend(f.result())
            except Exception: pass
    return rows


def fetch_global_v7(when):
    qs=[
        '(Turkey OR Türkiye) (industry OR manufacturing OR technology OR semiconductor OR defense OR aerospace OR automotive)',
        '(Turkish OR Turkey) (Baykar OR ASELSAN OR TUSAŞ OR ROKETSAN OR HAVELSAN OR KAAN OR drone OR missile)',
    ]
    rows=[]
    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as ex:
        fs=[ex.submit(gdelt_news,q,250,when) for q in qs]
        for f in concurrent.futures.as_completed(fs):
            try: rows.extend(f.result())
            except Exception: pass
    return rows


def fetch_news_v7(user_query, time_hours=24, broad=True, negative_boost=True,
                  include_greek=True, include_global=False, include_social=True):
    cutoff=datetime.now(timezone.utc)-timedelta(hours=time_hours)
    when=_period_to_when(time_hours)
    stats={"Ham sonuç":"başlıyor", "Türk RSS":0, "Negatif RSS":0, "Yunan":0, "Sosyal":0, "Global":0,
           "Zaman filtresi":0, "Konu filtresi":0, "Kaynak filtresi":0, "Sonuç":0}

    # Türk ana akış + negatif akış + özel Yunan/global/sosyal paralel.
    jobs=[]
    jobs.append(("tr", lambda: fetch_targeted_turkish_sources_v7(user_query, when)))
    if negative_boost: jobs.append(("neg", lambda: fetch_negative_turkish_sources_v7(when)))
    if include_greek: jobs.append(("gr", lambda: fetch_greek_v7(when)))
    if include_social: jobs.append(("social", lambda: fetch_social_v7(when)))
    if include_global: jobs.append(("global", lambda: fetch_global_v7(when)))

    records=[]
    with concurrent.futures.ThreadPoolExecutor(max_workers=min(5,len(jobs))) as ex:
        fs={ex.submit(fn): name for name,fn in jobs}
        for f in concurrent.futures.as_completed(fs):
            name=fs[f]
            try:
                data=f.result() or []
                stats[{'tr':'Türk RSS','neg':'Negatif RSS','gr':'Yunan','social':'Sosyal','global':'Global'}[name]]+=len(data)
                records.extend(data)
            except Exception: pass
    stats["Ham sonuç"]=len(records)

    normalized=[]; seen_urls=set(); seen_titles=set()
    for r in records:
        url=(r.get("url") or r.get("url_mobile") or r.get("link") or "").strip()
        title=html.unescape((r.get("title") or "").strip())
        if not url or not title: continue
        dt=parse_dt(r.get("date") or r.get("seendate") or r.get("publishedAt"))
        if dt and dt < cutoff:
            stats["Zaman filtresi"]+=1; continue
        snippet=html.unescape((r.get("body") or r.get("snippet") or r.get("description") or "").strip())
        candidate_domain=domain_of(r.get("source_url") or "") or domain_of(url)
        # Sosyal/Global sonuçları kaynak filtresinden geçirmek yerine kendi sınıflarında tut.
        is_gr=any(x in candidate_domain for x in GR_DOMAINS)
        if is_gr:
            if not is_greek_turkish_defense({"domain":candidate_domain,"title":title,"snippet":snippet}):
                stats["Konu filtresi"]+=1; continue
        elif not (is_turkish_priority_source(candidate_domain) or include_global):
            stats["Kaynak filtresi"]+=1; continue

        if not is_gr and not is_relevant_industry_tech(title,snippet,split_query(user_query)):
            # Global ve sosyal için de aynı konu filtresi; konu dışı gündemi engeller.
            stats["Konu filtresi"]+=1; continue

        nt=normalize_title(title); title_key=" ".join(nt.split()[:14])
        if url in seen_urls or title_key in seen_titles: continue
        group=source_group(candidate_domain)
        item={"title":title,"url":url,"published_dt":dt,"published":fmt_dt(dt),
              "source":str(r.get("source") or candidate_domain or "Açık Kaynak"),
              "domain":candidate_domain,"snippet":snippet,
              "image_url":r.get("image") or r.get("socialimage") or "",
              "Kaynak_Grubu":group}
        normalized.append(item); seen_urls.add(url); seen_titles.add(title_key)

    normalized.sort(key=lambda x:(x["published_dt"] is not None, x["published_dt"] or datetime.min.replace(tzinfo=timezone.utc)), reverse=True)
    stats["Sonuç"]=len(normalized)
    return normalized, stats

# -------------------------
# STREAMLIT
# -------------------------
st.title("🛡️ Sanayi & Teknoloji Açık Kaynak / Negatif Haber Radarı")
st.caption("İlk bakışta hızlı haber akışı · negatif/yüksek risk ayrıştırma · kronolojik saat/tarih · Türk medya önceliği · Yunan/Türk savunma · seçilen haberlerden DOCX")

with st.sidebar:
    st.header("⚙️ Tarama Ayarları")
    default_query=(
        'sanayi OR teknoloji OR üretim OR imalat OR fabrika OR OSB OR Ar-Ge OR patent OR '
        'yapay zeka OR siber güvenlik OR çip OR yarı iletken OR elektronik OR otomasyon OR robotik OR '
        'savunma sanayii OR ASELSAN OR TUSAŞ OR ROKETSAN OR HAVELSAN OR Baykar OR Bayraktar OR İHA OR SİHA OR KAAN OR '
        'havacılık OR uzay OR uydu OR otomotiv OR TOGG OR batarya OR enerji OR hidrojen OR '
        'TÜBİTAK OR KOSGEB OR teknopark OR yerlileştirme OR tedarik zinciri'
    )
    query=st.text_area("Geniş sanayi / teknoloji sorgusu:",value=default_query,height=180)
    broad=st.checkbox("🔎 Geniş Sanayi & Teknoloji evrenini otomatik genişlet",value=True)
    negative_boost=st.checkbox("⚠️ Negatif haber taramasını güçlendir",value=True)
    include_greek=st.checkbox("🇬🇷 Yunan medyası — yalnızca Türk savunma sanayii",value=True)
    include_social=st.checkbox("📱 Türk açık sosyal kaynak / indeks sonuçları",value=True)
    include_global=st.checkbox("🌍 Global basın (opsiyonel)",value=False)
    enrich=st.checkbox("📖 Tarama sonrası seçilen/öncelikli haberlerde tam metin + görsel zenginleştir",value=False)
    period_label=st.selectbox("🕒 Haber dönemi",[
        "⚡ Anlık — son 3 saat","📅 Son 24 saat","📆 Son 48 saat","📆 Son 1 hafta","🗓️ Son 1 ay"
    ],index=1)
    hours={"⚡ Anlık — son 3 saat":3,"📅 Son 24 saat":24,"📆 Son 48 saat":48,"📆 Son 1 hafta":168,"🗓️ Son 1 ay":720}[period_label]
    run=st.button("🔍 TARAMAYI BAŞLAT / YENİLE",type="primary",use_container_width=True)
    if st.button("🧹 Önbelleği Temizle",use_container_width=True):
        st.cache_data.clear(); st.session_state.df=None; st.rerun()

if "df" not in st.session_state: st.session_state.df=None
if "last_query" not in st.session_state: st.session_state.last_query=""
if "last_scan_time" not in st.session_state: st.session_state.last_scan_time=None
if "scan_stats" not in st.session_state: st.session_state.scan_stats={}

# İlk açılışta ASLA otomatik tarama yok.
if run:
    scan_nonce=datetime.now(timezone.utc).isoformat()
    with st.status("🔎 Tarama başlatıldı — önce hızlı başlık/saat/kaynak akışı getiriliyor...",expanded=True) as status:
        st.write("🇹🇷 Türk medya + teknoloji/savunma + resmi kaynaklar taranıyor...")
        raw,stats=fetch_news_v7(query,hours,broad,negative_boost,include_greek,include_global,include_social)
        st.session_state.scan_stats=stats
        st.write(f"Toplanan ham sonuç: {stats.get('Ham sonuç',0)} → gösterilecek: {len(raw)}")
        # Hızlı ilk bakış: tam metin yok. İstenirse yalnızca negatif/riskli ve ilk haberler zenginleştirilebilir.
        enriched={}
        if enrich and raw:
            priority=[]
            for x in raw:
                probe=normalize_text(x.get('title','')+' '+x.get('snippet',''))
                if any(k in probe for k in NEGATIVE_EVENT_TERMS+HIGH_RISK_KEYWORDS): priority.append(x)
            priority += [x for x in raw if x not in priority]
            urls=[x['url'] for x in priority[:40] if x.get('url')]
            if urls:
                st.write(f"📖 {len(urls)} öncelikli haberin tam metni/görseli alınıyor...")
                enriched=extract_many(urls,workers=12)
        rows=[]
        for item in raw:
            extra=enriched.get(item['url'],{})
            body=extra.get('text') or item.get('snippet','')
            a=analyze_article(item['title'],body)
            rows.append({
                'Tarih_dt':item['published_dt'],'Tarih':item['published'],'Kaynak':item['source'],
                'Kaynak_Grubu':item['Kaynak_Grubu'],'Domain':item['domain'],'Kategori':a['category'],
                'Başlık':item['title'],'İçerik_Özeti':summarize_text(item['title'],body,2200),
                'Özet':summarize_text(item['title'],body,2200),'Duygu':a['sentiment'],'Skor':a['score'],
                'Risk_Durumu':a['risk'],'Negatif_Sinyaller':a['negative_hits'],'Risk_Sinyalleri':a['risk_hits'],
                'Pozitif_Sinyaller':a['positive_hits'],'URL':item['url'],'Görsel_URL':item['image_url'] or extra.get('image_url',''),
                'Tam_Metin_Uzunluğu':len(body or ''),'Seç':a['risk'] in ('Yüksek Risk','İzleme') or a['sentiment']=='Negatif'
            })
        df=pd.DataFrame(rows)
        if not df.empty: df=df.sort_values('Tarih_dt',ascending=False,na_position='last').reset_index(drop=True)
        st.session_state.df=df; st.session_state.last_query=query; st.session_state.last_scan_time=datetime.now().astimezone()
        status.update(label=f"✅ Tarama tamamlandı — {len(df)} haber",state="complete")

df=st.session_state.df
if df is None:
    st.info("👋 İlk bakış hazır. **TARAMAYI BAŞLAT / YENİLE** düğmesine bastığınızda seçtiğiniz zaman aralığındaki haberler yeni bir taramayla güncellenir.")
else:
    if st.session_state.last_scan_time:
        st.caption(f"Son tarama: {st.session_state.last_scan_time.strftime('%d.%m.%Y %H:%M:%S')}")
    with st.expander("🧪 Tarama teşhisi",expanded=False): st.json(st.session_state.scan_stats)
    if df.empty:
        st.warning("Seçilen zaman penceresinde kriterlere uygun haber bulunamadı.")
    else:
        total=len(df); neg=int((df.Duygu=='Negatif').sum()); high=int((df.Risk_Durumu=='Yüksek Risk').sum()); watch=int((df.Risk_Durumu=='İzleme').sum())
        k1,k2,k3,k4=st.columns(4); k1.metric('Toplam Haber',total); k2.metric('Negatif',neg); k3.metric('Yüksek Risk',high); k4.metric('İzleme',watch)
        tr_count=int(df.Kaynak_Grubu.astype(str).str.startswith('🇹🇷').sum()); gr_count=int(df.Kaynak_Grubu.astype(str).str.startswith('🇬🇷').sum()); social_count=int(df.Kaynak_Grubu.astype(str).str.startswith('📱').sum())
        k5,k6,k7=st.columns(3); k5.metric('🇹🇷 Türk Kaynakları',tr_count); k6.metric('🇬🇷 Yunan / Türk Savunma',gr_count); k7.metric('📱 Sosyal / Açık Kaynak',social_count)
        tab_all,tab_tr,tab_neg,tab_risk,tab_gr,tab_osint=st.tabs([f'📰 Kronolojik Akış ({total})',f'🇹🇷 Türk Medyası ({tr_count})',f'⚠️ Negatif ({neg})',f'🚨 Yüksek Risk ({high})',f'🇬🇷 Yunan / Türk Savunma ({gr_count})','🔎 Açık Kaynak Çalışma Masası'])
        display_cols=['Seç','Tarih','Kaynak_Grubu','Kaynak','Kategori','Başlık','İçerik_Özeti','Duygu','Skor','Risk_Durumu','URL']
        with tab_all:
            edited=st.data_editor(df[display_cols],column_config={'Seç':st.column_config.CheckboxColumn('Bilgi Notuna Ekle'),'URL':st.column_config.LinkColumn('Haber Linki'),'İçerik_Özeti':st.column_config.TextColumn('İçerik / Özet',width='large')},disabled=[c for c in display_cols if c!='Seç'],hide_index=True,use_container_width=True,height=650,key='news_editor_v7')
            df.loc[edited.index,'Seç']=edited['Seç']; st.session_state.df=df
        with tab_tr:
            t=df[df.Kaynak_Grubu.astype(str).str.startswith('🇹🇷')]
            st.dataframe(t[['Tarih','Kaynak','Kategori','Başlık','İçerik_Özeti','Duygu','Risk_Durumu','URL']],column_config={'URL':st.column_config.LinkColumn('Haber Linki'),'İçerik_Özeti':st.column_config.TextColumn('İçerik / Özet',width='large')},hide_index=True,use_container_width=True,height=650)
        with tab_neg:
            t=df[df.Duygu=='Negatif']; st.dataframe(t[['Tarih','Kaynak','Kategori','Başlık','İçerik_Özeti','Negatif_Sinyaller','URL']],column_config={'URL':st.column_config.LinkColumn('Haber Linki')},hide_index=True,use_container_width=True,height=650)
        with tab_risk:
            t=df[df.Risk_Durumu=='Yüksek Risk']; st.dataframe(t[['Tarih','Kaynak','Kategori','Başlık','İçerik_Özeti','Risk_Sinyalleri','URL']],column_config={'URL':st.column_config.LinkColumn('Haber Linki')},hide_index=True,use_container_width=True,height=650)
        with tab_gr:
            t=df[df.Kaynak_Grubu.astype(str).str.startswith('🇬🇷')]; st.dataframe(t[['Tarih','Kaynak','Kategori','Başlık','İçerik_Özeti','Duygu','Risk_Durumu','URL']],column_config={'URL':st.column_config.LinkColumn('Haber Linki')},hide_index=True,use_container_width=True,height=650)
        with tab_osint:
            st.subheader('🔎 Seçili Haber Üzerinden Açık Kaynak Çalışması'); selected=df[df['Seç']==True]
            if selected.empty: st.info('Önce haber listesinden bir veya daha fazla haber seçin.')
            else:
                choice=st.selectbox('İncelenecek haber:',list(selected.index),format_func=lambda i:selected.loc[i,'Başlık']); row=selected.loc[choice]
                st.markdown(f"### {row['Başlık']}"); st.write(row['İçerik_Özeti']); st.markdown(f"**Kaynak:** {row['Kaynak']} · **Tarih:** {row['Tarih']} · **Kategori:** {row['Kategori']}"); st.link_button('🌐 Haberi Aç',row['URL'])
                st.markdown('**Negatif sinyaller:** '+(', '.join(row['Negatif_Sinyaller']) or 'Yok')); st.markdown('**Risk sinyalleri:** '+(', '.join(row['Risk_Sinyalleri']) or 'Yok'))
                st.text_area('Analist notu',height=180,key=f'note_v7_{choice}')

        st.markdown('---'); st.subheader('📥 Çıktılar')
        selected_df=df[df['Seç']==True]; st.write(f'**{len(selected_df)} haber** bilgi notuna seçildi.')
        if 'briefing_docx' not in st.session_state: st.session_state.briefing_docx=None
        if st.button('📝 SEÇİLEN HABERLERDEN DOCX BİLGİ NOTU OLUŞTUR',type='primary',use_container_width=True):
            with st.spinner('Seçilen haberlerin geniş özeti ve görselleri hazırlanıyor...'):
                # DOCX öncesi seçilen haberlerin tam metin/görsellerini al.
                urls=list(dict.fromkeys(selected_df['URL'].tolist()))
                details=extract_many(urls,workers=12)
                doc_df=selected_df.copy()
                for idx,row in doc_df.iterrows():
                    ex=details.get(row['URL'],{}); body=ex.get('text') or row['Özet']; doc_df.at[idx,'Özet']=summarize_text(row['Başlık'],body,5000); doc_df.at[idx,'Görsel_URL']=row['Görsel_URL'] or ex.get('image_url','')
                st.session_state.briefing_docx=make_docx(doc_df,st.session_state.last_query,title='GÜNLÜK BİLGİ NOTU — SANAYİ & TEKNOLOJİ')
        if st.session_state.briefing_docx:
            st.download_button('⬇️ BİLGİ NOTUNU İNDİR (.DOCX)',st.session_state.briefing_docx,file_name=f'Gunluk_Bilgi_Notu_{date.today()}.docx',mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',type='primary',use_container_width=True)
        csv=df.drop(columns=['Tarih_dt'],errors='ignore').to_csv(index=False).encode('utf-8-sig')
        st.download_button('📊 HAM OSINT VERİSİNİ İNDİR (.CSV)',csv,file_name=f'Sanayi_Teknoloji_OSINT_{date.today()}.csv',mime='text/csv',use_container_width=True)

st.caption('Not: İlk açılışta otomatik tarama yoktur. Her TARAMAYI BAŞLAT / YENİLE tıklaması yeni bir haber taraması yapar. Türk kaynakları önceliklidir; Yunan kaynakları yalnızca Türk savunma sanayii bağlantılı sonuçlarda tutulur; global basın isteğe bağlıdır. Haberler yayın saatine göre en yeniden eskiye sıralanır. Negatif/risk sınıflandırması otomatik ön elemedir ve kritik bulgular teyit edilmelidir.')