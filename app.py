import streamlit as st
import pandas as pd
import requests
import concurrent.futures
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

st.set_page_config(page_title='Sanayi & Teknoloji OSINT Radarı', page_icon='🛡️', layout='wide')

HEADERS={'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/151 Safari/537.36'}

# -----------------------------
# KONU EVRENİ
# -----------------------------
TOPIC_TERMS = [
    # Sanayi / üretim
    'sanayi','sanayi üretimi','imalat','üretim','fabrika','tesis','organize sanayi','OSB','endüstri',
    'makine','makine sanayii','endüstriyel otomasyon','otomasyon','robotik','endüstri 4.0','mesleki üretim',
    'kapasite','kapasite kullanım','yatırım','yatırım teşvik','yerli üretim','yerlileştirme','millileştirme',
    'tedarik zinciri','tedarikçi','lojistik','depo','depolama','tersane','gemi inşa','denizcilik',
    # Teknoloji / dijital
    'teknoloji','teknolojik','Ar-Ge','Arge','araştırma geliştirme','inovasyon','patent','faydalı model',
    'dijital dönüşüm','endüstri 4.0','yapay zeka','yapay zekâ','makine öğrenmesi','derin öğrenme','yazılım',
    'siber güvenlik','siber saldırı','veri sızıntısı','veri merkezi','bulut','cloud','saas','yazılım şirketi',
    'çip','mikroçip','yarı iletken','semiconductor','işlemci','wafer','elektronik','pcb','sensör',
    'telekom','5G','6G','fiber','internet altyapısı','kuantum','blokzincir','blockchain','fintech',
    # İleri teknoloji / sağlık teknolojileri
    'biyoteknoloji','biyomedikal','nanoteknoloji','medikal cihaz','sağlık teknolojisi','gen tedavisi',
    'malzeme','ileri malzeme','kompozit','karbon fiber','3D yazıcı','eklemeli imalat','batarya teknolojisi',
    # Savunma / havacılık / uzay
    'savunma sanayii','savunma sanayi','savunma teknolojisi','ASELSAN','TUSAŞ','TUSAS','ROKETSAN','HAVELSAN',
    'Baykar','Bayraktar','İHA','SİHA','drone','insansız hava aracı','insansız deniz aracı','KAAN','Kızılelma',
    'HİSAR','SİPER','füze','roket','radar','elektronik harp','elektronik destek','komuta kontrol',
    'mühimmat','zırhlı araç','tank','denizaltı','fırkateyn','korvet','helikopter','havacılık','uçak',
    'havacılık sanayii','uzay','uydu','uydu teknolojisi','roket fırlatma','fırlatma sistemi','Türkiye Uzay Ajansı',
    # Otomotiv / mobilite
    'otomotiv','TOGG','elektrikli araç','hibrit araç','otonom araç','sürücüsüz araç','batarya','şarj',
    'şarj istasyonu','mobilite','raylı sistem','lokomotif','metro','demiryolu','lastik','yan sanayi',
    # Enerji / kimya / kaynak
    'enerji','enerji depolama','güneş enerjisi','solar','rüzgar enerjisi','hidrojen','yakıt hücresi',
    'nükleer enerji','nükleer santral','petrol','doğalgaz','LNG','elektrik üretimi','şebeke','kimya',
    'petrokimya','plastik','polimer','demir çelik','çelik','metal','alüminyum','bakır','madencilik','maden',
    # Diğer üretim sektörleri
    'tekstil','hazır giyim','gıda teknolojisi','gıda sanayii','tarım teknolojisi','akıllı tarım','seracılık',
    'su ürünleri','inşaat teknolojisi','çimento','cam','seramik','kağıt','ambalaj','mobilya',
    # Ekosistem / kamu / girişim
    'TÜBİTAK','KOSGEB','Sanayi ve Teknoloji Bakanlığı','TSE','TürkPatent','TEKNOFEST','teknopark',
    'girişim','girişimcilik','startup','start-up','venture capital','yatırım turu','teknoloji transferi',
    'teknoloji geliştirme bölgesi','Ar-Ge merkezi','tasarım merkezi','OSBÜK','ihracat','ithalat','yüksek teknoloji',
    'orta yüksek teknoloji','kritik teknoloji','stratejik ürün','stratejik yatırım'
]

NEGATIVE_TERMS = [
    'iflas','konkordato','zarar açıkladı','net zarar','üretim durdu','üretim durduruldu',
    'fabrika kapandı','fabrika kapanıyor','işten çıkarma','işçi çıkarma','toplu işten çıkarma',
    'grev','lokavt','soruşturma','dava açıldı','dava','ceza','para cezası','geri çağırma',
    'recall','arıza','kaza','patlama','yangın','siber saldırı','veri sızıntısı','hacklendi',
    'fidye yazılımı','ambargo','yaptırım','ihracat yasağı','ithalat yasağı','lisans reddi',
    'ruhsat iptali','sözleşme feshi','ihale iptal','ihale iptal edildi','askıya alındı',
    'ertelendi','gecikme','teslim edilemedi','testi geçemedi','kapasite kaybı','daralma',
    'sert düşüş','pazar kaybı','maliyet artışı','tedarik sorunu','tedarik krizi','çip krizi',
    'kıtlık','blokaj','güvenlik açığı','kritik zafiyet','zafiyet','usulsüzlük','yolsuzluk',
    'vurgun','casusluk','can kaybı','ölüm','yaralanma','ifşa edildi'
]
HIGH_RISK_TERMS = [
    'iflas','konkordato','üretim durdu','fabrika kapandı','toplu işten çıkarma','siber saldırı',
    'veri sızıntısı','fidye yazılımı','ambargo','yaptırım','ihracat yasağı','lisans reddi',
    'ruhsat iptali','sözleşme feshi','ihale iptal edildi','patlama','yangın','can kaybı','ölüm',
    'kritik zafiyet','yolsuzluk','usulsüzlük','casusluk','tedarik krizi','çip krizi'
]

CATEGORIES={
 'Savunma & Havacılık':['savunma','aselsan','tusaş','tusas','roketsan','havelsan','baykar','bayraktar','iha','siha','kaan','kızılelma','füze','roket','havacılık'],
 'Dijital & Yapay Zeka':['yapay zeka','yapay zekâ','siber','yazılım','5g','6g','veri merkezi','bulut','kuantum'],
 'Yarı İletken & Elektronik':['çip','mikroçip','yarı iletken','işlemci','elektronik','wafer','pcb'],
 'Otomotiv & Mobilite':['otomotiv','togg','elektrikli araç','batarya','şarj'],
 'Enerji':['enerji','hidrojen','güneş','rüzgar','nükleer','enerji depolama'],
 'Sanayi & Üretim':['sanayi','imalat','üretim','fabrika','osb','makine','robotik','otomasyon','demir çelik','kimya'],
 'Uzay & İleri Teknoloji':['uzay','uydu','tua','nanoteknoloji','biyoteknoloji'],
 'Kurumsal Ekosistem':['tübitak','kosgeb','sanayi ve teknoloji bakanlığı','türkpatent','teknopark','teknofest']
}

TR_MAIN=[
 'aa.com.tr','trthaber.com','ntv.com.tr','cnnturk.com','haberturk.com','hurriyet.com.tr','milliyet.com.tr',
 'sabah.com.tr','sozcu.com.tr','cumhuriyet.com.tr','karar.com','yenisafak.com','star.com.tr','aksam.com.tr',
 'turkiyegazetesi.com.tr','t24.com.tr','haber7.com','haberler.com','ensonhaber.com','gazeteduvar.com.tr',
 'odatv.com','medyascope.tv','tv100.com','tgrthaber.com.tr','mynet.com','dunya.com','ekonomim.com',
 'bloomberght.com','paraanaliz.com','bigpara.com','fortuneturkey.com','doviz.com','haberler.com'
]
TR_TECH=[
 'webrazzi.com','shiftdelete.net','donanimhaber.com','chip.com.tr','log.com.tr','technopat.net',
 'hardwareplus.com.tr','turk-internet.com','savunmasanayist.com','savunmatr.com','defenceturk.net',
 'defencehere.com','c4defence.com','savunmahaber.com','gdh.digital','stratejikortak.com','m5dergi.com','mavivatan.net'
]
TR_OFFICIAL=[
 'sanayi.gov.tr','tubitak.gov.tr','kosgeb.gov.tr','tse.org.tr','turkpatent.gov.tr','tua.gov.tr','ticaret.gov.tr',
 'uab.gov.tr','aselsan.com','tusas.com','roketsan.com.tr','havelsan.com.tr','baykartech.com','togg.com.tr','tei.com.tr','tai.com.tr'
]
GR=[
 'kathimerini.gr','protothema.gr','news247.gr','tovima.gr','enikos.gr','naftemporiki.gr','skai.gr','capital.gr',
 'defence-point.gr','defencereview.gr','militaire.gr','pronews.gr','newsbreak.gr','pentapostagma.gr','hellasjournal.com'
]
SOCIAL=['x.com','twitter.com','youtube.com','linkedin.com','facebook.com','instagram.com']

SOURCE_ALIASES={
 'aa':'aa.com.tr','anadolu ajansı':'aa.com.tr','anadolu agency':'aa.com.tr','trt haber':'trthaber.com','trt':'trthaber.com',
 'ntv':'ntv.com.tr','cnn türk':'cnnturk.com','cnn turk':'cnnturk.com','habertürk':'haberturk.com','hürriyet':'hurriyet.com.tr',
 'milliyet':'milliyet.com.tr','sabah':'sabah.com.tr','sözcü':'sozcu.com.tr','cumhuriyet':'cumhuriyet.com.tr','karar':'karar.com',
 'yeni şafak':'yenisafak.com','türkiye gazetesi':'turkiyegazetesi.com.tr','t24':'t24.com.tr','haberler':'haberler.com',
 'dünya':'dunya.com','ekonomim':'ekonomim.com','bloomberg ht':'bloomberght.com','webrazzi':'webrazzi.com',
 'shiftdelete':'shiftdelete.net','donanımhaber':'donanimhaber.com','technopat':'technopat.net','savunma sanayi st':'savunmasanayist.com',
 'savunma sanayi':'savunmasanayist.com','defence türk':'defenceturk.net','defence turk':'defenceturk.net','defencehere':'defencehere.com',
 'c4 defence':'c4defence.com','c4defence':'c4defence.com','sanayi ve teknoloji bakanlığı':'sanayi.gov.tr','tübitak':'tubitak.gov.tr',
 'kosgeb':'kosgeb.gov.tr','türkpatent':'turkpatent.gov.tr','türkiye uzay ajansı':'tua.gov.tr','aselsan':'aselsan.com',
 'tusaş':'tusas.com','tusas':'tusas.com','roketsan':'roketsan.com.tr','havelsan':'havelsan.com.tr','baykar':'baykartech.com','togg':'togg.com.tr'
}

def norm(s):
    return re.sub(r'\s+',' ',str(s or '').lower()).strip()

def title_key(s):
    return re.sub(r'[^\w\s]',' ',norm(s)).strip()[:180]

def domain(url):
    try: return urlparse(url).netloc.lower().replace('www.','')
    except: return ''

def parse_dt(v):
    if not v: return None
    s=str(v).strip()
    for x in (s.replace('Z','+00:00'),s):
        try:
            d=datetime.fromisoformat(x)
            if d.tzinfo is None: d=d.replace(tzinfo=timezone.utc)
            return d.astimezone(timezone.utc)
        except: pass
    try:
        d=parsedate_to_datetime(s)
        if d.tzinfo is None: d=d.replace(tzinfo=timezone.utc)
        return d.astimezone(timezone.utc)
    except: return None

def _to_utc_datetime(value):
    """datetime / pandas.Timestamp / string değerlerini güvenli biçimde UTC-aware datetime'a çevirir."""
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass

    try:
        if isinstance(value, pd.Timestamp):
            ts = value
            if ts.tzinfo is None:
                ts = ts.tz_localize('UTC')
            else:
                ts = ts.tz_convert('UTC')
            return ts.to_pydatetime()
    except Exception:
        pass

    if isinstance(value, datetime):
        if value.tzinfo is None:
            return value.replace(tzinfo=timezone.utc)
        return value.astimezone(timezone.utc)

    try:
        ts = pd.to_datetime(value, utc=True, errors='coerce')
        if pd.isna(ts):
            return None
        return ts.to_pydatetime()
    except Exception:
        return None


def fmt_dt(d):
    d = _to_utc_datetime(d)
    return d.astimezone().strftime('%d.%m.%Y %H:%M:%S') if d else 'Tarih/saat bilinmiyor'

def infer_source(source_name='',source_url='',article_url=''):
    d=domain(source_url)
    if d and d not in ('news.google.com','google.com'): return d
    n=norm(source_name)
    for a,d in SOURCE_ALIASES.items():
        if a in n: return d
    # domain adının yayıncı adına gömülü olması
    for d in TR_MAIN+TR_TECH+TR_OFFICIAL+GR:
        stem=d.split('.')[0]
        if stem and stem in re.sub(r'[^a-z0-9ğüşöçıİĞÜŞÖÇ]','',n): return d
    return domain(article_url)

def source_group(d):
    d=domain(d)
    if d in TR_OFFICIAL: return '🇹🇷 Resmi / Kurumsal'
    if d in TR_TECH: return '🇹🇷 Türk Teknoloji / Savunma'
    if d in TR_MAIN: return '🇹🇷 Türk Medyası / Ekonomi'
    if d in GR: return '🇬🇷 Yunan Medyası — Türk Savunma'
    if d in SOCIAL: return '📱 Açık Sosyal / İndeks'
    return '🌍 Diğer / Açık Kaynak'

def source_rank(d):
    d=domain(d)
    if d in TR_OFFICIAL: return 500
    if d in TR_TECH: return 450
    if d in TR_MAIN: return 400
    if d in GR: return 300
    if d in SOCIAL: return 250
    return 100

def relevant(text,user_query=''):
    t=norm(text)
    if any(x in t for x in TOPIC_TERMS): return True
    uq=re.split(r'\bOR\b|,|\n',user_query or '',flags=re.I)
    generic={'sanayi','teknoloji','üretim','yatırım','enerji','türkiye','türk','haber'}
    return any(len(x.strip())>2 and norm(x.strip()) not in generic and norm(x.strip()) in t for x in uq)

def greek_defense(text):
    t=norm(text)
    terms=['turkey','türkiye','turkish','türk','τουρκ','aselsan','tusaş','tusas','roketsan','havelsan','baykar','bayraktar','kaan','kızılelma','siper','hisar','iha','siha','drone','uav','missile','fighter','frigate','submarine','defense','defence','savunma','άμυνα']
    return any(x in t for x in terms)


OSB_FIRE_LOCATION_TERMS = [
    'osb','organize sanayi','organize sanayi bölgesi','organize sanayi bölgesinde',
    'organize sanayi bölgesindeki','organize sanayi sitesinde'
]
OSB_FIRE_EVENT_TERMS = [
    'yangın','yangını','yangin','alev','alevler','yanıyor','yaniyor','yandı','yandi',
    'fabrika yangını','tesis yangını'
]

def is_osb_fire(title, snippet=''):
    """Organize sanayi bölgesi + yangın bağlamı birlikte geçiyorsa özel alarm üretir."""
    t=norm(f'{title} {snippet}')
    return (
        any(term in t for term in OSB_FIRE_LOCATION_TERMS)
        and any(term in t for term in OSB_FIRE_EVENT_TERMS)
    )


def classify(title,snippet,source_domain=''):
    t=norm(f'{title} {snippet}')
    neg=[x for x in NEGATIVE_TERMS if x in t]
    risk=[x for x in HIGH_RISK_TERMS if x in t]
    cat='Genel Sanayi / Teknoloji'
    for c,ks in CATEGORIES.items():
        if any(k in t for k in ks): cat=c; break

    # Daha açıklanabilir 0-100 risk skoru.
    score=8
    reasons=[]
    if neg:
        score += min(30, 8*len(neg))
        reasons.append(f'{len(neg)} negatif sinyal')
    if risk:
        score += min(38, 13*len(risk))
        reasons.append(f'{len(risk)} kritik risk sinyali')
    # Etki alanı ağırlıkları
    if any(x in t for x in ['savunma','aselsan','tusaş','tusas','roketsan','havelsan','baykar','iha','siha','füze','radar']):
        score += 10; reasons.append('savunma/kritik teknoloji')
    if any(x in t for x in ['kritik altyapı','enerji şebeke','elektrik üretimi','doğalgaz','nükleer','çip','yarı iletken','siber','veri sızıntısı']):
        score += 12; reasons.append('kritik altyapı/teknoloji')
    if any(x in t for x in ['tedarik zinciri','tedarik krizi','çip krizi','lojistik','ambargo','yaptırım']):
        score += 10; reasons.append('tedarik/jeopolitik etki')
    if any(x in t for x in ['iflas','konkordato','fabrika kapandı','üretim durdu','can kaybı','ölüm']):
        score += 18; reasons.append('operasyonel/finansal ağır sonuç')
    score=min(100,score)
    sentiment='Negatif' if neg else 'Nötr'
    if score>=70 or risk or len(neg)>=3: status='Yüksek Risk'
    elif score>=30 or neg: status='Negatif'
    else: status='Normal'
    if not reasons: reasons=['olumsuz risk sinyali tespit edilmedi']
    return sentiment,score,status,neg,risk,cat,reasons

def rss(query, timeout=7):
    try:
        r=requests.get('https://news.google.com/rss/search',params={'q':query,'hl':'tr','gl':'TR','ceid':'TR:tr'},headers=HEADERS,timeout=timeout)
        r.raise_for_status(); root=ET.fromstring(r.content); out=[]
        for it in root.findall('.//item'):
            src=it.find('source')
            out.append({
                'title':html.unescape(it.findtext('title') or ''),
                'url':it.findtext('link') or '',
                'date':it.findtext('pubDate') or '',
                'snippet':BeautifulSoup(it.findtext('description') or '','html.parser').get_text(' ',strip=True),
                'source':src.text if src is not None else '',
                'source_url':src.get('url','') if src is not None else ''
            })
        return out
    except Exception:
        return []

def ddgs_text(q):
    try:
        from ddgs import DDGS
    except Exception:
        try: from duckduckgo_search import DDGS
        except Exception: return []
    try:
        with DDGS() as d: return list(d.text(q,region='tr-tr',timelimit='d',max_results=40))
    except Exception: return []

def gdelt(q, timespan):
    try:
        r=requests.get('https://api.gdeltproject.org/api/v2/doc/doc',params={'query':q,'mode':'artlist','maxrecords':250,'format':'json','sort':'HybridRel','timespan':timespan},headers=HEADERS,timeout=8)
        r.raise_for_status(); return r.json().get('articles',[]) or []
    except Exception: return []

def period_window(hours):
    if hours<=3: return '6h'
    if hours<=24: return '1d'
    if hours<=48: return '2d'
    if hours<=168: return '7d'
    return '30d'

def _query_terms(user_query):
    parts=re.split(r'\bOR\b|,|;|\n',user_query or '',flags=re.I)
    out=[]; seen=set()
    for x in parts:
        x=x.strip().strip('"').strip("'")
        if len(x)>=3 and norm(x) not in seen:
            seen.add(norm(x)); out.append(x)
    return out

def build_turkish_queries(when, user_query=''):
    # Geniş arama evreni: tek dev sorgu yerine konu kümeleri paralel taranır.
    # Böylece kapsam genişlerken Google News sorguları aşırı ağırlaşmaz.
    groups=[
        '(sanayi OR imalat OR üretim OR fabrika OR tesis OR OSB OR "organize sanayi" OR endüstri)',
        '(makine OR otomasyon OR robotik OR "endüstri 4.0" OR kapasite OR "kapasite kullanım")',
        '(teknoloji OR inovasyon OR "Ar-Ge" OR Arge OR patent OR "dijital dönüşüm" OR teknopark)',
        '("yapay zeka" OR "yapay zekâ" OR "makine öğrenmesi" OR yazılım OR SaaS OR bulut)',
        '("siber güvenlik" OR "siber saldırı" OR "veri sızıntısı" OR kuantum OR blockchain OR fintech)',
        '(çip OR mikroçip OR "yarı iletken" OR semiconductor OR işlemci OR wafer OR elektronik OR PCB OR sensör)',
        '("savunma sanayii" OR "savunma sanayi" OR ASELSAN OR TUSAŞ OR ROKETSAN OR HAVELSAN OR Baykar OR Bayraktar)',
        '(İHA OR SİHA OR drone OR KAAN OR Kızılelma OR HİSAR OR SİPER OR füze OR roket OR radar OR "elektronik harp")',
        '(havacılık OR "havacılık sanayii" OR uçak OR helikopter OR uzay OR uydu OR "roket fırlatma" OR "Türkiye Uzay Ajansı")',
        '(otomotiv OR TOGG OR "elektrikli araç" OR "hibrit araç" OR "otonom araç" OR batarya OR şarj OR mobilite)',
        '(enerji OR "enerji depolama" OR "güneş enerjisi" OR "rüzgar enerjisi" OR hidrojen OR "yakıt hücresi" OR "nükleer enerji")',
        '(kimya OR petrokimya OR plastik OR polimer OR "demir çelik" OR çelik OR metal OR alüminyum OR bakır)',
        '(madencilik OR maden OR tekstil OR "gıda teknolojisi" OR "gıda sanayii" OR "tarım teknolojisi" OR seracılık)',
        '(lojistik OR "tedarik zinciri" OR tersane OR "gemi inşa" OR denizcilik OR demiryolu OR "raylı sistem")',
        '(biyoteknoloji OR biyomedikal OR nanoteknoloji OR "medikal cihaz" OR "sağlık teknolojisi" OR "ileri malzeme" OR kompozit)',
        '(TÜBİTAK OR KOSGEB OR "Sanayi ve Teknoloji Bakanlığı" OR TürkPatent OR TEKNOFEST OR "yatırım teşvik" OR "teknoloji transferi")',
        '(startup OR "start-up" OR girişimcilik OR "yatırım turu" OR "venture capital" OR "Ar-Ge merkezi" OR "tasarım merkezi")',
        '(ihracat OR ithalat OR "yüksek teknoloji" OR "orta yüksek teknoloji" OR "kritik teknoloji" OR "stratejik ürün" OR yerlileştirme)'
    ]
    qs=[f'Türkiye {g} when:{when}' for g in groups]
    # Kullanıcının kutuya eklediği ÖZEL terimler ayrıca taranır.
    # Performans: varsayılan geniş evrende zaten bulunan terimleri ikinci kez sorgulamayız.
    # Böylece normal kullanımda 38 civarı sorgu yerine yaklaşık 18 ana sorgu çalışır;
    # kullanıcı gerçekten yeni bir terim eklerse yalnızca o terim(ler) ek sorgu olur.
    built_in={norm(x) for x in TOPIC_TERMS}
    generic={'sanayi','teknoloji','üretim','imalat','fabrika','türkiye','türk'}
    custom=[
        x for x in _query_terms(user_query)
        if norm(x) not in generic and norm(x) not in built_in
    ]
    for term in custom[:8]:
        qs.append(f'Türkiye ("{term}") when:{when}')
    return qs

def build_negative_queries(when):
    return [
        f'Türkiye (iflas OR konkordato OR "üretim durdu" OR "fabrika kapandı" OR "işten çıkarma" OR grev OR soruşturma OR dava OR ceza OR "geri çağırma" OR "siber saldırı" OR "veri sızıntısı" OR yaptırım OR ambargo OR "ihale iptal" OR ertelendi OR gecikme OR "tedarik krizi" OR daralma OR zafiyet OR usulsüzlük OR yolsuzluk) (sanayi OR teknoloji OR üretim OR fabrika OR savunma OR otomotiv OR enerji OR şirket OR tesis OR proje) when:{when}',
        f'Türkiye (OSB OR "organize sanayi" OR "organize sanayi bölgesi") (yangın OR yangını OR alev OR "fabrika yangını" OR "tesis yangını") when:{when}'
    ]

def build_greek_queries(when):
    site='('+' OR '.join('site:'+x for x in GR)+')'
    return [
        f'(Turkey OR Türkiye OR Turkish OR Τουρκία OR τουρκική) (defense OR defence OR savunma OR άμυνα OR arms) {site} when:{when}',
        f'(Baykar OR Bayraktar OR ASELSAN OR TUSAŞ OR Roketsan OR HAVELSAN OR KAAN OR Kızılelma OR SİPER OR HİSAR) {site} when:{when}',
        f'(Turkey OR Turkish OR Τουρκία) (drone OR UAV OR missile OR fighter OR frigate OR submarine OR defense industry) {site} when:{when}'
    ]

def build_social_queries(when):
    site='('+' OR '.join('site:'+x for x in SOCIAL)+')'
    return [
        f'(Türkiye OR Türk) (sanayi OR teknoloji OR üretim OR savunma OR yapay zeka OR siber) {site} when:{when}',
        f'(ASELSAN OR TUSAŞ OR ROKETSAN OR HAVELSAN OR Baykar OR TOGG OR TÜBİTAK) {site} when:{when}',
        f'(iflas OR üretim durdu OR fabrika kapandı OR soruşturma OR siber saldırı OR yaptırım) (sanayi OR teknoloji OR savunma) {site} when:{when}'
    ]

def normalize_rows(raw, cutoff, mode, user_query):
    out=[]; reasons={'zaman':0,'konu':0,'kaynak':0,'yunan':0,'gecersiz':0}
    for r in raw:
        url=(r.get('url') or r.get('link') or '').strip(); title=html.unescape((r.get('title') or '').strip())
        if not url or not title: reasons['gecersiz']+=1; continue
        dt=parse_dt(r.get('date') or r.get('publishedAt') or r.get('seendate'))
        # Tüm tarihleri UTC-aware datetime olarak karşılaştır. Bazı RSS/arama
        # sağlayıcıları timezone bilgisi olmadan tarih döndürebildiği için
        # doğrudan datetime karşılaştırması TypeError üretebilir.
        if dt:
            try:
                if dt.tzinfo is None:
                    dt=dt.replace(tzinfo=timezone.utc)
                else:
                    dt=dt.astimezone(timezone.utc)
                cutoff_utc = cutoff if cutoff.tzinfo is not None else cutoff.replace(tzinfo=timezone.utc)
                cutoff_utc = cutoff_utc.astimezone(timezone.utc)
                if dt < cutoff_utc:
                    reasons['zaman']+=1
                    continue
            except (TypeError, ValueError, AttributeError):
                # Tarih karşılaştırılamıyorsa haberi düşürme; aşağıda
                # bilinmeyen tarih olarak sıralanmasına izin ver.
                dt=None
        if not dt and mode=='turkish':
            # tarih yoksa hızlı bakışta atmayalım; sadece sıralamada alta al.
            pass
        snippet=html.unescape((r.get('snippet') or r.get('body') or r.get('description') or '').strip())
        src=r.get('source') or ''
        d=infer_source(src,r.get('source_url',''),url)
        t=f'{title} {snippet}'
        if mode=='greek':
            if d not in GR or not greek_defense(t): reasons['yunan']+=1; continue
        elif mode=='social':
            if d not in SOCIAL: reasons['kaynak']+=1; continue
            if not relevant(t,user_query): reasons['konu']+=1; continue
        elif mode=='global':
            if not relevant(t,user_query): reasons['konu']+=1; continue
        else:
            # Türk batch'inde kaynak filtresi YOK. Arama zaten Türkiye odaklı.
            # Bu, Google News'in yayıncı URL'sini Google domaininde tuttuğu durumlarda
            # Türk haberlerinin 0'a düşmesini engeller. Türk kaynakları sıralamada öne çıkar.
            if not relevant(t,user_query): reasons['konu']+=1; continue
        sentiment,score,status,neg,risk,cat,risk_reasons=classify(title,snippet,d)
        out.append({
            'Tarih_dt':dt,'Tarih':fmt_dt(dt),'Başlık':title,'İçerik_Özeti':snippet or title,
            'URL':url,'RSS_URL':url,'Kaynak':(src if norm(src) not in {'google haberler','google news','google'} else (d or src or 'Açık Kaynak')),
            'Yayıncı_URL':(r.get('source_url') or '').strip(),'Yayıncı':src or d or 'Açık Kaynak',
            'Domain':d,'Kaynak_Grubu':source_group(d),
            'Kategori':cat,'Duygu':sentiment,'Skor':score,'Risk_Skoru':score,'Risk_Durumu':status,
            'Risk_Gerekçesi':'; '.join(risk_reasons),'Negatif_Sinyaller':neg,'Risk_Sinyalleri':risk,
            'Seç':False,'Görsel_URL':'','_mode':mode
        })
    return out,reasons


def source_reliability(domain_name, source_name=''):
    d=domain(domain_name); n=norm(source_name)
    if d in TR_OFFICIAL: return '🟢 A — Birincil / resmî'
    if d in TR_MAIN or d in TR_TECH: return '🟢 A — Güvenilir medya'
    if d in GR: return '🔵 B — Yunan medya'
    if d in SOCIAL: return '🟠 C — Sosyal / indeks'
    return '🟡 B — Açık kaynak'



def dedupe(rows):
    """URL ve başlık anahtarına göre hızlı tekilleştirme; kronolojik sıralamayı korur."""
    out=[]
    urls=set()
    titles=set()
    for r in rows:
        u=str(r.get('URL','') or '')
        k=title_key(str(r.get('Başlık','') or ''))
        if u and u in urls:
            continue
        if k and k in titles:
            continue
        if u:
            urls.add(u)
        if k:
            titles.add(k)
        out.append(r)

    out.sort(
        key=lambda x:(
            x.get('Tarih_dt') is not None,
            _to_utc_datetime(x.get('Tarih_dt')) or datetime.min.replace(tzinfo=timezone.utc),
            source_rank(x.get('Domain',''))
        ),
        reverse=True
    )
    return out


def _title_tokens(text):
    """Başlıktan olay eşleştirmesi için anlamlı token kümesi üretir."""
    txt=norm(text)
    txt=re.sub(r'[^\wçğıöşüÇĞİÖŞÜ]+',' ',txt)
    stop={
        've','ile','bir','bu','da','de','için','son','yeni','türkiye','türk','haberi','haber',
        'açıklama','dedi','oldu','olarak','olan','milyon','milyar','bin','yüzde','ile ilgili'
    }
    return {x for x in txt.split() if len(x)>=3 and x not in stop}


def _event_signature(title):
    """
    Aynı/çok benzer haber başlıklarını hızlı gruplamak için deterministik imza.
    İlk 6 ayırt edici token kullanılır. O(n²) SequenceMatcher taraması yerine
    ters indeks kullanacağımız için yüzlerce haberde çok daha hızlıdır.
    """
    toks=sorted(_title_tokens(title))
    return ' '.join(toks[:6])


def _jaccard(a,b):
    if not a or not b:
        return 0.0
    inter=len(a & b)
    union=len(a | b)
    return inter/union if union else 0.0


def source_reliability(source_domain,source_name=''):
    d=domain(source_domain); n=norm(source_name)
    if d in TR_OFFICIAL: return '🟢 A — Birincil / resmî'
    if d in TR_MAIN or d in TR_TECH: return '🟢 A — Güvenilir medya'
    if d in GR: return '🔵 B — Yunan medya'
    if d in SOCIAL: return '🟠 C — Sosyal / indeks'
    return '🟡 B — Açık kaynak'


def enrich_rows(rows):
    """
    HIZLI analitik katman.
    Önceki sürümde her haber diğer bütün haberlerle SequenceMatcher üzerinden
    karşılaştırılıyordu ve doğrulama için ikinci kez O(n²) tarama yapılıyordu.
    Bu sürüm ters token indeksi + olay grubu istatistikleri kullanır.
    """
    if not rows:
        return rows

    # 1) Tarih + temel risk sınıflaması: O(n)
    for r in rows:
        r['Tarih_dt']=_to_utc_datetime(r.get('Tarih_dt'))
        sentiment,score,status,neg,risk,cat,reasons=classify(
            r.get('Başlık',''),r.get('İçerik_Özeti',''),r.get('Domain','')
        )
        r['Duygu']=sentiment
        r['Risk_Skoru']=score
        r['Risk_Durumu']=status
        r['Negatif_Sinyaller']=neg
        r['Risk_Sinyalleri']=risk
        r['Risk_Gerekçesi']='; '.join(reasons)
        r['Kaynak_Güvenilirliği']=source_reliability(r.get('Domain',''),r.get('Kaynak',''))
        r['_tokens']=_title_tokens(r.get('Başlık',''))

    # 2) Olay kümelemesi: ters token indeksi.
    # Her haber yalnızca ortak token taşıyan sınırlı sayıdaki önceki adayla karşılaştırılır.
    token_index={}
    event_representative={}
    next_event=1

    for idx,r in enumerate(rows):
        toks=r['_tokens']
        candidate_events=set()
        for tok in toks:
            candidate_events.update(token_index.get(tok,set()))

        best_event=None
        best_score=0.0
        for eid in candidate_events:
            rep_tokens=event_representative[eid]
            score=_jaccard(toks,rep_tokens)
            if score>best_score:
                best_score=score
                best_event=eid

        # Aynı olay için Jaccard eşiği. Çok kısa başlıklarda biraz daha sıkı.
        threshold=0.48 if len(toks)>=6 else 0.58
        if best_event is None or best_score < threshold:
            best_event=f'OLAY-{next_event:03d}'
            next_event+=1
            event_representative[best_event]=set(toks)

        r['Olay_ID']=best_event
        for tok in toks:
            token_index.setdefault(tok,set()).add(best_event)

    # 3) Olay istatistikleri bir kez hesaplanır: O(n)
    groups={}
    for r in rows:
        groups.setdefault(r['Olay_ID'],[]).append(r)

    event_meta={}
    for eid,g in groups.items():
        domains={x.get('Domain') for x in g if x.get('Domain')}
        times=[x.get('Tarih_dt') for x in g if x.get('Tarih_dt') is not None]
        official=any(domain(x.get('Domain','')) in TR_OFFICIAL for x in g)
        social_only=all(domain(x.get('Domain','')) in SOCIAL for x in g if x.get('Domain')) if domains else False

        if official:
            verification='🟢 Resmî açıklama / birincil kaynak'
        elif len(domains)>=2 or len(g)>=3:
            verification='🟢 Çoklu kaynakla destekleniyor'
        elif social_only:
            verification='🟠 Sosyal medya / tek kaynak'
        elif any(domain(x.get('Domain','')) in TR_MAIN+TR_TECH+GR for x in g):
            verification='🟡 Tek medya kaynağı'
        else:
            verification='🟡 Tek/açık kaynak'

        event_meta[eid]={
            'sources':len(domains),
            'first':fmt_dt(min(times)) if times else '',
            'last':fmt_dt(max(times)) if times else '',
            'verification':verification
        }

    for r in rows:
        meta=event_meta[r['Olay_ID']]
        r['Olay_Kaynak_Sayisi']=meta['sources']
        r['Olay_İlk_Görülme']=meta['first'] or r.get('Tarih','')
        r['Olay_Son_Görülme']=meta['last'] or r.get('Tarih','')
        r['Doğrulama']=meta['verification']
        r.pop('_tokens',None)

    return rows

def build_event_summary(df):
    if df.empty: return pd.DataFrame()
    items=[]
    for oid,g in df.groupby('Olay_ID',dropna=False):
        g=g.sort_values('Tarih_dt',ascending=False)
        head=str(g.iloc[0].get('Başlık',''))
        risk=int(g['Risk_Skoru'].max())
        cat=str(g.iloc[0].get('Kategori',''))
        sources=', '.join(dict.fromkeys(str(x) for x in g['Kaynak'].tolist()))
        items.append({'Olay_ID':oid,'Öne Çıkan Başlık':head,'Kategori':cat,'Haber Sayısı':len(g),'Kaynak Sayısı':g['Domain'].nunique(),'Risk':risk,'İlk Görülme':g['Olay_İlk_Görülme'].min(),'Son Görülme':g['Olay_Son_Görülme'].max(),'Kaynaklar':sources})
    return pd.DataFrame(items).sort_values(['Risk','Son Görülme'],ascending=[False,False])

def trend_table(df):
    if df.empty: return pd.DataFrame()
    x=df.copy(); x['Saat']=x['Tarih_dt'].apply(lambda d: d.strftime('%Y-%m-%d %H:00') if d else 'Bilinmiyor')
    return x.groupby(['Kategori']).size().reset_index(name='Haber').sort_values('Haber',ascending=False)

def watchlist_hits(df, terms):
    terms=[norm(x) for x in re.split(r',|\n|;',terms or '') if len(norm(x))>=2]
    if df.empty or not terms: return pd.DataFrame()
    mask=df.apply(lambda r:any(t in norm(f"{r.get('Başlık','')} {r.get('İçerik_Özeti','')}") for t in terms),axis=1)
    return df[mask].copy()

def _clean_note_text(value):
    text=BeautifulSoup(str(value or ''),'html.parser').get_text(' ',strip=True)
    text=re.sub(r'\s+',' ',text).strip()
    return text

def _sentence_split_tr(text):
    text=_clean_note_text(text)
    if not text:
        return []
    parts=re.split(r'(?<=[.!?])\s+(?=[A-ZÇĞİÖŞÜ0-9“"])',text)
    return [p.strip() for p in parts if len(p.strip())>20]

def _unique_sentences(sentences):
    out=[]; seen=set()
    for s in sentences:
        k=norm(s)
        if not k or k in seen:
            continue
        seen.add(k); out.append(s)
    return out

def _note_source_sentence(r):
    title=_clean_note_text(r.get('Başlık',''))
    source=_clean_note_text(r.get('Kaynak','Açık Kaynak'))
    when=_clean_note_text(r.get('Tarih',''))
    cat=_clean_note_text(r.get('Kategori',''))
    if when:
        return f"{when} tarihinde {source} tarafından yayımlanan “{title}” başlıklı içerik, {cat.lower() if cat else 'sanayi ve teknoloji'} alanındaki gelişmelere ilişkindir."
    return f"{source} tarafından yayımlanan “{title}” başlıklı içerik, {cat.lower() if cat else 'sanayi ve teknoloji'} alanındaki gelişmelere ilişkindir."

def _detail_sentences(text, title=''):
    """Haber gövdesinden bilgi taşıyan cümleleri temizler; ayrıntıyı korur."""
    text=_clean_note_text(text)
    if not text:
        return []
    raw=_sentence_split_tr(text)
    title_n=norm(title)
    boiler=[
        'çerez','cookie','abonelik','abone ol','reklam','tüm hakları saklıdır',
        'gizlilik politikası','kullanım koşulları','google news','bildirimleri aç',
        'uygulamamızı indirin','facebook','instagram','twitter','whatsapp',
        'son dakika haberleri için','haberlerimizi takip'
    ]
    out=[]; seen=set()
    for s in raw:
        sn=norm(s)
        if len(s)<28 or sn==title_n or any(b in sn for b in boiler):
            continue
        key=' '.join(sn.split()[:16])
        if key in seen:
            continue
        seen.add(key)
        out.append(s.strip())
    return out


def _sent_score(s):
    """Bilgi yoğun cümlelere öncelik verir."""
    n=norm(s)
    score=0
    if re.search(r'\b\d+(?:[.,]\d+)?\b', s): score+=3
    if any(x in n for x in ['tarih','saat','yıl','ay','gün','bugün','dün']): score+=2
    if any(x in n for x in ['bakan','başkan','valilik','belediye','şirket','kurum','bakanlık','müdür','yetkili','açıkladı','bildirdi','belirtti']): score+=3
    if any(x in n for x in ['nedeni','sebebi','sonucu','sonuç','etki','hasar','zarar','yaralı','hayatını kaybetti','tahliye','müdahale','kontrol altına']): score+=3
    if any(x in n for x in ['üretim','kapasite','yatırım','ihracat','ithalat','tesis','fabrika','osb','teknoloji','savunma','enerji']): score+=2
    return score


def _join_sentences_naturally(sentences):
    """Kaynak cümlelerini bilgi kaybı olmadan okunabilir paragraf akışına getirir."""
    if not sentences:
        return ''
    out=[]
    for s in sentences:
        s=s.strip()
        if not s:
            continue
        if s[-1] not in '.!?':
            s+='.'
        out.append(s)
    return ' '.join(out)


def _compose_single_article_note(row, detail):
    """
    Tek haberi 'haber özeti' gibi değil, ayrıntılı bilgi notu gibi ele alır:
    konu/olay -> gelişmeler -> açıklamalar/veriler -> mevcut durum/sonuç.
    Ara başlık kullanmaz.
    """
    title=_clean_note_text(detail.get('title') or row.get('Başlık',''))
    source=_clean_note_text(detail.get('source') or row.get('Kaynak','Açık Kaynak'))
    published=_clean_note_text(detail.get('published') or row.get('Tarih',''))
    fulltext=_clean_note_text(detail.get('text') or row.get('İçerik_Özeti','') or title)
    sentences=_detail_sentences(fulltext,title)

    # Haber sırasını esas al. İlk cümleler olayın başlangıcını çoğunlukla verir.
    # Çok uzun haberlerde bilgi yoğun cümleleri de mutlaka koru.
    if len(sentences)>45:
        first=sentences[:22]
        rest=sentences[22:]
        important=sorted(enumerate(rest), key=lambda z:_sent_score(z[1]), reverse=True)[:18]
        important=[s for _,s in sorted(important,key=lambda z:z[0])]
        sentences=first+important

    intro=(
        f"{published} tarihinde {source} tarafından yayımlanan “{title}” başlıklı haberde, "
        if published else
        f"{source} tarafından yayımlanan “{title}” başlıklı haberde, "
    )

    if not sentences:
        fallback=_clean_note_text(row.get('İçerik_Özeti','') or title)
        return intro + (fallback[0].lower()+fallback[1:] if len(fallback)>1 else fallback)

    # İlk 1-2 cümle olayın girişini oluşturur; geri kalanı kronolojik/haber sırasıyla devam eder.
    first=sentences[:2]
    remaining=sentences[2:]
    opening=_join_sentences_naturally(first)
    if opening:
        opening=opening[0].lower()+opening[1:]
    para1=intro+opening

    # Uzun haberlerde okunabilirlik için doğal paragraf bölmeleri.
    chunks=[]
    chunk_size=7
    for i in range(0,len(remaining),chunk_size):
        part=remaining[i:i+chunk_size]
        txt=_join_sentences_naturally(part)
        if txt:
            chunks.append(txt)

    parts=[para1]+chunks

    # Son cümlede yalnızca kaynakta aktarılan çerçeveye dayan.
    last_context=sentences[-3:] if len(sentences)>=3 else sentences
    conclusion=(
        "Bu çerçevede, haberde aktarılan son durum itibarıyla "
        + _join_sentences_naturally(last_context)
    )
    # Son üç cümleyi gövdede zaten kullandığımız için birebir tekrar çok fazlaysa genel, temkinli kapanış kullan.
    if len(norm(conclusion))>900:
        conclusion="Bu çerçevede gelişmenin, haberde aktarılan mevcut durum ve ilgili kurumların sonraki açıklamaları doğrultusunda izlenmesi önem taşımaktadır."
    parts.append(conclusion)

    return '\n\n'.join(parts)


def _compose_prose_note(df):
    """
    Seçilen gerçek haber sayfalarının tam metninden ayrıntılı bilgi notu oluşturur.
    'Giriş/Gelişme/Sonuç' başlıkları yazılmaz; anlatı doğal olarak bu sırada ilerler.
    """
    if df is None or df.empty:
        return '', []

    x=df.copy()
    if 'Tarih_dt' in x.columns:
        x['Tarih_dt']=pd.to_datetime(x['Tarih_dt'],utc=True,errors='coerce')
        x=x.sort_values('Tarih_dt',ascending=True,na_position='last')

    enriched=[]
    for _,r in x.iterrows():
        row=r.to_dict()
        detail=article_detail(row)
        enriched.append((row,detail))

    if len(enriched)==1:
        row,detail=enriched[0]
        note=_compose_single_article_note(row,detail)
        return note,enriched

    # Çoklu haberde kısa bir doğal giriş, ardından her haber kronolojik sırada ayrıntılı işlenir.
    source_names=[]
    for row,detail in enriched:
        s=_clean_note_text(detail.get('source') or row.get('Kaynak',''))
        if s and s not in source_names:
            source_names.append(s)

    opening=(
        f"Seçilen {len(enriched)} açık kaynak haberi birlikte değerlendirildiğinde, konuya ilişkin gelişmeler "
        f"{len(source_names)} farklı kaynağın aktardığı bilgiler çerçevesinde kronolojik bir seyir göstermektedir. "
        f"Aşağıdaki anlatımda haberlerde yer alan olaylar, açıklamalar, kişi ve kurumlar, teknik ve sayısal veriler, "
        f"neden-sonuç ilişkileri ile bildirilen etkiler mümkün olduğunca ayrıntılı biçimde korunmuştur."
    )

    blocks=[opening]
    for row,detail in enriched:
        blocks.append(_compose_single_article_note(row,detail))

    blocks.append(
        "Mevcut açık kaynak bilgileri birlikte değerlendirildiğinde, gelişmenin bundan sonraki seyri bakımından "
        "ilgili kurum ve kuruluşların yeni açıklamalarının, resmî duyuruların ve farklı açık kaynaklardan gelecek "
        "teyitlerin izlenmesi önem taşımaktadır. Bu bilgi notunda kaynak haberlerde yer almayan bir husus olgu olarak eklenmemiştir."
    )
    return '\n\n'.join(blocks), enriched

def make_analyst_docx(df, title='BİLGİ NOTU'):
    doc=Document()
    sec=doc.sections[0]
    sec.top_margin=Cm(2); sec.bottom_margin=Cm(2)
    sec.left_margin=Cm(2.5); sec.right_margin=Cm(2.5)

    styles=doc.styles
    styles['Normal'].font.name='Times New Roman'
    styles['Normal'].font.size=Pt(11)

    p=doc.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run(title); r.bold=True; r.font.size=Pt(14)

    p=doc.add_paragraph()
    p.add_run('Tarih/Saat: ').bold=True
    p.add_run(datetime.now().astimezone().strftime('%d.%m.%Y %H:%M:%S'))

    # Bilgi notu oluşturulurken seçili haberlerin gerçek sayfaları derinlemesine alınır.
    note,enriched=_compose_prose_note(df)

    # Kesintisiz düz yazı: ara başlık kullanılmaz.
    for block in note.split('\n\n'):
        if not block.strip(): continue
        bp=doc.add_paragraph()
        bp.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        bp.paragraph_format.first_line_indent=Cm(1.25)
        bp.paragraph_format.line_spacing=1.15
        bp.paragraph_format.space_after=Pt(8)
        bp.add_run(block.strip())

    # Kaynaklar; bilgi notunun anlatı akışından sonra.
    doc.add_paragraph()
    hp=doc.add_paragraph()
    rr=hp.add_run('KAYNAKLAR'); rr.bold=True

    for i,(row,detail) in enumerate(enriched,1):
        title_text=_clean_note_text(detail.get('title') or row.get('Başlık',''))
        source=_clean_note_text(detail.get('source') or row.get('Kaynak','Açık Kaynak'))
        real_url=detail.get('canonical') or row.get('Yayıncı_URL') or row.get('URL','')

        p=doc.add_paragraph()
        p.paragraph_format.left_indent=Cm(.5)
        p.add_run(f"{i}. {source} — {title_text}")
        if real_url:
            p.add_run(' — ')
            _word_hyperlink(p,real_url,'Haber linki')

    bio=BytesIO()
    doc.save(bio); bio.seek(0)
    return bio.getvalue()



# -----------------------------
# V33 — BİLGİ NOTU ADAYLARI + DÜNDEN BERİ NE DEĞİŞTİ?
# V32 çekirdek tarama motoruna dokunmaz.
# -----------------------------
_HISTORY_DB = Path(__file__).resolve().with_name("sanayi_teknoloji_osint_history.db")

def _history_connect():
    conn=sqlite3.connect(str(_HISTORY_DB),timeout=8)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    return conn

def _init_history_db():
    try:
        with _history_connect() as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS scans(
                    scan_id INTEGER PRIMARY KEY AUTOINCREMENT,
                    scanned_at TEXT NOT NULL,
                    period_hours INTEGER,
                    total_news INTEGER,
                    total_events INTEGER
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS event_snapshots(
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    scan_id INTEGER NOT NULL,
                    event_id TEXT,
                    title TEXT,
                    source TEXT,
                    url TEXT,
                    category TEXT,
                    summary TEXT,
                    risk_score INTEGER,
                    risk_status TEXT,
                    sentiment TEXT,
                    verification TEXT,
                    source_count INTEGER,
                    event_first_seen TEXT,
                    event_last_seen TEXT,
                    tokens_json TEXT,
                    FOREIGN KEY(scan_id) REFERENCES scans(scan_id)
                )
            """)
            conn.execute("CREATE INDEX IF NOT EXISTS idx_event_snapshots_scan ON event_snapshots(scan_id)")
            conn.execute("""
                CREATE TABLE IF NOT EXISTS shift_marks(
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    marked_at TEXT NOT NULL,
                    scan_id INTEGER,
                    label TEXT
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS important_basket(
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    added_at TEXT NOT NULL,
                    news_time TEXT,
                    title TEXT NOT NULL,
                    source TEXT,
                    url TEXT,
                    category TEXT,
                    risk_score INTEGER,
                    risk_status TEXT,
                    summary TEXT,
                    UNIQUE(url,title)
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS osint_report_basket(
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    added_at TEXT NOT NULL,
                    news_time TEXT,
                    title TEXT NOT NULL,
                    source TEXT,
                    url TEXT,
                    category TEXT,
                    risk_score INTEGER,
                    risk_status TEXT,
                    summary TEXT,
                    UNIQUE(url,title)
                )
            """)
            conn.commit()
        return True
    except Exception:
        return False

def _history_tokens(text):
    try:
        toks=_title_tokens(text)
        return sorted(toks)
    except Exception:
        txt=norm(text)
        return sorted({x for x in re.split(r'\W+',txt) if len(x)>=3})

def _save_scan_history(rows, scanned_at, period_hours):
    """Her taramanın olay özetini yerel SQLite dosyasına kaydeder."""
    if not rows or not _init_history_db():
        return None
    try:
        dfh=pd.DataFrame(rows)
        events=int(dfh['Olay_ID'].nunique()) if 'Olay_ID' in dfh.columns else len(dfh)
        with _history_connect() as conn:
            cur=conn.execute(
                "INSERT INTO scans(scanned_at,period_hours,total_news,total_events) VALUES(?,?,?,?)",
                (scanned_at.isoformat(),int(period_hours),len(dfh),events)
            )
            scan_id=int(cur.lastrowid)

            if 'Olay_ID' in dfh.columns:
                groups=dfh.groupby('Olay_ID',dropna=False)
            else:
                groups=[(f'ROW-{i}',dfh.iloc[[i]]) for i in range(len(dfh))]

            rows_to_insert=[]
            for oid,g in groups:
                g=g.copy()
                if 'Tarih_dt' in g.columns:
                    g['Tarih_dt']=pd.to_datetime(g['Tarih_dt'],utc=True,errors='coerce')
                    g=g.sort_values('Tarih_dt',ascending=False,na_position='last')
                r=g.iloc[0]
                title=str(r.get('Başlık','') or '')
                summary=' '.join(
                    str(x) for x in g.get('İçerik_Özeti',pd.Series(dtype=str)).tolist()
                    if str(x).strip()
                )[:8000]
                domains=set(str(x) for x in g.get('Domain',pd.Series(dtype=str)).tolist() if str(x).strip())
                source_count=max(
                    len(domains),
                    int(r.get('Olay_Kaynak_Sayisi',0) or 0)
                )
                rows_to_insert.append((
                    scan_id,str(oid),title,str(r.get('Kaynak','') or ''),
                    str(r.get('URL','') or ''),str(r.get('Kategori','') or ''),
                    summary,int(g.get('Risk_Skoru',pd.Series([0])).max() or 0),
                    str(r.get('Risk_Durumu','') or ''),str(r.get('Duygu','') or ''),
                    str(r.get('Doğrulama','') or ''),source_count,
                    str(r.get('Olay_İlk_Görülme','') or ''),
                    str(r.get('Olay_Son_Görülme','') or ''),
                    json.dumps(_history_tokens(title),ensure_ascii=False)
                ))

            conn.executemany("""
                INSERT INTO event_snapshots(
                    scan_id,event_id,title,source,url,category,summary,risk_score,risk_status,
                    sentiment,verification,source_count,event_first_seen,event_last_seen,tokens_json
                ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """,rows_to_insert)
            conn.commit()
        return scan_id
    except Exception:
        return None

def _previous_scan_id(current_scan_id=None):
    """Öncelik: bugünden önceki en son tarama; yoksa mevcut taramadan önceki en son tarama."""
    if not _init_history_db():
        return None
    try:
        today=datetime.now().astimezone().date().isoformat()
        with _history_connect() as conn:
            if current_scan_id:
                row=conn.execute(
                    "SELECT scan_id FROM scans WHERE scan_id < ? AND substr(scanned_at,1,10) < ? ORDER BY scanned_at DESC LIMIT 1",
                    (int(current_scan_id),today)
                ).fetchone()
                if not row:
                    row=conn.execute(
                        "SELECT scan_id FROM scans WHERE scan_id < ? ORDER BY scanned_at DESC LIMIT 1",
                        (int(current_scan_id),)
                    ).fetchone()
            else:
                row=conn.execute(
                    "SELECT scan_id FROM scans WHERE substr(scanned_at,1,10) < ? ORDER BY scanned_at DESC LIMIT 1",
                    (today,)
                ).fetchone()
            return int(row[0]) if row else None
    except Exception:
        return None

def _load_scan_events(scan_id):
    if not scan_id:
        return pd.DataFrame()
    try:
        with _history_connect() as conn:
            return pd.read_sql_query(
                """SELECT e.*,s.scanned_at,s.period_hours
                   FROM event_snapshots e JOIN scans s ON e.scan_id=s.scan_id
                   WHERE e.scan_id=?""",
                conn,params=(int(scan_id),)
            )
    except Exception:
        return pd.DataFrame()

def _token_jaccard_lists(a,b):
    sa=set(a or []); sb=set(b or [])
    if not sa or not sb:
        return 0.0
    return len(sa&sb)/len(sa|sb)

def _verification_rank(text):
    t=norm(text)
    if 'resmi' in t or 'resmî' in t or 'birincil' in t: return 4
    if 'coklu kaynak' in t or 'çoklu kaynak' in t: return 3
    if 'tek medya' in t: return 2
    if 'sosyal medya' in t: return 1
    return 1

def _risk_rank(status):
    t=norm(status)
    if 'yuksek risk' in t or 'yüksek risk' in t: return 3
    if 'negatif' in t: return 2
    return 1

def _current_event_frame(df):
    if df is None or df.empty:
        return pd.DataFrame()
    items=[]
    group_col='Olay_ID' if 'Olay_ID' in df.columns else None
    groups=df.groupby(group_col,dropna=False) if group_col else [(f'ROW-{i}',df.iloc[[i]]) for i in range(len(df))]
    for oid,g in groups:
        g=g.copy()
        if 'Tarih_dt' in g.columns:
            g['Tarih_dt']=pd.to_datetime(g['Tarih_dt'],utc=True,errors='coerce')
            g=g.sort_values('Tarih_dt',ascending=False,na_position='last')
        r=g.iloc[0]
        summary=' '.join(str(x) for x in g.get('İçerik_Özeti',pd.Series(dtype=str)).tolist() if str(x).strip())[:8000]
        items.append({
            'event_id':str(oid),
            'title':str(r.get('Başlık','') or ''),
            'source':str(r.get('Kaynak','') or ''),
            'url':str(r.get('URL','') or ''),
            'category':str(r.get('Kategori','') or ''),
            'summary':summary,
            'risk_score':int(g.get('Risk_Skoru',pd.Series([0])).max() or 0),
            'risk_status':str(r.get('Risk_Durumu','') or ''),
            'sentiment':str(r.get('Duygu','') or ''),
            'verification':str(r.get('Doğrulama','') or ''),
            'source_count':max(
                int(r.get('Olay_Kaynak_Sayisi',0) or 0),
                len(set(str(x) for x in g.get('Domain',pd.Series(dtype=str)).tolist() if str(x).strip()))
            ),
            'event_first_seen':str(r.get('Olay_İlk_Görülme','') or ''),
            'event_last_seen':str(r.get('Olay_Son_Görülme','') or ''),
            'tokens':_history_tokens(str(r.get('Başlık','') or ''))
        })
    return pd.DataFrame(items)

def _compare_since_previous(df,current_scan_id=None):
    """
    Olay bazında:
    🆕 yeni olay
    🔄 yeni bilgi/güncelleme
    ⚠️ risk arttı
    ✅ teyit güçlendi
    """
    current=_current_event_frame(df)
    prev_id=_previous_scan_id(current_scan_id)
    previous=_load_scan_events(prev_id)
    if current.empty:
        return pd.DataFrame(),None,None
    if previous.empty:
        return pd.DataFrame(),prev_id,None

    prev_records=[]
    for _,p in previous.iterrows():
        try: toks=json.loads(p.get('tokens_json') or '[]')
        except Exception: toks=_history_tokens(p.get('title',''))
        rec=p.to_dict(); rec['tokens']=toks; prev_records.append(rec)

    changes=[]
    for _,c in current.iterrows():
        best=None; best_sim=0.0
        for p in prev_records:
            sim=_token_jaccard_lists(c['tokens'],p['tokens'])
            # Kaynak/URL aynıysa eşleşmeyi kuvvetlendir.
            if c.get('url') and c.get('url')==p.get('url'):
                sim=max(sim,0.95)
            if sim>best_sim:
                best_sim=sim; best=p

        if best is None or best_sim < 0.42:
            changes.append({
                'Değişim':'🆕 YENİ OLAY',
                'Başlık':c['title'],'Kaynak':c['source'],'Kategori':c['category'],
                'Risk':c['risk_score'],'Önceki Risk':'—','Kaynak Sayısı':c['source_count'],
                'Açıklama':'Önceki karşılaştırma taramasında benzer olay tespit edilmedi.',
                'URL':c['url'],'_priority':100+c['risk_score']
            })
            continue

        prev_risk=int(best.get('risk_score') or 0)
        risk_up=(c['risk_score']>=prev_risk+15) or (_risk_rank(c['risk_status'])>_risk_rank(best.get('risk_status','')))
        verify_up=_verification_rank(c['verification'])>_verification_rank(best.get('verification',''))
        sources_up=int(c['source_count'] or 0)>int(best.get('source_count') or 0)

        prev_tokens=set(_history_tokens((best.get('title') or '')+' '+(best.get('summary') or '')))
        cur_tokens=set(_history_tokens(c['title']+' '+c['summary']))
        new_tokens=cur_tokens-prev_tokens
        materially_updated=len(new_tokens)>=6 or sources_up

        if risk_up:
            kind='⚠️ RİSK ARTTI'
            expl=f"Risk {prev_risk}/100 seviyesinden {c['risk_score']}/100 seviyesine yükseldi."
            priority=95+c['risk_score']
        elif verify_up:
            kind='✅ TEYİT GÜÇLENDİ'
            expl=f"Doğrulama seviyesi “{best.get('verification','')}” düzeyinden “{c['verification']}” düzeyine yükseldi."
            priority=90+c['risk_score']
        elif materially_updated:
            kind='🔄 YENİ BİLGİ'
            bits=[]
            if sources_up:
                bits.append(f"kaynak sayısı {int(best.get('source_count') or 0)} → {c['source_count']}")
            if len(new_tokens)>=6:
                sample=', '.join(sorted(list(new_tokens))[:8])
                bits.append(f"yeni içerik unsurları: {sample}")
            expl='; '.join(bits) if bits else 'Olayla ilgili yeni ayrıntılar tespit edildi.'
            priority=80+c['risk_score']
        else:
            continue

        changes.append({
            'Değişim':kind,'Başlık':c['title'],'Kaynak':c['source'],'Kategori':c['category'],
            'Risk':c['risk_score'],'Önceki Risk':prev_risk,'Kaynak Sayısı':c['source_count'],
            'Açıklama':expl,'URL':c['url'],'_priority':priority
        })

    out=pd.DataFrame(changes)
    if not out.empty:
        out=out.sort_values(['_priority','Risk'],ascending=[False,False]).drop(columns=['_priority'])
    prev_time=str(previous.iloc[0].get('scanned_at','')) if not previous.empty else None
    return out,prev_id,prev_time

def _note_candidate_reason(r,change_kind=''):
    reasons=[]
    risk=int(r.get('risk_score',0) or 0)
    if risk>=70: reasons.append('yüksek risk')
    elif risk>=45: reasons.append('dikkat gerektiren risk')
    if str(r.get('sentiment',''))=='Negatif': reasons.append('negatif etki')
    if int(r.get('source_count',0) or 0)>=2: reasons.append('çoklu kaynak')
    vr=norm(r.get('verification',''))
    if 'resmi' in vr or 'resmî' in vr or 'birincil' in vr: reasons.append('birincil/resmî teyit')
    elif 'coklu kaynak' in vr or 'çoklu kaynak' in vr: reasons.append('teyit güçlendi')
    cat=norm(r.get('category',''))
    if any(k in cat for k in ['savunma','yarı iletken','dijital','enerji','sanayi']): reasons.append('stratejik sektör')
    if is_osb_fire(r.get('title',''),r.get('summary','')): reasons.append('OSB yangını/kritik üretim olayı')
    if change_kind:
        reasons.append(change_kind.replace('🆕','').replace('🔄','').replace('⚠️','').replace('✅','').strip().lower())
    return ', '.join(dict.fromkeys(reasons)) or 'güncel ve sektörel önem'

def _information_note_candidates(df,current_scan_id=None,limit=10):
    events=_current_event_frame(df)
    if events.empty:
        return pd.DataFrame()

    changes,_,_=_compare_since_previous(df,current_scan_id)
    change_map={}
    if not changes.empty:
        for _,c in changes.iterrows():
            change_map[c['Başlık']]=c['Değişim']

    rows=[]
    for _,r in events.iterrows():
        score=0
        risk=int(r['risk_score'] or 0)
        score += min(45,int(risk*0.45))
        if r['risk_status']=='Yüksek Risk': score+=18
        elif r['sentiment']=='Negatif': score+=10
        score += min(int(r['source_count'] or 0)*4,16)

        vr=_verification_rank(r['verification'])
        score += {4:12,3:9,2:4,1:0}.get(vr,0)

        title_summary=norm(r['title']+' '+r['summary'])
        if is_osb_fire(r['title'],r['summary']): score+=18
        if any(x in title_summary for x in ['savunma','aselsan','tusaş','tusas','roketsan','baykar','havelsan','füze','iha','siha']): score+=10
        if any(x in title_summary for x in ['yatırım','yeni tesis','kapasite art','ihracat','kritik teknoloji','yarı iletken','çip','nükleer']): score+=9
        if any(x in title_summary for x in ['üretim durdu','fabrika kapandı','yangın','patlama','siber saldırı','ambargo','yaptırım']): score+=12

        change_kind=change_map.get(r['title'],'')
        if change_kind:
            score += 14 if 'YENİ OLAY' in change_kind else 12

        score=min(100,score)
        rows.append({
            'Aday Puanı':score,
            'Başlık':r['title'],
            'Kaynak':r['source'],
            'Kategori':r['category'],
            'Risk':risk,
            'Kaynak Sayısı':r['source_count'],
            'Doğrulama':r['verification'],
            'Değişim':change_kind or '—',
            'Neden Bilgi Notu?':_note_candidate_reason(r,change_kind),
            'URL':r['url']
        })

    out=pd.DataFrame(rows).sort_values(['Aday Puanı','Risk'],ascending=[False,False]).head(limit)
    return out.reset_index(drop=True)


# -----------------------------
# V34 — VARDİYA BAŞLANGIÇ ÖZETİ + ÖNEMLİ GELİŞMELER SEPETİ
# V33 çekirdeğine dokunmaz.
# -----------------------------
def _mark_shift_handover(scan_id=None, label='Devir noktası'):
    if not _init_history_db():
        return False
    try:
        now=datetime.now().astimezone().isoformat()
        with _history_connect() as conn:
            conn.execute(
                "INSERT INTO shift_marks(marked_at,scan_id,label) VALUES(?,?,?)",
                (now,int(scan_id) if scan_id else None,label)
            )
            conn.commit()
        return True
    except Exception:
        return False

def _latest_shift_mark():
    if not _init_history_db():
        return None
    try:
        with _history_connect() as conn:
            row=conn.execute(
                "SELECT marked_at,scan_id,label FROM shift_marks ORDER BY marked_at DESC LIMIT 1"
            ).fetchone()
        return {'marked_at':row[0],'scan_id':row[1],'label':row[2]} if row else None
    except Exception:
        return None

def _shift_baseline(current_scan_id=None):
    """
    Öncelik manuel devir noktasıdır.
    Hiç devir noktası yoksa V33'ün önceki taramasını baseline olarak kullanır.
    """
    mark=_latest_shift_mark()
    if mark:
        try:
            return pd.to_datetime(mark['marked_at'],utc=True),f"Devir noktası: {mark['marked_at']}",mark.get('scan_id')
        except Exception:
            pass

    prev_id=_previous_scan_id(current_scan_id)
    prev=_load_scan_events(prev_id)
    if not prev.empty:
        try:
            ts=pd.to_datetime(str(prev.iloc[0].get('scanned_at','')),utc=True)
            return ts,f"Önceki tarama: {prev.iloc[0].get('scanned_at','')}",prev_id
        except Exception:
            pass
    return None,"Henüz devir noktası yok",None

def _shift_start_summary(df,current_scan_id=None):
    """
    Son devir noktasından bu yana:
    - yeni haber
    - yeni önemli olay
    - yüksek riskli gelişme
    - risk artışı
    - teyit güçlenmesi
    - OSB olayı
    - sabah ilk bakılması gereken 5 gelişme
    """
    if df is None or df.empty:
        return {},pd.DataFrame(),""

    baseline,baseline_label,baseline_scan_id=_shift_baseline(current_scan_id)
    x=df.copy()
    x['Tarih_dt']=pd.to_datetime(x.get('Tarih_dt'),utc=True,errors='coerce')

    if baseline is not None:
        since=x[(x['Tarih_dt'].isna()) | (x['Tarih_dt']>=baseline)].copy()
    else:
        since=x.copy()

    changes,_,_=_compare_since_previous(df,current_scan_id)
    if not changes.empty:
        new_events=int(changes['Değişim'].astype(str).str.contains('YENİ OLAY').sum())
        risk_up=int(changes['Değişim'].astype(str).str.contains('RİSK ARTTI').sum())
        verify_up=int(changes['Değişim'].astype(str).str.contains('TEYİT').sum())
    else:
        new_events=0; risk_up=0; verify_up=0

    high=int((since.get('Risk_Durumu',pd.Series(dtype=str))=='Yüksek Risk').sum()) if not since.empty else 0
    osb=0
    for _,r in since.iterrows():
        if is_osb_fire(r.get('Başlık',''),r.get('İçerik_Özeti','')):
            osb+=1

    top=_daily_top_events(since,5) if not since.empty else pd.DataFrame()

    stats={
        'new_news':len(since),
        'new_important_events':new_events,
        'high_risk':high,
        'risk_up':risk_up,
        'verify_up':verify_up,
        'osb':osb,
        'baseline_label':baseline_label
    }
    return stats,top,baseline_label

def _add_rows_to_important_basket(rows):
    if rows is None or len(rows)==0 or not _init_history_db():
        return 0
    added=0
    try:
        with _history_connect() as conn:
            for row in rows:
                title=str(row.get('Başlık','') or '').strip()
                url=str(row.get('URL','') or '').strip()
                if not title:
                    continue
                cur=conn.execute("""
                    INSERT OR IGNORE INTO important_basket(
                        added_at,news_time,title,source,url,category,risk_score,risk_status,summary
                    ) VALUES(?,?,?,?,?,?,?,?,?)
                """,(
                    datetime.now().astimezone().isoformat(),
                    str(row.get('Tarih','') or ''),
                    title,
                    str(row.get('Kaynak','') or ''),
                    url,
                    str(row.get('Kategori','') or ''),
                    int(row.get('Risk_Skoru',0) or 0),
                    str(row.get('Risk_Durumu','') or ''),
                    str(row.get('İçerik_Özeti','') or '')[:8000]
                ))
                if cur.rowcount:
                    added+=1
            conn.commit()
        return added
    except Exception:
        return 0

def _load_important_basket():
    if not _init_history_db():
        return pd.DataFrame()
    try:
        with _history_connect() as conn:
            return pd.read_sql_query(
                "SELECT * FROM important_basket ORDER BY added_at ASC,id ASC",
                conn
            )
    except Exception:
        return pd.DataFrame()


def _add_rows_to_osint_basket(rows):
    if rows is None or len(rows)==0 or not _init_history_db():
        return 0
    added=0
    try:
        with _history_connect() as conn:
            for row in rows:
                title=str(row.get('Başlık','') or '').strip()
                url=str(row.get('URL','') or '').strip()
                if not title:
                    continue
                cur=conn.execute("""
                    INSERT OR IGNORE INTO osint_report_basket(
                        added_at,news_time,title,source,url,category,risk_score,risk_status,summary
                    ) VALUES(?,?,?,?,?,?,?,?,?)
                """,(
                    datetime.now().astimezone().isoformat(),
                    str(row.get('Tarih','') or ''),
                    title,
                    str(row.get('Kaynak','') or ''),
                    url,
                    str(row.get('Kategori','') or ''),
                    int(row.get('Risk_Skoru',0) or 0),
                    str(row.get('Risk_Durumu','') or ''),
                    str(row.get('İçerik_Özeti','') or '')[:8000]
                ))
                if cur.rowcount:
                    added+=1
            conn.commit()
        return added
    except Exception:
        return 0

def _load_osint_basket():
    if not _init_history_db():
        return pd.DataFrame()
    try:
        with _history_connect() as conn:
            return pd.read_sql_query(
                "SELECT * FROM osint_report_basket ORDER BY added_at ASC,id ASC",
                conn
            )
    except Exception:
        return pd.DataFrame()

def _remove_osint_basket_ids(ids):
    ids=[int(x) for x in ids if str(x).isdigit()]
    if not ids:
        return 0
    try:
        with _history_connect() as conn:
            q="DELETE FROM osint_report_basket WHERE id IN (" + ",".join("?" for _ in ids) + ")"
            cur=conn.execute(q,ids)
            conn.commit()
            return cur.rowcount
    except Exception:
        return 0

def _clear_osint_basket():
    try:
        with _history_connect() as conn:
            cur=conn.execute("DELETE FROM osint_report_basket")
            conn.commit()
            return cur.rowcount
    except Exception:
        return 0

def _remove_basket_ids(ids):
    ids=[int(x) for x in ids if str(x).isdigit()]
    if not ids:
        return 0
    try:
        with _history_connect() as conn:
            q="DELETE FROM important_basket WHERE id IN (" + ",".join("?" for _ in ids) + ")"
            cur=conn.execute(q,ids)
            conn.commit()
            return cur.rowcount
    except Exception:
        return 0

def _clear_important_basket():
    try:
        with _history_connect() as conn:
            cur=conn.execute("DELETE FROM important_basket")
            conn.commit()
            return cur.rowcount
    except Exception:
        return 0

def make_important_basket_docx(basket_df):
    doc=Document()
    sec=doc.sections[0]
    sec.top_margin=Cm(2); sec.bottom_margin=Cm(2)
    sec.left_margin=Cm(2.5); sec.right_margin=Cm(2.5)
    doc.styles['Normal'].font.name='Times New Roman'
    doc.styles['Normal'].font.size=Pt(11)

    p=doc.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('24 SAATLİK ÖNEMLİ GELİŞMELER')
    r.bold=True; r.font.size=Pt(14)

    p=doc.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(datetime.now().astimezone().strftime('%d.%m.%Y %H:%M'))

    if basket_df is None or basket_df.empty:
        doc.add_paragraph('Sepette kayıtlı önemli gelişme bulunmamaktadır.')
    else:
        for i,(_,r) in enumerate(basket_df.iterrows(),1):
            p=doc.add_paragraph()
            p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent=Cm(.75)
            run=p.add_run(f"{i}. {r.get('news_time','')} tarihinde {r.get('source','Açık Kaynak')} tarafından yayımlanan “{r.get('title','')}” başlıklı gelişme")
            run.bold=False
            summary=_clean_note_text(r.get('summary',''))
            if summary:
                p.add_run(f"; {summary}")
            if r.get('risk_score') is not None:
                p.add_run(f" Risk puanı {int(r.get('risk_score') or 0)}/100 olarak hesaplanmıştır.")
            if r.get('url'):
                p.add_run(' (')
                _word_hyperlink(p,r.get('url'),'Haber linki')
                p.add_run(')')

    bio=BytesIO()
    doc.save(bio); bio.seek(0)
    return bio.getvalue()

# -----------------------------
# GÜNLÜK DURUM ÖZETİ — V32 EK MODÜL
# V31 çekirdek tarama / risk / alarm / bilgi notu fonksiyonlarına dokunmaz.
# -----------------------------
def _daily_summary_stats(df):
    x=df.copy()
    if x.empty:
        return {}

    neg=int((x['Duygu']=='Negatif').sum()) if 'Duygu' in x else 0
    high=int((x['Risk_Durumu']=='Yüksek Risk').sum()) if 'Risk_Durumu' in x else 0

    osb=0
    if 'Başlık' in x:
        for _,r in x.iterrows():
            if is_osb_fire(r.get('Başlık',''),r.get('İçerik_Özeti','')):
                osb+=1

    def count_terms(terms):
        c=0
        for _,r in x.iterrows():
            text=norm(f"{r.get('Başlık','')} {r.get('İçerik_Özeti','')} {r.get('Kategori','')}")
            if any(t in text for t in terms):
                c+=1
        return c

    investment=count_terms(['yatırım','yatirim','fabrika aç','tesis aç','kapasite art','yeni tesis','teşvik','tesvik'])
    defence=count_terms(['savunma','aselsan','tusaş','tusas','roketsan','baykar','havelsan','saha expo','iha','siha','füze','fuze'])
    cyber=count_terms(['siber','veri sızınt','veri sizint','fidye yazılım','fidye yazilim','hack','siber saldır','siber saldir'])

    return {
        'total':len(x),
        'negative':neg,
        'high_risk':high,
        'osb_fire':osb,
        'investment':investment,
        'defence':defence,
        'cyber':cyber
    }


def _daily_top_events(df, n=5):
    """Risk + negatiflik + kaynak teyidi + güncellik ile günün önemli olaylarını seçer."""
    if df.empty:
        return df.copy()

    x=df.copy()
    x['Tarih_dt']=pd.to_datetime(x.get('Tarih_dt'),utc=True,errors='coerce')

    def importance(r):
        score=int(r.get('Risk_Skoru',0) or 0)
        if r.get('Duygu')=='Negatif': score+=15
        if r.get('Risk_Durumu')=='Yüksek Risk': score+=25
        if is_osb_fire(r.get('Başlık',''),r.get('İçerik_Özeti','')): score+=35
        try:
            score+=min(int(r.get('Olay_Kaynak_Sayisi',0) or 0)*4,20)
        except Exception:
            pass
        if 'resmî' in norm(r.get('Doğrulama','')) or 'çoklu kaynak' in norm(r.get('Doğrulama','')):
            score+=12
        return score

    x['_Önem']=x.apply(importance,axis=1)

    # Aynı olayın beş kez listeye girmesini engelle.
    if 'Olay_ID' in x.columns:
        x=x.sort_values(['_Önem','Tarih_dt'],ascending=[False,False],na_position='last')
        x=x.drop_duplicates(subset=['Olay_ID'],keep='first')
    else:
        x=x.sort_values(['_Önem','Tarih_dt'],ascending=[False,False],na_position='last')

    return x.head(n).drop(columns=['_Önem'],errors='ignore')


def _daily_summary_text(df):
    stats=_daily_summary_stats(df)
    top=_daily_top_events(df,5)
    if not stats:
        return '',top,stats

    intro=(
        f"Sanayi ve teknoloji alanında gerçekleştirilen güncel açık kaynak taramasında toplam {stats['total']} haber tespit edilmiştir. "
        f"Bunların {stats['negative']} adedi negatif içerik, {stats['high_risk']} adedi yüksek riskli gelişme olarak sınıflandırılmıştır. "
        f"Tarama kapsamında {stats['osb_fire']} organize sanayi bölgesi yangını, {stats['investment']} yatırım/kapasite gelişmesi, "
        f"{stats['defence']} savunma sanayii bağlantılı içerik ve {stats['cyber']} siber güvenlik bağlantılı içerik tespit edilmiştir."
    )

    paras=[intro]
    if not top.empty:
        paras.append(
            "Günün açık kaynak görünümünde risk seviyesi, negatif etki, kaynak teyidi ve olayın güncelliği birlikte dikkate alındığında "
            "öne çıkan gelişmeler aşağıdaki şekilde değerlendirilmektedir."
        )
        for i,(_,r) in enumerate(top.iterrows(),1):
            title=_clean_note_text(r.get('Başlık',''))
            source=_clean_note_text(r.get('Kaynak','Açık Kaynak'))
            when=_clean_note_text(r.get('Tarih',''))
            content=_clean_note_text(r.get('İçerik_Özeti',''))
            sentences=_detail_sentences(content,title)
            detail=_join_sentences_naturally(sentences[:5]) if sentences else content
            p=f"{i}. sırada {when} tarihinde {source} tarafından yayımlanan “{title}” başlıklı gelişme öne çıkmaktadır."
            if detail:
                p+=" "+detail
            risk=int(r.get('Risk_Skoru',0) or 0)
            if risk:
                p+=f" Gelişmenin sistem risk puanı {risk}/100 olarak hesaplanmıştır."
            paras.append(p)

    paras.append(
        "Genel görünüm itibarıyla, yüksek riskli ve negatif gelişmeler ile kritik üretim altyapısını etkileyebilecek olayların "
        "takibinin sürdürülmesi; özellikle yeni resmî açıklamalar, olayların üretim ve tedarik zincirine etkileri ile farklı "
        "açık kaynaklardan gelecek teyitlerin izlenmesi önem taşımaktadır."
    )
    return '\n\n'.join(paras),top,stats


def make_daily_summary_docx(df):
    text,top,stats=_daily_summary_text(df)

    doc=Document()
    sec=doc.sections[0]
    sec.top_margin=Cm(2); sec.bottom_margin=Cm(2)
    sec.left_margin=Cm(2.5); sec.right_margin=Cm(2.5)
    doc.styles['Normal'].font.name='Times New Roman'
    doc.styles['Normal'].font.size=Pt(11)

    p=doc.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('GÜNLÜK SANAYİ VE TEKNOLOJİ DURUM ÖZETİ')
    r.bold=True; r.font.size=Pt(14)

    p=doc.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(datetime.now().astimezone().strftime('%d.%m.%Y %H:%M'))

    for block in text.split('\n\n'):
        bp=doc.add_paragraph()
        bp.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        bp.paragraph_format.first_line_indent=Cm(1.25)
        bp.paragraph_format.line_spacing=1.15
        bp.paragraph_format.space_after=Pt(8)
        bp.add_run(block)

    if not top.empty:
        hp=doc.add_paragraph()
        rr=hp.add_run('ÖNE ÇIKAN GELİŞMELERİN KAYNAKLARI')
        rr.bold=True
        for i,(_,row) in enumerate(top.iterrows(),1):
            p=doc.add_paragraph()
            p.add_run(f"{i}. {_clean_note_text(row.get('Kaynak','Açık Kaynak'))} — {_clean_note_text(row.get('Başlık',''))}")
            if row.get('URL'):
                p.add_run(' — ')
                _word_hyperlink(p,row['URL'],'Haber linki')

    bio=BytesIO()
    doc.save(bio); bio.seek(0)
    return bio.getvalue()

# -----------------------------
# DOCX — AKT / Açık Kaynak Taraması formatı
# Tarama motoru korunur. Yalnızca seçilen haberlerin rapora aktarılması değiştirilmiştir.
# -----------------------------
@st.cache_data(ttl=1800, show_spinner=False)
def article_detail(row):
    """
    Seçilen kayıt için gerçek yayıncı URL'sini ve gerçek haber sayfasını bulur.
    Google News'in kodlanmış RSS bağlantıları doğrudan yayıncı adresi değilse
    sırasıyla decoder, HTTP redirect, GDELT ve DuckDuckGo üzerinden çözülür.
    """
    if isinstance(row, str):
        row = {"URL": row}

    original_url = str(row.get("URL") or "").strip()
    fallback_title = str(row.get("Başlık") or "").strip()
    fallback_snippet = str(row.get("İçerik_Özeti") or "").strip()
    publisher_url = str(row.get("Yayıncı_URL") or "").strip()
    publisher_name = str(row.get("Yayıncı") or row.get("Kaynak") or "").strip()

    out = {
        "title": fallback_title,
        "canonical": original_url,
        "published": str(row.get("Tarih") or ""),
        "text": fallback_snippet,
        "images": [],
        "source": publisher_name,
    }

    def is_google(u):
        try:
            h = urlparse(u).netloc.lower()
            return h == "news.google.com" or h.endswith(".google.com")
        except Exception:
            return False

    def valid_article_url(u):
        if not u or not u.startswith("http"):
            return False
        h = urlparse(u).netloc.lower()
        return h not in {"news.google.com", "www.google.com", "google.com"} and "google.com" not in h

    def fetch_page(u):
        try:
            rr = requests.get(
                u,
                headers={
                    **HEADERS,
                    "Accept-Language": "tr-TR,tr;q=0.9,en;q=0.7",
                    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                },
                timeout=12,
                allow_redirects=True,
            )
            if rr.status_code >= 400 or not rr.text:
                return None, None
            return rr, BeautifulSoup(rr.text, "html.parser")
        except Exception:
            return None, None

    def decode_with_package(u):
        try:
            from googlenewsdecoder import gnewsdecoder
            result = gnewsdecoder(u, interval=0.2)
            if isinstance(result, dict) and result.get("status"):
                decoded = result.get("decoded_url")
                if valid_article_url(decoded):
                    return decoded
        except Exception:
            pass
        return ""

    def decode_with_http(u):
        rr, soup = fetch_page(u)
        if rr and valid_article_url(rr.url):
            return rr.url

        if soup:
            for attrs in (
                {"property": "og:url"},
                {"name": "twitter:url"},
            ):
                tag = soup.find("meta", attrs=attrs)
                if tag and valid_article_url(tag.get("content", "")):
                    return requests.compat.urljoin(rr.url, tag["content"])

            tag = soup.find("link", rel=lambda x: x and "canonical" in str(x).lower())
            if tag and valid_article_url(requests.compat.urljoin(rr.url, tag.get("href", ""))):
                return requests.compat.urljoin(rr.url, tag.get("href"))

        return ""

    def decode_with_search(title):
        if not title:
            return ""

        # Önce GDELT: sonuçlar doğrudan yayıncı URL'si verir.
        try:
            q = '"' + title.replace('"', " ")[:240] + '"'
            r = requests.get(
                "https://api.gdeltproject.org/api/v2/doc/doc",
                params={
                    "query": q,
                    "mode": "artlist",
                    "maxrecords": 20,
                    "format": "json",
                    "sort": "HybridRel",
                    "timespan": "30d",
                },
                headers=HEADERS,
                timeout=8,
            )
            if r.ok:
                arts = r.json().get("articles", []) or []
                target = norm(title)
                for art in arts:
                    u = art.get("url") or ""
                    t = norm(art.get("title") or "")
                    if valid_article_url(u):
                        # Exact/near exact başlık eşleşmesi öncelikli.
                        if target and (target in t or t in target):
                            return u
                for art in arts:
                    u = art.get("url") or ""
                    if valid_article_url(u):
                        return u
        except Exception:
            pass

        # Son fallback: DuckDuckGo doğrudan yayıncı URL'si döndürebilir.
        try:
            from ddgs import DDGS
        except Exception:
            try:
                from duckduckgo_search import DDGS
            except Exception:
                DDGS = None

        if DDGS:
            try:
                with DDGS() as d:
                    results = list(d.text(f'"{title}"', region="tr-tr", timelimit="m", max_results=8))
                target = norm(title)
                for item in results:
                    u = item.get("href") or item.get("url") or ""
                    t = norm(item.get("title") or "")
                    if valid_article_url(u) and target and (target in t or t in target):
                        return u
                for item in results:
                    u = item.get("href") or item.get("url") or ""
                    if valid_article_url(u):
                        return u
            except Exception:
                pass

        return ""

    # 1) Google News bağlantısını çöz.
    real_url = ""
    if is_google(original_url):
        real_url = decode_with_package(original_url)
        if not real_url:
            real_url = decode_with_http(original_url)
        if not real_url:
            real_url = decode_with_search(fallback_title)
    elif valid_article_url(original_url):
        real_url = original_url
    else:
        real_url = decode_with_search(fallback_title)

    # 2) Gerçek sayfayı indir.
    rr, soup = fetch_page(real_url) if real_url else (None, None)

    if rr and soup:
        out["canonical"] = real_url or rr.url

        # Canonical
        can = soup.find("link", rel=lambda x: x and "canonical" in str(x).lower())
        if can and can.get("href"):
            out["canonical"] = requests.compat.urljoin(rr.url, can["href"])
        else:
            ogurl = soup.find("meta", attrs={"property": "og:url"})
            if ogurl and ogurl.get("content"):
                out["canonical"] = requests.compat.urljoin(rr.url, ogurl["content"])

        # Başlık
        for attrs in (
            {"property": "og:title"},
            {"name": "twitter:title"},
        ):
            t = soup.find("meta", attrs=attrs)
            if t and t.get("content"):
                out["title"] = t["content"].strip()
                break
        if not out["title"] and soup.title:
            out["title"] = soup.title.get_text(" ", strip=True)

        # Yayıncı
        for attrs in (
            {"property": "og:site_name"},
            {"name": "application-name"},
        ):
            t = soup.find("meta", attrs=attrs)
            if t and t.get("content"):
                out["source"] = t["content"].strip()
                break

        # Tarih
        for attrs in (
            {"property": "article:published_time"},
            {"itemprop": "datePublished"},
            {"name": "date"},
            {"name": "pubdate"},
        ):
            t = soup.find("meta", attrs=attrs)
            if t and t.get("content"):
                out["published"] = t["content"].strip()
                break

        bodies = []
        images = []

        def walk_json(obj):
            if isinstance(obj, dict):
                typ = str(obj.get("@type", "")).lower()
                if "article" in typ or "news" in typ:
                    if obj.get("headline"):
                        out["title"] = str(obj["headline"])
                    if obj.get("datePublished"):
                        out["published"] = str(obj["datePublished"])
                    if obj.get("articleBody"):
                        bodies.append(str(obj["articleBody"]))
                    pub = obj.get("publisher")
                    if isinstance(pub, dict) and pub.get("name"):
                        out["source"] = str(pub["name"])
                    im = obj.get("image") or obj.get("thumbnailUrl")
                    if isinstance(im, str):
                        images.append(im)
                    elif isinstance(im, list):
                        for x in im:
                            if isinstance(x, str):
                                images.append(x)
                            elif isinstance(x, dict) and x.get("url"):
                                images.append(str(x["url"]))
                    elif isinstance(im, dict) and im.get("url"):
                        images.append(str(im["url"]))
                for v in obj.values():
                    walk_json(v)
            elif isinstance(obj, list):
                for x in obj:
                    walk_json(x)

        for tag in soup.find_all("script", attrs={"type": re.compile(r"application/ld\+json", re.I)}):
            try:
                raw = tag.string or tag.get_text()
                if raw:
                    walk_json(json.loads(raw))
            except Exception:
                pass

        for attrs in (
            {"property": "og:image"},
            {"property": "og:image:url"},
            {"name": "twitter:image"},
            {"name": "twitter:image:src"},
        ):
            t = soup.find("meta", attrs=attrs)
            if t and t.get("content"):
                images.append(requests.compat.urljoin(rr.url, t["content"]))

        selectors = [
            '[itemprop="articleBody"]',
            "article",
            '[class*="article-body"]',
            '[class*="article-content"]',
            '[class*="news-content"]',
            '[class*="news-detail"]',
            '[class*="story-body"]',
            '[class*="post-content"]',
            '[class*="entry-content"]',
            '[class*="content-body"]',
            "main",
        ]
        for selector in selectors:
            for node in soup.select(selector)[:4]:
                parts = []
                for p in node.find_all(["p", "h2", "h3", "li"]):
                    txt = p.get_text(" ", strip=True)
                    if len(txt) >= 40:
                        parts.append(txt)
                if parts:
                    candidate = " ".join(parts)
                    if len(candidate) >= 250:
                        bodies.append(candidate)

        if not bodies:
            for p in soup.find_all("p"):
                txt = p.get_text(" ", strip=True)
                if len(txt) >= 45:
                    bodies.append(txt)

        for img in soup.find_all("img"):
            for attr in ("src", "data-src", "data-lazy-src", "data-original", "data-image"):
                value = img.get(attr)
                if value:
                    images.append(requests.compat.urljoin(rr.url, value))

        # Temizle
        seen=set()
        out["images"]=[]
        for u in images:
            if not isinstance(u,str): continue
            u=u.strip()
            if not u or u in seen: continue
            if any(x in u.lower() for x in ("favicon","sprite","avatar","logo")): continue
            seen.add(u); out["images"].append(u)
            if len(out["images"]) >= 20: break

        texts=[]
        seen_t=set()
        for body in bodies:
            body=re.sub(r"\s+"," ",html.unescape(body)).strip()
            if len(body)<120: continue
            key=norm(body[:700])
            if key in seen_t: continue
            seen_t.add(key); texts.append(body)
        texts.sort(key=len, reverse=True)
        if texts:
            out["text"]=" ".join(texts[:4])[:18000]

    # 3) Sayfa erişilemediyse bile RSS kaydını çöp etmiyoruz.
    # Generic Google Haberler adını asla gerçek yayıncı olarak rapora yazma.
    generic = {"google haberler","google news","google","google news rss","rss"}
    if norm(out["source"]) in generic:
        if publisher_name and norm(publisher_name) not in generic:
            out["source"] = publisher_name
        elif publisher_url:
            out["source"] = urlparse(publisher_url).netloc.replace("www.", "")
        else:
            out["source"] = "Açık Kaynak"

    # Başlık generic ise snippet/ekran başlığı kullan.
    if norm(out["title"]) in generic or not out["title"]:
        out["title"] = fallback_title or fallback_snippet

    if not out["text"] or len(out["text"]) < 250:
        out["text"] = fallback_snippet or out["title"]

    # Eğer gerçek URL çözüldüyse onu kullan; çözülmediyse Google News linkini rapora koyma.
    if not valid_article_url(out["canonical"]):
        out["canonical"] = publisher_url or original_url

    return out

def _download_report_image(url):
    if not url:
        return None
    try:
        rr = requests.get(
            url,
            headers={
                **HEADERS,
                "Accept": "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.7",
            },
            timeout=12,
        )
        if rr.status_code != 200 or len(rr.content) < 1200:
            return None

        im = Image.open(BytesIO(rr.content))
        if im.mode not in ("RGB", "L"):
            im = im.convert("RGB")
        im.thumbnail((1600, 1200), Image.LANCZOS)

        bio = BytesIO()
        im.save(bio, "JPEG", quality=88)
        bio.seek(0)
        return bio
    except Exception:
        return None


def _word_hyperlink(paragraph, url, label):
    if not url:
        paragraph.add_run(label)
        return

    try:
        rid = paragraph.part.relate_to(
            url,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            is_external=True,
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
        paragraph.add_run(url)


def _real_source(row, detail, real_url):
    generic = {"google haberler", "google news", "google", "google news rss", "rss"}

    for value in (
        detail.get("source"),
        row.get("Yayıncı"),
        row.get("Kaynak"),
    ):
        value = str(value or "").strip()
        if value and norm(value) not in generic:
            return value

    for value in (row.get("Yayıncı_URL"), real_url):
        value = str(value or "").strip()
        if valid_host := (urlparse(value).netloc.lower().replace("www.", "") if value else ""):
            if "google.com" not in valid_host:
                return valid_host

    return "Açık Kaynak"

def _expanded_report_text(title, body):
    body = re.sub(r"\s+", " ", (body or "")).strip()
    if not body:
        return title

    # Haber gövdesini mümkün olduğunca geniş tut; sadece çok uzun teknik/menü tekrarlarını kes.
    sentences = [
        s.strip()
        for s in re.split(r"(?<=[.!?])\s+", body)
        if len(s.strip()) >= 40
    ]

    if sentences:
        body = " ".join(sentences)

    return body[:15000]


def make_docx(rows):
    """
    Seçilen haberleri, yüklenen AKT/Açık Kaynak Taraması örneğinin
    anlatım yapısına yakın şekilde DOCX'e dönüştürür.
    """
    doc = Document()

    section = doc.sections[0]
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

    # Resmî rapor görünümü
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")

    # Başlık
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run("AÇIK KAYNAK TARAMA ÇALIŞMASI")
    r.bold = True
    r.font.name = "Times New Roman"
    r.font.size = Pt(14)

    # Üst bilgiler
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    p.add_run("Tarama Yapılan Görev Alanı ").bold = True
    p.add_run("Sanayi ve Teknoloji")

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    p.add_run("Tarih ").bold = True
    p.add_run(datetime.now().astimezone().strftime("%d.%m.%Y"))

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(10)
    p.add_run("Rapor Saati ").bold = True
    p.add_run(datetime.now().astimezone().strftime("%H:%M:%S"))

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    p.add_run("Bulgular:").bold = True

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Cm(0.75)
    p.paragraph_format.space_after = Pt(10)
    p.add_run(
        "Sanayi ve Teknoloji alanlarında yapılan açık kaynak taraması neticesinde, "
        f"seçilen {len(rows)} haber/İçeriğe ilişkin bulgular aşağıda sunulmuştur."
    )

    # Haberler
    for i, row in enumerate(rows, 1):
        detail = article_detail(row)

        real_url = detail.get("canonical") or row.get("Yayıncı_URL") or row.get("URL", "")
        title = (detail.get("title") or row.get("Başlık") or "").strip()
        source = _real_source(row, detail, real_url)
        body = detail.get("text") or row.get("İçerik_Özeti") or title
        expanded = _expanded_report_text(title, body)

        # Tek ana anlatım paragrafı — örnek rapordaki biçim.
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Cm(0.75)
        p.paragraph_format.space_after = Pt(5)

        nr = p.add_run(f"{i}. ")
        nr.bold = True

        p.add_run(
            f'“{source}” isimli internet sitesinde, “{title}” başlığıyla bir haber '
            "yayımlanmıştır. ("
        )
        _word_hyperlink(p, real_url, real_url if real_url else "Haber Linki")
        p.add_run(") Söz konusu haber içeriğinde, ")
        p.add_run(expanded)

        # Görsel başlığı
        image_stream = None
        image_url = ""
        for candidate in detail.get("images", []):
            image_stream = _download_report_image(candidate)
            if image_stream:
                image_url = candidate
                break

        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap.paragraph_format.space_before = Pt(5)
        cap.paragraph_format.space_after = Pt(5)

        cr = cap.add_run(
            f'Görsel {i}: “{source}” Sitesinde Yer Alan İçerik'
        )
        cr.bold = True
        cr.font.name = "Times New Roman"
        cr.font.size = Pt(11)

        if image_stream:
            ip = doc.add_paragraph()
            ip.alignment = WD_ALIGN_PARAGRAPH.CENTER
            ip.paragraph_format.space_after = Pt(2)
            ip.add_run().add_picture(image_stream, width=Cm(14.5))

            lp = doc.add_paragraph()
            lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            lp.paragraph_format.space_after = Pt(12)
            lp.add_run("(")
            _word_hyperlink(lp, image_url, image_url if image_url else "Görsel Linki")
            lp.add_run(")")
        else:
            lp = doc.add_paragraph()
            lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            lp.paragraph_format.space_after = Pt(12)
            lp.add_run("Görsel alınamadı.")
            if detail.get("images"):
                lp.add_run(" (")
                _word_hyperlink(lp, detail["images"][0], "Görsel Linki")
                lp.add_run(")")

    end = doc.add_paragraph()
    end.paragraph_format.space_before = Pt(10)
    end.add_run("Arz olunur.").bold = True

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()


def _section_select_table(section_key, data, columns, height=420):
    """Her haber bölümünde kutucuk gösterir; seçilenler doğrudan iki sepete eklenebilir."""
    if data is None or data.empty:
        return pd.DataFrame()

    tbl=data.copy()
    tbl['_row_key']=[
        str(r.get('URL','')) if str(r.get('URL','')).strip()
        else title_key(str(r.get('Başlık','')))
        for _,r in tbl.iterrows()
    ]
    selected_map=st.session_state.section_selections.get(section_key,{})
    if 'Seç' in tbl.columns:
        tbl['Seç']=[bool(selected_map.get(k, bool(v))) for k,v in zip(tbl['_row_key'], tbl['Seç'].tolist())]
    else:
        tbl.insert(0,'Seç',[bool(selected_map.get(k,False)) for k in tbl['_row_key']])

    show_cols=['Seç']+[c for c in columns if c in tbl.columns]
    edited=st.data_editor(
        tbl[show_cols+['_row_key']],
        column_config={
            'Seç':st.column_config.CheckboxColumn('Seç'),
            'URL':st.column_config.LinkColumn('Haber Linki'),
            'Risk_Skoru':st.column_config.NumberColumn('Risk',format='%d/100'),
            'Risk':st.column_config.NumberColumn('Risk',format='%d/100'),
            '_row_key':None
        },
        disabled=[c for c in show_cols if c!='Seç']+['_row_key'],
        hide_index=True,use_container_width=True,height=height,
        key=f'section_editor_{section_key}'
    )
    st.session_state.section_selections[section_key]={
        str(r['_row_key']):bool(r['Seç']) for _,r in edited.iterrows()
    }
    selected_keys={str(r['_row_key']) for _,r in edited.iterrows() if bool(r['Seç'])}
    selected=data[
        data.apply(
            lambda r:(str(r.get('URL','')) if str(r.get('URL','')).strip() else title_key(str(r.get('Başlık','')))) in selected_keys,
            axis=1
        )
    ].copy()

    # Kullanıcı kronolojiye dönmeden, seçtiği haberleri bulunduğu bölümden sepete atabilsin.
    if selected_keys:
        main_rows=pd.DataFrame(st.session_state.get('rows') or [])
        if not main_rows.empty:
            main_selected=main_rows[
                main_rows.apply(
                    lambda r:(str(r.get('URL','')) if str(r.get('URL','')).strip() else title_key(str(r.get('Başlık','')))) in selected_keys,
                    axis=1
                )
            ].copy()
        else:
            main_selected=pd.DataFrame()

        a1,a2=st.columns(2)
        with a1:
            if st.button('📌 Seçilenleri Önemli Gelişmeler Sepetine Ekle',key=f'imp_{section_key}',use_container_width=True):
                if main_selected.empty:
                    st.warning('Seçilen haber ana tarama kayıtlarıyla eşleştirilemedi.')
                else:
                    n=_add_rows_to_important_basket(main_selected.to_dict('records'))
                    st.success(f'{n} haber önemli gelişmeler sepetine eklendi.')
        with a2:
            if st.button('🗂️ Seçilenleri Açık Kaynak Tarama Sepetine Ekle',key=f'akt_{section_key}',use_container_width=True):
                if main_selected.empty:
                    st.warning('Seçilen haber ana tarama kayıtlarıyla eşleştirilemedi.')
                else:
                    n=_add_rows_to_osint_basket(main_selected.to_dict('records'))
                    st.success(f'{n} haber açık kaynak tarama sepetine eklendi.')

    return selected

def _collect_section_selected_from_main_df(df):
    if df is None or df.empty:
        return pd.DataFrame()
    keys=set()
    for selmap in st.session_state.section_selections.values():
        for k,v in selmap.items():
            if v:
                keys.add(str(k))
    if not keys:
        return pd.DataFrame()
    mask=df.apply(
        lambda r:(str(r.get('URL','')) if str(r.get('URL','')).strip() else title_key(str(r.get('Başlık','')))) in keys,
        axis=1
    )
    return df[mask].copy()

# -----------------------------
# UI
# -----------------------------
st.title('🛡️ Sanayi & Teknoloji Açık Kaynak / Negatif Haber Radarı')
st.caption('Hızlı ilk bakış · olay kümeleri · risk/negatif ayrımı · Türk medya önceliği · Yunan/Türk savunma · kaynak güvenilirliği · trend · alarm · seçilen haberlerden DOCX')
with st.sidebar:
    st.header('⚙️ Tarama Ayarları')
    default=('sanayi OR teknoloji OR üretim OR imalat OR fabrika OR OSB OR makine OR otomasyon OR robotik OR Ar-Ge OR patent OR yapay zeka OR yazılım OR siber güvenlik OR çip OR yarı iletken OR elektronik OR telekom OR kuantum OR biyoteknoloji OR nanoteknoloji OR savunma sanayii OR ASELSAN OR TUSAŞ OR ROKETSAN OR HAVELSAN OR Baykar OR İHA OR SİHA OR KAAN OR havacılık OR uzay OR uydu OR otomotiv OR TOGG OR batarya OR enerji OR hidrojen OR kimya OR petrokimya OR demir çelik OR madencilik OR tekstil OR gıda teknolojisi OR tarım teknolojisi OR lojistik OR tedarik zinciri OR TÜBİTAK OR KOSGEB OR teknopark OR yatırım teşvik OR yerlileştirme')
    query=st.text_area('Geniş sanayi / teknoloji sorgusu:',default,height=190)
    watch=st.text_area('⭐ Takip listesi (virgül / satır sonu):','ASELSAN, TUSAŞ, ROKETSAN, HAVELSAN, Baykar, TOGG, TÜBİTAK',height=90)
    neg=st.checkbox('⚠️ Negatif haberleri ayrıca tespit et',True)
    greek=st.checkbox('🇬🇷 Yunan medyası — yalnızca Türk savunma sanayii',True)
    social=st.checkbox('📱 Türk açık sosyal / indeks kaynakları',True)
    global_on=st.checkbox('🌍 Global basın (opsiyonel)',False)
    instant_alerts=st.checkbox('🔔 Tarama sırasında negatif/yüksek risk bildirimi göster',True,
                               help='Tarama devam ederken yeni negatif veya yüksek riskli içerik yakalanırsa ekranda anlık bildirim gösterir.')
    period=st.selectbox('🕒 Haber dönemi',['⚡ Son 3 saat','📅 Son 24 saat','📆 Son 48 saat','📆 Son 1 hafta','🗓️ Son 1 ay'],index=1)
    hours={'⚡ Son 3 saat':3,'📅 Son 24 saat':24,'📆 Son 48 saat':48,'📆 Son 1 hafta':168,'🗓️ Son 1 ay':720}[period]
    run=st.button('🔍 TARAMAYI BAŞLAT / YENİLE',type='primary',use_container_width=True)

if 'rows' not in st.session_state: st.session_state.rows=None
if 'scan_time' not in st.session_state: st.session_state.scan_time=None
if 'stats' not in st.session_state: st.session_state.stats={}
if 'last_scan_alerts' not in st.session_state: st.session_state.last_scan_alerts=[]
if 'docx_bytes' not in st.session_state: st.session_state.docx_bytes=None
if 'note_bytes' not in st.session_state: st.session_state.note_bytes=None

if 'current_scan_id' not in st.session_state: st.session_state.current_scan_id=None
if 'history_status' not in st.session_state: st.session_state.history_status=_init_history_db()
if 'basket_docx_bytes' not in st.session_state: st.session_state.basket_docx_bytes=None
if 'section_selections' not in st.session_state: st.session_state.section_selections={}


if run:
    cutoff=(datetime.now(timezone.utc)-timedelta(hours=hours)).astimezone(timezone.utc)
    when=period_window(hours)
    batches=[('🇹🇷 Türk medya / sanayi-teknoloji',build_turkish_queries(when,query),'turkish')]
    if neg: batches.append(('⚠️ Negatif haber taraması',build_negative_queries(when),'negative'))
    if greek: batches.append(('🇬🇷 Yunan medyası / Türk savunma',build_greek_queries(when),'greek'))
    if social: batches.append(('📱 Açık sosyal / indeks',build_social_queries(when),'social'))
    if global_on: batches.append(('🌍 Global basın',[
        f'(Turkey OR Türkiye) (industry OR manufacturing OR technology OR semiconductor OR defense OR aerospace OR automotive) timespan:{when}',
        f'(Turkey OR Turkish) (Baykar OR ASELSAN OR TUSAŞ OR ROKETSAN OR HAVELSAN OR KAAN OR drone OR missile) timespan:{when}'
    ],'global'))
    all_rows=[]; stat={'Ham sonuç':0,'Zaman dışı':0,'Konu dışı':0,'Yunan dışı':0,'Kaynak dışı':0,'Sonuç':0,'Olay':0}
    placeholder=st.empty()
    live_alarm_box=st.empty()
    status_box=st.status('🔎 Tarama başlıyor...',expanded=True)

    alerted_keys=set()
    live_alerts=[]
    toast_count=0
    MAX_TOASTS_PER_SCAN=2

    def _alert_key(row):
        return row.get('URL') or title_key(row.get('Başlık',''))

    def _register_alert(row):
        key=_alert_key(row)
        if not key or key in alerted_keys:
            return False
        alerted_keys.add(key)
        risk_score=int(row.get('Risk_Skoru',row.get('Skor',0)) or 0)
        osb_fire=is_osb_fire(row.get('Başlık',''),row.get('İçerik_Özeti',''))
        is_high=row.get('Risk_Durumu')=='Yüksek Risk' or risk_score>=70 or osb_fire
        live_alerts.insert(0,{
            'Tarih':str(row.get('Tarih','')),
            'Seviye':'🔥 OSB YANGINI' if osb_fire else ('YÜKSEK RİSK' if is_high else 'NEGATİF'),
            'Kaynak':str(row.get('Kaynak','Açık Kaynak')),
            'Başlık':str(row.get('Başlık','')),
            'Risk':risk_score,
            'URL':row.get('URL','')
        })
        del live_alerts[25:]
        return True

    def _merge_batch(raw,mode):
        nonlocal_dummy=None
        norm_rows,reasons=normalize_rows(raw,cutoff,mode,query)
        stat['Zaman dışı']+=reasons['zaman']
        stat['Konu dışı']+=reasons['konu']
        stat['Yunan dışı']+=reasons['yunan']
        stat['Kaynak dışı']+=reasons['kaynak']
        return norm_rows

    # 1) Türk ana taraması önce: kullanıcı ilk sonuçları en kısa sürede görsün.
    primary_label,primary_queries,primary_mode=batches[0]
    status_box.write(f'{primary_label} — {len(primary_queries)} sorgu / 10 eşzamanlı')
    primary_raw=[]
    with concurrent.futures.ThreadPoolExecutor(max_workers=min(10,len(primary_queries))) as ex:
        futures=[ex.submit(rss,q) for q in primary_queries]
        for f in concurrent.futures.as_completed(futures):
            try:
                primary_raw.extend(f.result() or [])
            except Exception:
                pass
    stat['Ham sonuç']+=len(primary_raw)
    all_rows=dedupe(_merge_batch(primary_raw,primary_mode))
    stat['Sonuç']=len(all_rows)

    if all_rows:
        pv=pd.DataFrame(all_rows)
        pv['Tarih_dt']=pd.to_datetime(pv['Tarih_dt'],utc=True,errors='coerce')
        pv=pv.sort_values(['Tarih_dt','Domain'],ascending=[False,True],na_position='last')
        fast=pv[['Tarih','Kaynak_Grubu','Kaynak','Başlık','İçerik_Özeti','Duygu','Risk_Skoru','Risk_Durumu','URL']].head(100).copy()
        fast['İçerik_Özeti']=fast['İçerik_Özeti'].astype(str).str.slice(0,260)
        placeholder.dataframe(
            fast,
            column_config={'URL':st.column_config.LinkColumn('Haber Linki'),'İçerik_Özeti':st.column_config.TextColumn('Kısa İçerik',width='large'),'Risk_Skoru':st.column_config.NumberColumn('Risk',format='%d/100')},
            hide_index=True,use_container_width=True,height=440
        )

    # 2) Negatif + Yunan + sosyal + global sorgularını TEK HAVUZDA paralel çalıştır.
    supplemental=batches[1:]
    jobs=[]
    for label,queries,mode in supplemental:
        for q in queries:
            jobs.append((label,q,mode))

    supplemental_raw_by_mode={}
    if jobs:
        status_box.write(f'⚡ Tamamlayıcı kaynaklar — {len(jobs)} sorgu / 12 eşzamanlı')
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(12,len(jobs))) as ex:
            future_map={ex.submit(rss,q):(label,mode) for label,q,mode in jobs}
            for fut in concurrent.futures.as_completed(future_map):
                label,mode=future_map[fut]
                try:
                    chunk=fut.result() or []
                except Exception:
                    chunk=[]
                stat['Ham sonuç']+=len(chunk)
                supplemental_raw_by_mode.setdefault(mode,[]).extend(chunk)

                # ÖZEL OSB YANGINI ALARMI:
                # İlgili sorgu döner dönmez tüm tamamlayıcı taramanın bitmesini beklemeden bildir.
                if instant_alerts and chunk:
                    quick_rows,_quick_reasons=normalize_rows(chunk,cutoff,mode,query)
                    for qr in quick_rows:
                        if is_osb_fire(qr.get('Başlık',''),qr.get('İçerik_Özeti','')):
                            qkey=_alert_key(qr)
                            if qkey not in alerted_keys:
                                _register_alert(qr)
                                st.toast(
                                    f'🔥 OSB YANGINI: {str(qr.get("Başlık",""))[:105]}',
                                    icon='🔥'
                                )
                                live_alarm_box.error(
                                    f'🔥 **ORGANİZE SANAYİ BÖLGESİ YANGIN ALARMI** — '
                                    f'{qr.get("Tarih","")} · {qr.get("Kaynak","Açık Kaynak")} · '
                                    f'{str(qr.get("Başlık",""))[:140]}'
                                )

    # Mode bazlı normalize + birleştirme.
    for mode,raw in supplemental_raw_by_mode.items():
        incoming=_merge_batch(raw,mode)
        old_keys={_alert_key(x) for x in all_rows}
        all_rows=dedupe(all_rows+incoming)
        for ar in all_rows:
            key=_alert_key(ar)
            if key in old_keys:
                continue
            if ar.get('Duygu')=='Negatif' or ar.get('Risk_Durumu')=='Yüksek Risk':
                if _register_alert(ar):
                    if instant_alerts and toast_count < MAX_TOASTS_PER_SCAN:
                        risk_score=int(ar.get('Risk_Skoru',ar.get('Skor',0)) or 0)
                        is_high=ar.get('Risk_Durumu')=='Yüksek Risk' or risk_score>=70
                        st.toast(
                            f'{"🚨 YÜKSEK RİSK" if is_high else "⚠️ NEGATİF"}: {str(ar.get("Başlık",""))[:100]}',
                            icon='🚨' if is_high else '⚠️'
                        )
                        toast_count+=1
        stat['Sonuç']=len(all_rows)

    if live_alerts:
        live_alarm_box.warning(
            f'🔔 {len(live_alerts)} negatif/riskli içerik yakalandı. Son: {live_alerts[0]["Başlık"][:100]}'
        )

    # 3) Analitik katman yalnızca bir kez ve artık hızlı ters indeks ile.
    if all_rows:
        status_box.write('🧩 Hızlı olay analizi hazırlanıyor...')
        all_rows=enrich_rows(all_rows)
        stat['Olay']=len({r.get('Olay_ID') for r in all_rows})
    else:
        stat['Olay']=0

    for ar in all_rows:
        if ar.get('Duygu')=='Negatif' or ar.get('Risk_Durumu')=='Yüksek Risk' or int(ar.get('Risk_Skoru',0) or 0)>=70:
            _register_alert(ar)

    status_box.update(
        label=f'✅ Tarama tamamlandı — {len(all_rows)} haber / {stat["Olay"]} olay',
        state='complete'
    )
    st.session_state.rows=all_rows
    st.session_state.scan_time=datetime.now().astimezone()
    st.session_state.stats=stat
    st.session_state.last_scan_alerts=live_alerts

    # V33 geçmiş karşılaştırma katmanı: tarama bittikten SONRA olay özetini kaydeder.
    # Tarama motoruna veya sıralamaya müdahale etmez.
    st.session_state.current_scan_id=_save_scan_history(
        all_rows,
        st.session_state.scan_time,
        hours
    )


rows=st.session_state.rows
if rows is None:
    st.info('👋 Hazır. Tarama başlamaz. Zaman aralığını seçip **TARAMAYI BAŞLAT / YENİLE** düğmesine basın.')
else:
    # Tarama sırasında satırlar zaten enrich_rows() ile zenginleştiriliyor.
    # Checkbox / sekme / buton gibi UI etkileşimlerinde pahalı analizi tekrar çalıştırmıyoruz.
    df=pd.DataFrame(rows)
    if not df.empty:
        df['Tarih_dt']=pd.to_datetime(df['Tarih_dt'],utc=True,errors='coerce')
        df=df.sort_values('Tarih_dt',ascending=False,na_position='last').reset_index(drop=True)
    st.caption(f'Son tarama: {st.session_state.scan_time.strftime("%d.%m.%Y %H:%M:%S") if st.session_state.scan_time else "-"}')
    with st.expander('🧪 Tarama teşhisi',False): st.json(st.session_state.stats)
    if df.empty:
        st.warning('Sonuç bulunamadı. Tarama teşhisini açarak hangi aşamada sonuçların azaldığını görebilirsiniz.')
    else:
        total=len(df); negc=int((df.Duygu=='Negatif').sum()); riskc=int((df.Risk_Durumu=='Yüksek Risk').sum()); trc=int(df.Kaynak_Grubu.astype(str).str.startswith('🇹🇷').sum()); grc=int(df.Kaynak_Grubu.astype(str).str.startswith('🇬🇷').sum()); events=df['Olay_ID'].nunique()
        a,b,c,d,e,f=st.columns(6); a.metric('Toplam',total); b.metric('Olay',events); c.metric('Negatif',negc); d.metric('Yüksek Risk',riskc); e.metric('🇹🇷 Türk',trc); f.metric('🇬🇷 Yunan',grc)



        # ---------------------------------------------------------
        # V34 — VARDİYA BAŞLANGIÇ ÖZETİ
        # ---------------------------------------------------------
        st.subheader('🌅 Vardiya Başlangıç Özeti')
        shift_stats,shift_top,shift_baseline_label=_shift_start_summary(
            df,
            st.session_state.get('current_scan_id')
        )
        if shift_stats:
            st.caption(shift_stats.get('baseline_label',''))
            s1,s2,s3,s4,s5,s6=st.columns(6)
            s1.metric('Son devirden beri yeni haber',shift_stats['new_news'])
            s2.metric('Yeni önemli olay',shift_stats['new_important_events'])
            s3.metric('Yüksek riskli gelişme',shift_stats['high_risk'])
            s4.metric('Risk artışı',shift_stats['risk_up'])
            s5.metric('Teyit güçlenmesi',shift_stats['verify_up'])
            s6.metric('OSB olayı',shift_stats['osb'])

            st.markdown('**Sabah ilk bakılması gereken 5 gelişme**')
            if shift_top.empty:
                st.info('Öne çıkan gelişme bulunamadı.')
            else:
                _section_select_table(
                    'shift_top',
                    shift_top,
                    ['Tarih','Kaynak','Kategori','Başlık','Risk_Skoru','Risk_Durumu','Doğrulama','URL'],
                    height=min(330,70+45*len(shift_top))
                )

        c_shift1,c_shift2=st.columns([2,1])
        with c_shift1:
            st.caption('Devir noktası, bir sonraki vardiya başlangıç özetinin başlangıç zamanını belirler.')
        with c_shift2:
            if st.button('📍 ŞİMDİYİ DEVİR NOKTASI OLARAK KAYDET',use_container_width=True):
                if _mark_shift_handover(st.session_state.get('current_scan_id'),'Manuel devir noktası'):
                    st.success('Devir noktası kaydedildi.')
                else:
                    st.error('Devir noktası kaydedilemedi.')

        # ---------------------------------------------------------
        # V34 — ÖNEMLİ GELİŞMELER SEPETİ
        # ---------------------------------------------------------
        st.subheader('📌 24 Saatlik Önemli Gelişmeler Sepeti')
        st.caption('Gün boyunca önemli gördüğünüz haberleri burada biriktirin; vardiya sonunda Word olarak alın.')

        selected_now=df[df.get('Seç',False)==True] if 'Seç' in df.columns else pd.DataFrame()
        selected_from_sections=_collect_section_selected_from_main_df(df)

        if st.button('➕ BÖLÜMLERDE İŞARETLEDİKLERİMİ ÖNEMLİ GELİŞMELER SEPETİNE EKLE',use_container_width=True):
            if selected_from_sections.empty:
                st.warning('Önce herhangi bir bölümde haberlerin yanındaki kutucuklardan seçim yapın.')
            else:
                added=_add_rows_to_important_basket(selected_from_sections.to_dict('records'))
                st.success(f'{added} yeni gelişme sepete eklendi.')
        if st.button('➕ SEÇİLİ HABERLERİ ÖNEMLİ GELİŞMELER SEPETİNE EKLE',use_container_width=True):
            if selected_now.empty:
                st.warning('Önce kronolojik görünümden haber seçin ve seçimleri kaydedin.')
            else:
                added=_add_rows_to_important_basket(selected_now.to_dict('records'))
                st.success(f'{added} yeni gelişme sepete eklendi.')

        basket=_load_important_basket()
        if basket.empty:
            st.info('Önemli gelişmeler sepeti şu anda boş.')
        else:
            basket_view=basket[['id','news_time','source','category','title','risk_score','risk_status','url']].copy()
            basket_view.insert(0,'Sil',False)
            with st.form('important_basket_form',clear_on_submit=False):
                edited_basket=st.data_editor(
                    basket_view,
                    column_config={
                        'Sil':st.column_config.CheckboxColumn('Sil'),
                        'url':st.column_config.LinkColumn('Haber Linki'),
                        'risk_score':st.column_config.NumberColumn('Risk',format='%d/100')
                    },
                    disabled=[c for c in basket_view.columns if c!='Sil'],
                    hide_index=True,use_container_width=True,height=min(430,80+36*len(basket_view))
                )
                remove_btn=st.form_submit_button('🗑️ İŞARETLENENLERİ SEPETTEN ÇIKAR',use_container_width=True)
            if remove_btn:
                ids=edited_basket.loc[edited_basket['Sil']==True,'id'].tolist()
                removed=_remove_basket_ids(ids)
                st.success(f'{removed} kayıt sepetten çıkarıldı.')

            b1,b2=st.columns(2)
            with b1:
                st.session_state.basket_docx_bytes=make_important_basket_docx(basket)
                st.download_button(
                    '⬇️ 24 SAATLİK ÖNEMLİ GELİŞMELER / WORD',
                    st.session_state.basket_docx_bytes,
                    file_name=f'24_Saatlik_Onemli_Gelismeler_{date.today()}.docx',
                    mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    use_container_width=True
                )
            with b2:
                if st.button('🧹 SEPETİ TAMAMEN TEMİZLE',use_container_width=True):
                    removed=_clear_important_basket()
                    st.success(f'{removed} kayıt silindi.')

        # ---------------------------------------------------------
        # V33 — BİLGİ NOTU ADAYLARI
        # ---------------------------------------------------------
        st.subheader('🎯 Bilgi Notu Adayları')
        st.caption('Mevcut taramadaki olayları risk, teyit, kaynak sayısı, stratejik önem ve yenilik açısından puanlar.')
        candidate_count=st.slider('Gösterilecek aday sayısı',5,15,10,1,key='candidate_count')
        candidates=_information_note_candidates(
            df,
            st.session_state.get('current_scan_id'),
            candidate_count
        )
        if candidates.empty:
            st.info('Bu taramada bilgi notu adayı oluşturulamadı.')
        else:
            _section_select_table(
                'candidates',
                candidates,
                ['Aday Puanı','Başlık','Kaynak','Kategori','Risk','Kaynak Sayısı','Doğrulama','Değişim','Neden Bilgi Notu?','URL'],
                height=min(470,65+36*len(candidates))
            )

        # ---------------------------------------------------------
        # V33 — DÜNDEN BERİ NE DEĞİŞTİ?
        # ---------------------------------------------------------
        st.subheader('🆕 Dünden Beri Ne Değişti?')
        changes,previous_scan_id,previous_scan_time=_compare_since_previous(
            df,
            st.session_state.get('current_scan_id')
        )
        if previous_scan_id is None:
            st.info(
                'Henüz karşılaştırılabilecek eski tarama bulunmuyor. '
                'Bu tarama yerel geçmişe kaydedildi; sonraki taramalarda yeni olaylar ve değişiklikler otomatik gösterilecek.'
            )
        else:
            if previous_scan_time:
                st.caption(f'Karşılaştırılan önceki tarama: {previous_scan_time}')
            if changes.empty:
                st.success('Önceki taramaya göre anlamlı yeni olay, risk artışı, teyit artışı veya içerik güncellemesi tespit edilmedi.')
            else:
                new_n=int(changes['Değişim'].astype(str).str.contains('YENİ OLAY').sum())
                upd_n=int(changes['Değişim'].astype(str).str.contains('YENİ BİLGİ').sum())
                risk_n=int(changes['Değişim'].astype(str).str.contains('RİSK ARTTI').sum())
                ver_n=int(changes['Değişim'].astype(str).str.contains('TEYİT').sum())
                q1,q2,q3,q4=st.columns(4)
                q1.metric('Yeni Olay',new_n)
                q2.metric('Yeni Bilgi',upd_n)
                q3.metric('Risk Artışı',risk_n)
                q4.metric('Teyit Güçlendi',ver_n)
                changes_view=changes.head(25).copy()
                _section_select_table(
                    'changes',
                    changes_view,
                    ['Değişim','Başlık','Kaynak','Kategori','Risk','Önceki Risk','Kaynak Sayısı','Açıklama','URL'],
                    height=min(560,70+35*min(len(changes_view),25))
                )

        # Son taramada anlık yakalanan bildirimlerin kalıcı özeti
        recent_alerts=st.session_state.get('last_scan_alerts',[])
        if recent_alerts:
            with st.expander(f'🔔 Son taramada yakalanan yeni negatif/riskli içerikler ({len(recent_alerts)})',False):
                alert_df=pd.DataFrame(recent_alerts)
                alert_view=alert_df.copy()
                if 'Risk' in alert_view.columns and 'Risk_Skoru' not in alert_view.columns:
                    alert_view['Risk_Skoru']=alert_view['Risk']
                _section_select_table(
                    'recent_alerts',
                    alert_view,
                    ['Tarih','Seviye','Kaynak','Başlık','Risk_Skoru','URL'],
                    height=min(420,42+35*len(alert_view))
                )

        # Organize Sanayi Bölgesi yangınları için özel kalıcı alarm bölümü
        osb_fire_mask=df.apply(
            lambda r:is_osb_fire(r.get('Başlık',''),r.get('İçerik_Özeti','')),
            axis=1
        )
        osb_fires=df[osb_fire_mask].sort_values('Tarih_dt',ascending=False)
        if not osb_fires.empty:
            st.error(f'🔥 **OSB YANGIN ALARMI — {len(osb_fires)} içerik tespit edildi**')
            _section_select_table(
                'osb_fires',
                osb_fires,
                ['Tarih','Kaynak','Başlık','Risk_Skoru','URL'],
                height=min(300,60+36*len(osb_fires))
            )

        # Alarm bandı
        alarms=df[(df.Risk_Skoru>=70) | (df.Duygu=='Negatif')].sort_values(['Risk_Skoru','Tarih_dt'],ascending=[False,False])
        if not alarms.empty:
            st.subheader('🚨 Yeni / Öncelikli Alarmlar')
            alarm_view=alarms.head(10).copy()
            _section_select_table(
                'priority_alarms',
                alarm_view,
                ['Tarih','Kaynak','Kategori','Başlık','Risk_Skoru','Risk_Gerekçesi','Doğrulama','URL'],
                height=min(470,70+38*len(alarm_view))
            )

        # Performans: Streamlit tabs içindeki TÜM içerikleri arka planda çalıştırır.
        # Bu nedenle tek seferde yalnızca seçilen görünümü üretiriz.
        view=st.radio(
            'Görünüm',
            ['📰 Kronolojik','⚠️ Negatif','🚨 Yüksek Risk','🇹🇷 Türk','🇬🇷 Yunan','🧩 Olaylar','📈 Trend / Analiz','⭐ Takip Listesi'],
            horizontal=True,
            key='main_view'
        )

        cols=['Seç','Tarih','Kaynak_Grubu','Kaynak','Kategori','Başlık','İçerik_Özeti','Duygu','Risk_Skoru','Risk_Durumu','Kaynak_Güvenilirliği','Doğrulama','URL']

        if view=='📰 Kronolojik':
            st.caption('☑️ Tüm haberler sistemde tutulur; hız için ekranda sayfa sayfa gösterilir.')
            page_size=75
            total_pages=max(1,(len(df)+page_size-1)//page_size)
            page_no=st.number_input('Sayfa',min_value=1,max_value=total_pages,value=1,step=1,key='news_page')
            start_i=(int(page_no)-1)*page_size
            end_i=min(start_i+page_size,len(df))
            page_df=df.iloc[start_i:end_i].copy()

            # Tarayıcıya çok uzun RSS özetleri göndermeyelim; tam içerik backend'de korunur.
            page_df['İçerik_Özeti']=page_df['İçerik_Özeti'].astype(str).str.slice(0,320)

            st.caption(f'{start_i+1}-{end_i} / {len(df)} haber')
            with st.form(key=f'selection_form_{st.session_state.scan_time}_{int(page_no)}', clear_on_submit=False):
                edited=st.data_editor(
                    page_df[cols],
                    column_config={
                        'Seç':st.column_config.CheckboxColumn('Seç'),
                        'URL':st.column_config.LinkColumn('Haber Linki'),
                        'İçerik_Özeti':st.column_config.TextColumn('Kısa İçerik',width='large'),
                        'Risk_Skoru':st.column_config.NumberColumn('Risk',format='%d/100')
                    },
                    disabled=[x for x in cols if x!='Seç'],
                    hide_index=True,use_container_width=True,height=560,
                    key=f'editor_{st.session_state.scan_time}_{int(page_no)}'
                )
                save_selection=st.form_submit_button('✅ BU SAYFADAKİ SEÇİMLERİ KAYDET',use_container_width=True)
            if save_selection:
                original_indices=df.index[start_i:end_i]
                df.loc[original_indices,'Seç']=edited['Seç'].astype(bool).to_numpy()
                st.session_state.rows=df.to_dict('records')
                st.success(f'✅ Toplam {int(df["Seç"].sum())} haber seçili.')

        elif view=='⚠️ Negatif':
            _section_select_table(
                'negative_view',
                df[df.Duygu=='Negatif'],
                ['Tarih','Kaynak','Kategori','Başlık','Risk_Skoru','Risk_Gerekçesi','Doğrulama','URL'],
                height=600
            )

        elif view=='🚨 Yüksek Risk':
            _section_select_table(
                'highrisk_view',
                df[df.Risk_Durumu=='Yüksek Risk'],
                ['Tarih','Kaynak','Kategori','Başlık','Risk_Skoru','Risk_Gerekçesi','Doğrulama','URL'],
                height=600
            )

        elif view=='🇹🇷 Türk':
            _section_select_table(
                'turkish_view',
                df[df.Kaynak_Grubu.astype(str).str.startswith('🇹🇷')],
                ['Tarih','Kaynak','Kategori','Başlık','Risk_Skoru','Duygu','URL'],
                height=600
            )

        elif view=='🇬🇷 Yunan':
            _section_select_table(
                'greek_view',
                df[df.Kaynak_Grubu.astype(str).str.startswith('🇬🇷')],
                ['Tarih','Kaynak','Kategori','Başlık','Risk_Skoru','Duygu','URL'],
                height=600
            )

        elif view=='🧩 Olaylar':
            ev=build_event_summary(df)
            st.dataframe(ev,hide_index=True,use_container_width=True,height=480)
            chosen=st.selectbox('Olay zaman çizelgesini göster:',ev['Olay_ID'].tolist() if not ev.empty else [])
            if chosen:
                g=df[df.Olay_ID==chosen].sort_values('Tarih_dt',ascending=True)
                _section_select_table(
                    f'event_{chosen}',
                    g,
                    ['Tarih','Kaynak','Kategori','Başlık','Risk_Skoru','Doğrulama','URL'],
                    height=min(500,80+40*len(g))
                )

        elif view=='📈 Trend / Analiz':
            st.subheader('📊 Konu yoğunluğu')
            tr=trend_table(df)
            if not tr.empty:
                st.bar_chart(tr.set_index('Kategori')['Haber'])
            st.subheader('📈 Gündem yoğunluğu')
            tmp=df[df['Tarih_dt'].notna()].copy()
            tmp['Saat']=tmp['Tarih_dt'].dt.strftime('%Y-%m-%d %H:00')
            if not tmp.empty:
                st.line_chart(tmp.groupby('Saat').size())
            st.subheader('🧭 Yoğun konular')
            for _,r in tr.head(10).iterrows():
                st.write(f"**{r['Kategori']}** — {int(r['Haber'])} haber")

        elif view=='⭐ Takip Listesi':
            hits=watchlist_hits(df,watch)
            st.write(f'Listede eşleşen: **{len(hits)}** haber')
            if not hits.empty:
                _section_select_table(
                    'watchlist_view',
                    hits,
                    ['Tarih','Kaynak','Kategori','Başlık','Risk_Skoru','Duygu','URL'],
                    height=550
                )

        st.markdown('---')
        st.subheader('🗂️ Açık Kaynak Tarama Çalışması Sepeti')
        st.caption('14:00 açık kaynak tarama raporuna girecek haberleri gün boyunca ayrı bir sepette biriktirin.')

        osint_selected_now=df[df.get('Seç',False)==True] if 'Seç' in df.columns else pd.DataFrame()
        osint_selected_sections=_collect_section_selected_from_main_df(df)

        o1,o2=st.columns(2)
        with o1:
            if st.button('➕ KRONOLOJİDE SEÇİLİ HABERLERİ AKT SEPETİNE EKLE',use_container_width=True):
                if osint_selected_now.empty:
                    st.warning('Önce kronolojik görünümden haber seçin ve seçimleri kaydedin.')
                else:
                    added=_add_rows_to_osint_basket(osint_selected_now.to_dict('records'))
                    st.success(f'{added} haber AKT sepetine eklendi.')
        with o2:
            if st.button('➕ BÖLÜMLERDE İŞARETLEDİKLERİMİ AKT SEPETİNE EKLE',use_container_width=True):
                if osint_selected_sections.empty:
                    st.warning('Önce herhangi bir bölümde seçim yapın.')
                else:
                    added=_add_rows_to_osint_basket(osint_selected_sections.to_dict('records'))
                    st.success(f'{added} haber AKT sepetine eklendi.')

        osint_basket=_load_osint_basket()
        if osint_basket.empty:
            st.info('Açık kaynak tarama çalışması sepeti boş.')
        else:
            osint_view=osint_basket[['id','news_time','source','category','title','risk_score','risk_status','url']].copy()
            osint_view.insert(0,'Sil',False)
            with st.form('osint_basket_form',clear_on_submit=False):
                edited_osint=st.data_editor(
                    osint_view,
                    column_config={
                        'Sil':st.column_config.CheckboxColumn('Sil'),
                        'url':st.column_config.LinkColumn('Haber Linki'),
                        'risk_score':st.column_config.NumberColumn('Risk',format='%d/100')
                    },
                    disabled=[c for c in osint_view.columns if c!='Sil'],
                    hide_index=True,use_container_width=True,height=min(430,80+36*len(osint_view))
                )
                remove_osint=st.form_submit_button('🗑️ İŞARETLENENLERİ AKT SEPETİNDEN ÇIKAR',use_container_width=True)
            if remove_osint:
                ids=edited_osint.loc[edited_osint['Sil']==True,'id'].tolist()
                removed=_remove_osint_basket_ids(ids)
                st.success(f'{removed} kayıt AKT sepetinden çıkarıldı.')

            osint_rows=[]
            for _,r in osint_basket.iterrows():
                osint_rows.append({
                    'Tarih':r.get('news_time',''),
                    'Kaynak':r.get('source',''),
                    'Başlık':r.get('title',''),
                    'İçerik_Özeti':r.get('summary',''),
                    'URL':r.get('url',''),
                    'Kategori':r.get('category',''),
                    'Risk_Skoru':r.get('risk_score',0),
                    'Risk_Durumu':r.get('risk_status',''),
                    'Yayıncı':r.get('source',''),
                    'Yayıncı_URL':''
                })

            ob1,ob2=st.columns(2)
            with ob1:
                if st.button('📝 AKT SEPETİNDEN WORD HAZIRLA',use_container_width=True):
                    with st.spinner('AKT sepetindeki haberler rapora hazırlanıyor...'):
                        st.session_state.docx_bytes=make_docx(osint_rows)
                if st.session_state.get('docx_bytes'):
                    st.download_button(
                        '⬇️ AKT SEPETİNDEN AÇIK KAYNAK RAPORU / WORD',
                        st.session_state.docx_bytes,
                        file_name=f'Sanayi_Teknoloji_Acik_Kaynak_Sepet_{date.today()}.docx',
                        mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                        use_container_width=True
                    )
            with ob2:
                if st.button('🧹 AKT SEPETİNİ TAMAMEN TEMİZLE',use_container_width=True):
                    removed=_clear_osint_basket()
                    st.success(f'{removed} kayıt silindi.')

        st.markdown('---')
        st.subheader('📊 Günlük Sanayi ve Teknoloji Durum Özeti')
        st.caption('Mevcut tarama sonuçlarının tamamından otomatik günlük görünüm ve günün öne çıkan 5 gelişmesini oluşturur.')

        daily_stats=_daily_summary_stats(df)
        if daily_stats:
            c1,c2,c3,c4,c5,c6=st.columns(6)
            c1.metric('Toplam Haber',daily_stats['total'])
            c2.metric('Negatif',daily_stats['negative'])
            c3.metric('Yüksek Risk',daily_stats['high_risk'])
            c4.metric('OSB Yangını',daily_stats['osb_fire'])
            c5.metric('Savunma',daily_stats['defence'])
            c6.metric('Siber',daily_stats['cyber'])

            daily_top_select=_daily_top_events(df,5)
            if not daily_top_select.empty:
                st.markdown('**Bugünün öne çıkan 5 gelişmesi**')
                _section_select_table(
                    'daily_top5',
                    daily_top_select,
                    ['Tarih','Kaynak','Kategori','Başlık','Risk_Skoru','Risk_Durumu','Doğrulama','URL'],
                    height=min(330,70+45*len(daily_top_select))
                )

        if st.button('📊 BUGÜNÜN DURUM ÖZETİNİ OLUŞTUR',use_container_width=True):
            with st.spinner('Günlük durum özeti hazırlanıyor...'):
                daily_text,daily_top,daily_stats=_daily_summary_text(df)
                st.session_state.daily_summary_text=daily_text
                st.session_state.daily_summary_bytes=make_daily_summary_docx(df)

        if st.session_state.get('daily_summary_text'):
            st.text_area(
                'Günlük Durum Özeti',
                st.session_state.daily_summary_text,
                height=360,
                key='daily_summary_preview'
            )
            st.download_button(
                '⬇️ GÜNLÜK DURUM ÖZETİNİ WORD OLARAK İNDİR',
                data=st.session_state.daily_summary_bytes,
                file_name=f'gunluk_sanayi_teknoloji_durum_ozeti_{datetime.now().strftime("%Y%m%d_%H%M")}.docx',
                mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                use_container_width=True
            )

        st.markdown('---'); st.subheader('📝 Seçilen haberlerden çıktı üret')
        # Form gönderildiyse session_state güncellenmiştir; aksi halde mevcut kayıtlı seçimleri kullan.
        current_rows=st.session_state.rows or []
        selected_df=pd.DataFrame(current_rows)
        if not selected_df.empty and 'Tarih_dt' in selected_df.columns:
            selected_df['Tarih_dt']=pd.to_datetime(selected_df['Tarih_dt'],utc=True,errors='coerce')
        selected=selected_df[selected_df.get('Seç',False)==True] if not selected_df.empty and 'Seç' in selected_df.columns else pd.DataFrame()
        st.write(f'{len(selected)} haber seçildi. Tarama sırasında görsel/tam metin indirilmez; yalnızca seçtiğiniz içerikler için derin zenginleştirme yapılır.')
        c1,c2=st.columns(2)
        with c1:
            if st.button('📝 AÇIK KAYNAK RAPORU / WORD',type='primary',use_container_width=True):
                if selected.empty: st.warning('Önce haber seçin.')
                else:
                    with st.spinner(f'{len(selected)} haber zenginleştiriliyor...'): st.session_state.docx_bytes=make_docx(selected.to_dict('records'))
            if st.session_state.docx_bytes: st.download_button('⬇️ Açık Kaynak Raporu DOCX',st.session_state.docx_bytes,file_name=f'Sanayi_Teknoloji_Acik_Kaynak_{date.today()}.docx',mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',use_container_width=True)
        with c2:
            if st.button('📌 AYRINTILI BİLGİ NOTU / WORD',use_container_width=True):
                if selected.empty: st.warning('Önce haber seçin.')
                else:
                    with st.spinner(f'{len(selected)} haberin tam haber metni okunuyor ve ayrıntılı bilgi notu hazırlanıyor...'):
                        st.session_state.note_bytes=make_analyst_docx(selected,title='SANAYİ & TEKNOLOJİ BİLGİ NOTU')
            if st.session_state.note_bytes: st.download_button('⬇️ Bilgi Notu DOCX',st.session_state.note_bytes,file_name=f'Sanayi_Teknoloji_Bilgi_Notu_{date.today()}.docx',mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',use_container_width=True)

st.caption('İlk açılışta otomatik tarama yoktur. Her yenileme yeni ağ taraması yapar. Haberler en yeni → en eski sıralanır; olay kümeleri, risk gerekçesi, kaynak güvenilirliği, doğrulama, trend ve takip listesi tarama sonucunda yer alır. DOCX aşamasında seçilen haberlerin gerçek yayıncı sayfası, görseli, linki ve geniş içeriği alınır.')
