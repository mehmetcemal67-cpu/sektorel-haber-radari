import streamlit as st
import pandas as pd
import requests
import concurrent.futures
import xml.etree.ElementTree as ET
import re, html, json
from datetime import datetime, timedelta, timezone, date
from urllib.parse import urlparse
from email.utils import parsedate_to_datetime
from io import BytesIO
from bs4 import BeautifulSoup
from PIL import Image
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml

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

def fmt_dt(d):
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

def classify(title,snippet):
    t=norm(f'{title} {snippet}')
    neg=[x for x in NEGATIVE_TERMS if x in t]
    risk=[x for x in HIGH_RISK_TERMS if x in t]
    score=max(-1.0,min(1.0,-0.24*len(neg)-0.32*len(risk)))
    if score<=-0.2: sentiment='Negatif'
    else: sentiment='Nötr'
    if risk or len(neg)>=3: status='Yüksek Risk'
    elif neg: status='Negatif'
    else: status='Normal'
    cat='Genel Sanayi / Teknoloji'
    for c,ks in CATEGORIES.items():
        if any(k in t for k in ks): cat=c; break
    return sentiment,round(score,2),status,neg,risk,cat

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
    # Kullanıcının kutuya eklediği özel terimler de ayrıca taranır.
    custom=[x for x in _query_terms(user_query) if norm(x) not in {'sanayi','teknoloji','üretim','imalat','fabrika','türkiye','türk'}]
    for term in custom[:20]:
        qs.append(f'Türkiye ("{term}") when:{when}')
    return qs

def build_negative_queries(when):
    return [f'Türkiye (iflas OR konkordato OR "üretim durdu" OR "fabrika kapandı" OR "işten çıkarma" OR grev OR soruşturma OR dava OR ceza OR "geri çağırma" OR "siber saldırı" OR "veri sızıntısı" OR yaptırım OR ambargo OR "ihale iptal" OR ertelendi OR gecikme OR "tedarik krizi" OR daralma OR zafiyet OR usulsüzlük OR yolsuzluk) (sanayi OR teknoloji OR üretim OR fabrika OR savunma OR otomotiv OR enerji OR şirket OR tesis OR proje) when:{when}']

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
        if dt and dt<cutoff: reasons['zaman']+=1; continue
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
        sentiment,score,status,neg,risk,cat=classify(title,snippet)
        out.append({
            'Tarih_dt':dt,'Tarih':fmt_dt(dt),'Başlık':title,'İçerik_Özeti':snippet or title,
            'URL':url,'Kaynak':src or d or 'Açık Kaynak','Domain':d,'Kaynak_Grubu':source_group(d),
            'Kategori':cat,'Duygu':sentiment,'Skor':score,'Risk_Durumu':status,
            'Negatif_Sinyaller':neg,'Risk_Sinyalleri':risk,'Seç':False,'Görsel_URL':'','_mode':mode
        })
    return out,reasons

def dedupe(rows):
    out=[]; urls=set(); titles=set()
    for r in rows:
        u=r['URL']; k=title_key(r['Başlık'])
        if u in urls or k in titles: continue
        urls.add(u); titles.add(k); out.append(r)
    out.sort(key=lambda x:(x['Tarih_dt'] is not None,x['Tarih_dt'] or datetime.min.replace(tzinfo=timezone.utc),source_rank(x['Domain'])),reverse=True)
    return out

# -----------------------------
# DOCX — sadece kullanıcı seçince
# Açık Kaynak Taraması rapor formatı
# -----------------------------
@st.cache_data(ttl=1800, show_spinner=False)
def article_detail(url):
    """Seçilen haber için gerçek sayfadan başlık, canonical URL, tarih, geniş metin ve görsel adaylarını çıkarır."""
    result = {
        'text': '', 'image': '', 'images': [], 'canonical': url,
        'title': '', 'published': '', 'source': ''
    }
    try:
        r = requests.get(url, headers=HEADERS, timeout=10, allow_redirects=True)
        if r.status_code != 200:
            return result

        result['canonical'] = r.url or url
        soup = BeautifulSoup(r.text, 'html.parser')

        # Canonical / başlık / tarih
        can = soup.find('link', rel=lambda x: x and 'canonical' in x.lower())
        if can and can.get('href'):
            result['canonical'] = requests.compat.urljoin(r.url, can.get('href'))

        og_title = soup.find('meta', attrs={'property':'og:title'})
        if og_title and og_title.get('content'):
            result['title'] = og_title['content'].strip()
        elif soup.title:
            result['title'] = soup.title.get_text(' ', strip=True)

        for prop in [
            ('meta', {'property':'article:published_time'}),
            ('meta', {'property':'article:modified_time'}),
            ('meta', {'name':'date'}),
            ('meta', {'name':'pubdate'}),
            ('meta', {'name':'publish-date'})
        ]:
            tag = soup.find(*prop)
            if tag and tag.get('content'):
                result['published'] = tag['content'].strip()
                break

        # JSON-LD: NewsArticle / Article içinden gövde ve görsel
        jsonld_bodies = []
        jsonld_images = []
        for tag in soup.find_all('script', attrs={'type': re.compile(r'application/ld\+json', re.I)}):
            raw = tag.string or tag.get_text()
            if not raw:
                continue
            try:
                data = json.loads(raw)
            except Exception:
                continue

            def walk(obj):
                if isinstance(obj, dict):
                    typ = str(obj.get('@type','')).lower()
                    if 'article' in typ or 'news' in typ:
                        if obj.get('articleBody'):
                            jsonld_bodies.append(str(obj['articleBody']))
                        if obj.get('headline') and not result['title']:
                            result['title'] = str(obj['headline'])
                        if obj.get('datePublished') and not result['published']:
                            result['published'] = str(obj['datePublished'])
                        im = obj.get('image') or obj.get('thumbnailUrl')
                        if isinstance(im, str): jsonld_images.append(im)
                        elif isinstance(im, list): jsonld_images.extend([str(x) for x in im if x])
                        elif isinstance(im, dict) and im.get('url'): jsonld_images.append(str(im['url']))
                    for v in obj.values():
                        walk(v)
                elif isinstance(obj, list):
                    for x in obj:
                        walk(x)
            walk(data)

        # Görsel adayları: OG/Twitter + JSON-LD + sayfadaki lazy/img/srcset
        image_candidates = []
        for attrs in [
            {'property':'og:image'}, {'property':'og:image:url'},
            {'name':'twitter:image'}, {'name':'twitter:image:src'}
        ]:
            tag = soup.find('meta', attrs=attrs)
            if tag and tag.get('content'):
                image_candidates.append(requests.compat.urljoin(r.url, tag['content'].strip()))

        image_candidates.extend([
            requests.compat.urljoin(r.url, x) for x in jsonld_images if x
        ])

        for img in soup.find_all('img'):
            for attr in ['src','data-src','data-lazy-src','data-original','data-image']:
                val = img.get(attr)
                if val:
                    image_candidates.append(requests.compat.urljoin(r.url, val))
            srcset = img.get('srcset') or img.get('data-srcset')
            if srcset:
                # En yüksek çözünürlüklü adayı sona yakın seç
                vals = [x.strip().split(' ')[0] for x in srcset.split(',') if x.strip()]
                image_candidates.extend([requests.compat.urljoin(r.url, x) for x in vals[-2:]])

        # Tekilleştir ve bariz logo/icon yollarını ele
        seen_img = set()
        clean_imgs = []
        for im in image_candidates:
            im = im.strip()
            if not im or im in seen_img:
                continue
            low = im.lower()
            if any(x in low for x in ['logo','favicon','icon','avatar','sprite']):
                continue
            seen_img.add(im)
            clean_imgs.append(im)
        result['images'] = clean_imgs[:20]
        result['image'] = clean_imgs[0] if clean_imgs else ''

        # Haber gövdesi: JSON-LD öncelikli, sonra article/main ve içerik sınıfları
        bodies = [x for x in jsonld_bodies if len(x.strip()) > 150]

        selectors = [
            'article',
            '[itemprop="articleBody"]',
            '[class*="article-body"]',
            '[class*="article-content"]',
            '[class*="news-content"]',
            '[class*="news-detail"]',
            '[class*="story-body"]',
            '[class*="post-content"]',
            '[class*="entry-content"]',
            'main'
        ]

        for selector in selectors:
            nodes = soup.select(selector)
            for node in nodes[:3]:
                parts = [
                    p.get_text(' ', strip=True)
                    for p in node.find_all(['p','h2','h3','li'])
                    if len(p.get_text(' ', strip=True)) >= 45
                ]
                if parts:
                    candidate = ' '.join(parts)
                    if len(candidate) > 250:
                        bodies.append(candidate)

        # Son çare: sayfadaki anlamlı paragraflar
        if not bodies:
            clone = BeautifulSoup(r.text, 'html.parser')
            for tag in clone(['script','style','nav','footer','header','aside','form','iframe','noscript','svg']):
                tag.decompose()
            parts = [
                p.get_text(' ', strip=True)
                for p in clone.find_all('p')
                if len(p.get_text(' ', strip=True)) >= 45
            ]
            bodies.append(' '.join(parts))

        # Temizle, tekrarları at
        seen = set()
        clean = []
        for text in bodies:
            text = re.sub(r'\s+', ' ', html.unescape(text)).strip()
            if len(text) < 120:
                continue
            key = norm(text[:500])
            if key in seen:
                continue
            seen.add(key)
            clean.append(text)

        clean.sort(key=len, reverse=True)
        result['text'] = ' '.join(clean[:3])[:20000]
        return result

    except Exception:
        return result


def broad_summary(title, body):
    """Kaynak metninden geniş, rapor diline uygun bir özet oluşturur."""
    body = re.sub(r'\s+', ' ', (body or '')).strip()
    if not body:
        return title

    # Haber metnini anlamlı cümlelere ayır; çok kısa/teknik olmayan parçaları at.
    sentences = [
        x.strip() for x in re.split(r'(?<=[.!?])\s+', body)
        if len(x.strip()) >= 45
    ]
    if not sentences:
        return body[:10000]

    # Kritik/negatif cümleleri kaybetmemek için öne al, geri kalanı kronolojik sırada koru.
    priority = [
        s for s in sentences
        if any(k in norm(s) for k in NEGATIVE_TERMS + HIGH_RISK_TERMS)
    ]
    normal = [s for s in sentences if s not in priority]

    ordered = []
    seen = set()
    for s in priority + normal:
        k = norm(s)
        if k in seen:
            continue
        seen.add(k)
        ordered.append(s)

    text = ' '.join(ordered)
    return text[:12000]


def download_image(url):
    if not url:
        return None
    try:
        r = requests.get(
            url, headers={**HEADERS, 'Accept':'image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8'},
            timeout=10
        )
        if r.status_code != 200 or len(r.content) < 1200:
            return None
        img = Image.open(BytesIO(r.content))
        if img.mode not in ('RGB','L'):
            img = img.convert('RGB')
        # Rapor görselleri için makul çözünürlük / dosya boyutu
        img.thumbnail((1600, 1200), Image.LANCZOS)
        b = BytesIO()
        img.save(b, 'JPEG', quality=88)
        b.seek(0)
        return b
    except Exception:
        return None


def add_link(paragraph, url, label='Haberi Aç'):
    """Word içinde tıklanabilir hyperlink."""
    try:
        rid = paragraph.part.relate_to(
            url,
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
            is_external=True
        )
        xml = (
            f'<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{rid}">'
            f'<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>{html.escape(label)}</w:t></w:r>'
            f'</w:hyperlink>'
        )
        paragraph._p.append(parse_xml(xml))
    except Exception:
        paragraph.add_run(url)


def _set_cell_shading(cell, fill):
    try:
        tcPr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="{fill}"/>')
        tcPr.append(shd)
    except Exception:
        pass


def make_docx(rows):
    """
    Kullanıcının yüklediği 'AÇIK KAYNAK TARAMA ÇALIŞMASI' raporunun
    düzenine yakın profesyonel bilgi notu:
    - başlık / görev alanı / tarih
    - Bulgular
    - her haber için numaralı madde
    - kaynak + başlık + tarih/saat
    - geniş içerik özeti
    - tıklanabilir haber linki
    - haber görseli ve görsel açıklaması
    """
    doc = Document()

    # Sayfa ayarları
    section = doc.sections[0]
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.85)
    section.right_margin = Inches(0.85)

    # Temel yazı tipi
    styles = doc.styles
    styles['Normal'].font.name = 'Arial'
    styles['Normal'].font.size = Pt(10.5)

    # Başlık
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run('AÇIK KAYNAK TARAMA ÇALIŞMASI')
    run.bold = True
    run.font.size = Pt(18)

    # Bilgi tablosu
    info = doc.add_table(rows=3, cols=2)
    info.style = 'Table Grid'
    metadata = [
        ('Tarama Yapılan Görev Alanı', 'Sanayi ve Teknoloji'),
        ('Tarih', datetime.now().astimezone().strftime('%d.%m.%Y')),
        ('Rapor Saati', datetime.now().astimezone().strftime('%H:%M:%S')),
    ]
    for i, (k, v) in enumerate(metadata):
        info.cell(i, 0).text = k
        info.cell(i, 1).text = v
        _set_cell_shading(info.cell(i, 0), 'E7E6E6')
        for rr in info.cell(i, 0).paragraphs[0].runs:
            rr.bold = True

    doc.add_paragraph()

    h = doc.add_paragraph()
    rr = h.add_run('Bulgular:')
    rr.bold = True
    rr.font.size = Pt(12)

    intro = doc.add_paragraph(
        'Sanayi ve Teknoloji alanlarında yapılan açık kaynak taraması neticesinde, '
        f'seçilen {len(rows)} haber/İçeriğe ilişkin bulgular aşağıda sunulmuştur.'
    )
    intro.paragraph_format.space_after = Pt(10)

    for i, row in enumerate(rows, 1):
        detail = article_detail(row.get('URL',''))
        real_url = detail.get('canonical') or row.get('URL','')
        title_text = detail.get('title') or row.get('Başlık') or 'Başlıksız Haber'
        body = detail.get('text') or row.get('İçerik_Özeti') or title_text
        summary = broad_summary(title_text, body)

        # Haber başlığı
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(f'{i}. “{title_text}”')
        r.bold = True
        r.font.size = Pt(11.5)

        # Kaynak / tarih / sınıflandırma
        meta = doc.add_paragraph()
        meta.add_run('Kaynak: ').bold = True
        meta.add_run(str(row.get('Kaynak','Açık Kaynak')))
        meta.add_run(' | Tarih/Saat: ').bold = True
        meta.add_run(str(row.get('Tarih','')))
        meta.add_run(' | Kategori: ').bold = True
        meta.add_run(str(row.get('Kategori','')))
        meta.add_run(' | Durum: ').bold = True
        meta.add_run(f"{row.get('Duygu','')} / {row.get('Risk_Durumu','')}")

        # Rapor örneğindeki anlatım biçimine yakın geniş özet
        content = doc.add_paragraph()
        content.add_run('Söz konusu haber içeriğinde, ').bold = True
        content.add_run(summary)

        # Link
        linkp = doc.add_paragraph()
        linkp.add_run('Haber Linki: ').bold = True
        add_link(linkp, real_url, 'Haberi Aç')
        raw = doc.add_paragraph()
        raw.add_run('Kaynak URL: ').bold = True
        raw.add_run(real_url)

        # Görsel
        image_stream = None
        selected_image_url = ''
        for image_url in detail.get('images', []):
            image_stream = download_image(image_url)
            if image_stream:
                selected_image_url = image_url
                break

        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap_run = cap.add_run(f'Görsel {i}: “{row.get("Kaynak","Kaynak")}” Sitesinde Yer Alan İçerik')
        cap_run.bold = True

        if image_stream:
            try:
                ip = doc.add_paragraph()
                ip.alignment = WD_ALIGN_PARAGRAPH.CENTER
                ip.add_run().add_picture(image_stream, width=Inches(5.9))

                gp = doc.add_paragraph()
                gp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                gp.add_run('Görsel kaynağı: ').bold = True
                add_link(gp, selected_image_url, 'Görseli Aç')
            except Exception:
                fp = doc.add_paragraph()
                fp.add_run('Görsel bağlantısı: ').bold = True
                add_link(fp, selected_image_url, 'Görseli Aç')
        else:
            fp = doc.add_paragraph()
            fp.add_run('Görsel: ').bold = True
            fp.add_run('Haber sayfasından Word içerisine gömülebilir bir görsel alınamadı. ')
            if detail.get('image'):
                add_link(fp, detail.get('image'), 'Görsel bağlantısını aç')
            else:
                fp.add_run('Kaynak sayfada erişilebilir görsel bağlantısı tespit edilemedi.')

        # Haberler arasında belirgin ayraç
        sep = doc.add_paragraph()
        sep.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sep.add_run('─' * 75)

    # Son not
    end = doc.add_paragraph()
    end.add_run('Arz olunur.').bold = True

    b = BytesIO()
    doc.save(b)
    b.seek(0)
    return b.getvalue()

# -----------------------------
# UI
# -----------------------------
st.title('🛡️ Sanayi & Teknoloji Açık Kaynak / Negatif Haber Radarı')
st.caption('Hızlı ilk bakış · geniş sanayi/teknoloji evreni · kronolojik saat/tarih · negatif/yüksek risk ayrımı · Türk medya önceliği · Yunan/Türk savunma · seçilen haberlerden DOCX')
with st.sidebar:
    st.header('⚙️ Tarama Ayarları')
    default=('sanayi OR teknoloji OR üretim OR imalat OR fabrika OR OSB OR makine OR otomasyon OR robotik OR Ar-Ge OR patent OR yapay zeka OR yazılım OR siber güvenlik OR çip OR yarı iletken OR elektronik OR telekom OR kuantum OR biyoteknoloji OR nanoteknoloji OR savunma sanayii OR ASELSAN OR TUSAŞ OR ROKETSAN OR HAVELSAN OR Baykar OR İHA OR SİHA OR KAAN OR havacılık OR uzay OR uydu OR otomotiv OR TOGG OR batarya OR enerji OR hidrojen OR kimya OR petrokimya OR demir çelik OR madencilik OR tekstil OR gıda teknolojisi OR tarım teknolojisi OR lojistik OR tedarik zinciri OR TÜBİTAK OR KOSGEB OR teknopark OR yatırım teşvik OR yerlileştirme')
    query=st.text_area('Geniş sanayi / teknoloji sorgusu:',default,height=190)
    neg=st.checkbox('⚠️ Negatif haberleri ayrıca tespit et',True)
    greek=st.checkbox('🇬🇷 Yunan medyası — yalnızca Türk savunma sanayii',True)
    social=st.checkbox('📱 Türk açık sosyal / indeks kaynakları',True)
    global_on=st.checkbox('🌍 Global basın (opsiyonel)',False)
    period=st.selectbox('🕒 Haber dönemi',['⚡ Son 3 saat','📅 Son 24 saat','📆 Son 48 saat','📆 Son 1 hafta','🗓️ Son 1 ay'],index=1)
    hours={'⚡ Son 3 saat':3,'📅 Son 24 saat':24,'📆 Son 48 saat':48,'📆 Son 1 hafta':168,'🗓️ Son 1 ay':720}[period]
    run=st.button('🔍 TARAMAYI BAŞLAT / YENİLE',type='primary',use_container_width=True)

if 'rows' not in st.session_state: st.session_state.rows=None
if 'scan_time' not in st.session_state: st.session_state.scan_time=None
if 'stats' not in st.session_state: st.session_state.stats={}

if run:
    cutoff=datetime.now(timezone.utc)-timedelta(hours=hours); when=period_window(hours)
    batches=[('🇹🇷 Türk medya / sanayi-teknoloji',build_turkish_queries(when,query),'turkish')]
    if neg: batches.append(('⚠️ Negatif haber taraması',build_negative_queries(when),'negative'))
    if greek: batches.append(('🇬🇷 Yunan medyası / Türk savunma',build_greek_queries(when),'greek'))
    if social: batches.append(('📱 Açık sosyal / indeks',build_social_queries(when),'social'))
    if global_on: batches.append(('🌍 Global basın',[
        f'(Turkey OR Türkiye) (industry OR manufacturing OR technology OR semiconductor OR defense OR aerospace OR automotive) timespan:{when}',
        f'(Turkey OR Turkish) (Baykar OR ASELSAN OR TUSAŞ OR ROKETSAN OR HAVELSAN OR KAAN OR drone OR missile) timespan:{when}'
    ],'global'))

    # İlk batch tek başına çalışır: kullanıcı çok kısa sürede ilk haberleri görür.
    all_rows=[]; stat={'Ham sonuç':0,'Zaman dışı':0,'Konu dışı':0,'Yunan dışı':0,'Kaynak dışı':0,'Sonuç':0}
    placeholder=st.empty(); status_box=st.status('🔎 Tarama başlıyor...',expanded=True)
    for idx,(label,queries,mode) in enumerate(batches):
        status_box.write(f'{label} — {len(queries)} paralel arama')
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(8,len(queries))) as ex:
            fs=[ex.submit(rss,q) for q in queries]
            raw=[]
            for f in concurrent.futures.as_completed(fs):
                try: raw.extend(f.result())
                except: pass
        stat['Ham sonuç']+=len(raw)
        norm_rows,reasons=normalize_rows(raw,cutoff,mode,query)
        stat['Zaman dışı']+=reasons['zaman']; stat['Konu dışı']+=reasons['konu']; stat['Yunan dışı']+=reasons['yunan']; stat['Kaynak dışı']+=reasons['kaynak']
        all_rows=dedupe(all_rows+norm_rows); stat['Sonuç']=len(all_rows)
        # İlk bakış: her batch sonrası anında güncellenir.
        if all_rows:
            pv=pd.DataFrame(all_rows)
            pv=pv.sort_values(['Tarih_dt','Domain'],ascending=[False,True],na_position='last')
            show=pv[['Tarih','Kaynak_Grubu','Kaynak','Başlık','İçerik_Özeti','Duygu','Risk_Durumu','URL']]
            placeholder.dataframe(show,column_config={'URL':st.column_config.LinkColumn('Haber Linki'),'İçerik_Özeti':st.column_config.TextColumn('İçerik / Özet',width='large')},hide_index=True,use_container_width=True,height=520)
        status_box.update(label=f'✅ {label} tamamlandı — toplam {len(all_rows)} haber',state='complete' if idx==len(batches)-1 else 'running')
    st.session_state.rows=all_rows; st.session_state.scan_time=datetime.now().astimezone(); st.session_state.stats=stat

rows=st.session_state.rows
if rows is None:
    st.info('👋 Hazır. Tarama başlamaz. Bir zaman aralığı seçip **TARAMAYI BAŞLAT / YENİLE** düğmesine basın.')
else:
    df=pd.DataFrame(rows)
    if not df.empty:
        df=df.sort_values('Tarih_dt',ascending=False,na_position='last').reset_index(drop=True)
        st.session_state.rows=df.to_dict('records')
    st.caption(f'Son tarama: {st.session_state.scan_time.strftime("%d.%m.%Y %H:%M:%S") if st.session_state.scan_time else "-"}')
    with st.expander('🧪 Tarama teşhisi',False): st.json(st.session_state.stats)
    if df.empty:
        st.warning('Sonuç bulunamadı. Tarama teşhisini açarak hangi aşamada sonuçların azaldığını görebilirsiniz.')
    else:
        total=len(df); negc=int((df.Duygu=='Negatif').sum()); riskc=int((df.Risk_Durumu=='Yüksek Risk').sum()); trc=int(df.Kaynak_Grubu.astype(str).str.startswith('🇹🇷').sum()); grc=int(df.Kaynak_Grubu.astype(str).str.startswith('🇬🇷').sum())
        a,b,c,d,e=st.columns(5); a.metric('Toplam',total); b.metric('Negatif',negc); c.metric('Yüksek Risk',riskc); d.metric('🇹🇷 Türk',trc); e.metric('🇬🇷 Yunan / Türk Savunma',grc)
        tabs=st.tabs([f'📰 Kronolojik ({total})',f'⚠️ Negatif ({negc})',f'🚨 Yüksek Risk ({riskc})',f'🇹🇷 Türk Medyası ({trc})',f'🇬🇷 Yunan ({grc})'])
        cols=['Seç','Tarih','Kaynak_Grubu','Kaynak','Kategori','Başlık','İçerik_Özeti','Duygu','Risk_Durumu','URL']
        with tabs[0]:
            edited=st.data_editor(df[cols],column_config={'Seç':st.column_config.CheckboxColumn('Seç'),'URL':st.column_config.LinkColumn('Haber Linki'),'İçerik_Özeti':st.column_config.TextColumn('İçerik / Özet',width='large')},disabled=[x for x in cols if x!='Seç'],hide_index=True,use_container_width=True,height=650,key=f'editor_{st.session_state.scan_time}')
            df.loc[edited.index,'Seç']=edited['Seç']; st.session_state.rows=df.to_dict('records')
        with tabs[1]:
            st.dataframe(df[df.Duygu=='Negatif'][['Tarih','Kaynak_Grubu','Kaynak','Kategori','Başlık','İçerik_Özeti','Risk_Durumu','URL']],column_config={'URL':st.column_config.LinkColumn('Haber Linki')},hide_index=True,use_container_width=True,height=650)
        with tabs[2]:
            st.dataframe(df[df.Risk_Durumu=='Yüksek Risk'][['Tarih','Kaynak_Grubu','Kaynak','Kategori','Başlık','İçerik_Özeti','Risk_Durumu','URL']],column_config={'URL':st.column_config.LinkColumn('Haber Linki')},hide_index=True,use_container_width=True,height=650)
        with tabs[3]:
            st.dataframe(df[df.Kaynak_Grubu.astype(str).str.startswith('🇹🇷')][['Tarih','Kaynak','Kategori','Başlık','İçerik_Özeti','Duygu','Risk_Durumu','URL']],column_config={'URL':st.column_config.LinkColumn('Haber Linki')},hide_index=True,use_container_width=True,height=650)
        with tabs[4]:
            st.dataframe(df[df.Kaynak_Grubu.astype(str).str.startswith('🇬🇷')][['Tarih','Kaynak','Kategori','Başlık','İçerik_Özeti','Duygu','Risk_Durumu','URL']],column_config={'URL':st.column_config.LinkColumn('Haber Linki')},hide_index=True,use_container_width=True,height=650)

        st.markdown('---'); st.subheader('📝 Seçilen haberlerden DOCX')
        selected=df[df['Seç']==True]
        st.write(f'{len(selected)} haber seçildi. Tarama sırasında görsel/tam metin indirilmez; sadece burada, sizin seçtiğiniz haberler için indirilir.')
        if st.button('📝 SEÇİLİ HABERLERLE WORD OLUŞTUR',type='primary',use_container_width=True):
            if selected.empty: st.warning('Önce haber tablosundan en az bir haber seçin.')
            else:
                with st.spinner(f'{len(selected)} seçili haber için görsel ve geniş özet hazırlanıyor...'):
                    st.session_state.docx_bytes=make_docx(selected.to_dict('records'))
        if st.session_state.get('docx_bytes'):
            st.download_button('⬇️ DOCX İNDİR',st.session_state.docx_bytes,file_name=f'Sanayi_Teknoloji_Bilgi_Notu_{date.today()}.docx',mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',type='primary',use_container_width=True)

st.caption('İlk açılışta otomatik tarama yoktur. Her TARAMAYI BAŞLAT / YENİLE yeni ağ taraması yapar. Haberler önce hızlı başlık/özet/saat/tarih/link olarak gelir; en yeni → en eski sıralanır. Negatif ve yüksek risk ayrı görünür. DOCX yalnızca sizin seçtiğiniz haberler için oluşturulur; görsel, haber sayfası ve geniş özet DOCX aşamasında alınır.')