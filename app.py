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
        sentiment,score,status,neg,risk,cat,reasons=classify(title,snippet,d)
        out.append({
            'Tarih_dt':dt,'Tarih':fmt_dt(dt),'Başlık':title,'İçerik_Özeti':snippet or title,
            'URL':url,'RSS_URL':url,'Kaynak':(src if norm(src) not in {'google haberler','google news','google'} else (d or src or 'Açık Kaynak')),
            'Yayıncı_URL':(r.get('source_url') or '').strip(),'Yayıncı':src or d or 'Açık Kaynak',
            'Domain':d,'Kaynak_Grubu':source_group(d),
            'Kategori':cat,'Duygu':sentiment,'Skor':score,'Risk_Skoru':score,'Risk_Durumu':status,
            'Risk_Gerekçesi':'; '.join(reasons),'Negatif_Sinyaller':neg,'Risk_Sinyalleri':risk,
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

def verification_status(row, all_rows=None):
    d=domain(row.get('Domain','')); source_group_name=str(row.get('Kaynak_Grubu',''))
    if d in TR_OFFICIAL: return '🟢 Resmî açıklama / birincil kaynak'
    if d in SOCIAL: return '🟠 Sosyal medya / tek kaynak'
    title=norm(row.get('Başlık',''))
    if all_rows:
        matches=[r for r in all_rows if r is not row and title_similarity(title,norm(r.get('Başlık',''))) >= 0.72]
        if len(matches)>=2: return '🟢 Çoklu kaynakla destekleniyor'
    if d in TR_MAIN or d in TR_TECH or d in GR: return '🟡 Tek medya kaynağı'
    return '🟡 Tek/açık kaynak'

def title_similarity(a,b):
    import difflib
    a=norm(a); b=norm(b)
    if not a or not b: return 0.0
    return difflib.SequenceMatcher(None,a,b).ratio()

def dedupe(rows):
    out=[]; urls=set(); titles=set()
    for r in rows:
        u=r['URL']; k=title_key(r['Başlık'])
        if u in urls or k in titles: continue
        urls.add(u); titles.add(k); out.append(r)
    out.sort(key=lambda x:(x['Tarih_dt'] is not None,x['Tarih_dt'] or datetime.min.replace(tzinfo=timezone.utc),source_rank(x['Domain'])),reverse=True)
    return out

def enrich_rows(rows):
    # Haber satırlarını analitik katmanla zenginleştirir; ağ çağrısı yapmaz.
    for r in rows:
        sentiment,score,status,neg,risk,cat,reasons=classify(r.get('Başlık',''),r.get('İçerik_Özeti',''),r.get('Domain',''))
        r['Duygu']=sentiment; r['Risk_Skoru']=score; r['Risk_Durumu']=status
        r['Negatif_Sinyaller']=neg; r['Risk_Sinyalleri']=risk; r['Risk_Gerekçesi']='; '.join(reasons)
        r['Kaynak_Güvenilirliği']=source_reliability(r.get('Domain',''),r.get('Kaynak',''))
    # Olay kümeleri: benzer başlıklar tek olay altında gruplanır.
    event_ids=[]
    for i,r in enumerate(rows):
        assigned=None
        for j,other in enumerate(rows[:i]):
            if r.get('Domain')==other.get('Domain') and norm(r.get('Başlık'))==norm(other.get('Başlık')):
                assigned=other.get('Olay_ID'); break
            if title_similarity(r.get('Başlık',''),other.get('Başlık',''))>=0.76:
                assigned=other.get('Olay_ID'); break
        if not assigned:
            assigned=f'OLAY-{len(event_ids)+1:03d}'
            event_ids.append(assigned)
        r['Olay_ID']=assigned
    # Her olay için kaynak ve zaman sayısı
    groups={}
    for r in rows: groups.setdefault(r['Olay_ID'],[]).append(r)
    for r in rows:
        g=groups.get(r['Olay_ID'],[])
        r['Olay_Kaynak_Sayisi']=len({x.get('Domain') for x in g if x.get('Domain')})
        times=[x.get('Tarih_dt') for x in g if x.get('Tarih_dt')]
        r['Olay_İlk_Görülme']=fmt_dt(min(times)) if times else r.get('Tarih','')
        r['Olay_Son_Görülme']=fmt_dt(max(times)) if times else r.get('Tarih','')
        r['Doğrulama']=verification_status(r,rows)
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

def make_analyst_docx(df, title='BİLGİ NOTU'):
    doc=Document(); sec=doc.sections[0]
    sec.top_margin=Cm(2); sec.bottom_margin=Cm(2); sec.left_margin=Cm(2.5); sec.right_margin=Cm(2.5)
    doc.styles['Normal'].font.name='Times New Roman'; doc.styles['Normal'].font.size=Pt(11)
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; r=p.add_run(title); r.bold=True; r.font.size=Pt(14)
    p=doc.add_paragraph(); p.add_run('Tarih/Saat: ').bold=True; p.add_run(datetime.now().astimezone().strftime('%d.%m.%Y %H:%M:%S'))
    events=build_event_summary(df)
    p=doc.add_paragraph(); p.add_run('Genel değerlendirme: ').bold=True; p.add_run(f'{len(df)} haber, {len(events)} olay kümesi incelenmiştir.')
    if not events.empty:
        p=doc.add_paragraph(); p.add_run('Öncelikli olaylar').bold=True
        for _,e in events.head(10).iterrows():
            p=doc.add_paragraph(style=None); p.paragraph_format.left_indent=Cm(.5)
            p.add_run(f"{e['Olay_ID']} — {e['Öne Çıkan Başlık']} | Risk {e['Risk']}/100 | {e['Haber Sayısı']} haber / {e['Kaynak Sayısı']} kaynak").bold=True
            p.add_run(f"\nKategori: {e['Kategori']}\nKaynaklar: {e['Kaynaklar']}")
            members=df[df['Olay_ID']==e['Olay_ID']].sort_values('Tarih_dt',ascending=True)
            for _,m in members.iterrows():
                p=doc.add_paragraph(); p.paragraph_format.left_indent=Cm(1)
                p.add_run(f"{m.get('Tarih','')} — {m.get('Kaynak','')}: ").bold=True
                p.add_run(str(m.get('Başlık',''))+' — '+str(m.get('Risk_Gerekçesi','')))
                if m.get('URL'): _word_hyperlink(p,m['URL'],'Haber linki')
    doc.add_paragraph('Risk değerlendirmesi, açık kaynak sinyallerine dayalı analitik bir önceliklendirmedir; tek başına doğrulama anlamına gelmez.')
    bio=BytesIO(); doc.save(bio); bio.seek(0); return bio.getvalue()

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
    period=st.selectbox('🕒 Haber dönemi',['⚡ Son 3 saat','📅 Son 24 saat','📆 Son 48 saat','📆 Son 1 hafta','🗓️ Son 1 ay'],index=1)
    hours={'⚡ Son 3 saat':3,'📅 Son 24 saat':24,'📆 Son 48 saat':48,'📆 Son 1 hafta':168,'🗓️ Son 1 ay':720}[period]
    run=st.button('🔍 TARAMAYI BAŞLAT / YENİLE',type='primary',use_container_width=True)

if 'rows' not in st.session_state: st.session_state.rows=None
if 'scan_time' not in st.session_state: st.session_state.scan_time=None
if 'stats' not in st.session_state: st.session_state.stats={}
if 'docx_bytes' not in st.session_state: st.session_state.docx_bytes=None
if 'note_bytes' not in st.session_state: st.session_state.note_bytes=None

if run:
    cutoff=datetime.now(timezone.utc)-timedelta(hours=hours)
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
    placeholder=st.empty(); status_box=st.status('🔎 Tarama başlıyor...',expanded=True)
    for idx,(label,queries,mode) in enumerate(batches):
        status_box.write(f'{label} — {len(queries)} paralel arama')
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(8,len(queries))) as ex:
            fs=[ex.submit(rss,q) for q in queries]; raw=[]
            for f in concurrent.futures.as_completed(fs):
                try: raw.extend(f.result())
                except: pass
        stat['Ham sonuç']+=len(raw)
        norm_rows,reasons=normalize_rows(raw,cutoff,mode,query)
        stat['Zaman dışı']+=reasons['zaman']; stat['Konu dışı']+=reasons['konu']; stat['Yunan dışı']+=reasons['yunan']; stat['Kaynak dışı']+=reasons['kaynak']
        all_rows=enrich_rows(dedupe(all_rows+norm_rows)); stat['Sonuç']=len(all_rows); stat['Olay']=len(set(r.get('Olay_ID') for r in all_rows))
        if all_rows:
            pv=pd.DataFrame(all_rows).sort_values(['Tarih_dt','Domain'],ascending=[False,True],na_position='last')
            show=pv[['Tarih','Kaynak_Grubu','Kaynak','Başlık','İçerik_Özeti','Duygu','Risk_Skoru','Risk_Durumu','URL']]
            placeholder.dataframe(show,column_config={'URL':st.column_config.LinkColumn('Haber Linki'),'İçerik_Özeti':st.column_config.TextColumn('İçerik / Özet',width='large'),'Risk_Skoru':st.column_config.NumberColumn('Risk',format='%d/100')},hide_index=True,use_container_width=True,height=520)
        status_box.update(label=f'✅ {label} tamamlandı — toplam {len(all_rows)} haber / {len(set(r.get("Olay_ID") for r in all_rows))} olay',state='complete' if idx==len(batches)-1 else 'running')
    st.session_state.rows=all_rows; st.session_state.scan_time=datetime.now().astimezone(); st.session_state.stats=stat

rows=st.session_state.rows
if rows is None:
    st.info('👋 Hazır. Tarama başlamaz. Zaman aralığını seçip **TARAMAYI BAŞLAT / YENİLE** düğmesine basın.')
else:
    df=pd.DataFrame(rows)
    if not df.empty:
        df=enrich_rows(df.to_dict('records')); df=pd.DataFrame(df).sort_values('Tarih_dt',ascending=False,na_position='last').reset_index(drop=True)
        st.session_state.rows=df.to_dict('records')
    st.caption(f'Son tarama: {st.session_state.scan_time.strftime("%d.%m.%Y %H:%M:%S") if st.session_state.scan_time else "-"}')
    with st.expander('🧪 Tarama teşhisi',False): st.json(st.session_state.stats)
    if df.empty:
        st.warning('Sonuç bulunamadı. Tarama teşhisini açarak hangi aşamada sonuçların azaldığını görebilirsiniz.')
    else:
        total=len(df); negc=int((df.Duygu=='Negatif').sum()); riskc=int((df.Risk_Durumu=='Yüksek Risk').sum()); trc=int(df.Kaynak_Grubu.astype(str).str.startswith('🇹🇷').sum()); grc=int(df.Kaynak_Grubu.astype(str).str.startswith('🇬🇷').sum()); events=df['Olay_ID'].nunique()
        a,b,c,d,e,f=st.columns(6); a.metric('Toplam',total); b.metric('Olay',events); c.metric('Negatif',negc); d.metric('Yüksek Risk',riskc); e.metric('🇹🇷 Türk',trc); f.metric('🇬🇷 Yunan',grc)

        # Alarm bandı
        alarms=df[(df.Risk_Skoru>=70) | (df.Duygu=='Negatif')].sort_values(['Risk_Skoru','Tarih_dt'],ascending=[False,False])
        if not alarms.empty:
            st.subheader('🚨 Yeni / Öncelikli Alarmlar')
            for _,r in alarms.head(8).iterrows():
                icon='🔴' if int(r['Risk_Skoru'])>=70 else '🟠'
                st.markdown(f"{icon} **{r['Tarih']} — {r['Başlık']}** — **{int(r['Risk_Skoru'])}/100**  \\n**{r['Kaynak']}** · {r['Kategori']} · {r['Risk_Gerekçesi']} · {r['Doğrulama']}")

        tabs=st.tabs([f'📰 Kronolojik ({total})',f'⚠️ Negatif ({negc})',f'🚨 Yüksek Risk ({riskc})',f'🇹🇷 Türk ({trc})',f'🇬🇷 Yunan ({grc})',f'🧩 Olaylar ({events})', '📈 Trend / Analiz', '⭐ Takip Listesi'])
        cols=['Seç','Tarih','Kaynak_Grubu','Kaynak','Kategori','Başlık','İçerik_Özeti','Duygu','Risk_Skoru','Risk_Durumu','Kaynak_Güvenilirliği','Doğrulama','URL']
        with tabs[0]:
            edited=st.data_editor(df[cols],column_config={'Seç':st.column_config.CheckboxColumn('Seç'),'URL':st.column_config.LinkColumn('Haber Linki'),'İçerik_Özeti':st.column_config.TextColumn('İçerik / Özet',width='large'),'Risk_Skoru':st.column_config.NumberColumn('Risk',format='%d/100')},disabled=[x for x in cols if x!='Seç'],hide_index=True,use_container_width=True,height=650,key=f'editor_{st.session_state.scan_time}')
            df.loc[edited.index,'Seç']=edited['Seç']; st.session_state.rows=df.to_dict('records')
        with tabs[1]:
            st.dataframe(df[df.Duygu=='Negatif'][['Tarih','Kaynak','Kategori','Başlık','Risk_Skoru','Risk_Gerekçesi','Doğrulama','URL']],column_config={'URL':st.column_config.LinkColumn('Haber Linki')},hide_index=True,use_container_width=True,height=650)
        with tabs[2]:
            st.dataframe(df[df.Risk_Durumu=='Yüksek Risk'][['Tarih','Kaynak','Kategori','Başlık','Risk_Skoru','Risk_Gerekçesi','Doğrulama','URL']],column_config={'URL':st.column_config.LinkColumn('Haber Linki')},hide_index=True,use_container_width=True,height=650)
        with tabs[3]:
            st.dataframe(df[df.Kaynak_Grubu.astype(str).str.startswith('🇹🇷')][['Tarih','Kaynak','Kategori','Başlık','Risk_Skoru','Duygu','URL']],column_config={'URL':st.column_config.LinkColumn('Haber Linki')},hide_index=True,use_container_width=True,height=650)
        with tabs[4]:
            st.dataframe(df[df.Kaynak_Grubu.astype(str).str.startswith('🇬🇷')][['Tarih','Kaynak','Kategori','Başlık','Risk_Skoru','Duygu','URL']],column_config={'URL':st.column_config.LinkColumn('Haber Linki')},hide_index=True,use_container_width=True,height=650)
        with tabs[5]:
            ev=build_event_summary(df); st.dataframe(ev,hide_index=True,use_container_width=True,height=500)
            chosen=st.selectbox('Olay zaman çizelgesini göster:',ev['Olay_ID'].tolist() if not ev.empty else [])
            if chosen:
                g=df[df.Olay_ID==chosen].sort_values('Tarih_dt',ascending=True)
                for _,r in g.iterrows(): st.markdown(f"**{r['Tarih']}** → **{r['Kaynak']}** — {r['Başlık']} — {r['Doğrulama']}")
        with tabs[6]:
            st.subheader('📊 Konu yoğunluğu')
            tr=trend_table(df)
            if not tr.empty: st.bar_chart(tr.set_index('Kategori')['Haber'])
            st.subheader('📈 Gündem yoğunluğu')
            tmp=df[df['Tarih_dt'].notna()].copy(); tmp['Saat']=tmp['Tarih_dt'].dt.strftime('%Y-%m-%d %H:00')
            if not tmp.empty: st.line_chart(tmp.groupby('Saat').size())
            st.subheader('🧭 En hızlı yükselen / yoğun konular')
            for _,r in tr.head(10).iterrows(): st.write(f"**{r['Kategori']}** — {int(r['Haber'])} haber")
        with tabs[7]:
            hits=watchlist_hits(df,watch)
            st.write(f'Listede eşleşen: **{len(hits)}** haber')
            if not hits.empty: st.dataframe(hits[['Tarih','Kaynak','Kategori','Başlık','Risk_Skoru','Duygu','URL']],column_config={'URL':st.column_config.LinkColumn('Haber Linki')},hide_index=True,use_container_width=True,height=550)

        st.markdown('---'); st.subheader('📝 Seçilen haberlerden çıktı üret')
        selected=df[df['Seç']==True]
        st.write(f'{len(selected)} haber seçildi. Tarama sırasında görsel/tam metin indirilmez; yalnızca seçtiğiniz içerikler için derin zenginleştirme yapılır.')
        c1,c2=st.columns(2)
        with c1:
            if st.button('📝 AÇIK KAYNAK RAPORU / WORD',type='primary',use_container_width=True):
                if selected.empty: st.warning('Önce haber seçin.')
                else:
                    with st.spinner(f'{len(selected)} haber zenginleştiriliyor...'): st.session_state.docx_bytes=make_docx(selected.to_dict('records'))
            if st.session_state.docx_bytes: st.download_button('⬇️ Açık Kaynak Raporu DOCX',st.session_state.docx_bytes,file_name=f'Sanayi_Teknoloji_Acik_Kaynak_{date.today()}.docx',mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',use_container_width=True)
        with c2:
            if st.button('📌 BİLGİ NOTU / ANALİTİK WORD',use_container_width=True):
                if selected.empty: st.warning('Önce haber seçin.')
                else: st.session_state.note_bytes=make_analyst_docx(selected,title='SANAYİ & TEKNOLOJİ BİLGİ NOTU')
            if st.session_state.note_bytes: st.download_button('⬇️ Bilgi Notu DOCX',st.session_state.note_bytes,file_name=f'Sanayi_Teknoloji_Bilgi_Notu_{date.today()}.docx',mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',use_container_width=True)

st.caption('İlk açılışta otomatik tarama yoktur. Her yenileme yeni ağ taraması yapar. Haberler en yeni → en eski sıralanır; olay kümeleri, risk gerekçesi, kaynak güvenilirliği, doğrulama, trend ve takip listesi tarama sonucunda yer alır. DOCX aşamasında seçilen haberlerin gerçek yayıncı sayfası, görseli, linki ve geniş içeriği alınır.')