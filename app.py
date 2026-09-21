import streamlit as st
import pandas as pd
import requests
import concurrent.futures
import xml.etree.ElementTree as ET
import re, html, json, math
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
# V120 QA-STABLE — V104+ çekirdeği / performans ve rapor kalite katmanı
# 1) Aynı olayın daha güçlü tekilleştirilmesi
# 2) Durum bilgisinin URL'ye değil olay kimliğine de dayanması
# 3) "Dünden Beri Ne Değişti?" yalnız gerçek/maddi değişiklikler
# 4) Vardiya Başlangıç Özeti: seçici, 5–8 tekil gelişme
# ============================================================

_V104_ENTITY_HINTS = {
    'tüik','tcmb','epdk','tübitak','kosgeb','tse','türkpatent','ssb','aselsan','tusaş','tusas',
    'roketsan','havelsan','baykar','togg','mke','kaan','kızılelma','kizilelma','hisar','siper',
    'teknofest','tcg','anadolu','turksat','türksat','thk','thy','turkcell','türk telekom'
}
_V104_GENERIC_EVENT_WORDS = {
    'türkiye','türk','sanayi','teknoloji','haber','son','yeni','ilk','bugün','dün','açıklama',
    'açıkladı','belirtti','duyurdu','başladı','gerçekleşti','oldu','edildi','yapıldı','kapsamında',
    'milyon','milyar','bin','yüzde','oran','veri','verileri','program','proje'
}

def _v104_event_tokens(title, summary=''):
    """
    Olay eşleştirmesinde yalnız başlık kelimelerine bağlı kalmaz.
    Kurum/ürün/yer/özgül sayı ve eylem çekirdeğini kısa bir imzaya dönüştürür.
    """
    title_n=norm(title)
    summary_n=norm(summary)
    title_tokens=set(_title_tokens(title))
    # Özellikle başlıkta geçen ayırt edici kurum/ürün kelimelerini koru.
    entities={x for x in re.findall(r'[a-z0-9çğıöşü]+',title_n)
              if x in _V104_ENTITY_HINTS or (len(x)>=5 and x not in _V104_GENERIC_EVENT_WORDS)}
    nums=set(re.findall(r'\b\d+(?:[.,]\d+)?\b',title_n))
    # Başlık çok kısaysa özetten sınırlı destek al.
    if len(title_tokens)<4:
        extra=[x for x in _title_tokens(summary_n) if x not in _V104_GENERIC_EVENT_WORDS]
        title_tokens.update(extra[:5])
    return set(title_tokens)|entities|nums

def _v104_event_similarity(a_title,a_summary,b_title,b_summary):
    a=_v104_event_tokens(a_title,a_summary)
    b=_v104_event_tokens(b_title,b_summary)
    if not a or not b:
        return 0.0
    jac=len(a&b)/max(1,len(a|b))
    # Aynı kurum/ürün çekirdeği varsa eşleşmeyi destekle; tek başına yeterli sayma.
    ae={x for x in a if x in _V104_ENTITY_HINTS}
    be={x for x in b if x in _V104_ENTITY_HINTS}
    entity_bonus=0.12 if (ae&be) else 0.0
    # Aynı özgül sayı/tarih varsa küçük destek.
    an={x for x in a if re.fullmatch(r'\d+(?:[.,]\d+)?',x)}
    bn={x for x in b if re.fullmatch(r'\d+(?:[.,]\d+)?',x)}
    number_bonus=0.06 if (an&bn) else 0.0
    return min(1.0,jac+entity_bonus+number_bonus)

def _v104_event_representatives(df):
    """Mevcut Olay_ID'leri ikinci kez birleştirerek yanlış çoğalmayı azaltır."""
    if df is None or df.empty:
        return df
    x=df.copy()
    if 'Tarih_dt' in x.columns:
        x['Tarih_dt']=pd.to_datetime(x['Tarih_dt'],utc=True,errors='coerce')
    # Önce mevcut kümelerden temsilci çıkar.
    if 'Olay_ID' in x.columns:
        reps=[]
        for oid,g in x.groupby('Olay_ID',dropna=False):
            g=g.sort_values('Tarih_dt',ascending=False,na_position='last')
            r=g.iloc[0].copy()
            r['_v104_members']=list(g.index)
            r['_v104_summary']=' '.join(g.get('İçerik_Özeti',pd.Series(dtype=str)).fillna('').astype(str).head(5))
            reps.append(r)
        reps=pd.DataFrame(reps)
    else:
        reps=x.copy()
        reps['_v104_members']=[[i] for i in reps.index]
        reps['_v104_summary']=reps.get('İçerik_Özeti','')

    # Ters indeks: performansı korur.
    token_index={}
    clusters=[]
    for _,r in reps.sort_values('Tarih_dt',ascending=False,na_position='last').iterrows():
        toks=_v104_event_tokens(r.get('Başlık',''),r.get('_v104_summary',''))
        candidates=set()
        for t in toks:
            candidates.update(token_index.get(t,set()))
        best=None; best_sim=0.0
        for ci in candidates:
            cr=clusters[ci]['rep']
            sim=_v104_event_similarity(
                r.get('Başlık',''),r.get('_v104_summary',''),
                cr.get('Başlık',''),cr.get('_v104_summary','')
            )
            if sim>best_sim:
                best_sim=sim; best=ci
        threshold=0.52 if len(toks)>=6 else 0.60
        if best is None or best_sim<threshold:
            best=len(clusters)
            clusters.append({'rep':r,'members':list(r['_v104_members'])})
        else:
            clusters[best]['members'].extend(r['_v104_members'])
            # En yeni temsilciyi koru.
            try:
                if pd.to_datetime(r.get('Tarih_dt'),utc=True,errors='coerce') > pd.to_datetime(clusters[best]['rep'].get('Tarih_dt'),utc=True,errors='coerce'):
                    clusters[best]['rep']=r
            except Exception:
                pass
        for t in toks:
            token_index.setdefault(t,set()).add(best)

    rows=[]
    for ci,c in enumerate(clusters,1):
        g=x.loc[list(dict.fromkeys(c['members']))].copy()
        g=g.sort_values('Tarih_dt',ascending=False,na_position='last')
        rep=g.iloc[0].copy()
        rep['Olay_ID']=f'V104-{ci:04d}'
        rep['Olay_Haber_Sayisi']=len(g)
        domains={str(v) for v in g.get('Domain',pd.Series(dtype=str)).tolist() if str(v).strip()}
        rep['Olay_Kaynak_Sayisi']=max(len(domains),int(pd.to_numeric(g.get('Olay_Kaynak_Sayisi',0),errors='coerce').fillna(0).max() or 0))
        # Aynı olayın en yüksek risk/teyit bilgisini kaybetme.
        rep['Risk_Skoru']=int(pd.to_numeric(g.get('Risk_Skoru',0),errors='coerce').fillna(0).max() or 0)
        rows.append(rep)
    return pd.DataFrame(rows).drop(columns=['_v104_members','_v104_summary'],errors='ignore')

def _v104_event_status_sets():
    """
    Durumu hem URL hem başlık hem de olay-token imzasıyla indeksler.
    Aynı olay farklı kaynaktan görünse bile kullanıcı işlemi kaybolmaz.
    """
    cached=st.session_state.get('_v104_status_cache')
    if cached is not None:
        return cached
    result={'imp':set(),'akt':set(),'notes':set(),'pres':set()}
    if not _init_history_db():
        return result
    try:
        with _history_connect() as conn:
            for table,key in [
                ('important_basket','imp'),('osint_report_basket','akt'),
                ('note_history','notes'),('presentation_basket','pres')
            ]:
                for title,url in conn.execute(f"SELECT title,url FROM {table}").fetchall():
                    title=str(title or ''); url=str(url or '').strip()
                    if url: result[key].add('U:'+url)
                    tk=title_key(title)
                    if tk: result[key].add('T:'+tk)
                    sig=' '.join(sorted(_v104_event_tokens(title,'')))
                    if sig: result[key].add('E:'+sig)
    except Exception:
        pass
    st.session_state['_v104_status_cache']=result
    return result

def _v104_row_status_keys(r):
    url=str(r.get('URL',r.get('url','')) or '').strip()
    title=str(r.get('Başlık',r.get('title','')) or '')
    summary=str(r.get('İçerik_Özeti',r.get('summary','')) or '')
    keys=set()
    if url: keys.add('U:'+url)
    tk=title_key(title)
    if tk: keys.add('T:'+tk)
    sig=' '.join(sorted(_v104_event_tokens(title,summary)))
    if sig: keys.add('E:'+sig)
    return keys

def _v63_add_status_badges(df):
    """V104 — tüm tablolarda olay bazlı güvenilir Durum sütunu."""
    if df is None or df.empty:
        return df
    out=df.copy()
    sets=_v104_event_status_sets()
    def badge(r):
        keys=_v104_row_status_keys(r)
        b=[]
        if keys & sets['pres']:  b.append('🖥️ Sunum Sepetinde')
        if keys & sets['imp']:   b.append('📌 Önemli Gelişmelerde')
        if keys & sets['notes']: b.append('📝 Bilgi Notu Yapıldı')
        if keys & sets['akt']:   b.append('📁 AKT Sepetinde')
        return ' • '.join(b) if b else '—'
    out['Durum']=out.apply(badge,axis=1)
    return out

def _v73_invalidate_status_cache():
    st.session_state.pop('_v73_status_sets_cache',None)
    st.session_state.pop('_v104_status_cache',None)

def _v104_material_change(prev,cur):
    """Kaynak sayısındaki sıradan artışı tek başına 'yeni bilgi' saymaz."""
    prev_risk=int(prev.get('risk_score') or 0)
    cur_risk=int(cur.get('risk_score') or 0)
    risk_up=(cur_risk>=prev_risk+15) or (_risk_rank(cur.get('risk_status',''))>_risk_rank(prev.get('risk_status','')))
    verify_up=_verification_rank(cur.get('verification',''))>_verification_rank(prev.get('verification',''))

    prev_text=(prev.get('title') or '')+' '+(prev.get('summary') or '')
    cur_text=(cur.get('title') or '')+' '+(cur.get('summary') or '')
    prev_tokens=set(_history_tokens(prev_text))
    cur_tokens=set(_history_tokens(cur_text))
    new_tokens=cur_tokens-prev_tokens

    # Gerçek yeni bilgi için yalnız genel kelimeler değil, sayı/kurum/özgül içerik aranır.
    prev_nums=set(re.findall(r'\b\d+(?:[.,]\d+)?\b',prev_text))
    cur_nums=set(re.findall(r'\b\d+(?:[.,]\d+)?\b',cur_text))
    new_nums=cur_nums-prev_nums
    meaningful=[t for t in new_tokens if len(t)>=5 and t not in _V104_GENERIC_EVENT_WORDS]
    materially_updated=(len(meaningful)>=8) or (len(new_nums)>=1 and len(meaningful)>=3)

    return risk_up,verify_up,materially_updated,meaningful,new_nums

def _compare_since_previous(df,current_scan_id=None):
    """
    V104 — yalnız gerçek değişiklikler:
    yeni olay / maddi yeni bilgi / risk artışı / teyit güçlenmesi.
    Aynı olayın yeni bir sitede tekrar yayımlanması tek başına değişiklik değildir.
    """
    current=_v104_event_representatives(df)
    prev_id=_previous_scan_id(current_scan_id)
    previous=_load_scan_events(prev_id)
    if current is None or current.empty:
        return pd.DataFrame(),None,None
    if previous.empty:
        return pd.DataFrame(),prev_id,None

    prev_records=[]
    for _,p in previous.iterrows():
        rec=p.to_dict()
        rec['tokens']=_history_tokens((rec.get('title') or '')+' '+(rec.get('summary') or ''))
        prev_records.append(rec)

    changes=[]
    for _,r in current.iterrows():
        c={
            'title':str(r.get('Başlık','') or ''),'source':str(r.get('Kaynak','') or ''),
            'url':str(r.get('URL','') or ''),'category':str(r.get('Kategori','') or ''),
            'summary':str(r.get('İçerik_Özeti','') or ''),'risk_score':int(r.get('Risk_Skoru',0) or 0),
            'risk_status':str(r.get('Risk_Durumu','') or ''),'verification':str(r.get('Doğrulama','') or ''),
            'source_count':int(r.get('Olay_Kaynak_Sayisi',1) or 1)
        }
        best=None; best_sim=0.0
        for p in prev_records:
            sim=_v104_event_similarity(c['title'],c['summary'],p.get('title',''),p.get('summary',''))
            if c['url'] and c['url']==str(p.get('url','') or ''):
                sim=max(sim,0.98)
            if sim>best_sim:
                best_sim=sim; best=p

        if best is None or best_sim<0.50:
            changes.append({
                'Değişim':'🆕 YENİ OLAY','Başlık':c['title'],'Kaynak':c['source'],'Kategori':c['category'],
                'Risk':c['risk_score'],'Önceki Risk':'—','Kaynak Sayısı':c['source_count'],
                'Açıklama':'Önceki taramada aynı olaya ilişkin yeterli eşleşme bulunmamıştır.',
                'URL':c['url'],'_priority':100+c['risk_score']
            })
            continue

        risk_up,verify_up,material,new_words,new_nums=_v104_material_change(best,c)
        if risk_up:
            kind='⚠️ RİSK ARTTI'
            expl=f"Risk {int(best.get('risk_score') or 0)}/100 seviyesinden {c['risk_score']}/100 seviyesine yükselmiştir."
            priority=95+c['risk_score']
        elif verify_up:
            kind='✅ TEYİT GÜÇLENDİ'
            expl=f"Doğrulama seviyesi {best.get('verification','')} düzeyinden {c['verification']} düzeyine yükselmiştir."
            priority=90+c['risk_score']
        elif material:
            kind='🔄 YENİ BİLGİ'
            bits=[]
            if new_nums: bits.append('yeni sayısal veri: '+', '.join(sorted(new_nums)[:4]))
            if new_words: bits.append('yeni içerik: '+', '.join(sorted(new_words)[:6]))
            expl='; '.join(bits) if bits else 'Olay hakkında maddi yeni bilgi tespit edilmiştir.'
            priority=80+c['risk_score']
        else:
            continue

        changes.append({
            'Değişim':kind,'Başlık':c['title'],'Kaynak':c['source'],'Kategori':c['category'],
            'Risk':c['risk_score'],'Önceki Risk':int(best.get('risk_score') or 0),
            'Kaynak Sayısı':c['source_count'],'Açıklama':expl,'URL':c['url'],'_priority':priority
        })

    out=pd.DataFrame(changes)
    if not out.empty:
        # Aynı olayın değişiklik listesinde de yalnız bir kez görünmesi.
        out['_sig']=out.apply(lambda r:' '.join(sorted(_v104_event_tokens(r.get('Başlık',''),r.get('Açıklama','')))),axis=1)
        out=out.sort_values(['_priority','Risk'],ascending=[False,False]).drop_duplicates('_sig',keep='first')
        out=out.drop(columns=['_priority','_sig'],errors='ignore')
    prev_time=str(previous.iloc[0].get('scanned_at','')) if not previous.empty else None
    return out,prev_id,prev_time

def _v104_shift_priority(r):
    text=norm(f"{r.get('Başlık','')} {r.get('İçerik_Özeti','')} {r.get('Kategori','')}")
    official=_is_official_radar_row(r) or _verification_rank(r.get('Doğrulama',''))>=4
    data_terms=['tüik','tcmb','epdk','istatistik','veri','endeks','oran','kapasite kullanım',
                'sanayi üretimi','ihracat','ithalat','istihdam','milyar','milyon','yüzde','%']
    strategic_terms=['yatırım','üretim','tesis','fabrika','çip','yarı iletken','yapay zeka','yapay zekâ',
                     'siber','ar-ge','arge','patent','teşvik','kritik teknoloji','otomotiv','enerji']
    defence_terms=['savunma','aselsan','tusaş','tusas','roketsan','havelsan','baykar','kaan','kızılelma',
                   'füze','iha','siha','uzay','uydu','teknofest']
    has_data=any(x in text for x in data_terms) or bool(re.search(r'\d',text))
    strategic=any(x in text for x in strategic_terms)
    defence=any(x in text for x in defence_terms)
    critical=(int(r.get('Risk_Skoru',0) or 0)>=70 or r.get('Duygu')=='Negatif' or
              bool(critical_industrial_incident(r.get('Başlık',''),r.get('İçerik_Özeti',''))))
    verified=_verification_rank(r.get('Doğrulama',''))>=3

    # Kullanıcının istediği açık öncelik sırası.
    if official and has_data: tier=5
    elif official: tier=5
    elif strategic: tier=4
    elif critical: tier=3
    elif defence: tier=2
    elif verified: tier=1
    else: tier=0

    score=tier*100
    score+=min(int(r.get('Risk_Skoru',0) or 0),100)
    score+=min(int(r.get('Olay_Kaynak_Sayisi',0) or 0)*5,20)
    if has_data: score+=15
    return score,tier

def _shift_start_summary(df,current_scan_id=None):
    """
    V104 — sabah ilk bakış: 5–8 adet gerçekten önemli, olay bazında tekil gelişme.
    Öncelik: resmî veri/açıklama > stratejik sanayi-teknoloji > kritik negatif >
    savunma/uzay/teknoloji programı > yüksek teyitli yeni gelişme.
    """
    if df is None or df.empty:
        return {},pd.DataFrame(),""

    baseline,baseline_label,baseline_scan_id=_shift_baseline(current_scan_id)
    x=df.copy()
    x['Tarih_dt']=pd.to_datetime(x.get('Tarih_dt'),utc=True,errors='coerce')
    since=x[(x['Tarih_dt'].isna()) | (x['Tarih_dt']>=baseline)].copy() if baseline is not None else x.copy()

    changes,_,_=_compare_since_previous(df,current_scan_id)
    new_events=int(changes['Tür'].astype(str).str.contains('YENİ OLAY').sum()) if not changes.empty else 0
    risk_up=int(changes['Tür'].astype(str).str.contains('RİSK ARTTI').sum()) if not changes.empty else 0
    verify_up=int(changes['Tür'].astype(str).str.contains('TEYİT').sum()) if not changes.empty else 0

    reps=_v104_event_representatives(since) if not since.empty else pd.DataFrame()
    if not reps.empty:
        scored=reps.apply(_v104_shift_priority,axis=1)
        reps['_V104_Puan']=[v[0] for v in scored]
        reps['_V104_Kademe']=[v[1] for v in scored]
        # Önemsiz/generic içerik sabah özetini doldurmasın.
        eligible=reps[reps['_V104_Kademe']>0].copy()
        eligible=eligible.sort_values(['_V104_Puan','Tarih_dt'],ascending=[False,False],na_position='last')
        # En az 5 uygun olay varsa 5; güçlü aday çoksa en fazla 8.
        top_n=min(8,max(5,min(len(eligible),8))) if len(eligible)>=5 else len(eligible)
        top=eligible.head(top_n).drop(columns=['_V104_Puan','_V104_Kademe'],errors='ignore')
    else:
        top=pd.DataFrame()

    high=0; osb=0
    if not reps.empty:
        high=int((reps.get('Risk_Durumu',pd.Series(dtype=str))=='Yüksek Risk').sum())
        osb=sum(bool(critical_industrial_incident(r.get('Başlık',''),r.get('İçerik_Özeti',''))) for _,r in reps.iterrows())

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

# ============================================================
# /V104
# ============================================================

# V106 — V105 düzeltmesi:
# Yalnızca görünür 'Resmî Açıklama – Medya Karşılaştırması' paneli kaldırılmıştır.
# Sorgu/negatif tarama yardımcı fonksiyonları ve V104 çekirdeği korunmuştur.

# V107 — Sepete eklemede aynı olayın mevcut taramadaki kaynakları otomatik zenginleştirilir.
# Yerel/kısa sürüm seçilse bile resmî/ana akım/daha ayrıntılı sürüm ve kritik veriler birleştirilir.
# Ek ağ isteği yapılmaz; V106 kararlı çekirdeği korunur.

# V108 — V107 zenginleştirme TypeError düzeltmesi:
# _v107_unique_sentences içindeki hashlenemeyen set, frozenset olarak saklanmaktadır.
# V107 olay/kaynak zenginleştirme mantığı korunmuştur.


# ============================================================
# V109 — PANEL GELİŞTİRMELERİ
# ============================================================

def _v109_numbers(text):
    return set(re.findall(
        r'\b\d+(?:[.,]\d+)?(?:\s*(?:%|yüzde|milyon|milyar|bin|adet|tl|dolar|euro|avro|mw|gw|gwh|mwh|km))?',
        norm(text)
    ))

def _v109_sentences(text):
    out=[]
    for s in _sentence_chunks(_clean_note_text(text)):
        s=_clean_note_text(s).strip()
        if len(s)>=35:
            out.append(s)
    return out

def _v109_direct_difference(prev,cur,kind):
    prev_text=_clean_note_text((prev.get('title') or '')+' '+(prev.get('summary') or ''))
    cur_text=_clean_note_text((cur.get('title') or '')+' '+(cur.get('summary') or ''))

    if 'RİSK ARTTI' in kind:
        return (
            f"Önceki taramada risk düzeyi {int(prev.get('risk_score') or 0)}/100 iken, "
            f"yeni taramada {int(cur.get('risk_score') or 0)}/100 seviyesine yükselmiştir."
        )
    if 'TEYİT GÜÇLENDİ' in kind:
        return (
            f"Önceki taramada doğrulama düzeyi “{prev.get('verification','')}” iken, "
            f"yeni taramada “{cur.get('verification','')}” seviyesine yükselmiştir."
        )
    if 'YENİ OLAY' in kind:
        return 'Önceki karşılaştırma taramasında aynı olaya ilişkin yeterli eşleşme bulunmamaktadır; gelişme yeni olay olarak değerlendirilmiştir.'

    prev_nums=_v109_numbers(prev_text)
    cur_nums=_v109_numbers(cur_text)
    new_nums=[x for x in cur_nums if x not in prev_nums]
    prev_tok=set(_history_tokens(prev_text))

    candidates=[]
    for s in _v109_sentences(cur.get('summary') or cur_text):
        toks=set(_history_tokens(s))
        fresh=[t for t in toks-prev_tok if len(t)>=5 and t not in _V104_GENERIC_EVENT_WORDS]
        nums=_v109_numbers(s)-prev_nums
        score=len(fresh)*2+len(nums)*5
        if score>0:
            candidates.append((score,s))
    if candidates:
        best=max(candidates,key=lambda x:x[0])[1]
        best=_v66_formalize_sentence_endings(best).strip()
        if best and best[-1] not in '.!?': best+='.'
        return (
            'Önceki taramada yer almayan yeni sayısal/olgusal bilgi tespit edilmiştir: '+best
            if new_nums else
            'Önceki taramaya göre yeni ayrıntı eklenmiştir: '+best
        )

    if new_nums:
        return 'Önceki taramada bulunmayan yeni sayısal bilgiler açıklanmıştır: '+', '.join(sorted(new_nums)[:5])+'.'
    return 'Olayın içeriğinde önceki taramaya göre anlamlı yeni ayrıntılar tespit edilmiştir.'

def _compare_since_previous(df,current_scan_id=None):
    """V109 — yalnız gerçek değişiklikler ve doğrudan 'Ne Değişti?' açıklaması."""
    current=_v104_event_representatives(df)
    prev_id=_previous_scan_id(current_scan_id)
    previous=_load_scan_events(prev_id)
    if current is None or current.empty:
        return pd.DataFrame(),None,None
    if previous.empty:
        return pd.DataFrame(),prev_id,None

    prev_records=[p.to_dict() for _,p in previous.iterrows()]
    changes=[]
    for _,r in current.iterrows():
        c={
            'title':str(r.get('Başlık','') or ''),'source':str(r.get('Kaynak','') or ''),
            'url':str(r.get('URL','') or ''),'category':str(r.get('Kategori','') or ''),
            'summary':str(r.get('İçerik_Özeti','') or ''),'risk_score':int(r.get('Risk_Skoru',0) or 0),
            'risk_status':str(r.get('Risk_Durumu','') or ''),'verification':str(r.get('Doğrulama','') or ''),
            'source_count':int(r.get('Olay_Kaynak_Sayisi',1) or 1)
        }

        best=None; best_sim=0.0
        for pr in prev_records:
            sim=_v104_event_similarity(c['title'],c['summary'],pr.get('title',''),pr.get('summary',''))
            if c['url'] and c['url']==str(pr.get('url','') or ''):
                sim=max(sim,0.98)
            if sim>best_sim:
                best_sim=sim; best=pr

        if best is None or best_sim<0.50:
            kind='🆕 YENİ OLAY'; priority=100+c['risk_score']; prev_risk='—'
            diff=_v109_direct_difference({},c,kind)
        else:
            risk_up,verify_up,material,_,_=_v104_material_change(best,c)
            if risk_up:
                kind='⚠️ RİSK ARTTI'; priority=95+c['risk_score']
            elif verify_up:
                kind='✅ TEYİT GÜÇLENDİ'; priority=90+c['risk_score']
            elif material:
                kind='🔄 YENİ BİLGİ'; priority=80+c['risk_score']
            else:
                continue
            prev_risk=int(best.get('risk_score') or 0)
            diff=_v109_direct_difference(best,c,kind)

        changes.append({
            'Değişim':kind,'Ne Değişti?':diff,'Başlık':c['title'],'Kaynak':c['source'],
            'Kategori':c['category'],'Risk':c['risk_score'],'Önceki Risk':prev_risk,
            'Kaynak Sayısı':c['source_count'],'URL':c['url'],'_priority':priority
        })

    out=pd.DataFrame(changes)
    if not out.empty:
        out['_sig']=out.apply(
            lambda r:' '.join(sorted(_v104_event_tokens(r.get('Başlık',''),r.get('Ne Değişti?','')))),
            axis=1
        )
        out=out.sort_values(['_priority','Risk'],ascending=[False,False]).drop_duplicates('_sig',keep='first')
        out=out.drop(columns=['_priority','_sig'],errors='ignore')
    prev_time=str(previous.iloc[0].get('scanned_at','')) if not previous.empty else None
    return out,prev_id,prev_time

def _v109_chronology_events(df):
    if df is None or df.empty:
        return pd.DataFrame()
    x=df.copy()
    x['Tarih_dt']=pd.to_datetime(x.get('Tarih_dt'),utc=True,errors='coerce')
    rows=[]
    groups=x.groupby('Olay_ID',dropna=False) if 'Olay_ID' in x.columns else [(f'ROW-{i}',x.iloc[[i]]) for i in range(len(x))]
    for oid,g in groups:
        g=g.sort_values('Tarih_dt',ascending=False,na_position='last')
        recs=g.to_dict('records')
        rep=max(recs,key=_v107_source_quality) if recs else g.iloc[0].to_dict()
        latest=g.iloc[0]
        rep=dict(rep)
        rep['Tarih_dt']=latest.get('Tarih_dt')
        rep['Tarih']=latest.get('Tarih')
        domains={str(v) for v in g.get('Domain',pd.Series(dtype=str)).tolist() if str(v).strip()}
        sources=list(dict.fromkeys(str(v) for v in g.get('Kaynak',pd.Series(dtype=str)).tolist() if str(v).strip()))
        rep['Kaynak Sayısı']=max(len(domains),len(sources),int(rep.get('Olay_Kaynak_Sayisi',0) or 0))
        rep['Haber Sayısı']=len(g)
        rep['Kaynaklar']=' • '.join(sources[:8])
        rep['_Olay_ID']=str(oid)
        rows.append(rep)
    return pd.DataFrame(rows).sort_values('Tarih_dt',ascending=False,na_position='last').reset_index(drop=True)

def _v109_event_sources(df,event_id):
    if df is None or df.empty or not event_id or 'Olay_ID' not in df.columns:
        return pd.DataFrame()
    x=df[df['Olay_ID'].astype(str)==str(event_id)].copy()
    if x.empty: return x
    x['Tarih_dt']=pd.to_datetime(x.get('Tarih_dt'),utc=True,errors='coerce')
    return x.sort_values('Tarih_dt',ascending=False,na_position='last')

def _v109_official_source_type(r):
    text=norm(f"{r.get('Kaynak','')} {r.get('Domain','')} {r.get('Başlık','')} {r.get('URL','')}")
    if 'tuik' in text or 'tüik' in text or 'turkiye istatistik' in text: return 'TÜİK'
    if 'tubitak' in text or 'tübitak' in text: return 'TÜBİTAK'
    if 'kosgeb' in text: return 'KOSGEB'
    if 'turkpatent' in text or 'türkpatent' in text: return 'TÜRKPATENT'
    if re.search(r'\btse\b',text) or 'türk standartları' in text: return 'TSE'
    if 'ssb.gov' in text or 'savunma sanayii başkan' in text or 'savunma sanayii baskan' in text: return 'SSB'
    if 'sanayi.gov' in text or 'sanayi ve teknoloji bakan' in text: return 'Bakanlık'
    if 'resmigazete' in text or 'resmî gazete' in text or 'resmi gazete' in text: return 'Resmî Gazete'
    return 'Diğer Resmî'

def _official_radar_rows(df):
    if df is None or df.empty:
        return pd.DataFrame()
    x=df[df.apply(_is_official_radar_row,axis=1)].copy()
    if x.empty: return x
    x['Kurum Türü']=x.apply(_v109_official_source_type,axis=1)
    x=x.sort_values('Tarih_dt',ascending=False,na_position='last')
    return x.drop_duplicates(subset=['URL','Başlık'])

# ============================================================
# /V109
# ============================================================

# V109 — doğrudan fark, olay bazlı kronoloji ve kurum türü filtreli Resmî Kaynak Radarı.
# V108 kaynak zenginleştirme ve V106 kararlı çekirdek korunmuştur.

st.set_page_config(page_title='Sanayi & Teknoloji OSINT Radarı', page_icon='🛡️', layout='wide')

# ============================================================
# V55 — ŞİFRE KORUMASI
# V54 STABLE işlevlerine dokunmaz; yalnızca uygulama girişini korur.
# Streamlit Secrets:
# APP_PASSWORD = "guclu-sifreniz"
# ============================================================
def _v55_password_gate():
    try:
        expected = str(st.secrets["APP_PASSWORD"])
    except Exception:
        st.error(
            "🔐 Uygulama şifresi tanımlanmamış. "
            "Streamlit App Settings → Secrets bölümüne APP_PASSWORD ekleyin."
        )
        st.stop()

    if st.session_state.get("_v55_authenticated", False):
        return

    st.title("STB-Açık Kaynak Tarama Merkezi")
    st.caption("Devam etmek için uygulama şifresini girin.")

    with st.form("_v55_login_form", clear_on_submit=False):
        entered = st.text_input("Şifre", type="password")
        submitted = st.form_submit_button("Giriş Yap", use_container_width=True)

    if submitted:
        import hmac
        if hmac.compare_digest(str(entered), expected):
            st.session_state["_v55_authenticated"] = True
            st.rerun()
        else:
            st.error("Şifre hatalı.")

    st.stop()

_v55_password_gate()



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
INDUSTRIAL_LOCATION_TERMS = [
    'osb','organize sanayi','organize sanayi bölgesi','organize sanayi sitesinde',
    'fabrika','fabrikada','fabrikasında','tesis','tesiste','üretim tesisi','sanayi tesisi',
    'imalathane','atölye','depo','üretim alanı','sanayi sitesi'
]
CRITICAL_INCIDENT_TERMS = [
    'yangın','yangını','yangin','alev','alevler','yanıyor','yaniyor','yandı','yandi',
    'patlama','patladı','patladi','infilak','infilak etti','parlama',
    'fabrika yangını','tesis yangını','fabrika patlaması','tesis patlaması'
]

def is_osb_fire(title, snippet=''):
    """Geriye dönük uyumluluk: OSB + yangın bağlamı."""
    t=norm(f'{title} {snippet}')
    return (
        any(term in t for term in OSB_FIRE_LOCATION_TERMS)
        and any(term in t for term in ['yangın','yangını','yangin','alev','alevler','yanıyor','yaniyor','yandı','yandi'])
    )

def critical_industrial_incident(title, snippet=''):
    """
    Özel kırmızı alarm için:
    - OSB içi yangın
    - OSB içi patlama
    - OSB dışı fabrika/tesis yangını
    - OSB dışı fabrika/tesis patlaması
    """
    t=norm(f'{title} {snippet}')
    has_location=any(term in t for term in INDUSTRIAL_LOCATION_TERMS)
    has_incident=any(_v89_has_term(t,term) for term in CRITICAL_INCIDENT_TERMS)
    if not (has_location and has_incident):
        return None

    osb=any(term in t for term in OSB_FIRE_LOCATION_TERMS)
    fire=any(term in t for term in ['yangın','yangını','yangin','alev','alevler','yanıyor','yaniyor','yandı','yandi'])
    explosion=any(term in t for term in ['patlama','patladı','patladi','infilak','infilak etti','parlama'])

    if osb and explosion:
        return '💥 OSB PATLAMA'
    if osb and fire:
        return '🔥 OSB YANGINI'
    if explosion:
        return '💥 FABRİKA/TESİS PATLAMASI'
    if fire:
        return '🔥 FABRİKA/TESİS YANGINI'
    return '🚨 KRİTİK SANAYİ OLAYI'


NEGATION_OR_RESOLUTION_PHRASES = [
    'olmadı','olmadığı','bulunmadı','bulunmadığı','yaşanmadı','gerçekleşmedi',
    'etkilenmedi','etkilenmediği','risk bulunmuyor','risk yok','tehdit yok',
    'iptal edilmedi','kapanmadı','durmadı','sona erdi','kaldırıldı',
    'giderildi','çözüldü','önlendi','engellendi','bertaraf edildi'
]

POSITIVE_SIGNAL_TERMS = [
    'arttı','artış','yükseldi','yükseliş','rekor','büyüdü','büyüme',
    'yatırım','yatırım kararı','yeni yatırım','ihracat arttı','ihracat artışı',
    'kapasite arttı','kapasite artışı','üretim arttı','üretim artışı',
    'devreye alındı','faaliyete geçti','başarıyla','başarılı',
    'anlaşma imzalandı','sözleşme imzalandı','teslim edildi',
    'teşvik','destek','hibe','istihdam artışı','yeni istihdam'
]

SEVERE_NEGATIVE_TERMS = {
    'iflas','konkordato','üretim durdu','fabrika kapandı','toplu işten çıkarma',
    'siber saldırı','veri sızıntısı','fidye yazılımı','ambargo','yaptırım',
    'ihracat yasağı','lisans reddi','ruhsat iptali','sözleşme feshi',
    'ihale iptal edildi','patlama','yangın','can kaybı','ölüm',
    'kritik zafiyet','yolsuzluk','usulsüzlük','casusluk','tedarik krizi','çip krizi'
}

def _term_regex(term):
    # Alt-string kaynaklı "ceza/cezasız", "dava/davalar" vb. yanlış eşleşmeleri azalt.
    escaped=re.escape(term)
    if ' ' in term:
        return re.compile(escaped,re.I)
    return re.compile(r'(?<!\w)'+escaped+r'(?!\w)',re.I)

def _physical_incident_is_real(term, context):
    if term not in {'yangın','patlama','ölüm'}:
        return True
    incident_markers=[
        'çıktı','çıkan','meydana geldi','meydana gelen','patladı','infilak',
        'alev','yaralandı','yaralı','hasar','müdahale','söndürüldü',
        'kontrol altına','tahliye','hayatını kaybetti','öldü'
    ]
    return any(x in context for x in incident_markers)

def _active_adverse_terms(terms, text):
    t=norm(text)
    active=[]
    for term in terms:
        rx=_term_regex(term)
        matches=list(rx.finditer(t))
        if not matches:
            continue

        term_active=False
        for m in matches:
            lo=max(0,m.start()-90); hi=min(len(t),m.end()+90)
            ctx=t[lo:hi]

            # Fiziksel olay kelimesi yalnız kavramsal/önleyici bir kullanımdaysa alarm verme.
            if not _physical_incident_is_real(term,ctx):
                continue

            # "yaptırım kaldırıldı", "ihlal yaşanmadı", "üretim durmadı" gibi bağlamları bastır.
            # Gerçekleşmiş yangın/patlama/can kaybı ise "kontrol altına alındı" gibi sonraki olumlu
            # gelişmeler olayın negatif niteliğini ortadan kaldırmaz.
            if term not in {'yangın','patlama','can kaybı','ölüm'}:
                if any(p in ctx for p in NEGATION_OR_RESOLUTION_PHRASES):
                    continue

            term_active=True
            break

        if term_active:
            active.append(term)
    return active

def _sentence_chunks(text):
    txt=re.sub(r'\s+',' ',str(text or '')).strip()
    if not txt:
        return []
    return [x.strip() for x in re.split(r'(?<=[.!?;:])\s+',txt) if x.strip()]

def _negated_in_context(term, sentence):
    s=norm(sentence)
    # Olumsuzluk/çözülme ifadeleri, ilgili risk kelimesinin yakın çevresindeyse baskılanır.
    negators=[
        'değil','değildir','olmadı','olmadığı','bulunmadı','bulunmadığı',
        'yaşanmadı','gerçekleşmedi','etkilenmedi','etkilenmediği',
        'risk yok','tehdit yok','iptal edilmedi','kapanmadı','durmadı',
        'giderildi','çözüldü','önlendi','engellendi','kaldırıldı','sona erdi'
    ]
    return any(n in s for n in negators)

def _positive_strength(text):
    t=norm(text)
    positive_terms=[
        'arttı','artış','yükseldi','yükseliş','rekor','büyüdü','büyüme',
        'yeni yatırım','yatırım kararı','yatırım yaptı','yatırım yapacak',
        'ihracat arttı','ihracat artışı','kapasite artışı','kapasite arttı',
        'üretim artışı','üretim arttı','devreye alındı','faaliyete geçti',
        'başarıyla','başarılı','anlaşma imzalandı','sözleşme imzalandı',
        'teslim edildi','teşvik','destek','hibe','istihdam artışı','yeni istihdam',
        'pazar payı arttı','gelir arttı','kâr arttı','kar arttı'
    ]
    return sum(1 for x in positive_terms if x in t)

# V48 — Ekonomik/operasyonel haber dilinde sık görülen, önceki sözlükte kolay
# kaçabilen negatif sinyaller. Bunlar tam sayfa indirmeden RSS içerik/özetinde aranır.
V48_NEGATIVE_PHRASES = [
    'düştü','düşüş','azaldı','azalış','geriledi','gerileme','daraldı','daralma',
    'zarar açıkladı','zarar etti','net zarar','faaliyet zararı','kayıp yaşadı',
    'satışlar düştü','satışlar azaldı','satışlarda düşüş','satışlarda azalma',
    'üretim düştü','üretim azaldı','üretimde düşüş','üretimde azalma',
    'ihracat düştü','ihracat azaldı','ihracatta düşüş','ihracatta gerileme',
    'siparişler düştü','siparişler azaldı','siparişlerde düşüş',
    'kapasite düştü','kapasite azaldı','kapasite kullanım oranı düştü',
    'istihdam azaldı','istihdam düştü','istihdam kaybı',
    'işten çıkarma','işçi çıkarma','personel azaltma','toplu işten çıkarma',
    'maliyet arttı','maliyet artışı','maliyet baskısı','girdi maliyetleri arttı',
    'fiyat baskısı','finansman maliyeti','nakit sıkıntısı','likidite sıkıntısı',
    'talep düştü','talep azaldı','talep daralması','talepte daralma',
    'pazar payı düştü','pazar payı kaybı','rekabet gücü kaybı',
    'beklentinin altında','beklentilerin altında','hedefin altında','hedefin gerisinde',
    'kriz','aksama','kesinti','arıza','gecikme','ertelendi','iptal edildi',
    'faaliyet durdu','üretim durdu','üretime ara verdi','üretime ara verildi',
    'fabrika kapandı','tesis kapandı','kapanma kararı',
    'iflas','konkordato','haciz','borç krizi',
    'soruşturma','inceleme başlatıldı','ceza verildi','para cezası',
    'yasaklandı','yasak','geri çağırma','ürün geri çağırma',
    'siber saldırı','veri sızıntısı','veri ihlali','kritik açık','güvenlik açığı',
    'tedarik sorunu','tedarik krizi','tedarik zinciri aksaması',
    'kaza','yangın','patlama','yaralandı','can kaybı'
]


# V49 — Yapısal / eleştirel negatiflik katmanı
# Haber sayfasına gitmez; yalnızca eldeki Başlık + RSS içerik/özet üzerinde çalışır.
V49_STRUCTURAL_NEGATIVE = [
    'tehlikeli gidiş','tehlikeli seyir','olumsuz gidiş','olumsuz seyir',
    'kötü gidiş','kötüye gidiş','kötüleşiyor','kötüleşme',
    'alarm veriyor','alarm zilleri','kan kaybediyor','kan kaybı',
    'ivme kaybediyor','ivme kaybı','güç kaybediyor','güç kaybı',
    'rekabet gücü zayıflıyor','rekabet gücü geriliyor','rekabet gücü kaybı',
    'zayıf seyir','zayıflama','zayıflıyor','yavaşlıyor','yavaşlama',
    'sıkıntılı süreç','sıkıntılı dönem','kritik süreç','kritik eşik',
    'sorun büyüyor','sorunlar büyüyor','sorun devam ediyor','sorun sürüyor',
    'risk artıyor','riskler artıyor','baskı artıyor','baskı altında',
    'istenilen seviyede değil','istenen seviyede değil',
    'beklenen seviyede değil','yeterli değil','yetersiz kaldı','yetersiz kalıyor',
    'teşvik yetmiyor','teşvikler yetmiyor','destek yetmiyor','destekler yetmiyor',
    'sadece teşvik vermekle olmuyor','sadece destek vermekle olmuyor',
    'çözüm olmuyor','çözüm değil','sürdürülebilir değil',
    'endişe yaratıyor','endişe veriyor','kaygı yaratıyor','kaygı veriyor',
    'uyarı geldi','uyarı yaptı','uyardı','dikkat çekti',
    'olumsuz tablo','karamsar tablo','zorlu görünüm','zayıf görünüm',
    'darboğaz','çıkmaz','kırılganlık','kırılgan hale geldi'
]


# V50 — Geniş Negatif Bölümü
# Kullanıcı açısından "Negatif" yalnızca gerçekleşmiş kötü olay değildir.
# Eleştirel, uyarıcı, yetersizlik bildiren, politika/sektör performansını sorgulayan
# ve yapısal sorun işaret eden haberler de aynı Negatif bölümüne girer.
V50_CRITICAL_NEGATIVE = [
    'eleştirdi','eleştiri','eleştirildi','tepki gösterdi','tepki çekti',
    'itiraz etti','itiraz','uyarıda bulundu','uyarı yaptı','uyardı',
    'dikkat çekti','dikkat çekiyor','endişesini dile getirdi','kaygısını dile getirdi',
    'yeterli değil','yetersiz','yetersiz kaldı','yetersiz kalıyor',
    'eksik kaldı','eksiklik','yetmiyor','yetmedi','karşılamıyor',
    'çözüm değil','çözüm olmadı','çözüm olmuyor','sonuç vermiyor','sonuç vermedi',
    'etkili değil','etkisiz','başarısız','başarısızlık',
    'hedefin gerisinde','hedeflerin gerisinde','beklentinin altında','beklentilerin altında',
    'istenen seviyede değil','istenilen seviyede değil','arzu edilen seviyede değil',
    'sorunlu','sorunlar','sorun devam ediyor','sorun sürüyor','sorun büyüyor',
    'risk taşıyor','risk oluşturuyor','risk yaratıyor','tehdit oluşturuyor',
    'sürdürülebilir değil','kırılgan','kırılganlık','darboğaz',
    'rekabet sorunu','rekabet gücü kaybı','rekabet gücü zayıflıyor',
    'verimlilik sorunu','finansmana erişim sorunu','nitelikli iş gücü sorunu',
    'maliyet baskısı','finansman baskısı','kur baskısı',
    'sanayici zorlanıyor','sektör zorlanıyor','firmalar zorlanıyor',
    'üretici zorlanıyor','ihracatçı zorlanıyor',
    'teşvikler yetersiz','destekler yetersiz','teşvik yetmiyor','destek yetmiyor',
    'politika yetersiz','politikalar yetersiz','düzenleme yetersiz',
    'önlem yetersiz','tedbir yetersiz','önlemler yetersiz','tedbirler yetersiz'
]

def _v50_critical_negative_signals(text):
    t = norm(text)
    found = set()
    for phrase in V50_CRITICAL_NEGATIVE:
        if phrase in t:
            # Açık biçimde reddedilen eleştirileri yanlış negatif yapma.
            idx = t.find(phrase)
            ctx = t[max(0, idx-65):idx+len(phrase)+65] if idx >= 0 else t
            if any(x in ctx for x in [
                'eleştiri yok','sorun yok','risk yok','yetersiz değil',
                'başarısız değil','kırılgan değil','zorlanmıyor'
            ]):
                continue
            found.add(phrase)
    return found

V49_PERSISTENCE_PATTERNS = [
    r'\b\d+\s*(?:çeyrektir|çeyrek boyunca|aydır|ay boyunca|yıldır|yıl boyunca|haftadır)\b',
    r'\buzun süredir\b',
    r'\bsüregelen\b',
    r'\bdevam eden\b',
    r'\bsürmekte olan\b',
    r'\bkronik\b'
]

def _v49_structural_negative_signals(text):
    t = norm(text)
    found = set()

    for phrase in V49_STRUCTURAL_NEGATIVE:
        if phrase in t:
            found.add(phrase)

    persistent = any(re.search(pat, t, re.I) for pat in V49_PERSISTENCE_PATTERNS)

    # Süre ifadesi tek başına negatif değildir. Ancak yapısal negatif bir ifade
    # veya başka bir negatif sinyal ile birlikteyse ağırlık kazanır.
    return found, persistent

V48_STRONG_NEGATIVE = [
    'üretim durdu','faaliyet durdu','fabrika kapandı','tesis kapandı',
    'toplu işten çıkarma','iflas','konkordato','siber saldırı','veri sızıntısı',
    'veri ihlali','yangın','patlama','can kaybı','hayatını kaybetti',
    'ihracat yasağı','yaptırım','ambargo'
]

V48_DIRECTION_PATTERNS = [
    r'(?:yüzde|%)\s*\d+(?:[.,]\d+)?\s*(?:düştü|azaldı|geriledi|daraldı)',
    r'\d+(?:[.,]\d+)?\s*(?:%|yüzde)\s*(?:düştü|azaldı|geriledi|daraldı)',
    r'(?:üretim|ihracat|satış|sipariş|istihdam|kapasite|talep|gelir|kâr|kar)\w*\s+.{0,55}\b(?:düştü|azaldı|geriledi|daraldı)\b',
    r'\b(?:geçen yıla|önceki yıla|geçen aya|önceki aya)\s+göre.{0,70}\b(?:düştü|azaldı|geriledi|daraldı)\b'
]


def _v89_has_term(text, term):
    """
    Negatif terimleri alt-dize ile değil kelime/ifade sınırıyla arar.
    Böylece 'kaza' -> 'kazandı/kazanç/kazanım' eşleşmesi oluşmaz.
    """
    t=norm(text)
    term=norm(term)
    if not term:
        return False
    # _term_regex mevcut negatif analiz motorunun güvenli eşleştiricisidir.
    try:
        return bool(_term_regex(term).search(t))
    except Exception:
        # Fallback: tek kelimede Unicode kelime sınırı, çok kelimede sınırlandırılmış ifade.
        return bool(re.search(r'(?<!\w)'+re.escape(term)+r'(?!\w)',t,re.I))

def _v48_extra_negative_signals(text):
    t=norm(text)
    found=set()
    for phrase in V48_NEGATIVE_PHRASES:
        if _v89_has_term(t,phrase):
            # "düşmedi / azalmadı / gerilemedi" gibi açık olumsuzlamaları alma.
            m=_term_regex(phrase).search(t)
            pos=m.start() if m else -1
            ctx=t[max(0,pos-60):pos+len(phrase)+60] if pos>=0 else t
            if any(x in ctx for x in [
                'düşmedi','azalmadı','gerilemedi','daralmadı','iptal edilmedi',
                'aksama olmadı','kesinti olmadı','etkilenmedi','risk yok'
            ]):
                continue
            found.add(phrase)

    directional=False
    for pat in V48_DIRECTION_PATTERNS:
        if re.search(pat,t,re.I):
            directional=True
            found.add('sayısal/yönsel düşüş')
            break
    return found,directional

def _negative_sentence_analysis(title, snippet):
    """
    V48 hızlı hassas analiz:
    - Başlık + RSS içerik/özet birlikte
    - mevcut negatif/risk sözlükleri
    - geniş ekonomik/operasyonel sözlük
    - sayısal/yönsel düşüş tespiti
    - olumsuzlama kontrolü
    """
    title_n=norm(title)
    full=f"{title}. {snippet}"
    sentences=_sentence_chunks(full)

    active_neg=set()
    active_risk=set()
    title_neg=set()
    title_risk=set()
    strong_event=False

    physical_terms={'yangın','patlama','can kaybı','ölüm'}
    physical_markers=[
        'çıktı','çıkan','meydana geldi','meydana gelen','patladı','infilak',
        'alev','yaralandı','yaralı','hasar','müdahale','söndürüldü',
        'kontrol altına','tahliye','hayatını kaybetti','öldü'
    ]

    for s in sentences:
        sn=norm(s)
        for term in NEGATIVE_TERMS:
            if not _term_regex(term).search(sn):
                continue
            if term in physical_terms:
                if not any(m in sn for m in physical_markers):
                    continue
            elif _negated_in_context(term,sn):
                continue
            active_neg.add(term)
            if _term_regex(term).search(title_n):
                title_neg.add(term)

        for term in HIGH_RISK_TERMS:
            if not _term_regex(term).search(sn):
                continue
            if term in physical_terms:
                if not any(m in sn for m in physical_markers):
                    continue
            elif _negated_in_context(term,sn):
                continue
            active_risk.add(term)
            if _term_regex(term).search(title_n):
                title_risk.add(term)

    extra,directional=_v48_extra_negative_signals(full)
    active_neg.update(extra)

    for phrase in extra:
        if phrase!='sayısal/yönsel düşüş' and phrase in title_n:
            title_neg.add(phrase)

    # V49: klasik düşüş/zarar kelimesi bulunmasa bile eleştirel ve yapısal
    # kötüleşme dili ayrıca yakalanır.
    structural,persistent=_v49_structural_negative_signals(full)
    active_neg.update(structural)
    for phrase in structural:
        if phrase in title_n:
            title_neg.add(phrase)

    # V50: eleştirel/uyarıcı/yetersizlik bildiren içerikler de doğrudan
    # mevcut Negatif havuzuna eklenir. Ayrı kategori oluşturulmaz.
    critical_negative=_v50_critical_negative_signals(full)
    active_neg.update(critical_negative)
    for phrase in critical_negative:
        if phrase in title_n:
            title_neg.add(phrase)

    if any(x in norm(full) for x in V48_STRONG_NEGATIVE):
        strong_event=True

    return active_neg,active_risk,title_neg,title_risk,strong_event,directional,structural,persistent,critical_negative

def classify(title,snippet,source_domain=''):
    full=f'{title} {snippet}'
    t=norm(full)

    neg_set,risk_set,title_neg,title_risk,strong_event,directional,structural,persistent,critical_negative=_negative_sentence_analysis(title,snippet)
    neg=sorted(neg_set)
    risk=sorted(risk_set)

    cat='Genel Sanayi / Teknoloji'
    for c,ks in CATEGORIES.items():
        if any(k in t for k in ks):
            cat=c
            break

    score=5
    reasons=[]

    if neg:
        score += min(30,6*len(neg))
        score += min(14,5*len(title_neg))
        reasons.append(f'{len(neg)} doğrulanmış negatif sinyal')

    if directional:
        score += 8
        reasons.append('ölçülebilir düşüş/gerileme')

    if structural:
        score += min(16, 7 + 3*len(structural))
        reasons.append('yapısal/eleştirel olumsuzluk')

    if critical_negative:
        # Eleştirel yaklaşım doğrudan Negatif bölümüne girecek kadar ağırlık alır,
        # fakat tek başına Yüksek Risk sayılmaz.
        score += min(15, 8 + 2*len(critical_negative))
        reasons.append('eleştirel/uyarıcı yaklaşım')

    if persistent and (structural or critical_negative or neg_set):
        score += 8
        reasons.append('olumsuzluğun sürekliliği')

    if risk:
        score += min(32,9*len(risk))
        score += min(14,5*len(title_risk))
        reasons.append(f'{len(risk)} yüksek risk sinyali')

    if strong_event:
        score += 14
        reasons.append('doğrudan ağır olumsuz olay')

    # Gerçek negatiflik varsa sektörel etki skoru eklenir.
    if neg or risk:
        if any(x in t for x in ['üretim','fabrika','tesis','istihdam','kapasite','ihracat','tedarik','satış','sipariş']):
            score += 6
            reasons.append('üretim/ekonomi etkisi')
        if any(x in t for x in ['savunma','kritik altyapı','enerji','siber','yarı iletken','çip']):
            score += 7
            reasons.append('stratejik/kritik sektör etkisi')

    positive_count=_positive_strength(full)
    severe_active=strong_event or any(x in norm(full) for x in V48_STRONG_NEGATIVE)

    # V48 farkı: olumlu sinyal gerçek negatifliği SİLMEZ.
    # Yalnızca ağır risk yoksa skoru sınırlı ölçüde dengeler.
    if positive_count and neg and not severe_active:
        score=max(0,score-min(8,2*positive_count))
        reasons.append('karma/olumlu unsurlar mevcut')

    score=max(0,min(100,score))

    # En kritik değişiklik: gerçek ve bağlamsal negatif sinyal bulunduysa,
    # yüksek risk olmasa dahi haber Negatif olabilir.
    sentiment='Negatif' if neg else 'Nötr'

    if severe_active and (risk or score>=55):
        status='Yüksek Risk'
    elif risk and score>=68:
        status='Yüksek Risk'
    elif (structural or critical_negative) and neg:
        status='Negatif'
    elif neg and score>=18:
        status='Negatif'
    else:
        status='Normal'

    if status=='Normal' and not neg:
        reasons=['olumsuz risk sinyali tespit edilmedi']

    # V87 — çok sınırlı yanlış-negatif koruması.
    # Açık başarı/madalya/ödül ve normal test ilerlemesi haberleri,
    # başlıkta gerçek bir olumsuzluk yoksa negatif değildir.
    _hn=norm(title)
    _positive_head=bool(re.search(
        r'(madalya\s+kazan|ödül\s+kazan|şampiyon|rekor\s+kır|başarıyla|'
        r'başarı\s+elde|testleri?\s+devam\s+ediyor|test\s+süreci\s+devam)',
        _hn,re.I
    ))
    _bad_head=bool(re.search(
        r'(başarısız|kaza|yangın|patlama|ölüm|yaralan|iptal|gecik|arıza|'
        r'iflas|saldırı|eleştir|yetersiz|kriz|sorun|tehlike|zarar|kayıp|'
        r'geriledi|azaldı|düştü|ceza|yaptırım)',
        _hn,re.I
    ))
    if _positive_head and not _bad_head:
        sentiment='Nötr'
        status='Normal'
        score=min(score,12)
        neg=[]
        risk=[]
        reasons=['açık başarı veya normal test/program ilerlemesi; negatif değildir']

    return sentiment,score,status,neg,risk,cat,reasons

def _v89_negative_selfcheck():
    """Basit regresyon kontrolleri; panelde gösterilmez."""
    cases=[
        ('Türk öğrenciler uluslararası yarışmada 15 madalya kazandı','Nötr'),
        ('Şirket yılın ilk yarısında güçlü kazanç açıkladı','Nötr'),
        ('Yeni teknoloji kazanımı ihracat kapasitesini artırdı','Nötr'),
    ]
    for h,expected in cases:
        try:
            sent,_,_,_,_,_,_=classify(h,'')
            if sent!=expected:
                return False
        except Exception:
            return False
    return True


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


# -----------------------------
# V41 — RESMÎ KAYNAK / İSTATİSTİK RADARI
# -----------------------------
OFFICIAL_RADAR_DOMAINS = [
    'sanayi.gov.tr','tubitak.gov.tr','kosgeb.gov.tr','turkpatent.gov.tr','tse.org.tr',
    'ssb.gov.tr','tuik.gov.tr','tcmb.gov.tr','ticaret.gov.tr','epdk.gov.tr','teias.gov.tr',
    'tua.gov.tr'
]

PRIMARY_STATS_DOMAINS = [
    'tuik.gov.tr','tcmb.gov.tr','ticaret.gov.tr','sanayi.gov.tr','ssb.gov.tr',
    'epdk.gov.tr','teias.gov.tr','tim.org.tr','osd.org.tr','odmd.org.tr'
]

STATISTIC_TERMS = [
    'sanayi üretim','sanayi üretimi','üretim endeksi','imalat sanayi',
    'kapasite kullanım','kapasite kullanım oranı','kko',
    'ihracat','dış ticaret','dış ticaret istatistik',
    'otomotiv üretim','otomotiv ihracat','araç üretim',
    'savunma ihracat','savunma ve havacılık ihracat',
    'elektrik üretim','enerji üretim','kurulu güç','tüketim',
    'yatırım teşvik','teşvik belgesi','sabit yatırım',
    'ar-ge','arge','araştırma geliştirme','yenilik','patent başvuru',
    'teknoloji istatistik','bilişim','girişim','yüksek teknoloji'
]

def build_official_radar_queries(when):
    """Genel medya taramasından ayrı, birincil/resmî kaynak sorguları."""
    gov_sites='('+' OR '.join('site:'+d for d in OFFICIAL_RADAR_DOMAINS)+')'
    return [
        f'(sanayi OR teknoloji OR üretim OR yatırım OR ihracat OR savunma OR Ar-Ge OR patent) {gov_sites} when:{when}',
        f'("basın açıklaması" OR duyuru OR açıklandı OR yayımlandı OR rapor OR veri OR istatistik) {gov_sites} when:{when}'
    ]

def build_statistics_queries(when):
    """Günlük sayısal veri yayımlarını yakalamaya dönük dar ve hızlı ek sorgular."""
    sites='('+' OR '.join('site:'+d for d in PRIMARY_STATS_DOMAINS)+')'
    return [
        f'("sanayi üretimi" OR "kapasite kullanım" OR ihracat OR "dış ticaret" OR "otomotiv üretimi") {sites} when:{when}',
        f'("savunma ihracatı" OR "enerji üretimi" OR "kurulu güç" OR "yatırım teşvik" OR "Ar-Ge") {sites} when:{when}'
    ]

def _is_official_radar_row(r):
    d=domain(r.get('Domain','') or r.get('URL',''))
    srcn=norm(r.get('Kaynak',''))
    if d in OFFICIAL_RADAR_DOMAINS or d in PRIMARY_STATS_DOMAINS:
        return True
    names=['sanayi ve teknoloji bakanlığı','tübitak','tubitak','kosgeb','türkpatent','turkpatent',
           'tse','savunma sanayii başkanlığı','ssb','tüik','tuik','tcmb','ticaret bakanlığı',
           'epdk','teiaş','teias','türkiye uzay ajansı']
    return any(x in srcn for x in names)


# -----------------------------
# V52 — GÜNÜN EN DEĞERLİ 10 GELİŞMESİ
# -----------------------------
V52_STRATEGIC_TERMS=[
    'savunma','savunma sanayii','tusaş','aselsan','roketsan','havelsan','baykar',
    'kaan','kızılelma','füze','hava savunma','siber','kritik altyapı',
    'yapay zeka','yarı iletken','çip','nükleer','enerji','otomotiv',
    'yatırım','fabrika','üretim','ihracat','arge','ar-ge','teknoloji yatırımı',
    'kritik mineral','nadir toprak','tedarik zinciri'
]

def _v52_event_value_table(df,n=10):
    """
    Olay bazlı 0-100 Değer Skoru.
    Gerçek okunma/tıklanma verisi mevcut akışta bulunmadığından uydurulmaz.
    Bunun yerine erişilebilen güçlü vekiller kullanılır:
    önem/risk, kaynak yayılımı, resmî teyit, güncellik, stratejik önem,
    negatif/eleştirel etki ve aynı olayın haber yoğunluğu.
    """
    cols=['Sıra','Değer_Skoru','Tarih','Gelişme','Neden_Değerli',
          'Kaynak_Sayısı','Haber_Sayısı','Resmî_Teyit','Risk','URL']
    if df is None or df.empty:
        return pd.DataFrame(columns=cols)

    now=pd.Timestamp.now(tz='UTC')
    items=[]

    for oid,g in df.groupby('Olay_ID',dropna=False):
        g=g.sort_values('Tarih_dt',ascending=False).copy()
        rep=g.iloc[0]
        maxrisk=int(pd.to_numeric(g.get('Risk_Skoru',0),errors='coerce').fillna(0).max())
        domains={domain(x) for x in g.get('Domain',pd.Series(dtype=str)).astype(str) if x}
        source_count=max(1,len(domains))
        news_count=len(g)
        official=any(_is_official_radar_row(r) for _,r in g.iterrows())

        latest=pd.to_datetime(g['Tarih_dt'],utc=True,errors='coerce').max()
        age_h=max(0.0,(now-latest).total_seconds()/3600) if pd.notna(latest) else 24.0
        recency=max(0.0,1.0-min(age_h,24.0)/24.0)

        text=norm(' '.join(
            (g['Başlık'].fillna('').astype(str)+' '+g['İçerik_Özeti'].fillna('').astype(str)).head(6).tolist()
        ))
        strategic_hits=sum(1 for x in V52_STRATEGIC_TERMS if x in text)
        strategic=min(1.0,strategic_hits/3.0)

        negative=bool(
            (g.get('Duygu',pd.Series(index=g.index,dtype=str))=='Negatif').any()
            or (g.get('Risk_Durumu',pd.Series(index=g.index,dtype=str))=='Yüksek Risk').any()
        )

        # 0-100: kullanıcının istediği kıstaslara göre dengeli ağırlık.
        risk_part=min(25.0,maxrisk*0.25)
        spread_part=min(20.0,5.0*source_count + max(0,news_count-source_count)*1.5)
        official_part=15.0 if official else 0.0
        recency_part=10.0*recency
        strategic_part=15.0*strategic
        impact_part=10.0 if negative else (5.0 if maxrisk>=35 else 0.0)

        # Gerçek click/read metriği yoksa "çok sayıda bağımsız kaynakta yankı"
        # popülerlik vekili olarak en fazla 5 puan taşır.
        popularity_proxy=min(5.0,max(0,source_count-1)*1.5 + max(0,news_count-2)*0.5)

        score=int(round(min(100,risk_part+spread_part+official_part+recency_part+
                            strategic_part+impact_part+popularity_proxy)))

        why=[]
        if source_count>=4: why.append(f'{source_count} farklı kaynakta geniş yankı')
        elif source_count>=2: why.append(f'{source_count} farklı kaynakta yer aldı')
        if official: why.append('resmî/birincil kaynak teyidi')
        if maxrisk>=70: why.append('yüksek risk/önem')
        elif maxrisk>=35: why.append('dikkat gerektiren etki')
        if strategic>=0.67: why.append('stratejik sanayi-teknoloji konusu')
        elif strategic>0: why.append('sanayi-teknoloji açısından ilgili')
        if negative: why.append('negatif/eleştirel etki')
        if recency>=0.75: why.append('çok güncel')
        if not why: why.append('güncel olay yoğunluğu')

        items.append({
            'Değer_Skoru':score,
            'Tarih':rep.get('Tarih',''),
            'Gelişme':rep.get('Başlık',''),
            'Neden_Değerli':' • '.join(why[:5]),
            'Kaynak_Sayısı':source_count,
            'Haber_Sayısı':news_count,
            'Resmî_Teyit':'Evet' if official else 'Hayır',
            'Risk':maxrisk,
            'URL':rep.get('URL','')
        })

    out=pd.DataFrame(items)
    if out.empty: return pd.DataFrame(columns=cols)
    out=out.sort_values(['Değer_Skoru','Kaynak_Sayısı','Haber_Sayısı','Tarih'],
                        ascending=[False,False,False,False]).head(n).reset_index(drop=True)
    out.insert(0,'Sıra',range(1,len(out)+1))
    return out[cols]


def _v53_find_event_row(df, value_row):
    """Top-10 satırını ana dataframe'deki temsilci haberle eşleştirir."""
    if df is None or df.empty:
        return None
    url=str(value_row.get('URL','') or '')
    title=norm(value_row.get('Gelişme',''))
    if url and 'URL' in df.columns:
        m=df[df['URL'].astype(str)==url]
        if not m.empty:
            return m.iloc[0]
    if title:
        m=df[df['Başlık'].astype(str).map(norm)==title]
        if not m.empty:
            return m.iloc[0]
    return None

def _v54_content_sentences(text,title=''):
    """Tam haber metninden menü/tekrar/gürültüyü azaltarak bilgi taşıyan cümleleri seçer."""
    clean=_clean_note_text(text)
    if not clean:
        return []
    title_n=norm(title)
    raw=_sentence_chunks(clean)
    out=[]; seen=set()
    noise=[
        'çerez','cookie','reklam','abonelik','bildirimleri aç','tüm hakları saklıdır',
        'gizlilik politikası','kullanım koşulları','facebook','instagram','twitter',
        'whatsapp','telegram','son dakika haberleri için'
    ]
    for s in raw:
        s=_clean_note_text(s)
        sn=norm(s)
        if len(s)<35 or len(s)>650: continue
        if any(x in sn for x in noise): continue
        if title_n and sn==title_n: continue
        key=re.sub(r'\W+','',sn)[:260]
        if not key or key in seen: continue
        seen.add(key)
        out.append(s)
    return out

def _v54_article_summary(detail,fallback_row,max_sentences=4):
    """
    Haberin içeriğini 2-4 bilgi yoğun cümlede özetler.
    Değer skoru, kaynak sayısı, resmî teyit gibi sıralama metadatasını özete katmaz.
    """
    title=_clean_note_text((detail or {}).get('title','') or fallback_row.get('Başlık',''))
    text=_clean_note_text((detail or {}).get('text','') or fallback_row.get('İçerik_Özeti',''))
    sents=_v54_content_sentences(text,title)

    if not sents:
        fallback=_clean_note_text(fallback_row.get('İçerik_Özeti',''))
        return fallback[:900].strip() if fallback else title

    # İlk anlamlı cümle bağlamı korur. Sonraki cümleler bilgi yoğunluğuna göre seçilir.
    selected=[sents[0]]
    candidates=[]
    for idx,s in enumerate(sents[1:],1):
        sn=norm(s)
        score=0
        if re.search(r'\b\d+(?:[.,]\d+)?\b',s): score+=4
        if any(x in sn for x in [
            'açıkladı','duyurdu','belirtti','bildirdi','ifade etti','kaydetti',
            'üretim','ihracat','ithalat','yatırım','istihdam','kapasite','satış',
            'sözleşme','anlaşma','teslim','tedarik','teşvik','destek','proje',
            'yangın','patlama','hasar','yaralı','kayıp','siber','veri',
            'arttı','azaldı','düştü','yükseldi','geriledi','başladı','tamamlandı'
        ]): score+=3
        if any(x in sn for x in [
            'bakanlık','tüik','ssb','şirket','firma','kurum','başkanlığı',
            'genel müdür','bakan','başkanı'
        ]): score+=1
        # Çok erken cümlelere hafif öncelik.
        score+=max(0,3-min(idx,3))
        candidates.append((score,idx,s))

    for _,_,s in sorted(candidates,key=lambda x:(-x[0],x[1])):
        if s not in selected:
            selected.append(s)
        if len(selected)>=max_sentences:
            break

    # Haber akışını bozmayacak şekilde özgün sıraya döndür.
    order={s:i for i,s in enumerate(sents)}
    selected=sorted(selected,key=lambda s:order.get(s,999))
    result=_join_sentences_naturally(selected)

    # Tek olayın 45 satırlık toplam özeti şişirmesini önle.
    return result[:1500].strip()

def _v54_deep_top10_summary(df,value10,max_lines=45):
    """
    Yalnızca Top-10 olayın temsilci haber sayfalarını butona basılınca zenginleştirir.
    Normal tarama hızını etkilemez. Her olay içerik odaklı özetlenir.
    """
    if value10 is None or value10.empty:
        return "Bugünün en değerli gelişmeleri arasında özet oluşturulabilecek içerik bulunamadı."

    lines=[
        f"Sanayi ve teknoloji gündeminde günün en değerli {len(value10)} gelişmesine ilişkin durum özeti aşağıda sunulmuştur.",
        ""
    ]

    for _,v in value10.head(10).iterrows():
        if len(lines)>=max_lines-3:
            break
        rank=int(v.get('Sıra',0) or 0)
        row=_v53_find_event_row(df,v)
        title=_clean_note_text(v.get('Gelişme',''))

        if row is None:
            detail_text=title
        else:
            try:
                # Ağ/tam metin işlemi SADECE özet butonuna basıldığında bu 10 haber için çalışır.
                detail=article_detail(row.to_dict() if hasattr(row,'to_dict') else row)
            except Exception:
                detail=None
            detail_text=_v54_article_summary(detail,row,4)

        lines.append(f"{rank}. {title}")
        if detail_text:
            lines.append(detail_text)
        lines.append("")

    lines.append(
        "Söz konusu gelişmelerin yeni açıklamalar ve ilave açık kaynak verileri doğrultusunda takip edilmesi önem taşımaktadır."
    )
    return '\n'.join(lines[:max_lines])

def make_v54_top10_summary_docx(df,value10,text=None):
    text=text or _v54_deep_top10_summary(df,value10,45)
    doc=Document()
    sec=doc.sections[0]
    sec.top_margin=Cm(2); sec.bottom_margin=Cm(2)
    sec.left_margin=Cm(2.5); sec.right_margin=Cm(2.5)
    doc.styles['Normal'].font.name='Times New Roman'
    doc.styles['Normal'].font.size=Pt(11)

    p=doc.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    rr=p.add_run('BUGÜNÜN SANAYİ VE TEKNOLOJİ DURUM ÖZETİ')
    rr.bold=True; rr.font.size=Pt(14)

    p=doc.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(datetime.now().astimezone().strftime('%d.%m.%Y %H:%M'))

    for line in text.splitlines():
        if not line.strip(): continue
        bp=doc.add_paragraph()
        bp.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        bp.paragraph_format.space_after=Pt(5)
        r=bp.add_run(line)
        if re.match(r'^\d+\.\s',line):
            r.bold=True

    bio=BytesIO()
    doc.save(bio); bio.seek(0)
    return bio.getvalue()



# -----------------------------
# V58 — ANALİTİK TAKİP ÜÇLÜSÜ
# 1) Olay Yaşam Döngüsü
# 2) Takip Edilecek Açık Hususlar
# 3) Teyit / Çelişki Matrisi
# Ek web isteği yapmaz; mevcut tarama sonuçlarını kullanır.
# -----------------------------

V58_RESOLUTION_TERMS=[
    'kontrol altına alındı','söndürüldü','sona erdi','tamamlandı','çözüldü',
    'giderildi','yeniden başladı','üretim yeniden başladı','faaliyet yeniden başladı',
    'normalleşti','normal seyrine döndü','tahliye sona erdi','arıza giderildi',
    'erişim sağlandı','sistem yeniden devreye alındı'
]

V58_ESCALATION_TERMS=[
    'arttı','büyüdü','genişledi','yayılıyor','devam ediyor','sürüyor',
    'üretim durdu','faaliyet durdu','tahliye','ikinci patlama','yeni patlama',
    'can kaybı','yaralı sayısı','hasar arttı','soruşturma başlatıldı',
    'acil durum','kriz','kesinti sürüyor'
]

def _v58_event_groups(df):
    if df is None or df.empty or 'Olay_ID' not in df.columns:
        return {}
    groups={}
    for oid,g in df.groupby('Olay_ID',dropna=False):
        groups[str(oid)]=g.sort_values('Tarih_dt',ascending=True).copy()
    return groups

def _v58_event_stage(g):
    """Olayın mevcut taramadaki izlerine göre yaşam döngüsü aşaması."""
    if g is None or g.empty:
        return 'İlk Sinyal'

    text=norm(' '.join(
        (g['Başlık'].fillna('').astype(str)+' '+g['İçerik_Özeti'].fillna('').astype(str)).tolist()
    ))
    source_count=max(1,g['Domain'].astype(str).replace('',pd.NA).dropna().nunique()) if 'Domain' in g.columns else 1
    news_count=len(g)
    official=any(_is_official_radar_row(r) for _,r in g.iterrows())

    if any(x in text for x in V58_RESOLUTION_TERMS):
        return '✅ Sonuçlandı'
    if official:
        return '🟢 Teyit Edildi'
    if source_count>=2 or news_count>=3 or any(x in text for x in V58_ESCALATION_TERMS):
        return '🟠 Gelişiyor'
    return '🔵 İlk Sinyal'

def _v58_stage_reason(g,stage):
    source_count=max(1,g['Domain'].astype(str).replace('',pd.NA).dropna().nunique()) if 'Domain' in g.columns else 1
    news_count=len(g)
    official=any(_is_official_radar_row(r) for _,r in g.iterrows())
    text=norm(' '.join(
        (g['Başlık'].fillna('').astype(str)+' '+g['İçerik_Özeti'].fillna('').astype(str)).tolist()
    ))
    reasons=[]
    if official: reasons.append('resmî/birincil açıklama mevcut')
    if source_count>=2: reasons.append(f'{source_count} farklı kaynak')
    if news_count>=3: reasons.append(f'{news_count} haber kaydı')
    if any(x in text for x in V58_RESOLUTION_TERMS): reasons.append('sonuç/normalleşme ifadesi')
    elif any(x in text for x in V58_ESCALATION_TERMS): reasons.append('devam/etki artışı sinyali')
    if not reasons: reasons.append('tek/erken kaynak sinyali')
    return ' • '.join(reasons)

def _v58_event_lifecycle_table(df,limit=25):
    cols=['Tarih','Aşama','Başlık','Kategori','Kaynak_Sayısı','Haber_Sayısı',
          'Doğrulama','Risk_Skoru','Aşama_Gerekçesi','URL']
    groups=_v58_event_groups(df)
    rows=[]
    for oid,g in groups.items():
        latest=g.sort_values('Tarih_dt',ascending=False).iloc[0]
        stage=_v58_event_stage(g)
        source_count=max(1,g['Domain'].astype(str).replace('',pd.NA).dropna().nunique()) if 'Domain' in g.columns else 1
        rows.append({
            'Tarih':latest.get('Tarih',''),
            'Aşama':stage,
            'Başlık':latest.get('Başlık',''),
            'Kategori':latest.get('Kategori',''),
            'Kaynak_Sayısı':source_count,
            'Haber_Sayısı':len(g),
            'Doğrulama':latest.get('Doğrulama',''),
            'Risk_Skoru':int(pd.to_numeric(g['Risk_Skoru'],errors='coerce').fillna(0).max()) if 'Risk_Skoru' in g.columns else 0,
            'Aşama_Gerekçesi':_v58_stage_reason(g,stage),
            'URL':latest.get('URL',''),
            '_stage_rank':{'🟠 Gelişiyor':4,'🟢 Teyit Edildi':3,'🔵 İlk Sinyal':2,'✅ Sonuçlandı':1}.get(stage,0),
            '_dt':pd.to_datetime(latest.get('Tarih_dt'),utc=True,errors='coerce')
        })
    if not rows:
        return pd.DataFrame(columns=cols)
    out=pd.DataFrame(rows).sort_values(['_stage_rank','Risk_Skoru','_dt'],ascending=[False,False,False])
    return out.head(limit).drop(columns=['_stage_rank','_dt'],errors='ignore')

def _v58_open_questions_for_group(g):
    """
    'Bilinmiyor' iddiası üretmez; mevcut içerikte ayrıca teyit/izleme gerektiren
    alanları analist kontrol listesi olarak önerir.
    """
    text=norm(' '.join(
        (g['Başlık'].fillna('').astype(str)+' '+g['İçerik_Özeti'].fillna('').astype(str)).tolist()
    ))
    qs=[]

    if any(_v89_has_term(text,x) for x in ['yangın','patlama','infilak','kaza']):
        qs += [
            'Olayın kesin nedeni ve teknik inceleme sonucu',
            'Can kaybı/yaralı ve maddi hasarın resmî bilançosu',
            'Üretim/faaliyet sürekliliğine etkisi ve normale dönüş takvimi'
        ]
    if any(x in text for x in ['siber','veri sızıntısı','veri ihlali','fidye','güvenlik açığı']):
        qs += [
            'Etkilenen sistem/veri kapsamının kesinleştirilmesi',
            'İhlalin kaynağı ve alınan düzeltici tedbirler',
            'Operasyonel hizmetlere etkisinin sürüp sürmediği'
        ]
    if any(x in text for x in ['yatırım','fabrika kurulacak','tesis kurulacak','teşvik']):
        qs += [
            'Yatırım tutarı, kapasitesi ve finansman yapısının teyidi',
            'Yatırım/üretime geçiş takvimi',
            'İstihdam ve yerli tedarik etkisinin netleşmesi'
        ]
    if any(x in text for x in ['ihracat','sözleşme','anlaşma','sipariş','teslimat','savunma']):
        qs += [
            'Sözleşme/anlaşmanın kapsamı ve parasal büyüklüğü',
            'Teslimat/uygulama takvimi',
            'Karşı taraf veya resmî makam teyidi'
        ]
    if any(x in text for x in ['üretim düştü','daralma','geriledi','azaldı','maliyet baskısı','rekabet gücü']):
        qs += [
            'Olumsuz eğilimin geçici mi yapısal mı olduğunun izlenmesi',
            'Bir sonraki resmî veri setinde eğilimin devam edip etmediği',
            'Sektör/şirket bazında üretim, ihracat ve istihdam etkisi'
        ]

    official=any(_is_official_radar_row(r) for _,r in g.iterrows())
    source_count=max(1,g['Domain'].astype(str).replace('',pd.NA).dropna().nunique()) if 'Domain' in g.columns else 1
    if not official:
        qs.append('Resmî/birincil kaynak açıklaması')
    if source_count<2:
        qs.append('İkinci bağımsız kaynaktan teyit')

    if not qs:
        qs=[
            'Gelişmenin kapsamının yeni açıklamalarla netleşmesi',
            'Resmî/birincil kaynak teyidi',
            'Sanayi/teknoloji alanındaki somut etkisinin izlenmesi'
        ]

    # Sıralı tekilleştirme, en fazla 4 açık husus.
    out=[]
    seen=set()
    for q in qs:
        k=norm(q)
        if k in seen: continue
        seen.add(k); out.append(q)
        if len(out)>=4: break
    return out

def _v58_open_issues_table(df,limit=20):
    cols=['Tarih','Başlık','Aşama','Takip_Edilecek_Açık_Hususlar','Risk_Skoru','Doğrulama','URL']
    groups=_v58_event_groups(df)
    rows=[]
    for oid,g in groups.items():
        latest=g.sort_values('Tarih_dt',ascending=False).iloc[0]
        stage=_v58_event_stage(g)
        # Sonuçlanan olaylar açık hususlar listesinin altında kalsın; aktif olaylar öne çıksın.
        qs=_v58_open_questions_for_group(g)
        risk=int(pd.to_numeric(g['Risk_Skoru'],errors='coerce').fillna(0).max()) if 'Risk_Skoru' in g.columns else 0
        rows.append({
            'Tarih':latest.get('Tarih',''),
            'Başlık':latest.get('Başlık',''),
            'Aşama':stage,
            'Takip_Edilecek_Açık_Hususlar':' • '.join(qs),
            'Risk_Skoru':risk,
            'Doğrulama':latest.get('Doğrulama',''),
            'URL':latest.get('URL',''),
            '_active':0 if stage=='✅ Sonuçlandı' else 1,
            '_dt':pd.to_datetime(latest.get('Tarih_dt'),utc=True,errors='coerce')
        })
    if not rows:
        return pd.DataFrame(columns=cols)
    out=pd.DataFrame(rows).sort_values(['_active','Risk_Skoru','_dt'],ascending=[False,False,False])
    return out.head(limit).drop(columns=['_active','_dt'],errors='ignore')

def _v58_numeric_claims(g):
    """Kaynak bazında başlık+özetten sayısal iddiaları çıkarır."""
    claims=[]
    pat=re.compile(r'(?:%\s*)?\b\d+(?:[.,]\d+)?\b(?:\s*%|\s*(?:milyon|milyar|bin|adet|kişi|yaralı|ölü|mw|gw|ton|tl|dolar|euro|avro))?',re.I)
    for _,r in g.iterrows():
        txt=f"{r.get('Başlık','')} {r.get('İçerik_Özeti','')}"
        nums={x.strip() for x in pat.findall(txt) if x.strip()}
        claims.append((str(r.get('Kaynak','')),nums))
    return claims

def _v58_conflict_status(g):
    source_count=max(1,g['Domain'].astype(str).replace('',pd.NA).dropna().nunique()) if 'Domain' in g.columns else 1
    official=any(_is_official_radar_row(r) for _,r in g.iterrows())
    claims=_v58_numeric_claims(g)

    nonempty=[nums for _,nums in claims if nums]
    numeric_conflict=False
    if len(nonempty)>=2:
        # Birden fazla kaynağın sayısal kümeleri tamamen ayrışıyorsa uyar.
        for i in range(len(nonempty)):
            for j in range(i+1,len(nonempty)):
                if nonempty[i] and nonempty[j] and nonempty[i].isdisjoint(nonempty[j]):
                    numeric_conflict=True
                    break
            if numeric_conflict: break

    text=norm(' '.join(
        (g['Başlık'].fillna('').astype(str)+' '+g['İçerik_Özeti'].fillna('').astype(str)).tolist()
    ))
    verbal_conflict=(
        ('can kaybı yok' in text and ('can kaybı' in text.replace('can kaybı yok','') or 'hayatını kaybetti' in text))
        or ('yaralı yok' in text and 'yaralandı' in text)
        or ('üretim durdu' in text and ('üretim devam ediyor' in text or 'üretim sürüyor' in text))
    )

    if numeric_conflict or verbal_conflict:
        return '🔴 Çelişkili Bilgi','Kaynaklar arasında sayı/olgu farklılığı tespit edildi; manuel teyit önerilir.'
    if official:
        return '🟢 Resmî Teyitli','Resmî/birincil kaynak mevcut.'
    if source_count>=2:
        return '🟢 Çoklu Kaynak','En az iki farklı kaynak aynı olayı destekliyor.'
    return '🟡 Tek Kaynak','İkinci bağımsız veya resmî teyit henüz görünmüyor.'

def _v58_verification_matrix(df,limit=25):
    cols=['Tarih','Başlık','Teyit_Durumu','Teyit_Açıklaması','Kaynak_Sayısı',
          'Haber_Sayısı','Risk_Skoru','URL']
    groups=_v58_event_groups(df)
    rows=[]
    rank={'🔴 Çelişkili Bilgi':4,'🟡 Tek Kaynak':3,'🟢 Çoklu Kaynak':2,'🟢 Resmî Teyitli':1}
    for oid,g in groups.items():
        latest=g.sort_values('Tarih_dt',ascending=False).iloc[0]
        status,reason=_v58_conflict_status(g)
        source_count=max(1,g['Domain'].astype(str).replace('',pd.NA).dropna().nunique()) if 'Domain' in g.columns else 1
        risk=int(pd.to_numeric(g['Risk_Skoru'],errors='coerce').fillna(0).max()) if 'Risk_Skoru' in g.columns else 0
        rows.append({
            'Tarih':latest.get('Tarih',''),
            'Başlık':latest.get('Başlık',''),
            'Teyit_Durumu':status,
            'Teyit_Açıklaması':reason,
            'Kaynak_Sayısı':source_count,
            'Haber_Sayısı':len(g),
            'Risk_Skoru':risk,
            'URL':latest.get('URL',''),
            '_rank':rank.get(status,0),
            '_dt':pd.to_datetime(latest.get('Tarih_dt'),utc=True,errors='coerce')
        })
    if not rows:
        return pd.DataFrame(columns=cols)
    out=pd.DataFrame(rows).sort_values(['_rank','Risk_Skoru','_dt'],ascending=[False,False,False])
    return out.head(limit).drop(columns=['_rank','_dt'],errors='ignore')

# -----------------------------
# V51 — RESMÎ AÇIKLAMA / MEDYA KARŞILAŞTIRMASI
# -----------------------------
_COMPARE_STOP={
    've','ile','bir','bu','şu','için','da','de','mi','mı','mu','mü','olan','olarak',
    'son','yeni','göre','daha','çok','ise','ile','ancak','fakat','tarafından','dedi',
    'açıkladı','açıklama','haber','gelişme','türkiye','türk'
}

def _compare_tokens(text):
    t=norm(text)
    toks=re.findall(r'[a-z0-9çğıöşü]{3,}',t)
    return {x for x in toks if x not in _COMPARE_STOP}

def _event_similarity(a,b):
    """Başlık + kısa içerik üzerinden hızlı olay benzerliği; ağ isteği yapmaz."""
    at=_compare_tokens(f"{a.get('Başlık','')} {str(a.get('İçerik_Özeti',''))[:500]}")
    bt=_compare_tokens(f"{b.get('Başlık','')} {str(b.get('İçerik_Özeti',''))[:500]}")
    if not at or not bt: return 0.0
    inter=len(at & bt)
    union=max(1,len(at | bt))
    j=inter/union
    title_a=_compare_tokens(a.get('Başlık',''))
    title_b=_compare_tokens(b.get('Başlık',''))
    tj=len(title_a & title_b)/max(1,min(len(title_a),len(title_b))) if title_a and title_b else 0
    return 0.55*tj+0.45*j

def _short_claim(r,limit=220):
    txt=_clean_note_text(r.get('İçerik_Özeti',''))
    if not txt or norm(txt)==norm(r.get('Başlık','')):
        txt=_clean_note_text(r.get('Başlık',''))
    sents=_sentence_chunks(txt)
    if sents:
        txt=' '.join(sents[:2])
    return txt[:limit].strip()

def _comparison_difference(media,official):
    """İki kısa metindeki belirgin yön/iddia farklarını özetler; LLM/ağ çağrısı yok."""
    mt=norm(f"{media.get('Başlık','')} {media.get('İçerik_Özeti','')}")
    ot=norm(f"{official.get('Başlık','')} {official.get('İçerik_Özeti','')}")
    pairs=[
        (['tamamen durdu','üretim durdu','faaliyet durdu'],['kısmi','belirli bölüm','geçici','kısa süre','devam ediyor'],'Medya daha geniş bir durma/aksama bildirirken resmî açıklama etkinin kısmi veya geçici olduğunu belirtiyor.'),
        (['yangın','patlama','kaza'],['kontrol altına','söndürüldü','müdahale edildi'],'Resmî açıklama olayın kontrol/müdahale durumuna ilişkin ek bilgi içeriyor.'),
        (['can kaybı','öldü','hayatını kaybetti'],['can kaybı yok','can kaybı bulunmuyor'],'Can kaybına ilişkin medya ve resmî açıklama arasında farklı ifade bulunuyor.'),
        (['yaralı','yaralandı'],['yaralı yok','yaralanan yok'],'Yaralanma bilgisine ilişkin farklı ifade bulunuyor.'),
        (['veri sızıntısı','veri ihlali'],['etkilenmedi','sınırlı','belirli kullanıcı'],'Resmî açıklama olayın kapsamını medya anlatımına göre sınırlandırıyor/netleştiriyor.'),
        (['kriz','tehlike','alarm'],['normal','rutin','planlandığı','devam ediyor'],'Medya daha olumsuz/uyarıcı bir çerçeve kullanırken resmî açıklama daha sınırlı veya olağan bir durum tarif ediyor.')
    ]
    for mkeys,okeys,msg in pairs:
        if any(x in mt for x in mkeys) and any(x in ot for x in okeys):
            return msg

    mn=set(re.findall(r'(?:%\s*)?\d+(?:[.,]\d+)?',mt))
    on=set(re.findall(r'(?:%\s*)?\d+(?:[.,]\d+)?',ot))
    if mn and on and mn!=on:
        return 'Medya ve resmî açıklamada yer alan sayısal bilgiler farklılık gösteriyor; rakamların ayrıca kontrol edilmesi önerilir.'

    return 'Aynı olaya ilişkin resmî açıklama bulundu. Belirgin bir çelişki otomatik olarak tespit edilmedi; ayrıntılar birlikte kontrol edilebilir.'

def _official_media_comparison(df):
    """
    Sabit panel için medya haberlerini aynı taramadaki resmî/birincil içeriklerle eşleştirir.
    Ek web isteği yoktur; mevcut Resmî Kaynak Radarı verisini kullanır.
    """
    cols=['Tarih','Medya_Kaynağı','Medya_Haberi','Resmî_Kaynak','Resmî_Açıklama',
          'Karşılaştırma','Eşleşme','Medya_URL','Resmî_URL']
    if df is None or df.empty:
        return pd.DataFrame(columns=cols)

    officials=df[df.apply(_is_official_radar_row,axis=1)].copy()
    media=df[~df.apply(_is_official_radar_row,axis=1)].copy()
    if officials.empty or media.empty:
        return pd.DataFrame(columns=cols)

    rows=[]
    # Performans için resmî havuz zaten küçüktür; medya tarafında en yeni 250 içerik yeterli.
    media=media.sort_values('Tarih_dt',ascending=False).head(250)
    officials=officials.sort_values('Tarih_dt',ascending=False).head(80)

    for _,m in media.iterrows():
        best=None; best_score=0.0
        mdt=pd.to_datetime(m.get('Tarih_dt'),utc=True,errors='coerce')
        for _,o in officials.iterrows():
            odt=pd.to_datetime(o.get('Tarih_dt'),utc=True,errors='coerce')
            if pd.notna(mdt) and pd.notna(odt):
                if abs((mdt-odt).total_seconds()) > 72*3600:
                    continue
            score=_event_similarity(m,o)
            if score>best_score:
                best_score=score; best=o
        # Yanlış eşleşmeyi azaltmak için ölçülü eşik.
        if best is None or best_score<0.30:
            continue

        rows.append({
            'Tarih':m.get('Tarih',''),
            'Medya_Kaynağı':m.get('Kaynak',''),
            'Medya_Haberi':_short_claim(m),
            'Resmî_Kaynak':best.get('Kaynak',''),
            'Resmî_Açıklama':_short_claim(best),
            'Karşılaştırma':_comparison_difference(m,best),
            'Eşleşme':int(round(best_score*100)),
            'Medya_URL':m.get('URL',''),
            'Resmî_URL':best.get('URL','')
        })

    if not rows:
        return pd.DataFrame(columns=cols)
    out=pd.DataFrame(rows).drop_duplicates(subset=['Medya_URL','Resmî_URL'])
    return out.sort_values(['Eşleşme','Tarih'],ascending=[False,False])

def _contains_number_or_rate(text):
    t=str(text or '')
    return bool(re.search(
        r'(?<!\w)(?:%\\s*)?\\d+(?:[.,]\\d+)?(?:\\s*%|\\s*(?:milyon|milyar|trilyon|bin|adet|ton|mw|gw|gwh|twh|tl|₺|dolar|euro|avro))?',
        t,flags=re.I
    ))

def _critical_numbers(text, limit=4):
    t=re.sub(r'\\s+',' ',str(text or ''))
    pats=re.findall(
        r'(?:%\\s*\\d+(?:[.,]\\d+)?|\\d+(?:[.,]\\d+)?\\s*%|'
        r'\\d+(?:[.,]\\d+)?\\s*(?:milyon|milyar|trilyon|bin)\\s*(?:TL|₺|dolar|euro|avro)?|'
        r'\\d+(?:[.,]\\d+)?\\s*(?:MW|GW|GWh|TWh|ton|adet))',
        t,flags=re.I
    )
    out=[]
    for p in pats:
        p=p.strip()
        if p and p not in out:
            out.append(p)
        if len(out)>=limit:
            break
    return ', '.join(out)

def _important_statistics_rows(df):
    """Bugün yayımlanan, sanayi/teknoloji açısından sayısal veri taşıyan içerikleri seçer."""
    if df is None or df.empty:
        return pd.DataFrame()

    x=df.copy()
    x['Tarih_dt']=pd.to_datetime(x.get('Tarih_dt'),utc=True,errors='coerce')
    local_tz=datetime.now().astimezone().tzinfo
    today_local=datetime.now().astimezone().date()

    def is_today(v):
        try:
            return v is not None and pd.notna(v) and v.tz_convert(local_tz).date()==today_local
        except Exception:
            return False

    def stat_match(r):
        text=norm(f"{r.get('Başlık','')} {r.get('İçerik_Özeti','')} {r.get('Kategori','')}")
        term_hit=any(term in text for term in STATISTIC_TERMS)
        number_hit=_contains_number_or_rate(f"{r.get('Başlık','')} {r.get('İçerik_Özeti','')}")
        return term_hit and number_hit

    mask=x.apply(stat_match,axis=1)
    today_mask=x['Tarih_dt'].apply(is_today)
    result=x[mask & today_mask].copy()

    # Eğer yayın saati eksik gelmişse ama resmî/statistik kaynağı ve veri içeriği varsa dışarıda bırakma.
    missing_date=x['Tarih_dt'].isna()
    fallback=x[mask & missing_date & x.apply(_is_official_radar_row,axis=1)].copy()
    result=pd.concat([result,fallback],ignore_index=False).drop_duplicates(subset=['URL','Başlık'])

    if result.empty:
        return result

    result['Kritik_Sayı']=result.apply(
        lambda r:_critical_numbers(f"{r.get('Başlık','')} {r.get('İçerik_Özeti','')}"),
        axis=1
    )
    result['Birincil_Kaynak']=result.apply(lambda r:'✅' if _is_official_radar_row(r) else '—',axis=1)
    result=result.sort_values('Tarih_dt',ascending=False,na_position='last')
    return result

def _official_radar_rows(df):
    if df is None or df.empty:
        return pd.DataFrame()
    x=df[df.apply(_is_official_radar_row,axis=1)].copy()
    if x.empty:
        return x
    x=x.sort_values('Tarih_dt',ascending=False,na_position='last')
    return x.drop_duplicates(subset=['URL','Başlık'])

def _two_sentence_summary(text):
    sents=_detail_sentences(str(text or ''),'')
    if not sents:
        raw=_clean_note_text(text)
        return raw[:500]
    return ' '.join(sents[:2])

def _presentation_candidates(df,n=5):
    """Sunuma girmeye değer 5 başlık: stratejik önem + risk + resmîlik + sayısal veri + güncellik."""
    if df is None or df.empty:
        return pd.DataFrame()
    x=df.copy()
    x['Tarih_dt']=pd.to_datetime(x.get('Tarih_dt'),utc=True,errors='coerce')

    def score(r):
        text=norm(f"{r.get('Başlık','')} {r.get('İçerik_Özeti','')} {r.get('Kategori','')}")
        s=int(r.get('Risk_Skoru',0) or 0)//3
        if r.get('Risk_Durumu')=='Yüksek Risk': s+=18
        if _is_official_radar_row(r): s+=18
        if _contains_number_or_rate(text): s+=8
        if any(k in text for k in ['yatırım','ihracat','kapasite','savunma','yarı iletken','çip','yapay zeka',
                                   'enerji','otomotiv','uzay','ar-ge','arge','üretim','teşvik']): s+=14
        if critical_industrial_incident(r.get('Başlık',''),r.get('İçerik_Özeti','')): s+=16
        try: s+=min(int(r.get('Olay_Kaynak_Sayisi',0) or 0)*3,12)
        except Exception: pass
        return s

    x['_Sunum_Puanı']=x.apply(score,axis=1)
    x=x.sort_values(['_Sunum_Puanı','Tarih_dt'],ascending=[False,False],na_position='last')
    if 'Olay_ID' in x.columns:
        x=x.drop_duplicates(subset=['Olay_ID'],keep='first')
    else:
        x=x.drop_duplicates(subset=['Başlık'],keep='first')
    x=x.head(n).copy()
    x['Sunum_Başlığı']=x['Başlık'].astype(str)
    x['2_Cümle_Özet']=x['İçerik_Özeti'].apply(_two_sentence_summary)
    x['Kritik_Sayı']=x.apply(
        lambda r:_critical_numbers(f"{r.get('Başlık','')} {r.get('İçerik_Özeti','')}") or '—',
        axis=1
    )
    return x.drop(columns=['_Sunum_Puanı'],errors='ignore')

def build_negative_queries(when):
    return [
        f'Türkiye (iflas OR konkordato OR "üretim durdu" OR "fabrika kapandı" OR "işten çıkarma" OR grev OR soruşturma OR dava OR ceza OR "geri çağırma" OR "siber saldırı" OR "veri sızıntısı" OR yaptırım OR ambargo OR "ihale iptal" OR ertelendi OR gecikme OR "tedarik krizi" OR daralma OR zafiyet OR usulsüzlük OR yolsuzluk) (sanayi OR teknoloji OR üretim OR fabrika OR savunma OR otomotiv OR enerji OR şirket OR tesis OR proje) when:{when}',
        f'Türkiye ((OSB OR "organize sanayi" OR fabrika OR tesis OR "sanayi sitesi") (yangın OR yangını OR alev OR patlama OR patladı OR infilak)) when:{when}'
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

def _repair_mojibake_utf8(text):
    """
    'TÃ¼rkiye', 'genÃ§', 'katÄ±lÄ±m', 'baÅarÄ±' gibi UTF-8'in yanlış
    Latin-1/Windows-1252 olarak çözülmesinden doğan bozulmaları düzeltir.
    Doğru Türkçe metne dokunmamaya çalışır.
    """
    s=str(text or '')
    if not s:
        return s

    suspicious=('Ã','Ä','Å','Â','â€','â€™','â€œ','â€','â€“','â€”','\x80','\x81','\x8d','\x8f','\x90','\x9d','\x9f')
    if not any(x in s for x in suspicious):
        return s

    # Önce Windows-1252 mojibake işaretlerini bayt değerlerine geri çevirebilmek
    # için özel karakter -> byte haritası oluştur.
    cp1252_rev={}
    for b in range(256):
        try:
            ch=bytes([b]).decode('cp1252')
            cp1252_rev[ch]=b
        except Exception:
            pass

    def char_to_byte(ch):
        o=ord(ch)
        if o <= 255:
            return o
        return cp1252_rev.get(ch)

    # UTF-8 olabilecek bayt dizilerini parça parça düzelt; doğru Unicode
    # karakterler (ör. gerçek “ ’ ğ ş) sınır olarak korunur.
    out=[]
    buf=[]
    def flush():
        nonlocal buf
        if not buf:
            return
        raw=bytes(buf)
        original=''.join(chr(b) for b in buf)
        try:
            decoded=raw.decode('utf-8')
            # Yalnız gerçekten mojibake işaretlerini azaltıyorsa kabul et.
            before=sum(original.count(x) for x in ('Ã','Ä','Å','Â'))
            after=sum(decoded.count(x) for x in ('Ã','Ä','Å','Â'))
            out.append(decoded if after < before else original)
        except Exception:
            out.append(original)
        buf=[]

    for ch in s:
        b=char_to_byte(ch)
        if b is None:
            flush()
            out.append(ch)
        else:
            buf.append(b)
    flush()
    fixed=''.join(out)

    # Çok katmanlı bozulma varsa en fazla iki tur daha dene.
    for _ in range(2):
        if not any(x in fixed for x in ('Ã','Ä','Å','Â')):
            break
        try:
            candidate=fixed.encode('latin1').decode('utf-8')
            if sum(candidate.count(x) for x in ('Ã','Ä','Å','Â')) < sum(fixed.count(x) for x in ('Ã','Ä','Å','Â')):
                fixed=candidate
            else:
                break
        except Exception:
            break
    return fixed

def _clean_note_text(value):
    """
    V78 Word-safe metin temizliği:
    - mojibake'i bayt düzeyinde onarır,
    - Türkçe karakterleri Unicode NFC biçiminde korur,
    - DOCX/XML açısından sorunlu kontrol/görünmez karakterleri temizler.
    """
    import html as _html
    import unicodedata as _unicodedata

    text=BeautifulSoup(str(value or ''),'html.parser').get_text(' ',strip=True)
    text=_html.unescape(text)
    text=_repair_mojibake_utf8(text)

    # Kalan yaygın tipografik bozulmalar.
    replacements={
        'â€™':'’','â€˜':'‘','â€œ':'“','â€':'”',
        'â€“':'–','â€”':'—','â€¦':'…','Â ':' ','Â':''
    }
    for bad,good in replacements.items():
        text=text.replace(bad,good)

    for bad in ('\u00ad','\u200b','\u200c','\u200d','\ufeff'):
        text=text.replace(bad,'')

    # XML 1.0 geçersiz kontrol karakterlerini at.
    text=''.join(
        ch for ch in text
        if ch in ('\t','\n','\r') or ord(ch)>=32
    )

    text=_unicodedata.normalize('NFC',text)
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
    """
    V66: Başlıksız üç aşamalı bilgi notu yapısı:
    1) İlk paragraf kısa özet,
    2) devam eden paragraf(lar) ayrıntı/rakam/istatistik/gelişme,
    3) son paragraf sonuç ve kısa değerlendirme.
    Metin daima 'Arz olunur.' ile tamamlanır.
    """
    doc=Document()
    sec=doc.sections[0]
    sec.top_margin=Cm(2); sec.bottom_margin=Cm(2)
    sec.left_margin=Cm(2.5); sec.right_margin=Cm(2.5)
    styles=doc.styles
    styles['Normal'].font.name='Times New Roman'; styles['Normal'].font.size=Pt(12)
    styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'),'Times New Roman')

    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run(_clean_note_text(title)); r.bold=True; r.font.size=Pt(14)
    p=doc.add_paragraph(); p.add_run('Tarih: ').bold=True
    p.add_run(datetime.now().astimezone().strftime('%d.%m.%Y'))

    enriched=[]
    x=df.copy() if df is not None else pd.DataFrame()
    if 'Tarih_dt' in x.columns:
        x['Tarih_dt']=pd.to_datetime(x['Tarih_dt'],utc=True,errors='coerce')
        x=x.sort_values('Tarih_dt',ascending=True,na_position='last')
    for _,rr in x.iterrows():
        row=rr.to_dict()
        try:
            detail=article_detail(row)
        except Exception:
            detail={}
        enriched.append((row,detail))

    all_sent=[]
    for row,detail in enriched:
        title_text=_clean_note_text(detail.get('title') or row.get('Başlık',''))
        body=_clean_note_text(detail.get('text') or row.get('İçerik_Özeti') or title_text)
        all_sent.extend(_akt_clean_sentences(title_text,body))

    # Yakın tekrarları temizle.
    uniq=[]; seen=[]
    for sent in all_sent:
        sent=_clean_note_text(sent)
        key=norm(sent)
        toks=set(key.split())
        if not key: continue
        dup=False
        for old in seen[-35:]:
            union=len(toks|old)
            if union and len(toks&old)/union>=0.78:
                dup=True; break
        if not dup:
            uniq.append(sent.strip()); seen.append(toks)

    def add_body(text):
        bp=doc.add_paragraph()
        bp.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        bp.paragraph_format.first_line_indent=Cm(1.25)
        bp.paragraph_format.line_spacing=1.15
        bp.paragraph_format.space_after=Pt(8)
        safe_text=_repair_mojibake_utf8(_clean_note_text(text))
        bp.add_run(_v66_formalize_sentence_endings(safe_text))

    if uniq:
        # İlk paragraf: haberin kısa özeti. Başlık yazılmaz.
        intro=_join_sentences_naturally(uniq[:2])
        add_body(intro)

        # Gelişme bölümü: başlık kullanılmadan, ayrıntı/rakam/istatistikler korunarak devam eder.
        detail_s=uniq[2:]
        if not detail_s:
            detail_s=uniq

        # Uzun haberlerde ayrıntıları iki paragraf halinde dağıtarak okunabilirliği koru.
        detail_s=detail_s[:18]
        if len(detail_s)<=9:
            add_body(_join_sentences_naturally(detail_s))
        else:
            add_body(_join_sentences_naturally(detail_s[:9]))
            add_body(_join_sentences_naturally(detail_s[9:18]))

        # Son paragraf: sonuç + kısa/temkinli değerlendirme; ayrı başlık yoktur.
        tail=_join_sentences_naturally(uniq[-3:])
        if tail:
            conclusion=(
                f"Mevcut bilgiler çerçevesinde, {tail[0].lower()+tail[1:]} "
                "Gelişmenin sanayi ve teknoloji alanındaki muhtemel etkilerinin, ilgili kurum ve kuruluşların "
                "yeni açıklamaları ile resmî veriler doğrultusunda takip edilmesinin uygun olacağı değerlendirilmektedir."
            )
        else:
            conclusion=(
                "Mevcut bilgiler çerçevesinde gelişmenin sanayi ve teknoloji alanındaki etkilerinin, ilgili kurum "
                "ve kuruluşların yeni açıklamaları ile resmî veriler doğrultusunda takip edilmesinin uygun olacağı değerlendirilmektedir."
            )
        add_body(conclusion)
    else:
        add_body('Seçilen habere ilişkin ayrıntılı içerik temin edilememiştir.')
        add_body(
            'Gelişmenin yeni açık kaynak bilgileri ile ilgili kurum ve kuruluşların resmî açıklamaları '
            'doğrultusunda takip edilmesinin uygun olacağı değerlendirilmektedir.'
        )

    endp=doc.add_paragraph()
    endp.paragraph_format.space_before=Pt(8)
    endp.add_run('Arz olunur.')

    if enriched:
        kp=doc.add_paragraph()
        kr=kp.add_run('Kaynak: '); kr.bold=True
        for i,(row,detail) in enumerate(enriched):
            source=_clean_note_text(detail.get('source') or row.get('Kaynak','Açık Kaynak'))
            url=detail.get('canonical') or row.get('Yayıncı_URL') or row.get('URL','')
            if i: kp.add_run('; ')
            kp.add_run(source)
            if url:
                kp.add_run(' ('); _word_hyperlink(kp,url,'Haber linki'); kp.add_run(')')

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
                CREATE TABLE IF NOT EXISTS presentation_basket(
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    added_at TEXT NOT NULL,
                    news_time TEXT,
                    title TEXT NOT NULL,
                    source TEXT,
                    url TEXT,
                    category TEXT,
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
            conn.execute("""
                CREATE TABLE IF NOT EXISTS app_visits(
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    visited_at TEXT NOT NULL
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS note_history(
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    created_at TEXT NOT NULL,
                    title TEXT NOT NULL,
                    url TEXT,
                    UNIQUE(url,title)
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS tomorrow_followup(
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    added_at TEXT NOT NULL,
                    title TEXT NOT NULL,
                    source TEXT,
                    url TEXT,
                    category TEXT,
                    reason TEXT,
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
            change_map[c['Başlık']]=c.get('Değişim',c.get('Tür',''))

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
        new_events=int(changes['Tür'].astype(str).str.contains('YENİ OLAY').sum())
        risk_up=int(changes['Tür'].astype(str).str.contains('RİSK ARTTI').sum())
        verify_up=int(changes['Tür'].astype(str).str.contains('TEYİT').sum())
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
    rows=_v107_enrich_selected_rows(rows)
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
        if added:
            _v73_invalidate_status_cache()
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
    rows=_v107_enrich_selected_rows(rows)
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
        if added:
            _v73_invalidate_status_cache()
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
            if cur.rowcount:
                _v73_invalidate_status_cache()
            return cur.rowcount
    except Exception:
        return 0

def _clear_osint_basket():
    try:
        with _history_connect() as conn:
            cur=conn.execute("DELETE FROM osint_report_basket")
            conn.commit()
            if cur.rowcount:
                _v73_invalidate_status_cache()
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
            if cur.rowcount:
                _v73_invalidate_status_cache()
            return cur.rowcount
    except Exception:
        return 0

def _v81_sentence_case_title(title):
    t=_clean_note_text(title).strip()
    letters=''.join(c for c in t if c.isalpha())
    if letters and sum(c.isupper() for c in letters)/max(1,len(letters))>.80:
        t=t.lower()
        t=t[:1].upper()+t[1:]
    return t

def _v84_hard_repair_text(text):
    """
    V84: Türkçe olmayan/mojibake karakterleri agresif biçimde temizler.
    Tam onarılamayan bozuk cümleler ÖGN özetine hiç alınmaz.
    """
    t=_clean_note_text(text)

    # Ek yaygın bozulmalar.
    fixes={
        'TÃ¼rkiye':'Türkiye','TÃ¼rk':'Türk','genÃ§':'genç','dÃ¼nya':'dünya',
        'Ã¼lke':'ülke','Ã¼stÃ¼n':'üstün','Ã¶ÄŸrenci':'öğrenci','Ã¶Ärenci':'öğrenci',
        'baÅŸar':'başar','katÄ±lÄ±m':'katılım','mÃ¼cadele':'mücadele',
        'saÄŸladÄ±ÄŸÄ±':'sağladığı','saÄladÄÄ±ÄÄ±':'sağladığı',
        'ettiÄŸi':'ettiği','ettiÄi':'ettiği','TÃ¼rkiyenin':"Türkiye'nin",
        'TÃ¼rkiyeyi':"Türkiye'yi",'Ã§':'ç','ÄŸ':'ğ','Ä±':'ı',
        'Ã¶':'ö','Ã¼':'ü','ÅŸ':'ş','Ã‡':'Ç','Äž':'Ğ','Ä°':'İ','Ã–':'Ö','Ãœ':'Ü','Åž':'Ş'
    }
    for a,b in fixes.items():
        t=t.replace(a,b)

    # Kalan açık mojibake işaretleri varsa cümle güvenilmez kabul edilir.
    return _clean_note_text(t)

def _v84_sentence_is_clean(s):
    bad=('Ã','Ä','Å','Â',' ','\ufffd','','',' ')
    return not any(x in s for x in bad)

def _v84_clean_article_sentences(text):
    """Haber gövdesinden yalnız güvenilir, tam ve kurumsal özetlemeye uygun cümleleri alır."""
    text=_v84_hard_repair_text(text)
    garbage=[
        'çerez','cookie','reklam','devamını oku','tıklayın','anasayfa','son dakika',
        'benzer haber','ilgili haber','foto galeri','video galeri','sıralamayı değiştirmek',
        'kartları yukarı','abone ol','bildirimleri aç','google news','whatsapp kanal',
        'instagram','facebook','twitter','ekonomi gazetesi »','doğru şarj alışkanlıklarını',
        'haberler (','bugün kocaeli gazetesi','açıklaması şöyle','şunları kaydetti',
        'şöyle konuştu','şöyle dedi'
    ]
    out=[]; seen=set()
    for s in _sentence_split_tr(text):
        s=_v84_hard_repair_text(s).strip(" ;:-[]'\"")
        ns=norm(s)
        if not _v84_sentence_is_clean(s):
            continue
        if len(s)<38 or len(s)>480 or any(g in ns for g in garbage):
            continue
        if s.endswith(('…','...')) or re.search(r'\bve k$',s,re.I):
            continue
        # Haber ortasından alınmış doğrudan konuşma/alıntı ile başlama.
        if s.startswith(('"','“',"'",'‘')) or re.match(r'^\d+\s',s):
            continue
        letters=''.join(c for c in s if c.isalpha())
        if letters and len(s)<135 and sum(c.isupper() for c in letters)/max(1,len(letters))>.76:
            continue
        k=title_key(s)
        if not k or k in seen:
            continue
        seen.add(k); out.append(s)
    return out

def _v84_formalize(s):
    """Yalnız cümle sonunu değil, yaygın haber dili kalıntılarını da resmîleştirir."""
    s=_v84_hard_repair_text(s).strip()
    replacements=[
        (r'\bifade etti\b','ifade etmiştir'),(r'\bifade ediyor\b','ifade etmektedir'),
        (r'\bbelirtti\b','belirtmiştir'),(r'\bbelirtiyor\b','belirtmektedir'),
        (r'\baçıkladı\b','açıklamıştır'),(r'\baçıklıyor\b','açıklamaktadır'),
        (r'\bduyurdu\b','duyurmuştur'),(r'\bduyuruyor\b','duyurmaktadır'),
        (r'\bgösterdi\b','göstermiştir'),(r'\bgösteriyor\b','göstermektedir'),
        (r'\bsağladı\b','sağlamıştır'),(r'\bsağlıyor\b','sağlamaktadır'),
        (r'\bhedefliyor\b','hedeflemektedir'),(r'\bplanlıyor\b','planlamaktadır'),
        (r'\bbaşladı\b','başlamıştır'),(r'\bbaşlıyor\b','başlamaktadır'),
        (r'\btamamladı\b','tamamlamıştır'),(r'\btamamladı\b','tamamlamıştır'),
        (r'\bkazandı\b','kazanmıştır'),(r'\bgerçekleşti\b','gerçekleşmiştir'),
        (r'\byükseldi\b','yükselmiştir'),(r'\bgeriledi\b','gerilemiştir'),
        (r'\barttı\b','artmıştır'),(r'\bazaldı\b','azalmıştır'),
        (r'\boldu\b','olmuştur'),(r'\bolacak\b','olacaktır'),
        (r'\byapılacak\b','yapılacaktır'),(r'\bsağlanacak\b','sağlanacaktır'),
        (r'\bbaşlayacak\b','başlayacaktır'),(r'\byer alacak\b','yer alacaktır'),
        (r'\bmücadele edecek\b','mücadele edecektir')
    ]
    for pat,val in replacements:
        s=re.sub(pat,val,s,flags=re.I)
    s=_v66_formalize_sentence_endings(s)
    s=re.sub(r'\bifade ettiği ifade etmiştir\b','ifade etmiştir',s,flags=re.I)
    s=re.sub(r'\bbelirttiği belirtmiştir\b','belirtmiştir',s,flags=re.I)
    return _v84_hard_repair_text(s)

def _v84_score_intro(s):
    ns=norm(s)
    actor=['cumhurbaşkan','bakan','bakanlık','tüik','tübitak','tcmb','tse','türkpatent',
           'ssb','valili','başkan','şirket','üniversite','nasa','ibm','türk telekom',
           'kardemir','togg','gezeravcı','zeytinoğlu']
    action=['açıklad','duyur','başlat','gerçekleştir','tamamla','imzala','yayımla',
            'düzenlen','üret','geliştir','test','ziyaret','göreve','başvuru','yatırım']
    place=['ankara','istanbul','kocaeli','antalya','amasya','astana','pekin','gölcük',
           'türkiye','abd','çin','kazakistan','avustralya','almanya','isveç']
    return 4*sum(x in ns for x in actor)+4*sum(x in ns for x in action)+sum(x in ns for x in place)+min(len(re.findall(r'\d',s)),2)

def _v84_score_detail(s):
    ns=norm(s)
    data=['%','yüzde','milyon','milyar','bin ','adet','mw','gwh','mwh','km','puan',
          'oran','endeks','kapasite','ciro','ihracat','üretim','satış','başvuru','rekor']
    return 4*sum(x in ns for x in data)+min(len(re.findall(r'\d',s)),5)

def _v84_score_result(s):
    ns=norm(s)
    result=['art','azal','gerile','yüksel','ulaş','hedef','plan','beklen','sağla',
            'kazandır','devreye','pilot','kullanıl','rekor','destek','katkı','başarı']
    return 4*sum(x in ns for x in result)+min(len(re.findall(r'\d',s)),3)

def _v80_reference_important_summary(title,summary,full_text=''):
    """
    V84: Önce düzgün bir giriş cümlesi, sonra kritik rakam/detay, sonra sonuç/önem.
    2-3 tam cümle; cümle ortasında kesme yok; Word'de yaklaşık 4 satır hedefi.
    """
    title=_v81_sentence_case_title(_v84_hard_repair_text(title))
    body=_v84_hard_repair_text(full_text or summary)
    good=_v84_clean_article_sentences(body)

    if len(good)<2:
        good=_v84_clean_article_sentences(str(summary)+' '+str(full_text))
    if not good:
        return _v84_formalize(title)

    # Giriş asla haberin ortasından başlamasın: aktör + eylem taşıyan cümleyi seç.
    intro_candidates=good[:12]
    intro=max(intro_candidates,key=lambda s:(_v84_score_intro(s),-good.index(s)))
    if _v84_score_intro(intro)<4:
        # Güçlü giriş bulunamazsa ilk temiz cümleyi kullan.
        intro=good[0]

    chosen=[intro]

    rem=[s for s in good if s not in chosen]
    if rem:
        detail=max(rem,key=lambda s:(_v84_score_detail(s),-good.index(s)))
        if _v84_score_detail(detail)>0:
            chosen.append(detail)

    rem=[s for s in good if s not in chosen]
    if rem:
        result=max(rem,key=lambda s:(_v84_score_result(s),-good.index(s)))
        if _v84_score_result(result)>0:
            chosen.append(result)

    # En az iki cümle olsun.
    if len(chosen)<2:
        for s in good:
            if s not in chosen:
                chosen.append(s)
                break

    chosen=sorted(chosen,key=lambda s:good.index(s))
    formal=[_v84_formalize(s) for s in chosen[:3] if _v84_sentence_is_clean(_v84_formalize(s))]
    text=_clean_note_text(' '.join(formal))

    # Çok uzun cümleler nedeniyle 4 satırı aşmaması için sıkı sınır:
    # 2 veya 3 TAM cümle, yaklaşık 500 karakter.
    sents=_sentence_split_tr(text)
    kept=[]; total=0
    for sent in sents:
        add=len(sent)+(1 if kept else 0)
        if kept and total+add>500:
            break
        kept.append(sent); total+=add
        if len(kept)>=3:
            break

    # Eğer ilk cümle tek başına çok uzunsa, güvenli cümle sınırında sıkıştır.
    if kept and len(' '.join(kept))>520:
        kept=kept[:2]

    result=' '.join(kept).strip()

    # Son güvenlik: bozuk yabancı karakter kalırsa o cümleyi düşür.
    final_sents=[s for s in _sentence_split_tr(result) if _v84_sentence_is_clean(s)]
    return ' '.join(final_sents[:3]).strip()


def _v87_safe_tr(text):
    """Only obvious mojibake repair; never drop the whole item."""
    t=_clean_note_text(text)
    fixes={
        'TÃ¼rkiye':'Türkiye','TÃ¼rk':'Türk','genÃ§':'genç','dÃ¼nya':'dünya',
        'Ã¼lke':'ülke','Ã¼stÃ¼n':'üstün','Ã¶ÄŸrenci':'öğrenci','Ã¶Ärenci':'öğrenci',
        'baÅŸar':'başar','katÄ±lÄ±m':'katılım','mÃ¼cadele':'mücadele',
        'Ä±':'ı','ÄŸ':'ğ','ÅŸ':'ş','Ã§':'ç','Ã¶':'ö','Ã¼':'ü',
        'Ä°':'İ','Äž':'Ğ','Åž':'Ş','Ã‡':'Ç','Ã–':'Ö','Ãœ':'Ü',
        'Â':'','â€™':'’','â€œ':'“','â€':'”','â€“':'–','â€”':'—'
    }
    for a,b in fixes.items():
        t=t.replace(a,b)
    return re.sub(r'\s+',' ',t).strip()


@st.cache_data(ttl=3600,show_spinner=False)
def _v88_cached_article_detail(title,source,url,fallback,news_time):
    """Same article is not fetched again for one hour."""
    try:
        return article_detail({
            'Başlık':title,
            'Kaynak':source,
            'URL':url,
            'Yayıncı_URL':url,
            'İçerik_Özeti':fallback,
            'Tarih':news_time
        })
    except Exception:
        return {
            'title':title,'source':source,'canonical':url,
            'published':news_time,'text':fallback,'images':[]
        }

def _v88_title_core(title,source=''):
    """Remove publisher suffixes and headline clutter."""
    t=_v87_safe_tr(title)
    source=_v87_safe_tr(source)
    # Common Google News/source suffix.
    if source:
        t=re.sub(r'\s*[-–—]\s*'+re.escape(source)+r'\s*$','',t,flags=re.I)
    t=re.sub(r'\s*[-–—]\s*(Haberler|Haber|Son Dakika|Gündem)\s*$','',t,flags=re.I)
    t=re.sub(r'\s+',' ',t).strip(' -–—|')
    return t

def _v88_sentence_bad(s):
    s=_v87_safe_tr(s)
    bad_chars=('Ã','Ä','Å','Â',' ',' ','','')
    if any(x in s for x in bad_chars):
        return True
    n=norm(s)
    noise=[
        'sıralamayı değiştirmek','kartları yukarı','tüvtürk en sık',
        'samsung sevilen modelin','benzer haber','ilgili haber',
        'devamını oku','çerez','cookie','reklam','foto galeri','video galeri',
        'ekonomi gazetesi »','araç sahipleri dikkat'
    ]
    if any(x in n for x in noise):
        return True
    if s.endswith(('…','...')) or re.search(r'\bve k$',s,re.I):
        return True
    return False

def _v88_clean_sentences(text):
    out=[]; seen=set()
    for s in _sentence_chunks(_v87_safe_tr(text)):
        s=_v87_safe_tr(s).strip(" []'\";-:")
        if len(s)<35 or len(s)>430 or _v88_sentence_bad(s):
            continue
        k=title_key(s)
        if not k or k in seen:
            continue
        seen.add(k); out.append(s)
    return out

def _v88_keywords(title):
    stop={'haber','haberi','son','dakika','bugün','yeni','ile','ve','bir','için','olan','oldu',
          'olacak','dedi','açıkladı','duyurdu','türkiye','türk'}
    words=[w for w in re.findall(r'[a-zçğıöşü0-9]+',norm(title)) if len(w)>=4 and w not in stop]
    return set(words[:12])

def _v88_formal(s):
    s=_v87_safe_tr(s)
    pairs=[
        (r'\baçıkladı\b','açıklamıştır'),(r'\bbelirtti\b','belirtmiştir'),
        (r'\bduyurdu\b','duyurmuştur'),(r'\bkaydetti\b','kaydetmiştir'),
        (r'\bifade etti\b','ifade etmiştir'),(r'\bbaşladı\b','başlamıştır'),
        (r'\btamamladı\b','tamamlamıştır'),(r'\bkazandı\b','kazanmıştır'),
        (r'\barttı\b','artmıştır'),(r'\bazaldı\b','azalmıştır'),
        (r'\bgeriledi\b','gerilemiştir'),(r'\byükseldi\b','yükselmiştir'),
        (r'\bulaştı\b','ulaşmıştır'),(r'\bgerçekleşti\b','gerçekleşmiştir'),
        (r'\boldu\b','olmuştur'),(r'\byer alacak\b','yer alacaktır'),
        (r'\bbaşlayacak\b','başlayacaktır'),(r'\bsağlanacak\b','sağlanacaktır'),
        (r'\bverilecek\b','verilecektir'),(r'\bseçilecek\b','seçilecektir'),
        (r'\bkazandırılacak\b','kazandırılacaktır'),(r'\bdevam ediyor\b','devam etmektedir'),
        (r'\bgösteriyor\b','göstermektedir'),(r'\bsağlıyor\b','sağlamaktadır'),
        (r'\bdikkat çekiyor\b','dikkat çekmektedir')
    ]
    for pat,val in pairs:
        s=re.sub(pat,val,s,flags=re.I)
    s=_v66_formalize_sentence_endings(s)
    s=_v87_safe_tr(s).strip()
    if s:
        s=s[0].upper()+s[1:]
    return s

def _v89_normalize_source_title(title,source):
    """Başlıktaki yayıncı/portal eklerini temizler; başlığı çıktı olarak kullanmaz."""
    t=_v88_title_core(title,source)
    t=re.sub(r'\s+[A-Za-zÇĞİÖŞÜçğıöşü0-9_.-]+\.(?:com|com\.tr|net|org|tr)\s*$','',t,flags=re.I)
    return _v87_safe_tr(t).strip(' -–—|')

def _v89_clause_from_sentence(s):
    """
    İkinci bir haber cümlesini ana resmî cümleye eklenebilir bilgi cümleciğine çevirir.
    Tam cümleyi parçalamaz; yalnız son noktayı kaldırır.
    """
    s=_v88_formal(_v87_safe_tr(s)).strip()
    return s.rstrip(' .;:')

def _v89_single_official_sentence(title,source,body,fallback):
    """
    Gerçek STB örneği mantığı:
    Her gelişme için TEK, TAM ve RESMÎ cümle.
    - kim/kurum + ne oldu ana cümlesi,
    - en kritik rakam/yer/tarih aynı cümlede,
    - gerekiyorsa sonuç/hedef ikinci cümlecik olarak noktalı virgülle bağlanır,
    - haber başlığı tek başına çıktı olmaz.
    """
    title=_v89_normalize_source_title(title,source)
    text=_v87_safe_tr(body or fallback)
    sents=_v88_clean_sentences(text)
    if len(sents)<2:
        sents=_v88_clean_sentences(fallback)

    if not sents:
        # Son çare: kaydedilmiş özet varsa onu kullan; sırf başlığı basma.
        fb=_v87_safe_tr(fallback)
        if len(fb)>=60:
            return _v88_formal(fb).rstrip(' .;')+'.'
        return ''

    keywords=_v88_keywords(title)
    actor_terms=['cumhurbaşkan','bakan','bakanlık','başkan','tüik','tübitak','tcmb','tse',
                 'türkpatent','ssb','valili','üniversite','şirket','genel müdür','türk telekom',
                 'kardemir','togg','aselsan','roketsan','gezeravcı','zeytinoğlu','kurum','takım']
    action_terms=['açıkla','duyur','başlat','gerçekleştir','tamamla','imzala','kazan','yatırım',
                  'test','görev','üret','satış','başvuru','düzenlen','ulaş','art','azal','gerile']
    detail_terms=['%','yüzde','milyon','milyar','bin ','adet','mw','gwh','mwh','km','puan',
                  'kapasite','ihracat','üretim','satış','hibe','öğrenci','madalya','rekor','tarih']
    result_terms=['hedef','beklen','sağla','katkı','devreye','plan','başarı','destek','başvuru',
                  'artış','azalış','yüksel','gerile','ulaş']

    def overlap(sent):
        ws=set(re.findall(r'[a-zçğıöşü0-9]+',norm(sent)))
        return len(keywords & ws)

    # Konuyla ilişkisiz "Samsung / TÜVTÜRK / başka haber" parçalarını devreden çıkar.
    related=[x for x in sents if overlap(x)>0]
    pool=related if related else sents[:8]

    def intro_score(x):
        n=norm(x)
        return 7*overlap(x)+4*sum(k in n for k in actor_terms)+4*sum(k in n for k in action_terms)

    intro=max(pool[:8],key=lambda x:(intro_score(x),-sents.index(x)))
    intro_formal=_v89_clause_from_sentence(intro)

    # Başlıkla neredeyse aynıysa, başka giriş ara.
    if title_key(intro_formal)==title_key(title):
        alternatives=[x for x in pool if title_key(x)!=title_key(title)]
        if alternatives:
            intro=max(alternatives,key=lambda x:(intro_score(x),-sents.index(x)))
            intro_formal=_v89_clause_from_sentence(intro)

    rem=[x for x in pool if x!=intro]
    detail=None
    if rem:
        def detail_score(x):
            n=norm(x)
            return 6*overlap(x)+5*sum(k in n for k in detail_terms)+min(len(re.findall(r'\d',x)),6)
        cand=max(rem,key=lambda x:(detail_score(x),-sents.index(x)))
        if detail_score(cand)>0:
            detail=cand

    rem=[x for x in rem if x!=detail]
    result=None
    if rem:
        def result_score(x):
            n=norm(x)
            return 5*overlap(x)+4*sum(k in n for k in result_terms)+2*sum(k in n for k in detail_terms)
        cand=max(rem,key=lambda x:(result_score(x),-sents.index(x)))
        if result_score(cand)>0:
            result=cand

    # Ana cümle doğal biçimde zaten gerekli rakamları içeriyorsa gereksiz tekrar ekleme.
    clauses=[intro_formal]
    intro_digits=set(re.findall(r'\d+(?:[.,]\d+)?',intro_formal))

    for extra in [detail,result]:
        if not extra:
            continue
        ef=_v89_clause_from_sentence(extra)
        if not ef or _v88_sentence_bad(ef):
            continue
        # Aynı olayı/rakamı tekrar eden cümleyi alma.
        nums=set(re.findall(r'\d+(?:[.,]\d+)?',ef))
        if nums and nums.issubset(intro_digits) and title_key(ef)[:80] in title_key(intro_formal):
            continue
        # Başlıkla ilişki şartı: unrelated site snippets cannot enter.
        if keywords and overlap(ef)==0:
            continue
        clauses.append(ef)
        if len(clauses)>=2:  # tek cümlede iki ana bilgi bloğu yeterli
            break

    # Tek resmî cümle: ilk tam cümle + ikinci bilgi bloğu noktalı virgülle.
    if len(clauses)==1:
        out=clauses[0]
    else:
        second=clauses[1]
        # İkinci bloğu küçük harfle doğal bağla; özel isimleri bozma.
        connector='; ayrıca, '
        out=clauses[0]+connector+second

    out=_v87_safe_tr(out).strip(' ;:.')
    # 4 satır hedefi: cümleyi kesmeden 500 karaktere yaklaş.
    if len(out)>500 and len(clauses)>1:
        out=clauses[0].strip(' ;:.')
    if len(out)>520:
        # Çok uzun tek giriş varsa noktalı virgül/virgül sınırından kısalt.
        cut=out[:520]
        candidates=[cut.rfind('; '),cut.rfind(', ')]
        k=max(candidates)
        if k>=300:
            out=cut[:k].rstrip(' ,;')
    return _v87_safe_tr(out)+'.'

# Keep name used by make_important_basket_docx, but route to V89.
def _v88_summary(title,source,body,fallback):
    return _v89_single_official_sentence(title,source,body,fallback)

def _v87_ogn_summary(title, body, fallback):
    """
    Simple, deterministic recovery summarizer:
    - never returns blank when fallback exists,
    - uses stable _akt_formal_summary,
    - keeps 2-3 complete sentences where available,
    - no extra web search and no experimental sentence dropping.
    """
    title=_v87_safe_tr(title)
    body=_v87_safe_tr(body or fallback or title)
    fallback=_v87_safe_tr(fallback)

    try:
        text=_akt_formal_summary(title,body,max_sentences=3,max_chars=700)
    except Exception:
        text=fallback or body or title

    text=_v87_safe_tr(text)
    if not text or title_key(text)==title_key(title):
        text=fallback if len(fallback)>=60 else (body if len(body)>=60 else title)

    # Formalize endings using existing stable routine.
    text=_v66_formalize_sentence_endings(text)

    # Keep max 3 COMPLETE sentences and roughly 4 Word lines.
    sents=_sentence_chunks(text)
    if not sents:
        return text[:520].strip()

    kept=[]; total=0
    for sent in sents:
        sent=_v87_safe_tr(sent).strip()
        if not sent: continue
        add=len(sent)+(1 if kept else 0)
        if kept and total+add>520:
            break
        kept.append(sent)
        total+=add
        if len(kept)>=3:
            break

    out=' '.join(kept).strip()
    return out or (fallback[:520].strip() if fallback else title[:520].strip())


V90_OGN_ENGINE_VERSION='v104_mevcut_bolum_olgunlastirma_20260822_1'

def _v90_clean_title(title,source=''):
    t=_v87_safe_tr(title)
    s=_v87_safe_tr(source)
    if s:
        t=re.sub(r'\s*[-–—]\s*'+re.escape(s)+r'\s*$','',t,flags=re.I)
    t=re.sub(r'\s*[-–—]\s*(Haberler|Haber|Son Dakika|Gündem)\s*$','',t,flags=re.I)
    t=re.sub(r'\s+',' ',t).strip(' -–—|')
    return t

def _v90_clean_sentence(s):
    s=_v87_safe_tr(s).strip(" []'\";-:")
    # Haber portalı / başka başlık / yarım snippet artıkları.
    noise=[
        'sıralamayı değiştirmek','kartları yukarı','devamını oku','benzer haber',
        'ilgili haber','çerez','cookie','reklam','foto galeri','video galeri',
        'google news','whatsapp','instagram','facebook','twitter',
        'araç sahipleri dikkat','samsung sevilen modelin','tüvtürk en sık',
        'ekonomi gazetesi »'
    ]
    ns=norm(s)
    if any(x in ns for x in noise):
        return ''
    if s.endswith(('…','...')) or re.search(r'\bve k$',s,re.I):
        return ''
    if any(x in s for x in ('Ã','Ä','Å',' ',' ','','')):
        return ''
    return s

def _v90_sentences(text):
    out=[]; seen=set()
    for raw in _sentence_chunks(_v87_safe_tr(text)):
        s=_v90_clean_sentence(raw)
        if len(s)<38 or len(s)>520:
            continue
        k=title_key(s)
        if not k or k in seen:
            continue
        seen.add(k)
        out.append(s)
    return out

def _v90_title_words(title):
    stop={
        'haber','haberi','son','dakika','bugün','yeni','ile','ve','bir','için','olan',
        'oldu','olacak','dedi','açıkladı','duyurdu','türkiye','türk','etti','başladı'
    }
    return set(
        w for w in re.findall(r'[a-zçğıöşü0-9]+',norm(title))
        if len(w)>=4 and w not in stop
    )

def _v90_formalize(s):
    s=_v87_safe_tr(s)
    replacements=[
        (r'\baçıkladı\b','açıklamıştır'),(r'\bbelirtti\b','belirtmiştir'),
        (r'\bduyurdu\b','duyurmuştur'),(r'\bkaydetti\b','kaydetmiştir'),
        (r'\bifade etti\b','ifade etmiştir'),(r'\bbaşladı\b','başlamıştır'),
        (r'\btamamladı\b','tamamlamıştır'),(r'\bkazandı\b','kazanmıştır'),
        (r'\barttı\b','artmıştır'),(r'\bazaldı\b','azalmıştır'),
        (r'\bgeriledi\b','gerilemiştir'),(r'\byükseldi\b','yükselmiştir'),
        (r'\bulaştı\b','ulaşmıştır'),(r'\bgerçekleşti\b','gerçekleşmiştir'),
        (r'\boldu\b','olmuştur'),(r'\bkırıldı\b','kırılmıştır'),
        (r'\byer alacak\b','yer alacaktır'),(r'\bbaşlayacak\b','başlayacaktır'),
        (r'\bsağlanacak\b','sağlanacaktır'),(r'\bverilecek\b','verilecektir'),
        (r'\bseçilecek\b','seçilecektir'),(r'\bkazandırılacak\b','kazandırılacaktır'),
        (r'\bdevam ediyor\b','devam etmektedir'),(r'\bgösteriyor\b','göstermektedir'),
        (r'\bsağlıyor\b','sağlamaktadır'),(r'\bdikkat çekiyor\b','dikkat çekmektedir')
    ]
    for pat,val in replacements:
        s=re.sub(pat,val,s,flags=re.I)
    s=_v66_formalize_sentence_endings(s)
    s=_v87_safe_tr(s).strip()
    if s:
        s=s[0].upper()+s[1:]
    return s

def _v90_item_summary(title,source,body,fallback):
    """
    Kurum örneğine yakın TEK, TAM, RESMÎ cümle:
    ana olay + kritik rakam/yer/tarih/kişi + gerekiyorsa tek tamamlayıcı bilgi.
    Başlık doğrudan Word'e yazılmaz.
    """
    title=_v90_clean_title(title,source)
    text=_v87_safe_tr(body or fallback)
    sents=_v90_sentences(text)
    if len(sents)<2:
        sents=_v90_sentences(fallback)

    # Başlık dışında gerçek içerik yoksa fallback'i kullan; yine başlığı tek başına basma.
    if not sents:
        fb=_v87_safe_tr(fallback)
        if len(fb)>=80 and title_key(fb)!=title_key(title):
            return _v90_formalize(fb).rstrip(' .;')+'.'
        return ''

    tw=_v90_title_words(title)

    actor_terms=[
        'cumhurbaşkan','bakan','bakanlık','başkan','tüik','tübitak','tcmb','tse',
        'türkpatent','ssb','valili','üniversite','şirket','genel müdür','türk telekom',
        'kardemir','togg','aselsan','roketsan','gezeravcı','zeytinoğlu','takım'
    ]
    action_terms=[
        'açıkla','duyur','başlat','gerçekleştir','tamamla','imzala','kazan','yatırım',
        'test','görev','üret','satış','başvuru','düzenlen','ulaş','art','azal','gerile'
    ]
    detail_terms=[
        '%','yüzde','milyon','milyar','bin ','adet','mw','gwh','mwh','km','puan',
        'kapasite','ihracat','üretim','satış','hibe','öğrenci','madalya','rekor',
        '2025','2026','2027'
    ]

    def overlap(s):
        words=set(re.findall(r'[a-zçğıöşü0-9]+',norm(s)))
        return len(words & tw)

    # Yalnız haberle ilişkili cümleleri tercih et.
    related=[s for s in sents if overlap(s)>0]
    pool=related if related else sents[:8]

    def intro_score(s):
        n=norm(s)
        return (
            8*overlap(s)
            + 4*sum(x in n for x in actor_terms)
            + 4*sum(x in n for x in action_terms)
            + min(len(re.findall(r'\d',s)),3)
        )

    # İlk cümle haberin ortasından değil, olayı tanımlayan cümle olsun.
    intro=max(pool[:8],key=lambda s:(intro_score(s),-sents.index(s)))
    intro=_v90_formalize(intro).rstrip(' .;')

    # Eğer intro başlıkla neredeyse aynıysa başka gövde cümlesi dene.
    if title_key(intro)==title_key(title):
        alternatives=[s for s in pool if title_key(s)!=title_key(title)]
        if alternatives:
            intro=_v90_formalize(
                max(alternatives,key=lambda s:(intro_score(s),-sents.index(s)))
            ).rstrip(' .;')

    # En kritik ikinci bilgi: rakam/tarih/ölçek; aynı olayla ilişkili olmak zorunda.
    remaining=[s for s in pool if title_key(_v90_formalize(s))!=title_key(intro)]
    detail=''
    if remaining:
        def detail_score(s):
            n=norm(s)
            return (
                7*overlap(s)
                + 5*sum(x in n for x in detail_terms)
                + min(len(re.findall(r'\d',s)),6)
            )
        cand=max(remaining,key=lambda s:(detail_score(s),-sents.index(s)))
        if detail_score(cand)>=5:
            detail=_v90_formalize(cand).rstrip(' .;')

    # Tek resmî cümle oluştur.
    out=intro
    if detail:
        # Aynı rakamları tekrar eden ayrıntıyı ekleme.
        n1=set(re.findall(r'\d+(?:[.,]\d+)?',intro))
        n2=set(re.findall(r'\d+(?:[.,]\d+)?',detail))
        if not (n2 and n2.issubset(n1) and len(detail)<180):
            out += '; ayrıca, ' + detail[0].lower()+detail[1:] if detail else ''

    out=_v87_safe_tr(out).strip(' ;:.')

    # Örnekteki yoğunluk: yaklaşık 4 Word satırı; cümleyi ortadan kesme.
    if len(out)>500 and detail:
        out=intro
    if len(out)>520:
        # Giriş tek başına çok uzunsa en yakın anlamlı virgül/noktalı virgül sınırında kısalt.
        cut=out[:520]
        k=max(cut.rfind('; '),cut.rfind(', '))
        if k>=330:
            out=cut[:k].rstrip(' ,;')

    return out.rstrip(' .;')+'.'

@st.cache_data(ttl=3600,show_spinner=False)
def _v90_fetch_detail(title,source,url,fallback,news_time):
    try:
        return article_detail({
            'Başlık':title,
            'Kaynak':source,
            'URL':url,
            'Yayıncı_URL':url,
            'İçerik_Özeti':fallback,
            'Tarih':news_time
        })
    except Exception:
        return {'title':title,'source':source,'text':fallback,'canonical':url,'images':[]}


def _v92_clean_news_text(text):
    t=_v87_safe_tr(text)
    # Portal artıkları / başka haber başlıkları / navigasyon.
    noise_patterns=[
        r'sıralamayı değiştirmek[^.!?]*[.!?]?',
        r'kartları yukarı[^.!?]*[.!?]?',
        r'devamını oku[^.!?]*[.!?]?',
        r'benzer haber(?:ler)?[^.!?]*[.!?]?',
        r'ilgili haber(?:ler)?[^.!?]*[.!?]?',
        r'google news[^.!?]*[.!?]?',
        r'whatsapp kanal[^.!?]*[.!?]?',
    ]
    for pat in noise_patterns:
        t=re.sub(pat,' ',t,flags=re.I)
    return re.sub(r'\s+',' ',t).strip()

def _v92_formal_sentence(s):
    """Yalnız haber dili yüklemlerini kurumsal geçmiş/şimdiki zamana çevirir."""
    s=_v92_clean_news_text(s).strip()
    pairs=[
        (r'\baçıkladı\b','açıklamıştır'),
        (r'\bbelirtti\b','belirtmiştir'),
        (r'\bduyurdu\b','duyurmuştur'),
        (r'\bkaydetti\b','kaydetmiştir'),
        (r'\bifade etti\b','ifade etmiştir'),
        (r'\bbaşladı\b','başlamıştır'),
        (r'\btamamladı\b','tamamlamıştır'),
        (r'\bkazandı\b','kazanmıştır'),
        (r'\barttı\b','artmıştır'),
        (r'\bazaldı\b','azalmıştır'),
        (r'\bgeriledi\b','gerilemiştir'),
        (r'\byükseldi\b','yükselmiştir'),
        (r'\bulaştı\b','ulaşmıştır'),
        (r'\bgerçekleşti\b','gerçekleşmiştir'),
        (r'\boldu\b','olmuştur'),
        (r'\bkırıldı\b','kırılmıştır'),
        (r'\byayımlandı\b','yayımlanmıştır'),
        (r'\byayınlandı\b','yayımlanmıştır'),
        (r'\bbaşlıyor\b','başlamaktadır'),
        (r'\bdevam ediyor\b','devam etmektedir'),
        (r'\bsağlıyor\b','sağlamaktadır'),
        (r'\bgösteriyor\b','göstermektedir'),
        (r'\bhedefleniyor\b','hedeflenmektedir'),
        (r'\bplanlanıyor\b','planlanmaktadır'),
    ]
    for pat,val in pairs:
        s=re.sub(pat,val,s,flags=re.I)
    return _v66_formalize_sentence_endings(s).strip()

def _v92_good_sentence(s):
    s=_v92_clean_news_text(s)
    if len(s)<35 or len(s)>650:
        return False
    if any(x in s for x in ('Ã','Ä','Å',' ',' ','','')):
        return False
    n=norm(s)
    noise=[
        'çerez','cookie','reklam','foto galeri','video galeri',
        'instagram','facebook','twitter','abone ol','bildirimleri aç',
        'sıralamayı değiştirmek','kartları yukarı','devamını oku',
        'benzer haber','ilgili haber'
    ]
    return not any(x in n for x in noise)

def _v92_summary(title, source, body, fallback):
    """
    STB örneğindeki mantık:
    1) Haberin ana gelişmesini veren giriş cümlesi.
    2) Hemen ardından gelen kritik veri/rakam/açıklama.
    3) Gerekliyse üçüncü ardışık cümle.
    Farklı yerlerden cümle toplayıp yapıştırmaz.
    """
    title=_v90_clean_title(title,source)
    raw=_v92_clean_news_text(body or fallback)

    sents=[]
    seen=set()
    for s in _sentence_chunks(raw):
        s=_v92_clean_news_text(s).strip(" []'\";-:")
        if not _v92_good_sentence(s):
            continue
        k=title_key(s)
        if not k or k in seen:
            continue
        seen.add(k)
        sents.append(s)

    if not sents:
        fb=_v92_clean_news_text(fallback)
        return _v92_formal_sentence(fb) if len(fb)>=50 else ''

    # Başlıktaki ayırt edici sözcükler.
    stop={'haber','haberi','son','dakika','bugün','yeni','ile','ve','bir','için',
          'olan','oldu','olacak','dedi','türkiye','türk'}
    title_words={
        w for w in re.findall(r'[a-zçğıöşü0-9]+',norm(title))
        if len(w)>=4 and w not in stop
    }

    def overlap(s):
        sw=set(re.findall(r'[a-zçğıöşü0-9]+',norm(s)))
        return len(sw & title_words)

    # Giriş: ilk 6 temiz cümle içinde başlıkla en ilişkili olanı bul.
    # Böylece sayfanın ortasından başlamaz; ama başlık tekrarına da mahkûm olmaz.
    first_pool=sents[:6]
    start=max(range(len(first_pool)), key=lambda i:(overlap(first_pool[i]), -i))

    # Eğer ilk cümle zaten makul derecede ilgiliyse onu tercih et.
    if overlap(first_pool[0])>0:
        start=0

    # Ardışık cümleler: bağlam korunur. Maksimum 3 cümle.
    chosen=[]
    total=0
    for s in sents[start:start+3]:
        fs=_v92_formal_sentence(s).strip()
        if not fs:
            continue
        if fs[-1] not in '.!?':
            fs+='.'
        # Yaklaşık 4 satır; cümleyi ortadan kesme.
        add=len(fs)+(1 if chosen else 0)
        if chosen and total+add>560:
            break
        chosen.append(fs)
        total+=add

    # Tek cümle çok kısa kaldıysa bir sonraki ilgili tam cümleyi ekle.
    if len(' '.join(chosen))<220:
        for s in sents[start+len(chosen):]:
            if overlap(s)<=0 and title_words:
                continue
            fs=_v92_formal_sentence(s).strip()
            if not fs:
                continue
            if fs[-1] not in '.!?':
                fs+='.'
            if len(' '.join(chosen+[fs]))<=560:
                chosen.append(fs)
            break

    out=' '.join(chosen).strip()

    # Başlığı tek başına "özet" kabul etme.
    if title_key(out)==title_key(title):
        fb=_v92_clean_news_text(fallback)
        if len(fb)>=80 and title_key(fb)!=title_key(title):
            out=_v92_formal_sentence(fb)

    return _v87_safe_tr(out).strip()


def _v93_sentence_split(text):
    t=_v92_clean_news_text(text)
    # Noktalı virgülü cümle gibi bölme; kurumsal örneklerde bağlı bilgi olabilir.
    return [x.strip() for x in re.split(r'(?<=[.!?])\s+',t) if x.strip()]

def _v93_is_noise(s):
    n=norm(s)
    bad=[
        'sıralamayı değiştirmek','kartları yukarı','benzer haber','ilgili haber',
        'devamını oku','son dakika','reklam','çerez','cookie','instagram',
        'facebook','twitter','whatsapp','youtube','google news',
        'tıklayın','abone ol','yorumlar','etiketler'
    ]
    return len(s)<30 or any(x in n for x in bad) or any(x in s for x in ('Ã','Ä','Å',' '))

def _v93_content_score(s, title_words):
    n=norm(s)
    words=set(re.findall(r'[a-zçğıöşü0-9]+',n))
    score=0
    score += 5*len(words & title_words)
    # Örnek notlarda öne çıkan unsurlar: kurum/kişi, tarih, rakam, oran, yer, karar/eylem.
    if re.search(r'\b\d+(?:[.,]\d+)?\b|%|yüzde|milyon|milyar|bin\b',n): score+=5
    if re.search(r'\b20\d{2}\b|\b\d{1,2}\s+(?:ocak|şubat|mart|nisan|mayıs|haziran|temmuz|ağustos|eylül|ekim|kasım|aralık)\b',n): score+=3
    if any(x in n for x in ['açıklad','duyur','yayımla','başlat','tamamla','imzala','seçilecek',
                             'gerçekleştir','üret','ihrac','yatırım','satın al','devreye al',
                             'denize indir','entegre','belirlen','güncellen','başvuru']): score+=4
    if any(x in n for x in ['bakanlık','tübitak','tüik','aselsan','roketsan','tusaş','havelsan',
                             'cumhurbaşkan','başkanlığı','ajansı','üniversitesi','şirketi','ofisi']): score+=3
    return score

def _v93_build_summary(title, source, body, fallback):
    """
    Tek haber -> tek olay anlatısı.
    Cümleler haberin başından/ortasından rastgele seçilmez:
    ana olay cümlesi bulunur ve yalnız onun çevresindeki aynı bağlam kullanılır.
    """
    title=_v90_clean_title(title,source)
    text=_v92_clean_news_text(body or fallback)
    sentences=[s for s in _v93_sentence_split(text) if not _v93_is_noise(s)]
    if not sentences:
        sentences=[s for s in _v93_sentence_split(fallback) if not _v93_is_noise(s)]
    if not sentences:
        return ''

    stop={'haber','haberi','son','dakika','bugün','yeni','ile','ve','bir','için','olan',
          'oldu','olacak','dedi','türkiye','türk','tarafından','kapsamında'}
    title_words={w for w in re.findall(r'[a-zçğıöşü0-9]+',norm(title)) if len(w)>=4 and w not in stop}

    # Ana olay cümlesi: yalnız ilk 7 cümle içinde aranır.
    # Böylece sayfanın sonundan/başka haberlerden içerik çekilmez.
    pool=sentences[:7]
    scores=[_v93_content_score(s,title_words) for s in pool]
    anchor=max(range(len(pool)),key=lambda i:(scores[i],-i)) if pool else 0

    # Başlangıç cümlesi başlıkla anlamlı örtüşüyorsa girişten başla.
    first_overlap=len(set(re.findall(r'[a-zçğıöşü0-9]+',norm(pool[0]))) & title_words) if pool else 0
    if first_overlap>=1:
        anchor=0

    chosen=[pool[anchor]]
    # Ana olayın hemen devamındaki en fazla iki cümleyi kullan.
    # Uzak cümle avlama kesinlikle yapılmaz.
    for s in sentences[anchor+1:anchor+3]:
        # Çok bariz biçimde yeni bir haber/konu başlıyorsa kes.
        ov=len(set(re.findall(r'[a-zçğıöşü0-9]+',norm(s))) & title_words)
        if ov==0 and len(chosen)>=2 and _v93_content_score(s,title_words)<5:
            break
        chosen.append(s)

    formal=[]
    for s in chosen:
        fs=_v92_formal_sentence(s).strip()
        if not fs:
            continue
        if fs[-1] not in '.!?':
            fs+='.'
        formal.append(fs)

    # 4 satır hedefi: yaklaşık 520 karakter; ASLA kelime/cümle ortasından kesme.
    result=[]
    for s in formal:
        candidate=' '.join(result+[s])
        if result and len(candidate)>520:
            break
        result.append(s)

    out=' '.join(result).strip()

    # Çok kısa ise yalnız bir sonraki ARDIŞIK cümleyi eklemeyi dene.
    if len(out)<170:
        next_i=anchor+len(result)
        if next_i<len(sentences):
            fs=_v92_formal_sentence(sentences[next_i]).strip()
            if fs and fs[-1] not in '.!?': fs+='.'
            if fs and len(out+' '+fs)<=520:
                out=(out+' '+fs).strip()

    # Başlık veya link asla eklenmez.
    out=re.sub(r'https?://\S+','',out)
    out=re.sub(r'\s+',' ',out).strip()
    return _v87_safe_tr(out)


def _v94_formal_summary_text(text):
    """
    Very simple and robust:
    - clean the saved news summary
    - take complete sentences in their original order
    - keep up to ~4 Word lines
    - formalize common journalistic endings
    - never return blank if usable text exists
    """
    t=_v87_safe_tr(text or '')
    t=re.sub(r'https?://\S+',' ',t)
    t=re.sub(r'\s+',' ',t).strip()
    if not t:
        return ''

    # Split only on real sentence endings; preserve original order.
    parts=[x.strip() for x in re.split(r'(?<=[.!?])\s+',t) if x.strip()]
    if not parts:
        parts=[t]

    noise=('devamını oku','benzer haber','ilgili haber','çerez','cookie',
           'reklam','instagram','facebook','twitter','whatsapp','google news')
    clean=[]
    for s in parts:
        ns=norm(s)
        if any(x in ns for x in noise):
            continue
        s=_v92_formal_sentence(s).strip()
        if not s:
            continue
        if s[-1] not in '.!?':
            s+='.'
        clean.append(s)

    if not clean:
        s=_v92_formal_sentence(t).strip()
        return (s.rstrip(' .;')+'.') if s else ''

    # Approx. four lines. Do not cut a sentence.
    chosen=[]
    for s in clean:
        candidate=' '.join(chosen+[s])
        if chosen and len(candidate)>560:
            break
        chosen.append(s)
        if len(chosen)>=3:
            break

    out=' '.join(chosen).strip()
    return out


def _v95_ogn_from_existing_engines(title, body):
    """
    Yeni bir özetleme algoritması yok.
    AKT'de başarılı çalışan:
      _akt_clean_sentences -> _akt_formal_summary
    ve Bilgi Notunda kullanılan kurumsal dil normalizasyonu kullanılır.
    """
    title=_clean_note_text(title)
    body=_clean_note_text(body or title)

    # AKT motoru haberi baştan sona değerlendirir; burada yalnız çıktı uzunluğu kısaltılır.
    text=_akt_formal_summary(
        title,
        body,
        max_sentences=3,
        max_chars=560
    )
    text=_clean_note_text(text)

    # AKT özetindeki ";" akışını ÖGN için tam cümlelere dönüştür.
    clauses=[_clean_note_text(x).strip(' ,;:.') for x in re.split(r'\s*;\s*',text) if _clean_note_text(x)]
    sentences=[]
    for clause in clauses[:3]:
        formal=_v66_formalize_sentence_endings(clause).strip()
        if not formal:
            continue
        if formal[-1] not in '.!?':
            formal+='.'
        sentences.append(formal)

    if not sentences:
        formal=_v66_formalize_sentence_endings(text).strip()
        if formal and formal[-1] not in '.!?':
            formal+='.'
        return formal

    # Yaklaşık 4 satır; tam cümleyi kesme.
    chosen=[]
    for s in sentences:
        candidate=' '.join(chosen+[s])
        if chosen and len(candidate)>560:
            break
        chosen.append(s)

    return ' '.join(chosen).strip()



def _v96_unique_sentences(title, body):
    """Bilgi notu motorundaki temizleme mantığını tek haber için uygular."""
    cleaned=_akt_clean_sentences(
        _clean_note_text(title),
        _clean_note_text(body)
    )
    uniq=[]
    seen=[]
    for sent in cleaned:
        sent=_repair_mojibake_utf8(_clean_note_text(sent)).strip()
        if not sent:
            continue
        toks=set(norm(sent).split())
        if not toks:
            continue
        dup=False
        for old in seen[-30:]:
            union=len(toks|old)
            if union and len(toks&old)/union>=0.78:
                dup=True
                break
        if not dup:
            uniq.append(sent)
            seen.append(toks)
    return uniq

def _v96_has_critical_data(s):
    n=norm(s)
    return bool(
        re.search(r'\b\d+(?:[.,]\d+)?\b|%|yüzde|milyon|milyar|trilyon|bin\b',s,re.I)
        or any(x in n for x in [
            'üretim','ihracat','ithalat','kapasite','yatırım','ciro','satış',
            'başvuru','öğrenci','personel','istihdam','menzil','adet','oran',
            'endeks','bütçe','hibe','destek','maliyet','gelir','zarar'
        ])
    )

def _v96_short_information_note(title, body):
    """
    ÖGN = kısaltılmış bilgi notu.
    En fazla 4 paragraf:
      1) ana gelişme/özet,
      2) kritik veri-rakam-istatistik,
      3) gerekiyorsa tamamlayıcı gelişme/sonuç,
      4) yalnız kaynakta anlamlı bir sonuç/son durum varsa.
    """
    uniq=_v96_unique_sentences(title,body)
    if not uniq:
        fallback=_repair_mojibake_utf8(_clean_note_text(body or title))
        if not fallback:
            return []
        return [_v66_formalize_sentence_endings(fallback)]

    # 1. paragraf: bilgi notundaki gibi ilk 1-2 cümlede olayın özü.
    intro_s=uniq[:2]
    intro=_join_sentences_naturally(intro_s)
    intro=_v66_formalize_sentence_endings(intro)

    # Kritik rakam/veri cümlelerini ASLA sırf kısa özet uğruna atlama.
    critical=[]
    for i,s in enumerate(uniq):
        if i<2:
            continue
        if _v96_has_critical_data(s):
            critical.append((i,_sent_score(s),s))

    # En yüksek bilgi yoğunluklu kritik cümleleri seç, fakat haber sırasını koru.
    chosen_detail_idx=set()
    for i,score,s in sorted(critical,key=lambda x:(x[1],-x[0]),reverse=True)[:6]:
        chosen_detail_idx.add(i)

    # İlk iki cümleden sonra konu akışını tamamlayan yüksek skorlu normal cümleler.
    remaining=[
        (i,_sent_score(s),s) for i,s in enumerate(uniq[2:],start=2)
        if i not in chosen_detail_idx
    ]
    for i,score,s in sorted(remaining,key=lambda x:(x[1],-x[0]),reverse=True)[:3]:
        if score>=3:
            chosen_detail_idx.add(i)

    ordered_details=[uniq[i] for i in sorted(chosen_detail_idx)]

    paragraphs=[intro] if intro else []

    # 2-3. paragraflar: kritik detayları bilgi notu gibi gruplandır.
    if ordered_details:
        if len(ordered_details)<=4:
            p2=_join_sentences_naturally(ordered_details)
            if p2:
                paragraphs.append(_v66_formalize_sentence_endings(p2))
        else:
            split=max(2,min(4,(len(ordered_details)+1)//2))
            p2=_join_sentences_naturally(ordered_details[:split])
            p3=_join_sentences_naturally(ordered_details[split:])
            if p2:
                paragraphs.append(_v66_formalize_sentence_endings(p2))
            if p3:
                paragraphs.append(_v66_formalize_sentence_endings(p3))

    # Son paragraf: generic değerlendirme yazma; kaynakta kalan gerçek son durumdan seç.
    used=set(intro_s+ordered_details)
    tail_candidates=[s for s in uniq[-5:] if s not in used]
    if tail_candidates and len(paragraphs)<4:
        # Sonuç/hedef/son durum taşıyan cümleyi tercih et.
        result_terms=[
            'hedef','beklen','plan','sonuç','bu kapsamda','bu çerçevede','devam',
            'başlayacak','tamamlanacak','uygulanacak','sağlanacak','öngör',
            'artıracak','azaltacak','katkı','etki','takvim'
        ]
        ranked=sorted(
            tail_candidates,
            key=lambda s:(
                sum(x in norm(s) for x in result_terms),
                _sent_score(s)
            ),
            reverse=True
        )
        tail=ranked[0] if ranked else ''
        if tail:
            paragraphs.append(
                _v66_formalize_sentence_endings(_join_sentences_naturally([tail]))
            )

    # Maksimum 4 paragraf; boşları temizle.
    out=[]
    for para in paragraphs[:4]:
        para=_repair_mojibake_utf8(_clean_note_text(para)).strip()
        if para and para not in out:
            out.append(para)
    return out



def _v97_compact_sentence(s,max_chars=260):
    """Çok uzun tek cümleyi, ana özne/eylem ve kritik sayısal parçaları koruyarak kısaltır."""
    s=_repair_mojibake_utf8(_clean_note_text(s)).strip()
    if len(s)<=max_chars:
        return s

    # Virgül/noktalı virgül ile ayrılmış anlamlı parçaları değerlendir.
    parts=[x.strip(' ,;:.') for x in re.split(r'\s*[;,]\s*',s) if x.strip()]
    if not parts:
        return s[:max_chars].rsplit(' ',1)[0].rstrip(' ,;:.')+'.'

    chosen=[parts[0]]
    # Kritik veri/rakam içeren parçaları öncelikle koru.
    critical=[x for x in parts[1:] if _v96_has_critical_data(x)]
    for x in critical:
        cand=', '.join(chosen+[x])
        if len(cand)<=max_chars:
            chosen.append(x)

    # Hâlâ çok kısa ise ikinci parçayı bağlam için ekle.
    if len(chosen)==1 and len(parts)>1:
        cand=', '.join(chosen+[parts[1]])
        if len(cand)<=max_chars:
            chosen.append(parts[1])

    out=', '.join(chosen).strip()
    if out and out[-1] not in '.!?':
        out+='.'
    return out

def _v97_analyst_summary(title,body):
    """
    Gerçek analist mantığıyla kısa ÖGN:
    - tek paragraf,
    - yaklaşık 4 Word satırı,
    - 2-3 tam cümle,
    - ana gelişme + kritik rakam/veri + sonuç/son durum,
    - bilgi notunun kısaltılmış hali.
    """
    uniq=_v96_unique_sentences(title,body)
    if not uniq:
        fallback=_repair_mojibake_utf8(_clean_note_text(body or title))
        fallback=_v66_formalize_sentence_endings(fallback)
        return _v97_compact_sentence(fallback,480)

    # 1) Ana gelişme: haberin ilk anlamlı cümlesi.
    intro=_v66_formalize_sentence_endings(uniq[0]).strip()
    intro=_v97_compact_sentence(intro,260)

    chosen=[intro] if intro else []
    used={0}

    # 2) En kritik veri/rakam/istatistik:
    critical=[]
    for i,s in enumerate(uniq[1:],start=1):
        if _v96_has_critical_data(s):
            # Rakam + kurumsal/sonuç bilgisi daha yüksek puan.
            score=_sent_score(s)
            score+=min(len(re.findall(r'\d+(?:[.,]\d+)?',s)),4)*2
            critical.append((score,i,s))
    if critical:
        _,idx,s=max(critical,key=lambda x:(x[0],-x[1]))
        fs=_v66_formalize_sentence_endings(s).strip()
        fs=_v97_compact_sentence(fs,250)
        if fs and title_key(fs)!=title_key(intro):
            chosen.append(fs); used.add(idx)

    # 3) Sonuç/hedef/sonraki adım taşıyan cümle.
    result_terms=[
        'hedef','beklen','plan','sonuç','başlayacak','tamamlanacak','uygulanacak',
        'sağlanacak','öngör','katkı','etki','devam','rekor','artış','azalış',
        'başvuru','tarih','takvim','yüksel','gerile'
    ]
    result_candidates=[]
    for i,s in enumerate(uniq[1:],start=1):
        if i in used:
            continue
        n=norm(s)
        score=sum(x in n for x in result_terms)*3 + _sent_score(s)
        if score>0:
            result_candidates.append((score,i,s))
    if result_candidates:
        _,idx,s=max(result_candidates,key=lambda x:(x[0],-x[1]))
        fs=_v66_formalize_sentence_endings(s).strip()
        fs=_v97_compact_sentence(fs,220)
        if fs and all(title_key(fs)!=title_key(x) for x in chosen):
            chosen.append(fs)

    # Haber sırasını korumak için seçilen cümleleri tekrar orijinal sıraya koymuyoruz:
    # intro daima ilk, ardından kritik veri, ardından sonuç.
    # Bu, bilgi notunun kısa analist akışıdır.

    # Tek paragraf ve yaklaşık 4 satır: tam cümleyi kesmeden 500 karakter.
    final=[]
    for s in chosen[:3]:
        s=_repair_mojibake_utf8(_clean_note_text(s)).strip()
        if not s:
            continue
        if s[-1] not in '.!?':
            s+='.'
        cand=' '.join(final+[s])
        if final and len(cand)>500:
            break
        final.append(s)

    out=' '.join(final).strip()

    # İlk cümle tek başına çok uzunsa güvenli biçimde sıkıştır.
    if len(out)>500:
        out=_v97_compact_sentence(out,490)

    return _repair_mojibake_utf8(_clean_note_text(out)).strip()



def _v98_strip_site_name(text, source=''):
    """Özetin sonunda kalan yayıncı/site adını temizler."""
    t=_repair_mojibake_utf8(_clean_note_text(text)).strip()
    s=_repair_mojibake_utf8(_clean_note_text(source)).strip()
    if s:
        # Cümle sonunda " - siteadı", " — siteadı", "(siteadı)" vb.
        t=re.sub(r'\s*[-–—|]\s*'+re.escape(s)+r'\s*$', '', t, flags=re.I)
        t=re.sub(r'\s*\(\s*'+re.escape(s)+r'\s*\)\s*$', '', t, flags=re.I)
        t=re.sub(r'\s*'+re.escape(s)+r'\s*$', '', t, flags=re.I)
    # Genel domain/site sonları.
    t=re.sub(r'\s*[-–—|]\s*[\w.-]+\.(?:com|com\.tr|net|org|gov\.tr|edu\.tr)\s*$', '', t, flags=re.I)
    return t.strip(' -–—|')

def _v98_remove_heading_like_text(text, title=''):
    """
    Başlık benzeri ALL CAPS parçalarını ve haber başlığının aynısını özetten çıkarır.
    """
    t=_repair_mojibake_utf8(_clean_note_text(text))
    title_n=title_key(title)

    parts=_sentence_chunks(t)
    out=[]
    for s in parts:
        s=_repair_mojibake_utf8(_clean_note_text(s)).strip()
        if not s:
            continue

        # Haber başlığının aynısı veya çok yakınsa alma.
        if title_n and title_key(s)==title_n:
            continue

        # Kısa ALL CAPS başlıkları alma.
        letters=''.join(ch for ch in s if ch.isalpha())
        if letters and len(s)<160:
            ratio=sum(ch.isupper() for ch in letters)/max(1,len(letters))
            if ratio>0.72:
                continue

        out.append(s)
    return ' '.join(out).strip()

def _v98_safe_tr(text):
    """
    Word'e girmeden önce Türkçe karakterleri güvenli hale getirir.
    Mojibake ve görünmez karakterleri temizler.
    """
    t=_repair_mojibake_utf8(_clean_note_text(text))
    # Kalan yaygın bozukluklar.
    fixes={
        'TÃ¼rkiye':'Türkiye','TÃ¼rk':'Türk','genÃ§':'genç','dÃ¼nya':'dünya',
        'Ã¼lke':'ülke','Ã¼stÃ¼n':'üstün','Ã¶ÄŸrenci':'öğrenci','Ã¶Ärenci':'öğrenci',
        'baÅŸar':'başar','katÄ±lÄ±m':'katılım','mÃ¼cadele':'mücadele',
        'Ä±':'ı','ÄŸ':'ğ','ÅŸ':'ş','Ã§':'ç','Ã¶':'ö','Ã¼':'ü',
        'Ä°':'İ','Äž':'Ğ','Åž':'Ş','Ã‡':'Ç','Ã–':'Ö','Ãœ':'Ü',
        'Â':'','â€™':'’','â€œ':'“','â€':'”','â€“':'–','â€”':'—'
    }
    for a,b in fixes.items():
        t=t.replace(a,b)

    # Word/XML açısından problemli görünmez karakterleri temizle.
    for bad in ('\u00ad','\u200b','\u200c','\u200d','\ufeff'):
        t=t.replace(bad,'')

    # Kalan açık mojibake sembollerini boşlukla değiştir.
    t=re.sub(r'[ÃÄÅÂ ]+',' ',t)
    t=re.sub(r'\s+',' ',t).strip()
    return t

def _v98_exact_four_line_summary(title, source, body):
    """
    V98 ÖGN:
    - her haber tek paragraf,
    - yaklaşık 4 Word satırı,
    - 2-3 tam cümle,
    - resmî dil,
    - kritik rakam/veri korunur,
    - başlık/site adı/bozuk karakter yok.
    """
    # V97 motorunu temel al.
    raw=_v97_analyst_summary(title,body)
    raw=_v98_safe_tr(raw)
    raw=_v98_remove_heading_like_text(raw,title)
    raw=_v98_strip_site_name(raw,source)

    if not raw:
        # Yedek: kayıtlı haber metninden doğrudan kısa resmî özet.
        raw=_v95_ogn_from_existing_engines(title,body)
        raw=_v98_safe_tr(raw)
        raw=_v98_remove_heading_like_text(raw,title)
        raw=_v98_strip_site_name(raw,source)

    # Cümleleri yeniden düzenle, tam cümleler dışında kesme yok.
    sents=[]
    for s in _sentence_chunks(raw):
        s=_v98_safe_tr(s).strip()
        if not s:
            continue
        s=_v66_formalize_sentence_endings(s).strip()
        if s and s[-1] not in '.!?':
            s+='.'
        sents.append(s)

    # Hedef yaklaşık 4 satır = 430-500 karakter bandı.
    chosen=[]
    total=0
    for s in sents[:4]:
        candidate=' '.join(chosen+[s])
        if chosen and len(candidate)>500:
            break
        chosen.append(s)
        total=len(candidate)
        if total>=430:
            break

    # Çok kısa kaldıysa sonraki tam cümleyi ekle.
    if len(' '.join(chosen))<300:
        for s in sents[len(chosen):]:
            candidate=' '.join(chosen+[s])
            if len(candidate)<=500:
                chosen.append(s)
            if len(' '.join(chosen))>=360:
                break

    out=' '.join(chosen).strip()

    # Son güvenlik katmanı.
    out=_v98_safe_tr(out)
    out=_v98_remove_heading_like_text(out,title)
    out=_v98_strip_site_name(out,source)

    # Sonunda nokta olsun, ama site adı olmasın.
    if out and out[-1] not in '.!?':
        out+='.'
    return out



# ========================= V99: ÖGN ANALİST FİLTRESİ =========================
_V99_UI_NOISE = [
    r'sıralamayı değiştirmek için kartları',
    r'kartları yukarı',
    r'kartları aşağı',
    r'view more',
    r'daha fazla göster',
    r'devamını oku',
    r'cookie',
    r'çerez',
    r'reklam',
    r'abone ol',
    r'bildirimleri aç',
    r'ana sayfa',
    r'son dakika',
    r'galeri',
    r'foto galeri',
    r'video galeri',
    r'yorumlar',
    r'paylaş',
]

_V99_NEWS_SPEECH = [
    (r'\bbaşladı\b', 'başlamıştır'),
    (r'\bbaşladıktan\b', 'başladıktan'),
    (r'\baçıkladı\b', 'açıklamıştır'),
    (r'\bbelirtti\b', 'belirtmiştir'),
    (r'\bkaydetti\b', 'kaydetmiştir'),
    (r'\bduyurdu\b', 'duyurmuştur'),
    (r'\bbildirdi\b', 'bildirmiştir'),
    (r'\bifade etti\b', 'ifade etmiştir'),
    (r'\bsöyledi\b', 'belirtmiştir'),
    (r'\bgerçekleşti\b', 'gerçekleşmiştir'),
    (r'\byükseldi\b', 'yükselmiştir'),
    (r'\bgeriledi\b', 'gerilemiştir'),
    (r'\barttı\b', 'artmıştır'),
    (r'\bazaldı\b', 'azalmıştır'),
    (r'\bkırıldı\b', 'kırılmıştır'),
    (r'\btamamlandı\b', 'tamamlanmıştır'),
    (r'\bgirdi\b', 'girmiştir'),
    (r'\bbaşlıyor\b', 'başlamaktadır'),
    (r'\bdevam ediyor\b', 'devam etmektedir'),
    (r'\bhedefliyor\b', 'hedeflemektedir'),
    (r'\bbekleniyor\b', 'beklenmektedir'),
    (r'\bsürüyor\b', 'sürmektedir'),
    (r'\bulaştı\b', 'ulaşmıştır'),
    (r'\bçıktı\b', 'çıkmıştır'),
    (r'\byapıldı\b', 'yapılmıştır'),
    (r'\bseçilecek\b', 'seçilecektir'),
    (r'\byetiştireceğiz\b', 'yetiştirilmesi planlanmaktadır'),
    (r'\bkazandıracağız\b', 'kazandırılması hedeflenmektedir'),
]

def _v99_clean_article_text(text, source=''):
    t=_v98_safe_tr(text)
    t=t.replace('»',' ').replace('›',' ').replace('',"'").replace('','ş').replace('¢','â')
    t=re.sub(r'\s+',' ',t).strip()

    # Web arayüzü / sayfa gürültüsü taşıyan cümleleri tamamen at.
    kept=[]
    for s in _sentence_chunks(t):
        n=norm(s)
        if any(re.search(p,n,re.I) for p in _V99_UI_NOISE):
            continue
        # Kaynak/site navigasyonu gibi çok kısa satırları at.
        if len(s.split()) < 4 and not re.search(r'\d',s):
            continue
        kept.append(s.strip())
    t=' '.join(kept)

    # Kaynak adını cümle başından/sonundan temizle.
    if source:
        ss=re.escape(_v98_safe_tr(source))
        t=re.sub(r'^\s*'+ss+r'\s*[-:»|]*\s*','',t,flags=re.I)
        t=re.sub(r'\s*[-:»|]*\s*'+ss+r'\s*$','',t,flags=re.I)

    # Genel domainleri ve tipik gazete navigasyonunu kaldır.
    t=re.sub(r'\b[\w.-]+\.(?:com\.tr|com|net|org|gov\.tr|edu\.tr)\b',' ',t,flags=re.I)
    t=re.sub(r'^[^.!?]{0,80}\b(?:Gazetesi|Gazete|Haber|Haberleri)\s*[»|:-]+\s*','',t,flags=re.I)
    t=re.sub(r'\s+',' ',t).strip()
    return t

def _v99_is_title_only(title, text):
    a=set(norm(title).split())
    b=set(norm(text).split())
    if not a or not b:
        return False
    return len(a & b)/max(1,len(a | b)) >= 0.72 and len(text) < 220

def _v99_officialize(text):
    t=_v98_safe_tr(text).strip()
    for pat,repl in _V99_NEWS_SPEECH:
        t=re.sub(pat,repl,t,flags=re.I)
    # Haber dilindeki doğrudan alıntı kalıplarını kurumsallaştır.
    t=re.sub(r'\bifadelerini kullandı\b','belirtmiştir',t,flags=re.I)
    t=re.sub(r'\bşunları kaydetti\b','açıklamada bulunmuştur',t,flags=re.I)
    t=re.sub(r'\bşöyle\b\s*:?', '', t, flags=re.I)
    t=_v66_formalize_sentence_endings(t)
    t=re.sub(r'\s+',' ',t).strip()
    return t

def _v99_sentence_value(s):
    n=norm(s)
    score=_sent_score(s)
    nums=len(re.findall(r'\b\d+(?:[.,]\d+)?\b|%',s))
    score += min(nums,5)*3
    # Kurumsal/stratejik bilgi yoğunluğu
    for term in [
        'tüik','tcmb','epdk','bakanlık','cumhurbaşkanı','tübitak','kosgeb','tse','türkpatent',
        'üretim','ihracat','ithalat','yatırım','kapasite','oran','yüzde','milyon','milyar',
        'teslimat','envanter','program','proje','hibe','destek','menzil','adet','öğrenci',
        'araştırmacı','tesis','fabrika','teknoloji','savunma','uzay','yapay zeka'
    ]:
        if term in n:
            score += 2
    return score

def _v99_analyst_ogn(title, source, body, fallback_summary=''):
    """
    İşyeri ÖGN örneğine göre:
    somut özne/kurum + gelişme + kritik rakam/ölçek + sonuç.
    Tek paragraf, 2-3 tam cümle, yaklaşık dört Word satırı.
    """
    title=_v98_safe_tr(title)
    source=_v98_safe_tr(source)
    body=_v99_clean_article_text(body,source)
    fallback_summary=_v99_clean_article_text(fallback_summary,source)

    # Gerçek içerik yoksa başlığı Word'e basmak yerine kayıtlı özeti dene.
    candidate_body=body
    if not candidate_body or _v99_is_title_only(title,candidate_body):
        candidate_body=fallback_summary

    # Hâlâ yalnız başlıksa bunu geçerli ÖGN sayma.
    if not candidate_body or _v99_is_title_only(title,candidate_body):
        return ''

    sentences=_v96_unique_sentences(title,candidate_body)
    clean=[]
    for s in sentences:
        s=_v99_clean_article_text(s,source)
        if not s or _v99_is_title_only(title,s):
            continue
        if any(re.search(p,norm(s),re.I) for p in _V99_UI_NOISE):
            continue
        clean.append(s)

    if not clean:
        return ''

    # İlk cümle: haberin gerçek ana gelişmesi. İlk 4 cümlede en yüksek değerli olanı seç.
    first_pool=list(enumerate(clean[:4]))
    first_idx, first=max(first_pool,key=lambda x:(_v99_sentence_value(x[1]),-x[0]))

    selected=[(first_idx,first)]

    # Kritik rakam/veri: ana cümlede yoksa ayrıca seç.
    critical=[]
    for i,s in enumerate(clean):
        if i==first_idx:
            continue
        if _v96_has_critical_data(s):
            critical.append((_v99_sentence_value(s),i,s))
    if critical:
        _,i,s=max(critical,key=lambda x:(x[0],-x[1]))
        selected.append((i,s))

    # Sonuç/hedef/son durum: yalnız kaynakta varsa.
    result_terms=['hedef','plan','beklen','öngör','teslim','envanter','başvuru','hibe',
                  'artış','azalış','rekor','üretim','faaliyete','tamamlan','başlam']
    results=[]
    used={i for i,_ in selected}
    for i,s in enumerate(clean):
        if i in used:
            continue
        sc=sum(x in norm(s) for x in result_terms)*4 + _v99_sentence_value(s)
        if sc>3:
            results.append((sc,i,s))
    if results:
        _,i,s=max(results,key=lambda x:(x[0],-x[1]))
        selected.append((i,s))

    # Analist akışı: ana gelişme -> veri -> sonuç. En fazla 3 cümle.
    out=[]
    for _,s in selected[:3]:
        s=_v99_officialize(s)
        s=_v98_strip_site_name(s,source)
        s=re.sub(r'^[A-ZÇĞİÖŞÜ0-9\s\-–—:]{8,80}(?=[A-ZÇĞİÖŞÜ][a-zçğıöşü])','',s).strip()
        if not s:
            continue
        if s[-1] not in '.!?':
            s+='.'
        if all(title_key(s)!=title_key(x) for x in out):
            out.append(s)

    # Dört satır hedefi: yaklaşık 560 karakter; cümle ortasında kesme yok.
    chosen=[]
    for s in out:
        cand=' '.join(chosen+[s])
        if chosen and len(cand)>560:
            break
        chosen.append(s)

    text=' '.join(chosen).strip()
    text=_v99_clean_article_text(text,source)
    text=_v99_officialize(text)
    text=_v98_strip_site_name(text,source)
    text=_v98_safe_tr(text)

    if _v99_is_title_only(title,text):
        return ''
    if text and text[-1] not in '.!?':
        text+='.'
    return text


# ========================= V100: KONU BÜTÜNLÜĞÜ =========================

def _v100_title_topic(title, source=''):
    """
    Başlıktan yalnızca konu/özne çekirdeğini çıkarır.
    Başlığı aynen basmaz; ALL CAPS ve site adını temizler.
    """
    t=_v98_safe_tr(title)
    t=_v98_strip_site_name(t,source)
    t=re.sub(r'\s*[-–—|]\s*[\w.-]+\.(?:com|com\.tr|net|org|gov\.tr|edu\.tr)\s*$','',t,flags=re.I)
    t=re.sub(r'!+$','',t).strip()

    # Tamamı büyük harfliyse normal cümle görünümüne getir.
    letters=''.join(ch for ch in t if ch.isalpha())
    if letters and sum(ch.isupper() for ch in letters)/max(1,len(letters)) > 0.72:
        t=t.lower()
        t=t[:1].upper()+t[1:]

    return t.strip(' -–—|:;')

def _v100_subject_hint(title, source=''):
    """
    Başlıktan self-contained paragraf için özne/konu ipucu çıkarır.
    Örn: 'KAAN...' -> 'KAAN projesi'
         'Antalya'da Savunma Sanayine Yatırım Fırsatı' -> 'Antalya'daki savunma sanayii yatırımları'
    """
    t=_v100_title_topic(title,source)
    n=norm(t)

    # Bilinen kalıplar.
    mappings=[
        (r'\bkaan\b', 'KAAN projesi'),
        (r'\bbayraktar kalkan\b', 'Bayraktar Kalkan DİHA'),
        (r'\bt10x\b|\bt10f\b|\btogg\b', 'Togg’un T10X ve T10F modelleri'),
        (r'\belektrikli araç.*şarj\b|\bşarj altyap', 'Türkiye’de elektrikli araç şarj altyapısı'),
        (r'\bkapasite kullanım\b', 'imalat sanayisi kapasite kullanım oranı'),
        (r'\bkardemir\b', 'Kardemir Çelik’in 2026 yılı ilk yarı finansal sonuçları'),
        (r'\bgezeravcı\b', 'Alper Gezeravcı’nın Amasya’daki görevi'),
        (r'\bgoogle.*ai plus\b|\bai plus\b', 'Google’ın üniversite öğrencilerine yönelik AI Plus programı'),
        (r'\bmoğolistan.*mühimmat\b|\bmke.*moğolistan\b', 'MKE’nin Moğolistan’daki mühimmat üretim tesisi'),
        (r'\btekno.*mavi vatan\b|\balkü\b|\bzeronetech\b', 'TEKNOFEST Mavi Vatan kapsamında ALKÜ Zeronetech Takımı'),
        (r'\bsavunma sanay.*yatırım.*antalya\b|\bantalya.*savunma sanay', 'Antalya’daki savunma sanayii yatırım fırsatları'),
        (r'\byapay zeka olimpiyat\b|\bbilim olimpiyat', 'TÜBİTAK Bilim Olimpiyatları kapsamında Türkiye’yi temsil eden öğrenciler'),
        (r'\btürk telekom\b|\bdijitalde hayat kolay\b', 'Türk Telekom’un Dijitalde Hayat Kolay projesi'),
        (r'\b5g.*robotik cerrahi\b|\btcg anadolu.*5g\b', 'TCG ANADOLU’da 5G destekli uzaktan robotik cerrahi uygulaması'),
    ]
    for pat,label in mappings:
        if re.search(pat,n,re.I):
            return label

    # Genel başlık: ilk 8-10 anlamlı kelimeyi konu ipucu olarak kullan.
    words=t.split()
    if not words:
        return ''
    return ' '.join(words[:10]).strip(' ,;:-')

def _v100_is_fragment(s):
    """Özne/bağlam içermeyen kırık veya yarım cümleleri tespit eder."""
    s=_v98_safe_tr(s).strip()
    if not s:
        return True
    n=norm(s)

    # Açık yarım cümleler / UI artıkları.
    if s.endswith((' ve',' ile',' için',' k',';',' :','...','…')):
        return True
    if re.search(r'\b(?:proje|program|şirket|kurum|bu kapsamda|bunun yanında|ayrıca)\b',n) and len(s.split())<7:
        return True
    if re.match(r'^(proje|program|şirket|kurum|bunun yanında|bu kapsamda|ayrıca)\b',n):
        return True
    if re.match(r'^\d+\s+(?:yıl|ay|gün)\b',n):
        return True
    return False

def _v100_contextualize(sentence, subject):
    """
    Paragrafın ilk cümlesi kendi başına anlaşılmıyorsa başlıktan gelen konu çekirdeğiyle
    bağlamı tamamlar. Başlığı olduğu gibi eklemez.
    """
    s=_v99_officialize(sentence).strip()
    if not s:
        return ''
    n=norm(s)

    weak_starts=(
        'proje ','program ','şirket ','kurum ','bunun yanında ',
        'bu kapsamda ','ayrıca ','ilk olarak ','daha sonra '
    )
    if any(n.startswith(x) for x in weak_starts) or _v100_is_fragment(s):
        if subject:
            # "Proje, Mart..." -> "KAAN projesinde üretim faaliyetleri Mart..."
            if n.startswith('proje '):
                s=re.sub(r'^\s*Proje\s*,?\s*', subject+' kapsamında ', s, flags=re.I)
            elif n.startswith('program '):
                s=re.sub(r'^\s*Program\s*,?\s*', subject+' kapsamında ', s, flags=re.I)
            elif n.startswith('şirket '):
                s=re.sub(r'^\s*Şirket\s*,?\s*', subject+' kapsamında ilgili şirket ', s, flags=re.I)
            else:
                s=subject.rstrip(' .')+' kapsamında '+s[:1].lower()+s[1:]
    return s

def _v100_pick_summary_sentences(title, source, body, fallback):
    """
    Self-contained analist özeti:
    1) ilk cümle mutlaka konuyu/özneyi açıklar,
    2) ikinci cümle kritik veri/rakam,
    3) üçüncü cümle sonuç/hedef/takvim.
    """
    subject=_v100_subject_hint(title,source)
    body=_v99_clean_article_text(body,source)
    fallback=_v99_clean_article_text(fallback,source)

    candidate=body
    if not candidate or _v99_is_title_only(title,candidate):
        candidate=fallback

    sents=_v96_unique_sentences(title,candidate)
    clean=[]
    for s in sents:
        s=_v99_clean_article_text(s,source)
        if not s or _v100_is_fragment(s):
            continue
        if _v99_is_title_only(title,s):
            continue
        clean.append(s)

    if not clean:
        return ''

    # 1) Ana olay: ilk 5 cümle içinde subject/title overlap + kurumsal eylem.
    tw=set(re.findall(r'[a-zçğıöşü0-9]+',norm(subject or title)))
    def overlap(s):
        sw=set(re.findall(r'[a-zçğıöşü0-9]+',norm(s)))
        return len(sw & tw)

    action_terms=['açıkla','duyur','başlat','gerçekleştir','tamamla','imzala','üret',
                  'satış','yatırım','test','görev','teslim','envanter','faaliyete','seç']
    def intro_score(s,idx):
        n=norm(s)
        return 7*overlap(s)+4*sum(x in n for x in action_terms)+_sent_score(s)-idx

    first_pool=list(enumerate(clean[:5]))
    idx0, intro=max(first_pool,key=lambda z:intro_score(z[1],z[0]))
    intro=_v100_contextualize(intro,subject)

    # Eğer ilk cümlede subject hiç yoksa, doğal konu cümlesiyle bağla.
    if subject and overlap(intro)==0:
        # Başlıktan birebir kopya değil, konu bağlamı.
        if intro:
            intro=subject.rstrip(' .')+' hakkında, '+intro[:1].lower()+intro[1:]

    selected=[intro] if intro else []
    used={idx0}

    # 2) Kritik veri/rakam.
    critical=[]
    for i,s in enumerate(clean):
        if i in used:
            continue
        if _v96_has_critical_data(s):
            critical.append((_v99_sentence_value(s)+3*overlap(s),i,s))
    if critical:
        _,i,s=max(critical,key=lambda x:(x[0],-x[1]))
        fs=_v99_officialize(s)
        if fs and not _v100_is_fragment(fs):
            selected.append(fs); used.add(i)

    # 3) Sonuç/hedef/takvim.
    result_terms=['hedef','beklen','plan','başvuru','teslim','envanter','rekor',
                  'katkı','artış','azalış','faaliyete','tamamlan','başlam','uygulan']
    results=[]
    for i,s in enumerate(clean):
        if i in used:
            continue
        n=norm(s)
        sc=4*sum(x in n for x in result_terms)+_sent_score(s)+2*overlap(s)
        if sc>3:
            results.append((sc,i,s))
    if results:
        _,i,s=max(results,key=lambda x:(x[0],-x[1]))
        fs=_v99_officialize(s)
        if fs and not _v100_is_fragment(fs):
            selected.append(fs)

    # Tekrarlı sayısal cümleleri azalt.
    final=[]
    seen_nums=[]
    for s in selected[:3]:
        s=_v98_safe_tr(_v99_officialize(s)).strip()
        if not s:
            continue
        nums=set(re.findall(r'\d+(?:[.,]\d+)?',s))
        duplicate=False
        for prev,prev_nums in zip(final,seen_nums):
            if nums and nums==prev_nums and len(nums)>=1 and title_key(s)==title_key(prev):
                duplicate=True
                break
        if duplicate:
            continue
        if s[-1] not in '.!?':
            s+='.'
        final.append(s)
        seen_nums.append(nums)

    # 4 satır hedefi ~560 karakter; tam cümle kesilmez.
    out=[]
    for s in final:
        cand=' '.join(out+[s])
        if out and len(cand)>560:
            break
        out.append(s)

    text=' '.join(out).strip()
    text=_v99_clean_article_text(text,source)
    text=_v99_officialize(text)
    text=_v98_strip_site_name(text,source)
    text=_v98_safe_tr(text)

    # Son güvenlik: paragraf yine başlık seviyesindeyse geçersiz say.
    if _v99_is_title_only(title,text):
        return ''
    return text


# ========================= V101: ÖGN KONU BÜTÜNLÜĞÜ 2.0 =========================

def _v101_clean_unicode(text):
    """Türkçe metindeki mojibake/control karakterlerini Word öncesi agresif biçimde temizler."""
    import unicodedata
    t=_repair_mojibake_utf8(_clean_note_text(text))

    # Kullanıcı çıktısında görülen tek-byte/control bozulmaları.
    fixes={
        '\x9c':'Ü','\x9e':'Ş','\x9f':'ş','\x96':'Ö','\x91':"'",'\x92':"'",'\x93':'“','\x94':'”',
        'T BİTAK':'TÜBİTAK','T BİTAK':'TÜBİTAK',
        'ö şrenci':'öğrenci','ö şrenc':'öğrenc','yarı ş':'yarış','etti şi':'ettiği',
        'ba şarı':'başarı','gümü ş':'gümüş',' zekinci':' Özekinci',
        'Yapay Zek â':'Yapay Zekâ','veyaşağı':'veya aşağı'
    }
    for a,b in fixes.items():
        t=t.replace(a,b)

    # Harflerin arasına sızmış boşluklu ş/ğ/ü bozulmaları.
    t=re.sub(r'\bö\s+şrenc', 'öğrenc', t, flags=re.I)
    t=re.sub(r'\byarı\s+ş', 'yarış', t, flags=re.I)
    t=re.sub(r'\bba\s+şar', 'başar', t, flags=re.I)
    t=re.sub(r'\bgümü\s+ş', 'gümüş', t, flags=re.I)
    t=re.sub(r'\betti\s+şi\b', 'ettiği', t, flags=re.I)

    # Kontrol/görünmez karakterleri kaldır.
    t=''.join(ch for ch in t if unicodedata.category(ch) not in {'Cc','Cf'} or ch in '\t\n\r')
    t=unicodedata.normalize('NFC',t)
    t=re.sub(r'\s+',' ',t).strip()
    return t

def _v101_title_terms(title):
    stop={'ve','ile','için','bir','bu','da','de','ile','olan','olarak','son','yeni',
          'türkiye','türk','haber','haberi','başladı','oldu','edildi','açıkladı'}
    return [w for w in re.findall(r'[a-zçğıöşü0-9]+',norm(title)) if len(w)>2 and w not in stop]

def _v101_sentence_overlap(title,s):
    tw=set(_v101_title_terms(title))
    sw=set(re.findall(r'[a-zçğıöşü0-9]+',norm(s)))
    return len(tw & sw)

def _v101_bad_sentence(s):
    s=_v101_clean_unicode(s).strip()
    n=norm(s)
    if not s or len(s)<35:
        return True
    if _v100_is_fragment(s):
        return True
    if any(re.search(p,n,re.I) for p in _V99_UI_NOISE):
        return True
    # Başlık/navigation artıkları.
    if re.search(r'\b(?:yarışıyor|tükendi|fırsatı|büyüyor)\s*!?\s*$',n) and len(s)<120:
        return True
    if re.search(r'\b(?:gazetesi|gazete|haberleri?)\s*[»|:-]',n):
        return True
    return False

def _v101_formal_sentence(s):
    """Haber dilini resmî bilgi notu diline yaklaştırır; doğrudan alıntı kırıntılarını atar."""
    s=_v101_clean_unicode(s)
    # Açık/kapanmamış tırnakları temizle.
    s=s.replace('"','').replace('“','').replace('”','')
    s=re.sub(r'\b(?:dedi|diyor|diye konuştu)\b','belirtmiştir',s,flags=re.I)
    s=_v99_officialize(s)
    # Kalan yaygın haber dili.
    repl=[
        (r'\bedecek\b','edecektir'),(r'\bolacak\b','olacaktır'),
        (r'\byarışacak\b','yarışacaktır'),(r'\bsunuluyor\b','sunulmaktadır'),
        (r'\büretiyor\b','üretmektedir'),(r'\bilerliyor\b','ilerlemektedir'),
        (r'\bdevreye alındı\b','devreye alınmıştır'),
        (r'\bonay alındı\b','onay alınmıştır'),
        (r'\bgerçekleştirildi\b','gerçekleştirilmiştir'),
        (r'\bdahil etti\b','dahil etmiştir'),
        (r'\btest etti\b','test etmiştir'),
        (r'\badım attı\b','adım atmıştır'),
    ]
    for a,b in repl:
        s=re.sub(a,b,s,flags=re.I)
    return re.sub(r'\s+',' ',s).strip()

def _v101_build_intro(title,source,sents):
    """
    İlk cümle mutlaka 'kim/ne + ne oldu' bilgisini taşır.
    Başlığın kendisini yazmaz; haber gövdesindeki en iyi bağlam cümlesini seçer.
    """
    subject=_v100_subject_hint(title,source)
    scored=[]
    for i,s in enumerate(sents[:8]):
        n=norm(s)
        action=sum(x in n for x in [
            'açıkla','duyur','başlat','başvur','gerçekleştir','üret','teslim','envanter',
            'faaliyete','satış','seç','program','proje','oran','veri','rekor','görev'
        ])
        sc=8*_v101_sentence_overlap(title,s)+4*action+_v99_sentence_value(s)-i
        scored.append((sc,i,s))
    if not scored:
        return '',-1
    _,idx,s=max(scored,key=lambda x:x[0])
    s=_v101_formal_sentence(s)

    # "Proje...", "Tüketim...", "Bunun yanında..." gibi referansı belirsiz başlangıçları konuya bağla.
    n=norm(s)
    ambiguous=re.match(r'^(proje|program|tüketim|şirket|bu kapsamda|bunun yanında|ayrıca|yoğun ilgi)',n)
    if subject and (ambiguous or _v101_sentence_overlap(title,s)==0):
        if n.startswith('proje'):
            s=re.sub(r'^\s*Proje\s*,?\s*',subject+' kapsamında ',s,flags=re.I)
        elif n.startswith('program'):
            s=re.sub(r'^\s*Program\s*,?\s*',subject+' kapsamında ',s,flags=re.I)
        elif n.startswith('tüketim'):
            s=subject.rstrip(' .')+' kapsamında '+s[:1].lower()+s[1:]
        else:
            s=subject.rstrip(' .')+' kapsamında '+s[:1].lower()+s[1:]

    return s,idx

def _v101_semantic_key(s):
    """Yakın tekrarları, özellikle aynı rakamı tekrarlayan cümleleri azaltır."""
    n=norm(s)
    nums=tuple(re.findall(r'\d+(?:[.,]\d+)?',n))
    words=[w for w in re.findall(r'[a-zçğıöşü]+',n) if len(w)>4]
    return set(words),set(nums)

def _v101_is_near_duplicate(s,chosen):
    sw,sn=_v101_semantic_key(s)
    for prev in chosen:
        pw,pn=_v101_semantic_key(prev)
        word_sim=len(sw&pw)/max(1,len(sw|pw))
        num_sim=(bool(sn) and bool(pn) and len(sn&pn)/max(1,len(sn|pn))>=0.75)
        if word_sim>=0.48 or (num_sim and word_sim>=0.25):
            return True
    return False

def _v101_analyst_paragraph(title,source,body,fallback=''):
    """
    Her haber için bağımsız okunabilen 4 satırlık mini bilgi notu:
    GİRİŞ: kim/ne, hangi gelişme
    GELİŞME: kritik veri/rakam/yer/tarih
    SONUÇ: hedef, sonuç, mevcut durum veya sonraki aşama
    """
    title=_v101_clean_unicode(title)
    source=_v101_clean_unicode(source)
    body=_v99_clean_article_text(_v101_clean_unicode(body),source)
    fallback=_v99_clean_article_text(_v101_clean_unicode(fallback),source)

    candidate=body
    if not candidate or _v99_is_title_only(title,candidate):
        candidate=fallback

    raw=_v96_unique_sentences(title,candidate)
    sents=[]
    for s in raw:
        s=_v99_clean_article_text(_v101_clean_unicode(s),source)
        if _v101_bad_sentence(s) or _v99_is_title_only(title,s):
            continue
        sents.append(s)

    # İçerik gerçekten yoksa başlığı "özet" diye basma.
    if not sents:
        return ''

    intro,intro_idx=_v101_build_intro(title,source,sents)
    if not intro:
        return ''

    chosen=[intro]
    used={intro_idx}

    # Gelişme: rakam/istatistik/ölçek/yer/tarih taşıyan en güçlü cümle.
    detail_candidates=[]
    for i,s in enumerate(sents):
        if i in used: continue
        n=norm(s)
        data=bool(re.search(r'\d|%|yüzde|milyon|milyar|bin|adet|oran|tarih',s,re.I))
        sc=_v99_sentence_value(s)+5*data+2*_v101_sentence_overlap(title,s)
        if sc>3:
            detail_candidates.append((sc,i,s))
    for _,i,s in sorted(detail_candidates,reverse=True):
        fs=_v101_formal_sentence(s)
        if fs and not _v101_is_near_duplicate(fs,chosen):
            chosen.append(fs); used.add(i); break

    # Sonuç: hedef/son durum/takvim/etki.
    result_candidates=[]
    for i,s in enumerate(sents):
        if i in used: continue
        n=norm(s)
        hits=sum(x in n for x in [
            'hedef','plan','beklen','teslim','başvuru','envanter','faaliyete',
            'tamamlan','başlam','katkı','sonuç','rekor','artış','azalış','dönem',
            'seviye','ulaş','oluştur','sağla'
        ])
        sc=5*hits+_v99_sentence_value(s)+_v101_sentence_overlap(title,s)
        if hits:
            result_candidates.append((sc,i,s))
    for _,i,s in sorted(result_candidates,reverse=True):
        fs=_v101_formal_sentence(s)
        if fs and not _v101_is_near_duplicate(fs,chosen):
            chosen.append(fs); break

    # 2-3 tam cümle; 4 Word satırı hedefi. Cümle ortasında kesme yapılmaz.
    final=[]
    for s in chosen[:3]:
        s=_v101_formal_sentence(s).strip()
        if not s: continue
        if s[-1] not in '.!?': s+='.'
        cand=' '.join(final+[s])
        if final and len(cand)>620:
            break
        final.append(s)

    text=' '.join(final)
    text=_v101_clean_unicode(_v98_strip_site_name(text,source))
    # Cümle sonunda kalan başlık/site artıkları.
    text=re.sub(r'\s+[A-ZÇĞİÖŞÜ0-9][A-ZÇĞİÖŞÜ0-9\s\'’-]{8,}\s*!?\s*(?=\.|$)','',text)
    text=re.sub(r'\s+',' ',text).strip()
    return text

def make_important_basket_docx_v101(basket_df):
    doc=Document()
    sec=doc.sections[0]
    sec.top_margin=Cm(2); sec.bottom_margin=Cm(2)
    sec.left_margin=Cm(2.5); sec.right_margin=Cm(2.5)

    normal=doc.styles['Normal']
    normal.font.name='Times New Roman'
    normal.font.size=Pt(12)
    normal._element.rPr.rFonts.set(qn('w:eastAsia'),'Times New Roman')

    now=datetime.now().astimezone()
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    p.add_run(now.strftime('%d/%m/%Y'))
    p=doc.add_paragraph()
    p.add_run('Konu: ').bold=True
    p.add_run('STB Temsilciliği Önemli Gelişmeler Notu')

    rows=[] if basket_df is None else basket_df.to_dict('records')

    def process_row(r):
        title=_v101_clean_unicode(r.get('title',''))
        source=_v101_clean_unicode(r.get('source',''))
        summary=_v101_clean_unicode(r.get('summary',''))
        url=str(r.get('url','') or '')
        news_time=_v101_clean_unicode(r.get('news_time',''))

        detail={}
        try:
            detail=article_detail({
                'Başlık':title,'Kaynak':source,'URL':url,'Yayıncı_URL':url,
                'İçerik_Özeti':summary,'Tarih':news_time
            }) or {}
        except Exception:
            pass

        text=_v101_analyst_paragraph(title,source,detail.get('text') or '',summary)
        if not text and summary:
            text=_v101_analyst_paragraph(title,source,summary,summary)
        return text

    outputs=['']*len(rows)
    if rows:
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(6,len(rows))) as ex:
            jobs={ex.submit(process_row,r):i for i,r in enumerate(rows)}
            for fut in concurrent.futures.as_completed(jobs):
                idx=jobs[fut]
                try: outputs[idx]=fut.result()
                except Exception: outputs[idx]=''

    for idx,r in enumerate(rows):
        text=_v101_clean_unicode(outputs[idx] if idx<len(outputs) else '')
        if not text:
            # İçeriksiz başlığı sahte bir 4 satırlık özet haline getirmiyoruz.
            # Kullanıcıya Word içinde başlık kalıntısı göstermek yerine bu kayıt atlanır.
            continue
        p=doc.add_paragraph()
        p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing=1.0
        p.paragraph_format.space_after=Pt(7)
        p.add_run(text.rstrip(' .;')+' (STB).')

    doc.add_paragraph('Arz olunur.')
    bio=BytesIO(); doc.save(bio); bio.seek(0)
    return bio.getvalue()

# ======================= /V101: ÖGN KONU BÜTÜNLÜĞÜ 2.0 =========================
def make_important_basket_docx_v100(basket_df):
    """
    V100: Her haber kendi içinde anlamlı bir bütün oluşturur.
    Başlıktan konu bağlamı alınır ama başlık Word'e ayrı yazılmaz.
    'Proje...' gibi bağlamsız başlangıçlar düzeltilir.
    """
    doc=Document()
    sec=doc.sections[0]
    sec.top_margin=Cm(2); sec.bottom_margin=Cm(2)
    sec.left_margin=Cm(2.5); sec.right_margin=Cm(2.5)

    normal=doc.styles['Normal']
    normal.font.name='Times New Roman'
    normal.font.size=Pt(12)
    normal._element.rPr.rFonts.set(qn('w:eastAsia'),'Times New Roman')

    now=datetime.now().astimezone()
    p=doc.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    p.add_run(now.strftime('%d/%m/%Y'))

    p=doc.add_paragraph()
    p.add_run('Konu: ').bold=True
    p.add_run('STB Temsilciliği Önemli Gelişmeler Notu')

    rows=[] if basket_df is None else basket_df.to_dict('records')

    def process_row(r):
        title=_v98_safe_tr(r.get('title',''))
        source=_v98_safe_tr(r.get('source',''))
        summary=_v98_safe_tr(r.get('summary',''))
        url=str(r.get('url','') or '')
        news_time=_v98_safe_tr(r.get('news_time',''))

        detail={}
        try:
            detail=article_detail({
                'Başlık':title,'Kaynak':source,'URL':url,'Yayıncı_URL':url,
                'İçerik_Özeti':summary,'Tarih':news_time
            }) or {}
        except Exception:
            pass

        body=detail.get('text') or ''
        text=_v100_pick_summary_sentences(title,source,body,summary)

        # Tam metin sorunluysa yalnız kayıtlı özetle tekrar dene.
        if not text and summary:
            text=_v100_pick_summary_sentences(title,source,summary,summary)

        # Son çare: başlığı "başlık" olarak değil, konu cümlesine dönüştür.
        if not text:
            subject=_v100_subject_hint(title,source)
            if subject:
                text=f'{subject} kapsamında gelişmenin ayrıntılarına ilişkin haber içeriği sınırlı olduğundan, mevcut açık kaynak metninden doğrulanabilir ek bilgi çıkarılamamıştır.'
                text=_v99_officialize(text)
        return text

    outputs=['']*len(rows)
    if rows:
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(6,len(rows))) as ex:
            jobs={ex.submit(process_row,r):i for i,r in enumerate(rows)}
            for fut in concurrent.futures.as_completed(jobs):
                idx=jobs[fut]
                try:
                    outputs[idx]=fut.result()
                except Exception:
                    outputs[idx]=''

    for idx,r in enumerate(rows):
        text=_v98_safe_tr(outputs[idx] if idx<len(outputs) else '')
        if not text:
            continue

        p=doc.add_paragraph()
        p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing=1.0
        p.paragraph_format.space_after=Pt(7)
        p.add_run(text.rstrip(' .;')+' (STB).')

    doc.add_paragraph('Arz olunur.')
    bio=BytesIO(); doc.save(bio); bio.seek(0)
    return bio.getvalue()

# ======================= /V100: KONU BÜTÜNLÜĞÜ =========================
def make_important_basket_docx_v99(basket_df):
    """
    V99 ÖGN Word:
    - sepetteki her haber ayrı işlenir;
    - web sayfası/UI gürültüsü atılır;
    - başlık/site adı basılmaz;
    - içerik bulunamayan haber başlık olarak Word'e sokulmaz;
    - resmî, analitik, 4 satıra yakın özet üretilir;
    - Türkçe mojibake temizliği son katmanda tekrar uygulanır.
    """
    doc=Document()
    sec=doc.sections[0]
    sec.top_margin=Cm(2); sec.bottom_margin=Cm(2)
    sec.left_margin=Cm(2.5); sec.right_margin=Cm(2.5)

    normal=doc.styles['Normal']
    normal.font.name='Times New Roman'
    normal.font.size=Pt(12)
    normal._element.rPr.rFonts.set(qn('w:eastAsia'),'Times New Roman')

    now=datetime.now().astimezone()
    p=doc.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    p.add_run(now.strftime('%d/%m/%Y'))

    p=doc.add_paragraph()
    p.add_run('Konu: ').bold=True
    p.add_run('STB Temsilciliği Önemli Gelişmeler Notu')

    rows=[] if basket_df is None else basket_df.to_dict('records')

    def process_row(r):
        title=_v98_safe_tr(r.get('title',''))
        source=_v98_safe_tr(r.get('source',''))
        summary=_v98_safe_tr(r.get('summary',''))
        url=str(r.get('url','') or '')
        news_time=_v98_safe_tr(r.get('news_time',''))

        detail={}
        try:
            detail=article_detail({
                'Başlık':title,'Kaynak':source,'URL':url,'Yayıncı_URL':url,
                'İçerik_Özeti':summary,'Tarih':news_time
            }) or {}
        except Exception:
            pass

        body=detail.get('text') or ''
        text=_v99_analyst_ogn(title,source,body,summary)

        # İçerik çekilememişse bir kez daha mevcut summary ile dene.
        if not text and summary:
            text=_v99_analyst_ogn(title,source,summary,summary)

        return text

    outputs=['']*len(rows)
    if rows:
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(6,len(rows))) as ex:
            jobs={ex.submit(process_row,r):i for i,r in enumerate(rows)}
            for fut in concurrent.futures.as_completed(jobs):
                idx=jobs[fut]
                try:
                    outputs[idx]=fut.result()
                except Exception:
                    outputs[idx]=''

    for idx,r in enumerate(rows):
        text=_v98_safe_tr(outputs[idx] if idx<len(outputs) else '')
        if not text:
            # Başlığı çıktı olarak kullanma. İçeriği olmayan haberi uyarı metniyle bozmak yerine atla.
            continue

        p=doc.add_paragraph()
        p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing=1.0
        p.paragraph_format.space_after=Pt(7)
        p.add_run(text.rstrip(' .;')+' (STB).')

    doc.add_paragraph('Arz olunur.')

    bio=BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()
# ======================= /V99: ÖGN ANALİST FİLTRESİ =========================
def make_important_basket_docx_v98(basket_df):
    """
    ÖGN Word:
    1) Sepetteki HER haber Word'e aktarılır.
    2) Her haber tek paragraf / yaklaşık 4 satır.
    3) Resmî dil.
    4) Türkçe karakter temizliği.
    5) Başlık yok.
    6) Site/yayıncı adı yok.
    """
    doc=Document()
    sec=doc.sections[0]
    sec.top_margin=Cm(2); sec.bottom_margin=Cm(2)
    sec.left_margin=Cm(2.5); sec.right_margin=Cm(2.5)

    normal=doc.styles['Normal']
    normal.font.name='Times New Roman'
    normal.font.size=Pt(12)
    normal._element.rPr.rFonts.set(qn('w:eastAsia'),'Times New Roman')

    now=datetime.now().astimezone()
    p=doc.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    p.add_run(now.strftime('%d/%m/%Y'))

    p=doc.add_paragraph()
    p.add_run('Konu: ').bold=True
    p.add_run('STB Temsilciliği Önemli Gelişmeler Notu')

    rows=[] if basket_df is None else basket_df.to_dict('records')

    def process_row(r):
        title=_v98_safe_tr(r.get('title',''))
        source=_v98_safe_tr(r.get('source',''))
        summary=_v98_safe_tr(r.get('summary',''))
        url=str(r.get('url','') or '')
        news_time=_v98_safe_tr(r.get('news_time',''))

        try:
            detail=article_detail({
                'Başlık':title,
                'Kaynak':source,
                'URL':url,
                'Yayıncı_URL':url,
                'İçerik_Özeti':summary,
                'Tarih':news_time
            })
        except Exception:
            detail={}

        body=_v98_safe_tr(detail.get('text') or summary or title)
        out=_v98_exact_four_line_summary(title,source,body)

        # Her haber mutlaka çıksın.
        if not out:
            out=_v98_exact_four_line_summary(title,source,summary or title)

        # Son çare: başlığı doğrudan yazmak yerine kısa resmî cümleye çevir.
        if not out:
            clean_title=_v98_remove_heading_like_text(title,title)
            if not clean_title:
                clean_title=title.capitalize() if title else 'Gelişme'
            out=_v66_formalize_sentence_endings(clean_title).strip()
            if out and out[-1] not in '.!?':
                out+='.'

        return out

    outputs=['']*len(rows)
    if rows:
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(6,len(rows))) as ex:
            jobs={ex.submit(process_row,r):i for i,r in enumerate(rows)}
            for fut in concurrent.futures.as_completed(jobs):
                idx=jobs[fut]
                try:
                    outputs[idx]=fut.result()
                except Exception:
                    r=rows[idx]
                    outputs[idx]=_v98_exact_four_line_summary(
                        r.get('title',''),
                        r.get('source',''),
                        r.get('summary','') or r.get('title','')
                    )

    # Her sepet haberi sırayla Word'e yazılır.
    for idx,r in enumerate(rows):
        text=outputs[idx] if idx<len(outputs) else ''
        text=_v98_safe_tr(text)

        # Hiçbir haber sessizce atlanmasın.
        if not text:
            title=_v98_safe_tr(r.get('title',''))
            text=_v66_formalize_sentence_endings(title.capitalize()).strip()
            if text and text[-1] not in '.!?':
                text+='.'

        p=doc.add_paragraph()
        p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing=1.0
        p.paragraph_format.space_after=Pt(7)

        # Başlık/link/site ayrıca yazılmaz.
        p.add_run(text.rstrip(' .;')+' (STB).')

    doc.add_paragraph('Arz olunur.')

    bio=BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()

def make_important_basket_docx_v97(basket_df):
    """
    Önemli Gelişmeler Word:
    Her haber TEK paragraf, yaklaşık 4 satır.
    Ana gelişme + kritik veri/rakam + sonuç.
    Başlık ve haber linki yazılmaz.
    """
    doc=Document()
    sec=doc.sections[0]
    sec.top_margin=Cm(2); sec.bottom_margin=Cm(2)
    sec.left_margin=Cm(2.5); sec.right_margin=Cm(2.5)

    normal=doc.styles['Normal']
    normal.font.name='Times New Roman'
    normal.font.size=Pt(12)
    normal._element.rPr.rFonts.set(qn('w:eastAsia'),'Times New Roman')

    now=datetime.now().astimezone()
    p=doc.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    p.add_run(now.strftime('%d/%m/%Y'))

    p=doc.add_paragraph()
    p.add_run('Konu: ').bold=True
    p.add_run('STB Temsilciliği Önemli Gelişmeler Notu')

    rows=[] if basket_df is None else basket_df.to_dict('records')

    def process_row(r):
        title=_clean_note_text(r.get('title',''))
        source=_clean_note_text(r.get('source',''))
        summary=_clean_note_text(r.get('summary',''))
        url=str(r.get('url','') or '')
        news_time=_clean_note_text(r.get('news_time',''))

        try:
            detail=article_detail({
                'Başlık':title,
                'Kaynak':source,
                'URL':url,
                'Yayıncı_URL':url,
                'İçerik_Özeti':summary,
                'Tarih':news_time
            })
        except Exception:
            detail={}

        body=_clean_note_text(detail.get('text') or summary or title)
        out=_v97_analyst_summary(title,body)

        if not out:
            out=_v97_analyst_summary(title,summary or title)

        return out

    outputs=['']*len(rows)
    if rows:
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(6,len(rows))) as ex:
            jobs={ex.submit(process_row,r):i for i,r in enumerate(rows)}
            for fut in concurrent.futures.as_completed(jobs):
                idx=jobs[fut]
                try:
                    outputs[idx]=fut.result()
                except Exception:
                    r=rows[idx]
                    outputs[idx]=_v97_analyst_summary(
                        r.get('title',''),
                        r.get('summary','') or r.get('title','')
                    )

    for text in outputs:
        text=_repair_mojibake_utf8(_clean_note_text(text)).strip()
        if not text:
            continue
        p=doc.add_paragraph()
        p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing=1.0
        p.paragraph_format.space_after=Pt(7)
        # Her haber tek bütünleşik paragraf.
        p.add_run(text.rstrip(' .;')+' (STB).')

    doc.add_paragraph('Arz olunur.')

    bio=BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()

def make_important_basket_docx_v96(basket_df):
    """
    Önemli Gelişmeler Word = her haber için ayrı ayrı kısaltılmış bilgi notu.
    Başlık/link yazılmaz. Her haber en fazla 4 paragraftır.
    Kritik veri, rakam ve istatistikler yüksek öncelikle korunur.
    """
    doc=Document()
    sec=doc.sections[0]
    sec.top_margin=Cm(2); sec.bottom_margin=Cm(2)
    sec.left_margin=Cm(2.5); sec.right_margin=Cm(2.5)

    normal=doc.styles['Normal']
    normal.font.name='Times New Roman'
    normal.font.size=Pt(12)
    normal._element.rPr.rFonts.set(qn('w:eastAsia'),'Times New Roman')

    now=datetime.now().astimezone()
    p=doc.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    p.add_run(now.strftime('%d/%m/%Y'))

    p=doc.add_paragraph()
    p.add_run('Konu: ').bold=True
    p.add_run('STB Temsilciliği Önemli Gelişmeler Notu')

    rows=[] if basket_df is None else basket_df.to_dict('records')

    def process_row(r):
        title=_clean_note_text(r.get('title',''))
        source=_clean_note_text(r.get('source',''))
        summary=_clean_note_text(r.get('summary',''))
        url=str(r.get('url','') or '')
        news_time=_clean_note_text(r.get('news_time',''))

        # Bilgi Notu ile aynı içerik alma sistemi.
        try:
            detail=article_detail({
                'Başlık':title,
                'Kaynak':source,
                'URL':url,
                'Yayıncı_URL':url,
                'İçerik_Özeti':summary,
                'Tarih':news_time
            })
        except Exception:
            detail={}

        body=_clean_note_text(detail.get('text') or summary or title)
        paras=_v96_short_information_note(title,body)

        # Tam metin sorunluysa kayıtlı özet üzerinde aynı motoru tekrar çalıştır.
        if not paras:
            paras=_v96_short_information_note(title,summary or title)

        return paras

    outputs=[[] for _ in rows]
    if rows:
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(6,len(rows))) as ex:
            jobs={ex.submit(process_row,r):i for i,r in enumerate(rows)}
            for fut in concurrent.futures.as_completed(jobs):
                idx=jobs[fut]
                try:
                    outputs[idx]=fut.result()
                except Exception:
                    r=rows[idx]
                    outputs[idx]=_v96_short_information_note(
                        r.get('title',''),
                        r.get('summary','') or r.get('title','')
                    )

    for item_index,paras in enumerate(outputs):
        if not paras:
            continue

        for para_index,text in enumerate(paras[:4]):
            text=_repair_mojibake_utf8(_clean_note_text(text)).strip()
            if not text:
                continue
            p=doc.add_paragraph()
            p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent=Cm(1.25)
            p.paragraph_format.line_spacing=1.0
            p.paragraph_format.space_after=Pt(4 if para_index<len(paras)-1 else 8)
            # (STB) yalnız haberin son paragrafında.
            suffix=' (STB).' if para_index==len(paras[:4])-1 else ''
            p.add_run(text.rstrip(' .;')+suffix)

    doc.add_paragraph('Arz olunur.')

    bio=BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()

def make_important_basket_docx_v95(basket_df):
    """
    Önemli Gelişmeler Word:
    - Haber işleme: Bilgi Notu gibi article_detail()
    - Özetleme: AKT gibi _akt_formal_summary()
    - Dil: mevcut V66 resmî dil normalizasyonu
    - Çıktı: başlıksız, linksiz, yaklaşık 4 satır.
    """
    doc=Document()
    sec=doc.sections[0]
    sec.top_margin=Cm(2); sec.bottom_margin=Cm(2)
    sec.left_margin=Cm(2.5); sec.right_margin=Cm(2.5)

    normal=doc.styles['Normal']
    normal.font.name='Times New Roman'
    normal.font.size=Pt(12)
    normal._element.rPr.rFonts.set(qn('w:eastAsia'),'Times New Roman')

    now=datetime.now().astimezone()
    p=doc.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    p.add_run(now.strftime('%d/%m/%Y'))

    p=doc.add_paragraph()
    p.add_run('Konu: ').bold=True
    p.add_run('STB Temsilciliği Önemli Gelişmeler Notu')

    rows=[] if basket_df is None else basket_df.to_dict('records')

    def process_row(r):
        title=_clean_note_text(r.get('title',''))
        source=_clean_note_text(r.get('source',''))
        summary=_clean_note_text(r.get('summary',''))
        url=str(r.get('url','') or '')
        news_time=_clean_note_text(r.get('news_time',''))

        # Bilgi Notu ile aynı yaklaşım: mümkünse gerçek haber metnini al.
        try:
            detail=article_detail({
                'Başlık':title,
                'Kaynak':source,
                'URL':url,
                'Yayıncı_URL':url,
                'İçerik_Özeti':summary,
                'Tarih':news_time
            })
        except Exception:
            detail={}

        body=_clean_note_text(detail.get('text') or summary or title)

        # AKT'nin çalışan özet motorunu doğrudan kullan.
        result=_v95_ogn_from_existing_engines(title,body)

        # Tam metin tarafı sonuç vermezse AKT motorunu kayıtlı özet üzerinde çalıştır.
        if not result or len(result)<50:
            result=_v95_ogn_from_existing_engines(title,summary or title)

        return result

    # Bilgi notu/AKT içerik yaklaşımı korunurken Word beklemesini azaltmak için paralel oku.
    outputs=['']*len(rows)
    if rows:
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(6,len(rows))) as ex:
            jobs={ex.submit(process_row,r):i for i,r in enumerate(rows)}
            for fut in concurrent.futures.as_completed(jobs):
                idx=jobs[fut]
                try:
                    outputs[idx]=fut.result()
                except Exception:
                    r=rows[idx]
                    outputs[idx]=_v95_ogn_from_existing_engines(
                        r.get('title',''),
                        r.get('summary','') or r.get('title','')
                    )

    for text in outputs:
        text=_clean_note_text(text)
        if not text:
            continue
        p=doc.add_paragraph()
        p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after=Pt(6)
        p.paragraph_format.line_spacing=1.0
        # Haber başlığı, kaynak adı veya URL ayrıca yazılmaz.
        p.add_run(text.rstrip(' .;')+' (STB).')

    doc.add_paragraph('Arz olunur.')

    bio=BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()

def make_important_basket_docx_v94(basket_df):
    """
    Reliable ÖGN Word generator.
    No headline, no URL, no bullet.
    Every basket row produces one paragraph.
    Primary source = summary already stored when the news was scanned.
    """
    doc=Document()
    sec=doc.sections[0]
    sec.top_margin=Cm(2); sec.bottom_margin=Cm(2)
    sec.left_margin=Cm(2.5); sec.right_margin=Cm(2.5)

    stl=doc.styles['Normal']
    stl.font.name='Times New Roman'
    stl.font.size=Pt(12)
    stl._element.rPr.rFonts.set(qn('w:eastAsia'),'Times New Roman')

    now=datetime.now().astimezone()
    p0=doc.add_paragraph()
    p0.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    p0.add_run(now.strftime('%d/%m/%Y'))

    p1=doc.add_paragraph()
    p1.add_run('Konu: ').bold=True
    p1.add_run('STB Temsilciliği Önemli Gelişmeler Notu')

    rows=[] if basket_df is None else basket_df.to_dict('records')

    for r in rows:
        title=_v87_safe_tr(r.get('title',''))
        source=_v87_safe_tr(r.get('source',''))
        summary=_v87_safe_tr(r.get('summary',''))
        url=str(r.get('url','') or '')
        news_time=_v87_safe_tr(r.get('news_time',''))

        # Use the text already saved in the basket first: fast and stable.
        text=summary

        # Only if the stored summary is genuinely empty/too short, try the article.
        if len(text.strip())<60 and url:
            try:
                d=article_detail({
                    'Başlık':title,'Kaynak':source,'URL':url,'Yayıncı_URL':url,
                    'İçerik_Özeti':summary,'Tarih':news_time
                })
                fetched=_v87_safe_tr((d or {}).get('text',''))
                if len(fetched)>len(text):
                    text=fetched
            except Exception:
                pass

        out=_v94_formal_summary_text(text)

        # Absolute fallback: do not create an empty Word document.
        # If scan stored no summary and article could not be read, use a cleaned
        # sentence from the title rather than silently omitting the news.
        if not out:
            out=_v92_formal_sentence(title).strip()
            if out and out[-1] not in '.!?':
                out+='.'

        if not out:
            continue

        p=doc.add_paragraph()
        p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after=Pt(6)
        p.paragraph_format.line_spacing=1.0
        p.add_run(out.rstrip(' .;')+' (STB).')

    doc.add_paragraph('Arz olunur.')
    bio=BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()

def make_important_basket_docx_v93(basket_df):
    """STB referansına göre sıfırdan yazılmış Önemli Gelişmeler Word motoru."""
    doc=Document()
    sec=doc.sections[0]
    sec.top_margin=Cm(2); sec.bottom_margin=Cm(2)
    sec.left_margin=Cm(2.5); sec.right_margin=Cm(2.5)

    style=doc.styles['Normal']
    style.font.name='Times New Roman'
    style.font.size=Pt(12)
    style._element.rPr.rFonts.set(qn('w:eastAsia'),'Times New Roman')

    now=datetime.now().astimezone()
    p=doc.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    p.add_run(now.strftime('%d/%m/%Y'))

    p=doc.add_paragraph()
    p.add_run('Konu: ').bold=True
    p.add_run('STB Temsilciliği Önemli Gelişmeler Notu')

    rows=[] if basket_df is None else basket_df.to_dict('records')

    def one(r):
        title=_v87_safe_tr(r.get('title',''))
        source=_v87_safe_tr(r.get('source',''))
        fallback=_v87_safe_tr(r.get('summary',''))
        url=str(r.get('url','') or '')
        news_time=_v87_safe_tr(r.get('news_time',''))
        try:
            d=_v90_fetch_detail(title,source,url,fallback,news_time)
            body=_v87_safe_tr((d or {}).get('text','') or fallback)
        except Exception:
            body=fallback
        return _v93_build_summary(title,source,body,fallback)

    summaries=['']*len(rows)
    if rows:
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(8,len(rows))) as ex:
            jobs={ex.submit(one,r):i for i,r in enumerate(rows)}
            for f in concurrent.futures.as_completed(jobs):
                i=jobs[f]
                try: summaries[i]=f.result()
                except Exception: summaries[i]=''

    for text in summaries:
        if not text:
            continue
        p=doc.add_paragraph()
        p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after=Pt(6)
        p.paragraph_format.line_spacing=1.0
        # Başlık yok, URL yok, madde imi yok.
        p.add_run(text.rstrip(' .;')+' (STB).')

    doc.add_paragraph('Arz olunur.')
    bio=BytesIO(); doc.save(bio); bio.seek(0)
    return bio.getvalue()

def make_important_basket_docx_v92(basket_df):
    """
    V92: ÖGN için sade ve izlenebilir akış.
    Haberleri paralel alır; her haber için tek gövde + ardışık 2-3 cümle.
    """
    doc=Document()
    sec=doc.sections[0]
    sec.top_margin=Cm(2); sec.bottom_margin=Cm(2)
    sec.left_margin=Cm(2.5); sec.right_margin=Cm(2.5)

    normal=doc.styles['Normal']
    normal.font.name='Times New Roman'
    normal.font.size=Pt(12)
    normal._element.rPr.rFonts.set(qn('w:eastAsia'),'Times New Roman')

    now=datetime.now().astimezone()
    p=doc.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    p.add_run(now.strftime('%d/%m/%Y')).bold=True

    p=doc.add_paragraph()
    p.add_run('Konu: ').bold=True
    p.add_run('STB Temsilciliği Önemli Gelişmeler Notu')

    records=[] if basket_df is None else basket_df.to_dict('records')

    def process(item):
        title=_v87_safe_tr(item.get('title',''))
        source=_v87_safe_tr(item.get('source',''))
        fallback=_v87_safe_tr(item.get('summary',''))
        url=str(item.get('url','') or '')
        news_time=_v87_safe_tr(item.get('news_time',''))

        detail=_v90_fetch_detail(title,source,url,fallback,news_time)
        body=_v87_safe_tr((detail or {}).get('text','') or fallback)

        txt=_v92_summary(title,source,body,fallback)
        if not txt:
            txt=_v92_summary(title,source,fallback,fallback)
        return txt

    summaries=['']*len(records)
    if records:
        workers=min(8,max(1,len(records)))
        with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as ex:
            fmap={ex.submit(process,r):i for i,r in enumerate(records)}
            for fut in concurrent.futures.as_completed(fmap):
                idx=fmap[fut]
                try:
                    summaries[idx]=fut.result()
                except Exception:
                    summaries[idx]=''

    for txt in summaries:
        if not txt:
            continue
        p=doc.add_paragraph()
        p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after=Pt(6)
        p.paragraph_format.line_spacing=1.0
        p.add_run(_v87_safe_tr(txt).rstrip(' .;')+' (STB).')

    p=doc.add_paragraph()
    p.paragraph_format.space_before=Pt(8)
    p.add_run('Arz olunur.')

    bio=BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()

def make_important_basket_docx_v90(basket_df):
    """
    V90 ÖGN motoru. Eski Word baytlarını/fonksiyonlarını kullanmaz.
    Her haber için gerçek metni paralel alır, sırayı korur.
    """
    doc=Document(); sec=doc.sections[0]
    sec.top_margin=Cm(2); sec.bottom_margin=Cm(2)
    sec.left_margin=Cm(2.5); sec.right_margin=Cm(2.5)
    normal=doc.styles['Normal']
    normal.font.name='Times New Roman'; normal.font.size=Pt(12)
    normal._element.rPr.rFonts.set(qn('w:eastAsia'),'Times New Roman')

    now=datetime.now().astimezone()
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    p.add_run(now.strftime('%d/%m/%Y')).bold=True
    p=doc.add_paragraph()
    p.add_run('Konu: ').bold=True
    p.add_run('STB Temsilciliği Önemli Gelişmeler Notu')

    records=[] if basket_df is None else basket_df.to_dict('records')
    if not records:
        doc.add_paragraph('Kayıtlı önemli gelişme bulunmamaktadır.')
    else:
        def process(item):
            title=_v87_safe_tr(item.get('title',''))
            source=_v87_safe_tr(item.get('source',''))
            fallback=_v87_safe_tr(item.get('summary',''))
            url=str(item.get('url','') or '')
            news_time=_v87_safe_tr(item.get('news_time',''))

            detail=_v90_fetch_detail(title,source,url,fallback,news_time)
            body=_v87_safe_tr((detail or {}).get('text','') or fallback)
            txt=_v90_item_summary(title,source,body,fallback)

            # Asla eski başlık-çıktı davranışına dönme.
            if not txt:
                # fallback gövdesinden tek resmî cümle oluşturmayı tekrar dene.
                txt=_v90_item_summary(title,source,fallback,fallback)
            if not txt:
                # Son çare: başlığı değil, açıklayıcı bir kurum cümlesi oluştur.
                clean_title=_v90_clean_title(title,source)
                txt=f'{clean_title} konusuna ilişkin gelişme açık kaynaklarda yer almıştır.'
                txt=_v90_formalize(txt)
            return txt

        summaries=['']*len(records)
        workers=min(8,max(1,len(records)))
        with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as ex:
            fmap={ex.submit(process,r):i for i,r in enumerate(records)}
            for fut in concurrent.futures.as_completed(fmap):
                idx=fmap[fut]
                try:
                    summaries[idx]=fut.result()
                except Exception:
                    summaries[idx]=''

        for rr,txt in zip(records,summaries):
            if not txt:
                continue
            p=doc.add_paragraph()
            p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.space_after=Pt(6)
            p.paragraph_format.line_spacing=1.0
            p.add_run(_v87_safe_tr(txt).rstrip(' .;')+' (STB).')

    p=doc.add_paragraph()
    p.paragraph_format.space_before=Pt(8)
    p.add_run('Arz olunur.')

    bio=BytesIO(); doc.save(bio); bio.seek(0)
    return bio.getvalue()
def make_important_basket_docx(basket_df):
    """
    V88:
    - article fetches run in parallel instead of one-by-one,
    - results cached for 1 hour,
    - output order remains basket order,
    - no item is silently dropped.
    """
    doc=Document(); sec=doc.sections[0]
    sec.top_margin=Cm(2); sec.bottom_margin=Cm(2); sec.left_margin=Cm(2.5); sec.right_margin=Cm(2.5)
    normal=doc.styles['Normal']
    normal.font.name='Times New Roman'; normal.font.size=Pt(12)
    normal._element.rPr.rFonts.set(qn('w:eastAsia'),'Times New Roman')

    now=datetime.now().astimezone()
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    p.add_run(now.strftime('%d/%m/%Y')).bold=True
    p=doc.add_paragraph(); p.add_run('Konu: ').bold=True
    p.add_run('STB Temsilciliği Önemli Gelişmeler Notu')

    if basket_df is None or basket_df.empty:
        doc.add_paragraph('Kayıtlı önemli gelişme bulunmamaktadır.')
    else:
        records=basket_df.to_dict('records')

        def fetch_one(item):
            title=_v87_safe_tr(item.get('title',''))
            source=_v87_safe_tr(item.get('source',''))
            fallback=_v87_safe_tr(item.get('summary',''))
            url=str(item.get('url','') or '')
            news_time=_v87_safe_tr(item.get('news_time',''))

            # If saved summary is already substantial, don't delay Word just to fetch again.
            # Full article is requested mainly for short/snippet-like summaries.
            detail={}
            if len(fallback)<380:
                detail=_v88_cached_article_detail(title,source,url,fallback,news_time)
            body=_v87_safe_tr((detail or {}).get('text','') or fallback)
            return _v88_summary(title,source,body,fallback)

        summaries=['']*len(records)
        max_workers=min(6,max(1,len(records)))
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as ex:
            futmap={ex.submit(fetch_one,r):idx for idx,r in enumerate(records)}
            for fut in concurrent.futures.as_completed(futmap):
                idx=futmap[fut]
                try:
                    summaries[idx]=fut.result()
                except Exception:
                    rr=records[idx]
                    summaries[idx]=_v88_summary(
                        rr.get('title',''),rr.get('source',''),
                        rr.get('summary',''),rr.get('summary','')
                    )

        for rr,txt in zip(records,summaries):
            if not txt:
                txt=_v88_formal(_v87_safe_tr(rr.get('summary','') or rr.get('title','')))
            p=doc.add_paragraph()
            p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.space_after=Pt(5)
            p.paragraph_format.line_spacing=1.0
            p.add_run(_v87_safe_tr(txt).rstrip(' .;')+' (STB).')

    p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(8)
    p.add_run('Arz olunur.')
    bio=BytesIO(); doc.save(bio); bio.seek(0)
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
    """
    V62: Sabah ilk bakılacak gelişmeleri seçer.
    Negatiflik tek başına belirleyici değildir. Stratejik sanayi-teknoloji ilgisi,
    ekonomik/kurumsal etki, resmî teyit, çoklu kaynak, yenilik ve risk birlikte puanlanır.
    """
    if df.empty:
        return df.copy()

    x=df.copy()
    x['Tarih_dt']=pd.to_datetime(x.get('Tarih_dt'),utc=True,errors='coerce')
    strategic_terms=[
        'yatırım','üretim','ihracat','ithalat','kapasite','fabrika','tesis','osb',
        'savunma','tusaş','aselsan','roketsan','havelsan','baykar','kaan',
        'yapay zeka','yapay zekâ','çip','yarı iletken','siber','teknoloji',
        'arge','ar-ge','tübitak','kosgeb','patent','togg','otomotiv','enerji',
        'kritik mineral','uzay','uydu','teknofest','sanayi üretimi'
    ]
    high_value_terms=[
        'milyar','milyon','rekor','anlaşma','sözleşme','yatırım','teşvik',
        'ihracat','üretim','kapasite','lansman','ilk kez','yeni tesis',
        'stratejik','program','eylem planı','resmi gazete','resmî gazete'
    ]
    low_relevance_terms=[
        'trafik kazası','magazin','spor','dualarla anıldı','hayatını kaybeden muhabir'
    ]

    def importance(r):
        text=norm(f"{r.get('Başlık','')} {r.get('İçerik_Özeti','')} {r.get('Kategori','')}")
        score=0

        # Sanayi-teknoloji alanına doğrudan ilgi en güçlü ölçüt.
        score += min(sum(1 for k in strategic_terms if k in text)*8,40)
        score += min(sum(1 for k in high_value_terms if k in text)*5,20)

        cat=norm(r.get('Kategori',''))
        if any(k in cat for k in ['savunma','sanayi','üretim','dijital','yapay zeka','yapay zekâ',
                                  'otomotiv','uzay','enerji','teknoloji']):
            score+=18

        # Risk önemlidir ama negatiflik listeyi ele geçirmez.
        risk=int(r.get('Risk_Skoru',0) or 0)
        score+=min(risk//4,20)
        if r.get('Risk_Durumu')=='Yüksek Risk':
            score+=12
        if r.get('Duygu')=='Negatif':
            score+=5

        if critical_industrial_incident(r.get('Başlık',''),r.get('İçerik_Özeti','')):
            score+=25

        try:
            score+=min(int(r.get('Olay_Kaynak_Sayisi',0) or 0)*5,20)
        except Exception:
            pass

        verification=norm(r.get('Doğrulama',''))
        if 'resmi' in verification or 'resmî' in verification or 'birincil' in verification:
            score+=22
        elif 'çoklu kaynak' in verification or 'coklu kaynak' in verification:
            score+=14

        if any(k in text for k in low_relevance_terms):
            score-=30

        return score

    x['_Önem']=x.apply(importance,axis=1)

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
        f"{stats['defence']} savunma sanayii bağlantılı içerik ve {stats['cyber']} siber güvenlik bağlantılı içerik belirlenmiştir."
    )

    paras=[intro]
    if not top.empty:
        paras.append(
            "Günün genel görünümünde öne çıkan gelişmeler; güncellik, risk düzeyi, kaynak teyidi ve sanayi-teknoloji alanına muhtemel etkileri "
            "birlikte dikkate alınarak aşağıda özetlenmiştir."
        )
        for i,(_,r) in enumerate(top.iterrows(),1):
            title=_clean_note_text(r.get('Başlık',''))
            source=_clean_note_text(r.get('Kaynak','Açık Kaynak'))
            when=_clean_note_text(r.get('Tarih',''))
            content=_clean_note_text(r.get('İçerik_Özeti',''))

            # Başlığı tekrar etmek yerine içerikten anlamlı cümleleri seç.
            sents=_detail_sentences(content,title)
            useful=[]
            seen=set()
            for s in sents:
                s=_clean_note_text(s)
                key=norm(s)
                if not s or len(s)<35 or key in seen:
                    continue
                seen.add(key)
                useful.append(s)
                if len(useful)>=4:
                    break

            detail=_join_sentences_naturally(useful) if useful else content[:700].strip()
            risk=int(r.get('Risk_Skoru',0) or 0)
            status=_clean_note_text(r.get('Risk_Durumu',''))
            category=_clean_note_text(r.get('Kategori',''))

            p=f"{i}. {when} tarihinde {source} kaynaklı gelişmede, {detail}" if detail else f"{i}. {when} tarihinde {source} kaynaklı “{title}” başlıklı gelişme öne çıkmıştır."
            if p and p[-1] not in '.!?':
                p+='.'
            if category:
                p+=f" Gelişme sistemde {category} başlığı altında izlenmektedir."
            if risk:
                p+=f" Risk puanı {risk}/100"
                if status:
                    p+=f" ve risk durumu {status}"
                p+=" olarak değerlendirilmiştir."
            paras.append(p)

    # Günlük tabloya dair kısa analitik kapanış.
    emphasis=[]
    if stats['high_risk']:
        emphasis.append(f"{stats['high_risk']} yüksek riskli gelişmenin")
    if stats['negative']:
        emphasis.append(f"{stats['negative']} negatif içeriğin")
    if stats['investment']:
        emphasis.append(f"{stats['investment']} yatırım/kapasite gelişmesinin")
    if stats['defence']:
        emphasis.append(f"{stats['defence']} savunma sanayii gelişmesinin")
    if stats['cyber']:
        emphasis.append(f"{stats['cyber']} siber güvenlik gelişmesinin")

    if emphasis:
        focus=', '.join(emphasis[:-1]) + ((' ve '+emphasis[-1]) if len(emphasis)>1 else emphasis[0])
        conclusion=(
            f"Günlük görünümde özellikle {focus} takip edilmesi gereken başlıklar arasında bulunduğu değerlendirilmektedir. "
            "Yeni resmî açıklamalar, üretim ve tedarik zincirine olası etkiler ile farklı açık kaynaklardan gelecek teyitlerin izlenmesi önem taşımaktadır."
        )
    else:
        conclusion=(
            "Günlük görünümde belirgin bir yüksek risk yoğunlaşması görülmemekle birlikte, yeni resmî açıklamalar ile üretim, yatırım, "
            "tedarik zinciri ve teknoloji alanındaki gelişmelerin izlenmesinin sürdürülmesi önem taşımaktadır."
        )
    paras.append(conclusion)
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
        # V92: Birden fazla article/main/content bloğunu BİRLEŞTİRME.
        # Aynı sayfadaki önerilen haberler ve tekrar blokları ÖGN'ye karışmasın.
        # Tek, en kapsamlı gövdeyi kullan.
        texts.sort(key=len, reverse=True)
        if texts:
            out["text"]=texts[0][:18000]

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


# -----------------------------
# V43 — TAM HABER METNİNE GÖRE NEGATİF/RİSK ANALİZİ
# -----------------------------
def _deep_negative_reclassify(rows, max_workers=14):
    """
    Her haberi mümkünse gerçek haber sayfasındaki tam metinle yeniden sınıflandırır.
    Sayfaya erişilemezse mevcut başlık + kısa içerik fallback olur.

    Yalnızca negatif/risk alanları güncellenir; kategori, olay kümeleri ve diğer
    çalışan modüller korunur.
    """
    if not rows:
        return rows, {'tam_metin':0,'kisa_icerik':0,'hata':0}

    results=[None]*len(rows)
    stats={'tam_metin':0,'kisa_icerik':0,'hata':0}

    def one(idx,row):
        try:
            detail=article_detail(row)
            full_text=re.sub(r'\s+',' ',str(detail.get('text') or '')).strip()
            snippet=re.sub(r'\s+',' ',str(row.get('İçerik_Özeti') or '')).strip()

            # article_detail erişemezse fallback olarak snippet döndürebilir.
            is_full=bool(full_text) and len(full_text)>=max(450,len(snippet)+180)
            analysis_text=full_text if is_full else (snippet or full_text or row.get('Başlık',''))

            sentiment,score,status,neg,risk,_cat,reasons=classify(
                row.get('Başlık',''),
                analysis_text,
                row.get('Domain','')
            )

            return idx,{
                'Duygu':sentiment,
                'Skor':score,
                'Risk_Skoru':score,
                'Risk_Durumu':status,
                'Risk_Gerekçesi':'; '.join(reasons),
                'Negatif_Sinyaller':neg,
                'Risk_Sinyalleri':risk,
                'Negatif_Analiz_Kapsamı':'Tam haber metni' if is_full else 'Başlık + kısa içerik',
                '_is_full':is_full
            }
        except Exception:
            sentiment,score,status,neg,risk,_cat,reasons=classify(
                row.get('Başlık',''),
                row.get('İçerik_Özeti',''),
                row.get('Domain','')
            )
            return idx,{
                'Duygu':sentiment,
                'Skor':score,
                'Risk_Skoru':score,
                'Risk_Durumu':status,
                'Risk_Gerekçesi':'; '.join(reasons),
                'Negatif_Sinyaller':neg,
                'Risk_Sinyalleri':risk,
                'Negatif_Analiz_Kapsamı':'Başlık + kısa içerik',
                '_is_full':False,
                '_error':True
            }

    with concurrent.futures.ThreadPoolExecutor(max_workers=min(max_workers,len(rows))) as ex:
        futures=[ex.submit(one,i,r.copy()) for i,r in enumerate(rows)]
        for fut in concurrent.futures.as_completed(futures):
            try:
                idx,data=fut.result()
                results[idx]=data
            except Exception:
                pass

    out=[]
    for i,row in enumerate(rows):
        r=row.copy()
        data=results[i]
        if data:
            if data.pop('_is_full',False):
                stats['tam_metin']+=1
            else:
                stats['kisa_icerik']+=1
            if data.pop('_error',False):
                stats['hata']+=1
            r.update(data)
        else:
            stats['kisa_icerik']+=1
            r['Negatif_Analiz_Kapsamı']='Başlık + kısa içerik'
        out.append(r)

    return out,stats


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

def _akt_clean_sentences(title, body):
    text=BeautifulSoup(str(body or ''),'html.parser').get_text(' ',strip=True)
    text=re.sub(r'\s+',' ',text).strip()
    if not text:
        return []

    raw=re.split(r'(?<=[.!?])\s+',text)
    title_n=norm(title)
    boiler=[
        'çerez','cookie','abonelik','abone ol','reklam','tüm hakları saklıdır',
        'gizlilik politikası','kullanım koşulları','google news','bildirimleri aç',
        'uygulamamızı indirin','facebook','instagram','whatsapp','twitter',
        'son dakika haberleri için','haberlerimizi takip','ilgili haberler',
        'öne çıkan haberler','etiketler','yorumlar'
    ]

    kept=[]
    token_sets=[]
    for s in raw:
        s=re.sub(r'\s+',' ',s).strip()
        sn=norm(s)
        if len(s)<32 or sn==title_n:
            continue
        if any(b in sn for b in boiler):
            continue

        toks={x for x in re.findall(r'\w+',sn) if len(x)>2}
        if not toks:
            continue

        duplicate=False
        for old in token_sets[-20:]:
            inter=len(toks & old); union=len(toks | old)
            if union and inter/union>=0.78:
                duplicate=True
                break
        if duplicate:
            continue

        kept.append(s)
        token_sets.append(toks)

    return kept

def _akt_sentence_score(s):
    n=norm(s)
    score=0
    if re.search(r'\b\d+(?:[.,]\d+)?\b',s): score+=4
    if '%' in s or 'yüzde' in n: score+=3
    if any(x in n for x in ['açıkladı','belirtti','bildirdi','kaydetti','duyurdu','ifade etti','vurguladı']): score+=2
    if any(x in n for x in ['arttı','azaldı','geriledi','yükseldi','düştü','ulaştı','çıktı','indi','daraldı','büyüdü']): score+=3
    if any(x in n for x in ['üretim','ihracat','ithalat','istihdam','kapasite','yatırım','hasar','etkilendi','müşteri','tesis','fabrika']): score+=2
    if any(x in n for x in ['nedeni','sonucu','buna göre','bu kapsamda','öte yandan','ayrıca','son olarak']): score+=1
    return score

def _akt_formal_summary(title, body, max_sentences=10, max_chars=2800):
    """
    Haber başından sonuna okunur:
    - tekrar/menü temizlenir,
    - başlangıçtan ilk önemli bilgiler,
    - ortadaki en güçlü veri/açıklamalar,
    - sondaki sonuç/son durum birlikte seçilir,
    - orijinal haber sırası korunur.
    """
    sentences=_akt_clean_sentences(title,body)
    if not sentences:
        fallback=re.sub(r'\s+',' ',str(body or title or '')).strip()
        return fallback[:max_chars].rstrip(' .;')

    n=len(sentences)
    chosen=set(range(min(2,n)))  # başlangıç

    # son durum / sonuç
    for i in range(max(0,n-2),n):
        chosen.add(i)

    # gövdedeki en vurucu sayısal/kurumsal bilgiler
    ranked=sorted(
        [(i,_akt_sentence_score(s)) for i,s in enumerate(sentences)],
        key=lambda z:(z[1],-z[0]),
        reverse=True
    )
    for i,_ in ranked:
        if len(chosen)>=max_sentences:
            break
        chosen.add(i)

    ordered=[sentences[i] for i in sorted(chosen)]

    clauses=[]
    total=0
    for s in ordered:
        s=s.strip().rstrip(' .;:')
        if not s:
            continue
        if total+len(s)>max_chars and clauses:
            break
        clauses.append(s)
        total+=len(s)+2

    if not clauses:
        clauses=[sentences[0].strip().rstrip(' .;:')]

    # Örnekteki resmî AKT anlatımına yakın tek akış.
    text='; '.join(clauses)
    if text:
        first=text[0]
        if first.isalpha() and not text[:5].isupper():
            text=first.lower()+text[1:]
    return text

def _expanded_report_text(title, body):
    # Geriye dönük uyumluluk: AKT artık ham tam metni değil, resmî ve tekrarsız özeti kullanır.
    return _akt_formal_summary(title,body)


# -----------------------------
# V66 — KURUMSAL RESMÎ DİL NORMALİZASYONU
# -----------------------------
def _v66_formalize_sentence_endings(text):
    """
    V67: Önemli Gelişmeler ve Bilgi Notunda cümle sonlarındaki haber dili
    (-yor/-dı) yerine kurumsal resmî dil (-maktadır/-miştir) kullanılır.
    """
    t=re.sub(r'\s+',' ',str(text or '')).strip()
    if not t:
        return t

    exact=[
        ('açıklıyor','açıklamaktadır'),('belirtiyor','belirtmektedir'),
        ('bildiriyor','bildirmektedir'),('duyuruyor','duyurmaktadır'),
        ('söylüyor','söylemektedir'),('ifade ediyor','ifade etmektedir'),
        ('vurguluyor','vurgulamaktadır'),('gösteriyor','göstermektedir'),
        ('işaret ediyor','işaret etmektedir'),('ortaya koyuyor','ortaya koymaktadır'),
        ('öne çıkarıyor','öne çıkarmaktadır'),('öne çıkıyor','öne çıkmaktadır'),
        ('yer alıyor','yer almaktadır'),('devam ediyor','devam etmektedir'),
        ('sürüyor','sürmektedir'),('yürütülüyor','yürütülmektedir'),
        ('sürdürülüyor','sürdürülmektedir'),('yapılıyor','yapılmaktadır'),
        ('gerçekleştiriliyor','gerçekleştirilmektedir'),('kullanılıyor','kullanılmaktadır'),
        ('sayılıyor','sayılmaktadır'),('belirtiliyor','belirtilmektedir'),
        ('açıklanıyor','açıklanmaktadır'),('bildiriliyor','bildirilmektedir'),
        ('duyuruluyor','duyurulmaktadır'),('değerlendiriliyor','değerlendirilmektedir'),
        ('bekleniyor','beklenmektedir'),('planlanıyor','planlanmaktadır'),
        ('hedefleniyor','hedeflenmektedir'),('öngörülüyor','öngörülmektedir'),
        ('çalışılıyor','çalışılmaktadır'),('gerçekleşiyor','gerçekleşmektedir'),
        ('sağlıyor','sağlamaktadır'),('oluşturuyor','oluşturmaktadır'),
        ('taşıyor','taşımaktadır'),('sunuyor','sunmaktadır'),('koruyor','korumaktadır'),
        ('dolduruyor','doldurmaktadır'),('geçiyor','geçmektedir'),
        ('vuruyor','vurmaktadır'),('tamamlıyor','tamamlamaktadır'),
        ('artıyor','artmaktadır'),('azalıyor','azalmaktadır'),
    ]
    past=[
        ('yapıldı','yapılmıştır'),('gerçekleştirildi','gerçekleştirilmiştir'),
        ('açıklandı','açıklanmıştır'),('duyuruldu','duyurulmuştur'),
        ('yayımlandı','yayımlanmıştır'),('yayınlandı','yayımlanmıştır'),
        ('başladı','başlamıştır'),('tamamlandı','tamamlanmıştır'),
        ('sona erdi','sona ermiştir'),('arttı','artmıştır'),('azaldı','azalmıştır'),
        ('düştü','düşmüştür'),('yükseldi','yükselmiştir'),('geriledi','gerilemiştir'),
        ('ulaştı','ulaşmıştır'),('çıktı','çıkmıştır'),('geldi','gelmiştir'),
        ('verildi','verilmiştir'),('belirlendi','belirlenmiştir'),
        ('kaydedildi','kaydedilmiştir'),('tespit edildi','tespit edilmiştir'),
        ('bildirildi','bildirilmiştir'),('belirtildi','belirtilmiştir'),
        ('ifade edildi','ifade edilmiştir'),('vurgulandı','vurgulanmıştır'),
        ('kararlaştırıldı','kararlaştırılmıştır'),('onaylandı','onaylanmıştır'),
        ('imzalandı','imzalanmıştır'),('kuruldu','kurulmuştur'),
        ('devreye alındı','devreye alınmıştır'),('duyurdu','duyurmuştur'),
        ('açıkladı','açıklamıştır'),('belirtti','belirtmiştir'),
        ('bildirdi','bildirmiştir'),('gösterdi','göstermiştir'),
        ('sağladı','sağlamıştır'),('geçti','geçmiştir'),('vurdu','vurmuştur'),
    ]
    pairs=exact+past
    parts=re.split(r'(?<=[.!?])\s+',t)
    out=[]
    for s in parts:
        s=s.strip()
        if not s: continue
        punct=s[-1] if s[-1] in '.!?' else '.'
        core=s[:-1].rstrip() if s[-1] in '.!?' else s
        low=core.lower()
        for old,newv in sorted(pairs,key=lambda x:len(x[0]),reverse=True):
            if low.endswith(old):
                core=core[:-len(old)]+newv
                break
        out.append(core.rstrip(' .;:')+punct)
    return ' '.join(out)


def _v66_limit_important_paragraph(text,max_chars=520,max_sentences=3):
    """
    Önemli gelişmeler notunda her gelişmeyi Word üzerinde yaklaşık dört satırı
    aşmayacak yoğunlukta tutar. Öncelik ilk bilgi taşıyan cümlelere verilir.
    """
    clean=_v66_formalize_sentence_endings(text)
    sents=_sentence_chunks(clean)
    chosen=[]
    total=0
    for s in sents:
        s=s.strip()
        if not s: continue
        if total+len(s)>max_chars and chosen:
            break
        chosen.append(s)
        total+=len(s)+1
        if len(chosen)>=max_sentences:
            break
    result=' '.join(chosen).strip()
    if len(result)>max_chars:
        cut=result[:max_chars].rsplit(' ',1)[0].rstrip(' ,;:')
        # Kurumsal kapanış; kesilmiş yarım yüklem bırakma.
        if cut and cut[-1] not in '.!?':
            cut+='.'
        result=cut
    return result

def _akt_topic_labels(rows):
    joined=norm(' '.join(
        f"{r.get('Başlık','')} {r.get('İçerik_Özeti','')} {r.get('Kategori','')}"
        for r in rows
    ))
    mapping=[
        ('istihdam','istihdam'),
        ('sanayi üret','sanayi üretimi'),
        ('otomotiv','otomotiv üretimi'),
        ('yapay zeka','yapay zeka'),
        ('yapay zekâ','yapay zeka'),
        ('veri sızınt','veri sızıntısı'),
        ('siber saldır','siber güvenlik'),
        ('ihracat','ihracat'),
        ('yatırım','yatırım'),
        ('kapasite kullanım','kapasite kullanım oranı'),
        ('savunma','savunma sanayii'),
        ('enerji','enerji'),
        ('yangın','sanayi tesisi yangını'),
        ('patlama','sanayi tesisi patlaması'),
        ('ar-ge','Ar-Ge'),
        ('arge','Ar-Ge')
    ]
    out=[]
    for key,label in mapping:
        if key in joined and label not in out:
            out.append(label)
        if len(out)>=6:
            break
    return out

def _akt_findings_intro(rows):
    topics=_akt_topic_labels(rows)
    if topics:
        if len(topics)==1:
            topic_text=f'“{topics[0]}”'
        else:
            topic_text=', '.join(f'“{x}”' for x in topics[:-1]) + f' ve “{topics[-1]}”'
        return (
            "Sanayi ve Teknoloji alanlarında yapılan açık kaynak taraması neticesinde bazı haber "
            f"bültenlerinde {topic_text} konu başlıklarıyla ilgili içerikler hazırlandığı tespit edilmiştir. "
            "İçeriklerin hangi internet sitesinde yer aldığı, başlığı, bağlantı adresi, içeriğin detaylı özeti "
            "ve görseli aşağıda yer almaktadır."
        )
    return (
        "Sanayi ve Teknoloji alanlarında yapılan açık kaynak taraması neticesinde seçilen haber içerikleri "
        "tespit edilmiştir. İçeriklerin hangi internet sitesinde yer aldığı, başlığı, bağlantı adresi, "
        "içeriğin detaylı özeti ve görseli aşağıda yer almaktadır."
    )


def _v67_akt_reported_content(text):
    """
    AKT'de haber içeriğini dolaylı anlatı biçimine çevirir:
    açıklıyor -> açıkladığı, duyurdu -> duyurduğu, belirtiyor -> belirttiği vb.
    Son kapanış tek kez 'hususları ifade edilmektedir.' olur.
    """
    t=re.sub(r'\s+',' ',str(text or '')).strip().rstrip(' .;:')
    if not t: return t

    conv=[
        ('ifade ediyor','ifade ettiği'),('ifade etti','ifade ettiği'),
        ('açıklıyor','açıkladığı'),('açıkladı','açıkladığı'),
        ('belirtiyor','belirttiği'),('belirtti','belirttiği'),
        ('bildiriyor','bildirdiği'),('bildirdi','bildirdiği'),
        ('duyuruyor','duyurduğu'),('duyurdu','duyurduğu'),
        ('vurguluyor','vurguladığı'),('vurguladı','vurguladığı'),
        ('gösteriyor','gösterdiği'),('gösterdi','gösterdiği'),
        ('işaret ediyor','işaret ettiği'),('işaret etti','işaret ettiği'),
        ('ortaya koyuyor','ortaya koyduğu'),('ortaya koydu','ortaya koyduğu'),
        ('sağlıyor','sağladığı'),('sağladı','sağladığı'),
        ('dolduruyor','doldurduğu'),('doldurdu','doldurduğu'),
        ('yer alıyor','yer aldığı'),('yer aldı','yer aldığı'),
        ('devam ediyor','devam ettiği'),('devam etti','devam ettiği'),
        ('sürüyor','sürdüğü'),('sürdü','sürdüğü'),
        ('tamamladı','tamamladığı'),('tamamlıyor','tamamladığı'),
        ('vuruyor','vurduğu'),('vurdu','vurduğu'),
        ('geçiyor','geçtiği'),('geçti','geçtiği'),
        ('yapıldı','yapıldığı'),('gerçekleştirildi','gerçekleştirildiği'),
        ('açıklandı','açıklandığı'),('duyuruldu','duyurulduğu'),
        ('yayımlandı','yayımlandığı'),('başladı','başladığı'),
        ('tamamlandı','tamamlandığı'),('ulaştı','ulaştığı'),
        ('arttı','arttığı'),('azaldı','azaldığı'),
        ('oldu','olduğu'),('oluyor','olduğu'),
        ('sahiptir','sahip olduğu'),('dayanmaktadır','dayandığı'),
        ('değişebilir','değişebileceği'),
    ]

    clauses=[x.strip(' ,;:.') for x in re.split(r'\s*;\s*',t) if x.strip()]
    out=[]
    for c in clauses:
        low=c.lower()
        changed=False
        for old,newv in sorted(conv,key=lambda x:len(x[0]),reverse=True):
            # Haber özetindeki yüklem çoğunlukla cümlecik sonundadır.
            if low.endswith(old):
                c=c[:-len(old)]+newv
                changed=True
                break
        # Nokta ile birleşmiş kısa cümlelerde de son yüklemi dönüştür.
        if not changed:
            for old,newv in sorted(conv,key=lambda x:len(x[0]),reverse=True):
                c=re.sub(r'\b'+re.escape(old)+r'(?=\s*$)',newv,c,flags=re.I)
        out.append(c.rstrip(' .;:'))
    return '; '.join(out)

def make_docx(rows):
    """
    Kullanıcının ilettiği STB AKT örneğine yakın resmî format:
    Başlık -> görev alanı -> tarih -> bulgular -> numaralı haber/özet/link -> görsel -> Arz olunur.
    """
    doc=Document()
    section=doc.sections[0]
    section.top_margin=Cm(2.0)
    section.bottom_margin=Cm(2.0)
    section.left_margin=Cm(2.5)
    section.right_margin=Cm(2.5)

    normal=doc.styles["Normal"]
    normal.font.name="Times New Roman"
    normal.font.size=Pt(12)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"),"Times New Roman")

    p=doc.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after=Pt(10)
    r=p.add_run("AÇIK KAYNAK TARAMA ÇALIŞMASI")
    r.bold=True
    r.font.name="Times New Roman"
    r.font.size=Pt(14)

    p=doc.add_paragraph()
    p.paragraph_format.space_after=Pt(0)
    p.add_run("Tarama Yapılan Görev Alanı: ").bold=True
    p.add_run("Sanayi ve Teknoloji")

    p=doc.add_paragraph()
    p.paragraph_format.space_after=Pt(8)
    p.add_run("Tarih: ").bold=True
    p.add_run(datetime.now().astimezone().strftime("%d.%m.%Y"))

    p=doc.add_paragraph()
    p.paragraph_format.space_after=Pt(3)
    p.add_run("Bulgular: ").bold=True
    p.add_run(_akt_findings_intro(rows))

    for i,row in enumerate(rows,1):
        detail=article_detail(row)

        real_url=detail.get("canonical") or row.get("Yayıncı_URL") or row.get("URL","")
        title=(detail.get("title") or row.get("Başlık") or "").strip()
        source=_real_source(row,detail,real_url)
        body=detail.get("text") or row.get("İçerik_Özeti") or title
        summary=_v67_akt_reported_content(_akt_formal_summary(title,body))

        p=doc.add_paragraph()
        p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent=Cm(0.75)
        p.paragraph_format.space_before=Pt(4)
        p.paragraph_format.space_after=Pt(6)

        nr=p.add_run(f"{i}. ")
        nr.bold=True

        sr=p.add_run(f'“{source}”')
        sr.bold=True
        p.add_run(' isimli internet sitesinde, ')
        tr=p.add_run(f'“{title}”')
        tr.bold=True
        p.add_run(' başlığıyla bir haber yayımlanmıştır. (')
        _word_hyperlink(p,real_url,real_url if real_url else "Haber Linki")
        p.add_run(') Söz konusu haber içeriğinde, ')
        p.add_run(summary)
        p.add_run(' hususları ifade edilmiştir.')

        image_stream=None
        image_url=""
        for candidate in detail.get("images",[]):
            image_stream=_download_report_image(candidate)
            if image_stream:
                image_url=candidate
                break

        if image_stream or detail.get("images"):
            cap=doc.add_paragraph()
            cap.alignment=WD_ALIGN_PARAGRAPH.CENTER
            cap.paragraph_format.space_before=Pt(4)
            cap.paragraph_format.space_after=Pt(4)
            cr=cap.add_run(f'Görsel {i}: “{source}” Sitesinde Yer Alan Görsel')
            cr.bold=True
            cr.font.name="Times New Roman"
            cr.font.size=Pt(11)

        if image_stream:
            ip=doc.add_paragraph()
            ip.alignment=WD_ALIGN_PARAGRAPH.CENTER
            ip.paragraph_format.space_after=Pt(10)
            ip.add_run().add_picture(image_stream,width=Cm(14.5))
        elif detail.get("images"):
            # Örnekte görsel esas; indirilemediyse raporu gereksiz teknik metinle doldurma.
            lp=doc.add_paragraph()
            lp.alignment=WD_ALIGN_PARAGRAPH.CENTER
            lp.paragraph_format.space_after=Pt(8)
            _word_hyperlink(lp,detail["images"][0],"Görseli Aç")

    endp=doc.add_paragraph()
    endp.paragraph_format.space_before=Pt(8)
    endp.add_run("Arz olunur.")

    bio=BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()


# -----------------------------
# V63 — İŞ AKIŞI HAFIZASI / İKİNCİ GÖZ / YARINA TAKİP
# -----------------------------
def _v63_mark_notes(rows):
    if rows is None or len(rows)==0 or not _init_history_db():
        return
    try:
        with _history_connect() as conn:
            for row in rows:
                title=str(row.get('Başlık','') or '').strip()
                if not title: continue
                conn.execute(
                    "INSERT OR IGNORE INTO note_history(created_at,title,url) VALUES(?,?,?)",
                    (datetime.now().astimezone().isoformat(),title,str(row.get('URL','') or '').strip())
                )
            conn.commit()
        _v73_invalidate_status_cache()
    except Exception:
        pass

def _v73_invalidate_status_cache():
    st.session_state.pop('_v73_status_sets_cache',None)

def _v63_status_sets():
    """
    V102 — Tek sorguda dört işlem durumunu okur:
    Önemli Gelişmeler, AKT, Sunum ve hazırlanmış Bilgi Notu.
    """
    cached=st.session_state.get('_v73_status_sets_cache')
    # V102 geçiş güvenliği: V101 açık oturumlarında cache 3 elemanlıdır.
    # Yeni sürüm 4 durum kümesi kullandığı için eski cache'i otomatik geçersiz kıl.
    if cached is not None:
        try:
            if isinstance(cached,(tuple,list)) and len(cached)==4:
                return cached
        except Exception:
            pass
        st.session_state.pop('_v73_status_sets_cache',None)

    imp=set(); akt=set(); notes=set(); pres=set()
    if not _init_history_db():
        return imp,akt,notes,pres
    try:
        with _history_connect() as conn:
            for table,target in [
                ('important_basket',imp),
                ('osint_report_basket',akt),
                ('note_history',notes),
                ('presentation_basket',pres)
            ]:
                rows=conn.execute(f"SELECT title,url FROM {table}").fetchall()
                for title,url in rows:
                    target.add(str(url).strip() if str(url or '').strip() else title_key(str(title or '')))
    except Exception:
        pass
    result=(imp,akt,notes,pres)
    st.session_state['_v73_status_sets_cache']=result
    return result

def _v63_add_status_badges(df):
    """Her haber tablosuna tek bir Durum sütunu ekler."""
    if df is None or df.empty:
        return df
    out=df.copy()
    imp,akt,notes,pres=_v63_status_sets()

    def badge(r):
        # Ana tarama tabloları Türkçe; sepet tabloları İngilizce kolon adları kullanır.
        url=str(r.get('URL',r.get('url','')) or '').strip()
        title=str(r.get('Başlık',r.get('title','')) or '')
        k=url or title_key(title)
        b=[]
        if k in pres:  b.append('🖥️ Sunum Sepetinde')
        if k in imp:   b.append('📌 Önemli Gelişmelerde')
        if k in notes: b.append('📝 Bilgi Notu Yapıldı')
        if k in akt:   b.append('📁 AKT Sepetinde')
        return ' • '.join(b) if b else '—'

    out['Durum']=out.apply(badge,axis=1)
    return out

def _v63_missed_candidates(df,limit=12):
    """Yüksek değerli fakat iki sepette de olmayan olayları ikinci göz olarak gösterir."""
    if df is None or df.empty: return pd.DataFrame()
    value=_v52_event_value_table(df,max(30,limit*2))
    if value.empty: return value
    imp,akt,notes,pres=_v63_status_sets()
    rows=[]
    for _,v in value.iterrows():
        url=str(v.get('URL','') or '').strip()
        key=url or title_key(str(v.get('Gelişme','')))
        if key in imp or key in akt: continue
        # İkinci göz eşiği: güçlü değer skoru veya belirgin risk.
        if int(v.get('Değer_Skoru',0) or 0)<55 and int(v.get('Risk',0) or 0)<60:
            continue
        rows.append(v.to_dict())
        if len(rows)>=limit: break
    return pd.DataFrame(rows)

def _v63_tomorrow_candidates(df,limit=15):
    """Sonuçlanmamış, stratejik/riskli ve takip değeri olan olayları yarın için önerir."""
    if df is None or df.empty: return pd.DataFrame()
    life=_v58_event_lifecycle_table(df,40)
    if life.empty: return pd.DataFrame()
    out=life[life['Aşama']!='✅ Sonuçlandı'].copy()
    out=out[(pd.to_numeric(out['Risk_Skoru'],errors='coerce').fillna(0)>=35) |
            (pd.to_numeric(out['Kaynak_Sayısı'],errors='coerce').fillna(0)>=2)]
    if out.empty: return out
    out['Takip_Gerekçesi']=out.apply(
        lambda r:(
            'Olay gelişiyor; yeni açıklama/sonuç bekleniyor.'
            if 'Gelişiyor' in str(r.get('Aşama','')) else
            'Teyit edildi; uygulama/sonuç etkisi izlenmeli.'
            if 'Teyit' in str(r.get('Aşama','')) else
            'İlk sinyal; ikinci kaynak veya resmî teyit izlenmeli.'
        ),axis=1
    )
    return out.head(limit)

def _v63_add_tomorrow(rows):
    if rows is None or len(rows)==0 or not _init_history_db(): return 0
    added=0
    try:
        with _history_connect() as conn:
            for row in rows:
                title=str(row.get('Başlık','') or '').strip()
                if not title: continue
                cur=conn.execute("""
                    INSERT OR IGNORE INTO tomorrow_followup(
                        added_at,title,source,url,category,reason
                    ) VALUES(?,?,?,?,?,?)
                """,(
                    datetime.now().astimezone().isoformat(),title,
                    str(row.get('Kaynak','') or ''),str(row.get('URL','') or ''),
                    str(row.get('Kategori','') or ''),str(row.get('Takip_Gerekçesi','') or '')
                ))
                added+=int(bool(cur.rowcount))
            conn.commit()
    except Exception: pass
    return added

def _v63_load_tomorrow():
    if not _init_history_db(): return pd.DataFrame()
    try:
        with _history_connect() as conn:
            return pd.read_sql_query("SELECT * FROM tomorrow_followup ORDER BY added_at DESC",conn)
    except Exception:
        return pd.DataFrame()



# -----------------------------
# V68 — KONTROL MERKEZİ / SONRAKİ EN İYİ İŞLEM
# -----------------------------
def _v68_control_center(df,limit=8):
    """
    V114 Kontrol Merkezi:
    - 09:00–17:30 Bilgi Notu: veri/istatistik, resmî açıklama, ürün/teknoloji tanıtımı vb.
    - 09:00–17:30 AKT: negatif, eleştirel, yapısal eleştiri, propaganda/dezenformasyon niteliği taşıyan olumsuz içerikler.
    - 17:30 sonrası: yalnız kritik/acil gelişmeler.
    - Sunum: resmî veri/istatistik, resmî açıklama veya resmî teyitli bilgi.
    """
    cols=['Öncelik','Önerilen_İşlem','Tarih','Başlık','Neden','Durum','Değer_Skoru','Risk_Skoru','URL']
    if df is None or df.empty:
        return pd.DataFrame(columns=cols), 'Veri Yok', ''

    try:
        from zoneinfo import ZoneInfo
        now_tr=datetime.now(ZoneInfo('Europe/Istanbul'))
    except Exception:
        now_tr=datetime.now().astimezone()
    hour=now_tr.hour + now_tr.minute/60

    if 9 <= hour < 14:
        phase='09:00–14:00 | Bilgi notu • AKT hazırlığı • sunum/veri kontrolü'
        phase_hint='Bilgi notunda resmî/veri odaklı içerikler; AKT’de negatif-eleştirel içerikler; sunumda ise resmî ve teyitli bilgiler önceliklendirilmektedir.'
    elif 14 <= hour < 17.5:
        phase='14:00–17:30 | Bilgi notu • sunum • önemli gelişmeleri zenginleştirme'
        phase_hint='Yeni resmî veri/açıklamalar bilgi notu ve sunum için; negatif-eleştirel içerikler AKT takibi için değerlendirilmektedir.'
    else:
        phase='17:30 sonrası | Kritik takip modu'
        phase_hint='Rutin bilgi notu ve sunum önerileri durdurulmakta; yalnızca kritik/acil gelişmeler öne çıkarılmaktadır.'

    value=_v52_event_value_table(df,max(40,limit*5))
    if value.empty:
        return pd.DataFrame(columns=cols),phase,phase_hint

    imp,akt,notes,pres=_v63_status_sets()
    actions=[]

    data_terms=[
        'istatistik','veri','oran','endeks','sanayi üretimi','kapasite kullanım',
        'ihracat','ithalat','ciro','istihdam','işsizlik','büyüme','yatırım teşvik',
        'arge','ar-ge','patent','başvuru','milyar','milyon','yüzde','%'
    ]
    product_terms=[
        'ürün tanıt','tanıtıldı','tanıttı','yeni ürün','yeni teknoloji','prototip',
        'seri üretim','ilk teslimat','envantere','platform','sistem geliştir',
        'füze','uydu','çip','yarı iletken','yapay zeka','yapay zekâ'
    ]
    propaganda_terms=[
        'propaganda','dezenformasyon','manipülasyon','iddia','suçlama','eleştiri',
        'eleştirel','tepki','kriz','başarısız','skandal','zarar','kayıp','çöküş',
        'iflas','işten çıkar','üretim durdu','üretimi durdur','gecikme','yaptırım',
        'ambargo','boykot','bağımlılık','risk','tehdit'
    ]

    for _,v in value.iterrows():
        row=_v53_find_event_row(df,v)
        if row is None:
            continue

        title=str(row.get('Başlık','') or v.get('Gelişme',''))
        url=str(row.get('URL','') or v.get('URL','')).strip()
        key=url or title_key(title)
        score=int(v.get('Değer_Skoru',0) or 0)
        risk=int(row.get('Risk_Skoru',v.get('Risk',0)) or 0)
        text=norm(f"{title} {row.get('İçerik_Özeti','')} {row.get('Kategori','')} {row.get('Doğrulama','')}")
        critical=bool(critical_industrial_incident(title,row.get('İçerik_Özeti','')))
        official=_is_official_radar_row(row)
        verification=norm(row.get('Doğrulama',''))
        officially_verified=official or any(x in verification for x in ['resmi','resmî','birincil','teyit'])
        negative=(str(row.get('Duygu',''))=='Negatif' or
                  str(row.get('Risk_Durumu',''))=='Yüksek Risk' or
                  any(x in text for x in propaganda_terms))
        data_stat=any(x in text for x in data_terms)
        product_intro=any(x in text for x in product_terms)
        multi=int(v.get('Kaynak_Sayısı',0) or 0)>=2

        badges=[]
        if key in imp: badges.append('📌 Önemli Gelişmelerde')
        if key in akt: badges.append('📁 AKT’de')
        if key in notes: badges.append('📝 Bilgi Notu Hazırlandı')
        status=' • '.join(badges) if badges else 'Henüz işleme alınmadı'

        proposals=[]

        # 17:30 sonrası: yalnız kritik gelişme.
        if hour >= 17.5 or hour < 9:
            if critical or risk>=75 or score>=88:
                why=[]
                if critical: why.append('kritik sanayi olayı')
                if risk>=75: why.append('çok yüksek risk')
                if score>=88: why.append('çok yüksek analitik değer')
                if officially_verified: why.append('resmî/teyitli bilgi')
                proposals.append((120+risk,'🚨 KRİTİK GELİŞME — ACİL DEĞERLENDİR',why))
        else:
            # Bilgi Notu: veri/istatistik, resmî açıklama, ürün/teknoloji tanıtımı.
            if key not in notes and (data_stat or official or product_intro):
                why=[]
                if data_stat: why.append('veri/istatistiki bilgi')
                if official: why.append('resmî açıklama/birincil kaynak')
                if product_intro: why.append('ürün/teknoloji tanıtımı veya somut teknolojik gelişme')
                if multi: why.append(f"{int(v.get('Kaynak_Sayısı',0) or 0)} farklı kaynak")
                proposals.append((105+score,'📝 Bilgi Notu Değerlendir',why))

            # AKT: negatif, eleştirel, propaganda/dezenformasyon/olumsuz içerik.
            if key not in akt and negative:
                why=[]
                if str(row.get('Duygu',''))=='Negatif': why.append('negatif/olumsuz içerik')
                if any(x in text for x in ['eleştiri','eleştirel','tepki','suçlama']): why.append('eleştirel dil/yapısal eleştiri')
                if any(x in text for x in ['propaganda','dezenformasyon','manipülasyon','iddia']): why.append('propaganda/manipülasyon iddiası veya niteliği')
                if risk>=55: why.append('dikkat gerektiren risk/etki')
                proposals.append((100+score+risk//5,'📁 AKT Sepetine Almayı Değerlendir',why))

            # Sunum: yalnız resmî veri/istatistik veya resmî/teyitli bilgi.
            if (data_stat and officially_verified) or official or (officially_verified and score>=55):
                why=[]
                if data_stat: why.append('resmî/teyitli veri veya istatistik')
                if official: why.append('resmî açıklama')
                elif officially_verified: why.append('resmî teyitli bilgi')
                proposals.append((85+score,'🖥️ Sunuma Eklemeyi Değerlendir',why))

            # Mevcut önemli gelişme yeni kaynaklarla zenginleşmişse ayrıca hatırlat.
            if key in imp and multi:
                proposals.append((78+score,'🔄 Önemli Gelişmeyi Zenginleştir',
                                  ['önemli gelişme sepetinde','yeni/çoklu kaynak desteği mevcut']))

        for priority,action,reason in proposals:
            actions.append({
                'Öncelik':priority,
                'Önerilen_İşlem':action,
                'Tarih':row.get('Tarih',''),
                'Başlık':title,
                'Neden':' • '.join(dict.fromkeys(reason)) if reason else 'analist değerlendirmesi önerilmektedir',
                'Durum':status,
                'Değer_Skoru':score,
                'Risk_Skoru':risk,
                'URL':url
            })

    if not actions:
        return pd.DataFrame(columns=cols),phase,phase_hint

    out=pd.DataFrame(actions)
    # Aynı haber aynı işlem için yalnız bir kez gösterilsin.
    out=out.sort_values(['Öncelik','Değer_Skoru','Risk_Skoru'],ascending=[False,False,False])
    out=out.drop_duplicates(subset=['Önerilen_İşlem','URL','Başlık'],keep='first').head(limit).reset_index(drop=True)
    out['Öncelik']=range(1,len(out)+1)
    return out[cols],phase,phase_hint


def _v73_row_keys(df):
    """apply(axis=1) yerine hızlı, vektörize haber anahtarı üretir."""
    if df is None or df.empty:
        return pd.Series(dtype=str)
    urls=df['URL'].fillna('').astype(str).str.strip() if 'URL' in df.columns else pd.Series('',index=df.index)
    titles=df['Başlık'].fillna('').astype(str) if 'Başlık' in df.columns else pd.Series('',index=df.index)
    # title_key yalnız URL'siz satırlarda çalışır.
    fallback=titles.map(title_key)
    return urls.where(urls.ne(''),fallback)



# ============================================================
# V107 — OLAY BAZLI KAYNAK ZENGİNLEŞTİRME
# Kullanıcı yerel/kısa bir haberi seçse bile aynı taramadaki aynı olayın
# ana akım/resmî/daha ayrıntılı sürümleri birleştirilerek sepete aktarılır.
# Ek web isteği YOKTUR; yalnız mevcut tarama havuzu kullanılır.
# ============================================================

def _v107_source_quality(row):
    """Bir olay kümesinde hangi haber sürümünün ana taşıyıcı olacağını puanlar."""
    d=str(row.get('Domain','') or '')
    source=str(row.get('Kaynak','') or '')
    text=_clean_note_text(row.get('İçerik_Özeti',''))
    score=0
    if d in TR_OFFICIAL:
        score+=80
    elif d in TR_MAIN:
        score+=55
    elif d in TR_TECH:
        score+=45
    elif d in SOCIAL:
        score+=5
    else:
        score+=20
    # Ayrıntılı gövdeyi ödüllendir; aşırı uzun portal artıklarına sınırsız puan verme.
    score+=min(len(text)//180,30)
    score+=min(len(re.findall(r'\b\d+(?:[.,]\d+)?\b',text))*2,16)
    if _verification_rank(row.get('Doğrulama',''))>=3:
        score+=12
    if source and norm(source) not in {'google','google news','google haberler','açık kaynak'}:
        score+=4
    return score

def _v107_same_event(selected, candidate):
    """V104 olay mantığını kullanarak seçili haberle gerçekten aynı olayı sınar."""
    su=str(selected.get('URL','') or '').strip()
    cu=str(candidate.get('URL','') or '').strip()
    if su and cu and su==cu:
        return True
    soid=str(selected.get('Olay_ID','') or '').strip()
    coid=str(candidate.get('Olay_ID','') or '').strip()
    if soid and coid and soid==coid:
        return True
    sim=_v104_event_similarity(
        selected.get('Başlık',''),selected.get('İçerik_Özeti',''),
        candidate.get('Başlık',''),candidate.get('İçerik_Özeti','')
    )
    return sim>=0.50

def _v107_unique_sentences(rows, max_chars=8000):
    """
    En iyi kaynaklardan gelen tamamlayıcı cümleleri birleştirir.
    Aynı cümleyi/çok benzer bilgiyi tekrar eklemez; kritik rakamları korur.
    """
    out=[]; seen=set()
    for row in rows:
        raw=_v84_hard_repair_text(row.get('İçerik_Özeti',''))
        sentences=_v84_clean_article_sentences(raw)
        if not sentences and raw:
            sentences=[raw]
        for s in sentences:
            s=_clean_note_text(s).strip()
            if len(s)<25:
                continue
            k=title_key(s)
            toks=set(_history_tokens(s))
            duplicate=False
            for oldk,oldtoks in seen:
                if k==oldk:
                    duplicate=True; break
                oldtoks=set(oldtoks)
                if toks and oldtoks:
                    jac=len(toks&oldtoks)/max(1,len(toks|oldtoks))
                    if jac>=0.72:
                        duplicate=True; break
            if duplicate:
                continue
            out.append(s)
            # V108 düzeltmesi: set nesnesi hashlenemez; frozenset olarak saklanır.
            seen.add((k,frozenset(toks)))
            if len(' '.join(out))>=max_chars:
                break
        if len(' '.join(out))>=max_chars:
            break
    return ' '.join(out)[:max_chars]

def _v107_enrich_selected_rows(rows):
    """
    Seçilen her haberi, st.session_state['rows'] içindeki aynı olayın diğer
    sürümleriyle zenginleştirir. Bir olaydan sepete yalnız bir kayıt gider.
    """
    if rows is None or len(rows)==0:
        return []
    selected=[dict(r) for r in rows]
    pool=st.session_state.get('rows') or []
    if not pool:
        return selected

    enriched=[]
    used_event_sigs=set()

    for sel in selected:
        matches=[]
        for cand in pool:
            try:
                if _v107_same_event(sel,cand):
                    matches.append(dict(cand))
            except Exception:
                continue
        if not matches:
            matches=[sel]

        # Yanlış geniş kümeyi engelle: en fazla en güçlü 8 kaynak.
        matches=sorted(matches,key=_v107_source_quality,reverse=True)[:8]
        best=matches[0].copy()

        # Seçilen kaydın işlem/risk bağlamını kaybetme.
        best['Risk_Skoru']=max([int(x.get('Risk_Skoru',0) or 0) for x in matches] or [int(sel.get('Risk_Skoru',0) or 0)])
        if sel.get('Risk_Durumu'):
            best['Risk_Durumu']=sel.get('Risk_Durumu')
        if sel.get('Kategori'):
            best['Kategori']=sel.get('Kategori')

        merged=_v107_unique_sentences(matches)
        if merged:
            best['İçerik_Özeti']=merged

        domains=[]
        sources=[]
        urls=[]
        for x in matches:
            d=str(x.get('Domain','') or '').strip()
            s=_clean_note_text(x.get('Kaynak','')).strip()
            u=str(x.get('URL','') or '').strip()
            if d and d not in domains: domains.append(d)
            if s and s not in sources: sources.append(s)
            if u and u not in urls: urls.append(u)

        best['Olay_Kaynak_Sayisi']=max(len(domains),len(sources),int(best.get('Olay_Kaynak_Sayisi',0) or 0))
        best['Zenginleştirme_Kaynakları']=' | '.join(sources[:8])
        best['Zenginleştirme_URLleri']=' | '.join(urls[:8])
        best['Zenginleştirildi']='Evet' if len(matches)>1 else 'Hayır'

        # Aynı olay kullanıcı tarafından iki farklı satırdan seçildiyse sepete iki kez gitmesin.
        sig=' '.join(sorted(_v104_event_tokens(best.get('Başlık',''),best.get('İçerik_Özeti',''))))
        if sig and sig in used_event_sigs:
            continue
        if sig: used_event_sigs.add(sig)
        enriched.append(best)

    return enriched

def _v74_bulk_add_basket(rows,table_name):
    """
    V74: Kronoloji hızlı işlemleri için tek SQLite executemany çağrısı.
    Satır satır execute yerine toplu INSERT OR IGNORE kullanır.
    """
    if rows is None or len(rows)==0 or not _init_history_db():
        return 0
    if table_name not in ('important_basket','osint_report_basket'):
        return 0
    payload=[]
    now_iso=datetime.now().astimezone().isoformat()
    for row in rows:
        title=str(row.get('Başlık','') or '').strip()
        if not title:
            continue
        payload.append((
            now_iso,
            str(row.get('Tarih','') or ''),
            title,
            str(row.get('Kaynak','') or ''),
            str(row.get('URL','') or '').strip(),
            str(row.get('Kategori','') or ''),
            int(row.get('Risk_Skoru',0) or 0),
            str(row.get('Risk_Durumu','') or ''),
            str(row.get('İçerik_Özeti','') or '')[:8000]
        ))
    if not payload:
        return 0
    try:
        with _history_connect() as conn:
            before=conn.total_changes
            conn.executemany(f"""
                INSERT OR IGNORE INTO {table_name}(
                    added_at,news_time,title,source,url,category,risk_score,risk_status,summary
                ) VALUES(?,?,?,?,?,?,?,?,?)
            """,payload)
            conn.commit()
            added=conn.total_changes-before
        if added:
            _v73_invalidate_status_cache()
        return int(added)
    except Exception:
        return 0

def _v74_fast_add_important(rows):
    return _v74_bulk_add_basket(_v107_enrich_selected_rows(rows),'important_basket')

def _v74_fast_add_osint(rows):
    return _v74_bulk_add_basket(_v107_enrich_selected_rows(rows),'osint_report_basket')

def _v80_add_presentation(rows):
    """Her bölümden seçilen haberleri sunum sepetine toplu ekler."""
    if rows is None or len(rows)==0 or not _init_history_db():
        return 0
    payload=[]
    now_iso=datetime.now().astimezone().isoformat()
    for row in rows:
        title=_clean_note_text(row.get('Başlık',''))
        if not title:
            continue
        payload.append((
            now_iso,
            str(row.get('Tarih','') or ''),
            title,
            _clean_note_text(row.get('Kaynak','')),
            str(row.get('URL','') or '').strip(),
            _clean_note_text(row.get('Kategori','')),
            _clean_note_text(row.get('İçerik_Özeti',''))[:5000]
        ))
    try:
        with _history_connect() as conn:
            before=conn.total_changes
            conn.executemany("""
                INSERT OR IGNORE INTO presentation_basket(
                    added_at,news_time,title,source,url,category,summary
                ) VALUES(?,?,?,?,?,?,?)
            """,payload)
            conn.commit()
            added=int(conn.total_changes-before)
        if added:
            _v73_invalidate_status_cache()
        return added
    except Exception:
        return 0

def _v80_load_presentation():
    if not _init_history_db():
        return pd.DataFrame()
    try:
        with _history_connect() as conn:
            return pd.read_sql_query(
                "SELECT * FROM presentation_basket ORDER BY id DESC",conn
            )
    except Exception:
        return pd.DataFrame()

def _v80_clear_presentation():
    try:
        with _history_connect() as conn:
            cur=conn.execute("DELETE FROM presentation_basket")
            conn.commit()
            removed=cur.rowcount
        if removed:
            _v73_invalidate_status_cache()
        return removed
    except Exception:
        return 0

def _v81_remove_presentation_ids(ids):
    ids=[int(x) for x in ids if str(x).isdigit()]
    if not ids: return 0
    try:
        with _history_connect() as conn:
            marks=','.join('?' for _ in ids)
            cur=conn.execute(f"DELETE FROM presentation_basket WHERE id IN ({marks})",ids)
            conn.commit()
            removed=cur.rowcount
        if removed:
            _v73_invalidate_status_cache()
        return removed
    except Exception:
        return 0

def _v81_basket_to_rows(bdf):
    rows=[]
    if bdf is None or bdf.empty: return rows
    for _,r in bdf.iterrows():
        rows.append({'Tarih':_clean_note_text(r.get('news_time','')),'Kaynak':_clean_note_text(r.get('source','')),
        'Başlık':_clean_note_text(r.get('title','')),'İçerik_Özeti':_clean_note_text(r.get('summary','')),
        'URL':str(r.get('url','') or ''),'Kategori':_clean_note_text(r.get('category','')),
        'Risk_Skoru':r.get('risk_score',0),'Risk_Durumu':_clean_note_text(r.get('risk_status',''))})
    return rows


def _v73_main_selected(selected_keys):
    """
    Ana tarama DataFrame'ini yalnız kullanıcı gerçekten bir işlem butonuna bastığında oluşturur/eşleştirir.
    Checkbox işaretlemek artık yüzlerce satır üzerinde gereksiz tekrar filtrelemesi başlatmaz.
    """
    if not selected_keys:
        return pd.DataFrame()
    main_rows=st.session_state.get('rows') or []
    if not main_rows:
        return pd.DataFrame()
    main_df=pd.DataFrame(main_rows)
    keys=_v73_row_keys(main_df)
    return main_df[keys.isin(selected_keys)].copy()

def _section_select_table(section_key, data, columns, height=420):
    """
    V75 ULTRA HIZ:
    Tüm bölüm tablolarında checkbox değişikliği form içinde kalır.
    Streamlit yalnız kullanıcı işlem düğmesine bastığında rerun yapar.
    """
    if data is None or data.empty:
        return pd.DataFrame()

    tbl=_v63_add_status_badges(data.copy())
    if 'Durum' not in columns:
        columns=list(columns)
        insert_at=columns.index('Başlık')+1 if 'Başlık' in columns else 0
        columns.insert(insert_at,'Durum')

    tbl['_row_key']=_v73_row_keys(tbl).values
    selected_map=st.session_state.section_selections.get(section_key,{})
    tbl.insert(0,'Seç',[bool(selected_map.get(k,False)) for k in tbl['_row_key']]) if 'Seç' not in tbl.columns else None
    if 'Seç' in tbl.columns:
        tbl['Seç']=[bool(selected_map.get(k,bool(v))) for k,v in zip(tbl['_row_key'],tbl['Seç'].tolist())]

    show_cols=['Seç']+[c for c in columns if c in tbl.columns and c!='Seç']

    with st.form(key=f'v75_fast_section_form_{section_key}',clear_on_submit=False):
        edited=st.data_editor(
            tbl[show_cols+['_row_key']],
            column_config={
                'Seç':st.column_config.CheckboxColumn('Seç'),
                'URL':st.column_config.LinkColumn('Haber Linki'),
                'Medya_URL':st.column_config.LinkColumn('Medya'),
                'Resmî_URL':st.column_config.LinkColumn('Resmî Açıklama'),
                'Eşleşme':st.column_config.NumberColumn('Eşleşme',format='%d%%'),
                'Değer_Skoru':st.column_config.ProgressColumn('Değer Skoru',min_value=0,max_value=100,format='%d/100'),
                'Risk':st.column_config.NumberColumn('Risk',format='%d/100'),
                'Risk_Skoru':st.column_config.NumberColumn('Risk',format='%d/100'),
                'Durum':st.column_config.TextColumn('Durum',width='large'),
                '_row_key':None
            },
            disabled=[c for c in show_cols if c!='Seç']+['_row_key'],
            hide_index=True,use_container_width=True,height=height,
            key=f'v75_section_editor_{section_key}'
        )
        a1,a2,a3,a4=st.columns(4)
        with a1: do_imp=st.form_submit_button('📌 Önemli Gelişmelere Ekle',use_container_width=True)
        with a2: do_akt=st.form_submit_button('🗂️ AKT Sepetine Ekle',use_container_width=True)
        with a3: do_pres=st.form_submit_button('🖥️ Sunum Sepetine Ekle',use_container_width=True)
        with a4: do_note=st.form_submit_button('📝 Bilgi Notu Oluştur',use_container_width=True)

    selected_keys=set(edited.loc[edited['Seç'].astype(bool),'_row_key'].astype(str))
    st.session_state.section_selections[section_key]={k:(k in selected_keys) for k in edited['_row_key'].astype(str)}
    selected=data[_v73_row_keys(data).isin(selected_keys)].copy()

    if do_imp or do_akt or do_note or do_pres:
        if not selected_keys:
            st.warning('Önce en az bir haberi işaretleyin.')
        else:
            # Önce görünür bölüm verisini kullan: ana dataframe eşleştirmesine çoğu işlemde gerek yok.
            action_rows=selected.copy()
            if do_imp:
                n=_v74_fast_add_important(action_rows.to_dict('records'))
                st.success(f'✅ {n} yeni haber Önemli Gelişmeler Sepeti’ne eklenmiştir.')
            elif do_akt:
                n=_v74_fast_add_osint(action_rows.to_dict('records'))
                st.success(f'✅ {n} yeni haber AKT Sepeti’ne eklenmiştir.')
            elif do_pres:
                n=_v80_add_presentation(action_rows.to_dict('records'))
                st.success(f'✅ {n} yeni haber Sunum Sepeti’ne eklenmiştir.')
            elif do_note:
                # Bilgi notunda tam içerik gerekiyorsa yalnız burada ana tabloya dön.
                full=_v73_main_selected(selected_keys)
                if full.empty: full=action_rows
                with st.spinner(f'{len(full)} seçili haber için bilgi notu hazırlanmaktadır...'):
                    try:
                        st.session_state[f'section_note_bytes_{section_key}']=make_analyst_docx(
                            full,title='SANAYİ & TEKNOLOJİ BİLGİ NOTU'
                        )
                        _v63_mark_notes(full.to_dict('records'))
                        _v73_invalidate_status_cache()
                    except Exception as e:
                        st.session_state[f'section_note_bytes_{section_key}']=None
                        st.error(f'Bilgi notu hazırlanamadı: {e}')

    section_note_bytes=st.session_state.get(f'section_note_bytes_{section_key}')
    if section_note_bytes:
        st.download_button(
            '⬇️ Hazırlanan Bilgi Notunu İndir',
            data=section_note_bytes,
            file_name=f'Sanayi_Teknoloji_Bilgi_Notu_{section_key}_{date.today()}.docx',
            mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            use_container_width=True,
            key=f'v75_note_download_{section_key}'
        )
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
    mask=_v73_row_keys(df).isin(keys)
    return df[mask].copy()


# -----------------------------
# V60 — OTOMATİK GERİ DÖNÜŞ / ANOMALİ / GÜN SONU
# -----------------------------
def _v60_register_visit_once():
    """
    Yeni browser oturumunda bir kez çalışır.
    Önceki giriş zamanını alır, mevcut girişi kaydeder.
    Streamlit rerun'larında baseline değişmez.
    """
    if st.session_state.get('_v60_visit_initialized',False):
        return st.session_state.get('_v60_previous_visit')

    previous=None
    now=datetime.now().astimezone()
    if _init_history_db():
        try:
            with _history_connect() as conn:
                row=conn.execute(
                    "SELECT visited_at FROM app_visits ORDER BY visited_at DESC LIMIT 1"
                ).fetchone()
                if row:
                    previous=pd.to_datetime(row[0],utc=True,errors='coerce')
                conn.execute(
                    "INSERT INTO app_visits(visited_at) VALUES(?)",
                    (now.isoformat(),)
                )
                conn.commit()
        except Exception:
            previous=None

    st.session_state['_v60_visit_initialized']=True
    st.session_state['_v60_previous_visit']=previous
    st.session_state['_v60_this_visit']=now
    return previous

def _v60_auto_catchup(previous_visit,user_query=''):
    """
    Kullanıcı yeniden giriş yaptığında manuel buton gerektirmeden,
    son girişten bu yana gelişmeleri hafif bir sorgu setiyle kontrol eder.
    Tam tarama değildir; yalnızca dönüş brifingi içindir.
    """
    if previous_visit is None or pd.isna(previous_visit):
        return [],None

    now_utc=datetime.now(timezone.utc)
    prev_utc=previous_visit.to_pydatetime() if hasattr(previous_visit,'to_pydatetime') else previous_visit
    if prev_utc.tzinfo is None:
        prev_utc=prev_utc.replace(tzinfo=timezone.utc)
    else:
        prev_utc=prev_utc.astimezone(timezone.utc)

    delta_h=max(0.25,(now_utc-prev_utc).total_seconds()/3600)
    # Google/RSS tarafında geniş pencere kullanılır; kesin filtre aşağıda previous_visit ile yapılır.
    when=period_window(max(3,delta_h))

    queries=[
        f'Türkiye (sanayi OR teknoloji OR üretim OR fabrika OR tesis OR yatırım OR OSB) when:{when}',
        f'Türkiye (savunma OR ASELSAN OR TUSAŞ OR ROKETSAN OR HAVELSAN OR Baykar OR otomotiv OR TOGG) when:{when}',
        f'Türkiye ("yapay zeka" OR "yarı iletken" OR çip OR siber OR Ar-Ge OR TÜBİTAK OR KOSGEB) when:{when}',
    ]
    queries += build_negative_queries(when)
    queries += build_official_radar_queries(when)

    raw=[]
    with concurrent.futures.ThreadPoolExecutor(max_workers=min(7,len(queries))) as ex:
        futs=[ex.submit(rss,q) for q in queries]
        for f in concurrent.futures.as_completed(futs):
            try:
                raw.extend(f.result() or [])
            except Exception:
                pass

    rows,_=normalize_rows(raw,prev_utc,'turkish',user_query)
    rows=dedupe(rows)
    if rows:
        rows=enrich_rows(rows)
    return rows,delta_h

def _v60_now_to_know_table(rows,n=5):
    if not rows:
        return pd.DataFrame()
    df=pd.DataFrame(rows)
    if df.empty:
        return df
    value=_v52_event_value_table(df,max(n,10))
    if value.empty:
        return value
    return value.head(n).copy()

def _v60_anomaly_radar(df,current_hours,lookback_days=14):
    """
    Mevcut taramadaki kategori olay hızını, geçmiş günlerin son taramalarındaki
    saatlik olay hızıyla karşılaştırır. Ek web isteği yoktur.
    """
    cols=['Kategori','Şimdi','Beklenen','Normalin_Katı','Durum']
    if df is None or df.empty or not _init_history_db():
        return pd.DataFrame(columns=cols)

    try:
        cutoff=(datetime.now().astimezone()-timedelta(days=lookback_days)).isoformat()
        with _history_connect() as conn:
            hist=pd.read_sql_query("""
                SELECT s.scan_id,s.scanned_at,s.period_hours,e.category
                FROM scans s
                JOIN event_snapshots e ON e.scan_id=s.scan_id
                WHERE s.scanned_at>=?
                ORDER BY s.scanned_at DESC
            """,conn,params=(cutoff,))
    except Exception:
        return pd.DataFrame(columns=cols)

    if hist.empty:
        return pd.DataFrame(columns=cols)

    hist['day']=hist['scanned_at'].astype(str).str.slice(0,10)
    # Aynı gün çok tarama varsa yalnız o günün en son taraması baseline olur.
    last_scan_per_day=(
        hist[['day','scan_id','scanned_at']]
        .drop_duplicates()
        .sort_values('scanned_at')
        .groupby('day',as_index=False)
        .tail(1)[['day','scan_id']]
    )
    hist=hist.merge(last_scan_per_day,on=['day','scan_id'],how='inner')
    if hist['day'].nunique()<2:
        return pd.DataFrame(columns=cols)

    scan_hours=hist[['scan_id','period_hours']].drop_duplicates().set_index('scan_id')['period_hours'].to_dict()
    hc=hist.groupby(['scan_id','category']).size().reset_index(name='events')
    hc['rate']=hc.apply(
        lambda r:r['events']/max(1,float(scan_hours.get(r['scan_id'],24) or 24)),axis=1
    )
    baseline=hc.groupby('category')['rate'].agg(['mean','std','count']).reset_index()

    cur=df.groupby('Kategori')['Olay_ID'].nunique() if 'Olay_ID' in df.columns else df.groupby('Kategori').size()
    rows=[]
    for cat,current in cur.items():
        b=baseline[baseline['category']==cat]
        if b.empty:
            continue
        mean_rate=float(b.iloc[0]['mean'] or 0)
        expected=max(0.1,mean_rate*max(1,float(current_hours)))
        ratio=float(current)/expected if expected else 0
        # Hem göreli hem mutlak fark arıyoruz; küçük bazlarda sahte alarmı azaltır.
        if current>=3 and ratio>=1.8 and (current-expected)>=2:
            level='🔴 Çok Olağandışı' if ratio>=3 else '🟠 Olağandışı'
            rows.append({
                'Kategori':cat,
                'Şimdi':int(current),
                'Beklenen':round(expected,1),
                'Normalin_Katı':round(ratio,1),
                'Durum':level
            })

    if not rows:
        return pd.DataFrame(columns=cols)
    return pd.DataFrame(rows).sort_values(['Normalin_Katı','Şimdi'],ascending=[False,False])

def _v60_day_end_performance(df=None):
    """Bugünün operasyonel üretimini yerel geçmiş/sepet kayıtlarından özetler."""
    today=datetime.now().astimezone().date().isoformat()
    result={
        'Taramalar':0,'Benzersiz Olay':0,'Negatif':0,'Yüksek Risk':0,
        'Önemli Sepete Eklenen':0,'AKT Sepete Eklenen':0,'Kritik Sanayi':0
    }
    if _init_history_db():
        try:
            with _history_connect() as conn:
                result['Taramalar']=int(conn.execute(
                    "SELECT COUNT(*) FROM scans WHERE substr(scanned_at,1,10)=?",(today,)
                ).fetchone()[0] or 0)

                q="""SELECT COUNT(DISTINCT e.title)
                     FROM event_snapshots e JOIN scans s ON e.scan_id=s.scan_id
                     WHERE substr(s.scanned_at,1,10)=?"""
                result['Benzersiz Olay']=int(conn.execute(q,(today,)).fetchone()[0] or 0)

                qn="""SELECT COUNT(DISTINCT e.title)
                      FROM event_snapshots e JOIN scans s ON e.scan_id=s.scan_id
                      WHERE substr(s.scanned_at,1,10)=? AND e.sentiment='Negatif'"""
                result['Negatif']=int(conn.execute(qn,(today,)).fetchone()[0] or 0)

                qh="""SELECT COUNT(DISTINCT e.title)
                      FROM event_snapshots e JOIN scans s ON e.scan_id=s.scan_id
                      WHERE substr(s.scanned_at,1,10)=? AND e.risk_status='Yüksek Risk'"""
                result['Yüksek Risk']=int(conn.execute(qh,(today,)).fetchone()[0] or 0)

                result['Önemli Sepete Eklenen']=int(conn.execute(
                    "SELECT COUNT(*) FROM important_basket WHERE substr(added_at,1,10)=?",(today,)
                ).fetchone()[0] or 0)
                result['AKT Sepete Eklenen']=int(conn.execute(
                    "SELECT COUNT(*) FROM osint_report_basket WHERE substr(added_at,1,10)=?",(today,)
                ).fetchone()[0] or 0)
        except Exception:
            pass

    if df is not None and not df.empty:
        try:
            result['Kritik Sanayi']=int(df.apply(
                lambda r:bool(critical_industrial_incident(r.get('Başlık',''),r.get('İçerik_Özeti',''))),
                axis=1
            ).sum())
        except Exception:
            pass
    return result


# ============================================================
# V110 — V109 PANEL HATA DÜZELTMESİ
# NOT: Bu override'lar bütün eski fonksiyon tanımlarından SONRA,
# UI başlamadan hemen önce tanımlanır. Böylece eski V33/V51 fonksiyonları
# yeni davranışı ezemez.
# ============================================================

def _official_radar_rows(df):
    """V110 — Resmî Kaynak Radarı her zaman 'Kurum Türü' sütununu üretir."""
    if df is None or df.empty:
        return pd.DataFrame()
    x=df[df.apply(_is_official_radar_row,axis=1)].copy()
    if x.empty:
        # UI KeyError vermesin.
        x['Kurum Türü']=pd.Series(dtype=str)
        return x
    x['Kurum Türü']=x.apply(_v109_official_source_type,axis=1)
    x['Tarih_dt']=pd.to_datetime(x.get('Tarih_dt'),utc=True,errors='coerce')
    x=x.sort_values('Tarih_dt',ascending=False,na_position='last')
    return x.drop_duplicates(subset=['URL','Başlık'])

def _v110_merge_change_rows(out):
    """
    Aynı yeni olay farklı başlıklarla geldiyse 'Dünden beri' tablosunda tek satıra indirir.
    Özellikle aynı yer + aynı olay kelimelerini taşıyan haberlerde daha toleranslıdır.
    """
    if out is None or out.empty:
        return out
    rows=[]
    for _,r in out.iterrows():
        rd=r.to_dict()
        merged=False
        for i,old in enumerate(rows):
            # Yalnız aynı değişim sınıfında birleştir.
            if str(old.get('Değişim',''))!=str(rd.get('Değişim','')):
                continue
            sim=_v104_event_similarity(
                old.get('Başlık',''),old.get('Ne Değişti?',''),
                rd.get('Başlık',''),rd.get('Ne Değişti?','')
            )
            # Aynı olay türü/yer için başlık token örtüşmesi.
            a=set(_title_tokens(old.get('Başlık','')))
            b=set(_title_tokens(rd.get('Başlık','')))
            overlap=len(a&b)/max(1,min(len(a),len(b))) if a and b else 0
            if sim>=0.42 or overlap>=0.55:
                # Daha yüksek riskli / daha çok kaynaklı satırı temsilci tut.
                old_score=(int(old.get('Risk',0) or 0),int(old.get('Kaynak Sayısı',0) or 0))
                new_score=(int(rd.get('Risk',0) or 0),int(rd.get('Kaynak Sayısı',0) or 0))
                if new_score>old_score:
                    rd['Kaynak Sayısı']=max(int(rd.get('Kaynak Sayısı',1) or 1),int(old.get('Kaynak Sayısı',1) or 1))
                    rows[i]=rd
                else:
                    rows[i]['Kaynak Sayısı']=max(int(old.get('Kaynak Sayısı',1) or 1),int(rd.get('Kaynak Sayısı',1) or 1))
                merged=True
                break
        if not merged:
            rows.append(rd)
    return pd.DataFrame(rows)

def _compare_since_previous(df,current_scan_id=None):
    """
    V110 — gerçek değişiklikleri gösterir ve kullanıcıya doğrudan ne değiştiğini söyler.
    'Yeni Olay' yalnız sınıflandırma türüdür; asıl bilgi 'Ne Değişti?' sütunundadır.
    """
    current=_v104_event_representatives(df)
    prev_id=_previous_scan_id(current_scan_id)
    previous=_load_scan_events(prev_id)
    if current is None or current.empty:
        return pd.DataFrame(),None,None
    if previous.empty:
        return pd.DataFrame(),prev_id,None

    prev_records=[p.to_dict() for _,p in previous.iterrows()]
    changes=[]

    for _,r in current.iterrows():
        c={
            'title':str(r.get('Başlık','') or ''),
            'source':str(r.get('Kaynak','') or ''),
            'url':str(r.get('URL','') or ''),
            'category':str(r.get('Kategori','') or ''),
            'summary':str(r.get('İçerik_Özeti','') or ''),
            'risk_score':int(r.get('Risk_Skoru',0) or 0),
            'risk_status':str(r.get('Risk_Durumu','') or ''),
            'verification':str(r.get('Doğrulama','') or ''),
            'source_count':int(r.get('Olay_Kaynak_Sayisi',1) or 1)
        }

        best=None; best_sim=0.0
        for pr in prev_records:
            sim=_v104_event_similarity(
                c['title'],c['summary'],
                pr.get('title',''),pr.get('summary','')
            )
            if c['url'] and c['url']==str(pr.get('url','') or ''):
                sim=max(sim,0.98)
            if sim>best_sim:
                best_sim=sim; best=pr

        if best is None or best_sim<0.50:
            kind='🆕 YENİ OLAY'
            priority=100+c['risk_score']
            prev_risk='—'
            # Genel metin yerine olayın kendisini doğrudan söyle.
            concise=_clean_note_text(c['summary'])
            first=_v109_sentences(concise)
            if first:
                detail=_v66_formalize_sentence_endings(first[0]).strip()
                if detail and detail[-1] not in '.!?': detail+='.'
                diff='Önceki taramada bulunmayan yeni gelişme tespit edilmiştir: '+detail
            else:
                diff='Önceki taramada bulunmayan yeni gelişme tespit edilmiştir: '+_clean_note_text(c['title']).rstrip('.')+'.'
        else:
            risk_up,verify_up,material,_,_=_v104_material_change(best,c)
            if risk_up:
                kind='⚠️ RİSK ARTTI'; priority=95+c['risk_score']
            elif verify_up:
                kind='✅ TEYİT GÜÇLENDİ'; priority=90+c['risk_score']
            elif material:
                kind='🔄 YENİ BİLGİ'; priority=80+c['risk_score']
            else:
                continue
            prev_risk=int(best.get('risk_score') or 0)
            diff=_v109_direct_difference(best,c,kind)

        changes.append({
            'Ne Değişti?':diff,
            'Tür':kind,
            'Başlık':c['title'],
            'Kaynak':c['source'],
            'Kategori':c['category'],
            'Risk':c['risk_score'],
            'Önceki Risk':prev_risk,
            'Kaynak Sayısı':c['source_count'],
            'URL':c['url'],
            '_priority':priority
        })

    out=pd.DataFrame(changes)
    if not out.empty:
        # Birleştirme fonksiyonunun mevcut isimle çalışması için geçici Değişim alanı.
        out['Değişim']=out['Tür']
        out=_v110_merge_change_rows(out)
        if not out.empty:
            if 'Tür' not in out.columns and 'Değişim' in out.columns:
                out['Tür']=out['Değişim']
            out=out.sort_values(['_priority','Risk'],ascending=[False,False],na_position='last')
            out=out.drop(columns=['_priority','Değişim'],errors='ignore')

    prev_time=str(previous.iloc[0].get('scanned_at','')) if not previous.empty else None
    return out,prev_id,prev_time

# ============================================================
# /V110
# ============================================================

# V110 — V109 düzeltmesi: override sırası düzeltildi; Resmî Kaynak Radarı KeyError giderildi;
# Dünden Beri Ne Değişti tablosunda 'Ne Değişti?' ana sütun haline getirildi ve benzer yeni olaylar birleştirildi.


# ============================================================
# V111 — DÜNDEN BERİ / PANEL DEVAMLILIĞI DÜZELTMESİ
# ============================================================

def _compare_since_previous(df,current_scan_id=None):
    """
    V111:
    - 'Ne Değişti?' başlığı tekrar etmez.
    - Yeni olayda yalnızca olayın önceki taramada bulunmadığını söyler;
      olayın kendisi zaten Başlık sütununda görülür.
    - Eski modüller için 'Değişim' alanı korunur; yeni UI için 'Tür' de bulunur.
      Böylece Bilgi Notu Adayları ve panelin devamı KeyError ile durmaz.
    """
    current=_v104_event_representatives(df)
    prev_id=_previous_scan_id(current_scan_id)
    previous=_load_scan_events(prev_id)

    empty_cols=['Ne Değişti?','Tür','Değişim','Başlık','Kaynak','Kategori',
                'Risk','Önceki Risk','Kaynak Sayısı','URL']

    if current is None or current.empty:
        return pd.DataFrame(columns=empty_cols),None,None
    if previous.empty:
        return pd.DataFrame(columns=empty_cols),prev_id,None

    prev_records=[p.to_dict() for _,p in previous.iterrows()]
    changes=[]

    for _,r in current.iterrows():
        c={
            'title':str(r.get('Başlık','') or ''),
            'source':str(r.get('Kaynak','') or ''),
            'url':str(r.get('URL','') or ''),
            'category':str(r.get('Kategori','') or ''),
            'summary':str(r.get('İçerik_Özeti','') or ''),
            'risk_score':int(r.get('Risk_Skoru',0) or 0),
            'risk_status':str(r.get('Risk_Durumu','') or ''),
            'verification':str(r.get('Doğrulama','') or ''),
            'source_count':int(r.get('Olay_Kaynak_Sayisi',1) or 1)
        }

        best=None
        best_sim=0.0
        for pr in prev_records:
            sim=_v104_event_similarity(
                c['title'],c['summary'],
                pr.get('title',''),pr.get('summary','')
            )
            if c['url'] and c['url']==str(pr.get('url','') or ''):
                sim=max(sim,0.98)
            if sim>best_sim:
                best_sim=sim
                best=pr

        if best is None or best_sim<0.50:
            kind='🆕 YENİ OLAY'
            priority=100+c['risk_score']
            prev_risk='—'
            # Başlık zaten ayrı sütunda. Burada onu yeniden yazma.
            if c['source']:
                diff=(
                    f"Bu olaya ilişkin kayıt önceki taramada bulunmamaktadır; "
                    f"gelişme mevcut taramada ilk kez {c['source']} kaynağında tespit edilmiştir."
                )
            else:
                diff=(
                    "Bu olaya ilişkin kayıt önceki taramada bulunmamaktadır; "
                    "gelişme mevcut taramada ilk kez tespit edilmiştir."
                )
        else:
            risk_up,verify_up,material,_,_=_v104_material_change(best,c)

            if risk_up:
                kind='⚠️ RİSK ARTTI'
                priority=95+c['risk_score']
            elif verify_up:
                kind='✅ TEYİT GÜÇLENDİ'
                priority=90+c['risk_score']
            elif material:
                kind='🔄 YENİ BİLGİ'
                priority=80+c['risk_score']
            else:
                continue

            prev_risk=int(best.get('risk_score') or 0)
            diff=_v109_direct_difference(best,c,kind)

        changes.append({
            'Ne Değişti?':diff,
            'Tür':kind,
            # Geriye dönük uyumluluk: mevcut aday/özet fonksiyonları bunu kullanıyor.
            'Değişim':kind,
            'Başlık':c['title'],
            'Kaynak':c['source'],
            'Kategori':c['category'],
            'Risk':c['risk_score'],
            'Önceki Risk':prev_risk,
            'Kaynak Sayısı':c['source_count'],
            'URL':c['url'],
            '_priority':priority
        })

    out=pd.DataFrame(changes)
    if out.empty:
        return pd.DataFrame(columns=empty_cols),prev_id,str(previous.iloc[0].get('scanned_at',''))

    # Aynı yeni olayın farklı kaynaklarını tek satırda birleştir.
    out=_v110_merge_change_rows(out)

    # Birleştirme sonrası iki isim de kesinlikle bulunsun.
    if 'Tür' not in out.columns and 'Değişim' in out.columns:
        out['Tür']=out['Değişim']
    if 'Değişim' not in out.columns and 'Tür' in out.columns:
        out['Değişim']=out['Tür']

    out=out.sort_values(['_priority','Risk'],ascending=[False,False],na_position='last')
    out=out.drop(columns=['_priority'],errors='ignore')

    # UI başlığı tekrar etmesin; fakat eski fonksiyonlar Değişim'i kullanabilsin.
    ordered=[c for c in empty_cols if c in out.columns]
    extras=[c for c in out.columns if c not in ordered]
    out=out[ordered+extras]

    prev_time=str(previous.iloc[0].get('scanned_at','')) if not previous.empty else None
    return out,prev_id,prev_time

# ============================================================
# /V111
# ============================================================

# V111 — Bilgi Notu Adayları KeyError düzeltildi; panelin devamı artık yüklenir.
# 'Ne Değişti?' yeni olaylarda başlığı tekrar etmez; Değişim/Tür çift alanı ile geriye dönük uyumluluk sağlanmıştır.


# ============================================================
# V112 — DURUM TARİHİ + HIZ OPTİMİZASYONU
# 1) Durum rozetlerinde işlemin yapıldığı tarih/saat gösterilir.
# 2) Bilgi notu üretiminde mevcut olay zenginleştirmesi + detay cache kullanılır.
# 3) Sepet silme işlemleri tek hızlı SQLite transaction ile yapılır.
# 4) Aynı taramadaki pahalı "Dünden beri" karşılaştırması rerun'larda cache'lenir.
# ============================================================

def _v112_status_key_variants(title='',url='',summary=''):
    keys=set()
    url=str(url or '').strip()
    title=str(title or '')
    summary=str(summary or '')
    if url:
        keys.add('U:'+url)
    tk=title_key(title)
    if tk:
        keys.add('T:'+tk)
    try:
        sig=' '.join(sorted(_v104_event_tokens(title,summary)))
        if sig:
            keys.add('E:'+sig)
    except Exception:
        pass
    return keys

def _v112_parse_status_time(value):
    try:
        dt=pd.to_datetime(value,utc=True,errors='coerce')
        if pd.isna(dt):
            return None
        return dt
    except Exception:
        return None

def _v112_format_status_time(value):
    dt=_v112_parse_status_time(value)
    if dt is None:
        return ''
    try:
        local_tz=datetime.now().astimezone().tzinfo
        dt=dt.tz_convert(local_tz)
        return dt.strftime('%d.%m.%Y %H:%M')
    except Exception:
        try:
            return dt.strftime('%d.%m.%Y %H:%M')
        except Exception:
            return ''

def _v112_status_history():
    """
    Her işlem için olay/URL/başlık anahtarını EN SON işlem tarihine eşler.
    Tek sorgu grubu + session cache: tüm paneller aynı veriyi tekrar tekrar okumaz.
    """
    cached=st.session_state.get('_v112_status_history_cache')
    if cached is not None:
        return cached

    result={'imp':{},'akt':{},'notes':{},'pres':{}}
    if not _init_history_db():
        return result

    specs=[
        ('important_basket','added_at','imp'),
        ('osint_report_basket','added_at','akt'),
        ('note_history','created_at','notes'),
        ('presentation_basket','added_at','pres'),
    ]
    try:
        with _history_connect() as conn:
            for table,time_col,key in specs:
                rows=conn.execute(
                    f"SELECT {time_col},title,url FROM {table} ORDER BY {time_col} DESC"
                ).fetchall()
                for ts,title,url in rows:
                    for k in _v112_status_key_variants(title,url,''):
                        # Sorgu DESC olduğu için ilk değer en yeni tarihtir.
                        if k not in result[key]:
                            result[key][k]=str(ts or '')
    except Exception:
        pass

    st.session_state['_v112_status_history_cache']=result
    return result

def _v63_status_sets():
    """Eski kodlarla uyumluluk için durum anahtarlarını set olarak döndürür."""
    h=_v112_status_history()
    return set(h['imp']),set(h['akt']),set(h['notes']),set(h['pres'])

def _v104_event_status_sets():
    """V104 olay bazlı durum altyapısının V112 tarihli cache ile uyumlu hali."""
    h=_v112_status_history()
    return {
        'imp':set(h['imp']),
        'akt':set(h['akt']),
        'notes':set(h['notes']),
        'pres':set(h['pres'])
    }

def _v73_invalidate_status_cache():
    st.session_state.pop('_v73_status_sets_cache',None)
    st.session_state.pop('_v104_status_cache',None)
    st.session_state.pop('_v112_status_history_cache',None)

def _v63_add_status_badges(df):
    """
    Durum örneği:
    📌 ÖGN — 22.08.2026 13:42 • 📝 Bilgi Notu — 22.08.2026 14:05
    """
    if df is None or df.empty:
        return df
    out=df.copy()
    hist=_v112_status_history()

    def badge(r):
        title=str(r.get('Başlık',r.get('title','')) or '')
        url=str(r.get('URL',r.get('url','')) or '').strip()
        summary=str(r.get('İçerik_Özeti',r.get('summary','')) or '')
        keys=_v112_status_key_variants(title,url,summary)

        def latest(bucket):
            vals=[hist[bucket].get(k) for k in keys if hist[bucket].get(k)]
            if not vals:
                return ''
            parsed=[(_v112_parse_status_time(v),v) for v in vals]
            parsed=[x for x in parsed if x[0] is not None]
            if parsed:
                return max(parsed,key=lambda x:x[0])[1]
            return vals[0]

        b=[]
        ts=latest('pres')
        if ts: b.append(f"🖥️ Sunum — {_v112_format_status_time(ts)}")
        ts=latest('imp')
        if ts: b.append(f"📌 ÖGN — {_v112_format_status_time(ts)}")
        ts=latest('notes')
        if ts: b.append(f"📝 Bilgi Notu — {_v112_format_status_time(ts)}")
        ts=latest('akt')
        if ts: b.append(f"📁 AKT — {_v112_format_status_time(ts)}")
        return ' • '.join(b) if b else '—'

    out['Durum']=out.apply(badge,axis=1)
    return out

def _v112_fast_delete(table, ids=None):
    """Silme için tek transaction; satır satır işlem ve gereksiz sorgu yoktur."""
    allowed={'important_basket','osint_report_basket','presentation_basket'}
    if table not in allowed or not _init_history_db():
        return 0
    try:
        with _history_connect() as conn:
            # Web uygulamasında silme gecikmesini azaltmak için küçük transaction.
            conn.execute("PRAGMA synchronous=NORMAL")
            if ids is None:
                cur=conn.execute(f"DELETE FROM {table}")
            else:
                ids=[int(x) for x in ids if str(x).isdigit()]
                if not ids:
                    return 0
                marks=','.join('?' for _ in ids)
                cur=conn.execute(f"DELETE FROM {table} WHERE id IN ({marks})",ids)
            conn.commit()
            removed=max(int(cur.rowcount or 0),0)
        if removed:
            _v73_invalidate_status_cache()
        return removed
    except Exception:
        return 0

def _remove_basket_ids(ids):
    return _v112_fast_delete('important_basket',ids)

def _clear_important_basket():
    return _v112_fast_delete('important_basket',None)

def _remove_osint_basket_ids(ids):
    return _v112_fast_delete('osint_report_basket',ids)

def _clear_osint_basket():
    return _v112_fast_delete('osint_report_basket',None)

def _v81_remove_presentation_ids(ids):
    return _v112_fast_delete('presentation_basket',ids)

def _v80_clear_presentation():
    return _v112_fast_delete('presentation_basket',None)

def _v112_detail_cache_key(row):
    url=str(row.get('URL',row.get('url','')) or '').strip()
    if url:
        return 'U:'+url
    return 'T:'+title_key(str(row.get('Başlık',row.get('title','')) or ''))

def _v112_cached_article_detail(row):
    """
    Aynı haber için tekrar Word üretildiğinde sayfayı yeniden indirmez.
    Cache yalnız mevcut kullanıcı oturumunda tutulur.
    """
    cache=st.session_state.setdefault('_v112_article_detail_cache',{})
    key=_v112_detail_cache_key(row)
    if key in cache:
        return cache[key]
    try:
        detail=article_detail(row) or {}
    except Exception:
        detail={}
    # Cache'in sınırsız büyümesini önle.
    if len(cache)>160:
        try:
            for old in list(cache.keys())[:40]:
                cache.pop(old,None)
        except Exception:
            cache={}
            st.session_state['_v112_article_detail_cache']=cache
    cache[key]=detail
    return detail

def make_analyst_docx(df, title='BİLGİ NOTU'):
    """
    V112 hızlı bilgi notu:
    - Önce V107 olay zenginleştirmesini kullanır.
    - Yeterince zengin mevcut özet varsa yeniden web isteği yapmaz.
    - Gerekli haber detaylarını paralel ve oturum cache'li alır.
    - V66 resmî dil / belge yapısı korunur.
    """
    doc=Document()
    sec=doc.sections[0]
    sec.top_margin=Cm(2); sec.bottom_margin=Cm(2)
    sec.left_margin=Cm(2.5); sec.right_margin=Cm(2.5)
    styles=doc.styles
    styles['Normal'].font.name='Times New Roman'
    styles['Normal'].font.size=Pt(12)
    styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'),'Times New Roman')

    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run(_clean_note_text(title)); r.bold=True; r.font.size=Pt(14)
    p=doc.add_paragraph(); p.add_run('Tarih: ').bold=True
    p.add_run(datetime.now().astimezone().strftime('%d.%m.%Y'))

    x=df.copy() if df is not None else pd.DataFrame()
    if x.empty:
        rows=[]
    else:
        if 'Tarih_dt' in x.columns:
            x['Tarih_dt']=pd.to_datetime(x['Tarih_dt'],utc=True,errors='coerce')
            x=x.sort_values('Tarih_dt',ascending=True,na_position='last')
        rows=x.to_dict('records')

    # Aynı olayın mevcut taramadaki daha iyi kaynaklarını önce birleştir.
    try:
        enriched_rows=_v107_enrich_selected_rows(rows)
        if enriched_rows:
            rows=enriched_rows
    except Exception:
        pass

    enriched=[None]*len(rows)

    def get_one(i,row):
        summary=_clean_note_text(row.get('İçerik_Özeti',''))
        # V107 zenginleştirmesi yeterli içerik sağladıysa web fetch'i atla.
        rich_enough=(
            len(summary)>=650
            and len(_sentence_chunks(summary))>=3
        )
        if rich_enough:
            return i,row,{}
        return i,row,_v112_cached_article_detail(row)

    if rows:
        workers=min(6,len(rows))
        if workers<=1:
            for i,row in enumerate(rows):
                _,rr,dd=get_one(i,row)
                enriched[i]=(rr,dd)
        else:
            with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as ex:
                futs=[ex.submit(get_one,i,row) for i,row in enumerate(rows)]
                for fut in concurrent.futures.as_completed(futs):
                    try:
                        i,rr,dd=fut.result()
                        enriched[i]=(rr,dd)
                    except Exception:
                        pass

    enriched=[x for x in enriched if x is not None]

    all_sent=[]
    for row,detail in enriched:
        title_text=_clean_note_text(detail.get('title') or row.get('Başlık',''))
        body=_clean_note_text(detail.get('text') or row.get('İçerik_Özeti') or title_text)
        all_sent.extend(_akt_clean_sentences(title_text,body))

    uniq=[]; seen=[]
    for sent in all_sent:
        sent=_clean_note_text(sent)
        key=norm(sent)
        toks=set(key.split())
        if not key: continue
        dup=False
        for old in seen[-35:]:
            union=len(toks|old)
            if union and len(toks&old)/union>=0.78:
                dup=True; break
        if not dup:
            uniq.append(sent.strip()); seen.append(toks)

    def add_body(text):
        bp=doc.add_paragraph()
        bp.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        bp.paragraph_format.first_line_indent=Cm(1.25)
        bp.paragraph_format.line_spacing=1.15
        bp.paragraph_format.space_after=Pt(8)
        safe_text=_repair_mojibake_utf8(_clean_note_text(text))
        bp.add_run(_v66_formalize_sentence_endings(safe_text))

    if uniq:
        intro=_join_sentences_naturally(uniq[:2])
        add_body(intro)

        detail_s=uniq[2:] or uniq
        detail_s=detail_s[:18]
        if len(detail_s)<=9:
            add_body(_join_sentences_naturally(detail_s))
        else:
            add_body(_join_sentences_naturally(detail_s[:9]))
            add_body(_join_sentences_naturally(detail_s[9:18]))

        tail=_join_sentences_naturally(uniq[-3:])
        if tail:
            conclusion=(
                f"Mevcut bilgiler çerçevesinde, {tail[0].lower()+tail[1:]} "
                "Gelişmenin sanayi ve teknoloji alanındaki muhtemel etkilerinin, ilgili kurum ve kuruluşların "
                "yeni açıklamaları ile resmî veriler doğrultusunda takip edilmesinin uygun olacağı değerlendirilmektedir."
            )
        else:
            conclusion=(
                "Mevcut bilgiler çerçevesinde gelişmenin sanayi ve teknoloji alanındaki etkilerinin, ilgili kurum "
                "ve kuruluşların yeni açıklamaları ile resmî veriler doğrultusunda takip edilmesinin uygun olacağı değerlendirilmektedir."
            )
        add_body(conclusion)
    else:
        add_body('Seçilen habere ilişkin ayrıntılı içerik temin edilememiştir.')
        add_body(
            'Gelişmenin yeni açık kaynak bilgileri ile ilgili kurum ve kuruluşların resmî açıklamaları '
            'doğrultusunda takip edilmesinin uygun olacağı değerlendirilmektedir.'
        )

    endp=doc.add_paragraph()
    endp.paragraph_format.space_before=Pt(8)
    endp.add_run('Arz olunur.')

    if enriched:
        kp=doc.add_paragraph()
        kr=kp.add_run('Kaynak: '); kr.bold=True
        for i,(row,detail) in enumerate(enriched):
            source=_clean_note_text(detail.get('source') or row.get('Kaynak','Açık Kaynak'))
            url=detail.get('canonical') or row.get('Yayıncı_URL') or row.get('URL','')
            if i: kp.add_run('; ')
            kp.add_run(source)
            if url:
                kp.add_run(' ('); _word_hyperlink(kp,url,'Haber linki'); kp.add_run(')')

    bio=BytesIO()
    doc.save(bio); bio.seek(0)
    return bio.getvalue()

# Aynı taramadaki pahalı karşılaştırmayı sepet silme gibi rerun'larda tekrar hesaplama.
_v112_compare_impl=_compare_since_previous

def _v112_scan_cache_key(df,current_scan_id=None):
    try:
        n=len(df)
    except Exception:
        n=0
    sid=current_scan_id or st.session_state.get('current_scan_id') or 'none'
    return f'{sid}:{n}'

def _compare_since_previous(df,current_scan_id=None):
    key=_v112_scan_cache_key(df,current_scan_id)
    cache=st.session_state.setdefault('_v112_compare_cache',{})
    if key in cache:
        out,prev_id,prev_time=cache[key]
        return out.copy(),prev_id,prev_time
    result=_v112_compare_impl(df,current_scan_id)
    out,prev_id,prev_time=result
    cache.clear()
    cache[key]=(out.copy(),prev_id,prev_time)
    return out,prev_id,prev_time

# ============================================================
# /V112
# ============================================================

# V112 — Durum alanında işlem tarih/saatleri gösterilir; bilgi notu ve sepet silme akışları hızlandırılmıştır.


# ============================================================
# V114 — KONTROL MERKEZİ + EK HIZ OPTİMİZASYONU
# 1) "Analist Komuta Merkezi" adı "Kontrol Merkezi" olarak değiştirildi.
# 2) Açılıştaki otomatik geri-dönüş ağ taraması kaldırıldı; "Şu An Bilmen Gerekenler"
#    son ana tarama verisi üzerinden hazırlanır.
# 3) Türk ana + resmî + istatistik + negatif + opsiyonel kaynak sorguları
#    tek paralel havuzda çalıştırılır; sıralı ağ beklemesi kaldırılır.
# Mevcut tarama kapsamı, sınıflandırma, sepetler ve rapor üretimi korunur.
# ============================================================

# ============================================================
# V113 — PANEL HIZ / TEKRAR HESAP ÖNLEME
#
# Amaç:
# - V112 kararlı tarama çekirdeğine ve panel içeriğine dokunmadan,
#   aynı taramadaki pahalı türetilmiş tabloları bir kez hesaplayıp yeniden kullanmak.
# - Özellikle Vardiya Başlangıç Özeti sonrasında sıralı biçimde tekrar hesaplanan
#   olay/değer/kronoloji/resmî kaynak/yaşam döngüsü tablolarının bekleme süresini azaltmak.
#
# Değişmeyenler:
# - Tarama sorguları ve kaynaklar
# - Risk/negatif sınıflandırması
# - Bölüm sırası ve görünümü
# - Checkbox/form işleyişi
# - Sepetler ve Word çıktıları
# ============================================================

def _v113_scan_key(df=None, current_scan_id=None):
    """Aynı taramaya ait bütün panel cache'lerinde ortak anahtar."""
    sid = current_scan_id or st.session_state.get('current_scan_id') or 'none'
    try:
        n = len(df)
    except Exception:
        n = 0

    # Aynı scan_id ile dataframe gerçekten değişirse cache yanlış kalmasın.
    last_dt = ''
    try:
        if df is not None and not df.empty and 'Tarih_dt' in df.columns:
            s = pd.to_datetime(df['Tarih_dt'], utc=True, errors='coerce')
            mx = s.max()
            last_dt = '' if pd.isna(mx) else str(mx.value)
    except Exception:
        last_dt = ''
    return f'{sid}:{n}:{last_dt}'

def _v113_cache():
    cache = st.session_state.setdefault('_v113_panel_cache', {})
    active = st.session_state.get('_v113_panel_cache_active_key')
    cur_sid = st.session_state.get('current_scan_id') or 'none'

    # Yeni taramada eski ağır DataFrame'leri RAM'de tutma.
    if active is not None and not str(active).startswith(str(cur_sid)+':'):
        cache.clear()
        st.session_state['_v113_panel_cache_active_key'] = None
    return cache

def _v113_get_cached(name, key):
    return _v113_cache().get((name, key))

def _v113_set_cached(name, key, value):
    cache = _v113_cache()
    cache[(name, key)] = value
    st.session_state['_v113_panel_cache_active_key'] = key
    # Tek tarama için yeterli; cache şişmesini önle.
    if len(cache) > 24:
        keep = {}
        for k, v in list(cache.items())[-18:]:
            keep[k] = v
        cache.clear()
        cache.update(keep)
    return value

# ------------------------------------------------------------
# 1) OLAY ÇERÇEVESİ — Bilgi Notu Adayları vb. aynı groupby'ı tekrar yapmasın.
# ------------------------------------------------------------
_v113_current_event_frame_impl = _current_event_frame

def _current_event_frame(df):
    key = _v113_scan_key(df)
    cached = _v113_get_cached('current_event_frame', key)
    if cached is not None:
        return cached.copy()
    out = _v113_current_event_frame_impl(df)
    _v113_set_cached('current_event_frame', key, out.copy())
    return out

# ------------------------------------------------------------
# 2) DEĞER TABLOSU — Kontrol Merkezi + Top10 + İkinci Göz tek hesap kullansın.
# ------------------------------------------------------------
_v113_value_table_impl = _v52_event_value_table

def _v52_event_value_table(df, n=10):
    if df is None or df.empty:
        return _v113_value_table_impl(df, n)

    key = _v113_scan_key(df)
    cached = _v113_get_cached('value_table_full', key)

    if cached is None:
        try:
            # Fonksiyonun n parametresi yalnız son head() aşamasında kullanılıyor.
            # Olayların tamamını bir kez puanla; sonraki çağrılar yalnız slice yapsın.
            event_count = int(df['Olay_ID'].nunique(dropna=False)) if 'Olay_ID' in df.columns else len(df)
            full_n = max(60, event_count)
        except Exception:
            full_n = max(60, int(n or 10))

        cached = _v113_value_table_impl(df, full_n)
        _v113_set_cached('value_table_full', key, cached.copy())

    out = cached.head(int(n or 10)).copy().reset_index(drop=True)
    if 'Sıra' in out.columns:
        out['Sıra'] = range(1, len(out)+1)
    return out

# ------------------------------------------------------------
# 3) BİLGİ NOTU ADAYLARI — 15 aday bir kez, slider yalnız dilimlesin.
# ------------------------------------------------------------
_v113_information_candidates_impl = _information_note_candidates

def _information_note_candidates(df, current_scan_id=None, limit=10):
    key = _v113_scan_key(df, current_scan_id)
    cached = _v113_get_cached('information_candidates_15', key)
    if cached is None:
        cached = _v113_information_candidates_impl(df, current_scan_id, 15)
        _v113_set_cached('information_candidates_15', key, cached.copy())
    return cached.head(int(limit or 10)).copy().reset_index(drop=True)

# ------------------------------------------------------------
# 4) KRONOLOJİ OLAY TABLOSU — sayfa değişimlerinde yeniden groupby yapma.
# ------------------------------------------------------------
_v113_chronology_impl = _v109_chronology_events

def _v109_chronology_events(df):
    key = _v113_scan_key(df)
    cached = _v113_get_cached('chronology_events', key)
    if cached is not None:
        return cached.copy()
    out = _v113_chronology_impl(df)
    _v113_set_cached('chronology_events', key, out.copy())
    return out

# ------------------------------------------------------------
# 5) RESMÎ KAYNAK RADARI — aynı apply() her rerun'da tekrar çalışmasın.
# ------------------------------------------------------------
_v113_official_radar_impl = _official_radar_rows

def _official_radar_rows(df):
    key = _v113_scan_key(df)
    cached = _v113_get_cached('official_radar', key)
    if cached is not None:
        return cached.copy()
    out = _v113_official_radar_impl(df)
    _v113_set_cached('official_radar', key, out.copy())
    return out

# ------------------------------------------------------------
# 6) OLAY YAŞAM DÖNGÜSÜ — tam tablo bir kez hesaplanır, limit sonradan uygulanır.
# ------------------------------------------------------------
_v113_lifecycle_impl = _v58_event_lifecycle_table

def _v58_event_lifecycle_table(df, limit=25):
    key = _v113_scan_key(df)
    cached = _v113_get_cached('event_lifecycle_60', key)
    if cached is None:
        cached = _v113_lifecycle_impl(df, 60)
        _v113_set_cached('event_lifecycle_60', key, cached.copy())
    return cached.head(int(limit or 25)).copy().reset_index(drop=True)

# ------------------------------------------------------------
# 7) TREND TABLOSU — görünüm değiştirirken tekrar üretme.
# ------------------------------------------------------------
try:
    _v113_trend_impl = trend_table
    def trend_table(df):
        key = _v113_scan_key(df)
        cached = _v113_get_cached('trend_table', key)
        if cached is not None:
            return cached.copy()
        out = _v113_trend_impl(df)
        _v113_set_cached('trend_table', key, out.copy())
        return out
except Exception:
    pass

# ------------------------------------------------------------
# 8) VARDİYA BAŞLANGIÇ ÖZETİ — aynı taramada widget rerun'larında tekrar hesaplama.
# Devir noktası anahtara eklenir; kullanıcı yeni devir noktası kaydederse cache yenilenir.
# ------------------------------------------------------------
_v113_shift_summary_impl = _shift_start_summary

def _v113_shift_mark_key():
    try:
        mark = _latest_shift_mark()
        return str((mark or {}).get('marked_at',''))
    except Exception:
        return ''

def _shift_start_summary(df, current_scan_id=None):
    key = _v113_scan_key(df, current_scan_id) + ':' + _v113_shift_mark_key()
    cached = _v113_get_cached('shift_summary', key)
    if cached is not None:
        stats, top, label = cached
        return dict(stats), top.copy(), label
    stats, top, label = _v113_shift_summary_impl(df, current_scan_id)
    _v113_set_cached('shift_summary', key, (dict(stats), top.copy(), label))
    return stats, top, label

# ------------------------------------------------------------
# 9) KRİTİK SANAYİ OLAYLARI — aynı haber için incident fonksiyonu iki kez çağrılmasın.
# ------------------------------------------------------------
def _v113_critical_events_table(df):
    if df is None or df.empty:
        return pd.DataFrame()

    key = _v113_scan_key(df)
    cached = _v113_get_cached('critical_events', key)
    if cached is not None:
        return cached.copy()

    labels = []
    for title, summary in zip(
        df.get('Başlık', pd.Series('', index=df.index)).fillna('').astype(str),
        df.get('İçerik_Özeti', pd.Series('', index=df.index)).fillna('').astype(str)
    ):
        labels.append(critical_industrial_incident(title, summary) or '')

    mask = pd.Series([bool(x) for x in labels], index=df.index)
    out = df.loc[mask].copy()
    if not out.empty:
        out['Kritik_Olay'] = [labels[i] for i, flag in enumerate(mask.tolist()) if flag]
        if 'Tarih_dt' in out.columns:
            out = out.sort_values('Tarih_dt', ascending=False, na_position='last')

    _v113_set_cached('critical_events', key, out.copy())
    return out

# ============================================================
# /V113
# ============================================================


# ============================================================
# V115 — SUNUM ÖNCESİ KARARLILIK + HIZ + KURUMSAL ÇIKTI REVİZYONU
# 1) Vardiya Başlangıç Özeti içindeki ağır önceki-tarama eşleştirmesi ters indeksle hızlandırıldı.
# 2) Vardiya üst listesi hafif, olay-bazlı puanlama ile hazırlanır; sayfanın bu noktada beklemesi azaltılır.
# 3) Google News bağlantıları için ek paket gerektirmeyen doğrudan yayıncı arama fallback'i eklendi.
# 4) ÖGN, AKT ve Bilgi Notu Word motorları gönderilen kurumsal örneklere göre yeniden düzenlendi.
# 5) Mevcut tarama/risk/sepet iş kuralları korunur; yalnız türetilmiş görünüm ve çıktı katmanı override edilir.
# ============================================================


# V115 doğruluk düzeltmesi: önceki domain() yalnız tam URL'lerde çalışıyordu.
# normalize_rows() ise source_group/source_reliability/source_rank fonksiyonlarına
# çoğunlukla "aa.com.tr" gibi çıplak alan adı gönderiyor. Bu nedenle güvenilir
# Türk/resmî kaynaklar yanlışlıkla "Diğer / Açık Kaynak" sınıfına düşebiliyordu.
_v115_domain_base=domain
def domain(value):
    s=str(value or '').strip()
    if not s:
        return ''
    try:
        # Tam URL.
        if '://' in s:
            return urlparse(s).netloc.lower().split('@')[-1].split(':')[0].replace('www.','')
        # Çıplak domain veya domainsiz path.
        head=s.split('/')[0].strip().lower().split('@')[-1].split(':')[0].replace('www.','')
        if re.fullmatch(r'[a-z0-9çğıöşü.-]+\.[a-zçğıöşü]{2,}',head,re.I):
            return head
        # Son fallback mevcut davranış.
        return _v115_domain_base(s)
    except Exception:
        return ''


def _v115_is_direct_article_url(url):
    try:
        u=str(url or '').strip()
        if not u.startswith(('http://','https://')):
            return False
        p=urlparse(u)
        host=(p.netloc or '').lower().replace('www.','')
        if not host or 'google.com' in host or host=='news.google.com':
            return False
        # Ana sayfa değil, haber yolu olma ihtimali bulunan URL'leri tercih et.
        path=(p.path or '').strip('/')
        return bool(path and len(path)>=4)
    except Exception:
        return False


def _v115_search_direct_article(title, preferred_domain=''):
    """Ek paket kullanmadan, yalnız belge üretiminde gerçek yayıncı URL'si bulmaya çalışır."""
    title=_clean_note_text(title)
    if not title:
        return ''
    try:
        from urllib.parse import parse_qs, unquote
    except Exception:
        parse_qs=None; unquote=lambda x:x

    preferred=(preferred_domain or '').lower().replace('www.','').strip()
    if preferred.startswith('http'):
        try: preferred=urlparse(preferred).netloc.lower().replace('www.','')
        except Exception: preferred=''
    if 'google.com' in preferred:
        preferred=''

    q=f'"{title[:220]}"'
    if preferred:
        q+=f' site:{preferred}'

    def clean_href(href):
        href=str(href or '').strip()
        if not href:
            return ''
        try:
            if 'uddg=' in href and parse_qs is not None:
                qs=parse_qs(urlparse(href).query)
                if qs.get('uddg'):
                    href=unquote(qs['uddg'][0])
        except Exception:
            pass
        return href

    candidates=[]
    # DuckDuckGo HTML: harici Python paketi gerektirmez.
    try:
        rr=requests.get(
            'https://html.duckduckgo.com/html/',
            params={'q':q},
            headers={**HEADERS,'Accept-Language':'tr-TR,tr;q=0.9,en;q=0.6'},
            timeout=6
        )
        if rr.ok and rr.text:
            soup=BeautifulSoup(rr.text,'html.parser')
            for a in soup.select('a.result__a, a.result-link')[:10]:
                href=clean_href(a.get('href'))
                if _v115_is_direct_article_url(href):
                    candidates.append((href,_clean_note_text(a.get_text(' ',strip=True))))
    except Exception:
        pass

    # DDG erişilemezse Bing HTML tek yedek aramadır.
    if not candidates:
        try:
            rr=requests.get(
                'https://www.bing.com/search',
                params={'q':q,'setlang':'tr'},
                headers=HEADERS,
                timeout=6
            )
            if rr.ok and rr.text:
                soup=BeautifulSoup(rr.text,'html.parser')
                for a in soup.select('li.b_algo h2 a')[:10]:
                    href=str(a.get('href') or '').strip()
                    if _v115_is_direct_article_url(href):
                        candidates.append((href,_clean_note_text(a.get_text(' ',strip=True))))
        except Exception:
            pass

    if not candidates:
        return ''

    # Önce beklenen yayıncı alan adını, sonra başlık benzerliğini tercih et.
    title_toks=set(_title_tokens(title))
    best=''; best_score=-1.0
    for href,label in candidates:
        host=urlparse(href).netloc.lower().replace('www.','')
        lt=set(_title_tokens(label))
        sim=(len(title_toks&lt)/max(1,len(title_toks|lt))) if title_toks and lt else 0.0
        score=sim*10
        if preferred and (host==preferred or host.endswith('.'+preferred) or preferred.endswith('.'+host)):
            score+=20
        if score>best_score:
            best_score=score; best=href
    return best


_v115_article_detail_base=article_detail

def article_detail(row):
    """
    V115: mevcut ayrıntı çıkarma motorunu korur; Google News çözümlenemediğinde
    başlık + yayıncı alan adı ile gerçek haber sayfasını ek paket gerektirmeden bulur.
    """
    if isinstance(row,str):
        row={'URL':row}
    elif hasattr(row,'to_dict'):
        row=row.to_dict()
    elif row is None:
        row={}
    else:
        row=dict(row)
    detail=_v115_article_detail_base(row) or {}

    title=_clean_note_text(detail.get('title') or row.get('Başlık',''))
    original=str(row.get('URL','') or '').strip()
    publisher_url=str(row.get('Yayıncı_URL','') or '').strip()
    preferred=''
    try:
        preferred=urlparse(publisher_url).netloc.lower().replace('www.','')
    except Exception:
        preferred=''

    canonical=str(detail.get('canonical') or '').strip()
    body=_clean_note_text(detail.get('text') or '')
    images=detail.get('images') or []
    unresolved=(not _v115_is_direct_article_url(canonical))
    too_thin=(len(body)<380 or len(_akt_clean_sentences(title,body))<2)

    # RSS kaydı gerçek yayıncı sayfasına çözülemediyse veya gövde çok zayıfsa bir kez arama yap.
    if title and (unresolved or too_thin or not images):
        direct=''
        if _v115_is_direct_article_url(publisher_url):
            direct=publisher_url
        else:
            direct=_v115_search_direct_article(title,preferred)
        if direct and direct!=original:
            try:
                rr=dict(row)
                rr['URL']=direct
                rr['Yayıncı_URL']=direct
                better=_v115_article_detail_base(rr) or {}
                better_text=_clean_note_text(better.get('text') or '')
                # Daha zengin gövde, gerçek URL veya görsel sağlıyorsa yedek sonucu kabul et.
                if (_v115_is_direct_article_url(better.get('canonical')) and
                    (len(better_text)>len(body)+80 or len(better_text)>=450 or better.get('images'))):
                    detail=better
                    body=better_text
            except Exception:
                pass

    # Son güvenlik: gerçek yayıncı URL'si bulunamadıysa var olmayan bir ana sayfayı haber linki diye yazma.
    if not _v115_is_direct_article_url(detail.get('canonical')):
        if _v115_is_direct_article_url(publisher_url):
            detail['canonical']=publisher_url
        else:
            detail['canonical']=original
    return detail


def _v115_unique_fact_sentences(title,body,max_sentences=14,max_chars=4200):
    """Haber gövdesinden başlığı tekrarlamayan, sayısal/kurumsal ayrıntıları koruyan cümleleri seçer."""
    title=_clean_note_text(title)
    sents=_akt_clean_sentences(title,body)
    if not sents:
        fb=_clean_note_text(body or title)
        return [fb] if fb else []

    scores=[]
    for i,s in enumerate(sents):
        n=norm(s)
        score=_akt_sentence_score(s)+_sent_score(s)
        if i<3: score+=5
        if i>=max(0,len(sents)-3): score+=2
        if re.search(r'\b\d+(?:[.,]\d+)?\b',s): score+=3
        if any(k in n for k in ['bakanlık','kurum','şirket','türkiye','tüik','tübitak','kosgeb','kvkk','epdk']): score+=2
        scores.append((score,i,s))

    keep=set(range(min(2,len(sents))))
    for _,i,_ in sorted(scores,key=lambda z:(z[0],-z[1]),reverse=True):
        if len(keep)>=max_sentences:
            break
        keep.add(i)
    # Haberin son durumunu tamamen kaybetme.
    if len(sents)>2:
        keep.add(len(sents)-1)

    out=[]; total=0; old_sets=[]
    for i in sorted(keep):
        s=_clean_note_text(sents[i]).strip()
        if not s: continue
        toks=set(_history_tokens(s))
        dup=False
        for old in old_sets:
            union=len(toks|old)
            if union and len(toks&old)/union>=0.80:
                dup=True; break
        if dup: continue
        if out and total+len(s)>max_chars:
            break
        out.append(s); old_sets.append(toks); total+=len(s)+1
    return out


def _v115_formal_sentence(s):
    s=_clean_note_text(s).strip()
    if not s: return ''
    s=_v66_formalize_sentence_endings(s)
    s=re.sub(r'\s+',' ',s).strip()
    if s and s[-1] not in '.!?': s+='.'
    return s


def _v115_dedupe_rows(rows):
    """Aynı olayın farklı başlıklı kopyalarını Word çıktısına iki kez taşımamaya çalışır."""
    out=[]
    for r in rows or []:
        rd=dict(r)
        title=str(rd.get('title',rd.get('Başlık','')) or '')
        summary=str(rd.get('summary',rd.get('İçerik_Özeti','')) or '')
        url=str(rd.get('url',rd.get('URL','')) or '')
        duplicate=False
        for old in out:
            ot=str(old.get('title',old.get('Başlık','')) or '')
            os=str(old.get('summary',old.get('İçerik_Özeti','')) or '')
            ou=str(old.get('url',old.get('URL','')) or '')
            if url and ou and url==ou:
                duplicate=True; break
            if _v104_event_similarity(title,summary,ot,os)>=0.62:
                duplicate=True; break
        if not duplicate:
            out.append(rd)
    return out


# ------------------------------------------------------------------
# HIZ: önceki tarama eşleştirmesinde bütün previous x current çapraz döngüsünü
# ortak token taşıyan adaylarla sınırla. Çıktı şeması V111 ile aynıdır.
# ------------------------------------------------------------------
def _compare_since_previous(df,current_scan_id=None):
    empty_cols=['Ne Değişti?','Tür','Değişim','Başlık','Kaynak','Kategori',
                'Risk','Önceki Risk','Kaynak Sayısı','URL']
    key=_v112_scan_cache_key(df,current_scan_id)+':v115'
    cache=st.session_state.setdefault('_v115_compare_cache',{})
    if key in cache:
        out,pid,ptime=cache[key]
        return out.copy(),pid,ptime

    current=_v104_event_representatives(df)
    prev_id=_previous_scan_id(current_scan_id)
    previous=_load_scan_events(prev_id)
    if current is None or current.empty:
        return pd.DataFrame(columns=empty_cols),None,None
    if previous.empty:
        return pd.DataFrame(columns=empty_cols),prev_id,None

    prev_records=[]; token_index={}; url_index={}
    for pi,(_,p) in enumerate(previous.iterrows()):
        pr=p.to_dict()
        toks=_v104_event_tokens(pr.get('title',''),pr.get('summary',''))
        pr['_v115_tokens']=toks
        prev_records.append(pr)
        for t in toks:
            token_index.setdefault(t,set()).add(pi)
        u=str(pr.get('url','') or '').strip()
        if u: url_index.setdefault(u,set()).add(pi)

    changes=[]
    for _,r in current.iterrows():
        c={
            'title':str(r.get('Başlık','') or ''),'source':str(r.get('Kaynak','') or ''),
            'url':str(r.get('URL','') or ''),'category':str(r.get('Kategori','') or ''),
            'summary':str(r.get('İçerik_Özeti','') or ''),'risk_score':int(r.get('Risk_Skoru',0) or 0),
            'risk_status':str(r.get('Risk_Durumu','') or ''),'verification':str(r.get('Doğrulama','') or ''),
            'source_count':int(r.get('Olay_Kaynak_Sayisi',1) or 1)
        }
        toks=_v104_event_tokens(c['title'],c['summary'])
        cand=set()
        if c['url']: cand.update(url_index.get(c['url'],set()))
        for t in toks: cand.update(token_index.get(t,set()))

        best=None; best_sim=0.0
        for pi in cand:
            pr=prev_records[pi]
            sim=_v104_event_similarity(c['title'],c['summary'],pr.get('title',''),pr.get('summary',''))
            if c['url'] and c['url']==str(pr.get('url','') or ''):
                sim=max(sim,0.98)
            if sim>best_sim:
                best_sim=sim; best=pr

        if best is None or best_sim<0.50:
            kind='🆕 YENİ OLAY'; priority=100+c['risk_score']; prev_risk='—'
            diff=(f"Bu olaya ilişkin kayıt önceki taramada bulunmamaktadır; gelişme mevcut taramada ilk kez "
                  f"{c['source']} kaynağında tespit edilmiştir." if c['source'] else
                  "Bu olaya ilişkin kayıt önceki taramada bulunmamaktadır; gelişme mevcut taramada ilk kez tespit edilmiştir.")
        else:
            risk_up,verify_up,material,_,_=_v104_material_change(best,c)
            if risk_up:
                kind='⚠️ RİSK ARTTI'; priority=95+c['risk_score']
            elif verify_up:
                kind='✅ TEYİT GÜÇLENDİ'; priority=90+c['risk_score']
            elif material:
                kind='🔄 YENİ BİLGİ'; priority=80+c['risk_score']
            else:
                continue
            prev_risk=int(best.get('risk_score') or 0)
            diff=_v109_direct_difference(best,c,kind)

        changes.append({'Ne Değişti?':diff,'Tür':kind,'Değişim':kind,'Başlık':c['title'],
                        'Kaynak':c['source'],'Kategori':c['category'],'Risk':c['risk_score'],
                        'Önceki Risk':prev_risk,'Kaynak Sayısı':c['source_count'],'URL':c['url'],
                        '_priority':priority})

    out=pd.DataFrame(changes)
    prev_time=str(previous.iloc[0].get('scanned_at','')) if not previous.empty else None
    if out.empty:
        out=pd.DataFrame(columns=empty_cols)
    else:
        out=_v110_merge_change_rows(out)
        if 'Tür' not in out.columns and 'Değişim' in out.columns: out['Tür']=out['Değişim']
        if 'Değişim' not in out.columns and 'Tür' in out.columns: out['Değişim']=out['Tür']
        out=out.sort_values(['_priority','Risk'],ascending=[False,False],na_position='last')
        out=out.drop(columns=['_priority'],errors='ignore')
        ordered=[c for c in empty_cols if c in out.columns]
        out=out[ordered+[c for c in out.columns if c not in ordered]]

    cache.clear(); cache[key]=(out.copy(),prev_id,prev_time)
    return out,prev_id,prev_time


def _v115_fast_shift_top(since,n=8):
    if since is None or since.empty:
        return pd.DataFrame()
    x=since.copy()
    x['Tarih_dt']=pd.to_datetime(x.get('Tarih_dt'),utc=True,errors='coerce')
    txt=(x.get('Başlık',pd.Series('',index=x.index)).fillna('').astype(str)+' '+
         x.get('İçerik_Özeti',pd.Series('',index=x.index)).fillna('').astype(str)+' '+
         x.get('Kategori',pd.Series('',index=x.index)).fillna('').astype(str)).str.lower()
    risk=pd.to_numeric(x.get('Risk_Skoru',0),errors='coerce').fillna(0)
    sources=pd.to_numeric(x.get('Olay_Kaynak_Sayisi',0),errors='coerce').fillna(0)
    ver=x.get('Doğrulama',pd.Series('',index=x.index)).fillna('').astype(str).str.lower()
    strategic=txt.str.contains(r'yatırım|üretim|tesis|fabrika|çip|yarı iletken|yapay zeka|yapay zekâ|siber|ar-ge|patent|otomotiv|enerji|savunma|uzay|uydu',regex=True)
    official=ver.str.contains(r'resm|birincil',regex=True)
    neg=x.get('Duygu',pd.Series('',index=x.index)).fillna('').astype(str).eq('Negatif')
    x['_v115_shift_score']=risk + sources.clip(upper=4)*6 + strategic.astype(int)*24 + official.astype(int)*30 + neg.astype(int)*8
    x=x.sort_values(['_v115_shift_score','Tarih_dt'],ascending=[False,False],na_position='last')
    if 'Olay_ID' in x.columns:
        x=x.drop_duplicates('Olay_ID',keep='first')
    else:
        x=x.drop_duplicates('Başlık',keep='first')
    return x.head(int(n)).drop(columns=['_v115_shift_score'],errors='ignore')


def _shift_start_summary(df,current_scan_id=None):
    """V115: vardiya özeti ağ isteği yapmaz ve ağır tüm-çift eşleştirme kullanmaz."""
    if df is None or df.empty:
        return {},pd.DataFrame(),''
    key=_v113_scan_key(df,current_scan_id)+':'+_v113_shift_mark_key()+':v115'
    cached=_v113_get_cached('shift_summary_v115',key)
    if cached is not None:
        stats,top,label=cached
        return dict(stats),top.copy(),label

    baseline,baseline_label,_=_shift_baseline(current_scan_id)
    x=df.copy(); x['Tarih_dt']=pd.to_datetime(x.get('Tarih_dt'),utc=True,errors='coerce')
    since=x[(x['Tarih_dt'].isna()) | (x['Tarih_dt']>=baseline)].copy() if baseline is not None else x.copy()
    changes,_,_=_compare_since_previous(df,current_scan_id)
    type_col='Tür' if 'Tür' in changes.columns else ('Değişim' if 'Değişim' in changes.columns else None)
    if type_col:
        ct=changes[type_col].astype(str)
        new_events=int(ct.str.contains('YENİ OLAY').sum())
        risk_up=int(ct.str.contains('RİSK ARTTI').sum())
        verify_up=int(ct.str.contains('TEYİT').sum())
    else:
        new_events=risk_up=verify_up=0
    high=int((since.get('Risk_Durumu',pd.Series('',index=since.index))=='Yüksek Risk').sum())
    osb=sum(is_osb_fire(t,s) for t,s in zip(
        since.get('Başlık',pd.Series('',index=since.index)).fillna('').astype(str),
        since.get('İçerik_Özeti',pd.Series('',index=since.index)).fillna('').astype(str)))
    top=_v115_fast_shift_top(since,8)
    stats={'new_news':len(since),'new_important_events':new_events,'high_risk':high,
           'risk_up':risk_up,'verify_up':verify_up,'osb':int(osb),'baseline_label':baseline_label}
    _v113_set_cached('shift_summary_v115',key,(dict(stats),top.copy(),baseline_label))
    return stats,top,baseline_label


# ------------------------------------------------------------------
# WORD ORTAK AYARLARI
# ------------------------------------------------------------------
def _v115_doc_defaults(doc,top=2.5,bottom=2.5,left=2.5,right=2.5):
    sec=doc.sections[0]
    sec.top_margin=Cm(top); sec.bottom_margin=Cm(bottom)
    sec.left_margin=Cm(left); sec.right_margin=Cm(right)
    normal=doc.styles['Normal']
    normal.font.name='Times New Roman'; normal.font.size=Pt(12)
    normal._element.rPr.rFonts.set(qn('w:eastAsia'),'Times New Roman')
    return doc


def _v115_set_run_font(run,size=12,bold=None):
    run.font.name='Times New Roman'; run.font.size=Pt(size)
    run._element.rPr.rFonts.set(qn('w:eastAsia'),'Times New Roman')
    if bold is not None: run.bold=bool(bold)
    return run


def _v115_add_akt_info_table(doc,findings):
    table=doc.add_table(rows=3,cols=2)
    table.style='Table Grid'
    try:
        table.columns[0].width=Cm(5.7); table.columns[1].width=Cm(10.3)
        # Örnekte çerçeveler beyazdır; tablo hizasını korurken görünmez yap.
        tblPr=table._tbl.tblPr
        borders=OxmlElement('w:tblBorders')
        for edge in ('top','left','bottom','right','insideH','insideV'):
            e=OxmlElement('w:'+edge); e.set(qn('w:val'),'single'); e.set(qn('w:sz'),'4'); e.set(qn('w:color'),'FFFFFF')
            borders.append(e)
        tblPr.append(borders)
    except Exception:
        pass
    vals=[('Tarama Yapılan Görev Alanı:','Sanayi ve Teknoloji'),
          ('Tarih:',datetime.now().astimezone().strftime('%d.%m.%Y')),
          ('Bulgular:','Sanayi ve Teknoloji alanlarında yapılan açık kaynak taraması neticesinde,')]
    for row,(lab,val) in zip(table.rows,vals):
        for ci,textv in enumerate((lab,val)):
            p=row.cells[ci].paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.LEFT
            r=p.add_run(textv); _v115_set_run_font(r,12,bold=(ci==0))
    return table


def _v115_prepare_akt_row(row):
    detail=article_detail(row)
    real_url=str(detail.get('canonical') or row.get('Yayıncı_URL') or row.get('URL','') or '')
    title=_clean_note_text(detail.get('title') or row.get('Başlık',''))
    source=_clean_note_text(_real_source(row,detail,real_url))
    body=_clean_note_text(detail.get('text') or row.get('İçerik_Özeti') or title)
    summary=_akt_formal_summary(title,body,max_sentences=9,max_chars=2600)
    summary=_v67_akt_reported_content(summary)
    image=None
    for candidate in (detail.get('images') or [])[:8]:
        image=_download_report_image(candidate)
        if image: break
    return {'title':title,'source':source,'url':real_url,'summary':summary,'image':image}


def make_docx(rows):
    """V115 AKT: gönderilen kurumsal örneğin tablo + ayrıntılı içerik + görsel düzeni."""
    rows=_v115_dedupe_rows(rows or [])
    doc=_v115_doc_defaults(Document(),top=2.5,bottom=1.3,left=2.5,right=2.5)
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after=Pt(8)
    _v115_set_run_font(p.add_run('AÇIK KAYNAK TARAMA ÇALIŞMASI'),14,True)
    _v115_add_akt_info_table(doc,_akt_findings_intro(rows))

    intro=doc.add_paragraph(); intro.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    intro.paragraph_format.first_line_indent=Cm(0.4); intro.paragraph_format.space_after=Pt(0)
    _v115_set_run_font(intro.add_run(_akt_findings_intro(rows)),12)

    prepared=[None]*len(rows)
    if rows:
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(6,len(rows))) as ex:
            futs={ex.submit(_v115_prepare_akt_row,r):i for i,r in enumerate(rows)}
            for fut in concurrent.futures.as_completed(futs):
                try: prepared[futs[fut]]=fut.result()
                except Exception: prepared[futs[fut]]=None

    for i,item in enumerate(prepared,1):
        if not item: continue
        p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent=Cm(0.4); p.paragraph_format.space_before=Pt(4); p.paragraph_format.space_after=Pt(6)
        _v115_set_run_font(p.add_run(f'{i}. “{item["source"]}”'),12,True)
        _v115_set_run_font(p.add_run(' isimli internet sitesinde, '),12)
        _v115_set_run_font(p.add_run(f'“{item["title"]}”'),12,True)
        _v115_set_run_font(p.add_run(' başlığıyla bir haber yayımlanmıştır. ('),12)
        _word_hyperlink(p,item['url'],item['url'] or 'Haber Linki')
        _v115_set_run_font(p.add_run(') Söz konusu haber içeriğinde, '),12)
        summary=(item['summary'] or '').strip().rstrip(' .;')
        _v115_set_run_font(p.add_run(summary),12)
        _v115_set_run_font(p.add_run(' hususları ifade edilmiştir.'),12)

        if item.get('image'):
            cap=doc.add_paragraph(); cap.alignment=WD_ALIGN_PARAGRAPH.CENTER
            cap.paragraph_format.space_before=Pt(4); cap.paragraph_format.space_after=Pt(4)
            _v115_set_run_font(cap.add_run(f'Görsel {i}: “{item["source"]}” Sitesinde Yer Alan Görsel'),11,True)
            ip=doc.add_paragraph(); ip.alignment=WD_ALIGN_PARAGRAPH.CENTER; ip.paragraph_format.space_after=Pt(8)
            try: ip.add_run().add_picture(item['image'],width=Cm(14.5))
            except Exception: pass

    p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(8)
    _v115_set_run_font(p.add_run('Arz olunur.'),12)
    bio=BytesIO(); doc.save(bio); bio.seek(0); return bio.getvalue()


def _v115_prepare_note_item(row):
    detail=article_detail(row)
    title=_clean_note_text(detail.get('title') or row.get('Başlık',''))
    body=_clean_note_text(detail.get('text') or row.get('İçerik_Özeti') or title)
    facts=_v115_unique_fact_sentences(title,body,max_sentences=18,max_chars=6500)
    return row,detail,title,facts


def make_analyst_docx(df,title='BİLGİ NOTU'):
    """
    V115 Bilgi Notu: gönderilen örnekteki gibi başlığı tekrarlamayan,
    Times New Roman 12 punto, yalnız olgusal ve bütünlüklü paragraf akışı.
    Otomatik/genel değerlendirme cümlesi eklenmez.
    """
    doc=_v115_doc_defaults(Document(),2.5,2.5,2.5,2.5)
    x=df.copy() if df is not None else pd.DataFrame()
    if x.empty:
        rows=[]
    else:
        if 'Tarih_dt' in x.columns:
            x['Tarih_dt']=pd.to_datetime(x['Tarih_dt'],utc=True,errors='coerce')
            x=x.sort_values('Tarih_dt',ascending=True,na_position='last')
        rows=x.to_dict('records')
    try:
        er=_v107_enrich_selected_rows(rows)
        if er: rows=er
    except Exception:
        pass
    rows=_v115_dedupe_rows(rows)

    items=[None]*len(rows)
    if rows:
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(6,len(rows))) as ex:
            futs={ex.submit(_v115_prepare_note_item,r):i for i,r in enumerate(rows)}
            for fut in concurrent.futures.as_completed(futs):
                try: items[futs[fut]]=fut.result()
                except Exception: pass

    all_facts=[]; seen=[]
    for item in items:
        if not item: continue
        for s in item[3]:
            toks=set(_history_tokens(s)); dup=False
            for old in seen[-45:]:
                union=len(toks|old)
                if union and len(toks&old)/union>=0.80: dup=True; break
            if not dup:
                all_facts.append(s); seen.append(toks)

    if not all_facts:
        all_facts=['Seçilen habere ilişkin ayrıntılı içerik temin edilememiştir.']

    # Örnekteki yoğunluğa yakın: 4–5 doğal paragraf; her paragrafta 2–4 tam cümle.
    max_paras=5
    if len(all_facts)<=3:
        chunks=[all_facts]
    else:
        target=min(max_paras,max(2,(len(all_facts)+2)//3))
        size=max(2,(len(all_facts)+target-1)//target)
        chunks=[all_facts[i:i+size] for i in range(0,len(all_facts),size)][:max_paras]
        # Son kesim nedeniyle kalan cümle varsa son paragrafa ekle.
        used=sum(len(c) for c in chunks)
        if used<len(all_facts): chunks[-1].extend(all_facts[used:])

    for chunk in chunks:
        text=' '.join(_v115_formal_sentence(s) for s in chunk if _clean_note_text(s)).strip()
        if not text: continue
        p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after=Pt(0); p.paragraph_format.line_spacing=1.0
        _v115_set_run_font(p.add_run(text),12)

    bio=BytesIO(); doc.save(bio); bio.seek(0); return bio.getvalue()


def _v115_ogn_paragraph(title,source,body,summary):
    facts=_v115_unique_fact_sentences(title,body or summary,max_sentences=5,max_chars=900)
    if not facts:
        facts=[_clean_note_text(summary or title)]
    # İlk 2-4 bilgi taşıyan cümle. Başlığı tek başına paragraf olarak tekrar etme.
    text=' '.join(_v115_formal_sentence(s) for s in facts[:4] if s).strip()
    text=_v98_strip_site_name(text,source)
    return re.sub(r'\s+',' ',text).strip()


def make_important_basket_docx_v101(basket_df):
    """V115 ÖGN: kısa ancak gerçek haber içeriğine dayalı, başlık tekrarı olmayan kurumsal özet."""
    doc=_v115_doc_defaults(Document(),2.0,2.0,2.5,2.5)
    now=datetime.now().astimezone()
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    _v115_set_run_font(p.add_run(now.strftime('%d/%m/%Y')),12)
    p=doc.add_paragraph(); _v115_set_run_font(p.add_run('Konu: '),12,True)
    _v115_set_run_font(p.add_run('STB Temsilciliği Önemli Gelişmeler Notu'),12)

    raw=[] if basket_df is None else basket_df.to_dict('records')
    rows=_v115_dedupe_rows(raw)
    outputs=[None]*len(rows)
    def work(r):
        title=_clean_note_text(r.get('title','')); source=_clean_note_text(r.get('source',''))
        summary=_clean_note_text(r.get('summary','')); url=str(r.get('url','') or '')
        row={'Başlık':title,'Kaynak':source,'URL':url,'Yayıncı_URL':url,'İçerik_Özeti':summary,'Tarih':r.get('news_time','')}
        detail=article_detail(row)
        body=_clean_note_text(detail.get('text') or summary)
        return _v115_ogn_paragraph(title,source,body,summary)
    if rows:
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(6,len(rows))) as ex:
            futs={ex.submit(work,r):i for i,r in enumerate(rows)}
            for fut in concurrent.futures.as_completed(futs):
                try: outputs[futs[fut]]=fut.result()
                except Exception: pass
    for text in outputs:
        text=_clean_note_text(text)
        if not text: continue
        p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing=1.0; p.paragraph_format.space_after=Pt(7)
        if text.endswith('.'): text=text[:-1]
        _v115_set_run_font(p.add_run(text+' (STB).'),12)
    p=doc.add_paragraph(); _v115_set_run_font(p.add_run('Arz olunur.'),12)
    bio=BytesIO(); doc.save(bio); bio.seek(0); return bio.getvalue()

# ============================================================
# /V115
# ============================================================


# ============================================================
# V116 — REFERANS RAPOR ŞABLONLARI + GÜVENİLİR HABER AYRINTISI
# Kullanıcının ilettiği üç REFERANS Word belgesi esas alınmıştır:
#   1) STB -AKT Raporu
#   2) STB-Önemli Gelişmeler Notu
#   3) Bilgi Notu Taslak
# Amaç: raporları yalnız benzer başlıklarla değil; içerik yoğunluğu, sayfa düzeni,
# paragraf yapısı, gerçek yayıncı URL'si ve görsel kullanımıyla referansa yaklaştırmak.
# ============================================================

V116_REPORT_ENGINE_VERSION='V116_REFERENCE_DOCS_20260921'
if st.session_state.get('_v116_report_engine_version') != V116_REPORT_ENGINE_VERSION:
    st.session_state['_v116_report_engine_version']=V116_REPORT_ENGINE_VERSION
    # Önceki sürümde üretilmiş bytes ekranda kalmasın.
    for _k in (
        'docx_bytes','note_bytes','v90_ogn_docx_bytes','basket_docx_bytes',
        'v78_ogn_note_bytes','v79_akt_note_bytes','v81_pres_note_bytes'
    ):
        st.session_state.pop(_k,None)
    st.session_state.pop('_v116_article_cache',None)


def _v116_google_news_url(u):
    try:
        h=urlparse(str(u or '')).netloc.lower()
        p=urlparse(str(u or '')).path.lower()
        return ('news.google.com' in h and ('/rss/articles/' in p or '/articles/' in p or '/read/' in p))
    except Exception:
        return False


def _v116_valid_direct_url(u):
    try:
        u=str(u or '').strip()
        if not u.startswith(('http://','https://')): return False
        host=urlparse(u).netloc.lower().replace('www.','')
        if not host or 'google.com' in host: return False
        return True
    except Exception:
        return False


def _v116_google_article_id(u):
    try:
        path=urlparse(str(u or '')).path.rstrip('/')
        parts=[x for x in path.split('/') if x]
        for marker in ('articles','read'):
            if marker in parts:
                i=parts.index(marker)
                if i+1<len(parts): return parts[i+1]
        return parts[-1] if parts else ''
    except Exception:
        return ''


def _v116_decode_google_news_url(u):
    """Google News RSS URL'sini ek paket olmadan gerçek yayıncı URL'sine çözer.
    2024+ Google News bağlantıları opak kimlik kullandığı için sayfadaki signature/timestamp
    alınır ve Fbv4je batchexecute çağrısı yapılır.
    """
    u=str(u or '').strip()
    if not _v116_google_news_url(u):
        return u if _v116_valid_direct_url(u) else ''

    cache=st.session_state.setdefault('_v116_gnews_decode_cache',{})
    if u in cache:
        return cache[u]

    data_id=_v116_google_article_id(u)
    if not data_id:
        cache[u]=''; return ''

    headers={
        **HEADERS,
        'Accept-Language':'tr-TR,tr;q=0.9,en;q=0.6',
        'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    }
    decoded=''
    try:
        sess=requests.Session()
        r=sess.get(u,headers=headers,timeout=12,allow_redirects=True)
        if r.ok and _v116_valid_direct_url(r.url):
            decoded=r.url
        elif r.ok and r.text:
            soup=BeautifulSoup(r.text,'html.parser')
            node=soup.find(attrs={'data-n-a-id':data_id})
            if node is None:
                node=soup.find(attrs={'data-n-a-sg':True,'data-n-a-ts':True})
            signature=str(node.get('data-n-a-sg') or '').strip() if node else ''
            timestamp=str(node.get('data-n-a-ts') or '').strip() if node else ''
            node_id=str(node.get('data-n-a-id') or data_id).strip() if node else data_id
            if signature and timestamp and node_id:
                from urllib.parse import quote
                inner=(
                    '["garturlreq",[["X","X",["X","X"],null,null,1,1,"TR:tr",null,1,'
                    'null,null,null,null,null,0,1],"tr","TR",1,[1,1,1],1,1,null,0,0,null,0],'
                    f'"{node_id}",{timestamp},"{signature}"]'
                )
                payload=['Fbv4je',inner]
                data='f.req='+quote(json.dumps([[payload]],ensure_ascii=False,separators=(',',':')))
                rr=sess.post(
                    'https://news.google.com/_/DotsSplashUi/data/batchexecute',
                    data=data,
                    headers={
                        **HEADERS,
                        'Content-Type':'application/x-www-form-urlencoded;charset=UTF-8',
                        'Referer':'https://news.google.com/',
                        'Accept-Language':'tr-TR,tr;q=0.9,en;q=0.6',
                    },
                    timeout=12
                )
                if rr.ok and rr.text:
                    raw=rr.text.lstrip("\ufeff \r\n")
                    if raw.startswith(")]}'"):
                        raw=raw.split('\n',1)[1] if '\n' in raw else ''
                    try:
                        outer=json.loads(raw)
                    except Exception:
                        # Bazı yanıtlar ilk satırda uzunluk değeri içerir.
                        raw2='\n'.join(x for x in raw.splitlines() if x.lstrip().startswith('['))
                        outer=json.loads(raw2) if raw2 else []

                    def _walk(obj):
                        if isinstance(obj,list):
                            if len(obj)>=3 and obj[1]=='Fbv4je' and isinstance(obj[2],str):
                                try:
                                    inside=json.loads(obj[2])
                                    if isinstance(inside,list) and len(inside)>1 and inside[0]=='garturlres':
                                        cand=str(inside[1] or '')
                                        if _v116_valid_direct_url(cand): return cand
                                except Exception:
                                    pass
                            for z in obj:
                                hit=_walk(z)
                                if hit: return hit
                        elif isinstance(obj,dict):
                            for z in obj.values():
                                hit=_walk(z)
                                if hit: return hit
                        return ''
                    decoded=_walk(outer)
    except Exception:
        decoded=''

    cache[u]=decoded if _v116_valid_direct_url(decoded) else ''
    return cache[u]


def _v116_clean_headline(title,source=''):
    t=_clean_note_text(title).strip(' -–—|')
    src=_clean_note_text(source).strip()
    variants=[]
    if src:
        variants.extend([src,src.replace(' Gazetesi','').strip()])
    # Kaynak alanında domain varsa stem'i de kullan.
    try:
        if '.' in src:
            variants.append(src.split('.')[0])
    except Exception:
        pass
    for v in sorted({x for x in variants if len(x)>=3},key=len,reverse=True):
        t=re.sub(r'\s*(?:-|–|—|\|)\s*'+re.escape(v)+r'\s*$', '', t, flags=re.I).strip()
    return re.sub(r'\s+',' ',t).strip()


def _v116_detail_score(d):
    if not d: return -1
    txt=_clean_note_text(d.get('text') or '')
    score=min(len(txt),12000)
    if _v116_valid_direct_url(d.get('canonical')): score+=1200
    if d.get('images'): score+=min(len(d.get('images') or []),3)*250
    if len(_akt_clean_sentences(d.get('title',''),txt))>=3: score+=800
    return score


# V115 article_detail'i son fallback olarak sakla.
_v116_article_detail_fallback=article_detail


def article_detail(row):
    """V116: önce Google News kimliğini gerçek yayıncı URL'sine çözer, sonra haber sayfasını okur.
    Çözülemezse V115'in GDELT/DDG/Bing fallback zincirini kullanır. Sonuç oturumda cache'lenir.
    """
    if isinstance(row,str): row={'URL':row}
    elif hasattr(row,'to_dict'): row=row.to_dict()
    elif row is None: row={}
    else: row=dict(row)

    original=str(row.get('URL') or row.get('url') or '').strip()
    raw_title=_clean_note_text(row.get('Başlık') or row.get('title') or '')
    source=_clean_note_text(row.get('Kaynak') or row.get('source') or row.get('Yayıncı') or '')
    title=_v116_clean_headline(raw_title,source)
    cache=st.session_state.setdefault('_v116_article_cache',{})
    ck=hashlib.sha1((original+'|'+title+'|'+source).encode('utf-8','ignore')).hexdigest()
    if ck in cache:
        return dict(cache[ck])

    candidates=[]
    direct=''
    if _v116_google_news_url(original):
        direct=_v116_decode_google_news_url(original)
    elif _v116_valid_direct_url(original):
        direct=original

    # source_url gerçek haber URL'siyse onu da dene.
    pub=str(row.get('Yayıncı_URL') or '').strip()
    if not direct and _v116_valid_direct_url(pub):
        # Ana sayfa ise haber URL'si olarak kabul etme.
        if len(urlparse(pub).path.strip('/'))>=4:
            direct=pub

    if direct:
        rr=dict(row); rr['URL']=direct; rr['Yayıncı_URL']=direct
        try:
            d=_v115_article_detail_base(rr) or {}
            candidates.append(d)
        except Exception:
            pass

    # V115 zinciri: paket/redirect/GDELT/DDG/Bing ve mevcut fallback'ler.
    try:
        d=_v116_article_detail_fallback(row) or {}
        candidates.append(d)
    except Exception:
        pass

    # Başlıkla doğrudan arama: kaynak alan adı bilinirse önceliklendir.
    if title:
        preferred=''
        try:
            preferred=urlparse(pub).netloc.lower().replace('www.','') if pub else ''
        except Exception:
            preferred=''
        try:
            found=_v115_search_direct_article(title,preferred)
        except Exception:
            found=''
        if found and _v116_valid_direct_url(found) and all(str(x.get('canonical',''))!=found for x in candidates):
            rr=dict(row); rr['URL']=found; rr['Yayıncı_URL']=found
            try:
                candidates.append(_v115_article_detail_base(rr) or {})
            except Exception:
                pass

    if candidates:
        best=max(candidates,key=_v116_detail_score)
    else:
        best={
            'title':title or raw_title,
            'canonical':original,
            'published':str(row.get('Tarih') or ''),
            'text':_clean_note_text(row.get('İçerik_Özeti') or row.get('summary') or title),
            'images':[],
            'source':source,
        }

    best=dict(best)
    best['source']=_clean_note_text(best.get('source') or source or 'Açık Kaynak')
    best['title']=_v116_clean_headline(best.get('title') or title or raw_title,best['source'])
    # RSS başlığının aynısı, içerik sayılmaz; mümkünse mevcut özet ile değiştir.
    body=_clean_note_text(best.get('text') or '')
    if norm(body)==norm(raw_title) or norm(body)==norm(best['title']):
        fallback=_clean_note_text(row.get('İçerik_Özeti') or row.get('summary') or '')
        if fallback and norm(fallback) not in {norm(raw_title),norm(best['title'])}:
            best['text']=fallback
    cache[ck]=dict(best)
    return best


def _v116_topic_labels(rows):
    joined=norm(' '.join(f"{r.get('Başlık',r.get('title',''))} {r.get('İçerik_Özeti',r.get('summary',''))} {r.get('Kategori',r.get('category',''))}" for r in (rows or [])))
    mapping=[
        (('veri sızınt','veri ihlal'),'Veri Sızıntıları'),
        (('yapay zeka','yapay zekâ','gemini','openai'),'Yapay Zeka (YZ)'),
        (('siber saldır','siber güvenlik','casus yazılım'),'Siber Güvenlik'),
        (('elektrikli araç','şarj istasyon','şarj nokt'),'Elektrikli Araçlar ve Şarj Altyapısı'),
        (('ihracat',),'İhracat'),
        (('sanayi üret','imalat sanayi'),'Sanayi Üretimi'),
        (('savunma','aselsan','tusaş','roketsan','baykar'),'Savunma Sanayii'),
        (('uzay','havacılık','uydu'),'Havacılık ve Uzay'),
        (('yatırım','fabrika','tesis','üretim hatt'),'Yatırım ve Üretim'),
        (('çip','yarı iletken'),'Yarı İletkenler'),
        (('enerji','hidrojen','güneş','rüzgar'),'Enerji Teknolojileri'),
        (('ar-ge','arge','inovasyon'),'Ar-Ge ve İnovasyon'),
    ]
    out=[]
    for keys,label in mapping:
        if any(k in joined for k in keys): out.append(label)
        if len(out)>=4: break
    if not out:
        cats=[]
        for r in rows or []:
            c=_clean_note_text(r.get('Kategori',r.get('category','')))
            if c and c not in cats: cats.append(c)
        out=cats[:3]
    return out


def _v116_akt_intro_text(rows):
    topics=_v116_topic_labels(rows)
    if topics:
        if len(topics)==1: tt=f'“{topics[0]}”'
        else: tt=', '.join(f'“{x}”' for x in topics[:-1])+f' ve “{topics[-1]}”'
        return (f'Sanayi ve Teknoloji alanlarında yapılan açık kaynak taraması neticesinde bazı haber bültenlerinde {tt} '
                'konu başlıklarıyla ilgili olmak üzere içerikler hazırlandığı tespit edilmiştir. İçeriklerin hangi internet '
                'sitesinde yer aldığı, başlığı, bağlantı adresi, içeriğin detaylı özeti ve görseli aşağıda yer almaktadır.')
    return ('Sanayi ve Teknoloji alanlarında yapılan açık kaynak taraması neticesinde seçilen içerikler tespit edilmiştir. '
            'İçeriklerin hangi internet sitesinde yer aldığı, başlığı, bağlantı adresi, içeriğin detaylı özeti ve görseli aşağıda yer almaktadır.')


def _v116_doc_defaults(doc,top=2.5,bottom=2.5,left=2.5,right=2.5):
    sec=doc.sections[0]
    sec.top_margin=Cm(top); sec.bottom_margin=Cm(bottom)
    sec.left_margin=Cm(left); sec.right_margin=Cm(right)
    normal=doc.styles['Normal']
    normal.font.name='Times New Roman'; normal.font.size=Pt(12)
    normal._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'Times New Roman')
    return doc


def _v116_run(run,size=12,bold=None,italic=None):
    run.font.name='Times New Roman'; run.font.size=Pt(size)
    run._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'Times New Roman')
    if bold is not None: run.bold=bool(bold)
    if italic is not None: run.italic=bool(italic)
    return run


def _v116_hide_table_borders(table):
    try:
        tblPr=table._tbl.tblPr
        old=tblPr.find(qn('w:tblBorders'))
        if old is not None: tblPr.remove(old)
        borders=OxmlElement('w:tblBorders')
        for edge in ('top','left','bottom','right','insideH','insideV'):
            e=OxmlElement('w:'+edge); e.set(qn('w:val'),'nil'); borders.append(e)
        tblPr.append(borders)
    except Exception:
        pass


def _v116_add_akt_info_table(doc):
    table=doc.add_table(rows=3,cols=2)
    _v116_hide_table_borders(table)
    try:
        table.autofit=False
        table.columns[0].width=Cm(5.2); table.columns[1].width=Cm(10.8)
    except Exception: pass
    vals=[('Tarama Yapılan Görev Alanı:','Sanayi ve Teknoloji'),
          ('Tarih:',datetime.now().astimezone().strftime('%d.%m.%Y')),
          ('Bulgular:','Sanayi ve Teknoloji alanlarında yapılan açık kaynak taraması neticesinde,')]
    for row,(lab,val) in zip(table.rows,vals):
        for c in row.cells:
            c.margin_top=Cm(0); c.margin_bottom=Cm(0)
        p=row.cells[0].paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0)
        _v116_run(p.add_run(lab),12,True)
        p=row.cells[1].paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0)
        _v116_run(p.add_run(val),12)
    return table


def _v116_prepare_akt_row(row):
    detail=article_detail(row)
    source=_clean_note_text(_real_source(row,detail,detail.get('canonical','')))
    title=_v116_clean_headline(detail.get('title') or row.get('Başlık',''),source)
    body=_clean_note_text(detail.get('text') or row.get('İçerik_Özeti') or '')
    summary=_akt_formal_summary(title,body,max_sentences=9,max_chars=2300)
    summary=_v67_akt_reported_content(summary).strip(' .;')
    real_url=str(detail.get('canonical') or '').strip()
    if not _v116_valid_direct_url(real_url):
        # Son çare: yayıncı URL'si haber yoluysa kullan; aksi halde kullanıcı linki kaybolmasın.
        pu=str(row.get('Yayıncı_URL') or '').strip()
        real_url=pu if _v116_valid_direct_url(pu) and len(urlparse(pu).path.strip('/'))>=4 else str(row.get('URL') or '')
    image=None
    image_urls=list(detail.get('images') or [])
    if row.get('Görsel_URL'): image_urls.insert(0,str(row.get('Görsel_URL')))
    for candidate in image_urls[:12]:
        image=_download_report_image(candidate)
        if image: break
    return {'title':title,'source':source,'url':real_url,'summary':summary,'image':image,'body_len':len(body)}


def make_docx(rows):
    """V116 AKT — referans belgenin görünüm ve içerik mantığına göre."""
    rows=_v115_dedupe_rows(rows or [])
    doc=_v116_doc_defaults(Document(),top=2.5,bottom=1.25,left=2.5,right=2.5)
    sec=doc.sections[0]; sec.header_distance=Cm(1.25); sec.footer_distance=Cm(1.25)

    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after=Pt(0)
    _v116_run(p.add_run('AÇIK KAYNAK TARAMA ÇALIŞMASI'),16,True)
    _v116_add_akt_info_table(doc)

    intro=doc.add_paragraph(); intro.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    intro.paragraph_format.first_line_indent=Cm(0.63); intro.paragraph_format.line_spacing=1.5
    intro.paragraph_format.space_before=Pt(0); intro.paragraph_format.space_after=Pt(0)
    intro_text=_v116_akt_intro_text(rows)
    topics=_v116_topic_labels(rows)
    # Konu adlarını referanstaki gibi italik yaz.
    pos=0
    spans=[]
    for topic in topics:
        for q in (f'“{topic}”',topic):
            ix=intro_text.find(q,pos)
            if ix>=0:
                spans.append((ix,ix+len(q))); pos=ix+len(q); break
    cursor=0
    for a,b in sorted(spans):
        if a>cursor: _v116_run(intro.add_run(intro_text[cursor:a]),12)
        _v116_run(intro.add_run(intro_text[a:b]),12,italic=True)
        cursor=b
    if cursor<len(intro_text): _v116_run(intro.add_run(intro_text[cursor:]),12)

    prepared=[None]*len(rows)
    if rows:
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(3,len(rows))) as ex:
            futs={ex.submit(_v116_prepare_akt_row,r):i for i,r in enumerate(rows)}
            for fut in concurrent.futures.as_completed(futs):
                try: prepared[futs[fut]]=fut.result()
                except Exception: prepared[futs[fut]]=None

    for i,item in enumerate(prepared,1):
        if not item: continue
        p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent=Cm(0.63); p.paragraph_format.line_spacing=1.5
        p.paragraph_format.space_before=Pt(4); p.paragraph_format.space_after=Pt(6)
        _v116_run(p.add_run(f'{i}. “{item["source"]}”'),12,True)
        _v116_run(p.add_run(' isimli internet sitesinde, '),12)
        _v116_run(p.add_run(f'“{item["title"]}”'),12,True,True)
        _v116_run(p.add_run(' başlığıyla bir haber yayımlanmıştır. ('),12)
        _word_hyperlink(p,item['url'],item['url'] or 'Haber bağlantısı')
        _v116_run(p.add_run(') Söz konusu haber içeriğinde, '),12)
        summary=(item.get('summary') or '').strip().rstrip(' .;')
        _v116_run(p.add_run(summary),12)
        _v116_run(p.add_run(' hususları ifade edilmiştir.'),12)

        if item.get('image'):
            cap=doc.add_paragraph(); cap.alignment=WD_ALIGN_PARAGRAPH.CENTER
            cap.paragraph_format.first_line_indent=Cm(0.63); cap.paragraph_format.line_spacing=1.5
            cap.paragraph_format.space_before=Pt(4); cap.paragraph_format.space_after=Pt(4)
            _v116_run(cap.add_run(f'Görsel {i}: “{item["source"]}” Sitesinde Yer Alan Görsel'),12,True)
            ip=doc.add_paragraph(); ip.alignment=WD_ALIGN_PARAGRAPH.CENTER
            ip.paragraph_format.space_after=Pt(8)
            try:
                run=ip.add_run(); shape=run.add_picture(item['image'])
                max_w=Cm(15.3); max_h=Cm(12.7)
                scale=min(1.0, max_w/shape.width, max_h/shape.height)
                if scale<1.0:
                    shape.width=int(shape.width*scale); shape.height=int(shape.height*scale)
            except Exception:
                pass

    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing=1.5
    _v116_run(p.add_run('Arz olunur.'),12)
    bio=BytesIO(); doc.save(bio); bio.seek(0); return bio.getvalue()


def _v116_select_fact_sentences(title,body,limit=15,max_chars=6500):
    sents=_akt_clean_sentences(title,body)
    if not sents:
        fb=_clean_note_text(body or '')
        return [fb] if fb and norm(fb)!=norm(title) else []
    ranked=[]
    for i,s in enumerate(sents):
        score=_akt_sentence_score(s)+_sent_score(s)
        if i<3: score+=7
        if re.search(r'\b\d+(?:[.,]\d+)?\b',s): score+=4
        if any(k in norm(s) for k in ['buna göre','toplam','oran','yüzde','aynı dönemde','saat','tarih','açıklam','duyur','bildir']): score+=2
        ranked.append((score,i))
    keep=set(range(min(2,len(sents))))
    # Metnin farklı bölgelerinden bağlam koru.
    if len(sents)>=6:
        keep.update({len(sents)//3,(2*len(sents))//3,len(sents)-1})
    for _,i in sorted(ranked,reverse=True):
        if len(keep)>=limit: break
        keep.add(i)
    out=[]; total=0; seen=[]
    for i in sorted(keep):
        s=_v115_formal_sentence(sents[i])
        toks=set(_history_tokens(s)); dup=False
        for old in seen:
            u=len(toks|old)
            if u and len(toks&old)/u>=0.82: dup=True; break
        if dup: continue
        if out and total+len(s)>max_chars: break
        out.append(s); seen.append(toks); total+=len(s)+1
    return out


def make_analyst_docx(df,title='BİLGİ NOTU'):
    """V116 Bilgi Notu — referans taslak gibi yalnız içerik; başlık/tarih/kaynak/Arz olunur eklenmez."""
    doc=_v116_doc_defaults(Document(),2.5,2.5,2.5,2.5)
    sec=doc.sections[0]; sec.header_distance=Cm(1.25); sec.footer_distance=Cm(1.25)
    x=df.copy() if df is not None else pd.DataFrame()
    if x.empty:
        rows=[]
    else:
        if 'Tarih_dt' in x.columns:
            x['Tarih_dt']=pd.to_datetime(x['Tarih_dt'],utc=True,errors='coerce')
            x=x.sort_values('Tarih_dt',ascending=True,na_position='last')
        rows=x.to_dict('records')
    try:
        er=_v107_enrich_selected_rows(rows)
        if er: rows=er
    except Exception: pass
    rows=_v115_dedupe_rows(rows)

    facts=[]
    for row in rows:
        detail=article_detail(row)
        source=_clean_note_text(_real_source(row,detail,detail.get('canonical','')))
        ttl=_v116_clean_headline(detail.get('title') or row.get('Başlık',''),source)
        body=_clean_note_text(detail.get('text') or row.get('İçerik_Özeti') or '')
        facts.extend(_v116_select_fact_sentences(ttl,body,limit=15,max_chars=6500))

    # Çapraz tekrarları temizle.
    uniq=[]; seen=[]
    for s in facts:
        toks=set(_history_tokens(s)); dup=False
        for old in seen:
            u=len(toks|old)
            if u and len(toks&old)/u>=0.82: dup=True; break
        if not dup: uniq.append(s); seen.append(toks)
    facts=uniq[:15]

    if not facts:
        # Kötü bir başlığı "bilgi notu" diye tekrarlamak yerine erişim sorunu açıkça belirtilir.
        facts=['Seçilen habere ilişkin ayrıntılı haber metnine erişilemediğinden bilgi notu içeriği oluşturulamamıştır.']

    # Referanstaki gibi 4-5 paragraf. Paragraf sınırları cümle sırasını bozmadan oluşturulur.
    target=(5 if len(facts)>=8 else (4 if len(facts)>=6 else (3 if len(facts)>=4 else (2 if len(facts)>=2 else 1))))
    base=max(1,math.ceil(len(facts)/target))
    chunks=[]
    for i in range(0,len(facts),base): chunks.append(facts[i:i+base])
    if len(chunks)>5:
        tail=[]
        for ch in chunks[4:]: tail.extend(ch)
        chunks=chunks[:4]+[tail]

    for ch in chunks:
        text=' '.join(s for s in ch if s).strip()
        if not text: continue
        p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0)
        p.paragraph_format.line_spacing=1.0
        _v116_run(p.add_run(text),12)

    bio=BytesIO(); doc.save(bio); bio.seek(0); return bio.getvalue()


def _v116_ogn_item(row):
    title=_clean_note_text(row.get('title','')); source=_clean_note_text(row.get('source',''))
    summary=_clean_note_text(row.get('summary','')); url=str(row.get('url','') or '')
    rr={'Başlık':title,'Kaynak':source,'URL':url,'İçerik_Özeti':summary,'Tarih':row.get('news_time','')}
    detail=article_detail(rr)
    source2=_clean_note_text(_real_source(rr,detail,detail.get('canonical','')) or source)
    ttl=_v116_clean_headline(detail.get('title') or title,source2)
    body=_clean_note_text(detail.get('text') or summary)
    sents=_v116_select_fact_sentences(ttl,body,limit=5,max_chars=1500)
    if not sents:
        # Özet başlıktan farklıysa onu kullan.
        if summary and norm(summary)!=norm(title): sents=[_v115_formal_sentence(summary)]
        else: sents=[]
    # Referans ÖGN'de gelişmeler genellikle 1-3 cümle / tek paragraftır.
    chosen=[]; total=0
    for s in sents:
        if chosen and (len(chosen)>=3 or total+len(s)>620): break
        chosen.append(s); total+=len(s)+1
    text=' '.join(chosen).strip()
    text=_v98_strip_site_name(text,source2)
    return re.sub(r'\s+',' ',text).strip()


def _v116_add_page_field(paragraph):
    run=paragraph.add_run()
    fldChar=OxmlElement('w:fldChar'); fldChar.set(qn('w:fldCharType'),'begin')
    instr=OxmlElement('w:instrText'); instr.set(qn('xml:space'),'preserve'); instr.text=' PAGE '
    sep=OxmlElement('w:fldChar'); sep.set(qn('w:fldCharType'),'separate')
    text=OxmlElement('w:t'); text.text='1'
    end=OxmlElement('w:fldChar'); end.set(qn('w:fldCharType'),'end')
    run._r.extend([fldChar,instr,sep,text,end])
    _v116_run(run,8)


def _v116_ogn_footer(section):
    footer=section.footer
    # Mevcut boş paragrafı kullan.
    p=footer.paragraphs[0]
    p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0)
    # Referanstaki ince mavi çizgi.
    pPr=p._p.get_or_add_pPr(); pBdr=OxmlElement('w:pBdr'); top=OxmlElement('w:top')
    top.set(qn('w:val'),'single'); top.set(qn('w:sz'),'4'); top.set(qn('w:space'),'1'); top.set(qn('w:color'),'5B9BD5')
    pBdr.append(top); pPr.append(pBdr)
    _v116_run(p.add_run('DEVLET BİLGİ KOORDİNASYON MERKEZİ'),8)
    p2=footer.add_paragraph(); p2.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    p2.paragraph_format.space_before=Pt(0); p2.paragraph_format.space_after=Pt(0)
    _v116_add_page_field(p2)


def make_important_basket_docx_v101(basket_df):
    """V116 ÖGN — referans STB Önemli Gelişmeler Notu sayfa düzeni ve özet yoğunluğu."""
    doc=_v116_doc_defaults(Document(),top=2.25,bottom=1.50,left=1.905,right=1.905)
    sec=doc.sections[0]; sec.header_distance=Cm(0.1); sec.footer_distance=Cm(0.45)
    # No Spacing stilini referans gibi ayarla.
    try:
        ns=doc.styles['No Spacing']; ns.font.name='Times New Roman'; ns.font.size=Pt(12)
        ns._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'Times New Roman')
    except Exception: pass
    _v116_ogn_footer(sec)

    today=datetime.now().astimezone().date(); yesterday=today-timedelta(days=1)
    p=doc.add_paragraph(style='No Spacing'); p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.line_spacing=1.15; p.paragraph_format.space_before=Pt(6); p.paragraph_format.space_after=Pt(6)
    _v116_run(p.add_run(f'{yesterday.strftime("%d/%m/%Y")} – {today.strftime("%d/%m/%Y")}'),12)
    p=doc.add_paragraph(style='No Spacing'); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing=1.15; p.paragraph_format.space_before=Pt(6); p.paragraph_format.space_after=Pt(6)
    _v116_run(p.add_run('Konu: '),12,True); _v116_run(p.add_run('STB Temsilciliği Önemli Gelişmeler Notu'),12)

    rows=_v115_dedupe_rows([] if basket_df is None else basket_df.to_dict('records'))
    outputs=[None]*len(rows)
    if rows:
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(3,len(rows))) as ex:
            futs={ex.submit(_v116_ogn_item,r):i for i,r in enumerate(rows)}
            for fut in concurrent.futures.as_completed(futs):
                try: outputs[futs[fut]]=fut.result()
                except Exception: outputs[futs[fut]]=''

    for text in outputs:
        text=_clean_note_text(text)
        if not text: continue
        p=doc.add_paragraph(style='No Spacing'); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing=1.15; p.paragraph_format.space_before=Pt(6); p.paragraph_format.space_after=Pt(6)
        text=text.rstrip()
        if text.endswith('.'): text=text[:-1]
        _v116_run(p.add_run(text+' (STB).'),12)

    p=doc.add_paragraph(style='No Spacing'); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing=1.15; p.paragraph_format.space_before=Pt(6); p.paragraph_format.space_after=Pt(6)
    _v116_run(p.add_run('Arz olunur.'),12)
    bio=BytesIO(); doc.save(bio); bio.seek(0); return bio.getvalue()

# ============================================================
# /V116
# ============================================================

# ============================================================
# V117 — RAPOR İÇERİK MOTORU / REFERANS BELGEYE YAKINLAŞTIRMA
# 1) AKT: sayfa içi ara başlık, ilişkili haber, menü ve ALL-CAPS parçaları temizlenir.
# 2) AKT: yalnız seçilen haberle ilişkili, bilgi taşıyan cümleler rapora alınır.
# 3) ÖGN: basket satırı mevcut taramadaki aynı olayın diğer kaynaklarıyla yeniden zenginleştirilir;
#    başlık tek başına çıktı yapılmaz, mümkünse gerçek haber/özet içeriğinden 1-3 cümle kurulur.
# 4) Bilgi Notu: giriş-gelişme-sonuç mantığında 3-5 doğal paragraf; soru-cevap ve ham haber dili yoktur.
# 5) Görsel: yalnız haber sayfasının güçlü aday görseli kullanılır; ikon/logo/küçük görsel alınmaz.
# ============================================================

V117_REPORT_ENGINE_VERSION='V117_CONTENT_ENGINE_20260921'
if st.session_state.get('_v117_report_engine_version') != V117_REPORT_ENGINE_VERSION:
    st.session_state['_v117_report_engine_version']=V117_REPORT_ENGINE_VERSION
    for _k in (
        'docx_bytes','note_bytes','v90_ogn_docx_bytes','basket_docx_bytes',
        'v78_ogn_note_bytes','v79_akt_note_bytes','v81_pres_note_bytes'
    ):
        st.session_state.pop(_k,None)
    st.session_state.pop('_v117_article_cache',None)
    st.session_state.pop('_v117_page_cache',None)


def _v117_source_short(source,url=''):
    """Site adını referans rapordaki gibi kısa ve kurumsal gösterir."""
    s=_clean_note_text(source).strip(' -–—|')
    generic={'google haberler','google news','google','rss','açık kaynak'}
    if norm(s) in generic or not s:
        try:
            host=urlparse(str(url or '')).netloc.lower().replace('www.','')
            if host and 'google.com' not in host:
                s=host
        except Exception:
            pass
    # Site adından sonra gelen slogan / SEO uzantıları.
    for sep in (' | ',' – ',' — ',' - '):
        if sep in s:
            left=s.split(sep,1)[0].strip()
            right=s.split(sep,1)[1].strip()
            if 3<=len(left)<=55 and any(k in norm(right) for k in (
                'haber','son dakika','gündem','hava durumu','ekonomi','teknoloji','haberleri'
            )):
                s=left
                break
    # Çok uzun site adı hâlâ kaldıysa ilk anlamlı parçayı kullan.
    if len(s)>70:
        m=re.split(r'\s+(?:haberleri|son dakika|gündem|haber|gazetesi haberleri)\b',s,1,flags=re.I)
        if m and 3<=len(m[0].strip())<=60:
            s=m[0].strip()
    return re.sub(r'\s+',' ',s).strip() or 'Açık Kaynak'


def _v117_is_upper_token(token):
    letters=''.join(ch for ch in str(token or '') if ch.isalpha())
    return bool(letters) and all(ch.upper()==ch and ch.lower()!=ch for ch in letters)


def _v117_strip_uppercase_runs(text):
    """Haber sayfasından gelen '50 BİN TELEFON...' türü ara başlıkları gövdeden söker."""
    words=_clean_note_text(text).split()
    out=[]; i=0
    while i<len(words):
        j=i; run=[]; alpha=0; uppercase_words=0
        while j<len(words):
            w=words[j]
            upper=_v117_is_upper_token(w)
            numeric=bool(re.fullmatch(r'[\d%.,:+\-–—/]+',w))
            if not (upper or numeric):
                break
            run.append(w)
            if upper:
                uppercase_words+=1
                alpha+=sum(ch.isalpha() for ch in w)
            j+=1
        if uppercase_words>=4 and alpha>=18:
            i=j
            continue
        out.append(words[i]); i+=1
    return ' '.join(out)


_V117_NOISE_TERMS=(
    'çerez','cookie','abonelik','abone ol','reklam','tüm hakları saklıdır','gizlilik politikası',
    'kullanım koşulları','bildirimleri aç','uygulamamızı indirin','facebook','instagram','whatsapp',
    'twitter','x.com','son dakika haberleri için','haberlerimizi takip','ilgili haberler','öne çıkan haberler',
    'diğer haberler','çok okunanlar','en çok okunan','etiketler','yorumlar','foto galeri','video galeri',
    'anasayfa','ana sayfa','reklamı geç','devamını oku','tıklayınız','kaynakça','paylaş','yazarın diğer yazıları'
)


def _v117_tokens(text):
    stop={'ve','ile','için','bir','bu','şu','ise','olarak','olan','olduğu','daha','çok','son','yeni','ilk',
          'haber','haberde','tarafından','göre','ancak','ayrıca','öte','yandan','ilgili','söz','konusu'}
    return {x for x in re.findall(r'[a-z0-9çğıöşü]+',norm(text)) if len(x)>=3 and x not in stop}


def _v117_heading_like(s):
    s=_clean_note_text(s)
    if not s: return True
    letters=''.join(ch for ch in s if ch.isalpha())
    if letters and len(s)<180:
        ratio=sum(ch.isupper() for ch in letters)/max(1,len(letters))
        if ratio>=0.68: return True
    # Sayfa içi bağlantı başlıkları çoğunlukla kısa ve ':' ile biter.
    if s.rstrip().endswith(':') and len(s.split())<=18:
        return True
    # Cümle değil, peş peşe haber başlıkları görünümü.
    if len(s.split())<=10 and s[-1:] not in '.!?' and not re.search(
        r'\b(?:oldu|olmuştur|edildi|edilmiştir|açıkladı|açıklamıştır|bildirdi|bildirilmiştir|duyurdu|duyurmuştur|'
        r'ulaştı|ulaşmıştır|arttı|artmıştır|azaldı|azalmıştır|gerçekleşti|gerçekleşmiştir|bulunmaktadır|yer almaktadır)\b',norm(s)
    ):
        return True
    return False


def _v117_question_or_interview(s):
    n=norm(s)
    if '?' in s: return True
    if s.lstrip().startswith(('■','●','▪','►','- ')): return True
    # Birinci şahıs röportaj cümleleri bilgi notunda doğrudan kullanılmasın.
    if re.search(r'\b(?:bekliyoruz|düşünüyoruz|istiyoruz|inanıyoruz|hedefliyoruz|öngörüyoruz|bizim|bize|bizler)\b',n):
        return True
    return False


def _v117_fragment_body(text):
    t=_clean_note_text(text)
    if not t: return []
    t=_v117_strip_uppercase_runs(t)
    # Soru işaretli röportaj ara başlıklarını ayıkla.
    t=re.sub(r'\s*[■●▪►]\s*[^.!?]{0,220}\?\s*',' ',t)
    # Noktalama ve ';' AKT gövdesinde doğal olgu sınırlarıdır.
    raw=re.split(r"(?<=[.!?])\s+|;\s*|(?<=:)\s+(?=[“\"'A-ZÇĞİÖŞÜ])",t)
    return [_clean_note_text(x).strip(' ;') for x in raw if _clean_note_text(x).strip(' ;')]


def _v117_fact_candidates(title,body):
    title=_clean_note_text(title)
    tt=_v117_tokens(title)
    body=_clean_note_text(body)
    if title and norm(body).startswith(norm(title)):
        body=body[len(title):].lstrip(' -–—:;,.')
    raw=_v117_fragment_body(body)
    if not raw: return []
    # İlk üç gerçek cümle olayın varlık/kavram bağlamını verir.
    context=set(tt)
    for s in raw[:3]: context.update(_v117_tokens(s))
    result=[]; seen=[]
    for idx,s in enumerate(raw):
        n=norm(s)
        if len(s)<28 or n==norm(title): continue
        if n.startswith('fiyatı ') and 'fiyat' not in norm(title): continue
        if any(x in n for x in _V117_NOISE_TERMS): continue
        if _v117_heading_like(s) or _v117_question_or_interview(s): continue
        if s.startswith(('http://','https://','www.')): continue
        st=_v117_tokens(s)
        if not st: continue
        # Tekrar.
        duplicate=False
        for old in seen:
            union=len(st|old)
            if union and len(st&old)/union>=0.80:
                duplicate=True; break
        if duplicate: continue
        overlap=len(st&tt)
        context_overlap=len(st&context)
        relevance=overlap*2.0 + min(context_overlap,5)*0.55
        if tt and st:
            relevance += (len(st&tt)/max(1,len(st|tt)))*8
        info=_akt_sentence_score(s)+_sent_score(s)
        if re.search(r'\b\d+(?:[.,]\d+)?\b',s): info+=2
        # Çok ileride gelen ve konu bağını tamamen kaybetmiş sayfa kalıntısını bastır.
        if idx>=5 and overlap==0 and context_overlap<=1 and info<=2:
            continue
        result.append({'i':idx,'text':s,'rel':relevance,'info':info,'score':relevance+info})
        seen.append(st)
    return result


def _v117_pick_facts(title,body,limit=8,max_chars=2200,mode='akt'):
    cand=_v117_fact_candidates(title,body)
    if not cand:
        fb=_clean_note_text(body)
        if fb and norm(fb)!=norm(title) and not _v117_heading_like(fb):
            return [_v66_formalize_sentence_endings(fb)]
        return []

    # İlk iki gerçek olgu bağlamı kurar.
    chosen={c['i'] for c in cand[:min(2,len(cand))]}
    ranked=sorted(cand,key=lambda x:(x['score'], -x['i']),reverse=True)
    for c in ranked:
        if len(chosen)>=limit: break
        chosen.add(c['i'])

    # Bilgi notunda sonuca/ölçeğe işaret eden geç bir cümleyi de koru.
    if mode=='note' and len(cand)>=5:
        late=sorted(cand[max(2,len(cand)//2):],key=lambda x:(x['info']+x['rel'],x['i']),reverse=True)
        if late: chosen.add(late[0]['i'])
        chosen.add(cand[-1]['i'])

    lookup={c['i']:c for c in cand}
    out=[]; total=0
    for i in sorted(chosen):
        if i not in lookup: continue
        s=_v117_formal_sentence(lookup[i]['text'])
        if not s: continue
        if out and total+len(s)>max_chars: continue
        out.append(s); total+=len(s)+1
        if len(out)>=limit: break
    return out


def _v117_formal_sentence(s):
    s=_clean_note_text(s).strip(' ;')
    if not s: return ''
    s=_v66_formalize_sentence_endings(s)
    punct=s[-1] if s[-1:] in '.!?' else '.'
    core=s[:-1].rstrip() if s[-1:] in '.!?' else s.rstrip()

    # Yalnız cümle sonundaki gelecek/şimdiki zaman yüklemini resmî dile çevir.
    # Böylece "fırsatı sunacak IAC 2026" gibi sıfat-fiil yapıları bozulmaz.
    suffix_pairs=[
        ('düzenlenecek','düzenlenecektir'),('gerçekleştirilecek','gerçekleştirilecektir'),
        ('yapılacak','yapılacaktır'),('sağlanacak','sağlanacaktır'),('kurulacak','kurulacaktır'),
        ('açılacak','açılacaktır'),('sunulacak','sunulacaktır'),('olacak','olacaktır'),
        ('yapacak','yapacaktır'),('sunacak','sunacaktır'),('açacak','açacaktır'),
        ('gelecek','gelecektir'),('hedefliyor','hedeflemektedir'),('hazırlanıyor','hazırlanmaktadır'),
        ('buluşacak','bir araya gelecektir'),('çıkıyor','çıkmaktadır'),('bulunuyor','bulunmaktadır'),
        ('ulaşıyor','ulaşmaktadır'),('getiriyor','getirmektedir')
    ]
    low=core.lower()
    for old,newv in sorted(suffix_pairs,key=lambda x:len(x[0]),reverse=True):
        if low.endswith(old):
            core=core[:-len(old)]+newv
            break

    # Haber sayfasından yan-cümle biçiminde kopmuş '-dığı/-diği' sonlarını tamamla.
    if re.search(r'(?:olduğu|bulunduğu|gerçekleştiği|tamamlandığı|arttığı|azaldığı|duyurduğu|açıkladığı|belirttiği|bildirdiği|kaydettiği|vurguladığı|öğrenildiği)$',core,re.I):
        core += ' aktarılmıştır'
    elif re.search(r'öğrenildi$',core,re.I):
        core=re.sub(r'öğrenildi$','öğrenilmiştir',core,flags=re.I)
    s=re.sub(r'\s+',' ',core).strip()+punct
    return s

def _v117_record_to_tr(row):
    r=dict(row or {})
    return {
        'Başlık':r.get('Başlık',r.get('title','')),
        'Kaynak':r.get('Kaynak',r.get('source','')),
        'URL':r.get('URL',r.get('url','')),
        'Yayıncı_URL':r.get('Yayıncı_URL',r.get('url','')),
        'İçerik_Özeti':r.get('İçerik_Özeti',r.get('summary','')),
        'Tarih':r.get('Tarih',r.get('news_time','')),
        'Kategori':r.get('Kategori',r.get('category','')),
        'Risk_Skoru':r.get('Risk_Skoru',r.get('risk_score',0)),
        'Risk_Durumu':r.get('Risk_Durumu',r.get('risk_status','')),
    }


def _v117_local_context(row):
    """Sepet satırını mevcut taramadaki aynı olayın diğer kaynaklarıyla tekrar zenginleştirir."""
    tr=_v117_record_to_tr(row)
    candidates=[]
    original=_clean_note_text(tr.get('İçerik_Özeti',''))
    if original and norm(original)!=norm(tr.get('Başlık','')):
        candidates.append(original)
    try:
        enriched=_v107_enrich_selected_rows([tr]) or []
        if enriched:
            e=enriched[0]
            txt=_clean_note_text(e.get('İçerik_Özeti',''))
            if txt and norm(txt)!=norm(tr.get('Başlık','')):
                candidates.append(txt)
            # Daha iyi URL/kaynak varsa satıra geri aktar.
            if e.get('URL'): tr['URL']=e.get('URL')
            if e.get('Kaynak'): tr['Kaynak']=e.get('Kaynak')
    except Exception:
        pass

    # Oturum yeniden çizilmiş olsa bile son tarama geçmişinden aynı olaya ait zengin özeti bul.
    if (not candidates or max(map(len,candidates),default=0)<180) and _init_history_db():
        try:
            with _history_connect() as conn:
                hist=conn.execute(
                    "SELECT title,summary,source,url FROM event_snapshots ORDER BY scan_id DESC LIMIT 250"
                ).fetchall()
            best_txt=''; best_sim=0.0
            for ht,hs,hsrc,hu in hist:
                sim=_v104_event_similarity(tr.get('Başlık',''),original,str(ht or ''),str(hs or ''))
                if sim>best_sim and sim>=0.46:
                    best_sim=sim; best_txt=_clean_note_text(hs)
                    if hsrc and not tr.get('Kaynak'): tr['Kaynak']=hsrc
                    if hu and not tr.get('URL'): tr['URL']=hu
            if best_txt and norm(best_txt)!=norm(tr.get('Başlık','')):
                candidates.append(best_txt)
        except Exception:
            pass
    candidates=[x for x in candidates if x]
    candidates=sorted(candidates,key=len,reverse=True)
    return tr,(candidates[0] if candidates else '')


def _v117_search_snippets(title,preferred_domain=''):
    """Haber sayfası okunamazsa başlığı tek başına bırakmamak için arama sonucu özetlerini kullanır."""
    title=_clean_note_text(title)
    if not title: return []
    preferred=domain(preferred_domain)
    q=f'"{title[:220]}"' + (f' site:{preferred}' if preferred and 'google.com' not in preferred else '')
    out=[]
    try:
        rr=requests.get('https://html.duckduckgo.com/html/',params={'q':q},headers=HEADERS,timeout=6)
        if rr.ok:
            soup=BeautifulSoup(rr.text,'html.parser')
            for res in soup.select('.result')[:8]:
                a=res.select_one('a.result__a, a.result-link')
                sn=res.select_one('.result__snippet')
                label=_clean_note_text(a.get_text(' ',strip=True) if a else '')
                snippet=_clean_note_text(sn.get_text(' ',strip=True) if sn else '')
                if snippet and len(snippet)>=55:
                    sim=len(_v117_tokens(title)&_v117_tokens(label+' '+snippet))
                    if sim>=2: out.append((sim,snippet))
    except Exception:
        pass
    if not out:
        try:
            rr=requests.get('https://www.bing.com/search',params={'q':q,'setlang':'tr'},headers=HEADERS,timeout=6)
            if rr.ok:
                soup=BeautifulSoup(rr.text,'html.parser')
                for res in soup.select('li.b_algo')[:8]:
                    a=res.select_one('h2 a'); sn=res.select_one('.b_caption p, p')
                    label=_clean_note_text(a.get_text(' ',strip=True) if a else '')
                    snippet=_clean_note_text(sn.get_text(' ',strip=True) if sn else '')
                    if snippet and len(snippet)>=55:
                        sim=len(_v117_tokens(title)&_v117_tokens(label+' '+snippet))
                        if sim>=2: out.append((sim,snippet))
        except Exception:
            pass
    uniq=[]; seen=set()
    for _,s in sorted(out,reverse=True):
        k=title_key(s)
        if k not in seen:
            seen.add(k); uniq.append(s)
    return uniq[:3]


def _v117_fetch_clean_page(url,title=''):
    """Doğrudan haber sayfasından yalnız makale paragrafı + ana görsel adayını çıkarır."""
    if not _v116_valid_direct_url(url): return {'text':'','images':[],'source':'','title':''}
    cache=st.session_state.setdefault('_v117_page_cache',{})
    ck=str(url)
    if ck in cache: return dict(cache[ck])
    out={'text':'','images':[],'source':'','title':''}
    try:
        rr=requests.get(url,headers={**HEADERS,'Accept-Language':'tr-TR,tr;q=0.9,en;q=0.6'},timeout=10,allow_redirects=True)
        if not rr.ok or not rr.text:
            cache[ck]=out; return dict(out)
        soup=BeautifulSoup(rr.text,'html.parser')
        for tag in soup(['script','style','noscript','nav','footer','header','aside','form']):
            try: tag.decompose()
            except Exception: pass
        bad_rx=re.compile(r'(related|recommend|suggest|popular|sidebar|footer|header|navigation|menu|share|social|advert|banner|cookie|breadcrumb|tag-list|author-box)',re.I)
        for node in list(soup.find_all(True)):
            try:
                ident=' '.join([str(node.get('id') or ''),' '.join(node.get('class') or [])])
                if ident and bad_rx.search(ident): node.decompose()
            except Exception: pass

        # Başlık / site adı / açıklama.
        for attrs in ({'property':'og:title'},{'name':'twitter:title'}):
            t=soup.find('meta',attrs=attrs)
            if t and t.get('content'): out['title']=_clean_note_text(t['content']); break
        for attrs in ({'property':'og:site_name'},{'name':'application-name'}):
            t=soup.find('meta',attrs=attrs)
            if t and t.get('content'): out['source']=_clean_note_text(t['content']); break
        descriptions=[]
        for attrs in ({'property':'og:description'},{'name':'description'},{'name':'twitter:description'}):
            t=soup.find('meta',attrs=attrs)
            if t and t.get('content'):
                v=_clean_note_text(t['content'])
                if len(v)>=70: descriptions.append(v)

        json_bodies=[]; json_images=[]
        def walk(obj):
            if isinstance(obj,dict):
                typ=norm(obj.get('@type',''))
                if 'article' in typ or 'news' in typ:
                    if obj.get('articleBody'): json_bodies.append(_clean_note_text(obj.get('articleBody')))
                    if obj.get('description'): descriptions.append(_clean_note_text(obj.get('description')))
                    im=obj.get('image') or obj.get('thumbnailUrl')
                    if isinstance(im,str): json_images.append(im)
                    elif isinstance(im,list):
                        for z in im:
                            if isinstance(z,str): json_images.append(z)
                            elif isinstance(z,dict) and z.get('url'): json_images.append(str(z['url']))
                    elif isinstance(im,dict) and im.get('url'): json_images.append(str(im['url']))
                for v in obj.values(): walk(v)
            elif isinstance(obj,list):
                for v in obj: walk(v)
        for tag in soup.find_all('script',attrs={'type':re.compile(r'application/ld\+json',re.I)}):
            try:
                raw=tag.string or tag.get_text()
                if raw: walk(json.loads(raw))
            except Exception: pass

        selectors=['[itemprop="articleBody"]','article','[class*="article-body"]','[class*="article-content"]',
                   '[class*="news-content"]','[class*="news-detail"]','[class*="story-body"]',
                   '[class*="post-content"]','[class*="entry-content"]','[class*="content-body"]','main']
        bodies=list(json_bodies)
        lead_images=[]
        for selector in selectors:
            for node in soup.select(selector)[:3]:
                ps=[]
                for p in node.find_all('p'):
                    txt=_clean_note_text(p.get_text(' ',strip=True))
                    if len(txt)>=38 and not any(b in norm(txt) for b in _V117_NOISE_TERMS): ps.append(txt)
                if ps:
                    body=' '.join(ps)
                    if len(body)>=180: bodies.append(body)
                for img in node.find_all('img')[:8]:
                    for attr in ('src','data-src','data-lazy-src','data-original'):
                        if img.get(attr): lead_images.append(requests.compat.urljoin(rr.url,img.get(attr))); break
        for attrs in ({'property':'og:image'},{'property':'og:image:url'},{'name':'twitter:image'}):
            t=soup.find('meta',attrs=attrs)
            if t and t.get('content'): lead_images.insert(0,requests.compat.urljoin(rr.url,t['content']))
        lead_images=json_images+lead_images

        candidates=[]
        for b in descriptions+bodies:
            b=_clean_note_text(b)
            if not b: continue
            facts=_v117_fact_candidates(title or out['title'],b)
            score=len(facts)*300 + min(len(b),6000)
            if facts: candidates.append((score,b))
        if candidates:
            out['text']=max(candidates,key=lambda z:z[0])[1]
        seen=set()
        for im in lead_images:
            im=str(im or '').strip()
            low=im.lower()
            if not im or im in seen or any(x in low for x in ('logo','favicon','sprite','avatar','icon','pixel')): continue
            seen.add(im); out['images'].append(im)
            if len(out['images'])>=8: break
    except Exception:
        pass
    cache[ck]=dict(out)
    return dict(out)


_v117_article_detail_base=article_detail

def article_detail(row):
    """V117: V116 çözümlemesini korur; gövdeyi temiz makale metni / yerel olay bağlamı ile güçlendirir."""
    if isinstance(row,str): row={'URL':row}
    elif hasattr(row,'to_dict'): row=row.to_dict()
    elif row is None: row={}
    else: row=dict(row)
    tr,local_text=_v117_local_context(row)
    title0=_clean_note_text(tr.get('Başlık',''))
    source0=_clean_note_text(tr.get('Kaynak',''))
    key=hashlib.sha1((title0+'|'+str(tr.get('URL',''))+'|'+local_text[:800]).encode('utf-8','ignore')).hexdigest()
    cache=st.session_state.setdefault('_v117_article_cache',{})
    if key in cache: return dict(cache[key])
    try:
        base=_v117_article_detail_base(tr) or {}
    except Exception:
        base={}
    canonical=str(base.get('canonical') or tr.get('URL') or '').strip()
    title=_v116_clean_headline(base.get('title') or title0,base.get('source') or source0)
    source=_v117_source_short(base.get('source') or source0,canonical)
    page={'text':'','images':[],'source':'','title':''}
    # V116 gövdesi çok kısa/başlıkla aynı/noise içeriyorsa temiz p-tag çıkarımıyla ikinci okuma.
    base_text=_clean_note_text(base.get('text') or '')
    upper_noise=(_v117_strip_uppercase_runs(base_text)!=base_text)
    base_facts=_v117_fact_candidates(title,base_text)
    if _v116_valid_direct_url(canonical) and (len(base_facts)<3 or upper_noise):
        page=_v117_fetch_clean_page(canonical,title)
    candidates=[]
    for txt,bonus in ((page.get('text',''),500),(base_text,250),(local_text,350),(_clean_note_text(tr.get('İçerik_Özeti','')),100)):
        txt=_clean_note_text(txt)
        if not txt or norm(txt)==norm(title): continue
        facts=_v117_fact_candidates(title,txt)
        score=len(facts)*450 + min(len(txt),6500) + bonus
        if facts: candidates.append((score,txt))
    # Sayfa okunamadıysa arama snippet'i başlık-only çıktıyı engeller.
    if not candidates or max((len(_v117_fact_candidates(title,x[1])) for x in candidates),default=0)<2:
        preferred=canonical or tr.get('URL','')
        snippets=_v117_search_snippets(title,preferred)
        if snippets:
            stxt=' '.join(snippets)
            candidates.append((len(_v117_fact_candidates(title,stxt))*450+len(stxt)+200,stxt))
    best_text=max(candidates,key=lambda z:z[0])[1] if candidates else (local_text or base_text or _clean_note_text(tr.get('İçerik_Özeti','')))
    result=dict(base)
    result['title']=title or title0
    result['source']=_v117_source_short(page.get('source') or source,canonical)
    result['canonical']=canonical
    result['text']=best_text
    # İmajda ikinci çıkarım daha güvenilir; yoksa V116 adayları.
    result['images']=list(page.get('images') or base.get('images') or [])
    cache[key]=dict(result)
    return result


def _v117_valid_report_image(url):
    bio=_download_report_image(url)
    if not bio: return None
    try:
        im=Image.open(bio)
        w,h=im.size
        # Site ikonu / küçük thumbnail yerine gerçek haber görseli.
        if w<480 or h<260 or w*h<180000:
            return None
        bio.seek(0); return bio
    except Exception:
        return None


def _v117_akt_summary(title,body):
    facts=_v117_pick_facts(title,body,limit=6,max_chars=1750,mode='akt')
    if not facts: return ''
    clauses=[]
    for s in facts:
        s=s.strip().rstrip(' .;:')
        if not s: continue
        # Referans AKT tek akış kullanıyor; her cümleyi ayrı başlık gibi bırakma.
        if s and not s[:5].isupper():
            s=s[0].lower()+s[1:]
        clauses.append(s)
    return '; '.join(clauses)


def _v117_prepare_akt_row(row):
    tr,local_text=_v117_local_context(row)
    detail=article_detail(tr)
    real_url=str(detail.get('canonical') or tr.get('URL') or '').strip()
    source=_v117_source_short(_real_source(tr,detail,real_url),real_url)
    title=_v116_clean_headline(detail.get('title') or tr.get('Başlık',''),source)
    body=_clean_note_text(detail.get('text') or local_text or tr.get('İçerik_Özeti') or '')
    summary=_v117_akt_summary(title,body)
    if not summary:
        # Başlığı özet diye ikinci kez yazma; yerel özet yoksa kısa resmî ifade.
        fallback=_clean_note_text(local_text or tr.get('İçerik_Özeti',''))
        if fallback and norm(fallback)!=norm(title):
            summary=_v117_akt_summary(title,fallback) or _v117_formal_sentence(fallback).rstrip('.')
        else:
            summary=_v117_formal_sentence(title).rstrip('.')
    if not _v116_valid_direct_url(real_url):
        real_url=str(tr.get('URL') or '')
    image=None
    for candidate in list(detail.get('images') or [])[:8]:
        image=_v117_valid_report_image(candidate)
        if image: break
    return {'title':title,'source':source,'url':real_url,'summary':summary,'image':image}


def make_docx(rows):
    """V117 AKT — referans düzen + temiz, konuya bağlı ayrıntılı haber özeti."""
    rows0=[]
    for r in (rows or []):
        tr,_=_v117_local_context(r)
        rows0.append(tr)
    rows=_v115_dedupe_rows(rows0)
    doc=_v116_doc_defaults(Document(),top=2.5,bottom=1.25,left=2.5,right=2.5)
    sec=doc.sections[0]; sec.header_distance=Cm(1.25); sec.footer_distance=Cm(1.25)
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(0)
    _v116_run(p.add_run('AÇIK KAYNAK TARAMA ÇALIŞMASI'),16,True)
    _v116_add_akt_info_table(doc)
    intro=doc.add_paragraph(); intro.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    intro.paragraph_format.first_line_indent=Cm(0.63); intro.paragraph_format.line_spacing=1.5
    intro.paragraph_format.space_before=Pt(0); intro.paragraph_format.space_after=Pt(0)
    intro_text=_v116_akt_intro_text(rows); topics=_v116_topic_labels(rows)
    cursor=0; spans=[]; pos=0
    for topic in topics:
        q=f'“{topic}”'; ix=intro_text.find(q,pos)
        if ix>=0: spans.append((ix,ix+len(q))); pos=ix+len(q)
    for a,b in spans:
        if a>cursor: _v116_run(intro.add_run(intro_text[cursor:a]),12)
        _v116_run(intro.add_run(intro_text[a:b]),12,italic=True); cursor=b
    if cursor<len(intro_text): _v116_run(intro.add_run(intro_text[cursor:]),12)

    prepared=[None]*len(rows)
    if rows:
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(3,len(rows))) as ex:
            futs={ex.submit(_v117_prepare_akt_row,r):i for i,r in enumerate(rows)}
            for fut in concurrent.futures.as_completed(futs):
                try: prepared[futs[fut]]=fut.result()
                except Exception: prepared[futs[fut]]=None
    out_no=0
    for item in prepared:
        if not item: continue
        out_no+=1
        p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent=Cm(0.63); p.paragraph_format.line_spacing=1.5
        p.paragraph_format.space_before=Pt(4); p.paragraph_format.space_after=Pt(6)
        _v116_run(p.add_run(f'{out_no}. “{item["source"]}”'),12,True)
        _v116_run(p.add_run(' isimli internet sitesinde, '),12)
        _v116_run(p.add_run(f'“{item["title"]}”'),12,True,True)
        _v116_run(p.add_run(' başlığıyla bir haber yayımlanmıştır. ('),12)
        _word_hyperlink(p,item['url'],item['url'] or 'Haber bağlantısı')
        _v116_run(p.add_run(') Söz konusu haber içeriğinde, '),12)
        _v116_run(p.add_run((item.get('summary') or '').strip().rstrip(' .;')),12)
        _v116_run(p.add_run(' hususları ifade edilmiştir.'),12)
        if item.get('image'):
            cap=doc.add_paragraph(); cap.alignment=WD_ALIGN_PARAGRAPH.CENTER
            cap.paragraph_format.first_line_indent=Cm(0.63); cap.paragraph_format.line_spacing=1.5
            cap.paragraph_format.space_before=Pt(4); cap.paragraph_format.space_after=Pt(4)
            _v116_run(cap.add_run(f'Görsel {out_no}: “{item["source"]}” Sitesinde Yer Alan Görsel'),12,True)
            ip=doc.add_paragraph(); ip.alignment=WD_ALIGN_PARAGRAPH.CENTER; ip.paragraph_format.space_after=Pt(8)
            try:
                run=ip.add_run(); shape=run.add_picture(item['image'])
                max_w=Cm(15.3); max_h=Cm(12.7)
                scale=min(1.0,max_w/shape.width,max_h/shape.height)
                if scale<1.0:
                    shape.width=int(shape.width*scale); shape.height=int(shape.height*scale)
            except Exception: pass
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY; p.paragraph_format.line_spacing=1.5
    _v116_run(p.add_run('Arz olunur.'),12)
    bio=BytesIO(); doc.save(bio); bio.seek(0); return bio.getvalue()


def _v117_note_paragraphs(title,body):
    """Bilgi notunu ham alıntı değil, giriş-gelişme-sonuç akışında yoğunlaştırır."""
    cand=_v117_fact_candidates(title,body)
    if not cand:
        fallback=_v117_formal_sentence(title)
        return [fallback] if fallback else []

    selected=[]
    def add_candidate(c):
        st=_v117_tokens(c['text'])
        for old in selected:
            ot=_v117_tokens(old['text']); union=len(st|ot)
            if union and len(st&ot)/union>=0.55:
                # Aynı olguyu tekrar eden cümle, belirgin yeni sayısal veri taşımıyorsa alınmaz.
                nums=set(re.findall(r'\b\d+(?:[.,]\d+)?\b',c['text']))
                oldnums=set(re.findall(r'\b\d+(?:[.,]\d+)?\b',old['text']))
                if not (nums-oldnums):
                    return False
        selected.append(c); return True

    # Giriş için ilk iki gerçek olgu.
    for c in cand[:2]: add_candidate(c)
    # Sonuç için son bölümdeki en bağlamsal/etki cümlesini rezerve et.
    concl_terms=('sonuç','böylece','bu kapsamda','etki','katkı','imkân','sağlam','hedef','beklen',
                 'öngör','kapasite','ulaş','toplam','öne çık','planlan','merkez haline','altyapı')
    conclusion=None
    for c in reversed(cand):
        if any(k in norm(c['text']) for k in concl_terms):
            conclusion=c; break
    if conclusion is None: conclusion=cand[-1]

    pool=[c for c in cand[2:] if c['i']!=conclusion['i']]
    # Gelişmede sayı, kurum, kapasite, tarih gibi somut bilgi rel skorundan daha önemlidir.
    pool=sorted(pool,key=lambda c:(c['info']*2 + c['rel']*0.35,-c['i']),reverse=True)
    for c in pool:
        if len(selected)>=8: break
        add_candidate(c)
    add_candidate(conclusion)
    selected=sorted(selected,key=lambda c:c['i'])[:9]
    facts=[_v117_formal_sentence(c['text']) for c in selected]
    facts=[x for x in facts if x]
    if len(facts)<=2: return [' '.join(facts)]

    intro=facts[:2]
    # Son seçili cümle sonuç paragrafına ayrılır; kalanlar gelişmedir.
    conclusion_text=[facts[-1]]
    dev=facts[2:-1]
    paras=[' '.join(intro)]
    if dev:
        if len(dev)<=3:
            paras.append(' '.join(dev))
        else:
            cut=math.ceil(len(dev)/2)
            paras.append(' '.join(dev[:cut])); paras.append(' '.join(dev[cut:]))
    paras.append(' '.join(conclusion_text))
    out=[]; seen=set()
    for para in paras:
        para=_clean_note_text(para)
        k=title_key(para)
        if para and k not in seen:
            seen.add(k); out.append(para)
    return out[:5]

def make_analyst_docx(df,title='BİLGİ NOTU'):
    """V117 Bilgi Notu — başlıksız, giriş-gelişme-sonuç akışında resmî ve özetlenmiş metin."""
    doc=_v116_doc_defaults(Document(),2.5,2.5,2.5,2.5)
    sec=doc.sections[0]; sec.header_distance=Cm(1.25); sec.footer_distance=Cm(1.25)
    x=df.copy() if df is not None else pd.DataFrame()
    if x.empty: rows=[]
    else:
        if 'Tarih_dt' in x.columns:
            x['Tarih_dt']=pd.to_datetime(x['Tarih_dt'],utc=True,errors='coerce')
            x=x.sort_values('Tarih_dt',ascending=True,na_position='last')
        rows=x.to_dict('records')
    # Bilgi notu genellikle tek seçili gelişme; birden çok kaynak varsa olay bazında zenginleştir.
    rows=_v115_dedupe_rows(rows)
    all_paras=[]
    for r in rows:
        tr,local_text=_v117_local_context(r)
        detail=article_detail(tr)
        source=_v117_source_short(_real_source(tr,detail,detail.get('canonical','')),detail.get('canonical',''))
        ttl=_v116_clean_headline(detail.get('title') or tr.get('Başlık',''),source)
        body=_clean_note_text(detail.get('text') or local_text or tr.get('İçerik_Özeti') or '')
        pars=_v117_note_paragraphs(ttl,body)
        all_paras.extend(pars)
    if not all_paras:
        all_paras=['Seçilen gelişmeye ilişkin yeterli içerik elde edilemediğinden ayrıntılı bilgi notu oluşturulamamıştır.']
    for text in all_paras[:5]:
        p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0); p.paragraph_format.line_spacing=1.0
        _v116_run(p.add_run(text),12)
    bio=BytesIO(); doc.save(bio); bio.seek(0); return bio.getvalue()


def _v117_ogn_item(row):
    tr,local_text=_v117_local_context(row)
    detail=article_detail(tr)
    source=_v117_source_short(_real_source(tr,detail,detail.get('canonical','')),detail.get('canonical',''))
    title=_v116_clean_headline(detail.get('title') or tr.get('Başlık',''),source)
    body=_clean_note_text(detail.get('text') or local_text or tr.get('İçerik_Özeti') or '')
    facts=_v117_pick_facts(title,body,limit=4,max_chars=850,mode='ogn')
    # Başlık tek başına kaldıysa onu haber başlığı gibi değil, resmî olgu cümlesi haline getir.
    if not facts:
        facts=[_v117_formal_sentence(title)] if title else []
    # En fazla 3 cümle / yaklaşık 600-700 karakter: referans ÖGN yoğunluğu.
    chosen=[]; total=0
    for s in facts:
        s=_v117_formal_sentence(s)
        if not s: continue
        if chosen and (len(chosen)>=3 or total+len(s)>680): break
        chosen.append(s); total+=len(s)+1
    text=' '.join(chosen).strip()
    # Sırf başlığın aynısı ise daha kurumsal bildirime çevir.
    if text and title and title_key(text.rstrip('.'))==title_key(title.rstrip('.')):
        text=_v117_formal_sentence(title)
    return _v98_strip_site_name(text,source).strip()


def make_important_basket_docx_v101(basket_df):
    """V117 ÖGN — referans düzen; her madde başlık değil, gelişmenin kısa kurumsal özetidir."""
    doc=_v116_doc_defaults(Document(),top=2.25,bottom=1.50,left=1.905,right=1.905)
    sec=doc.sections[0]; sec.header_distance=Cm(0.1); sec.footer_distance=Cm(0.45)
    try:
        ns=doc.styles['No Spacing']; ns.font.name='Times New Roman'; ns.font.size=Pt(12)
        ns._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'Times New Roman')
    except Exception: pass
    _v116_ogn_footer(sec)
    today=datetime.now().astimezone().date(); yesterday=today-timedelta(days=1)
    p=doc.add_paragraph(style='No Spacing'); p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.line_spacing=1.15; p.paragraph_format.space_before=Pt(6); p.paragraph_format.space_after=Pt(6)
    _v116_run(p.add_run(f'{yesterday.strftime("%d/%m/%Y")} – {today.strftime("%d/%m/%Y")}'),12)
    p=doc.add_paragraph(style='No Spacing'); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing=1.15; p.paragraph_format.space_before=Pt(6); p.paragraph_format.space_after=Pt(6)
    _v116_run(p.add_run('Konu: '),12,True); _v116_run(p.add_run('STB Temsilciliği Önemli Gelişmeler Notu'),12)
    raw=[] if basket_df is None else basket_df.to_dict('records')
    rows=_v115_dedupe_rows(raw)
    outputs=[None]*len(rows)
    if rows:
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(3,len(rows))) as ex:
            futs={ex.submit(_v117_ogn_item,r):i for i,r in enumerate(rows)}
            for fut in concurrent.futures.as_completed(futs):
                try: outputs[futs[fut]]=fut.result()
                except Exception: outputs[futs[fut]]=''
    for text in outputs:
        text=_clean_note_text(text)
        if not text: continue
        p=doc.add_paragraph(style='No Spacing'); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing=1.15; p.paragraph_format.space_before=Pt(6); p.paragraph_format.space_after=Pt(6)
        if text.endswith('.'): text=text[:-1]
        _v116_run(p.add_run(text+' (STB).'),12)
    p=doc.add_paragraph(style='No Spacing'); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing=1.15; p.paragraph_format.space_before=Pt(6); p.paragraph_format.space_after=Pt(6)
    _v116_run(p.add_run('Arz olunur.'),12)
    bio=BytesIO(); doc.save(bio); bio.seek(0); return bio.getvalue()

# ============================================================
# /V117
# ============================================================

# ============================================================
# V118 — STANDART RAPOR MOTORU / SUNUM GÜVENLİ ÇIKTI
# 1) Şifre ekranı adı: STB-Açık Kaynak Tarama Merkezi
# 2) Haber gövdesinde konu dışı / karışmış satırlar daha katı ayıklanır.
# 3) Bilgi notunda tek-cümle/başlık çıktısı yerine asgari kalite standardı uygulanır.
# 4) ÖGN ve AKT için aynı zenginleştirilmiş kaynak havuzu kullanılır.
# ============================================================

_V118_ENGINE_VERSION='V118-2026-09-21-A'


def _v118_fix_surface(text):
    """Türkçe rapor çıktısındaki basit yüzey/boşluk/büyük-küçük harf hatalarını düzeltir."""
    s=_clean_note_text(text)
    if not s: return ''
    s=re.sub(r"\s+([,.;:!?])",r"\1",s)
    s=re.sub(r"\s+(['’])",r"\1",s)
    s=re.sub(r"(['’])\s+",r"\1",s)
    # Türkçe yer/ad hataları: "rusya 'da" -> "Rusya'da"
    for old,newv in (
        ('rusya','Rusya'),('türkiye','Türkiye'),('amerika birleşik devletleri','Amerika Birleşik Devletleri'),
        ('google','Google'),('apple','Apple'),('microsoft','Microsoft'),('openai','OpenAI'),
        ('anthropic','Anthropic'),('meta','Meta'),('ukrayna','Ukrayna'),('avrupa birliği','Avrupa Birliği')
    ):
        s=re.sub(r'(?<!\w)'+re.escape(old)+r'(?!\w)',newv,s,flags=re.I)
    s=re.sub(r'(?<!\w)aI(?!\w)','AI',s)
    # Her cümle/clauseda ilk harfi büyüt; kalan karakterlere dokunma.
    chars=list(s)
    need_cap=True
    for i,ch in enumerate(chars):
        if need_cap and ch.isalpha():
            chars[i]=ch.upper(); need_cap=False
        if ch in '.!?;':
            need_cap=True
    s=''.join(chars)
    return re.sub(r'\s+',' ',s).strip()


def _v118_title_relevance(title,sentence):
    tt=_v117_tokens(title); st=_v117_tokens(sentence)
    if not tt or not st: return 0.0
    inter=len(tt&st)
    jac=inter/max(1,len(tt|st))
    return inter*2.2+jac*10


def _v118_strict_facts(title,body,max_items=12):
    """Konu dışı sayfa parçalarını bastırıp aynı olaya bağlı olguları kronolojik sırayla döndürür."""
    title=_clean_note_text(title)
    raw=_v117_fragment_body(body)
    if not raw: return []
    tt=_v117_tokens(title)
    clean=[]
    for i,s in enumerate(raw):
        s=_v118_fix_surface(s)
        n=norm(s)
        if len(s)<26 or n==norm(title): continue
        if any(x in n for x in _V117_NOISE_TERMS): continue
        if _v117_heading_like(s) or _v117_question_or_interview(s): continue
        if s.startswith(('http://','https://','www.')): continue
        st=_v117_tokens(s)
        if not st: continue
        direct=len(st&tt)
        info=_akt_sentence_score(s)+_sent_score(s)+(2 if re.search(r'\b\d+(?:[.,]\d+)?\b',s) else 0)
        clean.append({'i':i,'text':s,'tok':st,'direct':direct,'rel':_v118_title_relevance(title,s),'info':info})
    if not clean: return []

    # Olayı başlatan anchor: başlığa en yakın ilk anlamlı cümle.
    early=clean[:min(8,len(clean))]
    anchor=max(early,key=lambda c:(c['rel']+c['info']*0.25,-c['i']))
    context=set(tt)|set(anchor['tok'])
    accepted=[]
    last_i=anchor['i']
    for c in clean:
        ctx=len(c['tok']&context)
        near=abs(c['i']-last_i)<=2
        # İlk sayfalardaki tamamen alakasız başka haber başlıklarını alma.
        keep=(c['direct']>=1 or ctx>=2 or (near and ctx>=1 and c['info']>=2))
        # Sayısal/kurumsal güçlü cümle, anchor'a çok yakınsa doğrudan başlık kelimesi geçmese de alınabilir.
        if not keep and abs(c['i']-anchor['i'])<=3 and c['info']>=6 and ctx>=1:
            keep=True
        if not keep: continue
        # Çok yüksek başlık benzerliğine sahip fakat kısa link/spot kalıntısını bastır.
        if len(c['text'].split())<6 and c['info']<3: continue
        # Tekrar kontrolü.
        dup=False
        for a in accepted:
            union=len(c['tok']|a['tok'])
            if union and len(c['tok']&a['tok'])/union>=0.76:
                dup=True; break
        if dup: continue
        accepted.append(c)
        context.update(c['tok'])
        last_i=c['i']
        if len(accepted)>=max_items: break
    if not accepted:
        accepted=[anchor]
    return accepted


def _v118_fact_count(title,text):
    return len(_v118_strict_facts(title,text,max_items=12))


def _v118_search_urls(title,preferred_url=''):
    """Başlıkla ilgili alternatif doğrudan haber sayfalarını bulur; tek kaynak bloklanırsa ikinci şans sağlar."""
    title=_clean_note_text(title)
    if not title: return []
    preferred_host=domain(preferred_url)
    q='"'+title[:210]+'"'
    results=[]
    try:
        rr=requests.get('https://www.bing.com/search',params={'q':q,'setlang':'tr'},headers=HEADERS,timeout=6)
        if rr.ok:
            soup=BeautifulSoup(rr.text,'html.parser')
            for a in soup.select('li.b_algo h2 a')[:10]:
                u=str(a.get('href') or '').strip(); label=_clean_note_text(a.get_text(' ',strip=True))
                if u.startswith('http') and 'bing.com' not in domain(u) and _v118_title_relevance(title,label)>=2:
                    results.append(u)
    except Exception:
        pass
    try:
        rr=requests.get('https://html.duckduckgo.com/html/',params={'q':q},headers=HEADERS,timeout=6)
        if rr.ok:
            soup=BeautifulSoup(rr.text,'html.parser')
            for a in soup.select('a.result__a, a.result-link')[:10]:
                u=str(a.get('href') or '').strip(); label=_clean_note_text(a.get_text(' ',strip=True))
                if 'uddg=' in u:
                    try:
                        from urllib.parse import parse_qs, urlparse as _uparse, unquote
                        u=unquote(parse_qs(_uparse(u).query).get('uddg',[''])[0]) or u
                    except Exception:
                        pass
                if u.startswith('http') and 'duckduckgo.com' not in domain(u) and _v118_title_relevance(title,label)>=2:
                    results.append(u)
    except Exception:
        pass
    uniq=[]; seen=set()
    # Önce aynı yayıncı domainindeki sonuçlar.
    results=sorted(results,key=lambda u:(0 if preferred_host and domain(u)==preferred_host else 1))
    for u in results:
        k=u.split('#')[0]
        if k not in seen:
            seen.add(k); uniq.append(k)
    return uniq[:5]


_v118_article_detail_base=article_detail

def article_detail(row):
    """V118: tek cümle riskini azaltmak için doğrudan sayfa + yerel olay + arama snippet + alternatif kaynak zinciri."""
    if isinstance(row,str): row={'URL':row}
    elif hasattr(row,'to_dict'): row=row.to_dict()
    elif row is None: row={}
    else: row=dict(row)
    tr,local_text=_v117_local_context(row)
    title0=_clean_note_text(tr.get('Başlık',''))
    try:
        base=_v118_article_detail_base(tr) or {}
    except Exception:
        base={}
    title=_v116_clean_headline(base.get('title') or title0,base.get('source') or tr.get('Kaynak',''))
    canonical=str(base.get('canonical') or tr.get('URL') or '').strip()
    candidate_texts=[]
    def add(txt,bonus,label=''):
        txt=_clean_note_text(txt)
        if not txt or norm(txt)==norm(title): return
        facts=_v118_strict_facts(title,txt,max_items=12)
        if facts:
            score=len(facts)*900+sum(min(c['info'],12) for c in facts)*25+min(len(txt),6000)+bonus
            candidate_texts.append((score,txt,label))
    add(base.get('text',''),300,'base')
    add(local_text,500,'local')
    add(tr.get('İçerik_Özeti',''),150,'summary')

    # Mevcut kaynak yetersizse temiz doğrudan sayfa okumasını zorla.
    current_best=max((_v118_fact_count(title,x[1]) for x in candidate_texts),default=0)
    if _v116_valid_direct_url(canonical) and current_best<4:
        pg=_v117_fetch_clean_page(canonical,title)
        add(pg.get('text',''),700,'page')
        if pg.get('source'): base['source']=pg.get('source')
        if pg.get('title'): title=_v116_clean_headline(pg.get('title'),base.get('source') or tr.get('Kaynak',''))
        if pg.get('images'): base['images']=pg.get('images')

    # Arama snippetleri — çoğu bloklu/JS haber için başlıktan fazlasını sağlar.
    current_best=max((_v118_fact_count(title,x[1]) for x in candidate_texts),default=0)
    if current_best<4:
        snippets=_v117_search_snippets(title,canonical or tr.get('URL',''))
        if snippets: add(' '.join(snippets),450,'snippets')

    # Hâlâ yetersizse alternatif doğrudan sonuçları okuyup en alakalı makale gövdesini seç.
    current_best=max((_v118_fact_count(title,x[1]) for x in candidate_texts),default=0)
    if current_best<3:
        for u in _v118_search_urls(title,canonical or tr.get('URL',''))[:3]:
            pg=_v117_fetch_clean_page(u,title)
            before=len(candidate_texts)
            add(pg.get('text',''),550,'alternate')
            if len(candidate_texts)>before and (not _v116_valid_direct_url(canonical) or 'google.com' in domain(canonical)):
                canonical=u
            if max((_v118_fact_count(title,x[1]) for x in candidate_texts),default=0)>=4:
                break

    best=max(candidate_texts,key=lambda z:z[0]) if candidate_texts else None
    result=dict(base)
    result['title']=title or title0
    result['canonical']=canonical
    result['source']=_v117_source_short(result.get('source') or tr.get('Kaynak',''),canonical)
    result['text']=best[1] if best else _clean_note_text(local_text or tr.get('İçerik_Özeti',''))
    result['quality_facts']=_v118_fact_count(title,result['text'])
    return result


def _v118_formal_facts(title,body,limit=9):
    facts=[]
    for c in _v118_strict_facts(title,body,max_items=limit):
        s=_v118_fix_surface(_v117_formal_sentence(c['text']))
        if s and title_key(s.rstrip('.'))!=title_key(title.rstrip('.')):
            facts.append(s)
    return facts


def _v118_title_fact(title):
    s=_v118_fix_surface(_v117_formal_sentence(title))
    # Site adı kalıntılarını kaldır.
    s=re.sub(r'\s+(?:-|–|—)\s+[A-ZÇĞİÖŞÜa-zçğıöşü0-9. ]{2,60}\.?$','.',s).strip()
    return s


def _v118_note_paragraphs(title,body,source='',category=''):
    """Her seçili haber için asgari 3 paragraflık, giriş-gelişme-sonuç standardı."""
    facts=_v118_formal_facts(title,body,limit=9)
    # Başlığın taşıdığı gerçek olgu gövdede yoksa girişe ekle.
    title_fact=_v118_title_fact(title)
    if title_fact and not any(_v118_title_relevance(title,x)>=6 for x in facts[:2]):
        facts.insert(0,title_fact)
    # Tekrarları yeniden temizle.
    clean=[]
    for f in facts:
        ft=_v117_tokens(f); dup=False
        for old in clean:
            ot=_v117_tokens(old); union=len(ft|ot)
            if union and len(ft&ot)/union>=0.70:
                dup=True; break
        if not dup: clean.append(f)
    facts=clean[:9]

    # Tam/iyi veri: 3-4 paragraf, olay akışı korunur.
    if len(facts)>=5:
        intro=' '.join(facts[:2])
        mid=facts[2:-1]
        dev1=' '.join(mid[:3])
        dev2=' '.join(mid[3:]) if len(mid)>3 else ''
        conclusion=facts[-1]
        return [x for x in (intro,dev1,dev2,conclusion) if x]
    if len(facts)==4:
        return [facts[0], ' '.join(facts[1:3]), facts[3]]
    if len(facts)==3:
        return [facts[0], facts[1], facts[2]]
    if len(facts)==2:
        # Veri sınırlıysa yine tek cümle bırakma; açıkça kaynak kapsamını belirt.
        tail='Mevcut açık kaynak içeriğinde gelişmeye ilişkin ayrıntılar sınırlı olmakla birlikte, yukarıdaki hususlar olayın mevcut görünümünü oluşturmaktadır.'
        return [facts[0],facts[1],tail]
    if len(facts)==1:
        context=[]
        if source: context.append(f'Gelişme, {_v117_source_short(source,"")} tarafından yayımlanan açık kaynak içeriğinde yer almaktadır.')
        if category: context.append(f'Konu, {category} başlığı altında takip edilmektedir.')
        if not context:
            context.append('Mevcut açık kaynak kaydında gelişmeye ilişkin ayrıntılı teknik veya sayısal bilgi sınırlı düzeydedir.')
        tail='Mevcut veriler çerçevesinde bilgi notu, doğrulanabilen hususlarla sınırlı tutulmuştur.'
        return [facts[0],context[0],tail]
    # Hiç gövde bulunamadığında dahi başlığı tek satır halinde bırakma.
    tf=title_fact or 'Seçilen gelişmeye ilişkin temel başlık bilgisi alınmıştır.'
    return [tf,
            'Mevcut açık kaynak kaydında gelişmenin ayrıntılarını destekleyecek yeterli gövde metni elde edilememiştir.',
            'Bilgi notu, doğrulanamayan ayrıntı eklenmemesi amacıyla mevcut verilerle sınırlı tutulmuştur.']


def make_analyst_docx(df,title='BİLGİ NOTU'):
    """V118 — hangi sepetten çağrılırsa çağrılsın aynı kalite standardında tek haber bilgi notu."""
    doc=_v116_doc_defaults(Document(),2.5,2.5,2.5,2.5)
    sec=doc.sections[0]; sec.header_distance=Cm(1.25); sec.footer_distance=Cm(1.25)
    x=df.copy() if df is not None else pd.DataFrame()
    rows=[] if x.empty else _v115_dedupe_rows(x.to_dict('records'))
    all_paras=[]
    for r in rows:
        tr,local_text=_v117_local_context(r)
        detail=article_detail(tr)
        source=_v117_source_short(_real_source(tr,detail,detail.get('canonical','')),detail.get('canonical',''))
        ttl=_v116_clean_headline(detail.get('title') or tr.get('Başlık',''),source)
        body=_clean_note_text(detail.get('text') or local_text or tr.get('İçerik_Özeti') or '')
        pars=_v118_note_paragraphs(ttl,body,source,tr.get('Kategori',''))
        all_paras.extend(pars)
    if not all_paras:
        all_paras=['Seçilen gelişmeye ilişkin temel kayıt alınmıştır.',
                   'Mevcut açık kaynak içeriğinde ayrıntılı gövde metni elde edilememiştir.',
                   'Bilgi notu, doğrulanabilen verilerle sınırlı tutulmuştur.']
    for text in all_paras[:5]:
        p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0); p.paragraph_format.line_spacing=1.0
        _v116_run(p.add_run(_v118_fix_surface(text)),12)
    bio=BytesIO(); doc.save(bio); bio.seek(0); return bio.getvalue()


def _v118_ogn_item(row):
    tr,local_text=_v117_local_context(row)
    detail=article_detail(tr)
    source=_v117_source_short(_real_source(tr,detail,detail.get('canonical','')),detail.get('canonical',''))
    title=_v116_clean_headline(detail.get('title') or tr.get('Başlık',''),source)
    body=_clean_note_text(detail.get('text') or local_text or tr.get('İçerik_Özeti') or '')
    facts=_v118_formal_facts(title,body,limit=5)
    if not facts:
        facts=[_v118_title_fact(title)] if title else []
    # Başlık + gövde ayrıntısı: tek başlık görünümünü önle.
    if len(facts)==1:
        snippets=_v117_search_snippets(title,detail.get('canonical') or tr.get('URL',''))
        if snippets:
            extra=_v118_formal_facts(title,' '.join(snippets),limit=3)
            for x in extra:
                if x not in facts: facts.append(x)
    chosen=[]; total=0
    for s in facts:
        s=_v118_fix_surface(s)
        if not s: continue
        if chosen and (len(chosen)>=3 or total+len(s)>720): break
        chosen.append(s); total+=len(s)+1
    if len(chosen)==1:
        # Son güvenlik ağı: başlığı çıplak bırakmak yerine eksikliği görünür kılan ikinci cümle.
        chosen.append('Mevcut açık kaynak kaydında gelişmeye ilişkin ayrıntılı bilgi sınırlı düzeydedir.')
    text=' '.join(chosen).strip()
    return _v98_strip_site_name(text,source).strip()


def make_important_basket_docx_v101(basket_df):
    """V118 ÖGN — karışmış cümleleri ayıklar; her madde en az anlamlı ve bütünlüklü iki cümle hedefler."""
    doc=_v116_doc_defaults(Document(),top=2.25,bottom=1.50,left=1.905,right=1.905)
    sec=doc.sections[0]; sec.header_distance=Cm(0.1); sec.footer_distance=Cm(0.45)
    try:
        ns=doc.styles['No Spacing']; ns.font.name='Times New Roman'; ns.font.size=Pt(12)
        ns._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'Times New Roman')
    except Exception: pass
    _v116_ogn_footer(sec)
    today=datetime.now().astimezone().date(); yesterday=today-timedelta(days=1)
    p=doc.add_paragraph(style='No Spacing'); p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.line_spacing=1.15; p.paragraph_format.space_before=Pt(6); p.paragraph_format.space_after=Pt(6)
    _v116_run(p.add_run(f'{yesterday.strftime("%d/%m/%Y")} – {today.strftime("%d/%m/%Y")}'),12)
    p=doc.add_paragraph(style='No Spacing'); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing=1.15; p.paragraph_format.space_before=Pt(6); p.paragraph_format.space_after=Pt(6)
    _v116_run(p.add_run('Konu: '),12,True); _v116_run(p.add_run('STB Temsilciliği Önemli Gelişmeler Notu'),12)
    rows=_v115_dedupe_rows([] if basket_df is None else basket_df.to_dict('records'))
    outputs=[None]*len(rows)
    if rows:
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(3,len(rows))) as ex:
            futs={ex.submit(_v118_ogn_item,r):i for i,r in enumerate(rows)}
            for fut in concurrent.futures.as_completed(futs):
                try: outputs[futs[fut]]=fut.result()
                except Exception: outputs[futs[fut]]=''
    for text in outputs:
        text=_v118_fix_surface(text)
        if not text: continue
        p=doc.add_paragraph(style='No Spacing'); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing=1.15; p.paragraph_format.space_before=Pt(6); p.paragraph_format.space_after=Pt(6)
        if text.endswith('.'): text=text[:-1]
        _v116_run(p.add_run(text+' (STB).'),12)
    p=doc.add_paragraph(style='No Spacing'); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing=1.15; p.paragraph_format.space_before=Pt(6); p.paragraph_format.space_after=Pt(6)
    _v116_run(p.add_run('Arz olunur.'),12)
    bio=BytesIO(); doc.save(bio); bio.seek(0); return bio.getvalue()


def _v118_akt_summary(title,body):
    facts=_v118_formal_facts(title,body,limit=6)
    if not facts:
        tf=_v118_title_fact(title)
        return tf.rstrip('.') if tf else ''
    clauses=[]
    for s in facts:
        s=_v118_fix_surface(s).strip().rstrip(' .;:')
        if s: clauses.append(s)
    return '; '.join(clauses)


def _v118_prepare_akt_row(row):
    tr,local_text=_v117_local_context(row)
    detail=article_detail(tr)
    real_url=str(detail.get('canonical') or tr.get('URL') or '').strip()
    source=_v117_source_short(_real_source(tr,detail,real_url),real_url)
    title=_v116_clean_headline(detail.get('title') or tr.get('Başlık',''),source)
    body=_clean_note_text(detail.get('text') or local_text or tr.get('İçerik_Özeti') or '')
    summary=_v118_akt_summary(title,body)
    if not summary:
        summary=_v118_title_fact(title).rstrip('.')
    if not _v116_valid_direct_url(real_url): real_url=str(tr.get('URL') or '')
    image=None
    for candidate in list(detail.get('images') or [])[:8]:
        image=_v117_valid_report_image(candidate)
        if image: break
    return {'title':title,'source':source,'url':real_url,'summary':summary,'image':image}


def make_docx(rows):
    """V118 AKT — referans görünüm; başlık/ara başlık kirliliği azaltılmış, cümle yüzeyi düzeltilmiş."""
    rows0=[]
    for r in (rows or []):
        tr,_=_v117_local_context(r); rows0.append(tr)
    rows=_v115_dedupe_rows(rows0)
    doc=_v116_doc_defaults(Document(),top=2.5,bottom=1.25,left=2.5,right=2.5)
    sec=doc.sections[0]; sec.header_distance=Cm(1.25); sec.footer_distance=Cm(1.25)
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(0)
    _v116_run(p.add_run('AÇIK KAYNAK TARAMA ÇALIŞMASI'),16,True)
    _v116_add_akt_info_table(doc)
    intro=doc.add_paragraph(); intro.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    intro.paragraph_format.first_line_indent=Cm(0.63); intro.paragraph_format.line_spacing=1.5
    intro.paragraph_format.space_before=Pt(0); intro.paragraph_format.space_after=Pt(0)
    intro_text=_v116_akt_intro_text(rows); topics=_v116_topic_labels(rows)
    cursor=0; spans=[]; pos=0
    for topic in topics:
        q=f'“{topic}”'; ix=intro_text.find(q,pos)
        if ix>=0: spans.append((ix,ix+len(q))); pos=ix+len(q)
    for a,b in spans:
        if a>cursor: _v116_run(intro.add_run(intro_text[cursor:a]),12)
        _v116_run(intro.add_run(intro_text[a:b]),12,italic=True); cursor=b
    if cursor<len(intro_text): _v116_run(intro.add_run(intro_text[cursor:]),12)
    prepared=[None]*len(rows)
    if rows:
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(3,len(rows))) as ex:
            futs={ex.submit(_v118_prepare_akt_row,r):i for i,r in enumerate(rows)}
            for fut in concurrent.futures.as_completed(futs):
                try: prepared[futs[fut]]=fut.result()
                except Exception: prepared[futs[fut]]=None
    out_no=0
    for item in prepared:
        if not item: continue
        out_no+=1
        p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent=Cm(0.63); p.paragraph_format.line_spacing=1.5
        p.paragraph_format.space_before=Pt(4); p.paragraph_format.space_after=Pt(6)
        _v116_run(p.add_run(f'{out_no}. “{item["source"]}”'),12,True)
        _v116_run(p.add_run(' isimli internet sitesinde, '),12)
        _v116_run(p.add_run(f'“{item["title"]}”'),12,True,True)
        _v116_run(p.add_run(' başlığıyla bir haber yayımlanmıştır. ('),12)
        _word_hyperlink(p,item['url'],item['url'] or 'Haber bağlantısı')
        _v116_run(p.add_run(') Söz konusu haber içeriğinde, '),12)
        _v116_run(p.add_run(_v118_fix_surface(item.get('summary') or '').strip().rstrip(' .;')),12)
        _v116_run(p.add_run(' hususları ifade edilmiştir.'),12)
        if item.get('image'):
            cap=doc.add_paragraph(); cap.alignment=WD_ALIGN_PARAGRAPH.CENTER
            cap.paragraph_format.first_line_indent=Cm(0.63); cap.paragraph_format.line_spacing=1.5
            cap.paragraph_format.space_before=Pt(4); cap.paragraph_format.space_after=Pt(4)
            _v116_run(cap.add_run(f'Görsel {out_no}: “{item["source"]}” Sitesinde Yer Alan Görsel'),12,True)
            ip=doc.add_paragraph(); ip.alignment=WD_ALIGN_PARAGRAPH.CENTER; ip.paragraph_format.space_after=Pt(8)
            try:
                run=ip.add_run(); shape=run.add_picture(item['image'])
                max_w=Cm(15.3); max_h=Cm(12.7); scale=min(1.0,max_w/shape.width,max_h/shape.height)
                if scale<1.0:
                    shape.width=int(shape.width*scale); shape.height=int(shape.height*scale)
            except Exception: pass
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY; p.paragraph_format.line_spacing=1.5
    _v116_run(p.add_run('Arz olunur.'),12)
    bio=BytesIO(); doc.save(bio); bio.seek(0); return bio.getvalue()


# Eski rapor byte'larının yeni motorla karışmasını önle.
if st.session_state.get('_report_engine_version') != _V118_ENGINE_VERSION:
    for _k in ('docx_bytes','note_bytes','basket_docx_bytes','v79_akt_note_bytes','v81_pres_note_bytes'):
        st.session_state.pop(_k,None)
    for _k in ('_v117_article_cache','_v117_page_cache'):
        st.session_state.pop(_k,None)
    st.session_state['_report_engine_version']=_V118_ENGINE_VERSION

# ============================================================
# /V118
# ============================================================

# ============================================================
# V119 — FINAL QUALITY / PERFORMANCE ENGINE
# ============================================================
# Goals
# -----
# 1. Vardiya Başlangıç Özeti must not block normal page rendering.
# 2. "Dünden Beri Ne Değişti?" uses an O(n) / indexed comparison path.
# 3. Report generation uses a common evidence-enrichment pipeline.
# 4. Weak, headline-only reports are rejected instead of producing filler text.
# 5. Turkish surface normalization must not corrupt numeric suffixes
#    (e.g. 476.934'e -> 476.934'E).
# ============================================================

_V119_ENGINE_VERSION = "V119-FINAL-2026-09-21-A"


class ReportQualityError(ValueError):
    """Raised when there is not enough verified text to build a proper report."""


def _v119_turkish_lower_char(char):
    table = str.maketrans({"I": "ı", "İ": "i"})
    return str(char).translate(table).lower()


def _v119_fix_surface(text):
    """Normalize report prose without damaging numbers, units or Turkish suffixes."""
    value = _clean_note_text(text)
    if not value:
        return ""

    value = re.sub(r"\s+([,.;:!?])", r"\1", value)
    value = re.sub(r"([({\[])\s+", r"\1", value)
    value = re.sub(r"\s+([)}\]])", r"\1", value)
    value = re.sub(r"\s+(['’])", r"\1", value)
    value = re.sub(r"(['’])\s+", r"\1", value)
    value = re.sub(r"\s{2,}", " ", value).strip()

    # A previous surface normalizer treated every period as a sentence boundary,
    # which turned 476.934'e into 476.934'E and 185.513 adet into 185.513 Adet.
    value = re.sub(
        r"(?<=\d)(['’])([A-ZÇĞİÖŞÜ])\b",
        lambda match: match.group(1) + _v119_turkish_lower_char(match.group(2)),
        value,
    )
    value = re.sub(
        r"(?<=\d\.\d{3})\s+(Adet|Olarak|Oranında|Seviyesinde|İse)\b",
        lambda match: " " + match.group(1).lower(),
        value,
        flags=re.IGNORECASE,
    )

    entities = (
        ("türkiye", "Türkiye"),
        ("rusya", "Rusya"),
        ("ukrayna", "Ukrayna"),
        ("amerika birleşik devletleri", "Amerika Birleşik Devletleri"),
        ("avrupa birliği", "Avrupa Birliği"),
        ("google", "Google"),
        ("apple", "Apple"),
        ("microsoft", "Microsoft"),
        ("openai", "OpenAI"),
        ("anthropic", "Anthropic"),
        ("meta", "Meta"),
    )
    for old, new in entities:
        value = re.sub(
            rf"(?<!\w){re.escape(old)}(?!\w)",
            new,
            value,
            flags=re.IGNORECASE,
        )
    value = re.sub(r"(?<!\w)aI(?!\w)", "AI", value)

    # Remove typical by-line remnants that were observed in generated documents.
    value = re.sub(
        r"\b[A-ZÇĞİÖŞÜ]{2,}(?:\s+[A-ZÇĞİÖŞÜ]{2,}){1,4}\s*[-–—]\s*"
        r"(?=[A-ZÇĞİÖŞÜ][a-zçğıöşü])",
        "",
        value,
    )

    # Capitalize only real sentence starts. Numeric decimal/thousand separators
    # are intentionally not treated as boundaries.
    value = re.sub(
        r"(^|(?<=[!?])\s+|(?<=\.)\s+)([a-zçğıöşü])",
        lambda match: match.group(1) + match.group(2).upper(),
        value,
    )
    return re.sub(r"\s{2,}", " ", value).strip()


# All later V117/V118 helpers resolve this name dynamically.
_v118_fix_surface = _v119_fix_surface


def _v119_direct_scan_time(scan_id):
    """Read only scan metadata; do not load a whole snapshot for one timestamp."""
    if not scan_id:
        return None
    try:
        with _history_connect() as conn:
            row = conn.execute(
                "SELECT scanned_at FROM scans WHERE scan_id=? LIMIT 1",
                (int(scan_id),),
            ).fetchone()
        return str(row[0]) if row else None
    except Exception:
        return None


def _shift_baseline(current_scan_id=None):
    """V119 lightweight shift baseline."""
    mark = _latest_shift_mark()
    if mark:
        try:
            timestamp = pd.to_datetime(mark["marked_at"], utc=True)
            return (
                timestamp,
                f"Devir noktası: {mark['marked_at']}",
                mark.get("scan_id"),
            )
        except Exception:
            pass

    previous_id = _previous_scan_id(current_scan_id)
    scanned_at = _v119_direct_scan_time(previous_id)
    if scanned_at:
        try:
            return (
                pd.to_datetime(scanned_at, utc=True),
                f"Önceki tarama: {scanned_at}",
                previous_id,
            )
        except Exception:
            pass
    return None, "Henüz devir noktası yok", None


def _v119_current_events(df):
    """Create one representative row per already-clustered event without reclustering."""
    columns = [
        "event_id",
        "title",
        "source",
        "url",
        "category",
        "summary",
        "risk_score",
        "risk_status",
        "verification",
        "source_count",
    ]
    if df is None or df.empty:
        return pd.DataFrame(columns=columns)

    frame = df.copy()

    def series(name, default=""):
        if name in frame.columns:
            return frame[name]
        return pd.Series(default, index=frame.index)

    frame["Tarih_dt"] = pd.to_datetime(
        series("Tarih_dt"), utc=True, errors="coerce"
    )
    frame = frame.sort_values("Tarih_dt", ascending=False, na_position="last")
    group_column = "Olay_ID" if "Olay_ID" in frame.columns else None

    if group_column:
        reps = frame.drop_duplicates(group_column, keep="first").copy()
        risk_values = pd.to_numeric(
            series("Risk_Skoru", 0), errors="coerce"
        ).fillna(0)
        source_values = pd.to_numeric(
            series("Olay_Kaynak_Sayisi", 1), errors="coerce"
        ).fillna(1)
        risk_max = risk_values.groupby(frame[group_column]).max()
        source_max = source_values.groupby(frame[group_column]).max()
        reps["_risk_max"] = reps[group_column].map(risk_max).fillna(0)
        reps["_source_max"] = reps[group_column].map(source_max).fillna(1)
        event_ids = reps[group_column].astype(str)
    else:
        reps = frame.copy()
        reps["_risk_max"] = pd.to_numeric(
            series("Risk_Skoru", 0), errors="coerce"
        ).fillna(0)
        reps["_source_max"] = pd.to_numeric(
            series("Olay_Kaynak_Sayisi", 1), errors="coerce"
        ).fillna(1)
        event_ids = pd.Series(
            [f"ROW-{index}" for index in reps.index], index=reps.index
        )

    def rep_series(name, default=""):
        if name in reps.columns:
            return reps[name].fillna(default).astype(str)
        return pd.Series(str(default), index=reps.index)

    return pd.DataFrame(
        {
            "event_id": event_ids.values,
            "title": rep_series("Başlık").values,
            "source": rep_series("Kaynak").values,
            "url": rep_series("URL").values,
            "category": rep_series("Kategori").values,
            "summary": rep_series("İçerik_Özeti").values,
            "risk_score": reps["_risk_max"].astype(int).values,
            "risk_status": rep_series("Risk_Durumu").values,
            "verification": rep_series("Doğrulama").values,
            "source_count": reps["_source_max"].astype(int).values,
        }
    )


def _v119_previous_events(scan_id):
    """Load only fields required for indexed previous-scan comparison."""
    if not scan_id:
        return pd.DataFrame()
    key = f"v119_prev:{scan_id}"
    cache = st.session_state.setdefault("_v119_previous_event_cache", {})
    if key in cache:
        return cache[key].copy()
    try:
        with _history_connect() as conn:
            previous = pd.read_sql_query(
                """
                SELECT
                    event_id, title, source, url, category,
                    substr(summary, 1, 2400) AS summary,
                    risk_score, risk_status, verification, source_count,
                    tokens_json
                FROM event_snapshots
                WHERE scan_id=?
                """,
                conn,
                params=(int(scan_id),),
            )
    except Exception:
        previous = pd.DataFrame()
    cache.clear()
    cache[key] = previous.copy()
    return previous


def _v119_candidate_indices(tokens, token_index, limit=14):
    """Use the rarest title tokens first to keep candidate sets very small."""
    ranked = sorted(
        (token for token in tokens if token in token_index),
        key=lambda token: len(token_index[token]),
    )
    candidates = set()
    for token in ranked[:6]:
        candidates.update(token_index[token])
        if len(candidates) >= limit:
            break
    if len(candidates) <= limit:
        return candidates
    return set(sorted(candidates)[:limit])


def _compare_since_previous(df, current_scan_id=None):
    """V119 indexed change comparison; avoids heavy event reclustering on page render."""
    columns = [
        "Ne Değişti?",
        "Tür",
        "Değişim",
        "Başlık",
        "Kaynak",
        "Kategori",
        "Risk",
        "Önceki Risk",
        "Kaynak Sayısı",
        "URL",
    ]
    cache_key = _v113_scan_key(df, current_scan_id) + ":v119"
    cache = st.session_state.setdefault("_v119_compare_cache", {})
    if cache_key in cache:
        output, previous_id, previous_time = cache[cache_key]
        return output.copy(), previous_id, previous_time

    current = _v119_current_events(df)
    previous_id = _previous_scan_id(current_scan_id)
    previous = _v119_previous_events(previous_id)
    previous_time = _v119_direct_scan_time(previous_id)

    if current.empty:
        return pd.DataFrame(columns=columns), None, None
    if previous.empty:
        result = pd.DataFrame(columns=columns)
        cache.clear()
        cache[cache_key] = (result.copy(), previous_id, previous_time)
        return result, previous_id, previous_time

    previous_records = []
    token_index = {}
    url_index = {}
    title_index = {}
    for index, row in previous.iterrows():
        record = row.to_dict()
        try:
            tokens = set(json.loads(record.get("tokens_json") or "[]"))
        except Exception:
            tokens = set(_history_tokens(record.get("title", "")))
        if not tokens:
            tokens = set(_history_tokens(record.get("title", "")))
        record["_tokens"] = tokens
        previous_records.append(record)
        record_index = len(previous_records) - 1
        for token in tokens:
            token_index.setdefault(token, set()).add(record_index)
        url = str(record.get("url") or "").strip()
        if url:
            url_index.setdefault(url, set()).add(record_index)
        key = title_key(record.get("title", ""))
        if key:
            title_index.setdefault(key, set()).add(record_index)

    changes = []
    for _, row in current.iterrows():
        current_record = row.to_dict()
        title = current_record.get("title", "")
        summary = current_record.get("summary", "")
        url = str(current_record.get("url") or "").strip()
        title_tokens = set(_history_tokens(title))

        candidates = set()
        if url:
            candidates.update(url_index.get(url, set()))
        current_title_key = title_key(title)
        if current_title_key:
            candidates.update(title_index.get(current_title_key, set()))
        candidates.update(
            _v119_candidate_indices(title_tokens, token_index, limit=14)
        )

        best = None
        best_similarity = 0.0
        for candidate_index in candidates:
            previous_record = previous_records[candidate_index]
            if url and url == str(previous_record.get("url") or ""):
                similarity = 1.0
            elif current_title_key and current_title_key == title_key(
                previous_record.get("title", "")
            ):
                similarity = 0.98
            else:
                previous_tokens = previous_record.get("_tokens", set())
                union = len(title_tokens | previous_tokens)
                jaccard = (
                    len(title_tokens & previous_tokens) / union if union else 0.0
                )
                similarity = jaccard
                if len(title_tokens & previous_tokens) >= 3:
                    similarity += 0.12
            if similarity > best_similarity:
                best_similarity = similarity
                best = previous_record

        risk_score = int(current_record.get("risk_score") or 0)
        if best is None or best_similarity < 0.43:
            kind = "🆕 YENİ OLAY"
            previous_risk = "—"
            difference = (
                "Bu olaya ilişkin kayıt önceki karşılaştırma taramasında "
                "bulunmamaktadır."
            )
            priority = 100 + risk_score
        else:
            comparable = {
                "title": title,
                "summary": summary,
                "risk_score": risk_score,
                "risk_status": str(current_record.get("risk_status") or ""),
                "verification": str(current_record.get("verification") or ""),
                "source_count": int(current_record.get("source_count") or 1),
            }
            risk_up, verify_up, material, _, _ = _v104_material_change(
                best, comparable
            )
            if risk_up:
                kind = "⚠️ RİSK ARTTI"
                priority = 95 + risk_score
            elif verify_up:
                kind = "✅ TEYİT GÜÇLENDİ"
                priority = 90 + risk_score
            elif material:
                kind = "🔄 YENİ BİLGİ"
                priority = 80 + risk_score
            else:
                continue
            previous_risk = int(best.get("risk_score") or 0)
            difference = _v109_direct_difference(best, comparable, kind)

        changes.append(
            {
                "Ne Değişti?": difference,
                "Tür": kind,
                "Değişim": kind,
                "Başlık": title,
                "Kaynak": current_record.get("source", ""),
                "Kategori": current_record.get("category", ""),
                "Risk": risk_score,
                "Önceki Risk": previous_risk,
                "Kaynak Sayısı": int(current_record.get("source_count") or 1),
                "URL": current_record.get("url", ""),
                "_priority": priority,
            }
        )

    output = pd.DataFrame(changes)
    if output.empty:
        output = pd.DataFrame(columns=columns)
    else:
        output = output.sort_values(
            ["_priority", "Risk"], ascending=[False, False]
        ).drop(columns=["_priority"], errors="ignore")
        output = output.drop_duplicates(
            subset=["Tür", "Başlık", "URL"], keep="first"
        ).reset_index(drop=True)
        output = output[columns]

    cache.clear()
    cache[cache_key] = (output.copy(), previous_id, previous_time)
    return output, previous_id, previous_time


def _shift_start_summary(df, current_scan_id=None):
    """V119 shift summary: all expensive work is cached/precomputed before render."""
    if df is None or df.empty:
        return {}, pd.DataFrame(), ""

    key = (
        _v113_scan_key(df, current_scan_id)
        + ":"
        + _v113_shift_mark_key()
        + ":v119"
    )
    cached = _v113_get_cached("shift_summary_v119", key)
    if cached is not None:
        stats, top, label = cached
        return dict(stats), top.copy(), label

    baseline, baseline_label, _ = _shift_baseline(current_scan_id)
    frame = df.copy()
    frame["Tarih_dt"] = pd.to_datetime(
        frame.get("Tarih_dt"), utc=True, errors="coerce"
    )
    if baseline is not None:
        since = frame[
            frame["Tarih_dt"].isna() | (frame["Tarih_dt"] >= baseline)
        ].copy()
    else:
        since = frame.copy()

    changes, _, _ = _compare_since_previous(df, current_scan_id)
    if changes.empty:
        new_events = risk_up = verify_up = 0
    else:
        kinds = changes["Tür"].astype(str)
        new_events = int(kinds.str.contains("YENİ OLAY").sum())
        risk_up = int(kinds.str.contains("RİSK ARTTI").sum())
        verify_up = int(kinds.str.contains("TEYİT").sum())

    high_risk = int(
        (
            since.get("Risk_Durumu", pd.Series("", index=since.index))
            == "Yüksek Risk"
        ).sum()
    )
    titles = since.get("Başlık", pd.Series("", index=since.index)).fillna("")
    summaries = since.get(
        "İçerik_Özeti", pd.Series("", index=since.index)
    ).fillna("")
    osb_count = sum(
        bool(is_osb_fire(title, summary))
        for title, summary in zip(titles.astype(str), summaries.astype(str))
    )
    top = _v115_fast_shift_top(since, 8)
    stats = {
        "new_news": len(since),
        "new_important_events": new_events,
        "high_risk": high_risk,
        "risk_up": risk_up,
        "verify_up": verify_up,
        "osb": int(osb_count),
        "baseline_label": baseline_label,
    }
    _v113_set_cached(
        "shift_summary_v119", key, (dict(stats), top.copy(), baseline_label)
    )
    return stats, top, baseline_label


def _v119_search_evidence(title, preferred_url=""):
    """Collect bounded fallback evidence for sources that expose only a headline."""
    title = _clean_note_text(title)
    if not title:
        return []

    cache_key = hashlib.sha1(
        f"{title}|{preferred_url}".encode("utf-8", "ignore")
    ).hexdigest()
    cache = st.session_state.setdefault("_v119_evidence_cache", {})
    if cache_key in cache:
        return list(cache[cache_key])

    evidence = []

    def add(text, url="", source="", weight=0):
        cleaned = _v119_fix_surface(text)
        if not cleaned or norm(cleaned) == norm(title):
            return
        if len(cleaned) < 45:
            return
        relevance = _v118_title_relevance(title, cleaned)
        if relevance < 1.8:
            return
        evidence.append(
            {
                "text": cleaned,
                "url": str(url or "").strip(),
                "source": _clean_note_text(source),
                "weight": int(weight),
                "relevance": float(relevance),
            }
        )

    # Google News RSS is already part of the application's normal source stack.
    # It is useful as a fallback because the selected publisher may block scraping.
    queries = [f'"{title[:210]}"']
    compact_terms = list(_v117_tokens(title))[:8]
    if compact_terms:
        queries.append(" ".join(compact_terms))

    seen_links = set()
    for query in queries[:2]:
        try:
            results = rss(query, timeout=5)[:8]
        except Exception:
            results = []
        for item in results:
            item_title = _clean_note_text(item.get("title", ""))
            if _v118_title_relevance(title, item_title) < 2.0:
                continue
            add(
                item.get("snippet", ""),
                item.get("url", ""),
                item.get("source", ""),
                250,
            )
            google_url = str(item.get("url") or "").strip()
            if not google_url or google_url in seen_links:
                continue
            seen_links.add(google_url)
            try:
                direct_url = _v116_decode_google_news_url(google_url)
            except Exception:
                direct_url = ""
            if direct_url and _v116_valid_direct_url(direct_url):
                page = _v117_fetch_clean_page(direct_url, title)
                add(
                    page.get("text", ""),
                    direct_url,
                    page.get("source", "") or item.get("source", ""),
                    700,
                )

    # Optional DDG package path, if available on the deployment.
    try:
        for result in ddgs_text(f'"{title[:210]}"')[:8]:
            add(
                result.get("body", ""),
                result.get("href", "") or result.get("url", ""),
                result.get("title", ""),
                300,
            )
    except Exception:
        pass

    # Existing HTML-search fallback. Fetch only a few highly relevant pages.
    try:
        for url in _v118_search_urls(title, preferred_url)[:3]:
            page = _v117_fetch_clean_page(url, title)
            add(page.get("text", ""), url, page.get("source", ""), 550)
    except Exception:
        pass

    # Existing search snippets are cheap and can rescue blocked publisher pages.
    try:
        for snippet in _v117_search_snippets(title, preferred_url)[:4]:
            add(snippet, "", "", 180)
    except Exception:
        pass

    evidence.sort(
        key=lambda item: (
            item["weight"],
            item["relevance"],
            len(item["text"]),
        ),
        reverse=True,
    )
    unique = []
    seen = set()
    for item in evidence:
        key = title_key(item["text"])
        if key in seen:
            continue
        seen.add(key)
        unique.append(item)
        if len(unique) >= 8:
            break

    cache.clear()
    cache[cache_key] = list(unique)
    return unique


_v119_article_detail_base = article_detail


def _v119_strip_title_echo(title, sentence):
    """Remove a headline or colon-suffix repeated at the start of article text."""
    clean_title = _clean_note_text(title)
    clean_sentence = _clean_note_text(sentence)
    segments = [clean_title]
    if ':' in clean_title:
        segments.append(clean_title.split(':', 1)[1].strip())
    for separator in (' - ', ' – ', ' — '):
        if separator in clean_title:
            segments.append(clean_title.split(separator, 1)[0].strip())
    for segment in sorted(set(segments), key=len, reverse=True):
        if len(segment) < 24:
            continue
        pattern = r'^\s*' + r'\s+'.join(
            re.escape(part) for part in segment.split()
        )
        match = re.match(pattern, clean_sentence, flags=re.IGNORECASE)
        if match:
            remainder = clean_sentence[match.end():].lstrip(' .;:–—-')
            if len(remainder) >= 30:
                return remainder
    return clean_sentence


def _v119_fact_texts(title, body, limit=12):
    """Return de-duplicated, formal, event-related factual sentences."""
    facts = []
    seen_tokens = []
    for candidate in _v118_strict_facts(title, body, max_items=limit * 2):
        raw_sentence = _v119_strip_title_echo(
            title, candidate.get("text", "")
        )
        sentence = _v119_fix_surface(
            _v117_formal_sentence(raw_sentence)
        )
        if not sentence or len(sentence) < 35:
            continue
        normalized = norm(sentence)
        if any(
            marker in normalized
            for marker in (
                "mevcut açık kaynak kaydında",
                "mevcut veriler çerçevesinde bilgi notu",
                "ayrıntılı bilgi sınırlı",
                "yeterli gövde metni",
            )
        ):
            continue
        if re.search(r"\b[A-ZÇĞİÖŞÜ]{3,}(?:\s+[A-ZÇĞİÖŞÜ]{3,}){2,}\b", sentence):
            continue
        tokens = _v117_tokens(sentence)
        duplicate = False
        for old_tokens in seen_tokens:
            union = len(tokens | old_tokens)
            if union and len(tokens & old_tokens) / union >= 0.72:
                duplicate = True
                break
        if duplicate:
            continue
        facts.append(sentence)
        seen_tokens.append(tokens)
        if len(facts) >= limit:
            break
    return facts


def article_detail(row):
    """V119 common evidence pipeline used by AKT, ÖGN and information notes."""
    if isinstance(row, str):
        row = {"URL": row}
    elif hasattr(row, "to_dict"):
        row = row.to_dict()
    elif row is None:
        row = {}
    else:
        row = dict(row)

    tr, local_text = _v117_local_context(row)
    title_hint = _clean_note_text(tr.get("Başlık", ""))
    cache_key = hashlib.sha1(
        (
            title_hint
            + "|"
            + str(tr.get("URL", ""))
            + "|"
            + local_text[:900]
            + "|v119"
        ).encode("utf-8", "ignore")
    ).hexdigest()
    cache = st.session_state.setdefault("_v119_article_cache", {})
    if cache_key in cache:
        return dict(cache[cache_key])

    try:
        base = _v119_article_detail_base(tr) or {}
    except Exception:
        base = {}

    source = _v117_source_short(
        base.get("source") or tr.get("Kaynak", ""),
        base.get("canonical") or tr.get("URL", ""),
    )
    title = _v116_clean_headline(
        base.get("title") or title_hint,
        source,
    )
    canonical = str(base.get("canonical") or tr.get("URL") or "").strip()

    text_candidates = []

    def add_candidate(text, bonus, candidate_url="", candidate_source=""):
        cleaned = _clean_note_text(text)
        if not cleaned or norm(cleaned) == norm(title):
            return
        facts = _v119_fact_texts(title, cleaned, limit=12)
        if not facts:
            return
        score = (
            len(facts) * 1200
            + min(len(cleaned), 6500)
            + int(bonus)
            + sum(40 for fact in facts if re.search(r"\d", fact))
        )
        text_candidates.append(
            {
                "score": score,
                "text": cleaned,
                "facts": facts,
                "url": candidate_url,
                "source": candidate_source,
            }
        )

    add_candidate(base.get("text", ""), 500, canonical, source)
    add_candidate(local_text, 650, tr.get("URL", ""), tr.get("Kaynak", ""))
    add_candidate(
        tr.get("İçerik_Özeti", ""),
        180,
        tr.get("URL", ""),
        tr.get("Kaynak", ""),
    )

    best_fact_count = max(
        (len(item["facts"]) for item in text_candidates), default=0
    )
    if best_fact_count < 4:
        for item in _v119_search_evidence(title, canonical or tr.get("URL", "")):
            add_candidate(
                item["text"],
                item["weight"],
                item.get("url", ""),
                item.get("source", ""),
            )

    if text_candidates:
        text_candidates.sort(key=lambda item: item["score"], reverse=True)
        best = text_candidates[0]
        # If the best source is still thin, safely merge only de-duplicated facts
        # from the next sources. This avoids headline-only notes without inventing facts.
        merged_facts = list(best["facts"])
        merged_tokens = [_v117_tokens(fact) for fact in merged_facts]
        for candidate in text_candidates[1:5]:
            for fact in candidate["facts"]:
                tokens = _v117_tokens(fact)
                if any(
                    len(tokens | old) > 0
                    and len(tokens & old) / len(tokens | old) >= 0.70
                    for old in merged_tokens
                ):
                    continue
                if _v118_title_relevance(title, fact) < 1.5:
                    continue
                merged_facts.append(fact)
                merged_tokens.append(tokens)
                if len(merged_facts) >= 10:
                    break
            if len(merged_facts) >= 10:
                break
        final_text = " ".join(merged_facts)
        if best.get("url") and _v116_valid_direct_url(best.get("url")):
            canonical = best["url"]
        if best.get("source"):
            source = _v117_source_short(best["source"], canonical)
    else:
        final_text = ""
        merged_facts = []

    result = dict(base)
    result["title"] = title or title_hint
    result["source"] = source
    result["canonical"] = canonical
    result["text"] = final_text
    result["quality_facts"] = len(merged_facts)
    result["quality_ready"] = (
        len(merged_facts) >= 3 and len(final_text) >= 180
    )
    cache[cache_key] = dict(result)
    return result


def _v119_note_paragraphs(title, body):
    """Build natural introduction-development-conclusion paragraphs from evidence only."""
    facts = _v119_fact_texts(title, body, limit=10)
    if len(facts) < 3:
        raise ReportQualityError(
            f'"{title}" için en az üç bağımsız olgu elde edilemedi. '
            "Başlık tekrarı veya yapay dolgu üretmek yerine rapor oluşturulmadı."
        )

    if len(facts) >= 8:
        groups = [facts[:2], facts[2:5], facts[5:7], facts[7:]]
    elif len(facts) >= 6:
        groups = [facts[:2], facts[2:4], facts[4:]]
    elif len(facts) >= 4:
        groups = [facts[:1], facts[1:3], facts[3:]]
    else:
        groups = [[facts[0]], [facts[1]], [facts[2]]]

    paragraphs = []
    for group in groups:
        paragraph = " ".join(item for item in group if item).strip()
        if paragraph:
            paragraphs.append(_v119_fix_surface(paragraph))
    return paragraphs[:5]


def make_analyst_docx(df, title="BİLGİ NOTU"):
    """V119: quality-gated information note with consistent 3–5 paragraph structure."""
    frame = df.copy() if df is not None else pd.DataFrame()
    rows = [] if frame.empty else _v115_dedupe_rows(frame.to_dict("records"))
    if not rows:
        raise ReportQualityError("Bilgi notu oluşturmak için haber bulunamadı.")

    all_paragraphs = []
    weak_titles = []
    for row in rows:
        tr, local_text = _v117_local_context(row)
        detail = article_detail(tr)
        source = _v117_source_short(
            _real_source(tr, detail, detail.get("canonical", "")),
            detail.get("canonical", ""),
        )
        clean_title = _v116_clean_headline(
            detail.get("title") or tr.get("Başlık", ""), source
        )
        body = _clean_note_text(
            detail.get("text") or local_text or tr.get("İçerik_Özeti") or ""
        )
        try:
            paragraphs = _v119_note_paragraphs(clean_title, body)
        except ReportQualityError:
            weak_titles.append(clean_title)
            continue
        all_paragraphs.extend(paragraphs)

    if weak_titles:
        names = "; ".join(weak_titles[:3])
        raise ReportQualityError(
            "Bilgi notu kalite eşiği sağlanamadı: " + names + ". "
            "Sistem eksik bilgiyi uydurmadı; farklı bir haber seçin veya yeniden deneyin."
        )
    if len(all_paragraphs) < 3:
        raise ReportQualityError(
            "Bilgi notu için yeterli doğrulanabilir içerik elde edilemedi."
        )

    document = _v116_doc_defaults(Document(), 2.5, 2.5, 2.5, 2.5)
    section = document.sections[0]
    section.header_distance = Cm(1.25)
    section.footer_distance = Cm(1.25)
    for paragraph_text in all_paragraphs[:5]:
        paragraph = document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(6)
        paragraph.paragraph_format.line_spacing = 1.15
        _v116_run(paragraph.add_run(_v119_fix_surface(paragraph_text)), 12)

    buffer = BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def _v119_ogn_item(row):
    tr, local_text = _v117_local_context(row)
    detail = article_detail(tr)
    source = _v117_source_short(
        _real_source(tr, detail, detail.get("canonical", "")),
        detail.get("canonical", ""),
    )
    title = _v116_clean_headline(
        detail.get("title") or tr.get("Başlık", ""), source
    )
    body = _clean_note_text(
        detail.get("text") or local_text or tr.get("İçerik_Özeti") or ""
    )
    facts = _v119_fact_texts(title, body, limit=5)
    if len(facts) < 2:
        raise ReportQualityError(
            f'ÖGN için "{title}" haberinde başlığın ötesinde yeterli içerik elde edilemedi.'
        )

    chosen = []
    total = 0
    for fact in facts:
        if chosen and (len(chosen) >= 3 or total + len(fact) > 760):
            break
        if title_key(fact.rstrip(".")) == title_key(title.rstrip(".")):
            continue
        chosen.append(fact)
        total += len(fact) + 1

    if len(chosen) < 2:
        # A rich first sentence may legitimately contain the whole development, but
        # a bare headline may not. Keep the quality threshold deterministic.
        if chosen and len(chosen[0]) >= 180 and re.search(r"\d", chosen[0]):
            return _v119_fix_surface(chosen[0])
        raise ReportQualityError(
            f'ÖGN için "{title}" maddesi yeterli ayrıntıya ulaşamadı.'
        )
    return _v119_fix_surface(" ".join(chosen))


def make_important_basket_docx_v101(basket_df):
    """V119 quality-gated ÖGN matching the supplied institutional template."""
    rows = _v115_dedupe_rows(
        [] if basket_df is None else basket_df.to_dict("records")
    )
    outputs = [None] * len(rows)
    errors = []
    if rows:
        with concurrent.futures.ThreadPoolExecutor(
            max_workers=min(4, len(rows))
        ) as executor:
            futures = {
                executor.submit(_v119_ogn_item, row): index
                for index, row in enumerate(rows)
            }
            for future in concurrent.futures.as_completed(futures):
                index = futures[future]
                try:
                    outputs[index] = future.result()
                except Exception as exc:
                    title = _clean_note_text(
                        rows[index].get("title", rows[index].get("Başlık", ""))
                    )
                    errors.append(f"{title}: {exc}")

    if errors:
        raise ReportQualityError(
            "Önemli Gelişmeler Notu oluşturulmadı. Kalite eşiğini geçemeyen "
            "maddeler: " + " | ".join(errors[:3])
        )

    document = _v116_doc_defaults(
        Document(), top=2.25, bottom=1.50, left=1.905, right=1.905
    )
    section = document.sections[0]
    section.header_distance = Cm(0.1)
    section.footer_distance = Cm(0.45)
    try:
        no_spacing = document.styles["No Spacing"]
        no_spacing.font.name = "Times New Roman"
        no_spacing.font.size = Pt(12)
        no_spacing._element.get_or_add_rPr().rFonts.set(
            qn("w:eastAsia"), "Times New Roman"
        )
    except Exception:
        pass
    _v116_ogn_footer(section)

    today = datetime.now().astimezone().date()
    yesterday = today - timedelta(days=1)
    paragraph = document.add_paragraph(style="No Spacing")
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    paragraph.paragraph_format.line_spacing = 1.15
    paragraph.paragraph_format.space_before = Pt(6)
    paragraph.paragraph_format.space_after = Pt(6)
    _v116_run(
        paragraph.add_run(
            f'{yesterday.strftime("%d/%m/%Y")} – {today.strftime("%d/%m/%Y")}'
        ),
        12,
    )

    paragraph = document.add_paragraph(style="No Spacing")
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    paragraph.paragraph_format.line_spacing = 1.15
    paragraph.paragraph_format.space_before = Pt(6)
    paragraph.paragraph_format.space_after = Pt(6)
    _v116_run(paragraph.add_run("Konu: "), 12, True)
    _v116_run(
        paragraph.add_run("STB Temsilciliği Önemli Gelişmeler Notu"), 12
    )

    for text in outputs:
        paragraph = document.add_paragraph(style="No Spacing")
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        paragraph.paragraph_format.line_spacing = 1.15
        paragraph.paragraph_format.space_before = Pt(6)
        paragraph.paragraph_format.space_after = Pt(6)
        clean = _v119_fix_surface(text).rstrip(".")
        _v116_run(paragraph.add_run(clean + " (STB)."), 12)

    paragraph = document.add_paragraph(style="No Spacing")
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    paragraph.paragraph_format.line_spacing = 1.15
    paragraph.paragraph_format.space_before = Pt(6)
    paragraph.paragraph_format.space_after = Pt(6)
    _v116_run(paragraph.add_run("Arz olunur."), 12)

    buffer = BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def _v119_akt_summary(title, body):
    facts = _v119_fact_texts(title, body, limit=6)
    if len(facts) < 3:
        raise ReportQualityError(
            f'AKT için "{title}" haberinde ayrıntılı içerik kalite eşiğinin altında kaldı.'
        )
    clauses = []
    for fact in facts:
        clean = _v119_fix_surface(fact).strip().rstrip(" .;:")
        if clean:
            clauses.append(clean)
    return "; ".join(clauses)


def _v119_prepare_akt_row(row):
    tr, local_text = _v117_local_context(row)
    detail = article_detail(tr)
    real_url = str(detail.get("canonical") or tr.get("URL") or "").strip()
    source = _v117_source_short(
        _real_source(tr, detail, real_url), real_url
    )
    title = _v116_clean_headline(
        detail.get("title") or tr.get("Başlık", ""), source
    )
    body = _clean_note_text(
        detail.get("text") or local_text or tr.get("İçerik_Özeti") or ""
    )
    summary = _v119_akt_summary(title, body)
    if not _v116_valid_direct_url(real_url):
        real_url = str(tr.get("URL") or "")

    image = None
    for candidate in list(detail.get("images") or [])[:8]:
        image = _v117_valid_report_image(candidate)
        if image:
            break
    return {
        "title": title,
        "source": source,
        "url": real_url,
        "summary": summary,
        "image": image,
    }


def make_docx(rows):
    """V119 quality-gated AKT report; weak headline-only items stop generation."""
    source_rows = []
    for row in rows or []:
        translated, _ = _v117_local_context(row)
        source_rows.append(translated)
    source_rows = _v115_dedupe_rows(source_rows)

    prepared = [None] * len(source_rows)
    errors = []
    if source_rows:
        with concurrent.futures.ThreadPoolExecutor(
            max_workers=min(4, len(source_rows))
        ) as executor:
            futures = {
                executor.submit(_v119_prepare_akt_row, row): index
                for index, row in enumerate(source_rows)
            }
            for future in concurrent.futures.as_completed(futures):
                index = futures[future]
                try:
                    prepared[index] = future.result()
                except Exception as exc:
                    title = _clean_note_text(source_rows[index].get("Başlık", ""))
                    errors.append(f"{title}: {exc}")

    if errors:
        raise ReportQualityError(
            "AKT raporu oluşturulmadı. Ayrıntılı içerik elde edilemeyen maddeler: "
            + " | ".join(errors[:3])
        )

    document = _v116_doc_defaults(
        Document(), top=2.5, bottom=1.25, left=2.5, right=2.5
    )
    section = document.sections[0]
    section.header_distance = Cm(1.25)
    section.footer_distance = Cm(1.25)

    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_after = Pt(0)
    _v116_run(paragraph.add_run("AÇIK KAYNAK TARAMA ÇALIŞMASI"), 16, True)
    _v116_add_akt_info_table(document)

    intro = document.add_paragraph()
    intro.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    intro.paragraph_format.first_line_indent = Cm(0.63)
    intro.paragraph_format.line_spacing = 1.5
    intro.paragraph_format.space_before = Pt(0)
    intro.paragraph_format.space_after = Pt(0)
    intro_text = _v116_akt_intro_text(source_rows)
    topics = _v116_topic_labels(source_rows)
    cursor = 0
    spans = []
    position = 0
    for topic in topics:
        quoted = f"“{topic}”"
        index = intro_text.find(quoted, position)
        if index >= 0:
            spans.append((index, index + len(quoted)))
            position = index + len(quoted)
    for start, end in spans:
        if start > cursor:
            _v116_run(intro.add_run(intro_text[cursor:start]), 12)
        _v116_run(intro.add_run(intro_text[start:end]), 12, italic=True)
        cursor = end
    if cursor < len(intro_text):
        _v116_run(intro.add_run(intro_text[cursor:]), 12)

    for number, item in enumerate(prepared, start=1):
        paragraph = document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        paragraph.paragraph_format.first_line_indent = Cm(0.63)
        paragraph.paragraph_format.line_spacing = 1.5
        paragraph.paragraph_format.space_before = Pt(4)
        paragraph.paragraph_format.space_after = Pt(6)
        _v116_run(paragraph.add_run(f'{number}. “{item["source"]}”'), 12, True)
        _v116_run(paragraph.add_run(" isimli internet sitesinde, "), 12)
        _v116_run(paragraph.add_run(f'“{item["title"]}”'), 12, True, True)
        _v116_run(paragraph.add_run(" başlığıyla bir haber yayımlanmıştır. ("), 12)
        _word_hyperlink(paragraph, item["url"], item["url"] or "Haber bağlantısı")
        _v116_run(paragraph.add_run(") Söz konusu haber içeriğinde, "), 12)
        _v116_run(
            paragraph.add_run(
                _v119_fix_surface(item["summary"]).strip().rstrip(" .;")
            ),
            12,
        )
        _v116_run(paragraph.add_run(" hususları ifade edilmiştir."), 12)

        if item.get("image"):
            caption = document.add_paragraph()
            caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
            caption.paragraph_format.first_line_indent = Cm(0.63)
            caption.paragraph_format.line_spacing = 1.5
            caption.paragraph_format.space_before = Pt(4)
            caption.paragraph_format.space_after = Pt(4)
            _v116_run(
                caption.add_run(
                    f'Görsel {number}: “{item["source"]}” Sitesinde Yer Alan Görsel'
                ),
                12,
                True,
            )
            image_paragraph = document.add_paragraph()
            image_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            image_paragraph.paragraph_format.space_after = Pt(8)
            try:
                run = image_paragraph.add_run()
                shape = run.add_picture(item["image"])
                max_width = Cm(15.3)
                max_height = Cm(12.7)
                scale = min(
                    1.0,
                    max_width / shape.width,
                    max_height / shape.height,
                )
                if scale < 1.0:
                    shape.width = int(shape.width * scale)
                    shape.height = int(shape.height * scale)
            except Exception:
                pass

    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    paragraph.paragraph_format.line_spacing = 1.5
    _v116_run(paragraph.add_run("Arz olunur."), 12)

    buffer = BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# Clear report and panel caches when V119 is first loaded.
if st.session_state.get("_report_engine_version") != _V119_ENGINE_VERSION:
    for key in (
        "docx_bytes",
        "note_bytes",
        "basket_docx_bytes",
        "v78_ogn_note_bytes",
        "v79_akt_note_bytes",
        "v90_ogn_docx_bytes",
        "v81_pres_note_bytes",
    ):
        st.session_state.pop(key, None)
    for key in (
        "_v117_article_cache",
        "_v117_page_cache",
        "_v119_article_cache",
        "_v119_evidence_cache",
        "_v119_compare_cache",
        "_v119_previous_event_cache",
    ):
        st.session_state.pop(key, None)
    st.session_state["_report_engine_version"] = _V119_ENGINE_VERSION

# ============================================================
# /V119
# ============================================================


# ============================================================
# V120 — QA-STABLE EVIDENCE / REPORT ENGINE
# ============================================================
# This final layer deliberately prefers correctness over aggressive enrichment.
# It prevents cross-event contamination, English search-result leakage, SEO/byline
# remnants and weak headline-only documents. The existing scan/risk/basket logic
# remains unchanged.
# ============================================================

_V120_ENGINE_VERSION = "V120-QA-STABLE-2026-09-22-A"

_V120_GENERIC_TOKENS = {
    "haber", "haberi", "son", "yeni", "ilk", "guncel", "güncel", "aciklandi",
    "açıklandı", "duyuruldu", "belirtildi", "ifade", "etti", "edildi", "olan",
    "olarak", "ile", "icin", "için", "ve", "veya", "bir", "bu", "da", "de",
    "turkiye", "türkiye", "sanayi", "teknoloji", "sirket", "şirket", "sistem",
    "sistemi", "siber", "saldiri", "saldırı", "yapay", "zeka", "zekâ", "konu",
    "baslik", "başlık", "sonrasi", "sonrası", "ilgili", "tarafindan", "tarafından",
    "gerceklesti", "gerçekleşti", "gerceklestirdi", "gerçekleştirdi",
}

_V120_ENGLISH_MARKERS = {
    "the", "and", "after", "following", "company", "shares", "revenue", "stock",
    "market", "with", "from", "for", "has", "have", "was", "were", "its", "due",
    "incident", "price", "earnings", "analysts", "expectations", "plunged", "million",
    "hours", "ago", "downgrade", "guidance", "fiscal", "year", "reported",
}

_V120_TURKISH_MARKERS = {
    "ve", "ile", "için", "olarak", "tarafından", "olduğu", "olduğunu", "açıklamıştır",
    "açıkladı", "belirtti", "duyurdu", "bildirdi", "gerçekleşti", "gerçekleştirildi",
    "ulaştı", "yüzde", "bin", "milyon", "şirket", "kurum", "türkiye", "türk",
    "sistem", "veri", "göre", "sonrası", "kapsamında", "bulunmaktadır", "edilmiştir",
}

_V120_SEO_NOISE_RE = re.compile(
    r"(?:\b\d+\s+(?:hours?|minutes?|days?)\s+ago\b\s*[·•:\-–—]*|"
    r"\b(?:forecasts?|revenue|earnings|analysts expectations|ratios)\b[^.]{0,220}\|[^.]{0,120}|"
    r"\b(?:son dakika|güncel haberler|teknoloji haberleri|haberleri)\s*[-–—|]\s*[A-Z0-9ÇĞİÖŞÜ ._-]{2,30}\s*$)",
    flags=re.IGNORECASE,
)


def _v120_clean_evidence_text(text):
    """Clean snippets/article text without inventing or translating content."""
    value = _clean_note_text(text)
    if not value:
        return ""

    value = re.sub(
        r"\b\d+\s+(?:hours?|minutes?|days?)\s+ago\b\s*[·•:\-–—]*",
        " ",
        value,
        flags=re.IGNORECASE,
    )
    value = value.replace("•", " ")
    value = re.sub(
        r"\b[A-ZÇĞİÖŞÜ]{2,}(?:\s+[A-ZÇĞİÖŞÜ]{2,}){1,4}\s*[-–—]\s*"
        r"(?=[A-ZÇĞİÖŞÜ][a-zçğıöşü])",
        " ",
        value,
    )
    value = re.sub(
        r"\b(?:Fore?casts?|Revenue|Earnings|Analysts Expectations|Ratios)\b[^.]{0,260}\|[^.]{0,180}",
        " ",
        value,
        flags=re.IGNORECASE,
    )
    value = re.sub(r"\s{2,}", " ", value).strip()
    return _v119_fix_surface(value)


def _v120_is_turkish_prose(text):
    """Reject English-dominant fallback text from Turkish institutional reports."""
    words = re.findall(r"[A-Za-zÇĞİÖŞÜçğıöşü]+", str(text or "").lower())
    if not words:
        return False
    english = sum(word in _V120_ENGLISH_MARKERS for word in words)
    turkish = sum(word in _V120_TURKISH_MARKERS for word in words)
    turkish_chars = len(re.findall(r"[çğıöşü]", str(text or "").lower()))
    if english >= 5 and english > max(2, turkish * 1.35) and turkish_chars < 3:
        return False
    return True


def _v120_signature(title):
    tokens = set(_v117_tokens(_clean_note_text(title)))
    meaningful = {
        token for token in tokens
        if len(token) >= 3 and token not in _V120_GENERIC_TOKENS
    }
    numbers = set(re.findall(r"\b\d+(?:[.,]\d+)?\b", str(title or "")))
    return meaningful, numbers


def _v120_event_overlap(title, text):
    sig, numbers = _v120_signature(title)
    tokens = set(_v117_tokens(_clean_note_text(text)))
    direct = len(sig & tokens)
    number_hits = len(numbers & set(re.findall(r"\b\d+(?:[.,]\d+)?\b", str(text or ""))))
    return direct, number_hits, _v118_title_relevance(title, text)


def _v120_sentence_is_noise(sentence):
    clean = _v120_clean_evidence_text(sentence)
    if len(clean) < 28:
        return True
    low = norm(clean)
    if any(term in low for term in _V117_NOISE_TERMS):
        return True
    if _v117_heading_like(clean) or _v117_question_or_interview(clean):
        return True
    if clean.startswith(("http://", "https://", "www.")):
        return True
    if _V120_SEO_NOISE_RE.search(clean):
        return True
    if not _v120_is_turkish_prose(clean):
        return True
    return False


def _v120_extract_facts(title, body, limit=12):
    """Extract one coherent event chain; never walk backward into another story."""
    title = _clean_note_text(title)
    body = _v120_clean_evidence_text(body)
    if not title or not body:
        return []

    fragments = _v117_fragment_body(body)
    candidates = []
    for index, raw in enumerate(fragments):
        sentence = _v120_clean_evidence_text(raw).strip(" ;")
        if _v120_sentence_is_noise(sentence):
            continue
        sentence = _v119_strip_title_echo(title, sentence)
        sentence = _v120_clean_evidence_text(_v117_formal_sentence(sentence))
        if _v120_sentence_is_noise(sentence):
            continue
        direct, number_hits, relevance = _v120_event_overlap(title, sentence)
        tokens = set(_v117_tokens(sentence))
        info = _akt_sentence_score(sentence) + _sent_score(sentence)
        if re.search(r"\b\d+(?:[.,]\d+)?\b", sentence):
            info += 2
        candidates.append(
            {
                "i": index,
                "text": sentence,
                "tokens": tokens,
                "direct": direct,
                "number_hits": number_hits,
                "relevance": relevance,
                "info": info,
            }
        )

    if not candidates:
        return []

    signature, _ = _v120_signature(title)
    anchor_pool = [
        item for item in candidates
        if item["direct"] >= 1 or item["relevance"] >= 3.2
    ]
    if not anchor_pool:
        return []

    anchor = max(
        anchor_pool[:12],
        key=lambda item: (
            item["direct"] * 8 + item["relevance"] + item["info"] * 0.25,
            -item["i"],
        ),
    )
    context = set(signature) | set(anchor["tokens"])
    accepted = []
    seen = []
    last_index = anchor["i"]
    misses = 0

    for item in candidates:
        if item["i"] < anchor["i"]:
            # Critical V120 rule: content before the event anchor cannot leak in.
            continue
        tokens = item["tokens"]
        context_overlap = len(tokens & context)
        near = item["i"] - last_index <= 3
        keep = (
            item["direct"] >= 1
            or item["relevance"] >= 3.2
            or (accepted and near and context_overlap >= 2)
            or (
                accepted
                and near
                and context_overlap >= 1
                and item["info"] >= 6
                and item["number_hits"] >= 1
            )
        )
        if not keep:
            if accepted:
                misses += 1
                if misses >= 2:
                    break
            continue
        misses = 0

        duplicate = False
        for old_tokens in seen:
            union = len(tokens | old_tokens)
            if union and len(tokens & old_tokens) / union >= 0.74:
                duplicate = True
                break
        if duplicate:
            continue

        text = item["text"]
        if accepted and norm(text).startswith(("buna göre", "bu kapsamda", "öte yandan")):
            # Connective clauses are fine after context exists.
            pass
        accepted.append(text)
        seen.append(tokens)
        context.update(tokens)
        last_index = item["i"]
        if len(accepted) >= limit:
            break

    return accepted


def _v120_row_to_tr(row):
    return _v117_record_to_tr(row)


def _v120_same_event_rows(row, max_rows=8):
    """Return only high-confidence rows belonging to the selected event."""
    selected = _v120_row_to_tr(row)
    title = _clean_note_text(selected.get("Başlık", ""))
    summary = _clean_note_text(selected.get("İçerik_Özeti", ""))
    selected_url = str(selected.get("URL") or "").strip()
    selected_sig, _ = _v120_signature(title)
    candidates = [(1000.0, selected)]

    def maybe_add(candidate, base_bonus=0.0):
        tr = _v120_row_to_tr(candidate)
        c_title = _clean_note_text(tr.get("Başlık", ""))
        if not c_title:
            return
        c_url = str(tr.get("URL") or "").strip()
        exact_url = bool(selected_url and c_url and selected_url == c_url)
        exact_title = title_key(c_title) == title_key(title)
        c_sig, _ = _v120_signature(c_title)
        sig_overlap = len(selected_sig & c_sig)
        similarity = _v104_event_similarity(
            title,
            summary,
            c_title,
            _clean_note_text(tr.get("İçerik_Özeti", "")),
        )
        if not (exact_url or exact_title or (similarity >= 0.60 and sig_overlap >= 1)):
            return
        score = base_bonus + similarity * 100 + sig_overlap * 12
        if exact_url:
            score += 250
        if exact_title:
            score += 180
        candidates.append((score, tr))

    try:
        for candidate in st.session_state.get("rows") or []:
            maybe_add(candidate, 100.0)
    except Exception:
        pass

    try:
        if _init_history_db():
            with _history_connect() as conn:
                history = conn.execute(
                    """
                    SELECT title, summary, source, url, category, risk_score, risk_status
                    FROM event_snapshots
                    ORDER BY scan_id DESC
                    LIMIT 350
                    """
                ).fetchall()
            for ht, hs, hsrc, hu, hcat, hrisk, hstatus in history:
                maybe_add(
                    {
                        "title": ht,
                        "summary": hs,
                        "source": hsrc,
                        "url": hu,
                        "category": hcat,
                        "risk_score": hrisk,
                        "risk_status": hstatus,
                    },
                    40.0,
                )
    except Exception:
        pass

    candidates.sort(key=lambda pair: pair[0], reverse=True)
    unique = []
    seen = set()
    for _, item in candidates:
        key = (
            title_key(item.get("Başlık", "")),
            str(item.get("URL") or "").split("#")[0],
        )
        if key in seen:
            continue
        seen.add(key)
        unique.append(item)
        if len(unique) >= max_rows:
            break
    return unique


def _v120_resolve_direct_url(tr):
    for value in (tr.get("Yayıncı_URL"), tr.get("URL")):
        url = str(value or "").strip()
        if not url:
            continue
        if _v116_valid_direct_url(url):
            return url
        if "news.google.com" in url:
            try:
                decoded = _v116_decode_google_news_url(url)
            except Exception:
                decoded = ""
            if decoded and _v116_valid_direct_url(decoded):
                return decoded
    return str(tr.get("URL") or "").strip()


def _v120_candidate_from_row(selected_title, tr, selected=False):
    """Build one evidence candidate while keeping its provenance attached."""
    direct_url = _v120_resolve_direct_url(tr)
    source = _v117_source_short(tr.get("Kaynak", ""), direct_url)
    row_title = _v116_clean_headline(tr.get("Başlık", ""), source)
    candidate_title = row_title or selected_title
    texts = []
    summary = _v120_clean_evidence_text(tr.get("İçerik_Özeti", ""))
    if summary:
        texts.append((summary, 150))

    page = {"text": "", "images": [], "source": "", "title": ""}
    if _v116_valid_direct_url(direct_url):
        try:
            page = _v117_fetch_clean_page(direct_url, selected_title) or page
        except Exception:
            page = {"text": "", "images": [], "source": "", "title": ""}
        page_title = _clean_note_text(page.get("title", ""))
        if page_title:
            _, _, page_rel = _v120_event_overlap(selected_title, page_title)
        else:
            page_rel = 9.0
        if page_rel >= 2.4:
            body = _v120_clean_evidence_text(page.get("text", ""))
            if body:
                texts.append((body, 700))

    best = None
    for text, bonus in texts:
        if not _v120_is_turkish_prose(text):
            continue
        facts = _v120_extract_facts(selected_title, text, limit=12)
        if not facts:
            continue
        score = len(facts) * 1500 + min(len(text), 7000) + bonus
        if selected:
            score += 3000
        try:
            score += min(source_rank(direct_url or source), 500)
        except Exception:
            pass
        item = {
            "score": score,
            "facts": facts,
            "text": text,
            "title": candidate_title,
            "source": _v117_source_short(page.get("source") or source, direct_url),
            "url": direct_url,
            "images": list(page.get("images") or []),
            "selected": bool(selected),
        }
        if best is None or item["score"] > best["score"]:
            best = item
    return best


def _v120_site_snippet_candidate(title, preferred_url):
    """Safe last resort: site-restricted Turkish snippets only; never generic full pages."""
    try:
        snippets = _v117_search_snippets(title, preferred_url)[:4]
    except Exception:
        snippets = []
    text = " ".join(_v120_clean_evidence_text(item) for item in snippets if item)
    if not text or not _v120_is_turkish_prose(text):
        return None
    facts = _v120_extract_facts(title, text, limit=8)
    if len(facts) < 2:
        return None
    return {
        "score": len(facts) * 1000 + 120,
        "facts": facts,
        "text": text,
        "title": title,
        "source": "",
        "url": preferred_url,
        "images": [],
        "selected": False,
    }


def _v120_event_evidence(row, purpose="note"):
    """Create a provenance-safe evidence bundle for one selected event."""
    tr = _v120_row_to_tr(row)
    selected_title = _v116_clean_headline(
        tr.get("Başlık", ""), tr.get("Kaynak", "")
    )
    if not selected_title:
        raise ReportQualityError("Seçilen kaydın haber başlığı bulunamadı.")

    cache_seed = (
        selected_title
        + "|"
        + str(tr.get("URL", ""))
        + "|"
        + purpose
        + "|v120"
    )
    cache_key = hashlib.sha1(cache_seed.encode("utf-8", "ignore")).hexdigest()
    cache = st.session_state.setdefault("_v120_event_evidence_cache", {})
    if cache_key in cache:
        return dict(cache[cache_key])

    rows = _v120_same_event_rows(tr)
    candidates = []
    for index, candidate_row in enumerate(rows):
        candidate = _v120_candidate_from_row(
            selected_title,
            candidate_row,
            selected=(index == 0),
        )
        if candidate:
            candidates.append(candidate)

    selected_url = _v120_resolve_direct_url(tr)
    best_fact_count = max((len(item["facts"]) for item in candidates), default=0)
    if best_fact_count < 4:
        snippet_candidate = _v120_site_snippet_candidate(selected_title, selected_url)
        if snippet_candidate:
            candidates.append(snippet_candidate)

    if not candidates:
        raise ReportQualityError(
            f'"{selected_title}" için Türkçe ve olayla uyumlu ayrıntılı içerik elde edilemedi.'
        )

    candidates.sort(key=lambda item: item["score"], reverse=True)
    best = candidates[0]

    # Keep the selected source when it already clears the quality threshold.
    selected_candidates = [item for item in candidates if item.get("selected")]
    if selected_candidates:
        selected_best = max(selected_candidates, key=lambda item: item["score"])
        minimum = 3 if purpose in {"akt", "ogn"} else 5
        if len(selected_best["facts"]) >= minimum:
            best = selected_best

    merged_facts = []
    seen_tokens = []
    source_pool = candidates[:4] if purpose in {"note", "ogn"} else [best]
    for candidate in source_pool:
        for fact in candidate["facts"]:
            direct, _, relevance = _v120_event_overlap(selected_title, fact)
            if direct < 1 and relevance < 3.0:
                # Secondary-source evidence must independently anchor to the event.
                continue
            tokens = set(_v117_tokens(fact))
            duplicate = False
            for old in seen_tokens:
                union = len(tokens | old)
                if union and len(tokens & old) / union >= 0.72:
                    duplicate = True
                    break
            if duplicate:
                continue
            merged_facts.append(_v120_clean_evidence_text(fact))
            seen_tokens.append(tokens)
            if len(merged_facts) >= 12:
                break
        if len(merged_facts) >= 12:
            break

    required = {"akt": 3, "ogn": 2, "note": 5}.get(purpose, 3)
    if len(merged_facts) < required:
        raise ReportQualityError(
            f'"{selected_title}" için {required} bağımsız ve aynı olaya bağlı olgu elde edilemedi.'
        )

    bundle = {
        "title": selected_title,
        "facts": merged_facts,
        "best": best,
        "source": best.get("source") or _v117_source_short(tr.get("Kaynak", ""), selected_url),
        "url": best.get("url") or selected_url,
        "images": list(best.get("images") or []),
        "fact_count": len(merged_facts),
    }
    cache[cache_key] = dict(bundle)
    return bundle


def article_detail(row):
    """V120 compatibility wrapper: return only provenance-safe evidence."""
    bundle = _v120_event_evidence(row, purpose="note")
    return {
        "title": bundle["title"],
        "source": bundle["source"],
        "canonical": bundle["url"],
        "text": " ".join(bundle["facts"]),
        "images": bundle["images"],
        "quality_facts": bundle["fact_count"],
        "quality_ready": bundle["fact_count"] >= 5,
    }


def _v120_title_as_context(title):
    sentence = _v120_clean_evidence_text(title).strip(" .;:-")
    if not sentence:
        return ""
    replacements = (
        (r"\bdüştü$", "düşmüştür"),
        (r"\barttı$", "artmıştır"),
        (r"\bazaldı$", "azalmıştır"),
        (r"\baçıklandı$", "açıklanmıştır"),
        (r"\bduyuruldu$", "duyurulmuştur"),
        (r"\bgerçekleşti$", "gerçekleşmiştir"),
        (r"\btest edildi$", "test edilmiştir"),
        (r"\bdüzenlenecek$", "düzenlenecektir"),
    )
    for pattern, replacement in replacements:
        sentence = re.sub(pattern, replacement, sentence, flags=re.IGNORECASE)
    if sentence[-1:] not in ".!?":
        sentence += "."
    return _v119_fix_surface(sentence)


def _v120_note_paragraphs(title, facts):
    """Build a stable, source-grounded 4–5 paragraph information note."""
    facts = [_v120_clean_evidence_text(item) for item in facts if item]
    facts = [item for item in facts if item and _v120_is_turkish_prose(item)]
    if len(facts) < 5:
        raise ReportQualityError(
            f'"{title}" için standart bilgi notu oluşturacak ayrıntı düzeyine ulaşılamadı.'
        )

    first = facts[0]
    if norm(first).startswith(("buna göre", "bu kapsamda", "öte yandan", "ayrıca")):
        context = _v120_title_as_context(title)
        if context and title_key(context) != title_key(first):
            facts.insert(0, context)

    count = len(facts)
    if count >= 10:
        groups = [facts[:2], facts[2:5], facts[5:7], facts[7:9], facts[9:]]
    elif count >= 8:
        groups = [facts[:2], facts[2:4], facts[4:6], facts[6:]]
    elif count >= 6:
        groups = [facts[:2], facts[2:4], facts[4:]]
    else:
        groups = [facts[:1], facts[1:3], facts[3:]]

    paragraphs = []
    for group in groups:
        text = " ".join(group).strip()
        if text:
            paragraphs.append(_v119_fix_surface(text))
    if len(paragraphs) < 3:
        raise ReportQualityError(
            f'"{title}" için giriş-gelişme-sonuç bütünlüğü kurulamadı.'
        )
    return paragraphs[:5]


def make_analyst_docx(df, title="BİLGİ NOTU"):
    """V120 stable information note: same quality standard for every entry point."""
    frame = df.copy() if df is not None else pd.DataFrame()
    rows = [] if frame.empty else _v115_dedupe_rows(frame.to_dict("records"))
    if not rows:
        raise ReportQualityError("Bilgi notu oluşturmak için haber bulunamadı.")

    paragraphs = []
    for row in rows:
        bundle = _v120_event_evidence(row, purpose="note")
        paragraphs.extend(_v120_note_paragraphs(bundle["title"], bundle["facts"]))

    if len(paragraphs) < 3:
        raise ReportQualityError("Bilgi notu kalite eşiğinin altında kaldı.")

    document = _v116_doc_defaults(Document(), 2.5, 2.5, 2.5, 2.5)
    section = document.sections[0]
    section.header_distance = Cm(1.25)
    section.footer_distance = Cm(1.25)
    for text in paragraphs[:5]:
        paragraph = document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(6)
        # Reference information note uses normal single spacing.
        paragraph.paragraph_format.line_spacing = 1.0
        _v116_run(paragraph.add_run(_v119_fix_surface(text)), 12)

    buffer = BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def _v120_ogn_item(row):
    bundle = _v120_event_evidence(row, purpose="ogn")
    facts = list(bundle["facts"])
    chosen = []
    total_chars = 0
    for fact in facts:
        clean = _v120_clean_evidence_text(fact)
        if not clean:
            continue
        if chosen and (len(chosen) >= 3 or total_chars + len(clean) > 760):
            break
        chosen.append(clean)
        total_chars += len(clean) + 1
    if len(chosen) < 2:
        raise ReportQualityError(
            f'ÖGN için "{bundle["title"]}" maddesi yeterli ayrıntıya ulaşamadı.'
        )
    if norm(chosen[0]).startswith(("buna göre", "bu kapsamda", "öte yandan", "ayrıca")):
        context = _v120_title_as_context(bundle["title"])
        if context and title_key(context) != title_key(chosen[0]):
            chosen.insert(0, context)
            chosen = chosen[:3]
    return _v119_fix_surface(" ".join(chosen))


def make_important_basket_docx_v101(basket_df):
    """V120 ÖGN: institutional format + strict event isolation."""
    rows = _v115_dedupe_rows(
        [] if basket_df is None else basket_df.to_dict("records")
    )
    if not rows:
        raise ReportQualityError("Önemli Gelişmeler Sepeti boş.")

    # Deliberately sequential: report generation may touch Streamlit session caches,
    # which are not guaranteed to be thread-safe. Deterministic quality is more
    # important here than shaving a few seconds from a user-triggered Word export.
    outputs = [None] * len(rows)
    errors = []
    for index, row in enumerate(rows):
        try:
            outputs[index] = _v120_ogn_item(row)
        except Exception as exc:
            title = _clean_note_text(row.get("title", row.get("Başlık", "")))
            errors.append(f"{title}: {exc}")

    if errors:
        raise ReportQualityError(
            "ÖGN oluşturulmadı; aşağıdaki maddeler kalite eşiğini geçemedi: "
            + " | ".join(errors[:4])
        )

    document = _v116_doc_defaults(Document(), 2.25, 1.5, 1.905, 1.905)
    section = document.sections[0]
    section.header_distance = Cm(1.25)
    section.footer_distance = Cm(1.0)
    _v116_add_ogn_footer(section)

    today = date.today()
    paragraph = document.add_paragraph(style="No Spacing")
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    paragraph.paragraph_format.line_spacing = 1.15
    _v116_run(
        paragraph.add_run(
            f'{(today - timedelta(days=1)).strftime("%d/%m/%Y")} – {today.strftime("%d/%m/%Y")}'
        ),
        12,
    )

    paragraph = document.add_paragraph(style="No Spacing")
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    paragraph.paragraph_format.line_spacing = 1.15
    paragraph.paragraph_format.space_before = Pt(6)
    paragraph.paragraph_format.space_after = Pt(6)
    _v116_run(paragraph.add_run("Konu: "), 12, True)
    _v116_run(paragraph.add_run("STB Temsilciliği Önemli Gelişmeler Notu"), 12)

    for text in outputs:
        paragraph = document.add_paragraph(style="No Spacing")
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        paragraph.paragraph_format.line_spacing = 1.15
        paragraph.paragraph_format.space_before = Pt(6)
        paragraph.paragraph_format.space_after = Pt(6)
        clean = _v119_fix_surface(text).rstrip(".")
        _v116_run(paragraph.add_run(clean + " (STB)."), 12)

    paragraph = document.add_paragraph(style="No Spacing")
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    paragraph.paragraph_format.line_spacing = 1.15
    paragraph.paragraph_format.space_before = Pt(6)
    paragraph.paragraph_format.space_after = Pt(6)
    _v116_run(paragraph.add_run("Arz olunur."), 12)

    buffer = BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def _v120_akt_clause_text(facts):
    clauses = []
    for fact in facts[:6]:
        clean = _v120_clean_evidence_text(fact).strip().rstrip(" .;:")
        if clean:
            clauses.append(clean)
    if len(clauses) < 3:
        return ""
    return "; ".join(clauses)


def _v120_prepare_akt_row(row):
    bundle = _v120_event_evidence(row, purpose="akt")
    best = bundle["best"]
    facts = _v120_extract_facts(bundle["title"], best.get("text", ""), limit=8)
    if len(facts) < 3:
        # Use only facts already tied to the same chosen source; do not mix generic web pages.
        facts = list(bundle["facts"][:6])
    summary = _v120_akt_clause_text(facts)
    if not summary:
        raise ReportQualityError(
            f'AKT için "{bundle["title"]}" haberinde yeterli ayrıntı elde edilemedi.'
        )

    image = None
    for candidate in list(best.get("images") or [])[:8]:
        try:
            image = _v117_valid_report_image(candidate)
        except Exception:
            image = None
        if image:
            break

    return {
        "title": bundle["title"],
        "source": _v117_source_short(bundle["source"], bundle["url"]),
        "url": bundle["url"],
        "summary": summary,
        "image": image,
    }


def make_docx(rows):
    """V120 AKT: coherent Turkish content, same-event source, matching image provenance."""
    source_rows = []
    for row in rows or []:
        source_rows.append(_v120_row_to_tr(row))
    source_rows = _v115_dedupe_rows(source_rows)
    if not source_rows:
        raise ReportQualityError("AKT raporu için haber bulunamadı.")

    prepared = [None] * len(source_rows)
    errors = []
    for index, row in enumerate(source_rows):
        try:
            prepared[index] = _v120_prepare_akt_row(row)
        except Exception as exc:
            title = _clean_note_text(row.get("Başlık", ""))
            errors.append(f"{title}: {exc}")

    if errors:
        raise ReportQualityError(
            "AKT raporu oluşturulmadı; kalite eşiğini geçemeyen maddeler: "
            + " | ".join(errors[:4])
        )

    document = _v116_doc_defaults(
        Document(), top=2.5, bottom=1.25, left=2.5, right=2.5
    )
    section = document.sections[0]
    section.header_distance = Cm(1.25)
    section.footer_distance = Cm(1.25)

    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_after = Pt(0)
    _v116_run(paragraph.add_run("AÇIK KAYNAK TARAMA ÇALIŞMASI"), 16, True)
    _v116_add_akt_info_table(document)

    intro = document.add_paragraph()
    intro.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    intro.paragraph_format.first_line_indent = Cm(0.63)
    intro.paragraph_format.line_spacing = 1.5
    intro.paragraph_format.space_before = Pt(0)
    intro.paragraph_format.space_after = Pt(0)
    intro_text = _v116_akt_intro_text(source_rows)
    topics = _v116_topic_labels(source_rows)
    cursor = 0
    position = 0
    spans = []
    for topic in topics:
        quoted = f"“{topic}”"
        index = intro_text.find(quoted, position)
        if index >= 0:
            spans.append((index, index + len(quoted)))
            position = index + len(quoted)
    for start, end in spans:
        if start > cursor:
            _v116_run(intro.add_run(intro_text[cursor:start]), 12)
        _v116_run(intro.add_run(intro_text[start:end]), 12, italic=True)
        cursor = end
    if cursor < len(intro_text):
        _v116_run(intro.add_run(intro_text[cursor:]), 12)

    for number, item in enumerate(prepared, start=1):
        paragraph = document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        paragraph.paragraph_format.first_line_indent = Cm(0.63)
        paragraph.paragraph_format.line_spacing = 1.5
        paragraph.paragraph_format.space_before = Pt(4)
        paragraph.paragraph_format.space_after = Pt(6)
        _v116_run(paragraph.add_run(f'{number}. “{item["source"]}”'), 12, True)
        _v116_run(paragraph.add_run(" isimli internet sitesinde, "), 12)
        _v116_run(paragraph.add_run(f'“{item["title"]}”'), 12, True, True)
        _v116_run(paragraph.add_run(" başlığıyla bir haber yayımlanmıştır. ("), 12)
        _word_hyperlink(paragraph, item["url"], item["url"] or "Haber bağlantısı")
        _v116_run(paragraph.add_run(") Söz konusu haber içeriğinde, "), 12)
        _v116_run(paragraph.add_run(item["summary"].rstrip(" .;")), 12)
        _v116_run(paragraph.add_run(" hususları ifade edilmiştir."), 12)

        if item.get("image"):
            caption = document.add_paragraph()
            caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
            caption.paragraph_format.first_line_indent = Cm(0.63)
            caption.paragraph_format.line_spacing = 1.5
            caption.paragraph_format.space_before = Pt(4)
            caption.paragraph_format.space_after = Pt(4)
            _v116_run(
                caption.add_run(
                    f'Görsel {number}: “{item["source"]}” Sitesinde Yer Alan Görsel'
                ),
                12,
                True,
            )
            image_paragraph = document.add_paragraph()
            image_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            image_paragraph.paragraph_format.space_after = Pt(8)
            try:
                run = image_paragraph.add_run()
                shape = run.add_picture(item["image"])
                max_width = Cm(15.3)
                max_height = Cm(12.7)
                scale = min(1.0, max_width / shape.width, max_height / shape.height)
                if scale < 1.0:
                    shape.width = int(shape.width * scale)
                    shape.height = int(shape.height * scale)
            except Exception:
                pass

    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    paragraph.paragraph_format.line_spacing = 1.5
    _v116_run(paragraph.add_run("Arz olunur."), 12)

    buffer = BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def _v120_fast_shift_counts(df, current_scan_id=None):
    """Cheap exact-key shift counters; no semantic all-pairs comparison on first render."""
    previous_id = _previous_scan_id(current_scan_id)
    previous = _v119_previous_events(previous_id)
    if previous.empty or df is None or df.empty:
        return 0, 0, 0

    previous_by_url = {}
    previous_by_title = {}
    for _, row in previous.iterrows():
        url = str(row.get("url") or "").strip()
        key = title_key(row.get("title", ""))
        if url:
            previous_by_url[url] = row
        if key:
            previous_by_title[key] = row

    current = _v119_current_events(df)
    new_events = risk_up = verify_up = 0
    for _, row in current.iterrows():
        url = str(row.get("url") or "").strip()
        key = title_key(row.get("title", ""))
        previous_row = previous_by_url.get(url) if url else None
        if previous_row is None and key:
            previous_row = previous_by_title.get(key)
        if previous_row is None:
            new_events += 1
            continue
        if int(row.get("risk_score") or 0) >= int(previous_row.get("risk_score") or 0) + 15:
            risk_up += 1
        if _verification_rank(row.get("verification", "")) > _verification_rank(previous_row.get("verification", "")):
            verify_up += 1
    return new_events, risk_up, verify_up


def _shift_start_summary(df, current_scan_id=None):
    """V120 immediate shift summary; removes the visible mid-page semantic-compare stall."""
    if df is None or df.empty:
        return {}, pd.DataFrame(), ""

    key = _v113_scan_key(df, current_scan_id) + ":" + _v113_shift_mark_key() + ":v120"
    cached = _v113_get_cached("shift_summary_v120", key)
    if cached is not None:
        stats, top, label = cached
        return dict(stats), top.copy(), label

    baseline, baseline_label, _ = _shift_baseline(current_scan_id)
    frame = df.copy()
    frame["Tarih_dt"] = pd.to_datetime(frame.get("Tarih_dt"), utc=True, errors="coerce")
    if baseline is not None:
        since = frame[frame["Tarih_dt"].isna() | (frame["Tarih_dt"] >= baseline)].copy()
    else:
        since = frame.copy()

    new_events, risk_up, verify_up = _v120_fast_shift_counts(df, current_scan_id)
    high_risk = int((since.get("Risk_Durumu", pd.Series("", index=since.index)) == "Yüksek Risk").sum())
    titles = since.get("Başlık", pd.Series("", index=since.index)).fillna("")
    summaries = since.get("İçerik_Özeti", pd.Series("", index=since.index)).fillna("")
    osb_count = sum(
        bool(is_osb_fire(title, summary))
        for title, summary in zip(titles.astype(str), summaries.astype(str))
    )
    top = _v115_fast_shift_top(since, 8)
    stats = {
        "new_news": len(since),
        "new_important_events": int(new_events),
        "high_risk": high_risk,
        "risk_up": int(risk_up),
        "verify_up": int(verify_up),
        "osb": int(osb_count),
        "baseline_label": baseline_label,
    }
    _v113_set_cached("shift_summary_v120", key, (dict(stats), top.copy(), baseline_label))
    return stats, top, baseline_label


# V120 cache/version reset.
if st.session_state.get("_report_engine_version") != _V120_ENGINE_VERSION:
    for key in (
        "docx_bytes",
        "note_bytes",
        "basket_docx_bytes",
        "v78_ogn_note_bytes",
        "v79_akt_note_bytes",
        "v90_ogn_docx_bytes",
        "v81_pres_note_bytes",
    ):
        st.session_state.pop(key, None)
    for key in (
        "_v120_event_evidence_cache",
        "_v117_page_cache",
        "_v119_compare_cache",
        "_v119_previous_event_cache",
    ):
        st.session_state.pop(key, None)
    st.session_state["_report_engine_version"] = _V120_ENGINE_VERSION

# ============================================================
# /V120
# ============================================================


# -----------------------------
# UI
# -----------------------------
st.title('🛡️ T.C. Sanayi ve Teknoloji Bakanlığı Açık Kaynak Tarama Merkezi')
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

# V60: Yeni browser oturumunda önceki giriş zamanı otomatik belirlenir.
_v60_previous_visit=_v60_register_visit_once()
if '_v60_catchup_done' not in st.session_state:
    st.session_state['_v60_catchup_done']=False
if '_v60_catchup_rows' not in st.session_state:
    st.session_state['_v60_catchup_rows']=[]
if '_v60_catchup_hours' not in st.session_state:
    st.session_state['_v60_catchup_hours']=None

if not st.session_state['_v60_catchup_done']:
    # V114: Uygulama açılışında otomatik web/RSS taraması yapılmaz.
    # Önceki giriş zamanı yalnız baseline olarak tutulur; "Şu An Bilmen Gerekenler"
    # kullanıcının başlattığı ana tarama tamamlandığında o taramanın verisinden üretilir.
    st.session_state['_v60_catchup_done']=True
    st.session_state['_v60_catchup_rows']=[]
    st.session_state['_v60_catchup_hours']=None


if run:
    import time as _v119_time
    _v119_scan_started=_v119_time.perf_counter()
    cutoff=(datetime.now(timezone.utc)-timedelta(hours=hours)).astimezone(timezone.utc)
    when=period_window(hours)
    batches=[('🇹🇷 Türk medya / sanayi-teknoloji',build_turkish_queries(when,query),'turkish')]
    # V41 bağımsız katmanları: yalnızca 4 ek sorgu; mevcut paralel havuzda çalışır.
    batches.append(('🏛️ Resmî kaynak radarı',build_official_radar_queries(when),'official'))
    batches.append(('📊 Önemli istatistik radarı',build_statistics_queries(when),'statistics'))
    if neg: batches.append(('⚠️ Negatif haber taraması',build_negative_queries(when),'negative'))
    if greek: batches.append(('🇬🇷 Yunan medyası / Türk savunma',build_greek_queries(when),'greek'))
    if social: batches.append(('📱 Açık sosyal / indeks',build_social_queries(when),'social'))
    if global_on: batches.append(('🌍 Global basın',[
        f'(Turkey OR Türkiye) (industry OR manufacturing OR technology OR semiconductor OR defense OR aerospace OR automotive) timespan:{when}',
        f'(Turkey OR Turkish) (Baykar OR ASELSAN OR TUSAŞ OR ROKETSAN OR HAVELSAN OR KAAN OR drone OR missile) timespan:{when}'
    ],'global'))
    all_rows=[]; stat={'Ham sonuç':0,'Zaman dışı':0,'Konu dışı':0,'Yunan dışı':0,'Kaynak dışı':0,'Sonuç':0,'Olay':0}
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
        critical_label=critical_industrial_incident(row.get('Başlık',''),row.get('İçerik_Özeti',''))
        is_high=row.get('Risk_Durumu')=='Yüksek Risk' or risk_score>=70 or bool(critical_label)
        live_alerts.insert(0,{
            'Tarih':str(row.get('Tarih','')),
            'Seviye':critical_label if critical_label else ('YÜKSEK RİSK' if is_high else 'NEGATİF'),
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

    # V114 — Bütün RSS sorguları tek paralel havuzda çalışır.
    # Önceki sürümde Türk ana taraması tamamen bittikten sonra tamamlayıcı kaynaklar
    # başlıyordu. Ana ekran sonuçları zaten tarama sonunda çizildiği için bu sıralı
    # bekleme kaldırıldı; kapsam ve normalize kuralları değişmeden korunur.
    jobs=[]
    mode_order=[]
    for label,queries,mode in batches:
        mode_order.append(mode)
        for q in queries:
            jobs.append((label,q,mode))

    raw_by_mode={}
    if jobs:
        workers=min(12,len(jobs))
        status_box.write(f'⚡ Paralel tarama — {len(jobs)} sorgu / {workers} eşzamanlı')
        with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as ex:
            future_map={ex.submit(rss,q):(label,mode) for label,q,mode in jobs}
            for fut in concurrent.futures.as_completed(future_map):
                label,mode=future_map[fut]
                try:
                    chunk=fut.result() or []
                except Exception:
                    chunk=[]
                stat['Ham sonuç']+=len(chunk)
                raw_by_mode.setdefault(mode,[]).extend(chunk)

                # ÖZEL KRİTİK SANAYİ OLAYI ALARMI:
                # OSB/OSB dışı fabrika-tesis yangın ve patlamalarında sorgu döner dönmez bildir.
                if instant_alerts and chunk:
                    quick_rows,_quick_reasons=normalize_rows(chunk,cutoff,mode,query)
                    for qr in quick_rows:
                        critical_label=critical_industrial_incident(qr.get('Başlık',''),qr.get('İçerik_Özeti',''))
                        if critical_label:
                            qkey=_alert_key(qr)
                            if qkey not in alerted_keys:
                                _register_alert(qr)
                                icon='💥' if 'PATLAMA' in critical_label else '🔥'
                                st.toast(
                                    f'{critical_label}: {str(qr.get("Başlık",""))[:105]}',
                                    icon=icon
                                )
                                live_alarm_box.error(
                                    f'{icon} **KRİTİK SANAYİ OLAYI ALARMI — {critical_label}** — '
                                    f'{qr.get("Tarih","")} · {qr.get("Kaynak","Açık Kaynak")} · '
                                    f'{str(qr.get("Başlık",""))[:140]}'
                                )

    # Mode bazlı normalize + birleştirme.
    for mode in mode_order:
        raw=raw_by_mode.get(mode,[])
        incoming=_merge_batch(raw,mode)
        old_keys={_alert_key(x) for x in all_rows}
        all_rows=dedupe(all_rows+incoming)
        # Genel negatif/yüksek risk alarmı bu aşamada verilmez.
        # Önce aşağıda gerçek haber sayfasının tam metni okunarak nihai sınıflandırma yapılır.
        # Kritik sanayi yangın/patlama anlık alarmı yukarıdaki özel blokta aynen devam eder.
        stat['Sonuç']=len(all_rows)

    # 3) Analitik katman — V44 performans düzenlemesi.
    # V43'teki her haber sayfasını tek tek indiren tam-metin negatif analizi kaldırıldı.
    # V42'deki hızlı ve bağlam duyarlı Başlık + RSS İçerik/Özet sınıflandırması kullanılır.
    if all_rows:
        status_box.write('🧩 Hızlı olay analizi hazırlanıyor...')
        all_rows=enrich_rows(all_rows)
        stat['Olay']=len({r.get('Olay_ID') for r in all_rows})
    else:
        stat['Olay']=0

    # Nihai alarm listesi mevcut hızlı sınıflandırmadan oluşturulur.
    # Kritik sanayi yangın/patlama alarmı aynen korunur.
    live_alerts=[]
    alerted_keys=set()
    final_toast_count=0
    for ar in all_rows:
        critical_label=critical_industrial_incident(ar.get('Başlık',''),ar.get('İçerik_Özeti',''))
        is_negative=(ar.get('Duygu')=='Negatif')
        is_high=(ar.get('Risk_Durumu')=='Yüksek Risk' or int(ar.get('Risk_Skoru',0) or 0)>=70)
        if critical_label or is_negative or is_high:
            if _register_alert(ar):
                if instant_alerts and not critical_label and final_toast_count < MAX_TOASTS_PER_SCAN:
                    st.toast(
                        f'{"🚨 YÜKSEK RİSK" if is_high else "⚠️ NEGATİF"}: {str(ar.get("Başlık",""))[:100]}',
                        icon='🚨' if is_high else '⚠️'
                    )
                    final_toast_count+=1

    if live_alerts:
        live_alarm_box.warning(
            f'🔔 {len(live_alerts)} negatif/riskli içerik yakalandı. Son: {live_alerts[0]["Başlık"][:100]}'
        )

    status_box.update(
        label='🧩 Tarama tamamlandı; panel özetleri hazırlanıyor...',
        state='running'
    )
    # V101 — Tarama sonucu daha session_state'e yazılmadan gerçek yayın tarih-saatine göre sıralanır.
    # Böylece özellikle Son 24 Saat taramasında Kronolojik ekran ilk açılışta en yeni -> en eski gelir.
    def _v101_row_dt(_r):
        _d=_to_utc_datetime(_r.get('Tarih_dt'))
        if _d is None:
            _d=_to_utc_datetime(_r.get('Tarih'))
        return _d or datetime.min.replace(tzinfo=timezone.utc)

    all_rows=sorted(all_rows,key=_v101_row_dt,reverse=True)
    st.session_state.rows=all_rows
    st.session_state.scan_time=datetime.now().astimezone()
    st.session_state.stats=stat
    st.session_state.last_scan_alerts=live_alerts

    # V114 — "Şu An Bilmen Gerekenler" için ikinci bir web taraması yapma.
    # Önceki girişten sonraki kayıtları, az önce tamamlanan ve zaten zenginleştirilmiş
    # ana tarama sonuçlarından filtrele.
    _prev_catch=st.session_state.get('_v60_previous_visit')
    if _prev_catch is not None and not pd.isna(_prev_catch):
        try:
            _prev_utc=pd.to_datetime(_prev_catch,utc=True).to_pydatetime()
            _now_utc=datetime.now(timezone.utc)
            st.session_state['_v60_catchup_hours']=max(0.0,(_now_utc-_prev_utc).total_seconds()/3600)
            _local_catch=[]
            for _r in all_rows:
                _rdt=_to_utc_datetime(_r.get('Tarih_dt'))
                if _rdt is None:
                    _rdt=_to_utc_datetime(_r.get('Tarih'))
                if _rdt is not None and _rdt>=_prev_utc:
                    _local_catch.append(_r)
            st.session_state['_v60_catchup_rows']=_local_catch
        except Exception:
            st.session_state['_v60_catchup_rows']=[]
            st.session_state['_v60_catchup_hours']=None
    else:
        st.session_state['_v60_catchup_rows']=[]
        st.session_state['_v60_catchup_hours']=None

    # V33 geçmiş karşılaştırma katmanı: tarama bittikten SONRA olay özetini kaydeder.
    # Tarama motoruna veya sıralamaya müdahale etmez.
    st.session_state.current_scan_id=_save_scan_history(
        all_rows,
        st.session_state.scan_time,
        hours
    )

    # V113 — Yeni tarama: yalnız türetilmiş panel cache'ini sıfırla.
    # Tarama sonucu ve geçmiş verisi korunur.
    st.session_state['_v113_panel_cache']={}
    st.session_state['_v113_panel_cache_active_key']=None
    st.session_state['_v119_compare_cache']={}
    st.session_state['_v119_previous_event_cache']={}

    # V119: expensive derived data is computed while the scan status is still
    # visible, before the page starts rendering section by section. The Vardiya
    # section therefore reads from cache instead of appearing to freeze midway.
    _v119_panel_started=_v119_time.perf_counter()
    try:
        _v119_scan_df=pd.DataFrame(all_rows)
        if not _v119_scan_df.empty:
            _v119_scan_df['Tarih_dt']=pd.to_datetime(
                _v119_scan_df.get('Tarih_dt'),utc=True,errors='coerce'
            )
            # V120: Vardiya özeti ilk görünümde yalnız hızlı exact-key sayımlarıyla
            # hazırlanır. Ayrıntılı semantik karşılaştırma sayfa ortasında/ön hesaplamada
            # çalıştırılmaz; böylece Vardiya Başlangıç Özeti görünürken donma oluşmaz.
            _shift_start_summary(
                _v119_scan_df,st.session_state.get('current_scan_id')
            )
    except Exception:
        pass

    stat['Panel ön hesaplama ms']=int(
        (_v119_time.perf_counter()-_v119_panel_started)*1000
    )
    stat['Toplam tarama sn']=round(
        _v119_time.perf_counter()-_v119_scan_started,2
    )
    st.session_state.stats=stat
    status_box.update(
        label=(
            f'✅ Tarama tamamlandı — {len(all_rows)} haber / {stat["Olay"]} olay '
            f'• panel {stat["Panel ön hesaplama ms"]} ms'
        ),
        state='complete'
    )


# V114 — ŞU AN BİLMEN GEREKENLER: ek ağ isteği yok; son ana tarama verisinden hazırlanır.
st.subheader('⚡ Şu An Bilmen Gerekenler')
_prev=st.session_state.get('_v60_previous_visit')
_catch_rows=st.session_state.get('_v60_catchup_rows') or []
_catch_hours=st.session_state.get('_v60_catchup_hours')

if _prev is None or pd.isna(_prev):
    st.info('İlk giriş kaydı oluşturuldu. Sonraki oturumlarda bu alan, başlattığınız ana tarama verisi üzerinden son girişten sonraki gelişmeleri gösterecektir.')
else:
    try:
        _prev_local=pd.to_datetime(_prev,utc=True).tz_convert(datetime.now().astimezone().tzinfo)
        st.caption(f'Son giriş: {_prev_local.strftime("%d.%m.%Y %H:%M")} — bu tarihten sonraki gelişmeler son ana tarama verisi üzerinden kontrol edilmektedir.')
    except Exception:
        pass

    _has_main_scan=st.session_state.get('rows') is not None
    _now5=_v60_now_to_know_table(_catch_rows,5) if _has_main_scan else pd.DataFrame()
    if not _has_main_scan:
        st.info('Bu alan ek tarama yapmaz. Ana taramayı başlattığınızda son girişten sonraki gelişmeler mevcut tarama sonuçlarından otomatik süzülecektir.')
    elif _now5.empty:
        st.success('Son girişinizden bu yana, seçtiğiniz ana tarama aralığında öncelikli yeni bir gelişme tespit edilmedi.')
    else:
        st.warning(f'Son girişinizden bu yana dikkat gerektiren {_now5.shape[0]} gelişme öne çıkıyor.')

        # V61: Bu bölümden doğrudan seçim/sepet/bilgi notu işlemleri yapılabilir.
        _catch_df=pd.DataFrame(_catch_rows)
        _know_rows=[]
        for _,_v in _now5.iterrows():
            _url=str(_v.get('URL','') or '')
            _title=norm(_v.get('Gelişme',''))
            _match=pd.DataFrame()
            if not _catch_df.empty and _url and 'URL' in _catch_df.columns:
                _match=_catch_df[_catch_df['URL'].astype(str)==_url]
            if _match.empty and not _catch_df.empty and 'Başlık' in _catch_df.columns:
                _match=_catch_df[_catch_df['Başlık'].astype(str).map(norm)==_title]
            if not _match.empty:
                _r=_match.iloc[0].to_dict()
            else:
                _r={
                    'Tarih':_v.get('Tarih',''),'Başlık':_v.get('Gelişme',''),
                    'URL':_v.get('URL',''),'Risk_Skoru':_v.get('Risk',0),
                    'İçerik_Özeti':'','Kaynak':'','Kategori':''
                }
            _r['Değer_Skoru']=int(_v.get('Değer_Skoru',0) or 0)
            _r['Neden_Değerli']=_v.get('Neden_Değerli','')
            _r['Kaynak_Sayısı']=int(_v.get('Kaynak_Sayısı',0) or 0)
            _know_rows.append(_r)

        _know_select=pd.DataFrame(_know_rows)
        if 'Seç' not in _know_select.columns:
            _know_select.insert(0,'Seç',False)

        _edited_know=st.data_editor(
            _know_select[['Seç','Tarih','Başlık','İçerik_Özeti','Değer_Skoru',
                          'Neden_Değerli','Kaynak_Sayısı','Risk_Skoru','URL']],
            column_config={
                'Seç':st.column_config.CheckboxColumn('Seç'),
                'Değer_Skoru':st.column_config.ProgressColumn('Değer Skoru',min_value=0,max_value=100,format='%d/100'),
                'Risk_Skoru':st.column_config.NumberColumn('Risk',format='%d/100'),
                'URL':st.column_config.LinkColumn('Haber Linki'),
                'İçerik_Özeti':st.column_config.TextColumn('Kısa İçerik',width='large')
            },
            disabled=['Tarih','Başlık','İçerik_Özeti','Değer_Skoru','Neden_Değerli',
                      'Kaynak_Sayısı','Risk_Skoru','URL'],
            hide_index=True,use_container_width=True,
            height=min(480,100+62*len(_know_select)),
            key='v61_now_to_know_editor'
        )

        _selected_idx=_edited_know.index[_edited_know['Seç'].astype(bool)].tolist()
        _selected_know=_know_select.loc[_selected_idx].copy() if _selected_idx else pd.DataFrame()

        k1,k2,k3,k4=st.columns(4)
        with k1:
            if st.button('📌 Önemli Gelişmelere Ekle',key='v82_know_imp',use_container_width=True):
                if _selected_know.empty:
                    st.warning('Önce en az bir gelişmeyi seçin.')
                else:
                    _n=_add_rows_to_important_basket(_selected_know.to_dict('records'))
                    st.success(f'{_n} haber önemli gelişmeler sepetine eklendi.')
        with k2:
            if st.button('🗂️ AKT Sepetine Ekle',key='v82_know_akt',use_container_width=True):
                if _selected_know.empty:
                    st.warning('Önce en az bir gelişmeyi seçin.')
                else:
                    _n=_add_rows_to_osint_basket(_selected_know.to_dict('records'))
                    st.success(f'{_n} haber açık kaynak tarama sepetine eklendi.')
        with k3:
            if st.button('🖥️ Sunum Sepetine Ekle',key='v82_know_pres',use_container_width=True):
                if _selected_know.empty:
                    st.warning('Önce en az bir gelişmeyi seçin.')
                else:
                    _n=_v80_add_presentation(_selected_know.to_dict('records'))
                    st.success(f'{_n} haber sunum sepetine eklendi.')
        with k4:
            if st.button('📝 Detaylı Bilgi Notu Oluştur',key='v82_know_note',use_container_width=True):
                if _selected_know.empty:
                    st.warning('Önce en az bir gelişmeyi seçin.')
                else:
                    with st.spinner(f'{len(_selected_know)} seçili gelişmenin ayrıntılı bilgi notu hazırlanıyor...'):
                        try:
                            st.session_state['v61_know_note_bytes']=make_analyst_docx(
                                _selected_know,
                                title='SANAYİ & TEKNOLOJİ BİLGİ NOTU'
                            )
                            _v63_mark_notes(_selected_know.to_dict('records'))
                        except Exception as _e:
                            st.session_state['v61_know_note_bytes']=None
                            st.error(f'Bilgi notu hazırlanamadı: {_e}')

        if st.session_state.get('v61_know_note_bytes'):
            st.download_button(
                '⬇️ Hazırlanan Bilgi Notunu İndir',
                data=st.session_state['v61_know_note_bytes'],
                file_name=f'Sanayi_Teknoloji_Bilgi_Notu_Su_An_Bilmen_Gerekenler_{date.today()}.docx',
                mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                use_container_width=True,
                key='v61_know_note_download'
            )

st.markdown('---')

# ============================================================
# V68 — KONTROL MERKEZİ
# ============================================================
st.caption('⚡ V120 performans/kalite modu: tarama sonrası panel özetleri önceden hesaplanır; seçim kutuları tek başına ağır analizleri yeniden çalıştırmaz.')
st.subheader('🎛️ Kontrol Merkezi')
st.caption(
    'Bu alan çalışma saatine ve içeriğin niteliğine göre işlem önermektedir: Bilgi Notu için veri/istatistik, '
    'resmî açıklama ve ürün/teknoloji gelişmeleri; AKT için negatif/eleştirel/olumsuz veya propaganda niteliğindeki '
    'içerikler; sunum için resmî veri, resmî açıklama ve teyitli bilgiler esas alınmaktadır.'
)

_cmd_rows=st.session_state.get('rows')
if _cmd_rows:
    _cmd_df=pd.DataFrame(_cmd_rows)
    if not _cmd_df.empty and 'Tarih_dt' in _cmd_df.columns:
        _cmd_df['Tarih_dt']=pd.to_datetime(_cmd_df['Tarih_dt'],utc=True,errors='coerce')

    _cmd,_phase,_phase_hint=_v68_control_center(_cmd_df,8)

    cphase1,cphase2=st.columns([1,2])
    with cphase1:
        st.info(f'**Çalışma Fazı**\n\n{_phase}')
    with cphase2:
        st.info(f'**Sistem Önceliği**\n\n{_phase_hint}')

    if _cmd.empty:
        st.success('Şu anda ayrıca işlem önerilecek yüksek öncelikli bir gelişme bulunmamaktadır.')
    else:
        st.markdown(f'**Şimdi yapılması önerilen {len(_cmd)} işlem**')
        _cmd_for_select=_cmd.copy()
        _section_select_table(
            'v68_command_center',
            _cmd_for_select,
            ['Öncelik','Önerilen_İşlem','Tarih','Başlık','Neden','Durum',
             'Değer_Skoru','Risk_Skoru','URL'],
            height=min(620,105+58*len(_cmd_for_select))
        )
        st.caption(
            'Öneriler karar yerine geçmemektedir. Haberi seçerek doğrudan Önemli Gelişmeler, AKT veya '
            'Bilgi Notu işlemlerini aynı bölümden uygulayabilirsiniz.'
        )
else:
    st.info('İlk ana tarama tamamlandığında Kontrol Merkezi otomatik olarak işlem önerileri oluşturacaktır.')

st.markdown('---')

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
        # V46 — ANA GÖRÜNÜM EN ÜSTTE
        # Tarama tamamlandığında ilk bölüm doğrudan Kronolojik / Negatif / Yüksek Risk vb. ana haber görünümüdür.
        # Performans: Streamlit tabs içindeki TÜM içerikleri arka planda çalıştırır.
        # Bu nedenle tek seferde yalnızca seçilen görünümü üretiriz.
        st.subheader('🌅 Vardiya Başlangıç Özeti')
        st.caption('Sabah ilk analitik bakış: aynı olay tekilleştirilir; resmî veri/açıklama, stratejik sanayi-teknoloji gelişmesi, kritik negatif, savunma/uzay/teknoloji programı ve yüksek teyitli yeni gelişmeler önceliklendirilir.')
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

            st.markdown('**Sabah ilk bakılması gereken 5–8 gelişme**')
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
                _change_type_col='Tür' if 'Tür' in changes.columns else 'Değişim'
                new_n=int(changes[_change_type_col].astype(str).str.contains('YENİ OLAY').sum())
                upd_n=int(changes[_change_type_col].astype(str).str.contains('YENİ BİLGİ').sum())
                risk_n=int(changes[_change_type_col].astype(str).str.contains('RİSK ARTTI').sum())
                ver_n=int(changes[_change_type_col].astype(str).str.contains('TEYİT').sum())
                q1,q2,q3,q4=st.columns(4)
                q1.metric('Yeni Olay',new_n)
                q2.metric('Yeni Bilgi',upd_n)
                q3.metric('Risk Artışı',risk_n)
                q4.metric('Teyit Güçlendi',ver_n)
                changes_view=changes.head(25).copy()
                _section_select_table(
                    'changes',
                    changes_view,
                    ['Ne Değişti?','Tür','Başlık','Kaynak','Kategori','Risk','Önceki Risk','Kaynak Sayısı','URL'],
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

        # Kritik sanayi olayları için SABİT bölüm.
        # Her zaman görünür; olay varsa içerik dolar, yoksa boş durum gösterilir.
        st.subheader('🚨 Kritik Sanayi Olayları — OSB / OSB Dışı Yangın ve Patlama')
        st.caption('OSB ve OSB dışındaki fabrika, tesis ve sanayi alanlarında tespit edilen yangın/patlama olayları burada sürekli izlenir.')

        # V113: her satırda incident fonksiyonunu iki kez çalıştırmak yerine
        # aynı tarama için bir kez hesaplanan cache'li tablo kullanılır.
        critical_events=_v113_critical_events_table(df)

        if not critical_events.empty:
            st.error(f'🚨 **KRİTİK SANAYİ OLAYI ALARMI — {len(critical_events)} içerik tespit edildi**')
            _section_select_table(
                'critical_industrial_events',
                critical_events,
                ['Tarih','Kritik_Olay','Kaynak','Başlık','Risk_Skoru','URL'],
                height=min(340,70+36*len(critical_events))
            )
        else:
            st.info('Bu tarama döneminde OSB / OSB dışı sanayi tesisi yangını veya patlaması tespit edilmedi.')

        st.markdown('---')
        st.subheader('🏆 Günün En Değerli 10 Gelişmesi')
        st.caption(
            'Aynı olaya ait haberlar tek gelişmede birleştirilir. Değer Skoru; önem/risk, farklı kaynak sayısı, '
            'resmî teyit, güncellik, stratejik sanayi-teknoloji önemi, negatif/eleştirel etki ve haber yoğunluğunu birlikte değerlendirir. '
            'Kaynak gerçek okunma/tıklanma verisi sağlıyorsa ileride ayrıca eklenebilir; mevcut sistem erişilemeyen okunma sayılarını tahmin etmez.'
        )
        value10=_v52_event_value_table(df,10)
        if value10.empty:
            st.info('Bu taramada sıralanabilecek gelişme bulunamadı.')
        else:
            _section_select_table(
                'daily_top10_value',
                value10,
                ['Sıra','Değer_Skoru','Tarih','Gelişme','Neden_Değerli',
                 'Kaynak_Sayısı','Haber_Sayısı','Resmî_Teyit','Risk','URL'],
                height=min(680,105+55*len(value10))
            )

            if st.button('📊 BUGÜNÜN DURUM ÖZETİNİ OLUŞTUR',use_container_width=True,key='v54_top10_summary_btn'):
                with st.spinner('Yalnızca en değerli 10 gelişmenin haber içerikleri okunuyor ve özetleniyor...'):
                    summary_text=_v54_deep_top10_summary(df,value10,45)
                    st.session_state.daily_summary_text=summary_text
                    st.session_state.daily_summary_bytes=make_v54_top10_summary_docx(df,value10,summary_text)

            if st.session_state.get('daily_summary_text'):
                st.text_area(
                    'Bugünün Durum Özeti — En Değerli 10 Gelişme',
                    st.session_state.daily_summary_text,
                    height=520,
                    key='v54_daily_summary_preview'
                )
                st.caption(
                    'Özet yalnızca yukarıdaki 10 gelişmenin haber içeriğini anlatır; değer skoru, kaynak sayısı ve '
                    'sıralama gerekçeleri metne eklenmez. Toplam çıktı 45 satırı geçmez.'
                )
                if st.session_state.get('daily_summary_bytes'):
                    st.download_button(
                        '⬇️ BUGÜNÜN DURUM ÖZETİNİ WORD OLARAK İNDİR',
                        data=st.session_state.daily_summary_bytes,
                        file_name=f'bugunun_durum_ozeti_top10_{datetime.now().strftime("%Y%m%d_%H%M")}.docx',
                        mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                        use_container_width=True,
                        key='v54_top10_summary_download'
                    )

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

        st.markdown('---')
        st.subheader('👀 Kaçırıyor Olabilir Miyim? — İkinci Göz')
        st.caption(
            'Mevcut taramada yüksek değer/risk taşıdığı hâlde henüz Önemli Gelişmeler veya '
            'Açık Kaynak Tarama sepetine alınmamış olayları otomatik gösterir.'
        )
        _missed=_v63_missed_candidates(df,12)
        if _missed.empty:
            st.success('Şu anda sepetler dışında kalan belirgin yüksek değerli bir gelişme görünmüyor.')
        else:
            st.warning(f'Henüz hiçbir sepete alınmamış {_missed.shape[0]} dikkat çekici gelişme var.')
            _section_select_table(
                'v63_missed',
                _missed.rename(columns={'Gelişme':'Başlık'}),
                ['Tarih','Başlık','Değer_Skoru','Neden_Değerli','Kaynak_Sayısı','Risk','URL'],
                height=min(620,100+48*len(_missed))
            )

        view=st.radio(
            'Görünüm',
            ['📰 Kronolojik','⚠️ Negatif','🚨 Yüksek Risk','🇹🇷 Türk','🇬🇷 Yunan','🧩 Olaylar','📈 Trend / Analiz','⭐ Takip Listesi'],
            horizontal=True,
            key='main_view'
        )

        cols=['Seç','Tarih','Kaynak_Grubu','Kaynak','Kategori','Başlık','İçerik_Özeti','Duygu','Risk_Skoru','Risk_Durumu','Kaynak_Güvenilirliği','Doğrulama','URL']

        if view=='📰 Kronolojik':
            st.caption(
                '☑️ Hızlı işlem modu: kutucuklara tıklarken sayfa yeniden çalıştırılmaz. '
                'Seçiminizi yaptıktan sonra aşağıdaki işlem düğmelerinden birine basmanız yeterlidir.'
            )
            group_events=st.toggle(
                '🧩 Aynı olayı tek satırda göster',
                value=True,
                key='v109_chron_group_events',
                help='Aynı gelişmenin farklı kaynaklardaki haberlerini tek olay satırında birleştirir.'
            )
            chronology_base=_v109_chronology_events(df) if group_events else df.copy()

            page_size=40
            total_pages=max(1,(len(chronology_base)+page_size-1)//page_size)
            page_no=st.number_input(
                'Sayfa',min_value=1,max_value=total_pages,value=1,step=1,
                key='news_page'
            )
            start_i=(int(page_no)-1)*page_size
            end_i=min(start_i+page_size,len(chronology_base))
            page_df=chronology_base.iloc[start_i:end_i].copy()

            page_df['İçerik_Özeti']=page_df['İçerik_Özeti'].astype(str).str.slice(0,220)
            page_df=_v63_add_status_badges(page_df)
            chron_cols=[
                'Seç','Tarih','Kaynak_Grubu','Kaynak','Kaynak Sayısı','Haber Sayısı',
                'Kategori','Başlık','Durum','İçerik_Özeti','Duygu','Risk_Skoru',
                'Risk_Durumu','Kaynak_Güvenilirliği','Doğrulama','URL'
            ]
            chron_cols=[c for c in chron_cols if c in page_df.columns]

            st.caption(f'{start_i+1}-{end_i} / {len(chronology_base)} kayıt' + (' (olay bazlı)' if group_events else ' haber'))

            # FORM: checkbox tıklamaları rerun yapmaz. Yalnız işlem butonuna basınca tek rerun olur.
            with st.form(
                key=f'v74_chronology_fast_form_{int(page_no)}',
                clear_on_submit=False
            ):
                edited=st.data_editor(
                    page_df[chron_cols],
                    column_config={
                        'Seç':st.column_config.CheckboxColumn('Seç'),
                        'URL':st.column_config.LinkColumn('Haber Linki'),
                        'İçerik_Özeti':st.column_config.TextColumn('Kısa İçerik',width='large'),
                        'Risk_Skoru':st.column_config.NumberColumn('Risk',format='%d/100'),
                        'Kaynak Sayısı':st.column_config.NumberColumn('Kaynak',format='%d'),
                        'Haber Sayısı':st.column_config.NumberColumn('Haber',format='%d'),
                        'Durum':st.column_config.TextColumn('Durum',width='large')
                    },
                    disabled=[x for x in chron_cols if x!='Seç'],
                    hide_index=True,
                    use_container_width=True,
                    height=535,
                    key=f'v74_chron_editor_{int(page_no)}'
                )

                st.markdown('### ⚡ Seçilen Haberlerle Hızlı İşlem')
                c1,c2,c3,c4=st.columns(4)
                with c1: do_imp=st.form_submit_button('📌 Önemli Gelişmelere Ekle',use_container_width=True)
                with c2: do_akt=st.form_submit_button('🗂️ AKT Sepetine Ekle',use_container_width=True)
                with c3: do_pres=st.form_submit_button('🖥️ Sunum Sepetine Ekle',use_container_width=True)
                with c4: do_note=st.form_submit_button('📝 Detaylı Bilgi Notu Oluştur',use_container_width=True)

            if do_imp or do_akt or do_note or do_pres:
                selected_mask=edited['Seç'].astype(bool).to_numpy()
                selected_page=page_df.loc[selected_mask].copy()

                if selected_page.empty:
                    st.warning('Önce en az bir haberi işaretleyin.')
                elif do_imp:
                    n=_v74_fast_add_important(selected_page.to_dict('records'))
                    st.success(f'✅ {n} yeni haber Önemli Gelişmeler Sepeti’ne eklenmiştir.')
                elif do_akt:
                    n=_v74_fast_add_osint(selected_page.to_dict('records'))
                    st.success(f'✅ {n} yeni haber Açık Kaynak Tarama Sepeti’ne eklenmiştir.')
                elif do_pres:
                    n=_v80_add_presentation(selected_page.to_dict('records'))
                    st.success(f'✅ {n} yeni haber Sunum Sepeti’ne eklenmiştir.')
                elif do_note:
                    with st.spinner(
                        f'{len(selected_page)} seçili haber için ayrıntılı bilgi notu hazırlanmaktadır...'
                    ):
                        try:
                            # Tam içerik için kısa page_df yerine ana df'deki aynı URL'leri kullan.
                            selected_urls=set(selected_page['URL'].fillna('').astype(str))
                            full_selected=df[df['URL'].fillna('').astype(str).isin(selected_urls)].copy()
                            if full_selected.empty:
                                full_selected=selected_page.copy()
                            st.session_state['v74_chron_note_bytes']=make_analyst_docx(
                                full_selected,
                                title='SANAYİ & TEKNOLOJİ BİLGİ NOTU'
                            )
                            _v63_mark_notes(full_selected.to_dict('records'))
                            _v73_invalidate_status_cache()
                            st.success('✅ Bilgi notu hazırlanmıştır.')
                        except Exception as e:
                            st.session_state['v74_chron_note_bytes']=None
                            st.error(f'Bilgi notu hazırlanamadı: {e}')

            if st.session_state.get('v74_chron_note_bytes'):
                st.download_button(
                    '⬇️ KRONOLOJİDEN HAZIRLANAN BİLGİ NOTUNU İNDİR',
                    data=st.session_state['v74_chron_note_bytes'],
                    file_name=f'Sanayi_Teknoloji_Bilgi_Notu_{date.today()}.docx',
                    mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    use_container_width=True,
                    key='v74_chron_note_download'
                )


            if group_events and not page_df.empty and '_Olay_ID' in page_df.columns:
                with st.expander('🔎 Olayın tüm kaynaklarını aç',False):
                    _event_options={
                        f"{str(r.get('Başlık',''))[:120]} — {int(r.get('Kaynak Sayısı',1) or 1)} kaynak":str(r.get('_Olay_ID',''))
                        for _,r in page_df.iterrows()
                    }
                    if _event_options:
                        _event_label=st.selectbox(
                            'Kaynaklarını görmek istediğiniz olay',
                            list(_event_options.keys()),
                            key=f'v109_event_source_select_{int(page_no)}'
                        )
                        _event_sources=_v109_event_sources(df,_event_options.get(_event_label))
                        if _event_sources.empty:
                            st.info('Bu olay için ayrıntılı kaynak kaydı bulunamadı.')
                        else:
                            st.dataframe(
                                _event_sources[['Tarih','Kaynak','Başlık','İçerik_Özeti','URL']].head(20),
                                column_config={
                                    'URL':st.column_config.LinkColumn('Haber Linki'),
                                    'İçerik_Özeti':st.column_config.TextColumn('Kısa İçerik',width='large')
                                },
                                hide_index=True,use_container_width=True,
                                height=min(520,80+42*len(_event_sources.head(20)))
                            )

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




        # V34 — ÖNEMLİ GELİŞMELER SEPETİ
        # ---------------------------------------------------------
        st.subheader('📌 24 Saatlik Önemli Gelişmeler Sepeti')
        st.caption('Gün boyunca önemli gördüğünüz haberleri burada biriktirin; vardiya sonunda Word olarak alın.')

        # V74: Kronoloji artık kendi hızlı işlem düğmelerine sahiptir.
        # Burada yalnız diğer bölümlerdeki seçili kayıtlar toplanır.
        selected_from_sections=_collect_section_selected_from_main_df(df)

        if st.button('➕ BÖLÜMLERDE İŞARETLEDİKLERİMİ ÖNEMLİ GELİŞMELER SEPETİNE EKLE',use_container_width=True):
            if selected_from_sections.empty:
                st.warning('Önce herhangi bir bölümde haberlerin yanındaki kutucuklardan seçim yapın.')
            else:
                added=_add_rows_to_important_basket(selected_from_sections.to_dict('records'))
                st.success(f'{added} yeni gelişme sepete eklendi.')

        basket=_load_important_basket()
        if basket.empty:
            st.info('Önemli gelişmeler sepeti şu anda boş.')
        else:
            # -----------------------------------------------------
            # V78 — ÖGN SEPETİ: SİLME ve BİLGİ NOTU TAMAMEN AYRI
            # -----------------------------------------------------
            basket_view=basket[['id','news_time','source','category','title','risk_score','risk_status','url']].copy()
            basket_view=_v63_add_status_badges(basket_view)

            # A) Sadece silme işlemi için checkbox.
            delete_view=basket_view.copy()
            delete_view.insert(0,'Sil',False)
            with st.form('v78_important_basket_delete_form',clear_on_submit=False):
                edited_delete=st.data_editor(
                    delete_view,
                    column_config={
                        'Sil':st.column_config.CheckboxColumn('Sil'),
                        'url':st.column_config.LinkColumn('Haber Linki'),
                        'risk_score':st.column_config.NumberColumn('Risk',format='%d/100'),
                        'Durum':st.column_config.TextColumn('Durum',width='large')
                    },
                    disabled=[c for c in delete_view.columns if c!='Sil'],
                    hide_index=True,use_container_width=True,
                    height=min(430,80+36*len(delete_view)),
                    key='v78_important_basket_delete_editor'
                )
                remove_btn=st.form_submit_button(
                    '🗑️ İŞARETLENENLERİ SEPETTEN ÇIKAR',
                    use_container_width=True
                )

            if remove_btn:
                ids=edited_delete.loc[edited_delete['Sil']==True,'id'].astype(int).tolist()
                removed=_remove_basket_ids(ids)
                st.success(f'{removed} kayıt sepetten çıkarıldı.')

            # B) Bilgi notunda TEK HABER seçilir. Sepetin tamamı hiçbir şekilde
            # make_analyst_docx'e gönderilmez.
            st.markdown('### 📝 Sepetten Seçilen Tek Haberden Detaylı Bilgi Notu')
            option_rows=[]
            for _,r in basket.iterrows():
                clean_title=_clean_note_text(r.get('title',''))
                option_rows.append((
                    int(r.get('id')),
                    f"{clean_title} — {_clean_note_text(r.get('source',''))}"
                ))

            label_to_id={label:rid for rid,label in option_rows}
            selected_label=st.selectbox(
                'Bilgi notu oluşturulacak haber',
                options=list(label_to_id.keys()),
                key='v78_ogn_note_single_select'
            ) if option_rows else None

            if st.button(
                '📝 SEÇİLEN TEK HABERDEN DETAYLI BİLGİ NOTU OLUŞTUR',
                use_container_width=True,
                key='v78_ogn_note_single_button'
            ):
                if not selected_label:
                    st.warning('Bilgi notu için bir haber seçin.')
                else:
                    selected_id=int(label_to_id[selected_label])
                    # Kesin tek satır: ID eşleşmesi + head(1).
                    selected_basket=basket[basket['id'].astype(int)==selected_id].head(1).copy()

                    if selected_basket.empty:
                        st.error('Seçilen haber sepette bulunamadı.')
                    else:
                        r=selected_basket.iloc[0]
                        important_note_rows=pd.DataFrame([{
                            'Tarih':_clean_note_text(r.get('news_time','')),
                            'Kaynak':_clean_note_text(r.get('source','')),
                            'Başlık':_clean_note_text(r.get('title','')),
                            'İçerik_Özeti':_clean_note_text(r.get('summary','')),
                            'URL':str(r.get('url','') or ''),
                            'Kategori':_clean_note_text(r.get('category','')),
                            'Risk_Skoru':r.get('risk_score',0),
                            'Risk_Durumu':_clean_note_text(r.get('risk_status',''))
                        }])

                        # Güvenlik kontrolü: make_analyst_docx'e asla 1'den fazla satır gitmesin.
                        important_note_rows=important_note_rows.head(1)

                        with st.spinner('Seçilen tek haberin tam metni okunuyor ve detaylı bilgi notu hazırlanıyor...'):
                            try:
                                st.session_state['v78_ogn_note_bytes']=make_analyst_docx(
                                    important_note_rows,
                                    title='SANAYİ & TEKNOLOJİ BİLGİ NOTU'
                                )
                                st.session_state['v78_ogn_note_title']=important_note_rows.iloc[0]['Başlık']
                                _v63_mark_notes(important_note_rows.to_dict('records'))
                                _v73_invalidate_status_cache()
                                st.success('✅ Bilgi notu yalnızca seçilen tek haberden hazırlanmıştır.')
                            except Exception as e:
                                st.session_state['v78_ogn_note_bytes']=None
                                st.error(f'Bilgi notu hazırlanamadı: {e}')

            if st.session_state.get('v78_ogn_note_bytes'):
                st.info(
                    'Bilgi notuna alınan tek haber: '
                    + _clean_note_text(st.session_state.get('v78_ogn_note_title',''))
                )
                st.download_button(
                    '⬇️ SEÇİLEN TEK HABERİN DETAYLI BİLGİ NOTUNU İNDİR',
                    data=st.session_state['v78_ogn_note_bytes'],
                    file_name=f'OGN_Secilen_Haber_Bilgi_Notu_{date.today()}.docx',
                    mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    use_container_width=True,
                    key='v78_ogn_note_download'
                )

            if selected_label and st.button('🖥️ SEÇİLEN ÖNEMLİ GELİŞMEYİ SUNUM SEPETİNE EKLE',use_container_width=True,key='v81_ogn_to_pres'):
                _one=basket[basket['id'].astype(int)==int(label_to_id[selected_label])].head(1)
                st.success(f"✅ {_v80_add_presentation(_v81_basket_to_rows(_one))} haber Sunum Sepeti’ne eklenmiştir.")

            # V90: önceki sürümlerden kalan Word bytes kesinlikle kullanılmaz.
            if st.session_state.get('_ogn_engine_version') != V90_OGN_ENGINE_VERSION:
                st.session_state['_ogn_engine_version']=V90_OGN_ENGINE_VERSION
                st.session_state.pop('v90_ogn_docx_bytes',None)
                st.session_state.pop('basket_docx_bytes',None)

            b1,b2=st.columns(2)
            with b1:
                if st.button('📄 ÖNEMLİ GELİŞMELER WORD OLUŞTUR',use_container_width=True,key='v90_make_ogn_word'):
                    # Her basışta eski çıktı silinir ve V90 motoruyla baştan hazırlanır.
                    st.session_state.pop('v90_ogn_docx_bytes',None)
                    with st.spinner('Önemli gelişmeler gerçek haber içeriklerinden resmî biçimde özetleniyor...'):
                        try:
                            st.session_state['v90_ogn_docx_bytes']=make_important_basket_docx_v101(basket)
                        except ReportQualityError as _quality_error:
                            st.session_state['v90_ogn_docx_bytes']=None
                            st.error(str(_quality_error))
                        except Exception as _ogn_error:
                            st.session_state['v90_ogn_docx_bytes']=None
                            st.error(f'Önemli Gelişmeler Notu hazırlanamadı: {_ogn_error}')
                if st.session_state.get('v90_ogn_docx_bytes'):
                    st.download_button(
                        '⬇️ 24 SAATLİK ÖNEMLİ GELİŞMELER / WORD',
                        st.session_state['v90_ogn_docx_bytes'],
                        file_name=f'STB_Onemli_Gelismeler_Notu_{date.today()}.docx',
                        mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                        use_container_width=True,
                        key='v90_download_ogn_word'
                    )
            with b2:
                if st.button('🧹 SEPETİ TAMAMEN TEMİZLE',use_container_width=True):
                    removed=_clear_important_basket()
                    st.success(f'{removed} kayıt silindi.')


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
            # -----------------------------------------------------
            # V79 — AKT SEPETİ: ÖGN İLE AYNI TEK HABER BİLGİ NOTU MANTIĞI
            # -----------------------------------------------------
            osint_view=osint_basket[['id','news_time','source','category','title','risk_score','risk_status','url']].copy()
            osint_view=_v63_add_status_badges(osint_view)

            # A) Silme işlemi ayrı checkbox formunda kalır.
            delete_osint_view=osint_view.copy()
            delete_osint_view.insert(0,'Sil',False)
            with st.form('v79_osint_basket_delete_form',clear_on_submit=False):
                edited_osint=st.data_editor(
                    delete_osint_view,
                    column_config={
                        'Sil':st.column_config.CheckboxColumn('Sil'),
                        'url':st.column_config.LinkColumn('Haber Linki'),
                        'risk_score':st.column_config.NumberColumn('Risk',format='%d/100'),
                        'Durum':st.column_config.TextColumn('Durum',width='large')
                    },
                    disabled=[c for c in delete_osint_view.columns if c!='Sil'],
                    hide_index=True,use_container_width=True,
                    height=min(430,80+36*len(delete_osint_view)),
                    key='v79_osint_basket_delete_editor'
                )
                remove_osint=st.form_submit_button(
                    '🗑️ İŞARETLENENLERİ AKT SEPETİNDEN ÇIKAR',
                    use_container_width=True
                )

            if remove_osint:
                ids=edited_osint.loc[edited_osint['Sil']==True,'id'].astype(int).tolist()
                removed=_remove_osint_basket_ids(ids)
                st.success(f'{removed} kayıt AKT sepetinden çıkarıldı.')

            # AKT raporu sepetin tamamından hazırlanabilir; bu davranış korunur.
            osint_rows=[]
            for _,r in osint_basket.iterrows():
                osint_rows.append({
                    'Tarih':_clean_note_text(r.get('news_time','')),
                    'Kaynak':_clean_note_text(r.get('source','')),
                    'Başlık':_clean_note_text(r.get('title','')),
                    'İçerik_Özeti':_clean_note_text(r.get('summary','')),
                    'URL':str(r.get('url','') or ''),
                    'Kategori':_clean_note_text(r.get('category','')),
                    'Risk_Skoru':r.get('risk_score',0),
                    'Risk_Durumu':_clean_note_text(r.get('risk_status','')),
                    'Yayıncı':_clean_note_text(r.get('source','')),
                    'Yayıncı_URL':''
                })

            _akt_pres_opts={f"{_clean_note_text(r.get('title',''))} — {_clean_note_text(r.get('source',''))}":int(r.get('id')) for _,r in osint_basket.iterrows()}
            _akt_pres_label=st.selectbox('Sunuma eklenecek AKT haberi',list(_akt_pres_opts.keys()),key='v81_akt_pres_select') if _akt_pres_opts else None
            if st.button('🖥️ SEÇİLEN AKT HABERİNİ SUNUM SEPETİNE EKLE',use_container_width=True,key='v81_akt_to_pres'):
                if _akt_pres_label:
                    _one=osint_basket[osint_basket['id'].astype(int)==_akt_pres_opts[_akt_pres_label]].head(1)
                    st.success(f"✅ {_v80_add_presentation(_v81_basket_to_rows(_one))} haber Sunum Sepeti’ne eklenmiştir.")

            ob1,ob2=st.columns(2)
            with ob1:
                if st.button('📝 AKT SEPETİNDEN WORD HAZIRLA',use_container_width=True,key='v79_akt_report'):
                    with st.spinner('AKT sepetindeki haberler rapora hazırlanıyor...'):
                        try:
                            st.session_state.docx_bytes=make_docx(osint_rows)
                        except ReportQualityError as _quality_error:
                            st.session_state.docx_bytes=None
                            st.error(str(_quality_error))
                        except Exception as _akt_error:
                            st.session_state.docx_bytes=None
                            st.error(f'AKT raporu hazırlanamadı: {_akt_error}')
                if st.session_state.get('docx_bytes'):
                    st.download_button(
                        '⬇️ AKT SEPETİNDEN AÇIK KAYNAK RAPORU / WORD',
                        st.session_state.docx_bytes,
                        file_name=f'Sanayi_Teknoloji_Acik_Kaynak_Sepet_{date.today()}.docx',
                        mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                        use_container_width=True,
                        key='v79_akt_report_download'
                    )
            with ob2:
                if st.button('🧹 AKT SEPETİNİ TAMAMEN TEMİZLE',use_container_width=True,key='v79_clear_akt'):
                    removed=_clear_osint_basket()
                    st.success(f'{removed} kayıt silindi.')

            # B) Bilgi notu için yalnız TEK HABER seçilir.
            st.markdown('### 📝 AKT Sepetinden Seçilen Tek Haberden Detaylı Bilgi Notu')

            akt_option_rows=[]
            for _,r in osint_basket.iterrows():
                clean_title=_clean_note_text(r.get('title',''))
                akt_option_rows.append((
                    int(r.get('id')),
                    f"{clean_title} — {_clean_note_text(r.get('source',''))}"
                ))

            akt_label_to_id={label:rid for rid,label in akt_option_rows}
            selected_akt_label=st.selectbox(
                'Bilgi notu oluşturulacak AKT haberi',
                options=list(akt_label_to_id.keys()),
                key='v79_akt_note_single_select'
            ) if akt_option_rows else None

            if st.button(
                '📝 SEÇİLEN TEK AKT HABERİNDEN DETAYLI BİLGİ NOTU OLUŞTUR',
                use_container_width=True,
                key='v79_akt_note_single_button'
            ):
                if not selected_akt_label:
                    st.warning('Bilgi notu için bir AKT haberi seçin.')
                else:
                    selected_akt_id=int(akt_label_to_id[selected_akt_label])
                    # Kesin tek satır: ID eşleşmesi ve head(1).
                    selected_akt=osint_basket[
                        osint_basket['id'].astype(int)==selected_akt_id
                    ].head(1).copy()

                    if selected_akt.empty:
                        st.error('Seçilen AKT haberi sepette bulunamadı.')
                    else:
                        r=selected_akt.iloc[0]
                        akt_note_df=pd.DataFrame([{
                            'Tarih':_clean_note_text(r.get('news_time','')),
                            'Kaynak':_clean_note_text(r.get('source','')),
                            'Başlık':_clean_note_text(r.get('title','')),
                            'İçerik_Özeti':_clean_note_text(r.get('summary','')),
                            'URL':str(r.get('url','') or ''),
                            'Kategori':_clean_note_text(r.get('category','')),
                            'Risk_Skoru':r.get('risk_score',0),
                            'Risk_Durumu':_clean_note_text(r.get('risk_status',''))
                        }]).head(1)

                        with st.spinner('Seçilen tek AKT haberinin tam metni okunuyor ve detaylı bilgi notu hazırlanıyor...'):
                            try:
                                st.session_state['v79_akt_note_bytes']=make_analyst_docx(
                                    akt_note_df,
                                    title='SANAYİ & TEKNOLOJİ BİLGİ NOTU'
                                )
                                st.session_state['v79_akt_note_title']=akt_note_df.iloc[0]['Başlık']
                                _v63_mark_notes(akt_note_df.to_dict('records'))
                                _v73_invalidate_status_cache()
                                st.success('✅ Bilgi notu yalnızca seçilen tek AKT haberinden hazırlanmıştır.')
                            except Exception as e:
                                st.session_state['v79_akt_note_bytes']=None
                                st.error(f'Bilgi notu hazırlanamadı: {e}')

            if st.session_state.get('v79_akt_note_bytes'):
                st.info(
                    'Bilgi notuna alınan tek AKT haberi: '
                    + _clean_note_text(st.session_state.get('v79_akt_note_title',''))
                )
                st.download_button(
                    '⬇️ SEÇİLEN TEK AKT HABERİNİN DETAYLI BİLGİ NOTUNU İNDİR',
                    data=st.session_state['v79_akt_note_bytes'],
                    file_name=f'AKT_Secilen_Haber_Bilgi_Notu_{date.today()}.docx',
                    mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    use_container_width=True,
                    key='v79_akt_note_download'
                )


        st.markdown('---')
        st.subheader('🖥️ Sunum Sepeti')
        st.caption('Önemli Gelişmeler ve AKT sepetlerinin hemen altında yer almaktadır.')
        _pb=_v80_load_presentation()
        if _pb.empty:
            st.info('Sunum sepeti boş.')
        else:
            _pv=_pb[['id','news_time','source','title','url']].copy()
            _pv=_v63_add_status_badges(_pv)
            _pv.insert(0,'Seç',False)
            with st.form('v81_presentation_basket_form',clear_on_submit=False):
                _ped=st.data_editor(_pv,column_config={
                    'Seç':st.column_config.CheckboxColumn('Seç'),
                    'url':st.column_config.LinkColumn('Haber Linki'),
                    'Durum':st.column_config.TextColumn('Durum',width='large')
                },
                    disabled=[c for c in _pv.columns if c!='Seç'],hide_index=True,use_container_width=True,height=min(420,80+36*len(_pv)))
                p1,p2,p3,p4=st.columns(4)
                with p1: _toimp=st.form_submit_button('📌 Önemli Gelişmelere Ekle',use_container_width=True)
                with p2: _toakt=st.form_submit_button('🗂️ AKT Sepetine Ekle',use_container_width=True)
                with p3: _pnote=st.form_submit_button('📝 Bilgi Notu Oluştur',use_container_width=True)
                with p4: _prem=st.form_submit_button('🗑️ Sepetten Çıkar',use_container_width=True)
            _ids=_ped.loc[_ped['Seç']==True,'id'].astype(int).tolist()
            _sel=_pb[_pb['id'].astype(int).isin(_ids)]
            _rows=_v81_basket_to_rows(_sel)
            if _toimp:
                if _rows: st.success(f"✅ {_v74_fast_add_important(_rows)} haber Önemli Gelişmeler Sepeti’ne eklenmiştir.")
                else: st.warning('Önce haber seçin.')
            if _toakt:
                if _rows: st.success(f"✅ {_v74_fast_add_osint(_rows)} haber AKT Sepeti’ne eklenmiştir.")
                else: st.warning('Önce haber seçin.')
            if _prem:
                st.success(f"✅ {_v81_remove_presentation_ids(_ids)} haber çıkarılmıştır.")
            if _pnote:
                if len(_rows)!=1: st.warning('Detaylı bilgi notu için yalnızca bir haber seçin.')
                else:
                    with st.spinner('Seçilen sunum haberinden detaylı bilgi notu hazırlanıyor...'):
                        st.session_state['v81_pres_note_bytes']=make_analyst_docx(pd.DataFrame(_rows).head(1),title='SANAYİ & TEKNOLOJİ BİLGİ NOTU')
            if st.session_state.get('v81_pres_note_bytes'):
                st.download_button('⬇️ SUNUM SEPETİ BİLGİ NOTUNU İNDİR',st.session_state['v81_pres_note_bytes'],
                    file_name=f'Sunum_Sepeti_Bilgi_Notu_{date.today()}.docx',mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    use_container_width=True,key='v81_pres_note_download')

        st.markdown('---')
        st.subheader('🏛️ Resmî Kaynak Radarı')
        st.caption('Sanayi ve Teknoloji Bakanlığı, TÜBİTAK, KOSGEB, TÜRKPATENT, TSE, SSB, TÜİK ve diğer birincil kamu kaynaklarından gelen içerikleri ayrı gösterir.')
        official_radar=_official_radar_rows(df)
        if official_radar.empty:
            st.info('Bu taramada resmî/birincil kaynaklardan eşleşen yeni içerik bulunamadı.')
        else:
            if 'Kurum Türü' not in official_radar.columns:
                official_radar=official_radar.copy()
                official_radar['Kurum Türü']=official_radar.apply(_v109_official_source_type,axis=1)
            _types=['Tümü']+[
                x for x in ['Bakanlık','TÜİK','TÜBİTAK','KOSGEB','TÜRKPATENT','TSE','SSB','Resmî Gazete','Diğer Resmî']
                if x in set(official_radar['Kurum Türü'].astype(str))
            ]
            _official_type=st.radio('Kaynak türü',_types,horizontal=True,key='v109_official_source_type')
            _official_show=official_radar if _official_type=='Tümü' else official_radar[official_radar['Kurum Türü']==_official_type]
            _section_select_table(
                'official_radar_'+re.sub(r'[^a-zA-Z0-9]+','_',norm(_official_type)),
                _official_show.head(30),
                ['Tarih','Kurum Türü','Kaynak','Kategori','Başlık','İçerik_Özeti','Risk_Skoru','Doğrulama','URL'],
                height=min(600,90+38*min(len(_official_show),30))
            )

        st.markdown('---')
        st.subheader('🧭 Olay Yaşam Döngüsü')
        st.caption(
            'Aynı olayın mevcut taramadaki gelişim aşamasını otomatik gösterir: '
            'İlk Sinyal → Gelişiyor → Teyit Edildi → Sonuçlandı. Bu alan sabittir.'
        )
        lifecycle=_v58_event_lifecycle_table(df,25)
        if lifecycle.empty:
            st.info('Bu taramada yaşam döngüsü oluşturulabilecek olay bulunamadı.')
        else:
            _section_select_table(
                'v58_event_lifecycle',
                lifecycle,
                ['Tarih','Aşama','Başlık','Kategori','Kaynak_Sayısı','Haber_Sayısı',
                 'Doğrulama','Risk_Skoru','Aşama_Gerekçesi','URL'],
                height=min(700,100+40*len(lifecycle))
            )


        st.markdown('---')
        st.subheader('📋 Gün Sonu Performans Özeti')
        st.caption('Bugün sistemde oluşan tarama ve çalışma çıktılarının operasyonel özeti.')
        _perf=_v60_day_end_performance(df)
        p1,p2,p3,p4,p5,p6,p7=st.columns(7)
        p1.metric('Tarama',_perf['Taramalar'])
        p2.metric('Benzersiz Olay',_perf['Benzersiz Olay'])
        p3.metric('Negatif',_perf['Negatif'])
        p4.metric('Yüksek Risk',_perf['Yüksek Risk'])
        p5.metric('Önemli Sepet',_perf['Önemli Sepete Eklenen'])
        p6.metric('AKT Sepet',_perf['AKT Sepete Eklenen'])
        p7.metric('Kritik Sanayi',_perf['Kritik Sanayi'])

        st.write(
            f"Bugün {_perf['Taramalar']} tarama gerçekleştirilmiş; geçmiş kayıtlarında "
            f"{_perf['Benzersiz Olay']} benzersiz olay, {_perf['Negatif']} negatif ve "
            f"{_perf['Yüksek Risk']} yüksek riskli gelişme kaydedilmiştir. "
            f"{_perf['Önemli Sepete Eklenen']} içerik önemli gelişmeler sepetine, "
            f"{_perf['AKT Sepete Eklenen']} içerik açık kaynak tarama sepetine eklenmiştir."
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
