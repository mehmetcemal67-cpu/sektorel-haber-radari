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
# V120 ADAY — V119 KARARLI tabanı + hedefli global basın genişletmesi
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
# V125 — GÖRSEL TEMA / ARKA PLAN
# İşlevsel mantığa dokunmaz. Ekteki teknoloji görseli panelin ve
# şifre ekranının ortak arka planı olarak gömülüdür; ayrı asset gerektirmez.
# ============================================================
_V125_BG_DATA = "data:image/jpeg;base64,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"

def _v125_apply_visual_theme():
    st.markdown(
        f"""
        <style>
        :root {
            --stb-bg-deep: #050b18;
            --stb-panel: rgba(6, 15, 35, .72);
            --stb-panel-strong: rgba(5, 13, 31, .88);
            --stb-border: rgba(115, 167, 255, .22);
            --stb-accent: #60a5fa;
            --stb-accent-2: #22d3ee;
            --stb-text: #eef5ff;
            --stb-muted: #a9b8cf;
        }

        html, body, .stApp, [data-testid="stAppViewContainer"] {
            min-height: 100%;
            color: var(--stb-text);
        }

        [data-testid="stAppViewContainer"] {
            background-image:
                linear-gradient(180deg, rgba(2, 7, 20, .77), rgba(3, 9, 24, .88)),
                url("{_V125_BG_DATA}");
            background-size: cover;
            background-position: center center;
            background-repeat: no-repeat;
            background-attachment: fixed;
        }

        [data-testid="stHeader"],
        [data-testid="stToolbar"],
        [data-testid="stDecoration"] {
            background: transparent !important;
        }

        [data-testid="stSidebar"] > div:first-child {
            background: rgba(3, 9, 23, .91) !important;
            border-right: 1px solid rgba(96, 165, 250, .16);
            backdrop-filter: blur(16px);
        }

        .main .block-container {
            background: rgba(4, 12, 30, .54);
            border: 1px solid rgba(115, 167, 255, .13);
            border-radius: 24px;
            box-shadow: 0 22px 70px rgba(0, 0, 0, .28);
            backdrop-filter: blur(9px);
            -webkit-backdrop-filter: blur(9px);
            margin-top: .8rem;
            margin-bottom: 1.8rem;
            padding-left: 2rem;
            padding-right: 2rem;
        }

        h1, h2, h3, h4, h5, h6,
        [data-testid="stMarkdownContainer"] p,
        [data-testid="stMarkdownContainer"] li,
        label, .stCaption {
            color: var(--stb-text);
        }

        [data-testid="stCaptionContainer"], .stCaption {
            color: var(--stb-muted) !important;
        }

        div[data-testid="stForm"],
        div[data-testid="stExpander"],
        div[data-testid="stMetric"],
        div[data-testid="stDataFrame"],
        div[data-testid="stTable"] {
            background: rgba(5, 14, 34, .68);
            border: 1px solid rgba(115, 167, 255, .16);
            border-radius: 16px;
            backdrop-filter: blur(10px);
        }

        div[data-testid="stMetric"] {
            padding: .75rem 1rem;
            box-shadow: inset 0 1px 0 rgba(255,255,255,.025);
        }

        .stTextInput input, .stTextArea textarea, .stNumberInput input,
        [data-baseweb="select"] > div {
            background: rgba(4, 12, 30, .82) !important;
            color: #f7fbff !important;
            border-color: rgba(96, 165, 250, .26) !important;
        }

        [data-baseweb="popover"], [data-baseweb="menu"] {
            background: #071427 !important;
            color: #eef5ff !important;
        }

        .stButton > button, .stDownloadButton > button,
        div[data-testid="stFormSubmitButton"] > button {
            background: linear-gradient(135deg, rgba(30, 64, 175, .88), rgba(8, 145, 178, .86));
            color: #ffffff;
            border: 1px solid rgba(125, 211, 252, .35);
            border-radius: 11px;
            box-shadow: 0 8px 24px rgba(3, 105, 161, .15);
        }

        .stButton > button:hover, .stDownloadButton > button:hover,
        div[data-testid="stFormSubmitButton"] > button:hover {
            border-color: rgba(125, 211, 252, .72);
            box-shadow: 0 10px 30px rgba(14, 165, 233, .24);
            transform: translateY(-1px);
        }

        [data-testid="stAlert"] {
            background: rgba(5, 14, 34, .78);
            border-radius: 14px;
            backdrop-filter: blur(10px);
        }

        a { color: #7dd3fc !important; }

        /* Plotly'nin dış kabı da panel temasıyla bütünleşsin. */
        [data-testid="stPlotlyChart"] {
            background: linear-gradient(180deg, rgba(3,10,25,.58), rgba(3,10,25,.34));
            border: 1px solid rgba(96,165,250,.16);
            border-radius: 18px;
            overflow: hidden;
            box-shadow: 0 18px 40px rgba(0,0,0,.20);
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

_v125_apply_visual_theme()
# ============================================================
# /V125 GÖRSEL TEMA
# ============================================================

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

    # V125 — Şifre ekranında aynı arka plan; giriş kartı sunumluk, odaklı ve yarı saydamdır.
    st.markdown(
        """
        <style>
        [data-testid="stAppViewContainer"] .main .block-container {
            max-width: 560px !important;
            padding-top: 12vh !important;
            padding-bottom: 3rem !important;
            background: rgba(4, 12, 30, .68) !important;
            border: 1px solid rgba(125, 211, 252, .24) !important;
            box-shadow: 0 28px 80px rgba(0, 0, 0, .42) !important;
            backdrop-filter: blur(14px) !important;
            -webkit-backdrop-filter: blur(14px) !important;
        }
        [data-testid="stAppViewContainer"] h1 {
            text-align: center;
            font-size: 2.05rem !important;
            letter-spacing: -.02em;
            text-shadow: 0 2px 26px rgba(59,130,246,.28);
        }
        [data-testid="stAppViewContainer"] [data-testid="stCaptionContainer"] {
            text-align: center;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

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
 'Savunma & Havacılık':['savunma','aselsan','tusaş','tusas','roketsan','havelsan','baykar','bayraktar','iha','siha','kaan','kızılelma','füze','roket','havacılık','defense','defence','aerospace','aviation','missile','uav'],
 'Dijital & Yapay Zeka':['yapay zeka','yapay zekâ','siber','yazılım','5g','6g','veri merkezi','bulut','kuantum','artificial intelligence','machine learning','cybersecurity','cyber attack','data breach','software','cloud','quantum'],
 'Yarı İletken & Elektronik':['çip','mikroçip','yarı iletken','işlemci','elektronik','wafer','pcb','chip','microchip','semiconductor','processor','electronics'],
 'Otomotiv & Mobilite':['otomotiv','togg','elektrikli araç','batarya','şarj','automotive','electric vehicle','electric vehicles','battery','mobility'],
 'Enerji':['enerji','hidrojen','güneş','rüzgar','nükleer','enerji depolama','energy','hydrogen','solar','wind','nuclear','energy storage'],
 'Sanayi & Üretim':['sanayi','imalat','üretim','fabrika','osb','makine','robotik','otomasyon','demir çelik','kimya','industry','industrial','manufacturing','factory','robotics','automation','supply chain','advanced manufacturing'],
 'Uzay & İleri Teknoloji':['uzay','uydu','tua','nanoteknoloji','biyoteknoloji','space','satellite','launch','nanotechnology','biotechnology','advanced materials'],
 'Kurumsal Ekosistem':['tübitak','kosgeb','sanayi ve teknoloji bakanlığı','türkpatent','teknopark','teknofest','innovation','startup','venture capital','research and development','r&d']
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

# V120 — Hedefli global sanayi/teknoloji kaynakları.
GLOBAL_TECH_SOURCES=[
 'technologyreview.com','spectrum.ieee.org','arstechnica.com'
]
GLOBAL_INDUSTRY_SOURCES=[
 'industryweek.com','automationworld.com','manufacturingtomorrow.com'
]
GLOBAL_ECON_TECH_SOURCES=[
 'ft.com','bloomberg.com','nikkei.com'
]
GLOBAL_DEFENSE_AERO_SOURCES=[
 'defensenews.com','aviationweek.com'
]
GLOBAL_PRIORITY_SOURCES=(
 GLOBAL_TECH_SOURCES+GLOBAL_INDUSTRY_SOURCES+
 GLOBAL_ECON_TECH_SOURCES+GLOBAL_DEFENSE_AERO_SOURCES
)

# İngilizce içeriklerin global modda konu dışı diye elenmesini önleyen kontrollü evren.
GLOBAL_TOPIC_TERMS=[
 'industry','industrial','manufacturing','factory','advanced manufacturing',
 'automation','robotics','supply chain','smart factory','digital twin','industrial software',
 'technology','innovation','artificial intelligence','machine learning','generative ai',
 'cybersecurity','cyber attack','data breach','cloud','quantum','telecom','5g','6g',
 'semiconductor','semiconductors','chip','chips','microchip','processor','electronics',
 'defense','defence','aerospace','aviation','missile','drone','uav','radar','space','satellite',
 'automotive','electric vehicle','electric vehicles','battery','mobility',
 'energy','nuclear','hydrogen','solar','wind','energy storage',
 'critical minerals','rare earth','advanced materials','export controls','sanctions',
 'startup','venture capital','research and development','r&d'
]

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
 'tusaş':'tusas.com','tusas':'tusas.com','roketsan':'roketsan.com.tr','havelsan':'havelsan.com.tr','baykar':'baykartech.com','togg':'togg.com.tr',
 'mit technology review':'technologyreview.com','technology review':'technologyreview.com',
 'ieee spectrum':'spectrum.ieee.org','ars technica':'arstechnica.com',
 'industryweek':'industryweek.com','industry week':'industryweek.com',
 'automation world':'automationworld.com','manufacturing tomorrow':'manufacturingtomorrow.com',
 'financial times':'ft.com','bloomberg technology':'bloomberg.com','bloomberg':'bloomberg.com',
 'nikkei asia':'nikkei.com','nikkei':'nikkei.com','defense news':'defensenews.com',
 'aviation week':'aviationweek.com'
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
    for d in TR_MAIN+TR_TECH+TR_OFFICIAL+GR+GLOBAL_PRIORITY_SOURCES:
        stem=d.split('.')[0]
        if stem and stem in re.sub(r'[^a-z0-9ğüşöçıİĞÜŞÖÇ]','',n): return d
    return domain(article_url)

def source_group(d):
    d=domain(d)
    if d in TR_OFFICIAL: return '🇹🇷 Resmi / Kurumsal'
    if d in TR_TECH: return '🇹🇷 Türk Teknoloji / Savunma'
    if d in TR_MAIN: return '🇹🇷 Türk Medyası / Ekonomi'
    if d in GLOBAL_PRIORITY_SOURCES: return '🌍 Global Sanayi / Teknoloji'
    if d in GR: return '🇬🇷 Yunan Medyası — Türk Savunma'
    if d in SOCIAL: return '📱 Açık Sosyal / İndeks'
    return '🌍 Diğer / Açık Kaynak'

def source_rank(d):
    d=domain(d)
    if d in TR_OFFICIAL: return 500
    if d in TR_TECH: return 450
    if d in TR_MAIN: return 400
    if d in GLOBAL_PRIORITY_SOURCES: return 360
    if d in GR: return 300
    if d in SOCIAL: return 250
    return 100

def relevant(text,user_query=''):
    t=norm(text)
    if any(x in t for x in TOPIC_TERMS): return True
    uq=re.split(r'\bOR\b|,|\n',user_query or '',flags=re.I)
    generic={'sanayi','teknoloji','üretim','yatırım','enerji','türkiye','türk','haber'}
    return any(len(x.strip())>2 and norm(x.strip()) not in generic and norm(x.strip()) in t for x in uq)


def global_relevant(text, user_query=''):
    """Global İngilizce kaynaklarda sanayi/teknoloji ilgisini koruyarak filtreler."""
    t=norm(text)
    if any(term in t for term in GLOBAL_TOPIC_TERMS):
        return True
    # Türkçe/özel kullanıcı sorgusu geçen global haberleri de kaybetme.
    return relevant(text,user_query)


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


def rss_global(query, timeout=7):
    """Global kaynaklar için İngilizce Google News dizinini tarar."""
    try:
        r=requests.get(
            'https://news.google.com/rss/search',
            params={'q':query,'hl':'en-US','gl':'US','ceid':'US:en'},
            headers=HEADERS,
            timeout=timeout
        )
        r.raise_for_status()
        root=ET.fromstring(r.content)
        out=[]
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


def build_global_queries(when):
    """V120 — hedefli global basını kaynak bazlı ve İngilizce terimlerle tarar."""
    source_topics={
        'technologyreview.com':'(artificial intelligence OR semiconductor OR quantum OR robotics OR cybersecurity OR biotechnology OR climate technology)',
        'spectrum.ieee.org':'(semiconductor OR electronics OR robotics OR artificial intelligence OR aerospace OR energy OR telecom)',
        'arstechnica.com':'(artificial intelligence OR cybersecurity OR semiconductor OR space OR technology OR energy)',
        'industryweek.com':'(manufacturing OR factory OR industrial OR automation OR supply chain OR workforce OR reshoring)',
        'automationworld.com':'(automation OR robotics OR industrial software OR smart factory OR manufacturing OR digital twin)',
        'manufacturingtomorrow.com':'(manufacturing OR automation OR robotics OR additive manufacturing OR supply chain OR factory)',
        'ft.com':'(technology OR semiconductor OR manufacturing OR industry OR supply chain OR artificial intelligence OR energy OR defense)',
        'bloomberg.com':'(technology OR semiconductor OR manufacturing OR artificial intelligence OR industry OR supply chain OR energy OR defense)',
        'nikkei.com':'(technology OR semiconductor OR manufacturing OR supply chain OR automotive OR battery OR artificial intelligence)',
        'defensenews.com':'(defense OR defence OR aerospace OR missile OR drone OR radar OR military technology OR space)',
        'aviationweek.com':'(aerospace OR aviation OR defense OR defence OR aircraft OR space OR missile OR drone)'
    }
    queries=[f'site:{domain_name} {terms} when:{when}' for domain_name,terms in source_topics.items()]

    # Hedef 11 yayın dışında kalan nitelikli küresel basını da kaçırmamak için
    # beş geniş tema sorgusu çalıştırılır. Kaynak önceliği hedef yayınlarda kalır,
    # ancak global katman yalnız bu 11 alan adıyla sınırlandırılmaz.
    queries.extend([
        f'(manufacturing OR industrial automation OR robotics OR supply chain OR smart factory) (investment OR production OR factory OR capacity OR technology) when:{when}',
        f'(semiconductor OR artificial intelligence OR cybersecurity OR quantum) (industry OR technology OR investment OR regulation OR production) when:{when}',
        f'(defense OR defence OR aerospace OR aviation OR space OR automotive OR electric vehicle OR battery OR energy) (technology OR industry OR manufacturing OR supply chain OR investment) when:{when}',
        f'(Turkey OR Türkiye OR Turkish) (industry OR manufacturing OR technology OR semiconductor OR aerospace OR defense OR automotive OR energy) when:{when}',
        f'(ASELSAN OR TUSAŞ OR TUSAS OR ROKETSAN OR HAVELSAN OR Baykar OR Bayraktar OR KAAN OR TOGG) when:{when}'
    ])
    return queries


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
            # Global katman hedef 11 yayını özel sorgularla önceliklendirir; ayrıca
            # diğer küresel kaynaklardan sanayi/teknoloji açısından ilgili haberleri kabul eder.
            if not global_relevant(t,user_query):
                reasons['konu']+=1
                continue
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
    if d in GLOBAL_PRIORITY_SOURCES: return '🟢 A — Hedefli global kaynak'
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
    if d in GLOBAL_PRIORITY_SOURCES: return '🟢 A — Hedefli global kaynak'
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
    tbl=_v122_add_source_verification(tbl)
    columns=list(columns)
    if 'Durum' not in columns:
        insert_at=columns.index('Başlık')+1 if 'Başlık' in columns else 0
        columns.insert(insert_at,'Durum')
    if 'Kaynak Teyidi' not in columns and 'Kaynak Teyidi' in tbl.columns:
        if 'Kaynak' in columns:
            insert_at=columns.index('Kaynak')+1
        elif 'Başlık' in columns:
            insert_at=columns.index('Başlık')
        else:
            insert_at=0
        columns.insert(insert_at,'Kaynak Teyidi')

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
                'Kaynak Teyidi':st.column_config.TextColumn('Kaynak Teyidi',width='medium'),
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
# V122 — YÖNETİCİ ÖZETİ / KAYNAK TEYİDİ / TÜRKİYE BAĞLANTILI GLOBAL
# V121 kararlı çekirdeği korunur; yalnız görünüm ve karar-destek katmanı eklenir.
# ============================================================

_V122_TURKEY_GLOBAL_TERMS = (
    'turkey', 'türkiye', 'turkiye', 'turkish', 'ankara', 'istanbul',
    'aselsan', 'tusaş', 'tusas', 'roketsan', 'havelsan', 'baykar',
    'bayraktar', 'togg', 'kaan', 'kizilelma', 'kızılelma', 'hisar', 'siper',
    'turkish aerospace', 'turkish defense', 'turkish defence',
    'turkish industry', 'turkish manufacturing', 'turkish technology'
)


def _v122_source_verification_lookup():
    'Mevcut taramayı URL/başlık üzerinden tek kez indeksler.'
    rows = st.session_state.get('rows') or []
    scan_id = st.session_state.get('current_scan_id')
    cache_key = (scan_id, len(rows))
    cached = st.session_state.get('_v122_verification_lookup')
    if cached and cached.get('key') == cache_key:
        return cached.get('lookup', {})

    lookup = {}
    if rows:
        rdf = pd.DataFrame(rows)
        for _, r in rdf.iterrows():
            try:
                source_count = int(r.get('Olay_Kaynak_Sayisi', 1) or 1)
            except Exception:
                source_count = 1
            verification = str(r.get('Doğrulama', '') or '')
            try:
                official = bool(_is_official_radar_row(r)) or 'resm' in norm(verification)
            except Exception:
                official = 'resm' in norm(verification)
            payload = (max(1, source_count), verification, official)
            url = str(r.get('URL', '') or '').strip()
            title = title_key(r.get('Başlık', ''))
            if url:
                old_payload = lookup.get('U:' + url)
                if old_payload is None or payload[0] > old_payload[0]:
                    lookup['U:' + url] = payload
            if title:
                old_payload = lookup.get('T:' + title)
                if old_payload is None or payload[0] > old_payload[0]:
                    lookup['T:' + title] = payload

    st.session_state['_v122_verification_lookup'] = {'key': cache_key, 'lookup': lookup}
    return lookup


def _v122_verification_payload(row):
    'Satır için kaynak sayısı, doğrulama etiketi ve resmî kaynak durumunu döndürür.'
    def _int_value(value, default=1):
        try:
            if pd.isna(value):
                return default
        except Exception:
            pass
        try:
            return max(default, int(float(value)))
        except Exception:
            return default

    source_count = 1
    for key in ('Kaynak Sayısı', 'Olay_Kaynak_Sayisi', 'source_count'):
        if key in row and str(row.get(key, '')).strip() not in ('', 'nan', 'None'):
            source_count = _int_value(row.get(key), 1)
            break

    verification = str(row.get('Doğrulama', row.get('verification', '')) or '')
    try:
        official = bool(_is_official_radar_row(row))
    except Exception:
        official = 'resm' in norm(verification)

    lookup = _v122_source_verification_lookup()
    candidates = []
    url = str(row.get('URL', row.get('url', '')) or '').strip()
    title = title_key(row.get('Başlık', row.get('title', '')))
    if url:
        candidates.append(lookup.get('U:' + url))
    if title:
        candidates.append(lookup.get('T:' + title))
    candidates = [x for x in candidates if x]
    if candidates:
        best = max(candidates, key=lambda x: x[0])
        source_count = max(source_count, int(best[0] or 1))
        if not verification:
            verification = str(best[1] or '')
        official = official or bool(best[2])

    return source_count, verification, official


def _v122_source_verification_badge(row):
    'Her panelde tek bakışta anlaşılır kaynak-teyit rozeti üretir.'
    source_count, verification, official = _v122_verification_payload(row)
    vrank = _verification_rank(verification)
    if official and source_count >= 2:
        return f'🏛️✅ Resmî + {source_count} kaynak'
    if official:
        return '🏛️ Resmî kaynak'
    if source_count >= 4:
        return f'✅ Güçlü teyit · {source_count} kaynak'
    if source_count >= 2:
        return f'🟢 Çoklu kaynak · {source_count}'
    if vrank >= 3:
        return '🟡 Güçlü tek kaynak'
    return '⚪ Tek kaynak'


def _v122_add_source_verification(data):
    if data is None or data.empty:
        return data
    out = data.copy()
    out['Kaynak Teyidi'] = out.apply(_v122_source_verification_badge, axis=1)
    return out


def _v122_is_global_row(row):
    group = str(row.get('Kaynak_Grubu', '') or '')
    mode = str(row.get('_mode', '') or '').lower()
    return group.startswith('🌍') or mode == 'global'


def _v122_is_turkey_linked_global(row):
    if not _v122_is_global_row(row):
        return False
    text = norm(
        f"{row.get('Başlık', '')} {row.get('İçerik_Özeti', '')} "
        f"{row.get('Kaynak', '')} {row.get('Kategori', '')}"
    )
    return any(term in text for term in _V122_TURKEY_GLOBAL_TERMS)


def _v122_unique_event_count(data):
    if data is None or data.empty:
        return 0
    if 'Olay_ID' in data.columns:
        s = data['Olay_ID'].fillna('').astype(str)
        nonempty = s[s.str.len() > 0]
        if not nonempty.empty:
            return int(nonempty.nunique())
    if 'Başlık' in data.columns:
        return int(data['Başlık'].fillna('').astype(str).map(title_key).nunique())
    return int(len(data))


def _v122_manager_metrics(df):
    'Yönetici kartları için mevcut taramadaki son 24 saati olay bazlı özetler.'
    if df is None or df.empty:
        return {
            'news': 0, 'events': 0, 'important': 0, 'high_risk': 0,
            'critical': 0, 'global_strategic': 0, 'turkey_global': 0
        }

    x = df.copy()
    x['Tarih_dt'] = pd.to_datetime(x.get('Tarih_dt'), utc=True, errors='coerce')
    now = pd.Timestamp.now(tz='UTC')
    cutoff = now - pd.Timedelta(hours=24)
    x24 = x[(x['Tarih_dt'].isna()) | (x['Tarih_dt'] >= cutoff)].copy()

    events = _v122_unique_event_count(x24)
    risk_series = x24.get('Risk_Durumu', pd.Series('', index=x24.index)).fillna('').astype(str)
    high = x24[risk_series.eq('Yüksek Risk')].copy()
    high_risk = _v122_unique_event_count(high)

    critical_ids = set()
    for idx, row in x24.iterrows():
        try:
            if critical_industrial_incident(row.get('Başlık', ''), row.get('İçerik_Özeti', '')):
                event_id = str(row.get('Olay_ID') or title_key(row.get('Başlık', '')) or idx)
                critical_ids.add(event_id)
        except Exception:
            pass

    global_df = x24[x24.apply(_v122_is_global_row, axis=1)].copy()
    turkey_global_df = x24[x24.apply(_v122_is_turkey_linked_global, axis=1)].copy()

    important = 0
    try:
        n_events = max(10, min(300, events or 10))
        value_tbl = _v52_event_value_table(x24, n=n_events)
        if not value_tbl.empty:
            scores = pd.to_numeric(value_tbl['Değer_Skoru'], errors='coerce').fillna(0)
            important = int((scores >= 55).sum())
    except Exception:
        important = high_risk

    return {
        'news': int(len(x24)),
        'events': int(events),
        'important': int(important),
        'high_risk': int(high_risk),
        'critical': int(len(critical_ids)),
        'global_strategic': int(_v122_unique_event_count(global_df)),
        'turkey_global': int(_v122_unique_event_count(turkey_global_df)),
    }


def _v122_render_manager_summary(df):
    metrics = _v122_manager_metrics(df)
    st.markdown('## 🧭 Yönetici Özeti')
    st.caption('Son 24 saatin yönetici bakışı: hacim, tekil olay, önem, risk ve global stratejik görünüm.')
    st.markdown(
        '''
        <style>
        .stb-manager-card {
            border: 1px solid rgba(96,165,250,.22);
            border-radius: 18px;
            padding: 18px 18px 15px 18px;
            min-height: 132px;
            background: linear-gradient(135deg, rgba(7,18,43,.84), rgba(15,46,88,.72));
            box-shadow: 0 12px 30px rgba(0,0,0,.22);
            transition: transform .18s ease, box-shadow .18s ease;
            margin-bottom: 10px;
        }
        .stb-manager-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 12px 28px rgba(15,23,42,.13);
        }
        .stb-manager-icon {font-size: 25px; line-height: 1; margin-bottom: 10px;}
        .stb-manager-value {font-size: 34px; font-weight: 800; line-height: 1.05; letter-spacing: -0.8px;}
        .stb-manager-label {font-size: 14px; font-weight: 700; margin-top: 7px;}
        .stb-manager-sub {font-size: 11px; opacity: .72; margin-top: 6px; line-height: 1.25;}
        </style>
        ''',
        unsafe_allow_html=True,
    )

    cards = [
        ('📰', metrics['news'], 'Son 24 Saat Haber', 'Tarama havuzundaki güncel haber hacmi'),
        ('🧩', metrics['events'], 'Tekil Olay', 'Aynı gelişmenin tekrarları tekilleştirilmiştir'),
        ('⭐', metrics['important'], 'Önemli Gelişme', 'Değer skoru 55 ve üzerindeki tekil gelişmeler'),
        ('🚨', metrics['high_risk'], 'Yüksek Risk', 'Yüksek risk sınıfındaki tekil olaylar'),
        ('🏭', metrics['critical'], 'Kritik Sanayi Olayı', 'Yangın / patlama gibi kritik endüstriyel olaylar'),
        ('🌍', metrics['global_strategic'], 'Global Stratejik Gelişme', f"Türkiye bağlantılı: {metrics['turkey_global']}"),
    ]

    for start in (0, 3):
        cols = st.columns(3)
        for col, card in zip(cols, cards[start:start + 3]):
            icon, value, label, sub = card
            with col:
                card_html = (
                    '<div class="stb-manager-card">'
                    f'<div class="stb-manager-icon">{icon}</div>'
                    f'<div class="stb-manager-value">{value}</div>'
                    f'<div class="stb-manager-label">{label}</div>'
                    f'<div class="stb-manager-sub">{sub}</div>'
                    '</div>'
                )
                st.markdown(card_html, unsafe_allow_html=True)

    st.caption(
        'Not: Yönetici Özeti mevcut tarama verisinin son 24 saatlik bölümünü kullanır. '
        '“Önemli Gelişme” mevcut değer skoru modelinde 55+; global stratejik gelişme ise '
        'global tarama kaynaklarındaki tekil olay sayısıdır.'
    )


# ============================================================
# V123 — TEK HARİTA / MOD SEÇİMLİ STRATEJİK COĞRAFİ GÖRÜNÜM
# V122 kararlı çekirdeği korunur. Harita yalnızca mevcut tarama verisini kullanır;
# ek ağ isteği ve coğrafi servis çağrısı yapmaz.
# ============================================================

_V123_COUNTRY_GEO = {
    'Türkiye': (39.0, 35.0, ('turkey', 'türkiye', 'turkiye')),
    'ABD': (38.0, -97.0, ('united states', 'u.s.', 'u.s.a', 'usa', 'american')),
    'Kanada': (56.1, -106.3, ('canada', 'canadian')),
    'Meksika': (23.6, -102.6, ('mexico', 'mexican')),
    'Brezilya': (-14.2, -51.9, ('brazil', 'brazilian')),
    'Arjantin': (-38.4, -63.6, ('argentina', 'argentine')),
    'Birleşik Krallık': (55.4, -3.4, ('united kingdom', 'britain', 'british', 'england', 'uk')),
    'İrlanda': (53.1, -8.2, ('ireland', 'irish')),
    'Fransa': (46.2, 2.2, ('france', 'french')),
    'Almanya': (51.2, 10.5, ('germany', 'german')),
    'İtalya': (41.9, 12.6, ('italy', 'italian')),
    'İspanya': (40.5, -3.7, ('spain', 'spanish')),
    'Portekiz': (39.4, -8.2, ('portugal', 'portuguese')),
    'Hollanda': (52.1, 5.3, ('netherlands', 'dutch')),
    'Belçika': (50.5, 4.5, ('belgium', 'belgian')),
    'İsviçre': (46.8, 8.2, ('switzerland', 'swiss')),
    'Avusturya': (47.5, 14.6, ('austria', 'austrian')),
    'Polonya': (51.9, 19.1, ('poland', 'polish')),
    'Çekya': (49.8, 15.5, ('czech republic', 'czechia', 'czech')),
    'Slovakya': (48.7, 19.7, ('slovakia', 'slovak')),
    'Macaristan': (47.2, 19.5, ('hungary', 'hungarian')),
    'Romanya': (45.9, 24.9, ('romania', 'romanian')),
    'Bulgaristan': (42.7, 25.5, ('bulgaria', 'bulgarian')),
    'Yunanistan': (39.1, 21.8, ('greece', 'greek', 'yunanistan')),
    'Ukrayna': (48.4, 31.2, ('ukraine', 'ukrainian')),
    'Rusya': (61.5, 105.3, ('russia', 'russian')),
    'İsveç': (60.1, 18.6, ('sweden', 'swedish')),
    'Norveç': (60.5, 8.5, ('norway', 'norwegian')),
    'Finlandiya': (61.9, 25.7, ('finland', 'finnish')),
    'Danimarka': (56.3, 9.5, ('denmark', 'danish')),
    'Estonya': (58.6, 25.0, ('estonia', 'estonian')),
    'Letonya': (56.9, 24.6, ('latvia', 'latvian')),
    'Litvanya': (55.2, 23.9, ('lithuania', 'lithuanian')),
    'İsrail': (31.0, 34.9, ('israel', 'israeli')),
    'Suudi Arabistan': (23.9, 45.1, ('saudi arabia', 'saudi')),
    'BAE': (23.4, 53.8, ('united arab emirates', 'uae', 'emirates')),
    'Katar': (25.3, 51.2, ('qatar', 'qatari')),
    'İran': (32.4, 53.7, ('iran', 'iranian')),
    'Irak': (33.2, 43.7, ('iraq', 'iraqi')),
    'Suriye': (34.8, 38.9, ('syria', 'syrian')),
    'Mısır': (26.8, 30.8, ('egypt', 'egyptian')),
    'Güney Afrika': (-30.6, 22.9, ('south africa', 'south african')),
    'Hindistan': (20.6, 79.0, ('india', 'indian')),
    'Pakistan': (30.4, 69.3, ('pakistan', 'pakistani')),
    'Bangladeş': (23.7, 90.4, ('bangladesh', 'bangladeshi')),
    'Çin': (35.9, 104.2, ('china', 'chinese')),
    'Japonya': (36.2, 138.3, ('japan', 'japanese')),
    'Güney Kore': (36.5, 127.9, ('south korea', 'korea', 'korean')),
    'Tayvan': (23.7, 121.0, ('taiwan', 'taiwanese')),
    'Singapur': (1.35, 103.82, ('singapore', 'singaporean')),
    'Malezya': (4.2, 101.98, ('malaysia', 'malaysian')),
    'Endonezya': (-0.8, 113.9, ('indonesia', 'indonesian')),
    'Vietnam': (14.1, 108.3, ('vietnam', 'vietnamese')),
    'Tayland': (15.9, 100.99, ('thailand', 'thai')),
    'Filipinler': (12.9, 121.8, ('philippines', 'filipino')),
    'Avustralya': (-25.3, 133.8, ('australia', 'australian')),
    'Yeni Zelanda': (-40.9, 174.9, ('new zealand',))
}

_V123_TR_CITY_GEO = {
    'Adana': (37.00, 35.32, ('adana',)),
    'Ankara': (39.93, 32.86, ('ankara',)),
    'Antalya': (36.89, 30.70, ('antalya',)),
    'Balıkesir': (39.65, 27.88, ('balıkesir', 'balikesir')),
    'Bilecik': (40.14, 29.98, ('bilecik',)),
    'Bolu': (40.74, 31.61, ('bolu',)),
    'Bursa': (40.20, 29.06, ('bursa',)),
    'Çanakkale': (40.15, 26.41, ('çanakkale', 'canakkale')),
    'Çorum': (40.55, 34.95, ('çorum', 'corum')),
    'Denizli': (37.78, 29.09, ('denizli',)),
    'Diyarbakır': (37.91, 40.24, ('diyarbakır', 'diyarbakir')),
    'Düzce': (40.84, 31.16, ('düzce', 'duzce')),
    'Elazığ': (38.68, 39.23, ('elazığ', 'elazig')),
    'Erzurum': (39.90, 41.27, ('erzurum',)),
    'Eskişehir': (39.77, 30.52, ('eskişehir', 'eskisehir')),
    'Gaziantep': (37.07, 37.38, ('gaziantep',)),
    'Hatay': (36.20, 36.16, ('hatay', 'iskenderun', 'antakya')),
    'İstanbul': (41.01, 28.98, ('istanbul', 'İstanbul')),
    'İzmir': (38.42, 27.14, ('izmir',)),
    'Kahramanmaraş': (37.58, 36.93, ('kahramanmaraş', 'kahramanmaras')),
    'Karabük': (41.20, 32.63, ('karabük', 'karabuk')),
    'Kayseri': (38.72, 35.49, ('kayseri',)),
    'Kırıkkale': (39.85, 33.52, ('kırıkkale', 'kirikkale')),
    'Kırklareli': (41.73, 27.23, ('kırklareli', 'kirklareli')),
    'Kocaeli': (40.77, 29.94, ('kocaeli', 'izmit', 'gebze', 'dilovası', 'dilovasi')),
    'Konya': (37.87, 32.49, ('konya',)),
    'Kütahya': (39.42, 29.98, ('kütahya', 'kutahya')),
    'Malatya': (38.35, 38.31, ('malatya',)),
    'Manisa': (38.62, 27.43, ('manisa',)),
    'Mersin': (36.81, 34.64, ('mersin', 'tarsus')),
    'Sakarya': (40.78, 30.40, ('sakarya', 'adapazarı', 'adapazari')),
    'Samsun': (41.29, 36.33, ('samsun',)),
    'Şanlıurfa': (37.17, 38.79, ('şanlıurfa', 'sanliurfa')),
    'Tekirdağ': (40.98, 27.51, ('tekirdağ', 'tekirdag', 'çerkezköy', 'cerkezkoy', 'çorlu', 'corlu')),
    'Trabzon': (41.00, 39.72, ('trabzon',)),
    'Zonguldak': (41.46, 31.80, ('zonguldak', 'ereğli', 'eregli'))
}


def _v123_alias_hits(text, aliases):
    """Kelime sınırlarını mümkün olduğunca koruyarak konum adı eşleşmesi sayar."""
    raw = str(text or '').lower()
    total = 0
    for alias in aliases:
        a = str(alias or '').lower().strip()
        if not a:
            continue
        if len(a) <= 3:
            total += len(re.findall(r'(?<!\w)' + re.escape(a) + r'(?!\w)', raw, flags=re.I))
        else:
            total += raw.count(a)
    return total


def _v123_subject_location(row, mode):
    """
    Kaynak kuruluşun merkezini değil, haber metninde açıkça geçen konu coğrafyasını döndürür.
    Belirsiz haberler haritaya zorla yerleştirilmez.
    """
    title = str(row.get('Başlık', '') or '')
    summary = str(row.get('İçerik_Özeti', '') or '')
    title_n = norm(title)
    summary_n = norm(summary)

    if mode == '🚨 Kritik Sanayi Olayları':
        best = None
        best_score = 0
        for city, (lat, lon, aliases) in _V123_TR_CITY_GEO.items():
            score = 5 * _v123_alias_hits(title_n, aliases) + _v123_alias_hits(summary_n, aliases)
            if score > best_score:
                best = (city, 'Türkiye', lat, lon)
                best_score = score
        return best

    best = None
    best_score = 0
    for country, (lat, lon, aliases) in _V123_COUNTRY_GEO.items():
        score = 5 * _v123_alias_hits(title_n, aliases) + _v123_alias_hits(summary_n, aliases)
        # Türkiye bağlantılı global görünümde yalnız "Turkey" kelimesi geçti diye harita
        # otomatik olarak Türkiye'ye yığılmasın; başka ülke açıkça geçiyorsa onu öne çıkar.
        if mode == '🇹🇷 Türkiye Bağlantılı Global' and country == 'Türkiye':
            score *= 0.65
        if score > best_score:
            best = (country, country, lat, lon)
            best_score = score
    return best


def _v123_turkey_link_reason(row):
    text = norm(
        f"{row.get('Başlık', '')} {row.get('İçerik_Özeti', '')} "
        f"{row.get('Kaynak', '')} {row.get('Kategori', '')}"
    )
    labels = [
        ('ASELSAN', ('aselsan',)), ('TUSAŞ', ('tusaş', 'tusas', 'turkish aerospace')),
        ('ROKETSAN', ('roketsan',)), ('HAVELSAN', ('havelsan',)),
        ('Baykar/Bayraktar', ('baykar', 'bayraktar')), ('TOGG', ('togg',)),
        ('KAAN', ('kaan',)), ('HİSAR/SİPER', ('hisar', 'siper')),
        ('Türkiye', ('turkey', 'türkiye', 'turkiye', 'turkish', 'ankara', 'istanbul')),
    ]
    found = [label for label, terms in labels if any(term in text for term in terms)]
    return ' · '.join(found[:3]) if found else 'Türkiye ile doğrudan bağlantı'


def _v123_map_event_key(row):
    event_id = str(row.get('Olay_ID', '') or '').strip()
    if event_id:
        return 'E:' + event_id
    return 'T:' + title_key(row.get('Başlık', ''))


def _v123_map_dataset(df, mode):
    """Seçilen harita modu için tekil ve haritalanabilir olay veri setini hazırlar."""
    if df is None or df.empty:
        return pd.DataFrame(), 0

    cache = st.session_state.setdefault('_v123_map_cache', {})
    scan_id = st.session_state.get('current_scan_id')
    cache_key = (str(scan_id), str(st.session_state.get('scan_time')), int(len(df)), str(mode))
    cached = cache.get(cache_key)
    if cached is not None:
        return pd.DataFrame(cached.get('records', [])), int(cached.get('unmapped', 0))

    if mode == '🌍 Global Stratejik Gelişmeler':
        subset = df[df.apply(_v122_is_global_row, axis=1)].copy()
    elif mode == '🇹🇷 Türkiye Bağlantılı Global':
        subset = df[df.apply(_v122_is_turkey_linked_global, axis=1)].copy()
    else:
        mask = df.apply(
            lambda r: bool(critical_industrial_incident(r.get('Başlık', ''), r.get('İçerik_Özeti', ''))),
            axis=1,
        )
        subset = df[mask].copy()

    if subset.empty:
        return pd.DataFrame(), 0

    subset['_v123_event_key'] = subset.apply(_v123_map_event_key, axis=1)
    # Aynı olay birçok kaynaktan geldiyse, haritada tek nokta gösterilir. Önce en yüksek
    # kaynak sayısı/risk, sonra en yeni kayıt tercih edilir.
    if 'Tarih_dt' in subset.columns:
        subset['Tarih_dt'] = pd.to_datetime(subset['Tarih_dt'], utc=True, errors='coerce')
    else:
        subset['Tarih_dt'] = pd.NaT
    subset['_v123_sources'] = subset.apply(lambda r: _v122_verification_payload(r)[0], axis=1)
    subset['_v123_risk'] = pd.to_numeric(subset.get('Risk_Skoru', 0), errors='coerce').fillna(0)
    subset = subset.sort_values(
        ['_v123_sources', '_v123_risk', 'Tarih_dt'],
        ascending=[False, False, False],
        na_position='last',
    ).drop_duplicates('_v123_event_key', keep='first')

    rows = []
    unmapped = 0
    for _, row in subset.iterrows():
        loc = _v123_subject_location(row, mode)
        if not loc:
            unmapped += 1
            continue
        location, country, lat, lon = loc
        source_count, verification, official = _v122_verification_payload(row)
        badge = _v122_source_verification_badge(row)
        risk = int(float(row.get('Risk_Skoru', 0) or 0))
        title = _clean_note_text(row.get('Başlık', ''))
        summary = _clean_note_text(row.get('İçerik_Özeti', ''))
        category = _clean_note_text(row.get('Kategori', '')) or 'Diğer'
        source = _clean_note_text(row.get('Kaynak', '')) or 'Açık Kaynak'
        time_text = _clean_note_text(row.get('Tarih', ''))
        if not time_text and row.get('Tarih_dt') is not None:
            time_text = fmt_dt(row.get('Tarih_dt'))
        turkey_link = _v123_turkey_link_reason(row) if _v122_is_turkey_linked_global(row) else '—'
        rows.append({
            'Konum': location,
            'Ülke': country,
            'lat': float(lat),
            'lon': float(lon),
            'Başlık': title,
            'Kısa Başlık': title[:92] + ('…' if len(title) > 92 else ''),
            'Özet': summary[:520] + ('…' if len(summary) > 520 else ''),
            'Kategori': category,
            'Kaynak': source,
            'Kaynak Teyidi': badge,
            'Kaynak Sayısı': int(source_count),
            'Doğrulama': verification,
            'Risk': risk,
            'Türkiye Bağlantısı': turkey_link,
            'Tarih': time_text,
            'URL': str(row.get('URL', '') or ''),
            'Boyut': max(9, min(28, 10 + risk * 0.10 + min(source_count, 5) * 1.5)),
        })

    result = pd.DataFrame(rows)
    if len(cache) > 8:
        cache.clear()
    cache[cache_key] = {'records': rows, 'unmapped': unmapped}
    return result, unmapped


def _v123_render_map_detail(row):
    if row is None:
        return
    st.markdown('#### 🔎 Seçili Gelişme')
    st.markdown(f"**{row.get('Başlık', '')}**")
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f"**Konum**  \n{row.get('Konum', '—')}")
    c2.markdown(f"**Kategori**  \n{row.get('Kategori', '—')}")
    c3.markdown(f"**Risk**  \n{row.get('Risk', 0)}/100")
    c4.markdown(f"**Teyit**  \n{row.get('Kaynak Teyidi', '—')}")
    st.caption(
        f"Kaynak: {row.get('Kaynak', '—')} · Tarih/Saat: {row.get('Tarih', '—')} · "
        f"Türkiye bağlantısı: {row.get('Türkiye Bağlantısı', '—')}"
    )
    if row.get('Özet'):
        st.write(row.get('Özet'))
    url = str(row.get('URL', '') or '').strip()
    if url.startswith(('http://', 'https://')):
        st.markdown(f'[🔗 Haberi aç]({url})')


def _v123_render_strategic_map(df):
    """Yönetici Özeti altında tek harita; mod değiştikçe veri katmanı değişir."""
    st.markdown('## 🌍 Küresel Sanayi ve Stratejik Teknoloji Haritası')
    st.caption(
        'Harita mevcut tarama dönemiyle otomatik senkronizedir. Konum, yayıncının merkezine göre değil; '
        'haber başlığı/özetinde açıkça geçen olay coğrafyasına göre belirlenir.'
    )
    mode = st.radio(
        'Harita modu',
        ['🌍 Global Stratejik Gelişmeler', '🇹🇷 Türkiye Bağlantılı Global', '🚨 Kritik Sanayi Olayları'],
        index=1,
        horizontal=True,
        key='v123_strategic_map_mode',
    )
    data, unmapped = _v123_map_dataset(df, mode)

    if data.empty:
        st.info(
            'Bu modda açık coğrafi konum içeren gelişme bulunamadı. '
            'Belirsiz konumlar yanlış nokta oluşturmamak için haritaya eklenmez.'
        )
        return

    categories = sorted(x for x in data['Kategori'].dropna().astype(str).unique() if x.strip())
    selected_categories = st.multiselect(
        'Kategori filtresi',
        categories,
        default=categories,
        key=f"v123_map_categories_{re.sub(r'[^a-zA-Z0-9]+', '_', mode)}",
        help='Harita tek kalır; seçilen kategorilere göre noktalar anlık filtrelenir.',
    )
    if selected_categories:
        data = data[data['Kategori'].astype(str).isin(selected_categories)].copy()
    else:
        data = data.iloc[0:0].copy()
    if data.empty:
        st.info('Seçilen kategori filtresinde haritalanabilir gelişme bulunmuyor.')
        return

    location_count = int(data['Konum'].nunique())
    event_count = int(len(data))
    turkey_count = int((data['Türkiye Bağlantısı'].astype(str) != '—').sum())
    high_count = int((pd.to_numeric(data['Risk'], errors='coerce').fillna(0) >= 70).sum())
    m1, m2, m3, m4 = st.columns(4)
    m1.metric('Haritalanan Konum', location_count)
    m2.metric('Tekil Gelişme', event_count)
    m3.metric('Türkiye Bağlantılı', turkey_count)
    m4.metric('Yüksek Risk', high_count)

    clicked_index = None
    try:
        import plotly.express as px

        fig = px.scatter_geo(
            data,
            lat='lat',
            lon='lon',
            color='Kategori',
            size='Boyut',
            hover_name='Kısa Başlık',
            hover_data={
                'Konum': True,
                'Kaynak': True,
                'Kaynak Teyidi': True,
                'Risk': True,
                'Türkiye Bağlantısı': True,
                'Tarih': True,
                'lat': False,
                'lon': False,
                'Boyut': False,
                'Kategori': False,
            },
            projection='natural earth',
            height=560,
            custom_data=['Başlık', 'Konum'],
        )
        if mode == '🚨 Kritik Sanayi Olayları':
            fig.update_geos(fitbounds='locations', visible=True, showcountries=True, showcoastlines=True)
        else:
            fig.update_geos(showcountries=True, showcoastlines=True, showland=True)
        fig.update_layout(
            margin=dict(l=0, r=0, t=8, b=0),
            legend_title_text='Kategori',
        )

        # Yeni Streamlit sürümlerinde nokta seçimi destekleniyorsa tıklanan gelişmeyi kartta aç.
        try:
            event = st.plotly_chart(
                fig,
                use_container_width=True,
                key='v123_strategic_geo_chart',
                on_select='rerun',
                selection_mode='points',
            )
            selection = getattr(event, 'selection', None)
            points = getattr(selection, 'points', None) if selection is not None else None
            if points:
                point = points[0]
                if isinstance(point, dict):
                    custom = point.get('customdata')
                    if custom is not None and len(custom) >= 2:
                        hit = data[
                            (data['Başlık'].astype(str) == str(custom[0]))
                            & (data['Konum'].astype(str) == str(custom[1]))
                        ]
                        if not hit.empty:
                            clicked_index = data.reset_index(drop=True).index[
                                data.reset_index(drop=True)['Başlık'].astype(str).eq(str(custom[0]))
                                & data.reset_index(drop=True)['Konum'].astype(str).eq(str(custom[1]))
                            ][0]
                    if clicked_index is None:
                        clicked_index = point.get('point_index', point.get('pointNumber'))
                else:
                    clicked_index = getattr(point, 'point_index', None)
        except TypeError:
            st.plotly_chart(fig, use_container_width=True, key='v123_strategic_geo_chart_fallback')
    except Exception:
        # Plotly kurulu olmayan ortamlarda Streamlit'in temel haritasına geri dön.
        st.map(data[['lat', 'lon']].rename(columns={'lat': 'latitude', 'lon': 'longitude'}))
        st.caption('Gelişmiş harita bileşeni kullanılamadığı için temel konum görünümü gösterilmektedir.')

    labels = [
        f"{row['Konum']} · {row['Kısa Başlık']}"
        for _, row in data.reset_index(drop=True).iterrows()
    ]
    default_index = 0
    try:
        if clicked_index is not None and 0 <= int(clicked_index) < len(labels):
            default_index = int(clicked_index)
    except Exception:
        default_index = 0

    st.caption('Haritada bir noktayı seçebilir veya aşağıdaki listeden gelişmeyi açabilirsiniz.')
    selected_label = st.selectbox(
        'Haritadaki gelişme',
        labels,
        index=default_index,
        key=f"v123_map_event_{re.sub(r'[^a-zA-Z0-9]+', '_', mode)}",
    )
    try:
        selected_idx = labels.index(selected_label)
    except ValueError:
        selected_idx = 0
    _v123_render_map_detail(data.reset_index(drop=True).iloc[selected_idx].to_dict())

    if unmapped:
        st.caption(
            f'ℹ️ {unmapped} tekil gelişmede güvenilir şehir/ülke ifadesi bulunmadığı için haritaya nokta eklenmedi.'
        )
    st.caption(
        'Konum çıkarımı yalnız açık metin eşleşmesine dayanır; tahminî geocoding yapılmaz. '
        'Bu yaklaşım sunumda yanlış ülke/şehir göstermeyi öncelikli olarak engeller.'
    )

# ============================================================
# /V123 STRATEJİK HARİTA
# ============================================================


# ============================================================
# V124 — HARİTA NOKTA SEÇİMİ / DETAY SENKRONU
# Plotly seçim olayı ile alttaki selectbox/session_state senkronize edilir.
# V123 harita görünümü ve veri mantığı korunur.
# ============================================================

def _v124_event_value(obj, key, default=None):
    """Streamlit PlotlyState hem dict-benzeri hem attribute-benzeri olabilir."""
    if obj is None:
        return default
    try:
        if isinstance(obj, dict):
            return obj.get(key, default)
    except Exception:
        pass
    try:
        return getattr(obj, key, default)
    except Exception:
        return default


def _v124_selected_plotly_point(event):
    """st.plotly_chart seçim sonucundan son seçilen noktayı güvenli biçimde döndürür."""
    selection = _v124_event_value(event, 'selection')
    points = _v124_event_value(selection, 'points', [])
    try:
        points = list(points or [])
    except Exception:
        points = []
    return points[-1] if points else None


def _v124_map_mode_key(mode):
    return re.sub(r'[^a-zA-Z0-9]+', '_', str(mode or '')).strip('_')


def _v124_map_row_id(row):
    """Aynı başlık/konum tekrar etse bile seçim için kararlı benzersiz kimlik."""
    raw = '|'.join([
        str(row.get('URL', '') or ''),
        str(row.get('Başlık', '') or ''),
        str(row.get('Konum', '') or ''),
        str(row.get('Tarih', '') or ''),
        str(row.get('Kaynak', '') or ''),
    ])
    return hashlib.sha1(raw.encode('utf-8', errors='ignore')).hexdigest()[:18]


_V125_MAP_CATEGORY_COLORS = {
    'Savunma & Havacılık': '#ff5d73',
    'Dijital & Yapay Zeka': '#8b8cff',
    'Yarı İletken & Elektronik': '#fb923c',
    'Otomotiv & Mobilite': '#38bdf8',
    'Enerji': '#22d3ee',
    'Sanayi & Üretim': '#34d399',
    'Uzay & İleri Teknoloji': '#f6c85f',
    'Kurumsal Ekosistem': '#c084fc',
    'Diğer': '#94a3b8',
}

def _v123_render_strategic_map(df):
    """V124 — tek harita; nokta tıklaması detay kartını doğrudan günceller."""
    st.markdown('## 🌍 Küresel Sanayi ve Stratejik Teknoloji Haritası')
    st.caption(
        'Harita mevcut tarama dönemiyle otomatik senkronizedir. Konum, yayıncının merkezine göre değil; '
        'haber başlığı/özetinde açıkça geçen olay coğrafyasına göre belirlenir.'
    )
    mode = st.radio(
        'Harita modu',
        ['🌍 Global Stratejik Gelişmeler', '🇹🇷 Türkiye Bağlantılı Global', '🚨 Kritik Sanayi Olayları'],
        index=1,
        horizontal=True,
        key='v123_strategic_map_mode',
    )
    data, unmapped = _v123_map_dataset(df, mode)

    if data.empty:
        st.info(
            'Bu modda açık coğrafi konum içeren gelişme bulunamadı. '
            'Belirsiz konumlar yanlış nokta oluşturmamak için haritaya eklenmez.'
        )
        return

    categories = sorted(x for x in data['Kategori'].dropna().astype(str).unique() if x.strip())
    selected_categories = st.multiselect(
        'Kategori filtresi',
        categories,
        default=categories,
        key=f"v123_map_categories_{_v124_map_mode_key(mode)}",
        help='Harita tek kalır; seçilen kategorilere göre noktalar anlık filtrelenir.',
    )
    if selected_categories:
        data = data[data['Kategori'].astype(str).isin(selected_categories)].copy()
    else:
        data = data.iloc[0:0].copy()
    if data.empty:
        st.info('Seçilen kategori filtresinde haritalanabilir gelişme bulunmuyor.')
        return

    data = data.reset_index(drop=True)
    data['_Map_ID'] = data.apply(_v124_map_row_id, axis=1)

    location_count = int(data['Konum'].nunique())
    event_count = int(len(data))
    turkey_count = int((data['Türkiye Bağlantısı'].astype(str) != '—').sum())
    high_count = int((pd.to_numeric(data['Risk'], errors='coerce').fillna(0) >= 70).sum())
    m1, m2, m3, m4 = st.columns(4)
    m1.metric('Haritalanan Konum', location_count)
    m2.metric('Tekil Gelişme', event_count)
    m3.metric('Türkiye Bağlantılı', turkey_count)
    m4.metric('Yüksek Risk', high_count)

    selected_map_id = None
    try:
        import plotly.express as px

        fig = px.scatter_geo(
            data,
            lat='lat',
            lon='lon',
            color='Kategori',
            size='Boyut',
            hover_name='Kısa Başlık',
            hover_data={
                'Konum': True,
                'Kaynak': True,
                'Kaynak Teyidi': True,
                'Risk': True,
                'Türkiye Bağlantısı': True,
                'Tarih': True,
                'lat': False,
                'lon': False,
                'Boyut': False,
                'Kategori': False,
                '_Map_ID': False,
            },
            projection='natural earth',
            height=560,
            color_discrete_map=_V125_MAP_CATEGORY_COLORS,
            # İlk alan benzersiz kimliktir. Kategori ayrı trace oluştursa bile
            # pointNumber'a güvenmek zorunda kalmayız.
            custom_data=['_Map_ID', 'Başlık', 'Konum'],
        )
        # V125 — yalnız görsel katman: koyu kurumsal / operasyon merkezi haritası.
        geo_style = dict(
            bgcolor='rgba(0,0,0,0)',
            showland=True,
            landcolor='#0d1a2e',
            showocean=True,
            oceancolor='#050b16',
            showlakes=True,
            lakecolor='#071321',
            showcountries=True,
            countrycolor='rgba(160,190,230,.26)',
            countrywidth=.7,
            showcoastlines=True,
            coastlinecolor='rgba(130,180,235,.34)',
            coastlinewidth=.8,
            showframe=False,
            showrivers=False,
            resolution=110,
        )
        if mode == '🚨 Kritik Sanayi Olayları':
            fig.update_geos(fitbounds='locations', visible=True, **geo_style)
        else:
            fig.update_geos(**geo_style)

        fig.update_traces(
            marker=dict(
                opacity=.92,
                line=dict(width=1.25, color='rgba(238,248,255,.72)'),
            ),
            selector=dict(type='scattergeo'),
        )
        fig.update_layout(
            margin=dict(l=0, r=0, t=8, b=0),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            font=dict(color='#eaf2ff', family='Arial, sans-serif'),
            legend_title_text='Kategori',
            legend=dict(
                bgcolor='rgba(5,12,28,.80)',
                bordercolor='rgba(96,165,250,.24)',
                borderwidth=1,
                font=dict(color='#eaf2ff', size=11),
                title=dict(font=dict(color='#bfdbfe', size=12)),
                x=1.01,
                y=1.0,
            ),
            hoverlabel=dict(
                bgcolor='#071426',
                bordercolor='#38bdf8',
                font=dict(color='#f8fbff', size=12, family='Arial, sans-serif'),
                namelength=0,
            ),
            clickmode='event+select',
            selectionrevision='v125',
        )

        try:
            event = st.plotly_chart(
                fig,
                use_container_width=True,
                key='v124_strategic_geo_chart',
                on_select='rerun',
                selection_mode='points',
            )
            point = _v124_selected_plotly_point(event)
            if point is not None:
                custom = _v124_event_value(point, 'customdata')
                try:
                    custom = list(custom or [])
                except Exception:
                    custom = []
                if custom:
                    selected_map_id = str(custom[0] or '')

                # Bazı Streamlit/Plotly sürümleri customdata döndürmez.
                # Bu durumda curve/point bilgisi yerine başlık/konum alanlarını dene.
                if not selected_map_id:
                    hover_text = _v124_event_value(point, 'hovertext')
                    if hover_text:
                        hit = data[data['Kısa Başlık'].astype(str) == str(hover_text)]
                        if len(hit) == 1:
                            selected_map_id = str(hit.iloc[0]['_Map_ID'])
        except TypeError:
            st.plotly_chart(fig, use_container_width=True, key='v124_strategic_geo_chart_fallback')
    except Exception:
        st.map(data[['lat', 'lon']].rename(columns={'lat': 'latitude', 'lon': 'longitude'}))
        st.caption('Gelişmiş harita bileşeni kullanılamadığı için temel konum görünümü gösterilmektedir.')

    labels = [
        f"{row['Konum']} · {row['Kısa Başlık']}"
        for _, row in data.iterrows()
    ]
    mode_key = _v124_map_mode_key(mode)
    select_key = f'v123_map_event_{mode_key}'
    click_state_key = f'_v124_last_map_click_{mode_key}'
    active_id_key = f'_v124_active_map_id_{mode_key}'

    # HARİTA -> SESSION STATE -> SELECTBOX
    # Widget daha önce oluşturulmuş olsa bile index parametresine güvenmeyiz;
    # seçilen noktanın etiketini widget state'e yazıyoruz.
    if selected_map_id and selected_map_id in set(data['_Map_ID'].astype(str)):
        old_click = str(st.session_state.get(click_state_key, '') or '')
        if selected_map_id != old_click:
            selected_idx = int(data.index[data['_Map_ID'].astype(str) == selected_map_id][0])
            selected_label_from_map = labels[selected_idx]
            st.session_state[click_state_key] = selected_map_id
            st.session_state[active_id_key] = selected_map_id
            st.session_state[select_key] = selected_label_from_map

    # Mod/kategori değişince daha önceki seçim artık listede yoksa ilk satıra dön.
    if st.session_state.get(select_key) not in labels:
        st.session_state[select_key] = labels[0]
        st.session_state[active_id_key] = str(data.iloc[0]['_Map_ID'])

    st.caption('Haritada bir noktayı seçebilir veya aşağıdaki listeden gelişmeyi açabilirsiniz.')
    selected_label = st.selectbox(
        'Haritadaki gelişme',
        labels,
        key=select_key,
    )
    try:
        selected_idx = labels.index(selected_label)
    except ValueError:
        selected_idx = 0

    selected_row = data.iloc[selected_idx]
    selected_id = str(selected_row['_Map_ID'])
    st.session_state[active_id_key] = selected_id
    _v123_render_map_detail(selected_row.to_dict())

    if unmapped:
        st.caption(
            f'ℹ️ {unmapped} tekil gelişmede güvenilir şehir/ülke ifadesi bulunmadığı için haritaya nokta eklenmedi.'
        )
    st.caption(
        'Konum çıkarımı yalnız açık metin eşleşmesine dayanır; tahminî geocoding yapılmaz. '
        'Bu yaklaşım sunumda yanlış ülke/şehir göstermeyi öncelikli olarak engeller.'
    )

# ============================================================
# /V124 HARİTA NOKTA SEÇİMİ
# ============================================================


# ============================================================
# V122 — YÖNETİCİ ÖZETİ + TEYİT + TÜRKİYE BAĞLANTILI GLOBAL
# 1) Ana başlık: STB-Açık Kaynak Tarama Merkezi
# 2) Ana haber görünümüne ayrı "Global Sanayi / Teknoloji" bölümü eklendi.
# 3) V120 global kaynak/tarama mantığı aynen korunur.
# ============================================================

# -----------------------------
# UI
# -----------------------------
# V125 — Görsel tema: ekteki teknoloji arka planı + koyu stratejik harita; işlevsel mantık V124 ile aynıdır.
st.title('🛡️ STB-Açık Kaynak Tarama Merkezi')
st.caption('Hızlı ilk bakış · mod seçimli stratejik harita · olay kümeleri · risk/negatif ayrımı · Türk medya önceliği · hedefli global sanayi/teknoloji basını · Yunan/Türk savunma · kaynak güvenilirliği · trend · alarm · seçilen haberlerden DOCX')
with st.sidebar:
    st.header('⚙️ Tarama Ayarları')
    default=('sanayi OR teknoloji OR üretim OR imalat OR fabrika OR OSB OR makine OR otomasyon OR robotik OR Ar-Ge OR patent OR yapay zeka OR yazılım OR siber güvenlik OR çip OR yarı iletken OR elektronik OR telekom OR kuantum OR biyoteknoloji OR nanoteknoloji OR savunma sanayii OR ASELSAN OR TUSAŞ OR ROKETSAN OR HAVELSAN OR Baykar OR İHA OR SİHA OR KAAN OR havacılık OR uzay OR uydu OR otomotiv OR TOGG OR batarya OR enerji OR hidrojen OR kimya OR petrokimya OR demir çelik OR madencilik OR tekstil OR gıda teknolojisi OR tarım teknolojisi OR lojistik OR tedarik zinciri OR TÜBİTAK OR KOSGEB OR teknopark OR yatırım teşvik OR yerlileştirme')
    query=st.text_area('Geniş sanayi / teknoloji sorgusu:',default,height=190)
    watch=st.text_area('⭐ Takip listesi (virgül / satır sonu):','ASELSAN, TUSAŞ, ROKETSAN, HAVELSAN, Baykar, TOGG, TÜBİTAK',height=90)
    neg=st.checkbox('⚠️ Negatif haberleri ayrıca tespit et',True)
    greek=st.checkbox('🇬🇷 Yunan medyası — yalnızca Türk savunma sanayii',True)
    social=st.checkbox('📱 Türk açık sosyal / indeks kaynakları',True)
    global_on=st.checkbox(
        '🌍 Global sanayi / teknoloji basını',
        True,
        help='MIT Technology Review, IEEE Spectrum, Ars Technica, IndustryWeek, Automation World, Manufacturing Tomorrow, Financial Times, Bloomberg, Nikkei Asia, Defense News ve Aviation Week hedefli taranır.'
    )
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
    if global_on:
        batches.append((
            '🌍 Global sanayi / teknoloji basını',
            build_global_queries(when),
            'global'
        ))
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
            future_map={
                ex.submit(rss_global if mode=='global' else rss,q):(label,mode)
                for label,q,mode in jobs
            }
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
            _compare_since_previous(
                _v119_scan_df,st.session_state.get('current_scan_id')
            )
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
        _know_select=_v122_add_source_verification(_know_select)
        if 'Seç' not in _know_select.columns:
            _know_select.insert(0,'Seç',False)

        _edited_know=st.data_editor(
            _know_select[['Seç','Tarih','Başlık','Kaynak Teyidi','İçerik_Özeti','Değer_Skoru',
                          'Neden_Değerli','Kaynak_Sayısı','Risk_Skoru','URL']],
            column_config={
                'Seç':st.column_config.CheckboxColumn('Seç'),
                'Değer_Skoru':st.column_config.ProgressColumn('Değer Skoru',min_value=0,max_value=100,format='%d/100'),
                'Risk_Skoru':st.column_config.NumberColumn('Risk',format='%d/100'),
                'URL':st.column_config.LinkColumn('Haber Linki'),
                'İçerik_Özeti':st.column_config.TextColumn('Kısa İçerik',width='large'),
                'Kaynak Teyidi':st.column_config.TextColumn('Kaynak Teyidi',width='medium')
            },
            disabled=['Tarih','Başlık','Kaynak Teyidi','İçerik_Özeti','Değer_Skoru','Neden_Değerli',
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
st.caption('⚡ V119 performans modu: tarama sonrası panel özetleri önceden hesaplanır; seçim kutuları tek başına ağır analizleri yeniden çalıştırmaz.')
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

        _v122_render_manager_summary(df)

        _v123_render_strategic_map(df)

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
            ['📰 Kronolojik','⚠️ Negatif','🚨 Yüksek Risk','🇹🇷 Türk','🇬🇷 Yunan','🇹🇷🌍 Türkiye Bağlantılı Global','🌍 Global Sanayi / Teknoloji','🧩 Olaylar','📈 Trend / Analiz','⭐ Takip Listesi'],
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
            page_df=_v122_add_source_verification(page_df)
            chron_cols=[
                'Seç','Tarih','Kaynak_Grubu','Kaynak','Kaynak Teyidi','Kaynak Sayısı','Haber Sayısı',
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
                        'Durum':st.column_config.TextColumn('Durum',width='large'),
                        'Kaynak Teyidi':st.column_config.TextColumn('Kaynak Teyidi',width='medium')
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
                                _v122_add_source_verification(_event_sources)[['Tarih','Kaynak','Kaynak Teyidi','Başlık','İçerik_Özeti','URL']].head(20),
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

        elif view=='🇹🇷🌍 Türkiye Bağlantılı Global':
            turkey_global_df=df[df.apply(_v122_is_turkey_linked_global,axis=1)].copy()
            if turkey_global_df.empty:
                st.info(
                    'Seçilen zaman aralığında global kaynaklarda Türkiye ile doğrudan bağlantılı '
                    'sanayi / teknoloji gelişmesi bulunamadı.'
                )
            else:
                turkey_global_df=turkey_global_df.sort_values(
                    'Tarih_dt',ascending=False,na_position='last'
                )
                st.caption(
                    f'{len(turkey_global_df)} haber · {_v122_unique_event_count(turkey_global_df)} tekil gelişme. '
                    'Türkiye, Türk şirketleri, savunma/havacılık programları veya stratejik sanayi-teknoloji '
                    'başlıklarıyla doğrudan bağlantılı global yayınlar.'
                )
                _section_select_table(
                    'turkey_linked_global_view',
                    turkey_global_df,
                    [
                        'Tarih','Kaynak','Kategori','Başlık','İçerik_Özeti',
                        'Risk_Skoru','Duygu','Kaynak_Güvenilirliği','Doğrulama','URL'
                    ],
                    height=650
                )

        elif view=='🌍 Global Sanayi / Teknoloji':
            global_df=df[
                df.Kaynak_Grubu.astype(str).eq('🌍 Global Sanayi / Teknoloji')
            ].copy()
            if global_df.empty:
                st.info(
                    'Seçilen zaman aralığında hedefli global sanayi / teknoloji '
                    'kaynaklarından sonuç bulunamadı.'
                )
            else:
                global_df=global_df.sort_values(
                    'Tarih_dt',ascending=False,na_position='last'
                )
                st.caption(
                    f'{len(global_df)} global sanayi / teknoloji haberi · '
                    'MIT Technology Review, IEEE Spectrum, Ars Technica, IndustryWeek, '
                    'Automation World, Manufacturing Tomorrow, Financial Times, Bloomberg, '
                    'Nikkei Asia, Defense News ve Aviation Week önceliklidir.'
                )
                _section_select_table(
                    'global_industry_tech_view',
                    global_df,
                    [
                        'Tarih','Kaynak','Kategori','Başlık','İçerik_Özeti',
                        'Risk_Skoru','Duygu','Kaynak_Güvenilirliği','Doğrulama','URL'
                    ],
                    height=650
                )

        elif view=='🧩 Olaylar':
            ev=build_event_summary(df)
            ev=_v122_add_source_verification(ev)
            st.dataframe(
                ev,
                column_config={'Kaynak Teyidi':st.column_config.TextColumn('Kaynak Teyidi',width='medium')},
                hide_index=True,use_container_width=True,height=480
            )
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
            basket_view=_v122_add_source_verification(basket_view)

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
                        'Durum':st.column_config.TextColumn('Durum',width='large'),
                        'Kaynak Teyidi':st.column_config.TextColumn('Kaynak Teyidi',width='medium')
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
            osint_view=_v122_add_source_verification(osint_view)

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
                        'Durum':st.column_config.TextColumn('Durum',width='large'),
                        'Kaynak Teyidi':st.column_config.TextColumn('Kaynak Teyidi',width='medium')
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
            _pv=_v122_add_source_verification(_pv)
            _pv.insert(0,'Seç',False)
            with st.form('v81_presentation_basket_form',clear_on_submit=False):
                _ped=st.data_editor(_pv,column_config={
                    'Seç':st.column_config.CheckboxColumn('Seç'),
                    'url':st.column_config.LinkColumn('Haber Linki'),
                    'Durum':st.column_config.TextColumn('Durum',width='large'),
                    'Kaynak Teyidi':st.column_config.TextColumn('Kaynak Teyidi',width='medium')
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
