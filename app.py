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
_V125_BG_DATA = "data:image/jpeg;base64,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AAADNmp1bWIAAAApanVtZGMyYXMAEQAQgAAAqgA4m3EDYzJwYS5hc3NlcnRpb25zAAAAAJxqdW1iAAAAKGp1bWRjYm9yABEAEIAAAKoAOJtxA2MycGEuaGFzaC5kYXRhAAAAAGxjYm9ypGpleGNsdXNpb25zgaJlc3RhcnQUZmxlbmd0aBkYymNhbGdmc2hhMjU2ZGhhc2hYIGcE4au4KY60SYSc/MNGTW5IyqGxmoPLtbYGRxZO7Xl2Y3BhZE4AAAAAAAAAAAAAAAAAAAAAAfhqdW1iAAAAKWp1bWRjYm9yABEAEIAAAKoAOJtxA2MycGEuYWN0aW9ucy52MgAAAAHHY2JvcqFnYWN0aW9uc4KkZmFjdGlvbmxjMnBhLmNyZWF0ZWRrZGVzY3JpcHRpb254IENyZWF0ZWQgYnkgR29vZ2xlIEdlbmVyYXRpdmUgQUkucWRpZ2l0YWxTb3VyY2VUeXBleEZodHRwOi8vY3YuaXB0Yy5vcmcvbmV3c2NvZGVzL2RpZ2l0YWxzb3VyY2V0eXBlL3RyYWluZWRBbGdvcml0aG1pY01lZGlhanBhcmFtZXRlcnOha2luZ3JlZGllbnRzgaJjdXJseC1zZWxmI2p1bWJmPWMycGEuYXNzZXJ0aW9ucy9jMnBhLmluZ3JlZGllbnQudjNkaGFzaFggQ89BBSZ+7rzB0gQct1E1zC8AYWx4J+2IJ7n7zntYOZajZmFjdGlvbmtjMnBhLmVkaXRlZGtkZXNjcmlwdGlvbngoQXBwbGllZCBpbXBlcmNlcHRpYmxlIFN5bnRoSUQgd2F0ZXJtYXJrLnFkaWdpdGFsU291cmNlVHlwZXhGaHR0cDovL2N2LmlwdGMub3JnL25ld3Njb2Rlcy9kaWdpdGFsc291cmNldHlwZS90cmFpbmVkQWxnb3JpdGhtaWNNZWRpYQAAAHFqdW1iAAAALGp1bWRjYm9yABEAEIAAAKoAOJtxA2MycGEuaW5ncmVkaWVudC52MwAAAAA9Y2JvcqJscmVsYXRpb25zaGlwZ2lucHV0VG9rZGVzY3JpcHRpb25ySW5wdXQgaW5ncmVkaWVudCAwAAAecWp1bWIAAABHanVtZGMybWEAEQAQgAAAqgA4m3EDdXJuOmMycGE6Y2ZlMTg5MTctNGVhZi04Y2MzLTg0MDgtNGJmZDRkMDRlNzBl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AAAAKWp1bWRjMmFzABEAEIAAAKoAOJtxA2MycGEuYXNzZXJ0aW9ucwAAAACcanVtYgAAAChqdW1kY2JvcgARABCAAACqADibcQNjMnBhLmhhc2guZGF0YQAAAABsY2JvcqRqZXhjbHVzaW9uc4GiZXN0YXJ0GRukZmxlbmd0aBk3O2NhbGdmc2hhMjU2ZGhhc2hYIGmSACu7+Dd6vLu6YsF21ieWK9Oa4g5blZv5tRtVJhGtY3BhZEwAAAAAAAAAAAAAAAAAAAHOanVtYgAAAClqdW1kY2JvcgARABCAAACqADibcQNjMnBhLmFjdGlvbnMudjIAAAABnWNib3KhZ2FjdGlvbnOCo2ZhY3Rpb25rYzJwYS5vcGVuZWRrZGVzY3JpcHRpb25wT3BlbmVkIGJ5IEdvb2dsZWpwYXJhbWV0ZXJzoWtpbmdyZWRpZW50c4GiY3VybHgtc2VsZiNqdW1iZj1jMnBhLmFzc2VydGlvbnMvYzJwYS5pbmdyZWRpZW50LnYzZGhhc2hYIKnOjLJisRbywiGUIoWEP5B0HWoyoMEpF3/FpA0xJsRso2ZhY3Rpb25vYzJwYS50cmFuc2NvZGVkcWRpZ2l0YWxTb3VyY2VUeXBleEZodHRwOi8vY3YuaXB0Yy5vcmcvbmV3c2NvZGVzL2RpZ2l0YWxzb3VyY2V0eXBlL2FsZ29yaXRobWljYWxseUVuaGFuY2VkanBhcmFtZXRlcnOha2luZ3JlZGllbnRzgaJjdXJseC1zZWxmI2p1bWJmPWMycGEuYXNzZXJ0aW9ucy9jMnBhLmluZ3JlZGllbnQudjNkaGFzaFggqc6MsmKxFvLCIZQihYQ/kHQdajKgwSkXf8WkDTEmxGwAAAZ4anVtYgAAACxqdW1kY2JvcgARABCAAACqADibcQNjMnBhLmluZ3JlZGllbnQudjMAAAAGRGNib3KlaWRjOmZvcm1hdGppbWFnZS9qcGVnbHJlbGF0aW9uc2hpcGhwYXJlbnRPZnF2YWxpZGF0aW9uUmVzdWx0c6FuYWN0aXZlTWFuaWZlc3SjZ2ZhaWx1cmWAZ3N1Y2Nlc3OKomRjb2Rlc3RpbWVTdGFtcC52YWxpZGF0ZWRjdXJseE1zZWxmI2p1bWJmPS9jMnBhL3VybjpjMnBhOmI4MGM5MDE5LWQ1MzEtOThhOC03ODhhLTRkNTRjMTU4ZDZiNC9jMnBhLnNpZ25hdHVyZaJkY29kZXF0aW1lU3RhbXAudHJ1c3RlZGN1cmx4TXNlbGYjanVtYmY9L2MycGEvdXJuOmMycGE6YjgwYzkwMTktZDUzMS05OGE4LTc4OGEtNGQ1NGMxNThkNmI0L2MycGEuc2lnbmF0dXJlomRjb2RleCFzaWduaW5nQ3JlZGVudGlhbC5vY3NwLm5vdFJldm9rZWRjdXJseE1zZWxmI2p1bWJmPS9jMnBhL3VybjpjMnBhOmI4MGM5MDE5LWQ1MzEtOThhOC03ODhhLTRkNTRjMTU4ZDZiNC9jMnBhLnNpZ25hdHVyZaJkY29kZXgZc2lnbmluZ0NyZWRlbnRpYWwudHJ1c3RlZGN1cmx4TXNlbGYjanVtYmY9L2MycGEvdXJuOmMycGE6YjgwYzkwMTktZDUzMS05OGE4LTc4OGEtNGQ1NGMxNThkNmI0L2MycGEuc2lnbmF0dXJlomRjb2RleB1jbGFpbVNpZ25hdHVyZS5pbnNpZGVWYWxpZGl0eWN1cmx4TXNlbGYjanVtYmY9L2MycGEvdXJuOmMycGE6YjgwYzkwMTktZDUzMS05OGE4LTc4OGEtNGQ1NGMxNThkNmI0L2MycGEuc2lnbmF0dXJlomRjb2RleBhjbGFpbVNpZ25hdHVyZS52YWxpZGF0ZWRjdXJseE1zZWxmI2p1bWJmPS9jMnBhL3VybjpjMnBhOmI4MGM5MDE5LWQ1MzEtOThhOC03ODhhLTRkNTRjMTU4ZDZiNC9jMnBhLnNpZ25hdHVyZaJkY29kZXgZYXNzZXJ0aW9uLmhhc2hlZFVSSS5tYXRjaGN1cmx4YXNlbGYjanVtYmY9L2MycGEvdXJuOmMycGE6YjgwYzkwMTktZDUzMS05OGE4LTc4OGEtNGQ1NGMxNThkNmI0L2MycGEuYXNzZXJ0aW9ucy9jMnBhLmluZ3JlZGllbnQudjOiZGNvZGV4GWFzc2VydGlvbi5oYXNoZWRVUkkubWF0Y2hjdXJseF5zZWxmI2p1bWJmPS9jMnBhL3VybjpjMnBhOmI4MGM5MDE5LWQ1MzEtOThhOC03ODhhLTRkNTRjMTU4ZDZiNC9jMnBhLmFzc2VydGlvbnMvYzJwYS5hY3Rpb25zLnYyomRjb2RleBlhc3NlcnRpb24uaGFzaGVk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"

def _v125_apply_visual_theme():
    # CSS içinde çok sayıda { } karakteri bulunduğu için f-string kullanmak
    # Python'un CSS değişkenlerini ifade olarak yorumlamasına ve NameError'a
    # yol açar. Arka plan verisi güvenli bir yer tutucu ile sonradan eklenir.
    css = """
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
                url("__V125_BG_DATA__");
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
        """
    css = css.replace("__V125_BG_DATA__", _V125_BG_DATA)
    st.markdown(css, unsafe_allow_html=True)

_v125_apply_visual_theme()
# ============================================================
# /V125 GÖRSEL TEMA
# ============================================================

# V126 — V125 görsel tema NameError düzeltmesi:
# CSS f-string yerine güvenli placeholder kullanılır. İşlevsel mantık değişmez.

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
