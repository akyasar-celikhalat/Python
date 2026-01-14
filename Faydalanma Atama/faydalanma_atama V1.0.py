import os
import glob
import pandas as pd
from datetime import datetime
import re
import math

BASE_DIR = os.path.dirname(__file__)

KEYS = {
    'faydalanma': ['faydalanma', 'faydal', 'faydalanma'],
    'kilavuz': ['kılavuz', 'kilavuz', 'kılavuz', 'kılavuz'],
    'tuketim': ['tuketim', 'tüketim', 'consumption']
}

CANDIDATE_BARKOD_COLS = ['BARKOD NUMARASI', 'BARKOD', 'Barkod Numarası', 'Barkod']
CANDIDATE_ISEMRI_COLS = ['İŞ EMRİ', 'IS EMRI', 'İŞ_EMRİ', 'IS_EMRI', 'İŞ EMRI']

# consumption candidate cols
C_TUKETIM_BARKOD = ['GİRİŞ ÜRÜN SAP BARKODU', 'GIRIS URUN SAP BARKODU', 'GİRİŞ ÜRÜN SAP BARKOD', 'GIRIS_BARKOD']
C_TUKETIM_ISEMRI = ['İŞ EMRİ', 'IS EMRI', 'İŞ_EMRİ']
C_TUKETIM_AMT = ['GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg', 'GIRIS URUN TUKETIM MIKTARI', 'TÜKETİM MİKTARI', 'Tuketim Miktari']

# kilavuz cols (Yeni sütun isimleri eklendi)
K_KOD = ['dssaptranswo.Matxt', 'ÜRÜN KODU', 'URUN KODU', 'ÜRÜN_KODU', 'URUNKODU']
K_ISEMRI = ['A360NO_02', 'İŞ EMRİ', 'IS EMRI', 'İŞ_EMRİ']


def find_files(search_dirs=None):
    search_dirs = search_dirs or [BASE_DIR, os.path.dirname(BASE_DIR)]
    found = {}
    for d in search_dirs:
        if not d or not os.path.isdir(d):
            continue
        for path in glob.glob(os.path.join(d, '*.xls*')):
            name = os.path.basename(path).lower()
            if name.startswith('~$'):
                continue
            for k, keys in KEYS.items():
                for kw in keys:
                    if kw in name:
                        found[k] = path
    return found


def safe_read(path, sheet_name=None):
    if not path:
        return None
    try:
        if sheet_name:
            try:
                return pd.read_excel(path, sheet_name=sheet_name, engine='openpyxl')
            except ValueError:
                sheets = pd.read_excel(path, sheet_name=None, engine='openpyxl')
                for k, df in sheets.items():
                    if k and k.casefold() == sheet_name.casefold():
                        return df
                return next(iter(sheets.values()))
        else:
            return pd.read_excel(path, engine='openpyxl')
    except Exception as e:
        print(f"Dosya okunamadi: {path} -> {e}")
        return None


def find_col(df, candidates):
    if df is None:
        return None
    cols = list(df.columns)
    low = {c.casefold(): c for c in cols}
    for c in candidates:
        if c in cols:
            return c
        if c.casefold() in low:
            return low[c.casefold()]
    # substring
    for ln, orig in low.items():
        for c in candidates:
            if c.casefold() in ln or ln in c.casefold():
                return orig
    return None


def to_num(x):
    try:
        if pd.isna(x):
            return 0.0
        return float(x)
    except:
        try:
            s = str(x).replace(',', '.')
            filtered = ''.join(ch for ch in s if ch.isdigit() or ch == '.')
            return float(filtered) if filtered not in ('', '.') else 0.0
        except:
            return 0.0


def parse_diameter_mm(text):
    if text is None:
        return None
    s = str(text)
    # try explicit mm pattern
    m = re.search(r"(\d{1,3}(?:[\.,]\d+)?)[ ]*mm\b", s, flags=re.IGNORECASE)
    if m:
        try:
            return float(m.group(1).replace(',', '.'))
        except:
            return None
    # try pattern like 'HT 1.05MM'
    m2 = re.search(r"(\d{1,3}(?:[\.,]\d+)?)[ ]*MM", s)
    if m2:
        try:
            return float(m2.group(1).replace(',', '.'))
        except:
            return None
    # try standalone number
    m3 = re.search(r"\b(\d{1,3}[\.,]\d+)\b", s)
    if m3:
        try:
            return float(m3.group(1).replace(',', '.'))
        except:
            return None
    return None


def linear_kg_per_m_from_d_mm(d_mm, density=7850.0):
    if d_mm is None:
        return None
    d_m = float(d_mm) / 1000.0
    area = math.pi * (d_m ** 2) / 4.0
    return area * density


def clean_isemri(isemri_val):
    """İş emrini normalize eder: 8 hane ve sonu 00 ise 6 haneye düşürür."""
    if isemri_val is None:
        return ""
    s = str(isemri_val).strip()
    # Eğer 8 hane uzunluğundaysa ve sonu 00 ile bitiyorsa, son iki karakteri sil
    if len(s) == 8 and s.endswith('00'):
        return s[:-2]
    return s


def main():
    files = find_files()
    print('Bulunan dosyalar:', files)

    fayd_file = files.get('faydalanma')
    kilavuz_file = files.get('kilavuz')
    tuk_file = files.get('tuketim')

    df_fayd = safe_read(fayd_file)
    df_kil = safe_read(kilavuz_file)
    df_tuk = safe_read(tuk_file)

    if df_fayd is None:
        print('Faydalanma dosyası bulunamadı veya okunamadı. Çıkılıyor.')
        return

    # detect columns
    col_f_barkod = find_col(df_fayd, CANDIDATE_BARKOD_COLS)
    col_f_isemri = find_col(df_fayd, CANDIDATE_ISEMRI_COLS)

    col_k_isemri = find_col(df_kil, K_ISEMRI) if df_kil is not None else None
    col_k_urunkod = find_col(df_kil, K_KOD) if df_kil is not None else None

    col_t_barkod = find_col(df_tuk, C_TUKETIM_BARKOD) if df_tuk is not None else None
    col_t_isemri = find_col(df_tuk, C_TUKETIM_ISEMRI) if df_tuk is not None else None
    col_t_amt = find_col(df_tuk, C_TUKETIM_AMT) if df_tuk is not None else None

    print('Faydalanma barkod kolonu:', col_f_barkod, 'İş emri kolonu:', col_f_isemri)
    print('Kılavuz iş emri kolonu:', col_k_isemri, 'ürün kodu kolonu:', col_k_urunkod)

    # prepare output columns
    df_fayd = df_fayd.copy()
    df_fayd['KILAVUZ_ÜRÜN_KODU'] = None
    df_fayd['TUKETIM_MIKTARI_KG'] = 0.0

    unmatched_isemri = []
    unmatched_barkod = []

    # build consumption map: key = (işemri, barkod) -> sum(miktar)
    cons_map = {}
    if df_tuk is not None and col_t_barkod:
        out_desc_col = find_col(df_tuk, ['ÇIKIŞ ÜRÜN AÇIKLAMA', 'CIKIS URUN ACIKLAMA', 'ÇIKIŞ_ÜRÜN_AÇIKLAMA', 'CIKIS_URUN_ACIKLAMA'])
        teyit_col = find_col(df_tuk, ['TEYİT MİKTARI Kg', 'TEYIT MIKTARI Kg', 'TEYIT_MIKTARI_KG', 'TEYIT_MIKTARI'])
        giris_tuketim_col = find_col(df_tuk, ['GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg', 'GIRIS URUN TUKETIM MIKTARI', 'GIRIS_URUN_TUKETIM_MIKTARI'])
        for _, r in df_tuk.iterrows():
            b = r.get(col_t_barkod)
            i = r.get(col_t_isemri) if col_t_isemri else None
            if b is None:
                continue
            
            out_desc_val = r.get(out_desc_col) if out_desc_col else None
            use_teyit = False
            try:
                if isinstance(out_desc_val, str) and out_desc_val.strip().upper().startswith('DMT'):
                    use_teyit = True
            except:
                use_teyit = False
            
            amt = to_num(r.get(teyit_col)) if use_teyit and teyit_col else to_num(r.get(giris_tuketim_col))
            
            # Tüketim dosyasındaki iş emrini de temizle (standart olması için)
            key = (clean_isemri(i), str(b).strip())
            cons_map[key] = cons_map.get(key, 0.0) + amt

    # build kilavuz map: işemri -> ürün kodu & çap
    kil_map = {}
    kil_diameter = {}
    if df_kil is not None and col_k_isemri and col_k_urunkod:
        cap_col = find_col(df_kil, ['ÇAP', 'CAP', 'ÇAPI'])
        for _, r in df_kil.iterrows():
            raw_isemri = r.get(col_k_isemri)
            if pd.isna(raw_isemri): continue
            
            # --- KRİTİK DÜZENLEME: Kılavuzdaki iş emrini 8->6 hane yap ---
            i_clean = clean_isemri(raw_isemri)
            
            kod = r.get(col_k_urunkod)
            kil_map[i_clean] = str(kod).strip() if kod is not None else None
            
            # Çap bulma
            d = None
            if cap_col:
                try:
                    dval = r.get(cap_col)
                    if pd.notna(dval): d = float(dval)
                except: d = None
            if d is None:
                for cc in r.index:
                    if isinstance(r.get(cc), str):
                        d = parse_diameter_mm(r.get(cc))
                        if d is not None: break
            if d is not None:
                kil_diameter[i_clean] = d

    # process each faydalanma row
    for idx, r in df_fayd.iterrows():
        isemri = r.get(col_f_isemri)
        barkod = r.get(col_f_barkod)
        
        # Faydalanma dosyasındaki iş emri zaten 6 hane ama standart olması için clean kullanıyoruz
        isemri_s = clean_isemri(isemri)
        barkod_s = str(barkod).strip() if barkod is not None else ''

        # kilavuz lookup
        if isemri_s in kil_map:
            df_fayd.at[idx, 'KILAVUZ_ÜRÜN_KODU'] = kil_map[isemri_s]
        else:
            unmatched_isemri.append(isemri_s)

        # consumption lookup
        key = (isemri_s, barkod_s)
        amt = cons_map.get(key, 0.0)
        df_fayd.at[idx, 'TUKETIM_MIKTARI_KG'] = amt
        if amt == 0.0:
            unmatched_barkod.append((isemri_s, barkod_s))

    # --- STOK_KG ve TUKETIM_KG hesapları ---
    stok_desc_candidates = ['STOK ÜRÜN TANIMI', 'STOK_ÜRÜN_TANIMI', 'STOK_URUN_TANIMI', 'STOK URUN TANIMI', 'ÜRÜN TANIMI', 'ÜRÜN_AÇIKLAMA']
    stok_desc_col = find_col(df_fayd, stok_desc_candidates)
    stok_amount_col = find_col(df_fayd, ['STOK MİKTARI', 'STOK_MIKTAR', 'STOK_MİKTAR', 'STOK', 'STOK_MIKTARI_METRE'])

    df_fayd['STOK_KG'] = None
    if stok_desc_col and stok_amount_col:
        def stok_row_stokkg(row):
            desc = row.get(stok_desc_col)
            d_mm = parse_diameter_mm(desc) or parse_diameter_mm(row.get('KILAVUZ_ÜRÜN_KODU'))
            if d_mm:
                kg_per_m = linear_kg_per_m_from_d_mm(d_mm)
                metres = to_num(row.get(stok_amount_col))
                return metres * kg_per_m
            return None
        df_fayd['STOK_KG'] = df_fayd.apply(stok_row_stokkg, axis=1)

    # Tüketim KG (metre * linear kg)
    df_fayd['TUKETIM_KG'] = None
    def tuketim_row_kg(row):
        metres = to_num(row.get('TUKETIM_MIKTARI_KG'))
        d_mm = parse_diameter_mm(row.get(stok_desc_col)) if stok_desc_col else None
        if not d_mm:
            isemri_key = clean_isemri(row.get(col_f_isemri))
            d_mm = kil_diameter.get(isemri_key)
            
        if d_mm and metres:
            kg_per_m = linear_kg_per_m_from_d_mm(d_mm)
            return metres * kg_per_m
        return 0.0
    df_fayd['TUKETIM_KG'] = df_fayd.apply(tuketim_row_kg, axis=1)

    # --- STOK_DURUM (PARÇA/TAM) ---
    def determine_stok_durum(row):
        name = str(row.get(stok_desc_col, '') or '')
        if not name.strip()[:2].upper() in ('TV', 'TG'): return None
        
        d_mm = parse_diameter_mm(name) or kil_diameter.get(clean_isemri(row.get(col_f_isemri)))
        if d_mm is None: return None
        
        if d_mm <= 2.30: cap = 400.0
        elif d_mm <= 5.5: cap = 850.0
        else: cap = 1500.0
        
        stokkg = to_num(row.get('STOK_KG'))
        return 'PARÇA' if stokkg < 0.25 * cap else 'TAM'

    df_fayd['STOK_DURUM'] = df_fayd.apply(determine_stok_durum, axis=1)

    # --- EŞMİ Sütunu ---
    prod_col = find_col(df_fayd, ['ÜRÜN', 'ÜRÜN ADI', 'MALZEME ADI', 'ÜRÜN_ADI', 'MALZEME_ADI'])
    def esmi_row(row):
        left = str(row.get(stok_desc_col, '') or '')[:3]
        right = str(row.get(prod_col, '') or '')[:3]
        return 'DOĞRU' if left == right and left != '' else 'YANLIŞ'
    df_fayd['EŞMİ'] = df_fayd.apply(esmi_row, axis=1)

    # Kayıt
    now = datetime.now().strftime('%Y%m%d_%H%M%S')
    out_file = os.path.join(BASE_DIR, f'sonuc_{now}.xlsx')

    rename_map = {
        'KILAVUZ_ÜRÜN_KODU': 'Eslenen_Kilavuz_Urun_Kodu',
        'TUKETIM_MIKTARI_KG': 'Tuketim_Metre_Orijinal (m)',
        'TUKETIM_KG': 'Tuketim_KG_Hesaplanmis (kg)',
        'STOK_KG': 'Stok_KG_Hesaplanmis (kg)',
        'STOK_DURUM': 'Stok_Durumu (PARCA/TAM)',
        'EŞMİ': 'UrunAdı_Eslesme (DOĞRU/YANLIŞ)'
    }
    df_out = df_fayd.rename(columns=rename_map)

    with pd.ExcelWriter(out_file, engine='openpyxl') as writer:
        df_out.to_excel(writer, sheet_name='Faydalanma_Atama_Sonuclari', index=False)
        pd.DataFrame({'unmatched_isemri': list(set(unmatched_isemri))}).to_excel(writer, sheet_name='Eksik_Is_Emri_Listesi', index=False)
        if unmatched_barkod:
            pd.DataFrame(list(set(unmatched_barkod)), columns=['IS_EMRI', 'BARKOD']).to_excel(writer, sheet_name='Eksik_Barkod_Eslestirmeleri', index=False)

    print(f'İşlem tamam. Rapor kaydedildi: {out_file}')

if __name__ == '__main__':
    main()