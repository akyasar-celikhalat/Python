import os
import glob
import pandas as pd
from datetime import datetime

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

# kilavuz cols
K_KOD = ['ÜRÜN KODU', 'URUN KODU', 'ÜRÜN_KODU', 'URUNKODU']
K_ISEMRI = ['İŞ EMRİ', 'IS EMRI', 'İŞ_EMRİ']


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
    """Extract diameter in mm from text like 'HT 1.05MM-CT-KY180' or numeric '0.88'."""
    if text is None:
        return None
    s = str(text)
    import re
    # try explicit mm pattern
    m = re.search(r"(\d{1,3}(?:[\.,]\d+)?)[ ]*mm\b", s, flags=re.IGNORECASE)
    if m:
        try:
            return float(m.group(1).replace(',', '.'))
        except:
            return None
    # try pattern like 'HT 1.05MM' or 'HT 1.05MM-CT'
    m2 = re.search(r"(\d{1,3}(?:[\.,]\d+)?)[ ]*MM", s)
    if m2:
        try:
            return float(m2.group(1).replace(',', '.'))
        except:
            return None
    # try standalone number with decimal (e.g. ' 1.05 ')
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
    import math
    area = math.pi * (d_m ** 2) / 4.0
    return area * density


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
    print('Tüketim barkod col:', col_t_barkod, 'iş emri col:', col_t_isemri, 'miktar col:', col_t_amt)

    # prepare output columns
    df_fayd = df_fayd.copy()
    df_fayd['KILAVUZ_ÜRÜN_KODU'] = None
    df_fayd['TUKETIM_MIKTARI_KG'] = 0.0

    unmatched_isemri = []
    unmatched_barkod = []

    # build consumption map: key = (işemri, barkod) -> sum(miktar)
    cons_map = {}
    if df_tuk is not None and col_t_barkod and col_t_amt:
        for _, r in df_tuk.iterrows():
            b = r.get(col_t_barkod)
            i = r.get(col_t_isemri) if col_t_isemri else None
            amt = to_num(r.get(col_t_amt))
            if b is None:
                continue
            key = (str(i).strip() if i is not None else '', str(b).strip())
            cons_map[key] = cons_map.get(key, 0.0) + amt

    # build kilavuz map: işemri -> ürün kodu (take first match)
    kil_map = {}
    if df_kil is not None and col_k_isemri and col_k_urunkod:
        for _, r in df_kil.iterrows():
            i = r.get(col_k_isemri)
            kod = r.get(col_k_urunkod)
            if i is None:
                continue
            kil_map[str(i).strip()] = str(kod).strip() if kod is not None else None

    # process each faydalanma row
    for idx, r in df_fayd.iterrows():
        isemri = r.get(col_f_isemri) if col_f_isemri else None
        barkod = r.get(col_f_barkod) if col_f_barkod else None
        isemri_s = str(isemri).strip() if isemri is not None else ''
        barkod_s = str(barkod).strip() if barkod is not None else ''

        # kilavuz lookup
        if isemri_s and kil_map.get(isemri_s) is not None:
            df_fayd.at[idx, 'KILAVUZ_ÜRÜN_KODU'] = kil_map.get(isemri_s)
        else:
            # not found
            df_fayd.at[idx, 'KILAVUZ_ÜRÜN_KODU'] = None
            unmatched_isemri.append(isemri_s)

        # consumption lookup by (işemri, barkod)
        key = (isemri_s, barkod_s)
        amt = cons_map.get(key, 0.0)
        df_fayd.at[idx, 'TUKETIM_MIKTARI_KG'] = amt
        if amt == 0.0:
            # if consumption not found, record
            unmatched_barkod.append((isemri_s, barkod_s))

    # Summarize unmatched
    unmatched_isemri = list({x for x in unmatched_isemri if x})
    unmatched_barkod = list({x for x in unmatched_barkod if x and x[1]})

    # --- Yeni: STOK_KG ve TUKETIM_KG hesapları (kullanıcı isteğine göre) ---
    # 1) STOK_KG: 'STOK ÜRÜN TANIMI' sütunundaki çapı parse et ve 'STOK MİKTARI' (metre) ile çarp
    # 2) TUKETIM_KG: 'KILAVUZ_ÜRÜN_KODU' veya kilavuz eşlemesinden çap al, 'TUKETIM_MIKTARI_KG' (metre) ile çarp

    # stok tanımı sütunu araması (kullanıcının belirttiği isim öncelikli)
    stok_desc_candidates = ['STOK ÜRÜN TANIMI', 'STOK_ÜRÜN_TANIMI', 'STOK_URUN_TANIMI', 'STOK URUN TANIMI', 'ÜRÜN TANIMI', 'ÜRÜN_AÇIKLAMA']
    stok_desc_col = None
    for c in stok_desc_candidates:
        if c in df_fayd.columns:
            stok_desc_col = c
            break

    # stok metre sütunu
    stok_amount_candidates = ['STOK MİKTARI', 'STOK_MIKTAR', 'STOK_MİKTAR', 'STOK', 'STOK_MIKTARI_METRE']
    stok_amount_col = find_col(df_fayd, stok_amount_candidates)

    df_fayd['STOK_KG'] = None
    if stok_desc_col and stok_amount_col:
        # parse çapı satır satır
        def stok_row_stokkg(row):
            desc = row.get(stok_desc_col)
            d_mm = parse_diameter_mm(desc)
            if d_mm is None:
                # fallback: eğer KILAVUZ_ÜRÜN_KODU varsa, deneyebiliriz
                kprod = row.get('KILAVUZ_ÜRÜN_KODU')
                if kprod:
                    d_mm = parse_diameter_mm(kprod)
            if d_mm is None:
                return None
            kg_per_m = linear_kg_per_m_from_d_mm(d_mm)
            metres = to_num(row.get(stok_amount_col))
            if metres is None or kg_per_m is None:
                return None
            return metres * kg_per_m
        df_fayd['STOK_KG'] = df_fayd.apply(stok_row_stokkg, axis=1)

    # Tüketim için metre sütunu adını bul
    cons_amount_col = None
    if 'TUKETIM_MIKTARI_KG' in df_fayd.columns:
        cons_amount_col = 'TUKETIM_MIKTARI_KG'
    else:
        cons_amount_col = find_col(df_fayd, ['TUKETIM', 'TÜKETİM', 'TUKETIM_MIKTARI', 'TUKETIM_MIKTARI_METRE'])

    df_fayd['TUKETIM_KG'] = None
    # hazırda varsa kilavuztan oluşturulmuş harita (iş emri -> çap) kullan
    kil_diameter = {}
    if 'df_kil' in locals():
        # df_kil'den numeric 'ÇAP' veya açıklamadan parse et
        cap_col = find_col(df_kil, ['ÇAP', 'CAP', 'ÇAPI'])
        for _, r in df_kil.iterrows():
            key = r.get(col_k_isemri) if col_k_isemri in r else None
            if pd.isna(key) or not key:
                continue
            kkey = str(key).strip()
            d = None
            if cap_col:
                try:
                    dval = r.get(cap_col)
                    if pd.notna(dval):
                        d = float(dval)
                except Exception:
                    d = None
            if d is None:
                # try parsing from any description columns
                for cc in r.index:
                    if isinstance(r.get(cc), str):
                        d = parse_diameter_mm(r.get(cc))
                        if d is not None:
                            break
            if d is not None:
                kil_diameter[kkey] = d

    if cons_amount_col:
        def tuketim_row_kg(row):
            metres = to_num(row.get(cons_amount_col))
            if metres is None:
                return None
            # if product code begins with DMT -> use STOK_KG as tüketim kg
            kprod = row.get('KILAVUZ_ÜRÜN_KODU')
            if kprod and str(kprod).upper().startswith('DMT'):
                stokkg = row.get('STOK_KG')
                try:
                    return float(stokkg) if stokkg is not None else None
                except Exception:
                    return None
            # öncelik 1: parse from KILAVUZ_ÜRÜN_KODU field itself
            d_mm = parse_diameter_mm(row.get('KILAVUZ_ÜRÜN_KODU'))
            # öncelik 2: use kilavuz eşlemesinden gelen çap (iş emri anahtarlı)
            if d_mm is None:
                isemri = row.get('İŞ EMRİ') if 'İŞ EMRİ' in row else row.get('IS EMRI')
                if isemri is not None:
                    d_mm = kil_diameter.get(str(isemri).strip())
            # öncelik 3: parse from genel açıklama sütunu varsa
            if d_mm is None:
                for c in ['ÜRÜN AÇIKLAMA', 'ÜRÜN_AÇIKLAMA', 'MALZEME ADI', 'MALZEME_ADI']:
                    if c in row:
                        d_mm = parse_diameter_mm(row.get(c))
                        if d_mm is not None:
                            break
            if d_mm is None:
                return None
            kg_per_m = linear_kg_per_m_from_d_mm(d_mm)
            if kg_per_m is None:
                return None
            return metres * kg_per_m
        df_fayd['TUKETIM_KG'] = df_fayd.apply(tuketim_row_kg, axis=1)

    # --- Yeni: STOK_DURUM hesapla (sadece STOK ÜRÜN TANIMI TV/TG ile başlayanlar için) ---
    df_fayd['STOK_DURUM'] = None
    stok_name_col = stok_desc_col if stok_desc_col else find_col(df_fayd, ['STOK ÜRÜN TANIMI', 'STOK_ÜRÜN_TANIMI'])

    def determine_stok_durum(row):
        name = str(row.get(stok_name_col, '') or '')
        if not name:
            return None
        pref = name.strip()[:2].upper()
        if pref not in ('TV', 'TG'):
            return None
        # find diameter
        d_mm = parse_diameter_mm(name)
        if d_mm is None:
            d_mm = parse_diameter_mm(row.get('KILAVUZ_ÜRÜN_KODU'))
        if d_mm is None and 'kil_diameter' in locals():
            isemri = row.get('İŞ EMRİ') if 'İŞ EMRİ' in row else row.get('IS EMRI')
            if isemri is not None:
                d_mm = kil_diameter.get(str(isemri).strip())
        if d_mm is None:
            return None
        # capacity mapping
        try:
            if float(d_mm) <= 2.30:
                cap = 400.0
            elif float(d_mm) <= 5.5:
                cap = 850.0
            else:
                cap = 1500.0
        except Exception:
            return None
        stokkg = row.get('STOK_KG')
        try:
            stokkg_val = float(stokkg) if stokkg is not None else None
        except Exception:
            stokkg_val = None
        if stokkg_val is None:
            return None
        return 'PARÇA' if stokkg_val < 0.25 * cap else 'TAM'

    df_fayd['STOK_DURUM'] = df_fayd.apply(determine_stok_durum, axis=1)

    # --- Yeni sütun: EşMi (STOK ÜRÜN TANIMI ilk 3 karakter == ürün ilk 3 karakter) ---
    prod_col = find_col(df_fayd, ['ÜRÜN', 'ÜRÜN ADI', 'MALZEME ADI', 'ÜRÜN_ADI', 'MALZEME_ADI'])
    stok_name_col = stok_desc_col if stok_desc_col else find_col(df_fayd, ['STOK ÜRÜN TANIMI', 'STOK_ÜRÜN_TANIMI', 'STOK URUN TANIMI'])

    def esmi_row(row):
        left = str(row.get(stok_name_col, '') or '')[:3]
        right = str(row.get(prod_col, '') or '')[:3]
        return 'DOĞRU' if left == right and left != '' else 'YANLIŞ'

    df_fayd['EşMi'] = df_fayd.apply(esmi_row, axis=1)


    now = datetime.now().strftime('%Y%m%d_%H%M%S')
    out_file = os.path.join(BASE_DIR, f'faydalanma_atama_sonuc_{now}.xlsx')

    with pd.ExcelWriter(out_file, engine='openpyxl') as writer:
        df_fayd.to_excel(writer, sheet_name='Faydalanma_Atama', index=False)

        # write summaries
        pd.DataFrame({'unmatched_isemri': unmatched_isemri}).to_excel(writer, sheet_name='Eksik_IsEmri', index=False)
        if unmatched_barkod:
            pd.DataFrame(unmatched_barkod, columns=['IS_EMRI','BARKOD']).to_excel(writer, sheet_name='Eksik_Barkod', index=False)

    print('İşlem tamam. Rapor kaydedildi:', out_file)
    print('Eksik İş Emirleri sayısı:', len(unmatched_isemri))
    print('Eksik Barkod eşleşmeleri sayısı:', len(unmatched_barkod))


if __name__ == '__main__':
    main()
