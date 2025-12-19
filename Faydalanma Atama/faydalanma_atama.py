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
