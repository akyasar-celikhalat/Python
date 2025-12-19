import os
import glob
import re
import pandas as pd

BASE_DIR = os.path.dirname(__file__)

# Eğer True ise stok değeri eksikse 0 ile doldurulur; False ise boş bırakılır
FILL_EMPTY_STOCK = True

KEYS = {
    'sayim': ['sayim', 'sayım', 'count'],
    'uretim': ['uretim', 'üretim', 'production'],
    'tuketim': ['tuketim', 'tüketim', 'consumption'],
    'stok': ['stok', 'stock']
}

def find_files():
    files = {}
    for path in glob.glob(os.path.join(BASE_DIR, '*.xls*')):
        name = os.path.basename(path)
        if name.startswith('~$'):
            continue
        lname = name.lower()
        for k, keys in KEYS.items():
            if any(key in lname for key in keys):
                files[k] = path
    return files

def safe_read(path, sheet_name=None):
    if not path:
        return None
    try:
        if sheet_name:
            try:
                return pd.read_excel(path, sheet_name=sheet_name, engine='openpyxl')
            except ValueError:
                # belirtilen sayfa yoksa tüm sayfaları oku ve isim eşleşmesi ara (case-insensitive)
                try:
                    sheets = pd.read_excel(path, sheet_name=None, engine='openpyxl')
                    for k, df in sheets.items():
                        if k and k.casefold() == sheet_name.casefold():
                            return df
                    # eşleşme yoksa ilk sayfayı döndür
                    return next(iter(sheets.values()))
                except Exception:
                    return None
        else:
            return pd.read_excel(path, engine='openpyxl')
    except PermissionError:
        print(f"Uyarı: Dosya açılamadı, atlanıyor: {path}")
        return None
    except Exception as e:
        print(f"Uyarı: Okuma hatası {path} -> {e}")
        return None

def find_col(df, names):
    if df is None:
        return None
    cols = list(df.columns)
    # use casefold for robust unicode insensitive matching
    low_map = {c.casefold(): c for c in cols}
    for n in names:
        if n in cols:
            return n
        nf = n.casefold()
        if nf in low_map:
            return low_map[nf]
    # substring fuzzy using casefold
    for ln, orig in low_map.items():
        for n in names:
            if n.casefold() in ln or ln in n.casefold():
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

def fmt_num_for_excel(x):
    # round to integer, use dot as thousands separator (Türkçe binlik ayraç isteğine göre)
    if x is None:
        return ''
    try:
        if pd.isna(x):
            return ''
    except:
        pass
    try:
        n = round(float(x))
    except:
        return ''
    s = "{:,.0f}".format(n)  # produces commas
    # convert comma grouping to dot (1,234 -> 1.234) for Turkish-style thousands separator
    return s.replace(',', '.')

def aggregate_add(d, key, amt):
    if key is None:
        return
    k = str(key).strip()
    if k == '':
        return
    d[k] = d.get(k, 0.0) + to_num(amt)

def build_prod_from_sayim(df):
    prod = {}
    if df is None:
        return prod
    col_bar = find_col(df, ['BOBİN', 'BOBIN', 'BOBIN ', 'Bobin', 'Barkod'])
    col_m = find_col(df, ['METRE', 'Metre', 'METRE '])
    if col_bar is None:
        col_bar = df.columns[0] if len(df.columns)>0 else None
    for _, r in df.iterrows():
        aggregate_add(prod, r.get(col_bar), r.get(col_m))
    return prod

def build_prod_from_uretim(df):
    prod = {}
    if df is None:
        return prod
    col_bar = find_col(df, ['ÜRETİLEN BARKOD', 'URETILEN BARKOD', 'BARKOD', 'Barkod'])
    col_m = find_col(df, ['METRE_02', 'METRE 02', 'METRE02', 'METRE'])
    for _, r in df.iterrows():
        aggregate_add(prod, r.get(col_bar), r.get(col_m))
    return prod

def build_cons_from_tuketim(df):
    cons = {}
    if df is None:
        return cons
    col_bar = find_col(df, ['GİRİŞ ÜRÜN SAP BARKODU', 'GIRIS URUN SAP BARKODU', 'Barkod', 'BARKOD'])
    col_desc = find_col(df, ['ÇIKIŞ ÜRÜN ACIKLAMA', 'CIKIS URUN ACIKLAMA', 'AÇIKLAMA', 'Aciklama'])
    col_teyit = find_col(df, ['TEYİT MİKTARI Kg', 'TEYIT MIKTARI Kg', 'TEYIT MIKTARI', 'TEYİT MİKTARI'])
    col_giris = find_col(df, ['GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg', 'GIRIS URUN TUKETIM MIKTARI', 'GİRİŞ ÜRÜN TÜKETİM MİKTARI', 'GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg'])
    for _, r in df.iterrows():
        code = r.get(col_bar)
        desc = '' if col_desc is None else str(r.get(col_desc) or '')
        desc_u = desc.strip().upper()
        if desc_u.startswith('H') or desc_u.startswith('H ') or desc_u.startswith('DMT'):
            amt = r.get(col_teyit) if col_teyit is not None else r.get(col_giris)
        else:
            amt = r.get(col_giris) if col_giris is not None else r.get(col_teyit)
        aggregate_add(cons, code, amt)
    return cons

def build_stock(df):
    stok = {}
    if df is None:
        return stok
    col_bar = find_col(df, ['BARKOD KODU', 'Barkod Kodu', 'BARKOD', 'Barkod'])
    col_amt = find_col(df, ['MİKTAR', 'MIKTAR', 'Miktar', 'MİKTAR '])
    for _, r in df.iterrows():
        aggregate_add(stok, r.get(col_bar), r.get(col_amt))
    return stok


def extract_meta(df, barcode_candidates, code_candidates, desc_candidates):
    """Return two dicts mapping barcode -> code and barcode -> description."""
    code_map = {}
    desc_map = {}
    if df is None:
        return code_map, desc_map
    col_bar = find_col(df, barcode_candidates)
    col_code = find_col(df, code_candidates)
    col_desc = find_col(df, desc_candidates)
    for _, r in df.iterrows():
        bar = r.get(col_bar)
        if bar is None:
            continue
        b = str(bar).strip()
        if b == '':
            continue
        if col_code is not None:
            c = r.get(col_code)
            if c is not None and str(c).strip() != '':
                code_map[b] = str(c).strip()
        if col_desc is not None:
            d = r.get(col_desc)
            if d is not None and str(d).strip() != '':
                desc_map[b] = str(d).strip()
    return code_map, desc_map

def build_single_table(prod, sayim, cons, stok, sayim_meta=None, uretim_meta=None, stok_meta=None, cons_meta=None):
    all_codes = set().union(prod.keys(), sayim.keys(), cons.keys(), stok.keys())
    rows = []
    for code in sorted(all_codes):
        # Barkod kontrolü: "-" tire işareti yoksa veya "M" ile başlıyorsa atla
        if '-' not in str(code) or str(code).startswith('M'):
            continue
        s_amt = to_num(sayim.get(code, 0.0))
        p_amt = to_num(prod.get(code, 0.0))
        c_amt = to_num(cons.get(code, 0.0))
        stok_amt = stok.get(code)
        stok_amt_n = None if stok_amt is None else to_num(stok_amt)

        expected = s_amt + p_amt - c_amt

        
        # Eğer sayım ve üretim değeri yoksa (her ikisi de 0) ve tüketim varsa,
        # "üretilmeden tüketilen" sütununda doğrudan tüketim gösterilsin.
        if (abs(s_amt) < 1e-9) and (abs(p_amt) < 1e-9) and (c_amt > 0):
            uretilmeden_tuketilen = c_amt
        else:
            uretilmeden_tuketilen = 0.0
        uretimden_fazla_tuketim = max(0.0, c_amt - p_amt) if c_amt > 0 else 0.0
        uretilip_kullanilmayan = max(0.0, p_amt - c_amt)
        stokta_uretilmemis_olan = stok_amt_n if (stok_amt_n is not None and p_amt <= 0 and stok_amt_n > 0) else 0.0
        # Eğer sayım ve üretim değeri yok (0) ve stokta kayıt yok veya 0 ise
        # ve tüketim varsa, "Stokta Yok Ama Tüketilmiş" sütununa tüketim yazılsın.
        if (abs(s_amt) < 1e-9) and (abs(p_amt) < 1e-9) and (stok_amt_n is None or abs(stok_amt_n) < 1e-9) and (c_amt > 0):
            stokta_yok_ama_tuketim = c_amt
        else:
            stokta_yok_ama_tuketim = 0.0
        stok_uyusmazligi = None
        if stok_amt_n is not None:
            stok_uyusmazligi = stok_amt_n - expected
        
        
        # Malzeme kodu ve açıklama: öncelik sayım -> üretim -> stok -> tüketim
        malzeme_kodu = None
        malzeme_aciklama = None
        if sayim_meta is not None:
            malzeme_kodu = sayim_meta[0].get(code) or malzeme_kodu
            malzeme_aciklama = sayim_meta[1].get(code) or malzeme_aciklama
        if uretim_meta is not None:
            malzeme_kodu = malzeme_kodu or uretim_meta[0].get(code)
            malzeme_aciklama = malzeme_aciklama or uretim_meta[1].get(code)
        if stok_meta is not None:
            malzeme_kodu = malzeme_kodu or stok_meta[0].get(code)
            malzeme_aciklama = malzeme_aciklama or stok_meta[1].get(code)
        if cons_meta is not None:
            malzeme_kodu = malzeme_kodu or cons_meta[0].get(code)
            malzeme_aciklama = malzeme_aciklama or cons_meta[1].get(code)

        rows.append({
            'Barkod': code,
            'Sayım': s_amt,
            'Üretim': p_amt,
            'Tüketim': c_amt,
            'Stok': stok_amt_n,
            'Malzeme Kodu': malzeme_kodu,
            'Malzeme Açıklama': malzeme_aciklama,
            'Beklenen Stok': expected,
            'Stok Farkı': stok_uyusmazligi,
            'Üretilmeden Tüketilen': uretilmeden_tuketilen,
            'Üretimden Fazla Tüketilen': uretimden_fazla_tuketim,
            'Üretilip Kullanılmayan': uretilip_kullanilmayan,
            'Stokta Üretilmemiş Olan': stokta_uretilmemis_olan,
            'Stokta Yok Ama Tüketilmiş': stokta_yok_ama_tuketim
        })
    df = pd.DataFrame(rows)
    # Malzeme bilgilerini ilk iki sütun yap: 'Malzeme Kodu', 'Malzeme Açıklama'
    cols = list(df.columns)
    new_order = []
    for c in ['Malzeme Kodu', 'Malzeme Açıklama']:
        if c in cols:
            new_order.append(c)
    for c in cols:
        if c not in new_order:
            new_order.append(c)
    df = df[new_order]

    # Numeric columns: convert to numbers so Excel can compute on them
    numeric_cols = [
        'Sayım', 'Üretim', 'Tüketim',
        'Beklenen Stok', 'Stok', 'Stok Farkı',
        'Üretilmeden Tüketilen', 'Üretimden Fazla Tüketilen', 'Üretilip Kullanılmayan',
        'Stokta Üretilmemiş Olan', 'Stokta Yok Ama Tüketilmiş'
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    # Stok eksikse davranış: doldur veya bırak
    if 'Stok' in df.columns:
        if FILL_EMPTY_STOCK:
            df['Stok'] = df['Stok'].fillna(0.0)

    return df, None  # return dataframe (numeric). formatted view not used to avoid strings

def main():
    files = find_files()
    # Sayım dosyasından "YARI MAMUL" sayfasını oku (yoksa ilk sayfa)
    df_sayim = safe_read(files.get('sayim'), sheet_name='YARI MAMUL')
    df_uretim = safe_read(files.get('uretim'))
    df_tuketim = safe_read(files.get('tuketim'))
    df_stok = safe_read(files.get('stok'))

    # üretim
    prod_from_uretim = build_prod_from_uretim(df_uretim)
    prod = {k: to_num(v) for k, v in prod_from_uretim.items()}

    # sayım (YARI MAMUL sayfasından alınmış df_sayim)
    sayim = {}
    if df_sayim is not None:
        col_bobin = find_col(df_sayim, ['BOBİN', 'BOBIN', 'Bobin', 'Barkod'])
        col_metre = find_col(df_sayim, ['METRE', 'Metre'])
        if col_bobin is None:
            col_bobin = df_sayim.columns[0] if len(df_sayim.columns)>0 else None
        for _, r in df_sayim.iterrows():
            aggregate_add(sayim, r.get(col_bobin), r.get(col_metre))

    cons = build_cons_from_tuketim(df_tuketim)
    stok = build_stock(df_stok)

    # Meta (malzeme kodu/açıklama) çıkar
    sayim_meta = extract_meta(df_sayim,
                              ['BOBİN', 'BOBIN', 'BOBIN ', 'Bobin', 'Barkod'],
                              ['ÜRÜN KODU', 'URUN KODU', 'ÜRÜN KOD', 'URUN_KODU', 'Ürün Kodu'],
                              ['ÜRÜN AÇIKLAMA', 'URUN ACIKLAMA', 'ÜRÜN AÇIKLAMA', 'ÜRÜN AÇIKLAMA'])

    uretim_meta = extract_meta(df_uretim,
                               ['ÜRETİLEN BARKOD', 'URETILEN BARKOD', 'BARKOD', 'Barkod'],
                               ['MALZEME NO', 'MALZEME_NO', 'MALZEME', 'MALZEME NO'],
                               ['MALZEME ADI', 'MALZEME_ADI', 'MALZEME ADI'])

    stok_meta = extract_meta(df_stok,
                             ['BARKOD KODU', 'Barkod Kodu', 'BARKOD', 'Barkod'],
                             ['ÜRÜN KODU', 'URUN KODU', 'Ürün Kodu'],
                             ['ÜRÜN', 'Ürün', 'ÜRÜN AÇIKLAMA'])

    cons_meta = extract_meta(df_tuketim,
                             ['GİRİŞ ÜRÜN SAP BARKODU', 'GIRIS URUN SAP BARKODU', 'Barkod', 'BARKOD'],
                             ['GİRİŞ ÜRÜN KODU', 'GIRIS URUN KODU', 'GIRIS_URUN_KODU'],
                             ['GİRİŞ ÜRÜN ACIKLAMA', 'GIRIS URUN ACIKLAMA', 'AÇIKLAMA', 'Aciklama'])

    df_raw, _ = build_single_table(prod, sayim, cons, stok, sayim_meta, uretim_meta, stok_meta, cons_meta)

    out_ts = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
    out_file = os.path.join(BASE_DIR, f"stok_kontrol_birlesik_{out_ts}.xlsx")

    # Sütun başlıklarındaki parantez içlerini kaldır
    df_write = df_raw.copy()
    df_write.columns = [re.sub(r"\s*\(.*?\)", "", c).strip() if isinstance(c, str) else c for c in df_write.columns]

    with pd.ExcelWriter(out_file, engine='openpyxl') as writer:
        df_write.to_excel(writer, sheet_name='Rapor', index=False)
        # set Excel number format for numeric columns
        workbook = writer.book
        worksheet = writer.sheets['Rapor']
        # map column names to excel col letters
        from openpyxl.utils import get_column_letter
        numeric_names = ['Sayım', 'Üretim', 'Tüketim', 'Beklenen Stok', 'Stok', 'Stok Farkı', 'Üretilmeden Tüketilen', 'Üretimden Fazla Tüketilen', 'Üretilip Kullanılmayan', 'Stokta Üretilmemiş Olan', 'Stokta Yok Ama Tüketilmiş']
        # normalize numeric name matching by checking start of header
        for idx, col in enumerate(df_write.columns, 1):
            for nk in numeric_names:
                if isinstance(col, str) and col.startswith(nk):
                    col_letter = get_column_letter(idx)
                    for row in range(2, 2 + len(df_write)):
                        cell = worksheet[f"{col_letter}{row}"]
                        cell.number_format = '#,##0'
                    break
    print(f"İşlem tamam. Tek tablo raporu: {out_file}")

if __name__ == '__main__':
    main()