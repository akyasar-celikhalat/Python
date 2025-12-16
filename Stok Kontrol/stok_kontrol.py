import os
import glob
import pandas as pd

BASE_DIR = os.path.dirname(__file__)

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

def build_single_table(prod, sayim, cons, stok):
    all_codes = set().union(prod.keys(), sayim.keys(), cons.keys(), stok.keys())
    rows = []
    for code in sorted(all_codes):
        s_amt = to_num(sayim.get(code, 0.0))
        p_amt = to_num(prod.get(code, 0.0))
        c_amt = to_num(cons.get(code, 0.0))
        stok_amt = stok.get(code)
        stok_amt_n = None if stok_amt is None else to_num(stok_amt)

        expected = s_amt + p_amt - c_amt

        # flags / computed
        uretilmeden_tuketilen = max(0.0, c_amt - (p_amt + s_amt))
        uretimden_fazla_tuketim = max(0.0, c_amt - p_amt) if c_amt > 0 else 0.0
        uretilip_kullanilmayan = max(0.0, p_amt - c_amt)
        stokta_uretilmemis_olan = stok_amt_n if (stok_amt_n is not None and p_amt <= 0 and stok_amt_n > 0) else 0.0
        stokta_yok_ama_tuketim = c_amt if (stok_amt_n is None and c_amt > 0) else 0.0
        stok_uyusmazligi = None
        if stok_amt_n is not None:
            stok_uyusmazligi = stok_amt_n - expected

        rows.append({
            'Barkod': code,
            'Sayım (başlangıç)': s_amt,
            'Üretim (METRE)': p_amt,
            'Tüketim (METRE/KG)': c_amt,
            'Beklenen Stok (Sayım+Üretim-Tüketim)': expected,
            'Stok (anlık)': stok_amt_n,
            'Stok Farkı (anlık - beklenen)': stok_uyusmazligi,
            'Üretilmeden Tüketilen': uretilmeden_tuketilen,
            'Üretimden Fazla Tüketilen': uretimden_fazla_tuketim,
            'Üretilip Kullanılmayan': uretilip_kullanilmayan,
            'Stokta Üretilmemiş Olan (miktar)': stokta_uretilmemis_olan,
            'Stokta Yok Ama Tüketilmiş': stokta_yok_ama_tuketim
        })
    df = pd.DataFrame(rows)
    # format numeric columns for Excel (rounded + binlik ayraç olarak nokta)
    numeric_cols = [
        'Sayım (başlangıç)', 'Üretim (METRE)', 'Tüketim (METRE/KG)',
        'Beklenen Stok (Sayım+Üretim-Tüketim)', 'Stok (anlık)', 'Stok Farkı (anlık - beklenen)',
        'Üretilmeden Tüketilen', 'Üretimden Fazla Tüketilen', 'Üretilip Kullanılmayan',
        'Stokta Üretilmemiş Olan (miktar)', 'Stokta Yok Ama Tüketilmiş'
    ]
    df_formatted = df.copy()
    for col in numeric_cols:
        if col in df_formatted.columns:
            df_formatted[col] = df_formatted[col].apply(lambda x: fmt_num_for_excel(x))
    return df, df_formatted  # return both raw numeric and formatted display

def main():
    files = find_files()
    # Sayım dosyasından "YARI MAMUL" sayfasını oku (yoksa ilk sayfa)
    df_sayim = safe_read(files.get('sayim'), sheet_name='YARI MAMUL')
    df_uretim = safe_read(files.get('uretim'))
    df_tuketim = safe_read(files.get('tuketim'))
    df_stok = safe_read(files.get('stok'))

    # debug: sayım dosyası okundu mu, kolonlar nedir
    if files.get('sayim'):
        if df_sayim is None:
            print(f"Uyarı: Sayım dosyası bulundu ama okunamadı: {files.get('sayim')}")
        else:
            print(f"Sayım dosyası (YARI MAMUL) okundu: {files.get('sayim')} - kolonlar: {list(df_sayim.columns)[:40]}")

    # prod_from_sayim = build_prod_from_sayim(df_sayim)
    prod_from_uretim = build_prod_from_uretim(df_uretim)
    # toplam üretim: yalnızca üretim dosyasından alınıyor (sayım artık ayrı tutuluyor)
    prod = {k: to_num(v) for k, v in prod_from_uretim.items()}

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

    df_raw, df_report = build_single_table(prod, sayim, cons, stok)

    out_ts = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
    out_file = os.path.join(BASE_DIR, f"stok_kontrol_birlesik_{out_ts}.xlsx")
    with pd.ExcelWriter(out_file, engine='openpyxl') as writer:
        # yazarken hem formatlı hem ham (isteğe göre) - ama tek sayfada formatlı göster
        df_report.to_excel(writer, sheet_name='Rapor', index=False)
    print(f"İşlem tamam. Tek tablo raporu: {out_file}")

if __name__ == '__main__':
    main()