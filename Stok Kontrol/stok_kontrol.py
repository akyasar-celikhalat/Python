import os
import glob
import pandas as pd

# Klasör yolunu script'in bulunduğu klasör olarak al (veya sabitleyin)
BASE_DIR = os.path.dirname(__file__)

# Dosya adlarında aranacak anahtarlar (küçük harfe çevirilerek kontrol edilir)
KEYS = {
    'sayim': ['sayim', 'sayım', 'count'],
    'uretim': ['uretim', 'üretim', 'production'],
    'tuketim': ['tuketim', 'tüketim', 'consumption'],
    'stok': ['stok', 'stock']
}

CANDIDATE_BARCODE_COLS = [
    'Barkod', 'barkod', 'GİRİŞ ÜRÜN SAP BARKODU', 'TEYİT VERİLEN BARKOD',
    'SAP ETİKET BARKODU', 'SAP ETIKET BARKODU', 'Ürün Kodu', 'ÜRÜN KODU'
]

CANDIDATE_QTY_COLS = [
    'Miktar', 'MIKTAR', 'TÜKETİM', 'Tuketim', 'GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg',
    'TEYİT MİKTARI Metre', 'Tüketim (Kg)', 'Adet', 'METRE'
]

def find_file_for_key():
    files = glob.glob(os.path.join(BASE_DIR, '*.xls*'))
    found = {}
    for f in files:
        name = os.path.basename(f)
        # Geçici Office dosyalarını atla
        if name.startswith('~$'):
            continue
        lname = name.lower()
        for k, keys in KEYS.items():
            if any(k0 in lname for k0 in keys):
                found[k] = f
    return found

def detect_col(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    # daha esnek eşleşme
    cols_lower = {col.lower(): col for col in df.columns}
    for cand in candidates:
        if cand.lower() in cols_lower:
            return cols_lower[cand.lower()]
    return None

def load_agg_qty(path, barcode_col=None, qty_col=None, default_count=False):
    try:
        df = pd.read_excel(path, engine='openpyxl')
    except PermissionError:
        print(f"Uyarı: Dosya açılamadı (izin/başka uygulama tarafından kilitli), atlanıyor: {path}")
        return {}, None, None
    except Exception as e:
        print(f"Uyarı: Dosya okunurken hata oluştu, atlanıyor: {path} -> {e}")
        return {}, None, None

    if barcode_col is None:
        barcode_col = detect_col(df, CANDIDATE_BARCODE_COLS) or df.columns[0]
    if qty_col is None:
        qty_col = detect_col(df, CANDIDATE_QTY_COLS)
    if qty_col:
        s = df.groupby(df[barcode_col].astype(str))[qty_col].sum()
    else:
        if default_count:
            s = df.groupby(df[barcode_col].astype(str)).size()
        else:
            s = pd.Series(dtype=float)
    s.index = s.index.astype(str)
    return s.to_dict(), barcode_col, qty_col

def main():
    files = find_file_for_key()
    if not files:
        print("Uyarı: Sayım/Üretim/Tüketim/Stok dosyaları bulunamadı.")
        return

    # Load aggregates (varsayılan: üretim/tüketim miktarı sütunu yoksa satır say)
    prod, prod_bar, prod_qty = ({}, None, None)
    cons, cons_bar, cons_qty = ({}, None, None)
    sayim, sayim_bar, sayim_qty = ({}, None, None)
    stok, stok_bar, stok_qty = ({}, None, None)

    if 'uretim' in files:
        prod, prod_bar, prod_qty = load_agg_qty(files['uretim'], default_count=True)
    if 'tuketim' in files:
        cons, cons_bar, cons_qty = load_agg_qty(files['tuketim'], default_count=True)
    if 'sayim' in files:
        sayim, sayim_bar, sayim_qty = load_agg_qty(files['sayim'], default_count=False)
    if 'stok' in files:
        stok, stok_bar, stok_qty = load_agg_qty(files['stok'], default_count=False)

    # Tüm barkodları birleştir
    all_codes = set().union(prod.keys(), cons.keys(), sayim.keys(), stok.keys())

    rows = []
    for code in sorted(all_codes):
        initial = float(sayim.get(code, 0))
        produced = float(prod.get(code, 0))
        consumed = float(cons.get(code, 0))
        current = stok.get(code, None)
        expected = initial + produced - consumed
        stok_mismatch = None
        if current is not None:
            try:
                stok_mismatch = float(current) - expected
            except:
                stok_mismatch = None

        issue = ''
        if consumed > initial + produced + 1e-6:
            issue += 'Tüketim > (Sayım+Üretim) ; '  # üretilmeden tüketilmiş olabilir
        if produced > consumed + 1e-6:
            issue += 'Üretilip kullanılmayan stok var ; '
        if stok_mismatch is not None and abs(stok_mismatch) > 1e-6:
            issue += 'Stok uyuşmazlığı ; '

        rows.append({
            'Barkod': code,
            'Sayım (başlangıç)': initial,
            'Üretim': produced,
            'Tüketim': consumed,
            'Stok (anlık)': current,
            'Beklenen Stok (sayım+üretim-tüketim)': expected,
            'Stok Farkı (anlık - beklenen)': stok_mismatch,
            'Uyarı/İşlem': issue.strip()
        })

    df_out = pd.DataFrame(rows)
    timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
    out_file = os.path.join(BASE_DIR, f"stok_kontrol_sonuc_{timestamp}.xlsx")
    df_out.to_excel(out_file, index=False, engine='openpyxl')
    print(f"İşlem tamam. Sonuçlar: {out_file}")

if __name__ == '__main__':
    main()