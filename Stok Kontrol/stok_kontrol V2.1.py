import os
import glob
import re
import pandas as pd
import numpy as np

BASE_DIR = os.path.dirname(__file__)
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
        if name.startswith('~$'): continue
        lname = name.lower()
        for k, keys in KEYS.items():
            if any(key in lname for key in keys):
                files[k] = path
    return files

def safe_read(path, sheet_name=None):
    if not path: return None
    try:
        if sheet_name:
            try:
                return pd.read_excel(path, sheet_name=sheet_name, engine='openpyxl')
            except ValueError:
                sheets = pd.read_excel(path, sheet_name=None, engine='openpyxl')
                for k, df in sheets.items():
                    if k and k.casefold() == sheet_name.casefold(): return df
                return next(iter(sheets.values()))
        return pd.read_excel(path, engine='openpyxl')
    except Exception as e:
        print(f"Uyarı: Okuma hatası {path} -> {e}")
        return None

def find_col(df, names):
    if df is None: return None
    cols = list(df.columns)
    low_map = {c.casefold(): c for c in cols}
    for n in names:
        if n in cols: return n
        if n.casefold() in low_map: return low_map[n.casefold()]
    for ln, orig in low_map.items():
        for n in names:
            if n.casefold() in ln or ln in n.casefold(): return orig
    return None

def to_num(x):
    try:
        if pd.isna(x): return 0.0
        if isinstance(x, (int, float)): return float(x)
        s = str(x).replace(',', '.')
        filtered = ''.join(ch for ch in s if ch.isdigit() or ch == '.')
        return float(filtered) if filtered not in ('', '.') else 0.0
    except: return 0.0

def aggregate_add(d, key, amt):
    if key is None: return
    k = str(key).strip()
    if k == '': return
    d[k] = d.get(k, 0.0) + to_num(amt)

# --- ANALİZ FONKSİYONLARI ---

def extract_cap(text):
    """Metin içerisinden çap bilgisini (ör: 2.30) ayıklar."""
    if not text or pd.isna(text): return None
    # Sayı+nokta/virgül+sayı formatını ara (Örn: 2,3 veya 5.5)
    match = re.search(r"(\d+[.,]\d+|\d+)", str(text))
    if match:
        return float(match.group(1).replace(',', '.'))
    return None

def get_row_status(row):
    s, p, c = row['Sayım'], row['Üretim'], row['Tüketim']
    stok, diff = row['Stok'], row['Stok Farkı']
    if abs(diff) < 1e-9 and (s + p + c) > 0: return "Tam Mutabakat"
    if p == 0 and c > 0 and s == 0: return "Girişsiz Tüketim"
    if s > 0 and stok == 0: return "Kayıt Dışı Fiziksel Stok"
    if stok > 0 and s == 0 and p == 0: return "Sanal Stok (Fizikte Yok)"
    return "Stok Açığı (Eksik)" if diff < 0 else "Stok Fazlası (Artı)"

def get_priority_by_cap(row):
    # Malzeme açıklaması ve kodu
    desc = str(row['Malzeme Açıklama'] or "").strip()
    
    # 1. MUAFİYET KONTROLÜ: İlk karakter rakam VEYA "H " ile başlıyorsa hesaba katma
    # startswith("H ") sayesinde "HT" gibi bitişik yazılan kodlar analize devam eder.
    if desc and (desc[0].isdigit() or desc.upper().startswith("H ")):
        return "5-ANALİZ DIŞI (AMBAR/YARDIMCI)"

    # 2. Çap Ayıklama
    cap = extract_cap(desc) 
    metraj_farki = abs(row['Stok Farkı'])
    
    if cap is None or cap == 0:
        return "4-BİLİNMEYEN ÇAP"
    
    # 3. Kilo Hesaplama (Sadeleştirilmiş Hassas Formül)
    # kg = cap^2 * 0.006165 * metre
    kg_farki = (cap**2) * 0.006165 * metraj_farki
    
    # 4. Çapa göre Referans Kilo Seçimi
    if 0 < cap <= 2.3:
        ref_kg = 400.0
    elif 2.3 < cap <= 5.5:
        ref_kg = 850.0
    else:
        ref_kg = 1500.0

    # 5. Yüzde Sapma ve Önem Derecesi
    yuzde_sapma = (kg_farki / ref_kg) * 100
    
    if yuzde_sapma < 20:
        return "3-NORMAL"
    elif 20 <= yuzde_sapma <= 60:
        return "2-YÜKSEK"
    else:
        return "1-KRİTİK"

# --- ANA SÜREÇ ---

def build_single_table(prod, sayim, cons, stok, sayim_meta, uretim_meta, stok_meta, cons_meta):
    all_codes = set().union(prod.keys(), sayim.keys(), cons.keys(), stok.keys())
    rows = []
    for code in sorted(all_codes):
        if '-' not in str(code) or str(code).startswith('M'): continue
        
        s_amt = to_num(sayim.get(code, 0.0))
        p_amt = to_num(prod.get(code, 0.0))
        c_amt = to_num(cons.get(code, 0.0))
        stok_amt_n = to_num(stok.get(code)) if code in stok else (0.0 if FILL_EMPTY_STOCK else None)
        
        expected = s_amt + p_amt - c_amt
        stok_farki = (stok_amt_n if stok_amt_n is not None else 0) - expected

        m_kodu = sayim_meta[0].get(code) or uretim_meta[0].get(code) or stok_meta[0].get(code) or cons_meta[0].get(code)
        m_aciklama = sayim_meta[1].get(code) or uretim_meta[1].get(code) or stok_meta[1].get(code) or cons_meta[1].get(code)

        rows.append({
            'Barkod': code, 'Malzeme Kodu': m_kodu, 'Malzeme Açıklama': m_aciklama,
            'Sayım': s_amt, 'Üretim': p_amt, 'Tüketim': c_amt,
            'Beklenen Stok': expected, 'Stok': stok_amt_n, 'Stok Farkı': stok_farki
        })
    
    df = pd.DataFrame(rows)
    df['Durum Tespiti'] = df.apply(get_row_status, axis=1)
    df['Önem Derecesi'] = df.apply(get_priority_by_cap, axis=1)
    
    return df

def main():
    files = find_files()
    df_s = safe_read(files.get('sayim'), sheet_name='YARI MAMUL')
    df_u = safe_read(files.get('uretim'))
    df_t = safe_read(files.get('tuketim'))
    df_st = safe_read(files.get('stok'))

    # Veri birleştirme ve meta çıkarma (Önceki mantıkla aynı)
    prod = {}
    if df_u is not None:
        c_b = find_col(df_u, ['BARKOD'])
        c_m = find_col(df_u, ['METRE_02', 'METRE'])
        for _, r in df_u.iterrows(): aggregate_add(prod, r.get(c_b), r.get(c_m))

    sayim = {}
    if df_s is not None:
        c_b = find_col(df_s, ['BOBİN', 'Barkod'])
        c_m = find_col(df_s, ['METRE', 'Metre'])
        for _, r in df_s.iterrows(): aggregate_add(sayim, r.get(c_b), r.get(c_m))

    # Tüketim ve Stok okuma...
    def extract_meta(df, bar_c, cod_c, des_c):
        c_m, d_m = {}, {}
        if df is None: return c_m, d_m
        cb, cc, cd = find_col(df, bar_c), find_col(df, cod_c), find_col(df, des_c)
        for _, r in df.iterrows():
            b = str(r.get(cb) or '').strip()
            if b:
                if cc: c_m[b] = str(r.get(cc) or '').strip()
                if cd: d_m[b] = str(r.get(cd) or '').strip()
        return c_m, d_m

    s_meta = extract_meta(df_s, ['BOBIN', 'Barkod'], ['ÜRÜN KODU'], ['ÜRÜN AÇIKLAMA'])
    u_meta = extract_meta(df_u, ['BARKOD'], ['MALZEME NO'], ['MALZEME ADI'])
    st_meta = extract_meta(df_st, ['BARKOD'], ['ÜRÜN KODU'], ['ÜRÜN'])
    c_meta = extract_meta(df_t, ['BARKOD'], ['GİRİŞ ÜRÜN KODU'], ['GİRİŞ ÜRÜN ACIKLAMA'])

    """cons = {} # Tüketim detayları build_cons_from_tuketim mantığıyla...
    # (Hızlıca basit tüketim toplama)
    if df_t is not None:
        cb = find_col(df_t, ['BARKOD'])
        cm = find_col(df_t, ['GİRİŞ ÜRÜN TÜKETİM MİKTARI', 'TEYİT MİKTARI'])
        for _, r in df_t.iterrows(): aggregate_add(cons, r.get(cb), r.get(cm))"""

    # --- Tüketim Hesaplama (Güncellenmiş Koşullu Mantık) ---
    cons = {}
    if df_t is not None:
        cb = find_col(df_t, ['BARKOD'])
        # Koşul için gerekli sütunları bulalım
        c_aciklama = find_col(df_t, ['ÇIKIŞ ÜRÜN ACIKLAMA'])
        c_teyit_kg = find_col(df_t, ['TEYİT MİKTARI Kg'])
        c_giris_kg = find_col(df_t, ['GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg'])
        
        for _, r in df_t.iterrows():
            # Açıklama metnini al (Boşsa boş string yap)
            aciklama = str(r.get(c_aciklama) or "").strip()
            
            # KOŞUL: "H " (boşluklu) veya "DMT" ile başlıyorsa
            if aciklama.startswith("H ") or aciklama.startswith("DMT"):
                miktar = r.get(c_teyit_kg)
            else:
                miktar = r.get(c_giris_kg)
                
            # Belirlenen miktarı barkod bazında topla
            aggregate_add(cons, r.get(cb), miktar)
            
    stok_map = {}
    if df_st is not None:
        cb = find_col(df_st, ['BARKOD'])
        cm = find_col(df_st, ['MİKTAR'])
        for _, r in df_st.iterrows(): aggregate_add(stok_map, r.get(cb), r.get(cm))

    df_final = build_single_table(prod, sayim, cons, stok_map, s_meta, u_meta, st_meta, c_meta)

    # Excel Yazımı
    out_file = os.path.join(BASE_DIR, f"stok_analiz_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    with pd.ExcelWriter(out_file, engine='openpyxl') as writer:
        df_final.to_excel(writer, sheet_name='Detayli_Rapor', index=False)
        # Özet sayfası
        ozet = df_final.groupby(['Önem Derecesi', 'Durum Tespiti']).size().unstack(fill_value=0)
        ozet.to_excel(writer, sheet_name='Analiz_Ozeti')

    print(f"Rapor hazır: {out_file}")
    print("\nÖnem Derecesine Göre Hata Dağılımı:")
    print(df_final['Önem Derecesi'].value_counts())

if __name__ == '__main__':
    main()