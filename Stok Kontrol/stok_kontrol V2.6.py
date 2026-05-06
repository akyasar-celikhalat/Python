import os
import glob
import re
import pandas as pd
import numpy as np
from openpyxl.styles import Alignment

BASE_DIR = os.path.dirname(__file__)
FILL_EMPTY_STOCK = True

# --- YARDIMCI FONKSİYONLAR ---

def safe_read(path, sheet_name=None):
    if not path: return None
    try:
        # engine='openpyxl' ile oku
        data = pd.read_excel(path, sheet_name=sheet_name, engine='openpyxl')
        if isinstance(data, dict):
            if not data: return None
            df = next(iter(data.values()))
        else:
            df = data
        
        # Sütun isimlerindeki boşlukları temizle (En kritik adım)
        if df is not None:
            df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception as e:
        print(f"Uyarı: {path} okunamadı -> {e}")
        return None

def clean_material_code(code):
    """Malzeme kodlarını temizler, .0 kısmını atar ve string yapar."""
    if pd.isna(code) or code == "": return None
    s = str(code).strip()
    if s.endswith('.0'): s = s[:-2]
    return s

def to_num(x):
    try:
        if pd.isna(x): return 0.0
        if isinstance(x, (int, float)): return float(x)
        s = str(x).replace(',', '.')
        filtered = ''.join(ch for ch in s if ch.isdigit() or ch == '.')
        return float(filtered) if filtered not in ('', '.') else 0.0
    except: return 0.0

def aggregate_add(d, key, amt):
    if key is None or pd.isna(key): return
    k = str(key).strip()
    if k == '': return
    d[k] = d.get(k, 0.0) + to_num(amt)

def extract_cap(text):
    if not text or pd.isna(text): return None
    match = re.search(r"(\d+[.,]\d+|\d+)", str(text))
    if match:
        return float(match.group(1).replace(',', '.'))
    return None

def get_row_status(row):
    s, p, c = row['Sayım'], row['Üretim'], row['Tüketim']
    stok, diff = row['Stok'], row['Stok Farkı']
    if abs(diff) < 1e-5: return "Tam Mutabakat"
    if p == 0 and c > 0 and s == 0: return "Girişsiz Tüketim"
    if s > 0 and c == 0 and stok == 0: return "Sayım Stok"
    if stok > 0 and s == 0 and p == 0: return "Sanal Stok (Fizikte Yok)"
    return "Stok Açığı (Eksik)" if diff < 0 else "Stok Fazlası (Artı)"

def get_priority_by_cap(row):
    desc = str(row['Malzeme Açıklama'] or "").strip()
    if desc and (desc[0].isdigit() or desc.upper().startswith("H ")):
        return "5-ANALİZ DIŞI (AMBAR/YARDIMCI)"
    if row['Durum Tespiti'] == "Tam Mutabakat": return "0-MUTABAKAT"
    cap = extract_cap(desc) 
    metraj_farki = abs(row['Stok Farkı'])
    if cap is None or cap == 0: return "4-BİLİNMEYEN ÇAP"
    kg_farki = (cap**2) * 0.006165 * metraj_farki
    ref_kg = 400.0 if cap <= 2.3 else (850.0 if cap <= 5.5 else 1500.0)
    yuzde_sapma = (kg_farki / ref_kg) * 100
    if yuzde_sapma < 20: return "3-NORMAL"
    elif 20 <= yuzde_sapma <= 60: return "2-YÜKSEK"
    else: return "1-KRİTİK"

# --- ANA ANALİZ SÜRECİ ---

def build_single_table(prod, sayim, cons, stok, eklen, silin, meta_desc, meta_code):
    all_codes = set().union(prod.keys(), sayim.keys(), cons.keys(), stok.keys(), eklen.keys(), silin.keys())
    rows = []
    for barkod in sorted(all_codes):
        # Filtre: Barkod formatı kontrolü
        if '-' not in str(barkod) or str(barkod).startswith('M'): 
            continue
        
        s_amt = sayim.get(barkod, 0.0)
        p_amt = prod.get(barkod, 0.0)
        c_amt = cons.get(barkod, 0.0)
        e_amt = eklen.get(barkod, 0.0)
        si_amt = silin.get(barkod, 0.0)
        stok_amt_n = stok.get(barkod) if barkod in stok else (0.0 if FILL_EMPTY_STOCK else None)
        
        expected = s_amt + p_amt - c_amt + e_amt - si_amt
        stok_farki = (stok_amt_n if stok_amt_n is not None else 0) - expected

        m_aciklama = meta_desc.get(barkod, "")
        m_kodu = meta_code.get(barkod, "")

        rows.append({
            'Barkod': barkod, 
            'Malzeme Kodu': str(m_kodu) if m_kodu else "", 
            'Malzeme Açıklama': m_aciklama,
            'Sayım': s_amt, 'Üretim': p_amt, 'Tüketim': c_amt,
            'Eklenen': e_amt, 'Silinen': si_amt,
            'Beklenen Stok': expected, 'Stok': stok_amt_n, 'Stok Farkı': stok_farki
        })
    
    df = pd.DataFrame(rows)
    if not df.empty:
        df['Durum Tespiti'] = df.apply(get_row_status, axis=1)
        df['Önem Derecesi'] = df.apply(get_priority_by_cap, axis=1)
    return df

def main():
    files = {
        'sayim': glob.glob(os.path.join(BASE_DIR, '*sayım*.xls*')) + glob.glob(os.path.join(BASE_DIR, '*sayim*.xls*')),
        'uretim': glob.glob(os.path.join(BASE_DIR, '*üretim*.xls*')) + glob.glob(os.path.join(BASE_DIR, '*uretim*.xls*')),
        'tuketim': glob.glob(os.path.join(BASE_DIR, '*tüketim*.xls*')) + glob.glob(os.path.join(BASE_DIR, '*tuketim*.xls*')),
        'stok': glob.glob(os.path.join(BASE_DIR, '*stok*.xls*')),
        'eklenen': glob.glob(os.path.join(BASE_DIR, '*eklenen*.xls*')),
        'silinen': glob.glob(os.path.join(BASE_DIR, '*silinen*.xls*'))
    }

    df_s = safe_read(next(iter(files['sayim']), None), sheet_name='YARI MAMUL')
    df_u = safe_read(next(iter(files['uretim']), None))
    df_t = safe_read(next(iter(files['tuketim']), None))
    df_st = safe_read(next(iter(files['stok']), None))
    df_ek = safe_read(next(iter(files['eklenen']), None))
    df_si = safe_read(next(iter(files['silinen']), None))

    meta_desc, meta_code = {}, {}

    def update_meta(b, code, desc):
        b = str(b).strip()
        if not b or b == 'nan': return
        
        # Kodu temizle (4000.0 -> 4000)
        cleaned_code = clean_material_code(code)
        
        if b not in meta_desc or not meta_desc[b]: 
            meta_desc[b] = str(desc).strip() if not pd.isna(desc) else None
        
        if b not in meta_code or not meta_code[b]: 
            meta_code[b] = cleaned_code

    # --- VERİ İŞLEME ---
    
    # 1. Sayım Verileri (Malzeme kodu eksikliği burada çözülüyor)
    sayim_map = {}
    if df_s is not None:
        for _, r in df_s.iterrows():
            b = r.get('BOBİN')
            if b:
                aggregate_add(sayim_map, b, r.get('METRE'))
                update_meta(b, r.get('ÜRÜN KODU'), r.get('ÜRÜN AÇIKLAMA'))

    # 2. Üretim Verileri
    prod_map = {}
    if df_u is not None:
        for _, r in df_u.iterrows():
            b = r.get('ÜRETİLEN BARKOD')
            if b:
                aggregate_add(prod_map, b, r.get('METRE_02'))
                update_meta(b, r.get('MALZEME NO'), r.get('MALZEME ADI'))

    # 3. Stok Verileri
    stok_map = {}
    if df_st is not None:
        for _, r in df_st.iterrows():
            b = r.get('BARKOD KODU') or r.get('BARKOD')
            if b:
                aggregate_add(stok_map, b, r.get('MİKTAR'))
                update_meta(b, r.get('ÜRÜN KODU'), r.get('ÜRÜN'))

    # 4. Eklenen/Silinen
    eklen_map = {}
    if df_ek is not None:
        for _, r in df_ek.iterrows():
            b = r.get('BARKOD NUMARASI')
            if b:
                aggregate_add(eklen_map, b, r.get('MİKTAR'))
                update_meta(b, r.get('ÜRÜN KODU / ÜRÜN TANIMI'), r.get('ÜRÜN KODU / ÜRÜN TANIMI'))

    silin_map = {}
    if df_si is not None:
        for _, r in df_si.iterrows():
            b = r.get('BARKOD KODU') or r.get('BARKOD')
            if b:
                aggregate_add(silin_map, b, r.get('MİKTAR'))
                update_meta(b, r.get('ÜRÜN KODU'), r.get('ÜRÜN'))

    # 5. Tüketim Verileri
    cons_map = {}
    if df_t is not None:
        for _, r in df_t.iterrows():
            b = r.get('GİRİŞ ÜRÜN SAP BARKODU')
            if not b: continue
            cikis_desc = str(r.get('ÇIKIŞ ÜRÜN ACIKLAMA') or "").strip()
            miktar = r.get('TEYİT MİKTARI Kg') if (cikis_desc.startswith("H ") or cikis_desc.startswith("DMT")) else r.get('GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg')
            aggregate_add(cons_map, b, miktar)
            update_meta(b, r.get('GİRİŞ ÜRÜN KODU'), r.get('GİRİŞ ÜRÜN ACIKLAMA'))

    # Raporu Oluştur
    df_final = build_single_table(prod_map, sayim_map, cons_map, stok_map, eklen_map, silin_map, meta_desc, meta_code)

    if not df_final.empty:
        # Sayısal değerleri yuvarla
        num_cols = ['Sayım', 'Üretim', 'Tüketim', 'Eklenen', 'Silinen', 'Beklenen Stok', 'Stok', 'Stok Farkı']
        for col in num_cols:
            if col in df_final.columns:
                df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(0).round(0).astype(int)

        out_file = os.path.join(BASE_DIR, f"analiz_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        
        with pd.ExcelWriter(out_file, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='Detayli_Rapor', index=False)
            
            workbook = writer.book
            worksheet = writer.sheets['Detayli_Rapor']
            thousand_format = '#,##0'
            
            for col_idx, col_name in enumerate(df_final.columns, 1):
                # Sayısal sütunlara binlik ayıraç uygula
                if col_name in num_cols:
                    for row_idx in range(2, len(df_final) + 2):
                        cell = worksheet.cell(row=row_idx, column=col_idx)
                        cell.number_format = thousand_format
                        cell.alignment = Alignment(horizontal='right')
                
                # Malzeme Kodu sütununu metin olarak ayarla
                if col_name == 'Malzeme Kodu' or col_name == 'Barkod':
                    for row_idx in range(2, len(df_final) + 2):
                        cell = worksheet.cell(row=row_idx, column=col_idx)
                        cell.alignment = Alignment(horizontal='left')

        print(f"Rapor hazır: {out_file}")
    else:
        print("Uyarı: Analiz edilecek veri bulunamadı.")

if __name__ == '__main__':
    main()