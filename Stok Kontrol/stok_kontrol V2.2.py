import os
import glob
import re
import pandas as pd
import numpy as np

BASE_DIR = os.path.dirname(__file__)
FILL_EMPTY_STOCK = True

# --- YARDIMCI FONKSİYONLAR ---

def safe_read(path, sheet_name=None):
    if not path: return None
    try:
        # Dosyayı oku
        data = pd.read_excel(path, sheet_name=sheet_name, engine='openpyxl')
        
        # Eğer sonuç bir sözlükse (birden fazla sayfa varsa) ilk sayfayı al
        if isinstance(data, dict):
            if not data: return None # Dosya boşsa
            return next(iter(data.values()))
        
        return data
    except Exception as e:
        print(f"Uyarı: {path} okunamadı -> {e}")
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
    if key is None or pd.isna(key): return
    # Senkronizasyon: Barkodu temiz metne çevir
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
    if abs(diff) < 1e-9 and (s + p + c) > 0: return "Tam Mutabakat"
    if p == 0 and c > 0 and s == 0: return "Girişsiz Tüketim"
    if s > 0 and c == 0 and stok == 0: return "Sayım Stok"
    if stok > 0 and s == 0 and p == 0: return "Sanal Stok (Fizikte Yok)"
    return "Stok Açığı (Eksik)" if diff < 0 else "Stok Fazlası (Artı)"

def get_priority_by_cap(row):
    desc = str(row['Malzeme Açıklama'] or "").strip()
    # Muafiyet: Rakamla başlıyorsa veya "H " ile başlıyorsa analiz dışı
    if desc and (desc[0].isdigit() or desc.upper().startswith("H ")):
        return "5-ANALİZ DIŞI (AMBAR/YARDIMCI)"

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

def build_single_table(prod, sayim, cons, stok, s_meta, u_meta, st_meta, c_meta):
    all_codes = set().union(prod.keys(), sayim.keys(), cons.keys(), stok.keys())
    rows = []
    for code in sorted(all_codes):
        # KRİTİK FİLTRE: Tire içermeyen veya M ile başlayanları atla
        if '-' not in str(code) or str(code).startswith('M'): 
            continue
        
        s_amt = sayim.get(code, 0.0)
        p_amt = prod.get(code, 0.0)
        c_amt = cons.get(code, 0.0)
        stok_amt_n = stok.get(code) if code in stok else (0.0 if FILL_EMPTY_STOCK else None)
        
        expected = s_amt + p_amt - c_amt
        stok_farki = (stok_amt_n if stok_amt_n is not None else 0) - expected

        # Meta verileri eşleştir
        m_kodu = code
        m_aciklama = s_meta.get(code) or u_meta.get(code) or st_meta.get(code) or c_meta.get(code)

        rows.append({
            'Barkod': code, 'Malzeme Açıklama': m_aciklama,
            'Sayım': s_amt, 'Üretim': p_amt, 'Tüketim': c_amt,
            'Beklenen Stok': expected, 'Stok': stok_amt_n, 'Stok Farkı': stok_farki
        })
    
    df = pd.DataFrame(rows)
    if not df.empty:
        df['Durum Tespiti'] = df.apply(get_row_status, axis=1)
        df['Önem Derecesi'] = df.apply(get_priority_by_cap, axis=1)
    return df

def main():
    # Dosyaları bul
    files = {
        'sayim': glob.glob(os.path.join(BASE_DIR, '*sayım*.xls*')) + glob.glob(os.path.join(BASE_DIR, '*sayim*.xls*')),
        'uretim': glob.glob(os.path.join(BASE_DIR, '*üretim*.xls*')) + glob.glob(os.path.join(BASE_DIR, '*uretim*.xls*')),
        'tuketim': glob.glob(os.path.join(BASE_DIR, '*tüketim*.xls*')) + glob.glob(os.path.join(BASE_DIR, '*tuketim*.xls*')),
        'stok': glob.glob(os.path.join(BASE_DIR, '*stok*.xls*'))
    }

    # Dosyaları oku (İlk bulunan dosyayı al)
    # main fonksiyonu içindeki ilgili satır:
    df_s = safe_read(next(iter(files['sayim']), None), sheet_name='YARI MAMUL')
    df_u = safe_read(next(iter(files['uretim']), None))
    df_t = safe_read(next(iter(files['tuketim']), None))
    df_st = safe_read(next(iter(files['stok']), None))

    # 1. SAYIM İŞLEME
    sayim_map, s_meta = {}, {}
    if df_s is not None:
        for _, r in df_s.iterrows():
            b = str(r.get('BOBİN') or '').strip()
            if b:
                aggregate_add(sayim_map, b, r.get('METRE'))
                s_meta[b] = r.get('ÜRÜN AÇIKLAMA')

    # 2. ÜRETİM İŞLEME
    prod_map, u_meta = {}, {}
    if df_u is not None:
        for _, r in df_u.iterrows():
            b = str(r.get('ÜRETİLEN BARKOD') or '').strip()
            if b:
                aggregate_add(prod_map, b, r.get('METRE_02'))
                u_meta[b] = r.get('MALZEME ADI')

    # 3. STOK İŞLEME
    stok_map, st_meta = {}, {}
    if df_st is not None:
        for _, r in df_st.iterrows():
            b = str(r.get('BARKOD KODU') or '').strip()
            if b:
                aggregate_add(stok_map, b, r.get('MİKTAR'))
                st_meta[b] = r.get('ÜRÜN')

    # 4. TÜKETİM İŞLEME (KOŞULLU MANTIK)
    cons_map, c_meta = {}, {}
    if df_t is not None:
        for _, r in df_t.iterrows():
            b = str(r.get('GİRİŞ ÜRÜN SAP BARKODU') or '').strip()
            if not b: continue
            
            # Koşul: ÇIKIŞ ÜRÜN ACIKLAMA kontrolü
            cikis_desc = str(r.get('ÇIKIŞ ÜRÜN ACIKLAMA') or "").strip()
            if cikis_desc.startswith("H ") or cikis_desc.startswith("DMT"):
                miktar = r.get('TEYİT MİKTARI Kg')
            else:
                miktar = r.get('GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg')
            
            aggregate_add(cons_map, b, miktar)
            c_meta[b] = r.get('GİRİŞ ÜRÜN ACIKLAMA')

    # Birleştirme ve Rapor Oluşturma
    df_final = build_single_table(prod_map, sayim_map, cons_map, stok_map, s_meta, u_meta, st_meta, c_meta)

    # Excel Yazımı
    if not df_final.empty:
        out_file = os.path.join(BASE_DIR, f"analiz_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        with pd.ExcelWriter(out_file, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='Detayli_Rapor', index=False)
            ozet = df_final.groupby(['Önem Derecesi', 'Durum Tespiti']).size().unstack(fill_value=0)
            ozet.to_excel(writer, sheet_name='Analiz_Ozeti')
        print(f"Rapor hazır: {out_file}")
    else:
        print("Uyarı: Analiz edilecek veri bulunamadı. Filtreleri veya sütun isimlerini kontrol edin.")

if __name__ == '__main__':
    main()