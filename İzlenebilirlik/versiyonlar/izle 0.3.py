import pandas as pd
import os

# Sonuçları toplamak için global liste
rapor_verisi = []

def geriye_donuk_takip_listele(df, hedef_barkod, seviye=0, yol_metni="", ziyaret_edilen=None):
    if ziyaret_edilen is None:
        ziyaret_edilen = set()
    
    if hedef_barkod in ziyaret_edilen:
        return
    ziyaret_edilen.add(hedef_barkod)

    # Üretim satırlarını bul
    girdiler = df[df['TEYİT VERİLEN BARKOD'] == hedef_barkod]

    if girdiler.empty:
        return

    for _, satir in girdiler.iterrows():
        # 1. Miktar Belirleme Mantığı
        proses = str(satir['PROSES']).strip()
        
        # Kural: BD ve DH için 'TEYİT MİKTARI Kg', diğerleri için 'GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg'
        if proses in ['BD', 'DH']:
            miktar = satir.get('TEYİT MİKTARI Kg', 0)
        else:
            miktar = satir.get('GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg', 0)

        # 2. Üretim Zinciri ve Görselleştirme
        mevcut_urun_adi = str(satir['ÇIKIŞ ÜRÜN ACIKLAMA'])
        yeni_yol = f"{yol_metni} → {mevcut_urun_adi}" if yol_metni else mevcut_urun_adi
        
        # Seviye çizgilerini en başa ekleyelim
        cizgiler = "-" * (seviye + 1)
        tam_zincir = f"{cizgiler} {yeni_yol}"

        rapor_verisi.append({
            "Üretim Akışı (Hiyerarşi)": tam_zincir,
            "Proses": proses,
            "İşlem Miktarı (Kg)": miktar,
            "Üretilen Barkod": hedef_barkod,
            "Tüketilen Barkod": satir['GİRİŞ ÜRÜN SAP BARKODU'],
            "Makine No": satir['MAKİNE NO'],
            "Zaman": satir['OLUŞTURMA ZAMANI']
        })
        
        # Özyineleme: Bir alt seviyeye in
        geriye_donuk_takip_listele(
            df, 
            str(satir['GİRİŞ ÜRÜN SAP BARKODU']), 
            seviye + 1, 
            yeni_yol, 
            ziyaret_edilen.copy()
        )

# --- ANA PROGRAM ---
dosya_adi = "veri.xlsx" 

if os.path.exists(dosya_adi):
    print("📊 Veriler analiz ediliyor, lütfen bekleyin...")
    df = pd.read_excel(dosya_adi)
    
    # Sütun temizliği
    df.columns = df.columns.str.strip()
    df['TEYİT VERİLEN BARKOD'] = df['TEYİT VERİLEN BARKOD'].astype(str)
    df['GİRİŞ ÜRÜN SAP BARKODU'] = df['GİRİŞ ÜRÜN SAP BARKODU'].astype(str)

    sorgu = input("\n🔍 İzlenecek barkodu girin: ").strip()
    
    rapor_verisi.clear()
    geriye_donuk_takip_listele(df, sorgu)

    if rapor_verisi:
        rapor_df = pd.DataFrame(rapor_verisi)
        cikti_adi = f"Detayli_Uretim_Analizi_{sorgu}.xlsx"
        
        # Excel'e kaydet
        rapor_df.to_excel(cikti_adi, index=False)
        
        print(f"\n✅ Analiz tamamlandı!")
        print(f"📁 Dosya: {cikti_adi}")
        print(f"ℹ️ BD/DH prosesleri için 'Teyit Miktarı', diğerleri için 'Tüketim Miktarı' baz alınmıştır.")
    else:
        print("\n❌ Girdiğiniz barkodun üretim geçmişi bulunamadı.")
else:
    print(f"❌ '{dosya_adi}' bulunamadı!")