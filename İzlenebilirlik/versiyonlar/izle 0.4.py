import pandas as pd
import os

# Sonuçları toplamak için global liste
rapor_verisi = []

def geriye_donuk_takip_listele(df, hedef_barkod, seviye=0, yol_metni="", ziyaret_edilen=None):
    if ziyaret_edilen is None:
        ziyaret_edilen = set()
    
    # Barkodu temizle
    hedef_barkod = str(hedef_barkod).strip()
    
    if hedef_barkod in ziyaret_edilen or hedef_barkod == "nan":
        return
    ziyaret_edilen.add(hedef_barkod)

    # Üretim satırlarını bul (Çıktı barkodu üzerinden)
    girdiler = df[df['TEYİT VERİLEN BARKOD'].str.strip() == hedef_barkod]

    if girdiler.empty:
        return

    for _, satir in girdiler.iterrows():
        # 1. Miktar Belirleme Mantığı
        proses = str(satir['PROSES']).strip()
        if proses in ['BD', 'DH']:
            miktar = satir.get('TEYİT MİKTARI Kg', 0)
        else:
            miktar = satir.get('GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg', 0)

        # 2. Üretim Zinciri ve Görselleştirme
        mevcut_urun_adi = str(satir['ÇIKIŞ ÜRÜN ACIKLAMA']).strip()
        yeni_yol = f"{yol_metni} → {mevcut_urun_adi}" if yol_metni else mevcut_urun_adi
        
        # Seviye çizgilerini en başa ekleyelim
        cizgiler = "-" * (seviye + 1)
        tam_zincir = f"{cizgiler} {yeni_yol}"

        # 3. İstenen Sütun Dizilimi
        rapor_verisi.append({
            "Üretilen Barkod": hedef_barkod,                   # 1. SÜTUN
            "ÇIKIŞ ÜRÜN ACIKLAMA": mevcut_urun_adi,            # 2. SÜTUN
            "Üretim Akışı (Hiyerarşi)": tam_zincir,
            "İşlem Miktarı (Kg)": miktar,
            "Tüketilen Barkod": str(satir['GİRİŞ ÜRÜN SAP BARKODU']).strip(),
            "Proses": proses,
            "Makine No": satir['MAKİNE NO'],
            "Zaman": satir['OLUŞTURMA ZAMANI']
        })
        
        # Özyineleme: Bir alt seviyeye in
        geriye_donuk_takip_listele(
            df, 
            satir['GİRİŞ ÜRÜN SAP BARKODU'], 
            seviye + 1, 
            yeni_yol, 
            ziyaret_edilen.copy()
        )

# --- ANA PROGRAM ---
dosya_adi = "veri.xlsx" 

if os.path.exists(dosya_adi):
    print("📊 Veriler yükleniyor...")
    df = pd.read_excel(dosya_adi)
    
    # Sütun isimlerini ve veri tiplerini temizle
    df.columns = df.columns.str.strip()
    df['TEYİT VERİLEN BARKOD'] = df['TEYİT VERİLEN BARKOD'].astype(str)
    df['GİRİŞ ÜRÜN SAP BARKODU'] = df['GİRİŞ ÜRÜN SAP BARKODU'].astype(str)

    sorgu = input("\n🔍 İzlenecek barkodu girin: ").strip()
    
    rapor_verisi.clear()
    geriye_donuk_takip_listele(df, sorgu)

    if rapor_verisi:
        rapor_df = pd.DataFrame(rapor_verisi)
        cikti_adi = f"Izlenebilirlik_Raporu_{sorgu}.xlsx"
        rapor_df.to_excel(cikti_adi, index=False)
        
        print(f"\n✅ Başarılı! {cikti_adi} dosyası oluşturuldu.")
    else:
        print(f"\n❌ '{sorgu}' barkodu için hiyerarşi bulunamadı.")
        print("İpucu: Barkodun Excel'de 'TEYİT VERİLEN BARKOD' sütununda olduğundan emin olun.")
else:
    print(f"❌ '{dosya_adi}' bulunamadı!")