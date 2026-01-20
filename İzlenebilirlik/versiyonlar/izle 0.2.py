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

    # Üretim satırlarını bul (Çıktı barkodu üzerinden)
    girdiler = df[df['TEYİT VERİLEN BARKOD'] == hedef_barkod]

    if girdiler.empty:
        return

    for _, satir in girdiler.iterrows():
        # Mevcut ürünün açıklaması
        mevcut_aciklama = str(satir['ÇIKIŞ ÜRÜN ACIKLAMA'])
        
        # Zinciri oluştur (Üst seviye açıklamalarını biriktirerek)
        yeni_yol = f"{yol_metni} > {mevcut_aciklama}" if yol_metni else mevcut_aciklama
        
        # Görsel Seviye (Rakam yerine '-' karakteri)
        gorsel_seviye = "-" * (seviye + 1)

        rapor_verisi.append({
            "Seviye": gorsel_seviye,
            "Üretilen Barkod (Çıktı)": hedef_barkod,
            "Ürün Açıklama": mevcut_aciklama,
            "Üretim Zinciri (Akış)": yeni_yol,
            "Giriş Barkodu (Tüketilen)": satir['GİRİŞ ÜRÜN SAP BARKODU'],
            "Proses": satir['PROSES'],
            "Zaman": satir['OLUŞTURMA ZAMANI'],
            "Makine No": satir['MAKİNE NO']
        })
        
        # Bir alt seviyeye (giriş barkoduna) in
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
    print("📊 Veri yükleniyor...")
    df = pd.read_excel(dosya_adi)
    
    # Sütun isimlerini temizle ve barkodları metne çevir
    df.columns = df.columns.str.strip()
    df['TEYİT VERİLEN BARKOD'] = df['TEYİT VERİLEN BARKOD'].astype(str)
    df['GİRİŞ ÜRÜN SAP BARKODU'] = df['GİRİŞ ÜRÜN SAP BARKODU'].astype(str)

    sorgu = input("\n🔍 İzlenecek ürün barkodunu girin: ").strip()
    
    rapor_verisi.clear()
    geriye_donuk_takip_listele(df, sorgu)

    if rapor_verisi:
        rapor_df = pd.DataFrame(rapor_verisi)
        
        # Raporu kaydet
        cikti_adi = f"Uretim_Akis_Raporu_{sorgu}.xlsx"
        rapor_df.to_excel(cikti_adi, index=False)
        
        print(f"\n✅ Rapor başarıyla oluşturuldu!")
        print(f"📂 Dosya: {cikti_adi}")
        print(f"📈 Toplam İşlem Sayısı: {len(rapor_verisi)}")
    else:
        print("\n❌ Kayıt bulunamadı. Barkodu ve veri dosyasını kontrol edin.")
else:
    print(f"❌ '{dosya_adi}' bulunamadı!")