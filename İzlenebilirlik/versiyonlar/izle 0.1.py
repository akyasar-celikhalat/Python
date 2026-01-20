import pandas as pd
import os

# Sonuçları toplamak için global bir liste
rapor_verisi = []

def geriye_donuk_takip_listele(df, hedef_barkod, seviye=0, ziyaret_edilen=None):
    if ziyaret_edilen is None:
        ziyaret_edilen = set()
    
    if hedef_barkod in ziyaret_edilen:
        return
    ziyaret_edilen.add(hedef_barkod)

    # Verinizdeki sütun isimlerini buraya birebir kopyaladım
    girdiler = df[df['TEYİT VERİLEN BARKOD'] == hedef_barkod]

    if girdiler.empty:
        return

    for _, satir in girdiler.iterrows():
        rapor_verisi.append({
            "Seviye": seviye,
            "Üretilen Barkod (Çıktı)": hedef_barkod,
            "Giriş Barkodu (Girdi)": satir['GİRİŞ ÜRÜN SAP BARKODU'],
            "Ürün Açıklama": satir['GİRİŞ ÜRÜN ACIKLAMA'],
            "Proses": satir['PROSES'],
            "Zaman": satir['OLUŞTURMA ZAMANI'],
            "Makine No": satir['MAKİNE NO']  # 'MAKINE NO' -> 'MAKİNE NO' olarak düzeltildi
        })
        
        geriye_donuk_takip_listele(df, str(satir['GİRİŞ ÜRÜN SAP BARKODU']), seviye + 1, ziyaret_edilen.copy())

# --- ANA PROGRAM ---
dosya_adi = "veri.xlsx" 

if os.path.exists(dosya_adi):
    df = pd.read_excel(dosya_adi)
    
    # Sütun isimlerindeki olası boşlukları temizleyelim (Güvenlik önlemi)
    df.columns = df.columns.str.strip()
    
    # Barkodları metne çevir
    df['TEYİT VERİLEN BARKOD'] = df['TEYİT VERİLEN BARKOD'].astype(str)
    df['GİRİŞ ÜRÜN SAP BARKODU'] = df['GİRİŞ ÜRÜN SAP BARKODU'].astype(str)

    sorgu = input("Raporunu almak istediğiniz barkodu girin: ").strip()
    
    rapor_verisi.clear()
    geriye_donuk_takip_listele(df, sorgu)

    if rapor_verisi:
        rapor_df = pd.DataFrame(rapor_verisi)
        cikti_adi = f"Izlenebilirlik_Raporu_{sorgu}.xlsx"
        rapor_df.to_excel(cikti_adi, index=False)
        print(f"\n✅ Başarılı! {cikti_adi} dosyası oluşturuldu.")
    else:
        print("\n❌ Bu barkoda ait bir alt bileşen bulunamadı. Lütfen barkodu doğru girdiğinizden emin olun.")
else:
    print(f"❌ '{dosya_adi}' bulunamadı!")