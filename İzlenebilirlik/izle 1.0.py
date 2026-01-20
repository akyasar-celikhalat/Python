import pandas as pd
import os

# Sonuçları toplamak için global liste
rapor_verisi = []

def geriye_donuk_takip_listele(df, hedef_barkod, seviye=0, ust_zincir="", ziyaret_edilen=None):
    if ziyaret_edilen is None:
        ziyaret_edilen = set()
    
    hedef_barkod = str(hedef_barkod).strip()
    if hedef_barkod in ziyaret_edilen or hedef_barkod in ["nan", "None", ""]:
        return
    ziyaret_edilen.add(hedef_barkod)

    # Bu barkodun üretim kaydını bul
    girdiler = df[df['TEYİT VERİLEN BARKOD'] == hedef_barkod]

    if girdiler.empty:
        return

    for _, satir in girdiler.iterrows():
        g_barkod = str(satir['GİRİŞ ÜRÜN SAP BARKODU']).strip()
        g_tanim = str(satir['GİRİŞ ÜRÜN ACIKLAMA']).strip()
        c_barkod = str(satir['TEYİT VERİLEN BARKOD']).strip()
        c_tanim = str(satir['ÇIKIŞ ÜRÜN ACIKLAMA']).strip()
        proses = str(satir['PROSES']).strip()
        
        # Miktar Mantığı
        if proses in ['BD', 'DH']:
            miktar = satir.get('TEYİT MİKTARI Kg', 0)
        else:
            miktar = satir.get('GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg', 0)

        # ÜRETİM AKIŞI: [Mamul] → [Girdi] (Çizgisiz temiz metin)
        if ust_zincir == "":
            yeni_yol = f"{c_tanim} → {g_tanim}"
        else:
            yeni_yol = f"{ust_zincir} → {g_tanim}"
        
        # SEVİYE ÇİZGİLERİ: Giriş ürün açıklamasının başına ekle
        cizgiler = "-" * (seviye + 1)
        giris_tanim_hiyerarsik = f"{cizgiler} {g_tanim}"

        rapor_verisi.append({
            "Üretilen Barkod (Çıktı)": c_barkod,
            "ÇIKIŞ ÜRÜN ACIKLAMA": c_tanim,
            "Üretim Akışı (Hiyerarşi)": yeni_yol,
            "Tüketilen Barkod (Giriş)": g_barkod,
            "GİRİŞ ÜRÜN ACIKLAMA": giris_tanim_hiyerarsik, # Çizgiler buraya eklendi
            "Tüketim (m)": miktar,
            "Proses": proses,
            "Makine No": satir.get('MAKİNE NO', satir.get('MAKINE NO', '')),
            "Zaman": satir['OLUŞTURMA ZAMANI']
        })
        
        # Derine git
        geriye_donuk_takip_listele(df, g_barkod, seviye + 1, yeni_yol, ziyaret_edilen.copy())

# --- ANA PROGRAM ---
dosya_adi = "veri.xlsx" 

if os.path.exists(dosya_adi):
    print("📊 Veri tabanı işleniyor...")
    df = pd.read_excel(dosya_adi)
    
    df.columns = df.columns.str.strip()
    df['TEYİT VERİLEN BARKOD'] = df['TEYİT VERİLEN BARKOD'].astype(str).str.strip()
    df['GİRİŞ ÜRÜN SAP BARKODU'] = df['GİRİŞ ÜRÜN SAP BARKODU'].astype(str).str.strip()

    sorgu = input("\n🔍 İzlenecek ürün barkodunu girin: ").strip()
    
    rapor_verisi.clear()
    geriye_donuk_takip_listele(df, sorgu)

    if rapor_verisi:
        rapor_df = pd.DataFrame(rapor_verisi)
        
        sutun_sirasi = [
            "Üretilen Barkod (Çıktı)", "ÇIKIŞ ÜRÜN ACIKLAMA",
            "Tüketilen Barkod (Giriş)", "GİRİŞ ÜRÜN ACIKLAMA", "Tüketim (m)",
            "Proses", "Makine No", "Zaman", "Üretim Akışı (Hiyerarşi)"
        ]
        rapor_df = rapor_df[sutun_sirasi]
        
        cikti_adi = f"Izlenebilirlik_Raporu_{sorgu}.xlsx"
        rapor_df.to_excel(cikti_adi, index=False)
        
        print(f"\n✅ Rapor hazırlandı: {cikti_adi}")
    else:
        print("\n❌ Kayıt bulunamadı.")
else:
    print(f"❌ '{dosya_adi}' dosyası bulunamadı!")