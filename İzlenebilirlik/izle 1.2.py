import pandas as pd
import os

# Rapor verilerini tutacak global liste
rapor_verisi = []

def geriye_donuk_takip_listele(df, hedef_barkod, ana_mamul_barkod, seviye=0, ust_zincir="", ziyaret_edilen=None):
    if ziyaret_edilen is None:
        ziyaret_edilen = set()
    
    hedef_barkod = str(hedef_barkod).strip()
    if hedef_barkod in ziyaret_edilen or hedef_barkod in ["nan", "None", ""]:
        return
    ziyaret_edilen.add(hedef_barkod)

    girdiler = df[df['TEYİT VERİLEN BARKOD'] == hedef_barkod]

    if girdiler.empty:
        return

    for _, satir in girdiler.iterrows():
        g_barkod = str(satir['GİRİŞ ÜRÜN SAP BARKODU']).strip()
        g_tanim = str(satir['GİRİŞ ÜRÜN ACIKLAMA']).strip()
        c_tanim = str(satir['ÇIKIŞ ÜRÜN ACIKLAMA']).strip()
        proses = str(satir['PROSES']).strip()
        
        # Miktar Mantığı
        miktar = satir.get('TEYİT MİKTARI Kg', 0) if proses in ['BD', 'DH'] else satir.get('GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg', 0)

        # Akış ve Hiyerarşi
        yeni_yol = f"{ust_zincir} → {g_tanim}" if ust_zincir else f"{c_tanim} → {g_tanim}"
        giris_tanim_hiyerarsik = ("-" * (seviye + 1)) + " " + g_tanim

        rapor_verisi.append({
            "Sorgulanan Ana Mamul": ana_mamul_barkod,
            "Üretilen Barkod (Çıktı)": hedef_barkod,
            "ÇIKIŞ ÜRÜN ACIKLAMA": c_tanim,
            "Tüketilen Barkod (Giriş)": g_barkod,
            "GİRİŞ ÜRÜN ACIKLAMA": giris_tanim_hiyerarsik,
            "İşlem Miktarı (Kg)": miktar,
            "Proses": proses,
            "Makine No": satir.get('MAKİNE NO', satir.get('MAKINE NO', '')),
            "Zaman": satir['OLUŞTURMA ZAMANI'],
            "Üretim Akışı (Hiyerarşi)": yeni_yol
        })
        
        geriye_donuk_takip_listele(df, g_barkod, ana_mamul_barkod, seviye + 1, yeni_yol, ziyaret_edilen.copy())

# --- ANA PROGRAM ---

# 1. Dosya Yollarını Belirle (Scriptin olduğu dizini baz al)
script_dizini = os.path.dirname(os.path.abspath(__file__))
girdi_dosyasi = os.path.join(script_dizini, "veri.xlsx")
cikti_dosyasi = os.path.join(script_dizini, "Izlenebilirlik_Raporu.xlsx")

if os.path.exists(girdi_dosyasi):
    print(f"✅ Dosya bulundu: {girdi_dosyasi}")
    
    # Veriyi yükle ve temizle
    df = pd.read_excel(girdi_dosyasi)
    df.columns = df.columns.str.strip()
    df['TEYİT VERİLEN BARKOD'] = df['TEYİT VERİLEN BARKOD'].astype(str).str.strip()
    df['GİRİŞ ÜRÜN SAP BARKODU'] = df['GİRİŞ ÜRÜN SAP BARKODU'].astype(str).str.strip()

    print("\n--- ÇOKLU İZLENEBİLİRLİK SİSTEMİ ---")
    giris = input("Barkodları aralarına virgül koyarak girin: ")
    sorgu_listesi = [b.strip() for b in giris.split(",")]
    
    rapor_verisi.clear()
    
    for barkod in sorgu_listesi:
        print(f"🔍 {barkod} analiz ediliyor...")
        geriye_donuk_takip_listele(df, barkod, barkod)

    if rapor_verisi:
        rapor_df = pd.DataFrame(rapor_verisi)
        
        # Raporu aynı dizine kaydet
        rapor_df.to_excel(cikti_dosyasi, index=False)
        
        print("\n" + "="*30)
        print(f"✅ İŞLEM BAŞARIYLA TAMAMLANDI")
        print(f"📊 Toplan Veri: {len(rapor_df)} satır")
        print(f"📁 Kayıt Yeri: {cikti_dosyasi}")
        print("="*30)
    else:
        print("\n❌ Girilen barkodlara ait alt bileşen verisi bulunamadı.")
else:
    print(f"\n❌ HATA: '{girdi_dosyasi}' dosyası bulunamadı!")
    print(f"Lütfen 'veri.xlsx' dosyasının script ile aynı klasörde olduğundan emin olun.")