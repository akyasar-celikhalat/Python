import pandas as pd
import os

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
            "Sorgulanan Ana Mamul": ana_mamul_barkod,  # Çoklu takipte karışmaması için
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

# Dosyayı script ile aynı klasörde ara
script_dizini = os.path.dirname(os.path.abspath(__file__))
dosya_adi = os.path.join(script_dizini, "veri.xlsx")

if not os.path.exists(dosya_adi):
    print(f"❌ HATA: '{dosya_adi}' bulunamadı!")
    print(f"Lütfen dosyanın şu klasörde olduğundan emin olun: {script_dizini}")
else:
    df = pd.read_excel(dosya_adi)

if os.path.exists(dosya_adi):
    df = pd.read_excel(dosya_adi)
    df.columns = df.columns.str.strip()
    df['TEYİT VERİLEN BARKOD'] = df['TEYİT VERİLEN BARKOD'].astype(str).str.strip()
    df['GİRİŞ ÜRÜN SAP BARKODU'] = df['GİRİŞ ÜRÜN SAP BARKODU'].astype(str).str.strip()

    print("\n--- ÇOKLU İZLENEBİLİRLİK SİSTEMİ ---")
    giris = input("Barkodları aralarına virgül koyarak girin (veya tek barkod): ")
    sorgu_listesi = [b.strip() for b in giris.split(",")]
    
    rapor_verisi.clear()
    
    for barkod in sorgu_listesi:
        print(f"🔍 {barkod} analiz ediliyor...")
        geriye_donuk_takip_listele(df, barkod, barkod)

    if rapor_verisi:
        rapor_df = pd.DataFrame(rapor_verisi)
        cikti_adi = "Izlenebilirlik_Raporu.xlsx"
        rapor_df.to_excel(cikti_adi, index=False)
        print(f"\n✅ İşlem tamam! {len(sorgu_listesi)} mamul için toplam {len(rapor_df)} satırlık rapor oluşturuldu.")
        print(f"📁 Dosya: {cikti_adi}")
    else:
        print("\n❌ Girilen barkodlara ait veri bulunamadı.")
else:
    print(f"❌ '{dosya_adi}' bulunamadı!")