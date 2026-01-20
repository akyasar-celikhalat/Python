import pandas as pd
import os

rapor_verisi = []

def geriye_donuk_takip_listele(df, hedef_barkod, seviye=0, alt_zincir="", ziyaret_edilen=None):
    if ziyaret_edilen is None:
        ziyaret_edilen = set()
    
    hedef_barkod = str(hedef_barkod).strip()
    if hedef_barkod in ziyaret_edilen or hedef_barkod in ["nan", "None", ""]:
        return
    ziyaret_edilen.add(hedef_barkod)

    # Bu barkodun hangi hammadde/yarı mamullerden üretildiğini (girdilerini) bul
    girdiler = df[df['TEYİT VERİLEN BARKOD'] == hedef_barkod]

    if girdiler.empty:
        # Eğer bu barkod bir çıktı değilse (yani hammadde ise), arama burada biter.
        return

    for _, satir in girdiler.iterrows():
        # Verileri güvenli al
        g_barkod = str(satir['GİRİŞ ÜRÜN SAP BARKODU']).strip()
        g_tanim = str(satir['GİRİŞ ÜRÜN ACIKLAMA']).strip()
        c_barkod = str(satir['TEYİT VERİLEN BARKOD']).strip()
        c_tanim = str(satir['ÇIKIŞ ÜRÜN ACIKLAMA']).strip()
        proses = str(satir['PROSES']).strip()
        
        # Miktar Mantığı (AKT, SW, TC, TCD, TV vs. BD, DH)
        if proses in ['BD', 'DH']:
            miktar = satir.get('TEYİT MİKTARI Kg', 0)
        else:
            miktar = satir.get('GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg', 0)

        # HİYERARŞİ OLUŞTURMA: 
        # Sizin istediğiniz: "En alttaki yarı mamulden başlayarak yukarı doğru"
        # Bu yüzden zinciri: [Giriş Tanımı] → [Mevcut Zincir] şeklinde kuruyoruz
        if alt_zincir == "":
            yeni_zincir = f"{g_tanim} → {c_tanim}"
        else:
            yeni_zincir = f"{g_tanim} → {alt_zincir}"
        
        gorsel_seviye = "-" * (seviye + 1)
        tam_akisi = f"{gorsel_seviye} {yeni_zincir}"

        # RAPOR SÜTUNLARI (Sıralama isteğinize göre)
        rapor_verisi.append({
            "Üretilen Barkod (Çıktı)": c_barkod,        # 1. SÜTUN
            "ÇIKIŞ ÜRÜN ACIKLAMA": c_tanim,             # 2. SÜTUN
            "Üretim Akışı (Hiyerarşi)": tam_akisi,
            "Tüketilen Barkod (Giriş)": g_barkod,
            "GİRİŞ ÜRÜN ACIKLAMA": g_tanim,
            "İşlem Miktarı (Kg)": miktar,
            "Proses": proses,
            "Makine No": satir.get('MAKİNE NO', satir.get('MAKINE NO', '')),
            "Zaman": satir['OLUŞTURMA ZAMANI']
        })
        
        # Bir adım daha geriye (Giriş barkodunun kaynağına) git
        geriye_donuk_takip_listele(df, g_barkod, seviye + 1, yeni_zincir, ziyaret_edilen.copy())

# --- PROGRAM BAŞLANGICI ---
dosya_adi = "veri.xlsx" 

if os.path.exists(dosya_adi):
    print("🚀 Veri işleniyor, lütfen bekleyin...")
    df = pd.read_excel(dosya_adi)
    
    # Sütun isimlerini ve verileri temizle
    df.columns = df.columns.str.strip()
    df['TEYİT VERİLEN BARKOD'] = df['TEYİT VERİLEN BARKOD'].astype(str).str.strip()
    df['GİRİŞ ÜRÜN SAP BARKODU'] = df['GİRİŞ ÜRÜN SAP BARKODU'].astype(str).str.strip()

    sorgu_barkod = input("\n🔍 İzlenecek barkodu girin: ").strip()
    
    rapor_verisi.clear()
    geriye_donuk_takip_listele(df, sorgu_barkod)

    if rapor_verisi:
        # Excel'e aktarırken sütun sırasını sabitleyelim
        cikti_df = pd.DataFrame(rapor_verisi)
        sutun_sirasi = [
            "Üretilen Barkod (Çıktı)", "ÇIKIŞ ÜRÜN ACIKLAMA", "Üretim Akışı (Hiyerarşi)",
            "Tüketilen Barkod (Giriş)", "GİRİŞ ÜRÜN ACIKLAMA", "İşlem Miktarı (Kg)",
            "Proses", "Makine No", "Zaman"
        ]
        cikti_df = cikti_df[sutun_sirasi]
        
        dosya_yolu = f"Tam_Izlenebilirlik_Raporu_{sorgu_barkod}.xlsx"
        cikti_df.to_excel(dosya_yolu, index=False)
        print(f"\n✅ İşlem Tamam! '{dosya_yolu}' dosyası oluşturuldu.")
        print(f"Toplam {len(rapor_verisi)} üretim adımı bulundu.")
    else:
        print(f"\n❌ Hata: '{sorgu_barkod}' barkodu için veri bulunamadı.")
else:
    print(f"❌ Hata: '{dosya_adi}' dosyası bulunamadı.")

    # Filmaşin, son yarı mamul gözükmüyor.