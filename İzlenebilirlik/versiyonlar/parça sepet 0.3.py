import pandas as pd
import os

def rapor_hazirla(df, sorgu_listesi):
    tum_sonuclar = []
    pi_sayisi = 3.14159
    # Çelik yoğunluğu (83HC malzemeler için 7.85 g/cm3)
    celik_yogunlugu = 7.85 

    for hedef_barkod in sorgu_listesi:
        hedef_barkod = str(hedef_barkod).strip()
        barkod_hareketleri = []
        
        # Malzeme Tanımını Bul
        malzeme_tanimi = ""
        tanim_bul = df[df['TEYİT VERİLEN BARKOD'] == hedef_barkod]
        if not tanim_bul.empty:
            malzeme_tanimi = tanim_bul.iloc[0]['ÇIKIŞ ÜRÜN ACIKLAMA']
        else:
            tanim_bul = df[df['GİRİŞ ÜRÜN SAP BARKODU'] == hedef_barkod]
            if not tanim_bul.empty:
                malzeme_tanimi = tanim_bul.iloc[0]['GİRİŞ ÜRÜN ACIKLAMA']

        # 1. ÜRETİM VERİLERİ (Girişler)
        uretimler = df[df['TEYİT VERİLEN BARKOD'] == hedef_barkod].copy()
        for _, satir in uretimler.iterrows():
            proses = str(satir['PROSES']).strip()
            
            # TV ve TG için Kg -> Metre dönüşümü (Çelik yoğunluğu ile)
            if proses in ['TV', 'TG']:
                agirlik_kg = satir.get('TEYİT MİKTARI Metre', 0)
                cap = satir.get('ÇAP', 2.90) # Dosyada ÇAP yoksa 2.90 varsayılır
                
                # Uzunluk Formülü: (Ağırlık * 1000) / ( (D^2 * PI / 4) * Yoğunluk )
                payda = ((cap**2) * pi_sayisi / 4) * celik_yogunlugu
                miktar = (agirlik_kg * 1000) / payda if payda > 0 else 0
            else:
                # Diğer prosesler (BD, DH vb.)
                miktar = satir.get('TEYİT MİKTARI Kg', 0)

            barkod_hareketleri.append({
                "Sorgulanan Barkod": hedef_barkod,
                "Malzeme Tanımı": malzeme_tanimi,
                "İşlem Tipi": "ÜRETİM",
                "Miktar": miktar,
                "İlişkili Barkod": satir['GİRİŞ ÜRÜN SAP BARKODU'],
                "İlişkili Tanım": satir['GİRİŞ ÜRÜN ACIKLAMA'],
                "Zaman": satir['OLUŞTURMA ZAMANI'],
                "Makine": satir.get('MAKİNE NO', satir.get('MAKINE NO', ''))
            })

        # 2. TÜKETİM VERİLERİ (Çıkışlar)
        tuketimler = df[df['GİRİŞ ÜRÜN SAP BARKODU'] == hedef_barkod]
        for _, satir in tuketimler.iterrows():
            miktar = satir.get('GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg', 0) # Tüketimler zaten metre
            
            barkod_hareketleri.append({
                "Sorgulanan Barkod": hedef_barkod,
                "Malzeme Tanımı": malzeme_tanimi,
                "İşlem Tipi": "TÜKETİM",
                "Miktar": miktar,
                "İlişkili Barkod": satir['TEYİT VERİLEN BARKOD'],
                "İlişkili Tanım": satir['ÇIKIŞ ÜRÜN ACIKLAMA'],
                "Zaman": satir['OLUŞTURMA ZAMANI'],
                "Makine": satir.get('MAKİNE NO', satir.get('MAKINE NO', ''))
            })

        # --- HAREKETLERİ ZAMANA GÖRE SIRALA VE STOK HESAPLA ---
        if barkod_hareketleri:
            # Zaman sütununa göre sırala (Kronolojik stok takibi için)
            barkod_hareketleri.sort(key=lambda x: x['Zaman'])
            
            mevcut_bakiye = 0
            for hareket in barkod_hareketleri:
                if hareket['İşlem Tipi'] == "ÜRETİM":
                    mevcut_bakiye += hareket['Miktar']
                else:
                    mevcut_bakiye -= hareket['Miktar']
                
                hareket['Stok Durumu'] = mevcut_bakiye
                tum_sonuclar.append(hareket)

    return tum_sonuclar

# --- ANA PROGRAM ---
script_dizini = os.path.dirname(os.path.abspath(__file__))
girdi_dosyasi = os.path.join(script_dizini, "veri.xlsx")
cikti_dosyasi = os.path.join(script_dizini, "Stok_Hareket_Raporu.xlsx")

if os.path.exists(girdi_dosyasi):
    df = pd.read_excel(girdi_dosyasi)
    df.columns = df.columns.str.strip()
    
    # Barkodları temizle
    for col in ['TEYİT VERİLEN BARKOD', 'GİRİŞ ÜRÜN SAP BARKODU']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    giris = input("Analiz edilecek barkodları girin (virgülle ayırarak): ")
    sorgu_listesi = [b.strip() for b in giris.split(",")]

    sonuc_listesi = rapor_hazirla(df, sorgu_listesi)

    if sonuc_listesi:
        rapor_df = pd.DataFrame(sonuc_listesi)
        # Sütunları düzenle
        kolonlar = ["Sorgulanan Barkod", "Malzeme Tanımı", "İşlem Tipi", "Miktar", "Stok Durumu", "İlişkili Barkod", "İlişkili Tanım", "Zaman", "Makine"]
        rapor_df = rapor_df[kolonlar]
        
        rapor_df.to_excel(cikti_dosyasi, index=False)
        print(f"\n✅ Rapor başarıyla oluşturuldu: {cikti_dosyasi}")
    else:
        print("\n❌ Kayıt bulunamadı.")
else:
    print(f"\n❌ Hata: {girdi_dosyasi} bulunamadı.")