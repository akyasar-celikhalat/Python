import pandas as pd
import os
import re

def cap_ayikla(tanim):
    """Malzeme tanımı içinden '4.60MM' gibi çap değerlerini bulur."""
    match = re.search(r"(\d+\.\d+|\d+)MM", str(tanim).upper())
    if match:
        return float(match.group(1))
    return None

def rapor_hazirla(df, sorgu_listesi):
    tum_sonuclar = []
    pi_sayisi = 3.14159265359
    celik_yogunlugu = 7.85 

    for hedef_barkod in sorgu_listesi:
        hedef_barkod = str(hedef_barkod).strip()
        barkod_hareketleri = []
        
        malzeme_tanimi = ""
        tanim_bul = df[df['TEYİT VERİLEN BARKOD'] == hedef_barkod]
        if not tanim_bul.empty:
            malzeme_tanimi = str(tanim_bul.iloc[0]['ÇIKIŞ ÜRÜN ACIKLAMA'])
        else:
            tanim_bul = df[df['GİRİŞ ÜRÜN SAP BARKODU'] == hedef_barkod]
            if not tanim_bul.empty:
                malzeme_tanimi = str(tanim_bul.iloc[0]['GİRİŞ ÜRÜN ACIKLAMA'])

        uretimler = df[df['TEYİT VERİLEN BARKOD'] == hedef_barkod].copy()
        for _, satir in uretimler.iterrows():
            proses = str(satir['PROSES']).strip()
            if proses in ['TV', 'TG']:
                agirlik_kg = satir.get('TEYİT MİKTARI Metre', 0)
                cap = cap_ayikla(malzeme_tanimi)
                if cap:
                    payda = ((cap**2) * pi_sayisi / 4) * celik_yogunlugu
                    miktar = (agirlik_kg * 1000) / payda
                else:
                    miktar = 0
            else:
                miktar = satir.get('TEYİT MİKTARI Kg', 0)

            barkod_hareketleri.append({
                "Sorgulanan Barkod": hedef_barkod,
                "Malzeme Tanımı": malzeme_tanimi,
                "İşlem Tipi": "ÜRETİM",
                "Miktar": round(miktar, 3),
                "İlişkili Barkod": satir['GİRİŞ ÜRÜN SAP BARKODU'],
                "İlişkili Tanım": satir['GİRİŞ ÜRÜN ACIKLAMA'],
                "Zaman": satir['OLUŞTURMA ZAMANI'],
                "Makine": satir.get('MAKİNE NO', satir.get('MAKINE NO', ''))
            })

        tuketimler = df[df['GİRİŞ ÜRÜN SAP BARKODU'] == hedef_barkod]
        for _, satir in tuketimler.iterrows():
            miktar = satir.get('GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg', 0)
            barkod_hareketleri.append({
                "Sorgulanan Barkod": hedef_barkod,
                "Malzeme Tanımı": malzeme_tanimi,
                "İşlem Tipi": "TÜKETİM",
                "Miktar": round(miktar, 3),
                "İlişkili Barkod": satir['TEYİT VERİLEN BARKOD'],
                "İlişkili Tanım": satir['ÇIKIŞ ÜRÜN ACIKLAMA'],
                "Zaman": satir['OLUŞTURMA ZAMANI'],
                "Makine": satir.get('MAKİNE NO', satir.get('MAKINE NO', ''))
            })

        if barkod_hareketleri:
            barkod_hareketleri.sort(key=lambda x: x['Zaman'])
            mevcut_bakiye = 0
            for hareket in barkod_hareketleri:
                if hareket['İşlem Tipi'] == "ÜRETİM":
                    mevcut_bakiye += hareket['Miktar']
                else:
                    mevcut_bakiye -= hareket['Miktar']
                hareket['Stok Durumu'] = round(mevcut_bakiye, 3)
                tum_sonuclar.append(hareket)

    return tum_sonuclar

# --- ANA PROGRAM ---

script_dizini = os.path.dirname(os.path.abspath(__file__))
girdi_dosyasi = os.path.join(script_dizini, "veri.xlsx")

if not os.path.exists(girdi_dosyasi):
    print(f"❌ HATA: {girdi_dosyasi} bulunamadı!")
else:
    print("🚀 Veri yükleniyor...")
    df = pd.read_excel(girdi_dosyasi)
    df.columns = df.columns.str.strip()
    for col in ['TEYİT VERİLEN BARKOD', 'GİRİŞ ÜRÜN SAP BARKODU']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    print("✅ Hazır.")

    while True:
        print("\n" + "-"*40)
        giris = input("📝 Barkod girin (Kapatmak için H veya ÇIKIŞ): ").strip()

        # Çıkış Kontrolü
        if giris.upper() in ['H', 'ÇIKIŞ', 'CIKIS', 'İPTAL', 'KAPAT']:
            print("👋 Program kapatıldı.")
            break
        
        if not giris:
            print("⚠️ Boş giriş yapıldı. Kapatmak için H yazın.")
            continue

        sorgu_listesi = [b.strip() for b in giris.split(",")]
        sonuc = rapor_hazirla(df, sorgu_listesi)

        if sonuc:
            rapor_df = pd.DataFrame(sonuc)
            kolonlar = ["Sorgulanan Barkod", "Malzeme Tanımı", "İşlem Tipi", "Miktar", 
                        #"Stok Durumu", 
                        "İlişkili Barkod", "İlişkili Tanım", "Zaman", "Makine"]
            
            # Dosya Adı Oluşturma (İlk barkod)
            ana_barkod = sorgu_listesi[0].replace("/", "-").replace("\\", "-")
            cikti_adi = f"Rapor_{ana_barkod}.xlsx"
            cikti_yolu = os.path.join(script_dizini, cikti_adi)
            
            rapor_df[kolonlar].to_excel(cikti_yolu, index=False)
            print(f"✅ Başarılı: {cikti_adi}")
        else:
            print("❌ Veri bulunamadı.")