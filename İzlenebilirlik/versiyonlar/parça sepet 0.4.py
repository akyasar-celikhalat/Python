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
    pi_sayisi = 3.14159
    celik_yogunlugu = 7.85 

    for hedef_barkod in sorgu_listesi:
        hedef_barkod = str(hedef_barkod).strip()
        barkod_hareketleri = []
        
        # Tanım bilgilerini al
        malzeme_tanimi = ""
        tanim_bul = df[df['TEYİT VERİLEN BARKOD'] == hedef_barkod]
        if not tanim_bul.empty:
            malzeme_tanimi = str(tanim_bul.iloc[0]['ÇIKIŞ ÜRÜN ACIKLAMA'])
        else:
            tanim_bul = df[df['GİRİŞ ÜRÜN SAP BARKODU'] == hedef_barkod]
            if not tanim_bul.empty:
                malzeme_tanimi = str(tanim_bul.iloc[0]['GİRİŞ ÜRÜN ACIKLAMA'])

        # 1. ÜRETİM VERİLERİ
        uretimler = df[df['TEYİT VERİLEN BARKOD'] == hedef_barkod].copy()
        for _, satir in uretimler.iterrows():
            proses = str(satir['PROSES']).strip()
            
            if proses in ['TV', 'TG']:
                agirlik_kg = satir.get('TEYİT MİKTARI Metre', 0)
                # Çapı doğrudan ürün açıklamasından çek
                cap = cap_ayikla(malzeme_tanimi)
                
                if cap:
                    payda = ((cap**2) * pi_sayisi / 4) * celik_yogunlugu
                    # (813 * 1000) / ((4.60^2 * 3.14 / 4) * 7.85) = ~6.232 m
                    miktar = (agirlik_kg * 1000) / payda
                else:
                    miktar = 0 # Çap bulunamazsa 0 döner
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

        # 2. TÜKETİM VERİLERİ
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

        # KRONOLOJİK SIRALAMA VE STOK HESABI
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
cikti_dosyasi = os.path.join(script_dizini, "Hatasiz_Stok_Raporu.xlsx")

if os.path.exists(girdi_dosyasi):
    df = pd.read_excel(girdi_dosyasi)
    df.columns = df.columns.str.strip()
    
    for col in ['TEYİT VERİLEN BARKOD', 'GİRİŞ ÜRÜN SAP BARKODU']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    giris = input("Barkodları girin: ")
    sorgu_listesi = [b.strip() for b in giris.split(",")]

    sonuc = rapor_hazirla(df, sorgu_listesi)

    if sonuc:
        rapor_df = pd.DataFrame(sonuc)
        kolonlar = ["Sorgulanan Barkod", "Malzeme Tanımı", "İşlem Tipi", "Miktar", "Stok Durumu", "İlişkili Barkod", "İlişkili Tanım", "Zaman", "Makine"]
        rapor_df[kolonlar].to_excel(cikti_dosyasi, index=False)
        print(f"\n✅ Hesaplama düzeltildi. Rapor: {cikti_dosyasi}")
    else:
        print("\n❌ Veri bulunamadı.")
else:
    print(f"\n❌ {girdi_dosyasi} bulunamadı.")