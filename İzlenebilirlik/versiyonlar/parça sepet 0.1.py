import pandas as pd
import os
import numpy as np

def kapsamli_hareket_raporu(df, sorgu_listesi):
    tum_hareketler = []
    pi_sayisi = 3.14159

    for hedef_barkod in sorgu_listesi:
        hedef_barkod = str(hedef_barkod).strip()
        
        # Sorgulanan barkodun tanımını bul (İlk bulduğunu alır)
        malzeme_tanimi = ""
        tanim_bul = df[df['TEYİT VERİLEN BARKOD'] == hedef_barkod]
        if not tanim_bul.empty:
            malzeme_tanimi = tanim_bul.iloc[0]['ÇIKIŞ ÜRÜN ACIKLAMA']
        else:
            tanim_bul = df[df['GİRİŞ ÜRÜN SAP BARKODU'] == hedef_barkod]
            if not tanim_bul.empty:
                malzeme_tanimi = tanim_bul.iloc[0]['GİRİŞ ÜRÜN ACIKLAMA']

        # 1. ÜRETİM VERİLERİ
        uretimler = df[df['TEYİT VERİLEN BARKOD'] == hedef_barkod].copy()
        
        for _, satir in uretimler.iterrows():
            proses = str(satir['PROSES']).strip()
            
            # TV ve TG için özel uzunluk hesaplama (Kg -> Metre)
            if proses in ['TV', 'TG']:
                agirlik_kg = satir.get('TEYİT MİKTARI Metre', 0) # Belirttiğiniz üzere kg burada
                cap = satir.get('ÇAP', 1) # Çap kolonu yoksa 1 kabul eder (hata önleme)
                ozgul_agirlik = satir.get('ÖZGÜL AĞIRLIK', 8.89) # Standart bakır yoğunluğu
                
                # Formül: (Ağırlık * 1000) / ( (Çap^2 * PI / 4) * Özgül Ağırlık )
                if cap > 0:
                    hesaplanan_metre = (agirlik_kg * 1000) / (( (cap**2) * pi_sayisi / 4) * ozgul_agirlik)
                else:
                    hesaplanan_metre = 0
                miktar = hesaplanan_metre
            else:
                # Diğer proseslerde standart miktar
                miktar = satir.get('TEYİT MİKTARI Kg', 0) if proses in ['BD', 'DH'] else satir.get('GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg', 0)

            tum_hareketler.append({
                "Sorgulanan Barkod": hedef_barkod,
                "Malzeme Tanımı": malzeme_tanimi,
                "İşlem Tipi": f"ÜRETİM ({proses})",
                "Miktar": miktar,
                "İlişkili Barkod": satir['GİRİŞ ÜRÜN SAP BARKODU'],
                "İlişkili Tanım": satir['GİRİŞ ÜRÜN ACIKLAMA'],
                "Zaman": satir['OLUŞTURMA ZAMANI'],
                "Makine": satir.get('MAKİNE NO', satir.get('MAKINE NO', ''))
            })

        # 2. TÜKETİM VERİLERİ
        tuketimler = df[df['GİRİŞ ÜRÜN SAP BARKODU'] == hedef_barkod]
        
        for _, satir in tuketimler.iterrows():
            # Belirttiğiniz üzere tüketim sütunundaki değerler zaten metre
            miktar = satir.get('GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg', 0)
            
            tum_hareketler.append({
                "Sorgulanan Barkod": hedef_barkod,
                "Malzeme Tanımı": malzeme_tanimi,
                "İşlem Tipi": "TÜKETİM",
                "Miktar": miktar,
                "İlişkili Barkod": satir['TEYİT VERİLEN BARKOD'],
                "İlişkili Tanım": satir['ÇIKIŞ ÜRÜN ACIKLAMA'],
                "Zaman": satir['OLUŞTURMA ZAMANI'],
                "Makine": satir.get('MAKİNE NO', satir.get('MAKINE NO', ''))
            })

    return tum_hareketler

# --- ANA PROGRAM ---

script_dizini = os.path.dirname(os.path.abspath(__file__))
girdi_dosyasi = os.path.join(script_dizini, "veri.xlsx")
cikti_dosyasi = os.path.join(script_dizini, "Guncel_Izlenebilirlik_Raporu.xlsx")

if os.path.exists(girdi_dosyasi):
    df = pd.read_excel(girdi_dosyasi)
    df.columns = df.columns.str.strip()
    
    # Barkod kolonlarını string yap
    for col in ['TEYİT VERİLEN BARKOD', 'GİRİŞ ÜRÜN SAP BARKODU']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    giris = input("Barkodları girin (virgülle ayırarak): ")
    sorgu_listesi = [b.strip() for b in giris.split(",")]

    sonuc = kapsamli_hareket_raporu(df, sorgu_listesi)

    if sonuc:
        rapor_df = pd.DataFrame(sonuc)
        # İstenen kolon düzeni
        kolon_sirasi = ["Sorgulanan Barkod", "Malzeme Tanımı", "İşlem Tipi", "Miktar", "İlişkili Barkod", "İlişkili Tanım", "Zaman", "Makine"]
        rapor_df = rapor_df[kolon_sirasi]
        
        rapor_df.to_excel(cikti_dosyasi, index=False)
        print(f"\n✅ İşlem Tamamlandı. Rapor oluşturuldu: {cikti_dosyasi}")
    else:
        print("\n❌ Veri bulunamadı.")
else:
    print(f"❌ HATA: {girdi_dosyasi} dosyası bulunamadı.")