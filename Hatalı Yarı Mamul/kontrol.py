import pandas as pd

# Excel dosyalarını okuma
urun_agaci = pd.read_excel("urun_agaci.xlsx")
tuketim = pd.read_excel("tuketim.xlsx")

# Rapor için kullanılacak liste
rapor = []

# Tüketim dosyasındaki her bir satır için kontrol
for index, row in tuketim.iterrows():
    cikis_aciklama = row["ÇIKIŞ ÜRÜN ACIKLAMA"]
    
    # Özel durum: Çıkış ürün açıklaması "H " veya "DMT" ile başlıyorsa kontrol yapma ve rapora ekleme
    if cikis_aciklama.startswith("H ") or cikis_aciklama.startswith("DMT") or cikis_aciklama.startswith("MN") :
        continue  # Bu kaydı atla ve sonraki kayda geç
    
    cikis_kodu = row["ÇIKIŞ ÜRÜN KODU"]
    giris_kodu = row["GİRİŞ ÜRÜN KODU"]
    
    # Ürün ağacı dosyasında ilgili çıkış kodunu bul
    urun_agaci_filtre = urun_agaci[urun_agaci["Çıkış Kod"] == cikis_kodu]
    
    if urun_agaci_filtre.empty:
        # Eğer çıkış kodu ürün ağacında yoksa
        rapor.append({
            "OLUŞTURMA ZAMANI": row["OLUŞTURMA ZAMANI"],
            "MAKİNE": row["MAKİNE NO"],
            "İŞ EMRİ": row["İŞ EMRİ"],
            "ÇIKIŞ ÜRÜN": cikis_aciklama,
            "GİRİŞ ÜRÜN": row["GİRİŞ ÜRÜN ACIKLAMA"],
            "DURUM": "Hatalı",
            "NEDEN": "Çıkış kodu ürün ağacında bulunamadı."
        })
    else:
        # Ürün ağacında bulunan giriş ürünlerini ve açıklamalarını al
        beklenen_giris_kodlari = urun_agaci_filtre["Giriş Kod"].tolist()
        beklenen_giris_aciklamalari = urun_agaci_filtre["Giriş Ürün"].tolist()
        
        if giris_kodu in beklenen_giris_kodlari:
            rapor.append({
                "OLUŞTURMA ZAMANI": row["OLUŞTURMA ZAMANI"],
                "MAKİNE": row["MAKİNE NO"],
                "İŞ EMRİ": row["İŞ EMRİ"],
                "ÇIKIŞ ÜRÜN": cikis_aciklama,
                "GİRİŞ ÜRÜN": row["GİRİŞ ÜRÜN ACIKLAMA"],
                "DURUM": "Doğru",
                "NEDEN": "Giriş ürünü ürün ağacı ile eşleşiyor."
            })
        else:
            # Hatalı durumda, beklenen giriş açıklamalarını birleştir
            rapor.append({
                "OLUŞTURMA ZAMANI": row["OLUŞTURMA ZAMANI"],
                "MAKİNE": row["MAKİNE NO"],
                "İŞ EMRİ": row["İŞ EMRİ"],
                "ÇIKIŞ ÜRÜN": cikis_aciklama,
                "GİRİŞ ÜRÜN": row["GİRİŞ ÜRÜN ACIKLAMA"],
                "DURUM": "Hatalı",
                "NEDEN": f"Beklenen giriş ürünleri: {', '.join(beklenen_giris_aciklamalari)}"
            })

# Raporu DataFrame'e dönüştürme
rapor_df = pd.DataFrame(rapor)

# Raporu Excel olarak kaydetme
rapor_df.to_excel("uretim_raporu.xlsx", index=False)

print("Rapor oluşturuldu ve 'uretim_raporu.xlsx' olarak kaydedildi.")