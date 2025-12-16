import pandas as pd

# Excel dosyasını okuyun
file_path = "Faydalanma Kontrol.xlsx"
faydalanma_df = pd.read_excel(file_path, sheet_name="FAYDALANMA")  # FAYDALANMA tablosu
hepsi_df = pd.read_excel(file_path, sheet_name="HEPSI")  # HEPSI tablosu

# Yeni bir sütun oluşturmak için liste
durum_listesi = []

# Her bir satır için kontrol yap
for _, f_row in faydalanma_df.iterrows():
    ym_kodu = f_row["YM KODU"]
    stok_urun_tanimi = f_row["STOK ÜRÜN TANIMI"]
    
    # HEPSI tablosunda YM KODU'na göre filtrele
    hepsi_filtre = hepsi_df[hepsi_df["KOD"] == ym_kodu]
    
    # Ana Bileşen ve Yeni Bileşen sütunlarını birleştir
    bileşenler = hepsi_filtre.iloc[:, 2:9].values.flatten()  # ANA BİLEŞEN ve YENİ BİLEŞEN (1-6)
    
    # STOK ÜRÜN TANIMI'nın bileşenler arasında olup olmadığını kontrol et
    if any(stok_urun_tanimi in str(bileşen) for bileşen in bileşenler):
        durum_listesi.append("VAR")
    else:
        durum_listesi.append("YOK")

# Yeni sütunu FAYDALANMA tablosuna ekle
faydalanma_df["Özel"] = durum_listesi

# Güncellenmiş veriyi yeni bir Excel dosyasına kaydet
output_file = "Guncellenmis_Faydalanma.xlsx"
faydalanma_df.to_excel(output_file, index=False)

print(f"İşlem tamamlandı. Güncellenmiş dosya '{output_file}' olarak kaydedildi.")