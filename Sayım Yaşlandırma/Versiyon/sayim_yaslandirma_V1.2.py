import pandas as pd

# Excel dosyasını yükleme
file_path = "sayim_verisi.xlsx"  # Excel dosyanızın yolu
df = pd.read_excel(file_path)

# Veriyi filtreleme: "SAYILDI MI" sütununun "SAYILDI" olduğu satırlar
filtered_df = df[df["SAYILDI MI"] == "SAYILDI"]

# En yüksek SAYIM NO değerini bulma (son ay)
en_yuksek_sayim_no = filtered_df["SAYIM NO"].max()

# Son sayımdaki verileri filtreleme
son_sayim_df = filtered_df[filtered_df["SAYIM NO"] == en_yuksek_sayim_no]

# Her bir BOBİN BARKODU'nun kaç farklı sayımda yer aldığını hesaplama
stok_yasi = filtered_df.groupby("BOBİN BARKODU")["SAYIM NO"].nunique().reset_index()
stok_yasi.rename(columns={"SAYIM NO": "SAYIM_YASI"}, inplace=True)

# En az iki sayımda yer alan stokları filtreleme
eski_stoklar = stok_yasi[stok_yasi["SAYIM_YASI"] >= 2]

# Son sayımdaki stoklarla eski stokları birleştirme
sonuc_df = son_sayim_df.merge(eski_stoklar, on="BOBİN BARKODU", how="inner")

# Gereksiz sütunları kaldırma (isteğe bağlı)
sonuc_df = sonuc_df[["BOBİN BARKODU", "ÜRÜN KODU", "ÜRÜN AÇIKLAMASI","SAYILAN MİKTAR (METRE)", "SAYILAN MİKTAR (KİLO)",  "SAYIM_YASI"]]

# Veri tipini kontrol etme ve dönüştürme
if not pd.api.types.is_numeric_dtype(sonuc_df["SAYIM_YASI"]):
    sonuc_df["SAYIM_YASI"] = pd.to_numeric(sonuc_df["SAYIM_YASI"], errors="coerce")

# Boş değerleri temizleme
sonuc_df = sonuc_df.dropna(subset=["SAYIM_YASI"])

# SAYIM_YASI'na göre büyükten küçüğe sıralama
sonuc_df = sonuc_df.sort_values(by="SAYIM_YASI", ascending=False)

# İndeksleri sıfırlama
sonuc_df = sonuc_df.reset_index(drop=True)

# Sonuçları yeni bir Excel dosyasına kaydetme
output_file_path = "eski_stoklar_listesi_sirali.xlsx"
sonuc_df.to_excel(output_file_path, index=False)

print(f"Analiz tamamlandı. Sonuçlar '{output_file_path}' dosyasına kaydedildi.")