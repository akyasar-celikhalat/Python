import pandas as pd

# Excel dosyasını okuma
file_path = r"C:\Users\yak\Desktop\Python\Provizyon.xlsx"
df = pd.read_excel(file_path)

# Giriş ve çıkış ürün kodlarını içeren sütunları seçme
giris_urun_kodu = df['GİRİŞ ÜRÜN KODU']
cikis_urun_kodu = df['ÇIKIŞ ÜRÜN KODU']

# Son "-83HC" ifadesini kontrol etmek için fonksiyon tanımlama
def check(giris_kod, cikis_kod):
    giris_suffix = giris_kod.split('-')[-1]
    cikis_suffix = cikis_kod.split('-')[-1]
    return giris_suffix == cikis_suffix

# Her satır için kontrol yapma ve sonuçları yeni bir sütuna ekleyme
df['83HC_Eşleşti'] = df.apply(lambda row: check(row['GİRİŞ ÜRÜN KODU'], row['ÇIKIŞ ÜRÜN KODU']), axis=1)

# Sonuçları ekrana yazdırma veya yeni bir Excel dosyasına kaydetme
print(df[['GİRİŞ ÜRÜN KODU', 'ÇIKIŞ ÜRÜN KODU', 'Eşleşti']])

# Yeni Excel dosyasına kaydetme (isteğe bağlı)
output_file_path = 'output.xlsx'
df.to_excel(output_file_path, index=False)