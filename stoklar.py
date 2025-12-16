import pandas as pd

# Excel dosyasını oku
file_path = 'stoklar.xlsx'  # Excel dosyanızın adını ve yolunu buraya girin
df = pd.read_excel(file_path)

# "Ürün" sütununda ilk karakterin sayı olup olmadığını kontrol et
def check_first_character_is_digit(product_name):
    if len(product_name) > 0:
        return product_name[0].isdigit()
    return False

# Yeni bir sütun ekleyerek ilk karakterin sayı olup olmadığını belirt
df['İLK KARAKTER SAYI MI?'] = df['ÜRÜN'].apply(check_first_character_is_digit)

# "Barkod Kodu" sütununda tekrarlanan kayıtları kontrol et
duplicate_barcodes = df[df.duplicated(subset='BARKOD KODU', keep=False)]

# Tekrarlanan barkod kodlarını yeni bir Excel dosyasına yaz
if not duplicate_barcodes.empty:
    duplicate_file_path = 'tekrarlanan_barkod_kodlari.xlsx'
    duplicate_barcodes.to_excel(duplicate_file_path, index=False)
    print(f"Tekrarlanan barkod kodları {duplicate_file_path} dosyasına kaydedildi.")
else:
    print("Tekrarlanan barkod kodu bulunmamaktadır.")

# Sonuçları yeni bir Excel dosyasına yaz
output_file_path = 'kontrol_edilen_stoklar.xlsx'
df.to_excel(output_file_path, index=False)

print(f"Kontrol sonuçları {output_file_path} dosyasına kaydedildi.")