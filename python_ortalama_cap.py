import openpyxl
import re

# Excel dosyasını aç
wb = openpyxl.load_workbook('ExcelReport (6).xlsx')
ws = wb['rphaldemtuketim']

# Çap bilgilerini toplamak için liste oluştur
cap_list = []

# Verileri oku
for row in ws.iter_rows(min_row=2, values_only=True):  # min_row=2 ilk satırı başlık olarak atlıyoruz
    product_definition = row[6]  # GİRİŞ ÜRÜN TANIMI sütunu
    
    if product_definition:
        # Çap bilgisini regular expression ile çıkart
        match = re.search(r'(\d+\.\d+)MM', product_definition)
        if match:
            cap = float(match.group(1))
            cap_list.append(cap)

# Çap bilgilerini yazdır
print("Çap Bilgileri:")
for cap in cap_list:
    print(cap)

# Çapların ortalamasını hesapla
if cap_list:
    average_cap = sum(cap_list) / len(cap_list)
    print(f"Ortalama Çap: {average_cap}")
else:
    print("Çap bilgisi bulunamadı.")