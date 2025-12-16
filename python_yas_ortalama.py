import openpyxl

# Excel dosyasını aç
wb = openpyxl.load_workbook('ornek_dosya.xlsx')
ws = wb.active

# Yaşların tutulacağı liste
ages = []

# Çalışma sayfasındaki tüm hücreleri oku
for row in ws.iter_rows(min_row=2, values_only=True):  # min_row=2 ilk satırı başlık olarak atlıyoruz
    age = row[1]  # Yaşı ikinci sütundan alıyoruz (index 1)
    if isinstance(age, (int, float)):  # Yaşa sayısal bir değer olup olmadığını kontrol ediyoruz
        ages.append(age)

# Yaş ortalamasını hesapla
if ages:
    average_age = sum(ages) / len(ages)
    print(f"Yaş ortalaması: {average_age}")
else:
    print("Yaş bilgisi bulunamadı.")
