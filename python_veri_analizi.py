import openpyxl
from collections import defaultdict

# Excel dosyasını aç
wb = openpyxl.load_workbook('ExcelReport (6).xlsx')
ws = wb['rphaldemtuketim']

# Verileri toplamak için sözlük oluştur
machine_counts = defaultdict(int)
product_stocks = []
stock_diffs = []

# Verileri oku
for row in ws.iter_rows(min_row=2, values_only=True):  # min_row=2 ilk satırı başlık olarak atlıyoruz
    machine_no = row[2]
    product_stock = row[7]
    stock_diff = row[9]
    
    if isinstance(product_stock, (int, float)):
        product_stocks.append(product_stock)
    
    if isinstance(stock_diff, (int, float)):
        stock_diffs.append(stock_diff)
    
    machine_counts[machine_no] += 1

# Makine numaralarının dağılımını yazdır
print("Makine Dağılımı:")
for machine, count in machine_counts.items():
    print(f"{machine}: {count} kayıt")

# Ürün stoğunun ortalamasını hesapla
if product_stocks:
    avg_product_stock = sum(product_stocks) / len(product_stocks)
    print(f"Ortalama Ürün Stoku: {avg_product_stock}")

# Stok farkının ortalamasını hesapla
if stock_diffs:
    avg_stock_diff = sum(stock_diffs) / len(stock_diffs)
    print(f"Ortalama Stok Farkı: {avg_stock_diff}")