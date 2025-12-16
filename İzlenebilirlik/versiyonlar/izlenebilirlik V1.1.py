import pandas as pd

# Excel dosyasını okuma
file_name = 'Tüketim.xlsx'
df = pd.read_excel(file_name, engine='openpyxl')

# Parent-Child ilişkilerini saklamak için veri yapısı
product_graph = {}

# Ürün açıklamalarını saklamak için sözlük
product_descriptions = {}

# Sütun isimlerini al
columns = df.columns.tolist()
giris_index = columns.index('GİRİŞ ÜRÜN SAP BARKODU')
giris_aciklama_index = columns.index('GİRİŞ ÜRÜN ACIKLAMA')
cikis_index = columns.index('TEYİT VERİLEN BARKOD')
cikis_aciklama_index = columns.index('ÇIKIŞ ÜRÜN ACIKLAMA')
proses_index = columns.index('PROSES')

# Veri çerçevesini dolaşma
for row in df.values:
    giriş_barkod = row[giris_index]
    giriş_aciklama = row[giris_aciklama_index]
    çıkış_barkod = row[cikis_index]
    çıkış_aciklama = row[cikis_aciklama_index]
    işlem = row[proses_index]
    
    # Çıkış barkodunun parent'larını kaydet
    if çıkış_barkod not in product_graph:
        product_graph[çıkış_barkod] = []
    product_graph[çıkış_barkod].append({
        'parent': giriş_barkod,
        'process': işlem
    })
    
    # Ürün açıklamalarını kaydet
    if çıkış_barkod not in product_descriptions:
        product_descriptions[çıkış_barkod] = çıkış_aciklama
    
    if giriş_barkod not in product_descriptions:
        product_descriptions[giriş_barkod] = giriş_aciklama

# Belirli bir ürün için izlenebilirlik bilgilerini almak için fonksiyon
def get_product_trace(product_code, graph):
    all_paths = []
    
    # BFS ile tüm yolları bulalım
    queue = [(product_code, [])]  # (current_product, path)
    
    while queue:
        current, path = queue.pop(0)
        
        # Her ürün ve yol çiftini sonuçlara ekle
        all_paths.append((current, path))
        
        # Eğer bu ürün, başka bir ürünün üretiminde kullanılıyorsa
        if current in graph:
            for entry in graph[current]:
                parent = entry['parent']
                process = entry['process']
                new_path = path + [(process, parent)]
                queue.append((parent, new_path))
    
    return all_paths

# Belirli bir ürün kodu için izlenebilirlik bilgilerini alma
product_code = '67301600-59'
trace = get_product_trace(product_code, product_graph)

# Sonuçları DataFrame'e dönüştürme
data = []
for item in trace:
    ürün_kodu = item[0]
    path = item[1]
    process_path = " -> ".join([f"{p} ({b})" for p, b in path])
    ürün_aciklama = product_descriptions.get(ürün_kodu, "Açıklama Bulunamadı")
    
    # Adım sayısını belirlemek için
    adim_sayisi = len(path)
    adim_belirtici = "-" * adim_sayisi
    
    # Ürün açıklamasını güncelle
    ürün_aciklama_güncellenmiş = f"{adim_belirtici} {ürün_aciklama}"
    
    data.append({
        'Ürün Kodu': ürün_kodu,
        'Ürün Açıklaması': ürün_aciklama_güncellenmiş,
        'İşlem Döngüsü': process_path
    })

# DataFrame oluşturma
df_result = pd.DataFrame(data)

# Sonuçları Excel dosyasına yazma
output_file = f'izlenebilirlik_{product_code}.xlsx'
df_result.to_excel(output_file, index=False, engine='openpyxl')

print(f"İşlem tamamlandı. Sonuçlar '{output_file}' dosyasına kaydedildi.")