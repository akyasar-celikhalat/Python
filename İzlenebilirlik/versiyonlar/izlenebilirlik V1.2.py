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

# Belirli bir ürün için izlenebilirlik ağacını oluşturacak fonksiyon
def build_traceability_tree(product_code, graph):
    result = []
    
    # Kök ürünü ekle
    result.append({
        'Ürün Kodu': product_code,
        'Ürün Açıklaması': f" {product_descriptions.get(product_code, 'Açıklama Bulunamadı')}",
        'İşlem Döngüsü': '',
        'depth': 0,
        'parent_id': None,
        'path_id': 0
    })
    
    # Eğer bu ürün için giriş ürünleri yoksa bitir
    if product_code not in graph:
        return result
    
    path_counter = 1
    
    # Her bir giriş ürününü tara
    for i, entry in enumerate(graph[product_code]):
        parent_id = 0  # Kök ürünün ID'si
        current_path_id = path_counter
        path_counter += 1
        
        # İlk seviye giriş ürünlerini ekle
        parent_code = entry['parent']
        process = entry['process']
        
        result.append({
            'Ürün Kodu': parent_code,
            'Ürün Açıklaması': f"- {product_descriptions.get(parent_code, 'Açıklama Bulunamadı')}",
            'İşlem Döngüsü': f"{process} ({parent_code})",
            'depth': 1,
            'parent_id': parent_id,
            'path_id': current_path_id
        })
        
        # Ürünün alt ürünlerini recursive olarak ekle
        if parent_code in graph:
            sub_path_counter = 0
            for sub_entry in graph[parent_code]:
                sub_parent_code = sub_entry['parent']
                sub_process = sub_entry['process']
                
                # Mevcut sub_entry'nin indeksini bul 
                current_index = len(result) - 1
                
                result.append({
                    'Ürün Kodu': sub_parent_code,
                    'Ürün Açıklaması': f"-- {product_descriptions.get(sub_parent_code, 'Açıklama Bulunamadı')}",
                    'İşlem Döngüsü': f"{process} ({parent_code}) -> {sub_process} ({sub_parent_code})",
                    'depth': 2,
                    'parent_id': current_index,
                    'path_id': current_path_id
                })
                sub_path_counter += 1
    
    # Sonuçları sırala - önce ana yola (path_id) göre, sonra derinliğe göre
    result.sort(key=lambda x: (x['path_id'], x['depth']))
    
    return result

# Belirli bir ürün kodu için izlenebilirlik bilgilerini alma
product_code = '77518600-2'
trace_data = build_traceability_tree(product_code, product_graph)

# DataFrame için sadece gerekli alanları al
data = [{
    'Ürün Kodu': item['Ürün Kodu'],
    'Ürün Açıklaması': item['Ürün Açıklaması'],
    'İşlem Döngüsü': item['İşlem Döngüsü']
} for item in trace_data]

# DataFrame oluşturma
df_result = pd.DataFrame(data)

# Sonuçları Excel dosyasına yazma
output_file = f'izlenebilirlik_{product_code}.xlsx'
df_result.to_excel(output_file, index=False, engine='openpyxl')

print(f"İşlem tamamlandı. Sonuçlar '{output_file}' dosyasına kaydedildi.")