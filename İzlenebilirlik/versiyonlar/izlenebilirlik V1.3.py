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
def build_traceability_tree(product_code, graph, depth=0, parent_id=None, path_id=0, path=[]):
    result = []
    
    # Ürün açıklamasını al
    ürün_aciklama = product_descriptions.get(product_code, 'Açıklama Bulunamadı')
    adim_belirtici = "-" * depth
    
    # Ürün bilgilerini ekle
    result.append({
        'Ürün Kodu': product_code,
        'Ürün Açıklaması': f"{adim_belirtici} {ürün_aciklama}",
        'İşlem Döngüsü': " -> ".join(path),
        'depth': depth,
        'parent_id': parent_id,
        'path_id': path_id
    })
    
    # Eğer bu ürün için giriş ürünleri yoksa bitir
    if product_code not in graph:
        return result
    
    # Ürünün parent'larını ekle
    for i, entry in enumerate(graph[product_code]):
        parent_code = entry['parent']
        process = entry['process']
        
        # Yeni path oluştur
        new_path = path.copy()
        new_path.append(f"{process} ({parent_code})")
        
        # Alt ürünleri recursive olarak ekle
        sub_result = build_traceability_tree(
            parent_code, 
            graph, 
            depth + 1, 
            parent_id=len(result) - 1, 
            path_id=path_id, 
            path=new_path
        )
        
        result.extend(sub_result)
    
    return result

# Belirli bir ürün kodu için izlenebilirlik bilgilerini alma
product_code = '77608000-8'
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