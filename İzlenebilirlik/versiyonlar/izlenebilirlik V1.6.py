import pandas as pd

file_name = 'Tüketim.xlsx'
df = pd.read_excel(file_name, engine='openpyxl')

# Veri yapısı ve sözlükler
product_graph = {}
product_descriptions = {}
product_machines = {}
product_times = {}
product_processes = {}  # Prosesserin depolanacağı yeni sözlük

columns = df.columns.tolist()
giris_index = columns.index('GİRİŞ ÜRÜN SAP BARKODU')
giris_aciklama_index = columns.index('GİRİŞ ÜRÜN ACIKLAMA')
cikis_index = columns.index('TEYİT VERİLEN BARKOD')
cikis_aciklama_index = columns.index('ÇIKIŞ ÜRÜN ACIKLAMA')
proses_index = columns.index('PROSES')
makine_no_index = columns.index('MAKİNE NO')
olusturma_zamani_index = columns.index('OLUŞTURMA ZAMANI')

for row in df.values:
    giriş_barkod = row[giris_index]
    giriş_aciklama = row[giris_aciklama_index]
    çıkış_barkod = row[cikis_index]
    çıkış_aciklama = row[cikis_aciklama_index]
    işlem = row[proses_index]
    makine_no = row[makine_no_index]
    olusturma_zamani = row[olusturma_zamani_index]
    
    # Çıkış ürününün makine, zaman ve proses bilgilerini kaydet
    if çıkış_barkod not in product_machines:
        product_machines[çıkış_barkod] = makine_no
        product_times[çıkış_barkod] = olusturma_zamani
        product_processes[çıkış_barkod] = işlem  # Prosese ekleme
    
    # Çıkış barkodunun parent'larını kaydet
    if çıkış_barkod not in product_graph:
        product_graph[çıkış_barkod] = []
    product_graph[çıkış_barkod].append({
        'parent': giriş_barkod,
        'process': işlem  # Prosese ekleme
    })
    
    # Ürün açıklamalarını kaydet
    if çıkış_barkod not in product_descriptions:
        product_descriptions[çıkış_barkod] = çıkış_aciklama
    if giriş_barkod not in product_descriptions:
        product_descriptions[giriş_barkod] = giriş_aciklama

# Izlenebilirlik ağacı oluşturma fonksiyonu
def build_traceability_tree(product_code, root_code, graph, depth=0, process_chain=None):
    result = []
    if process_chain is None:
        process_chain = []
    
    ürün_aciklama = product_descriptions.get(product_code, 'Açıklama Bulunamadı')
    depth_indicator = "-" * depth
    
    # Ürünün kendi proses bilgisini al
    proses = product_processes.get(product_code, "")
    
    # Makine ve zaman bilgilerini al
    makine = product_machines.get(product_code, "")
    zaman = product_times.get(product_code, "")
    
    # İşlem döngüsünü oluştur
    process_path = " -> ".join(process_chain) if process_chain else ""
    
    # Ürün bilgilerini ekle
    result.append({
        'Ürün Kodu': product_code,
        'Ürün Açıklaması': f"{depth_indicator} {ürün_aciklama}" if depth > 0 else ürün_aciklama,
        'Makine No': makine,
        'Oluşturma Zamanı': zaman,
        'İşlem Döngüsü': process_path
    })
    
    # Eğer ürünün parent'ları yoksa döndür
    if product_code not in graph:
        return result
    
    # Parent'ları ekle
    for entry in graph[product_code]:
        parent_code = entry['parent']
        # Prosese ekleme: ürünün kendi prosesini kullan
        new_process_chain = process_chain.copy()
        new_process_chain.append(f"{proses} ({product_code})")  # KENDİ PROSESİ KULLANILDI
        sub_result = build_traceability_tree(
            parent_code,
            root_code,
            graph,
            depth + 1,
            new_process_chain
        )
        result.extend(sub_result)
    
    return result

# Ana ürün kodu
product_code = '77608000-7'

# Ağacı oluştur
trace_data = build_traceability_tree(product_code, product_code, product_graph)

# DataFrame ve Excel kaydetme
df_result = pd.DataFrame(trace_data)
df_result = df_result[['Ürün Kodu', 'Ürün Açıklaması', 'Makine No', 'Oluşturma Zamanı', 'İşlem Döngüsü']]
output_file = f'izlenebilirlik_{product_code}.xlsx'
df_result.to_excel(output_file, index=False, engine='openpyxl')
print(f"İşlem tamamlandı. Sonuçlar '{output_file}' dosyasına kaydedildi.")