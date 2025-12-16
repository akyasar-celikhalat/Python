import pandas as pd

# Excel dosyasını okuma
file_name = 'Tüketim.xlsx'
df = pd.read_excel(file_name, engine='openpyxl')

# Veri yapısı ve sözlükler
product_graph = {}
product_descriptions = {}
product_machines = {}
product_times = {}
product_processes = {}

# Sütun indekslerini belirleme
columns = df.columns.tolist()
olusturma_zamani_index = columns.index('OLUŞTURMA ZAMANI')
proses_index = columns.index('PROSES')
makine_no_index = columns.index('MAKİNE NO')
giris_index = columns.index('GİRİŞ ÜRÜN SAP BARKODU')
giris_aciklama_index = columns.index('GİRİŞ ÜRÜN ACIKLAMA')
cikis_index = columns.index('SAP ETİKET BARKODU')
cikis_aciklama_index = columns.index('ÇIKIŞ ÜRÜN ACIKLAMA')

# Veriyi işleme
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
        product_processes[çıkış_barkod] = işlem
    
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

# İzlenebilirlik ağacı oluşturma fonksiyonu
def build_traceability_tree(product_code, root_code, graph, depth=0, process_chain=None):
    result = []
    if process_chain is None:
        process_chain = []
    
    ürün_aciklama = product_descriptions.get(product_code, 'Açıklama Bulunamadı')
    depth_indicator = "-" * depth
    
    # Makine, zaman ve proses bilgilerini al
    makine = product_machines.get(product_code, "")
    zaman = product_times.get(product_code, "")
    proses = product_processes.get(product_code, "")
    
    # İşlem döngüsünü oluştur
    process_path = " -> ".join(process_chain) if process_chain else ""
    
    # Ürün bilgilerini ekle
    result.append({
        'Ürün Kodu': product_code,
        'Ürün Açıklaması': f"{depth_indicator} {ürün_aciklama}" if depth > 0 else ürün_aciklama,
        'Makine No': makine,
        'Oluşturma Zamanı': zaman,
        'Proses': proses,  
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
        new_process_chain.append(f"{proses} ({product_code})")
        sub_result = build_traceability_tree(
            parent_code,
            root_code,
            graph,
            depth + 1,
            new_process_chain
        )
        result.extend(sub_result)
    
    return result

# Ana ürün kodunu veya belirli bir varyasyonu bulma ve işleme
def process_products(base_product_code=None, specific_product_code=None):
    results = []
    
    if base_product_code:
        # Tüm etiketleri bulma (örneğin, 78378300-1, 78378300-2, vb.)
        all_etiketler = [code for code in product_graph.keys() if code.startswith(base_product_code)]
        
        for etiket in all_etiketler:
            print(f"İşleniyor: {etiket}")
            trace_data = build_traceability_tree(etiket, etiket, product_graph)
            results.extend(trace_data)
    
    elif specific_product_code:
        # Belirli bir varyasyonu işleme
        if specific_product_code in product_graph:
            print(f"İşleniyor: {specific_product_code}")
            trace_data = build_traceability_tree(specific_product_code, specific_product_code, product_graph)
            results.extend(trace_data)
        else:
            print(f"Belirtilen ürün kodu ({specific_product_code}) bulunamadı.")
    
    return results

# Kullanıcıdan veri girişi alma
base_product_code = input("Tüm varyasyonlarını aramak istediğiniz ana ürün kodunu girin (örn. 77359201) (boş bırakabilirsiniz): ").strip()
specific_product_code = input("Ya da sadece belirli bir varyasyonu aramak için tam ürün kodunu girin (örn. 77359201-3) (boş bırakabilirsiniz): ").strip()

# Ağacı oluştur ve sonuçları birleştirme
trace_data = process_products(base_product_code=base_product_code, specific_product_code=specific_product_code)

# DataFrame ve Excel kaydetme
if trace_data:
    df_result = pd.DataFrame(trace_data)
    df_result = df_result[['Ürün Kodu', 'Ürün Açıklaması', 'Makine No', 'Oluşturma Zamanı', 'Proses', 'İşlem Döngüsü']] 
    output_file = f'izlenebilirlik_{base_product_code or specific_product_code}.xlsx'
    df_result.to_excel(output_file, index=False, engine='openpyxl')
    print(f"İşlem tamamlandı. Sonuçlar '{output_file}' dosyasına kaydedildi.")
else:
    print("Belirtilen ürün kodu için veri bulunamadı veya geçersiz giriş yapıldı.")