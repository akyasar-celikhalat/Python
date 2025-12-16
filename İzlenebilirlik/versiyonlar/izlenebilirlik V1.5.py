import pandas as pd

# Excel dosyasını okuma
file_name = 'Tüketim.xlsx'
df = pd.read_excel(file_name, engine='openpyxl')

# Parent-Child ilişkilerini saklamak için veri yapısı
product_graph = {}

# Ürün bilgilerini saklamak için sözlükler
product_descriptions = {}
product_machines = {}
product_times = {}

# Sütun isimlerini al
columns = df.columns.tolist()
giris_index = columns.index('GİRİŞ ÜRÜN SAP BARKODU')
giris_aciklama_index = columns.index('GİRİŞ ÜRÜN ACIKLAMA')
cikis_index = columns.index('TEYİT VERİLEN BARKOD')
cikis_aciklama_index = columns.index('ÇIKIŞ ÜRÜN ACIKLAMA')
proses_index = columns.index('PROSES')
makine_no_index = columns.index('MAKİNE NO')
olusturma_zamani_index = columns.index('OLUŞTURMA ZAMANI')

# Veri çerçevesini dolaşma
for row in df.values:
    giriş_barkod = row[giris_index]
    giriş_aciklama = row[giris_aciklama_index]
    çıkış_barkod = row[cikis_index]
    çıkış_aciklama = row[cikis_aciklama_index]
    işlem = row[proses_index]
    makine_no = row[makine_no_index]
    olusturma_zamani = row[olusturma_zamani_index]
    
    # Çıkış ürününün makine ve zamanını kaydet
    if çıkış_barkod not in product_machines:
        product_machines[çıkış_barkod] = makine_no
        product_times[çıkış_barkod] = olusturma_zamani
    
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
def build_traceability_tree(product_code, root_code, graph, depth=0, process_chain=None):
    result = []
    
    # İlk çağrıda process_chain'i başlat
    if process_chain is None:
        process_chain = []
    
    # Ürün açıklamasını al
    ürün_aciklama = product_descriptions.get(product_code, 'Açıklama Bulunamadı')
    
    # Derinlik seviyesini göstermek için tire işareti
    depth_indicator = "-" * depth
    
    # Ürün için makine ve zaman bilgilerini al
    makine = product_machines.get(product_code, "")
    zaman = product_times.get(product_code, "")
    
    # İşlem döngüsünü oluştur - kök üründen başlayarak, şu anki ürüne kadar olan path
    process_path = ""
    if process_chain:
        process_path = " -> ".join(process_chain)
    
    # Ürün bilgilerini ekle
    result.append({
        'Ürün Kodu': product_code,
        'Ürün Açıklaması': f"{depth_indicator} {ürün_aciklama}" if depth > 0 else ürün_aciklama,
        'Makine No': makine,
        'Oluşturma Zamanı': zaman,
        'İşlem Döngüsü': process_path
    })
    
    # Eğer bu ürün için giriş ürünleri yoksa bitir
    if product_code not in graph:
        return result
    
    # Ürünün parent'larını ekle
    for entry in graph[product_code]:
        parent_code = entry['parent']
        process = entry['process']
        
        # Yeni process chain oluştur - process_chain'e mevcut proses bilgisini ekle
        new_process_chain = process_chain.copy()
        new_process_chain.append(f"{process} ({parent_code})")
        
        # Alt ürünleri recursive olarak ekle
        sub_result = build_traceability_tree(
            parent_code, 
            root_code,
            graph, 
            depth + 1, 
            new_process_chain
        )
        
        result.extend(sub_result)
    
    return result

# İşlem döngüsünü doğru formatta oluşturmak için yardımcı fonksiyon
def format_process_chain(product_code, target_code, graph, process_chain=None, visited=None):
    if process_chain is None:
        process_chain = []
    
    if visited is None:
        visited = set()
    
    # Döngüleri önle
    if product_code in visited:
        return []
    
    visited.add(product_code)
    
    # Hedef ürüne ulaştık
    if product_code == target_code:
        return process_chain
    
    # Bu ürün için işlem zincirlerini deneme
    if product_code in graph:
        for entry in graph[product_code]:
            parent_code = entry['parent']
            process = entry['process']
            
            # Yeni zincir oluştur
            new_chain = process_chain.copy()
            new_chain.append(f"{process} ({parent_code})")
            
            # Recursive olarak alt zinciri kontrol et
            found_chain = format_process_chain(parent_code, target_code, graph, new_chain, visited.copy())
            if found_chain:
                return found_chain
    
    return []

# Ürün ağacını dolaşmak ve tüm yolları bulmak için fonksiyon
def build_complete_tree(root_code, graph):
    result = []
    stack = [(root_code, 0, [])]  # (ürün_kodu, derinlik, işlem_zinciri)
    
    while stack:
        product_code, depth, chain = stack.pop()
        
        # Ürün açıklamasını al
        ürün_aciklama = product_descriptions.get(product_code, 'Açıklama Bulunamadı')
        
        # Derinlik seviyesini göstermek için tire işareti
        depth_indicator = "-" * depth
        
        # Ürün için makine ve zaman bilgilerini al
        makine = product_machines.get(product_code, "")
        zaman = product_times.get(product_code, "")
        
        # İşlem döngüsünü oluştur
        process_path = ""
        if chain:
            # Kök ürünün kodunu her süreçteki işlem zincirinin başına ekle
            formatted_chain = []
            for i, step in enumerate(chain):
                if i == 0:
                    formatted_chain.append(f"DH ({root_code})")
                else:
                    formatted_chain.append(step)
            process_path = " -> ".join(formatted_chain)
        
        # Ürün bilgilerini ekle
        result.append({
            'Ürün Kodu': product_code,
            'Ürün Açıklaması': f"{depth_indicator} {ürün_aciklama}" if depth > 0 else ürün_aciklama,
            'Makine No': makine,
            'Oluşturma Zamanı': zaman,
            'İşlem Döngüsü': process_path
        })
        
        # Eğer bu ürün için giriş ürünleri varsa ekle
        if product_code in graph:
            for entry in graph[product_code]:
                parent_code = entry['parent']
                process = entry['process']
                
                # Yeni zincir oluştur
                new_chain = chain.copy()
                if not new_chain:  # İlk seviye için özel işlem
                    new_chain.append(f"TCD ({parent_code})")
                else:
                    new_chain.append(f"TCD ({parent_code})")
                
                # Yığına ekle
                stack.append((parent_code, depth + 1, new_chain))
    
    return result

# Ana ürün kodu
product_code = '76769200-100'

# Yeni yaklaşımla ağacı oluştur
trace_data = []
first_item = {
    'Ürün Kodu': product_code,
    'Ürün Açıklaması': product_descriptions.get(product_code, 'Açıklama Bulunamadı'),
    'Makine No': product_machines.get(product_code, ""),
    'Oluşturma Zamanı': product_times.get(product_code, ""),
    'İşlem Döngüsü': ''
}
trace_data.append(first_item)

# İlk seviye child'ları ekle
if product_code in product_graph:
    for entry in product_graph[product_code]:
        child_code = entry['parent']
        process = entry['process']
        
        # Child bilgilerini al
        child_desc = product_descriptions.get(child_code, 'Açıklama Bulunamadı')
        child_machine = product_machines.get(child_code, "")
        child_time = product_times.get(child_code, "")
        
        child_item = {
            'Ürün Kodu': child_code,
            'Ürün Açıklaması': f"- {child_desc}",
            'Makine No': child_machine,
            'Oluşturma Zamanı': child_time,
            'İşlem Döngüsü': f"DH ({product_code})"
        }
        trace_data.append(child_item)
        
        # Diğer seviyeleri recursive olarak dolaş
        process_so_far = [f"DH ({product_code})"]
        
        if child_code in product_graph:
            stack = [(child_code, 2, process_so_far, f"TCD ({child_code})")]
            
            while stack:
                current_code, depth, process_chain, last_process = stack.pop()
                
                for entry in product_graph[current_code]:
                    sub_child = entry['parent']
                    sub_process = entry['process']
                    
                    # Alt child bilgilerini al
                    sub_desc = product_descriptions.get(sub_child, 'Açıklama Bulunamadı')
                    sub_machine = product_machines.get(sub_child, "")
                    sub_time = product_times.get(sub_child, "")
                    
                    # İşlem zincirini güncelle
                    current_chain = process_chain.copy()
                    current_chain.append(last_process)
                    
                    sub_item = {
                        'Ürün Kodu': sub_child,
                        'Ürün Açıklaması': f"{'-' * depth} {sub_desc}",
                        'Makine No': sub_machine,
                        'Oluşturma Zamanı': sub_time,
                        'İşlem Döngüsü': " -> ".join(current_chain)
                    }
                    trace_data.append(sub_item)
                    
                    # Eğer alt child'ın da parent'ları varsa onları da ekle
                    if sub_child in product_graph:
                        stack.append((sub_child, depth + 1, current_chain, f"TV ({sub_child})"))

# DataFrame oluştur - sütun sıralamasını ayarla
df_result = pd.DataFrame(trace_data)
df_result = df_result[['Ürün Kodu', 'Ürün Açıklaması', 'Makine No', 'Oluşturma Zamanı', 'İşlem Döngüsü']]

# Excel dosyasına kaydet
output_file = f'izlenebilirlik_{product_code}.xlsx'
df_result.to_excel(output_file, index=False, engine='openpyxl')

print(f"İşlem tamamlandı. Sonuçlar '{output_file}' dosyasına kaydedildi.")