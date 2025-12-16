import pandas as pd

# Excel dosyasını okuma
file_name = 'kardemir.xlsx'
df = pd.read_excel(file_name, engine='openpyxl')

# Veri yapısı ve sözlükler
product_graph = {}  # Bu sefer giriş ürününden çıkış ürününü bulmak için
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

# Veriyi işleme - GRAPH YAPISINI TERS ÇEVİR: GİRİŞ'TEN ÇIKIŞ'A DOĞRU
for row in df.values:
    giriş_barkod = row[giris_index]
    giriş_aciklama = row[giris_aciklama_index]
    çıkış_barkod = row[cikis_index]
    çıkış_aciklama = row[cikis_aciklama_index]
    işlem = row[proses_index]
    makine_no = row[makine_no_index]
    olusturma_zamani = row[olusturma_zamani_index]

    # GİRİŞ ürününün bilgilerini de kaydet (mamul olmayan ürünler için)
    if giriş_barkod not in product_descriptions:
        product_descriptions[giriş_barkod] = giriş_aciklama

    # ÇIKIŞ ürününün makine, zaman ve proses bilgilerini kaydet
    if çıkış_barkod not in product_machines:
        product_machines[çıkış_barkod] = makine_no
        product_times[çıkış_barkod] = olusturma_zamani
        product_processes[çıkış_barkod] = işlem
    if çıkış_barkod not in product_descriptions:
        product_descriptions[çıkış_barkod] = çıkış_aciklama

    # GİRİŞ barkodunun child'larını kaydet (tersine çevirilmiş graf)
    if giriş_barkod not in product_graph:
        product_graph[giriş_barkod] = []
    product_graph[giriş_barkod].append({
        'child': çıkış_barkod,
        'process': işlem
    })

# HAMMADDEDEN MAMULE DOĞRU İZLEYEBİLMEK İÇİN AĞAÇ OLUŞTURMA FONKSİYONU
def build_traceability_tree_from_raw_material(product_code, root_code, graph, depth=0, process_chain=None):
    result = []
    if process_chain is None:
        process_chain = []

    ürün_aciklama = product_descriptions.get(product_code, 'Açıklama Bulunamadı')
    depth_indicator = "-" * depth

    # Makine, zaman ve proses bilgilerini al
    # Root hammadde için bu bilgiler olmayabilir
    makine = product_machines.get(product_code, "")
    zaman = product_times.get(product_code, "")
    proses = product_processes.get(product_code, "")

    # İşlem döngüsünü oluştur
    process_path = " -> ".join(process_chain) if process_chain else ""

    # Ürün bilgilerini ekle
    result.append({
        'Barkod': product_code,
        'Ürün Açıklaması': f"{depth_indicator} {ürün_aciklama}" if depth > 0 else ürün_aciklama,
        'Makine': makine,
        'Oluşturma Zamanı': zaman,
        'Proses': proses,
        'İşlem Döngüsü': process_path
    })

    # Eğer ürünün child'ları (üretilen ürünler) yoksa döndür
    if product_code not in graph:
        return result

    # Child'ları ekle (üretilen ürünler)
    for entry in graph[product_code]:
        child_code = entry['child']
        child_process = entry['process']
        # Prosese ekleme: child'ın prosesini kullan
        new_process_chain = process_chain.copy()
        new_process_chain.append(f"{child_process} ({child_code})")
        sub_result = build_traceability_tree_from_raw_material(
            child_code,
            root_code,
            graph,
            depth + 1,
            new_process_chain
        )
        result.extend(sub_result)

    return result

# Ana ürün kodunu veya belirli bir varyasyonu bulma ve işleme
def find_root_materials():
    # Tüm giriş (giriş ürünleri) ve çıkış (çıkış ürünleri) kodlarını topla
    all_giris_codes = set(product_graph.keys())
    all_cikis_codes = set()
    for v_list in product_graph.values():
        for entry in v_list:
            all_cikis_codes.add(entry['child'])

    # Çıkış olan ama giriş olmayanlar (yani son ürünler)
    # VEYA Giriş olan ama çıkış olmayanlar (yani hammadde)
    # Bize hammadde lazım: GİRİŞ olan ama asla ÇIKIŞ olarak kullanılmayanlar
    # YANİ: all_giris_codes içinde ama all_cikis_codes içinde OLMAYANLAR hammadde olabilir
    # AMA BU KODUN ÇOCUKLARI OLMALI (yani bir üretimde kullanılmış olmalı)

    # root materials: bir prosesin girdisi olup ama başka bir prosesin çıktısı olmayan ürünler
    root_materials = all_giris_codes.difference(all_cikis_codes)
    # Ancak bu root material'ların çocukları olmalı ki anlamlı olsun
    filtered_root_materials = [rm for rm in root_materials if rm in product_graph]
    return filtered_root_materials

def process_products_from_raw_material(base_product_code=None, specific_product_code=None):
    results = []

    if base_product_code:
        # Tüm varyasyonları bulma (örneğin, 78378300-1, 78378300-2, vb.)
        # Burada sadece giriş ürünlerinden (hammadde) başlayabiliriz
        all_input_codes = [code for code in product_graph.keys() if code.startswith(base_product_code)]
        for input_code in all_input_codes:
            print(f"İşleniyor: {input_code}")
            trace_data = build_traceability_tree_from_raw_material(input_code, input_code, product_graph)
            results.extend(trace_data)

    elif specific_product_code:
        # Belirli bir hammaddeyi takip et
        if specific_product_code in product_graph:
            print(f"İşleniyor: {specific_product_code}")
            trace_data = build_traceability_tree_from_raw_material(specific_product_code, specific_product_code, product_graph)
            results.extend(trace_data)
        else:
            print(f"Belirtilen hammadde kodu ({specific_product_code}) bulunamadı veya herhangi bir prosesin girdisi değil.")

    return results

# Hammadde kodu girişi
print("Hammadde ürünleri:")
roots = find_root_materials()
print(roots)

specific_product_code = input("Hammadde ürün kodunu girin (örn. 77359201-1) (boş bırakırsanız örnek hammadde kodları gösterilir): ").strip()

if not specific_product_code:
    print("Örnek hammadde kodları:")
    for rm in roots:
        print(rm)
    specific_product_code = input("Birini seçin ve girin: ").strip()

# Ağacı oluştur ve sonuçları birleştirme
trace_data = process_products_from_raw_material(specific_product_code=specific_product_code)

# DataFrame ve Excel kaydetme
if trace_data:
    # Sütun isimlerini orijinal koddaki gibi düzelt
    df_result = pd.DataFrame(trace_data)
    df_result = df_result.rename(columns={
        'Barkod': 'Ürün Kodu',
        'Makine': 'Makine No',
        'Oluşturma Zamanı': 'Oluşturma Zamanı',
        'Proses': 'Proses',
        'İşlem Döngüsü': 'İşlem Döngüsü',
        'Ürün Açıklaması': 'Ürün Açıklaması'
    })
    df_result = df_result[['Ürün Kodu', 'Ürün Açıklaması', 'Makine No', 'Oluşturma Zamanı', 'Proses', 'İşlem Döngüsü']]
    output_file = f'hammadde_mamul_izlenebilirlik_{specific_product_code}.xlsx'
    df_result.to_excel(output_file, index=False, engine='openpyxl')
    print(f"İşlem tamamlandı. Sonuçlar '{output_file}' dosyasına kaydedildi.")
else:
    print("Belirtilen ürün kodu için veri bulunamadı veya geçersiz giriş yapıldı.")