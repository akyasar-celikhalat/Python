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
product_quantities = {} # TEYİT MİKTARI Metre bilgisi için

# Sütun indekslerini belirleme
columns = df.columns.tolist()
olusturma_zamani_index = columns.index('OLUŞTURMA ZAMANI')
proses_index = columns.index('PROSES')
makine_no_index = columns.index('MAKİNE NO')
giris_index = columns.index('GİRİŞ ÜRÜN SAP BARKODU')
giris_aciklama_index = columns.index('GİRİŞ ÜRÜN ACIKLAMA')
cikis_index = columns.index('SAP ETİKET BARKODU')
cikis_aciklama_index = columns.index('ÇIKIŞ ÜRÜN ACIKLAMA')
teyit_miktari_index = columns.index('TEYİT MİKTARI Metre') # Yeni sütun

# Veriyi işleme - GRAPH YAPISINI TERS ÇEVİR: GİRİŞ'TEN ÇIKIŞ'A DOĞRU
for row in df.values:
    giriş_barkod = row[giris_index]
    giriş_aciklama = row[giris_aciklama_index]
    çıkış_barkod = row[cikis_index]
    çıkış_aciklama = row[cikis_aciklama_index]
    işlem = row[proses_index]
    makine_no = row[makine_no_index]
    olusturma_zamani = row[olusturma_zamani_index]
    teyit_miktari = row[teyit_miktari_index]

    # GİRİŞ ürününün bilgilerini de kaydet (mamul olmayan ürünler için)
    if giriş_barkod not in product_descriptions:
        product_descriptions[giriş_barkod] = giriş_aciklama

    # ÇIKIŞ ürününün diğer bilgilerini kaydet
    if çıkış_barkod not in product_machines:
        product_machines[çıkış_barkod] = makine_no
        product_times[çıkış_barkod] = olusturma_zamani
        product_processes[çıkış_barkod] = işlem
        product_quantities[çıkış_barkod] = teyit_miktari # Miktarı da kaydet
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

    # Makine, zaman, proses ve miktar bilgilerini al
    # Root hammadde için bu bilgiler olmayabilir
    makine = product_machines.get(product_code, "")
    zaman = product_times.get(product_code, "")
    proses = product_processes.get(product_code, "")
    miktar = product_quantities.get(product_code, "") # Yeni alan

    # İşlem döngüsünü oluştur
    process_path = " -> ".join(process_chain) if process_chain else ""

    # Ürün bilgilerini ekle
    result.append({
        'Barkod': product_code,
        'Ürün Açıklaması': f"{depth_indicator} {ürün_aciklama}" if depth > 0 else ürün_aciklama,
        'Makine': makine,
        'Oluşturma Zamanı': zaman,
        'Proses': proses,
        'Miktar (Metre)': miktar, # Yeni alan
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


def process_products_from_raw_material_list(raw_material_codes):
    results = []
    for code in raw_material_codes:
        code = code.strip()
        if code in product_graph:
            print(f"İşleniyor: {code}")
            trace_data = build_traceability_tree_from_raw_material(code, code, product_graph)
            results.extend(trace_data)
        else:
            print(f"Belirtilen hammadde kodu ({code}) bulunamadı veya herhangi bir prosesin girdisi değil.")
    return results


# Kullanıcıdan virgülle ayrılmış hammadde kodlarını alma
raw_material_input = input("Hammadde ürün kodlarını virgülle ayırarak girin (örn. 77359201-1, 40001234-5): ").strip()

if not raw_material_input:
    print("Hiçbir kod girilmedi.")
    exit()

raw_material_codes = [code.strip() for code in raw_material_input.split(',')]

# Ağacı oluştur ve sonuçları birleştirme
trace_data = process_products_from_raw_material_list(raw_material_codes)

# DataFrame ve Excel kaydetme
if trace_data:
    df_result = pd.DataFrame(trace_data)
    df_result = df_result.rename(columns={
        'Barkod': 'Ürün Kodu',
        'Makine': 'Makine No',
        'Oluşturma Zamanı': 'Oluşturma Zamanı',
        'Proses': 'Proses',
        'İşlem Döngüsü': 'İşlem Döngüsü',
        'Ürün Açıklaması': 'Ürün Açıklaması',
        'Miktar (Metre)': 'Miktar' # Yeni sütun ismi
    })
    # Miktar sütununu uygun yere yerleştir
    df_result = df_result[['Ürün Kodu', 'Ürün Açıklaması', 'Miktar', 'Makine No', 'Oluşturma Zamanı', 'Proses', 'İşlem Döngüsü']]
    output_file = f'hammadde_mamul_izlenebilirlik_{len(raw_material_codes)}_hammadde.xlsx'
    df_result.to_excel(output_file, index=False, engine='openpyxl')
    print(f"İşlem tamamlandı. Sonuçlar '{output_file}' dosyasına kaydedildi.")
else:
    print("Belirtilen ürün kodları için veri bulunamadı veya geçersiz giriş yapıldı.")