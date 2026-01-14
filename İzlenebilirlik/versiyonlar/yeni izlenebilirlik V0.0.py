import pandas as pd

def track_production_backwards(file_path, final_product_barcodes):
    """
    Tracks production backwards from final products using the provided Excel data.

    Args:
        file_path (str): Path to the Excel file.
        final_product_barcodes (list): A list of final product barcode(s) to start tracking from.

    Returns:
        pd.DataFrame: A DataFrame containing the tracked production steps.
    """
    # Load data
    df = pd.read_excel(file_path, sheet_name='VERİ')
    
    # Ensure barcodes are strings for consistent comparison
    df['GİRİŞ ÜRÜN SAP BARKODU'] = df['GİRİŞ ÜRÜN SAP BARKODU'].astype(str)
    df['TEYİT VERİLEN BARKOD'] = df['TEYİT VERİLEN BARKOD'].astype(str)
    df['SAP ETİKET BARKODU'] = df['SAP ETİKET BARKODU'].astype(str)

    # Create a mapping from output barcode to rows (as there can be multiple inputs per output)
    output_to_input_map = {}
    for _, row in df.iterrows():
        output_barkod = row['TEYİT VERİLEN BARKOD']
        if output_barkod not in output_to_input_map:
            output_to_input_map[output_barkod] = []
        output_to_input_map[output_barkod].append(row)
        
    # Also map SAP ETİKET BARKODU to rows for finding the initial step
    # Handle duplicate index by keeping the first occurrence
    try:
        sap_etiket_map = df.drop_duplicates(subset='SAP ETİKET BARKODU').set_index('SAP ETİKET BARKODU').to_dict('index')
    except ValueError: # Fallback if drop_duplicates doesn't resolve it or another issue
        # Create the map manually, keeping the first occurrence
        sap_etiket_map = {}
        for _, row in df.iterrows():
            barkod = row['SAP ETİKET BARKODU']
            if barkod not in sap_etiket_map:
                sap_etiket_map[barkod] = row

    results = []

    def find_path(current_barcodes, current_path, visited):
        """Recursively finds the path backwards."""
        for current_barcode in current_barcodes:
            if current_barcode in visited:
                continue # Prevent cycles
            visited.add(current_barcode)

            # Check if this barcode is a final product (i.e., it's an output we are looking for)
            # This handles the case where the initial barcode might not be in 'TEYİT VERİLEN BARKOD'
            # but is a 'SAP ETİKET BARKODU'. We prioritize 'TEYİT VERİLEN BARKOD' for mapping.
            rows_for_output = output_to_input_map.get(current_barcode, [])
            
            # If not found by TEYİT VERİLEN BARKOD, check if it's a direct output (SAP ETİKET BARKODU)
            if not rows_for_output and current_barcode in sap_etiket_map:
                 # This is likely a terminal node in our trace, but we still want its details
                 # We use the SAP ETİKET BARKODU row for details
                 row = sap_etiket_map[current_barcode]
                 results.append({
                     "Barkod": current_barcode,
                     "Ürün Açıklaması": row.get('ÇIKIŞ ÜRÜN ACIKLAMA', '') if pd.notna(row.get('ÇIKIŞ ÜRÜN ACIKLAMA', '')) else row.get('GİRİŞ ÜRÜN ACIKLAMA', ''),
                     "Makine": row.get('MAKİNE NO', ''),
                     "Tüketim": row.get('GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg', 0) if pd.notna(row.get('GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg')) else 0,
                     "Oluşturma Zamanı": row.get('OLUŞTURMA ZAMANI', ''),
                     "Proses": row.get('PROSES', ''),
                     "İşlem Döngüsü": " -> ".join(current_path + [current_barcode])
                 })
                 # This is a starting point, no need to go further back.
                 continue

            # If there are input rows for this output barcode, process them
            if rows_for_output:
                # Get output product info from the first representative row (they should be the same for the same output)
                representative_row = rows_for_output[0]
                results.append({
                    "Barkod": current_barcode,
                    "Ürün Açıklaması": representative_row.get('ÇIKIŞ ÜRÜN ACIKLAMA', representative_row.get('GİRİŞ ÜRÜN ACIKLAMA', '')),
                    "Makine": representative_row.get('MAKİNE NO', ''),
                    "Tüketim": representative_row.get('GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg', 0) if pd.notna(representative_row.get('GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg')) else 0,
                    "Oluşturma Zamanı": representative_row.get('OLUŞTURMA ZAMANI', ''),
                    "Proses": representative_row.get('PROSES', ''),
                    "İşlem Döngüsü": " -> ".join(current_path + [current_barcode])
                })
                
                # Find all input barcodes for this step and recurse
                input_barcodes = [row['GİRİŞ ÜRÜN SAP BARKODU'] for row in rows_for_output if pd.notna(row['GİRİŞ ÜRÜN SAP BARKODU'])]
                find_path(input_barcodes, current_path + [current_barcode], visited)
            else:
                # If no info is found for this barcode, still add it to the list
                # (can be useful for data inconsistency cases)
                 results.append({
                     "Barkod": current_barcode,
                     "Ürün Açıklaması": "Bilinmiyor",
                     "Makine": "Bilinmiyor",
                     "Tüketim": 0,
                     "Oluşturma Zamanı": "",
                     "Proses": "Bilinmiyor",
                     "İşlem Döngüsü": " -> ".join(current_path + [current_barcode])
                 })

    # Start tracking for each final product barcode
    for barcode in final_product_barcodes:
        find_path([barcode], [], set())
        
    # Convert results to DataFrame and sort by the length of the process cycle
    # (longest cycle first)
    result_df = pd.DataFrame(results)
    if not result_df.empty:
        # Sort by the length of the process cycle (descending)
        result_df['Döngü_Uzunluğu'] = result_df['İşlem Döngüsü'].str.count(' -> ')
        result_df.sort_values(by='Döngü_Uzunluğu', ascending=False, inplace=True)
        result_df.drop('Döngü_Uzunluğu', axis=1, inplace=True)
        result_df.reset_index(drop=True, inplace=True)
        
    return result_df

# Dosya yolu ve izlemeye başlanacak son ürün barkodu
file_path = '79528600-33.xlsx' # Yüklediğiniz dosya adı
final_barcodes = ['79528600-33'] # Dosya adından gelen barkod

# Fonksiyonu çağır ve sonucu al
df_sonuc = track_production_backwards(file_path, final_barcodes)

# Sonucu yazdır
print("Ham sonucun ilk 5 satırı:")
print(df_sonuc.head().to_string(index=False, justify='center'))
print("\n" + "="*50 + "\n")

# İşlem sırasını ters çevir (ilk satır en uzun döngüyü yani son ürünü göstersin)
df_sonuc_ters = df_sonuc.copy()
df_sonuc_ters = df_sonuc_ters.iloc[::-1] # Tüm satırları ters sırala
df_sonuc_ters.reset_index(drop=True, inplace=True) # İndeksleri sıfırla

# İşlem Döngüsünü de ters çevir (İlk adım en başta olsun)
def reverse_cycle(cycle_str):
    parts = cycle_str.split(' -> ')
    return ' -> '.join(reversed(parts))

df_sonuc_ters['İşlem Döngüsü'] = df_sonuc_ters['İşlem Döngüsü'].apply(reverse_cycle)

# Sonucu bir Excel dosyasına kaydet
output_file = 'uretim_izleme_sonucu_ters_sira.xlsx'
df_sonuc_ters.to_excel(output_file, index=False)

print("İşlem sırası ters çevrilmiş sonuç:")
print(df_sonuc_ters.to_string(index=False, justify='center'))
print(f"\nSonuç '{output_file}' dosyasına kaydedildi.")
