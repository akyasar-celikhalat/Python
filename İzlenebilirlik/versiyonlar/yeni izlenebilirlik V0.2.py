import pandas as pd
import os
from datetime import datetime

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
    # Create a mapping from input barcode to rows for finding raw material descriptions
    input_barkod_to_row_map = {} 
    
    for _, row in df.iterrows():
        # Map for output tracking
        output_barkod = row['TEYİT VERİLEN BARKOD']
        if output_barkod not in output_to_input_map:
            output_to_input_map[output_barkod] = []
        output_to_input_map[output_barkod].append(row)
        
        # Map for input description lookup (store the first occurrence)
        input_barkod = row['GİRİŞ ÜRÜN SAP BARKODU']
        if input_barkod not in input_barkod_to_row_map:
             input_barkod_to_row_map[input_barkod] = row

    # Also map SAP ETİKET BARKODU to rows for finding the initial step
    # Handle duplicate index by keeping the first occurrence
    try:
        sap_etiket_map = df.drop_duplicates(subset='SAP ETİKET BARKODU').set_index('SAP ETİKET BARKODU').to_dict('index')
    except ValueError:
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
                 
                 # For raw materials, try to get description from input_barkod_to_row_map first
                 urun_aciklamasi = row.get('GİRİŞ ÜRÜN ACIKLAMA', '')
                 if not urun_aciklamasi or str(urun_aciklamasi).lower() == "nan":
                      # Fallback to the description from sap_etiket_map if needed
                      urun_aciklamasi = row.get('ÇIKIŞ ÜRÜN ACIKLAMA', row.get('GİRİŞ ÜRÜN ACIKLAMA', ''))
                 
                 # Format the cycle path to include process for the current raw material step
                 # This is the final step in the path, so we add it to the path
                 current_step_formatted = f"{row.get('PROSES', '?')} ({current_barcode})"
                 formatted_path = current_path + [current_step_formatted] # Append to existing path
                 
                 results.append({
                     "Barkod": current_barcode,
                     "Ürün Açıklaması": urun_aciklamasi,
                     "Makine": row.get('MAKİNE NO', ''),
                     "Tüketim": row.get('GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg', 0) if pd.notna(row.get('GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg')) else 0,
                     "Oluşturma Zamanı": row.get('OLUŞTURMA ZAMANI', ''),
                     "Proses": row.get('PROSES', ''),
                     "İşlem Döngüsü": " -> ".join(formatted_path)
                 })
                 # This is a starting point, no need to go further back.
                 continue

            # If there are input rows for this output barcode, process them
            if rows_for_output:
                # Get output product info from the first representative row (they should be the same for the same output)
                representative_row = rows_for_output[0]
                
                # For intermediate/finished products, prefer output description
                urun_aciklamasi = representative_row.get('ÇIKIŞ ÜRÜN ACIKLAMA', representative_row.get('GİRİŞ ÜRÜN ACIKLAMA', ''))
                
                # Format the cycle path to include process for the current step
                # This step is added to the path, and we will recurse for its inputs
                current_step_formatted = f"{representative_row.get('PROSES', '?')} ({current_barcode})"
                formatted_path = current_path + [current_step_formatted] # Append to existing path
                
                results.append({
                    "Barkod": current_barcode,
                    "Ürün Açıklaması": urun_aciklamasi,
                    "Makine": representative_row.get('MAKİNE NO', ''),
                    "Tüketim": representative_row.get('GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg', 0) if pd.notna(representative_row.get('GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg')) else 0,
                    "Oluşturma Zamanı": representative_row.get('OLUŞTURMA ZAMANI', ''),
                    "Proses": representative_row.get('PROSES', ''),
                    "İşlem Döngüsü": " -> ".join(formatted_path)
                })
                
                # Find all input barcodes for this step and recurse
                # Pass the updated path (including current step) to the next recursion level
                input_barcodes = [row['GİRİŞ ÜRÜN SAP BARKODU'] for row in rows_for_output if pd.notna(row['GİRİŞ ÜRÜN SAP BARKODU'])]
                find_path(input_barcodes, formatted_path, visited) # Pass formatted_path
            else:
                 # If no info is found for this barcode in main maps, check input_barkod_to_row_map for raw material info
                 urun_aciklamasi = "Bilinmiyor"
                 makine = "Bilinmiyor"
                 tuketim = 0
                 olusturma_zamani = ""
                 proses = "Bilinmiyor"
                 
                 if current_barcode in input_barkod_to_row_map:
                     row = input_barkod_to_row_map[current_barcode]
                     urun_aciklamasi = row.get('GİRİŞ ÜRÜN ACIKLAMA', 'Bilinmiyor')
                     makine = row.get('MAKİNE NO', 'Bilinmiyor')
                     tuketim = row.get('GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg', 0) if pd.notna(row.get('GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg')) else 0
                     olusturma_zamani = row.get('OLUŞTURMA ZAMANI', '')
                     proses = row.get('PROSES', 'Bilinmiyor')
                 
                 # If no info is found for this barcode at all, still add it to the list
                 # (can be useful for data inconsistency cases)
                 current_step_formatted = f"{proses} ({current_barcode})"
                 formatted_path = current_path + [current_step_formatted] # Append to existing path
                 results.append({
                     "Barkod": current_barcode,
                     "Ürün Açıklaması": urun_aciklamasi,
                     "Makine": makine,
                     "Tüketim": tuketim,
                     "Oluşturma Zamanı": olusturma_zamani,
                     "Proses": proses,
                     "İşlem Döngüsü": " -> ".join(formatted_path)
                 })

    # Start tracking for each final product barcode
    # Initial path is empty for the final product
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


# İşlem sırasını ters çevir (ilk satır en uzun döngüyü yani son ürünü göstersin)
df_sonuc_ters = df_sonuc.copy()
df_sonuc_ters = df_sonuc_ters.iloc[::-1] # Tüm satırları ters sırala
df_sonuc_ters.reset_index(drop=True, inplace=True) # İndeksleri sıfırla

# İşlem Döngüsünü de ters çevir (İlk adım en başta olsun)
def reverse_cycle(cycle_str):
    parts = cycle_str.split(' -> ')
    return ' -> '.join(reversed(parts))

df_sonuc_ters['İşlem Döngüsü'] = df_sonuc_ters['İşlem Döngüsü'].apply(reverse_cycle)

# Benzersiz bir dosya adı oluştur (tarih ve saat ekleyerek)
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
output_file = f"uretim_izleme_sonucu_{timestamp}.xlsx"

try:
    # Sonucu bir Excel dosyasına kaydet
    df_sonuc_ters.to_excel(output_file, index=False)
    print(f"\nSonuç '{output_file}' dosyasına kaydedildi.")
except PermissionError as e:
    print(f"Hata: Dosyaya yazma izniniz yok. Lütfen '{output_file}' dosyasının açık olmadığından ve yazma izniniz olduğundan emin olun.")
    print(f"Detaylı hata: {e}")

"""
# Ekstra: Sadece işlem döngüsünü gösteren kısmı yazdır
print("\n--- İşlem Döngüleri (Özet) ---")
for index, row in df_sonuc_ters.iterrows():
    print(row['İşlem Döngüsü'])
"""