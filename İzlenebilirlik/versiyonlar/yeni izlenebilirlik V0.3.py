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

    # Create mappings
    output_to_input_map = {}
    input_barkod_to_row_map = {}
    sap_etiket_to_row_map = {}
    
    for _, row in df.iterrows():
        # Map for output tracking
        output_barkod = row['TEYİT VERİLEN BARKOD']
        if output_barkod not in output_to_input_map:
            output_to_input_map[output_barkod] = []
        output_to_input_map[output_barkod].append(row)
        
        # Map for input description lookup
        input_barkod = row['GİRİŞ ÜRÜN SAP BARKODU']
        if input_barkod not in input_barkod_to_row_map:
            input_barkod_to_row_map[input_barkod] = row
            
        # Map for SAP etiket lookup
        sap_barkod = row['SAP ETİKET BARKODU']
        if sap_barkod not in sap_etiket_to_row_map:
            sap_etiket_to_row_map[sap_barkod] = row

    results = []

    def get_product_info(barcode, context_row=None):
        """Get product information for a given barcode"""
        # First check if this barcode appears as an output (TEYİT VERİLEN BARKOD)
        if barcode in output_to_input_map:
            # This is an intermediate/final product
            representative_row = output_to_input_map[barcode][0]
            return {
                'urun_aciklamasi': representative_row.get('ÇIKIŞ ÜRÜN ACIKLAMA', ''),
                'makine': representative_row.get('MAKİNE NO', ''),
                'tuketim': representative_row.get('GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg', 0),
                'olusturma_zamani': representative_row.get('OLUŞTURMA ZAMANI', ''),
                'proses': representative_row.get('PROSES', ''),
                'row': representative_row
            }
        
        # If not found as output, check as SAP ETİKET BARKODU (could be raw material or intermediate)
        elif barcode in sap_etiket_to_row_map:
            row = sap_etiket_to_row_map[barcode]
            # For raw materials, we want the input information from the row where this barcode is used
            return {
                'urun_aciklamasi': row.get('GİRİŞ ÜRÜN ACIKLAMA', ''),
                'makine': '',  # Raw materials don't have production machine
                'tuketim': 0,  # Raw materials don't have consumption, they are consumed
                'olusturma_zamani': '',  # Raw materials don't have creation time in this context
                'proses': row.get('PROSES', '').replace('PR', 'TF').replace('TV', 'TF') if 'TF' in str(row.get('GİRİŞ ÜRÜN ACIKLAMA', '')) else row.get('PROSES', ''),
                'row': row
            }
        
        # If not found in SAP etiket, check if it's used as input somewhere
        elif barcode in input_barkod_to_row_map:
            row = input_barkod_to_row_map[barcode]
            # Determine the correct process based on product description
            proses = 'TF'  # Default for raw materials
            if 'TF' in str(row.get('GİRİŞ ÜRÜN ACIKLAMA', '')):
                proses = 'TF'
            elif context_row:
                proses = context_row.get('PROSES', 'TF')
                
            return {
                'urun_aciklamasi': row.get('GİRİŞ ÜRÜN ACIKLAMA', ''),
                'makine': '',  # Raw materials don't have production machine
                'tuketim': 0,
                'olusturma_zamani': '',
                'proses': proses,
                'row': row
            }
        
        # If not found anywhere, return default values
        else:
            return {
                'urun_aciklamasi': 'Bilinmiyor',
                'makine': '',
                'tuketim': 0,
                'olusturma_zamani': '',
                'proses': 'Bilinmiyor',
                'row': None
            }

    def format_value(value, value_type='string'):
        """Format values properly"""
        if value_type == 'float':
            if pd.notna(value):
                return float(value)
            return 0
        elif value_type == 'datetime':
            if pd.notna(value):
                if isinstance(value, datetime):
                    return value.strftime('%Y-%m-%d %H:%M:%S')
                else:
                    return str(value)
            return ''
        else:
            if pd.notna(value):
                return str(value)
            return ''

    def find_path(current_barcodes, current_path, visited):
        """Recursively finds the path backwards."""
        for current_barcode in current_barcodes:
            if current_barcode in visited:
                continue
            visited.add(current_barcode)

            # Get product information
            product_info = get_product_info(current_barcode)
            
            # Check if this barcode has inputs (is a produced item)
            rows_for_output = output_to_input_map.get(current_barcode, [])
            
            if rows_for_output:
                # This is a produced item, get its production info
                representative_row = rows_for_output[0]
                
                results.append({
                    "Barkod": current_barcode,
                    "Ürün Açıklaması": format_value(product_info['urun_aciklamasi']),
                    "Makine": format_value(product_info['makine']),
                    "Tüketim": format_value(representative_row.get('GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg', 0), 'float'),
                    "Oluşturma Zamanı": format_value(representative_row.get('OLUŞTURMA ZAMANI', ''), 'datetime'),
                    "Proses": format_value(product_info['proses']),
                    "İşlem Döngüsü": " -> ".join(current_path + [f"{product_info['proses']} ({current_barcode})"])
                })
                
                # Find input barcodes and continue recursion
                input_barcodes = [row['GİRİŞ ÜRÜN SAP BARKODU'] for row in rows_for_output if pd.notna(row['GİRİŞ ÜRÜN SAP BARKODU'])]
                new_path = current_path + [f"{product_info['proses']} ({current_barcode})"]
                find_path(input_barcodes, new_path, visited)
            else:
                # This is likely a raw material
                results.append({
                    "Barkod": current_barcode,
                    "Ürün Açıklaması": format_value(product_info['urun_aciklamasi']),
                    "Makine": format_value(product_info['makine']),
                    "Tüketim": format_value(product_info['tuketim'], 'float'),
                    "Oluşturma Zamanı": format_value(product_info['olusturma_zamani'], 'datetime'),
                    "Proses": format_value(product_info['proses']),
                    "İşlem Döngüsü": " -> ".join(current_path + [f"{product_info['proses']} ({current_barcode})"])
                })

    # Start tracking for each final product barcode
    for barcode in final_product_barcodes:
        find_path([barcode], [], set())
        
    # Convert results to DataFrame and sort by the length of the process cycle
    result_df = pd.DataFrame(results)
    if not result_df.empty:
        result_df['Döngü_Uzunluğu'] = result_df['İşlem Döngüsü'].str.count(' -> ')
        result_df.sort_values(by='Döngü_Uzunluğu', ascending=False, inplace=True)
        result_df.drop('Döngü_Uzunluğu', axis=1, inplace=True)
        result_df.reset_index(drop=True, inplace=True)
        
    return result_df

# Dosya yolu ve izlemeye başlanacak son ürün barkodu
file_path = '79528600-33.xlsx'
final_barcodes = ['79528600-33']

# Fonksiyonu çağır ve sonucu al
df_sonuc = track_production_backwards(file_path, final_barcodes)

# İşlem sırasını ters çevir (ham maddeler en üstte)
df_sonuc_ters = df_sonuc.copy()
df_sonuc_ters = df_sonuc_ters.iloc[::-1]
df_sonuc_ters.reset_index(drop=True, inplace=True)

# İşlem Döngüsünü de ters çevir
def reverse_cycle(cycle_str):
    parts = cycle_str.split(' -> ')
    return ' -> '.join(reversed(parts))

df_sonuc_ters['İşlem Döngüsü'] = df_sonuc_ters['İşlem Döngüsü'].apply(reverse_cycle)

# Benzersiz bir dosya adı oluştur
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
output_file = f"uretim_izleme_sonucu_{timestamp}.xlsx"

try:
    # Excel writer ile sütun genişliklerini ayarla
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_sonuc_ters.to_excel(writer, index=False, sheet_name='Sheet1')
        
        # Sütun genişliklerini ayarla
        worksheet = writer.sheets['Sheet1']
        column_widths = {
            'A': 15,  # Barkod
            'B': 35,  # Ürün Açıklaması
            'C': 10,  # Makine
            'D': 15,  # Tüketim
            'E': 20,  # Oluşturma Zamanı
            'F': 10,  # Proses
            'G': 80   # İşlem Döngüsü
        }
        
        for column, width in column_widths.items():
            worksheet.column_dimensions[column].width = width
    
    print(f"\nSonuç '{output_file}' dosyasına kaydedildi.")
    
   
    
except PermissionError as e:
    print(f"Hata: Dosyaya yazma izniniz yok. Lütfen '{output_file}' dosyasının açık olmadığından emin olun.")
    print(f"Detaylı hata: {e}")
