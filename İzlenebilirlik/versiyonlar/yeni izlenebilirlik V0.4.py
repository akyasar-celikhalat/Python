import pandas as pd
from datetime import datetime
import re

def track_production_backwards(file_path, search_term):
    """
    Tracks production backwards from final products using the provided Excel data.
    Supports two search types:
    1. Exact barcode: "79528600-33" - tracks only this specific barcode
    2. Base barcode: "79528600" - tracks all barcodes starting with this base (e.g., 79528600-1, 79528600-2, ..., 79528600-33)

    Args:
        file_path (str): Path to the Excel file.
        search_term (str): Either an exact barcode or a base barcode.

    Returns:
        pd.DataFrame: A DataFrame containing the tracked production steps.
    """
    # Load data
    print("Veriler yükleniyor...")
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

    # Determine search type and find target barcodes
    target_barcodes = []
    if '-' in search_term and search_term.endswith(tuple('0123456789')):
        # Exact barcode search
        print(f"Tam barkod araması yapılıyor: {search_term}")
        if search_term in output_to_input_map:
            target_barcodes = [search_term]
        else:
            print(f"Uyarı: {search_term} barkodu bulunamadı.")
            return pd.DataFrame()
    else:
        # Base barcode search - find all barcodes starting with this base
        print(f"Temel barkod araması yapılıyor: {search_term}")
        pattern = f"^{re.escape(search_term)}-\\d+$"
        target_barcodes = [barcode for barcode in output_to_input_map.keys() if re.match(pattern, barcode)]
        if not target_barcodes:
            print(f"Uyarı: {search_term} ile başlayan barkod bulunamadı.")
            return pd.DataFrame()
        print(f"Bulunan barkodlar: {target_barcodes}")

    results = []

    def get_product_info(barcode, context_row=None):
        """Get product information for a given barcode"""
        # First check if this barcode appears as an output (TEYİT VERİLEN BARKOD)
        if barcode in output_to_input_map:
            # This is an intermediate/final product
            representative_row = output_to_input_map[barcode][0]
            return {
                'urun_aciklamasi': representative_row.get('ÇIKIŞ ÜRÜN ACIKLAMA', representative_row.get('GİRİŞ ÜRÜN ACIKLAMA', '')),
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
                'urun_aciklamasi': row.get('GİRİŞ ÜRÜN ACIKLAMA', 'Bilinmiyor'),
                'makine': row.get('MAKİNE NO', ''),
                'tuketim': row.get('GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg', 0) if pd.notna(row.get('GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg')) else 0,
                'olusturma_zamani': row.get('OLUŞTURMA ZAMANI', ''),
                'proses': row.get('PROSES', ''),
                'row': row
            }
        
        # If not found in SAP etiket, check if it's used as input somewhere
        elif barcode in input_barkod_to_row_map:
            row = input_barkod_to_row_map[barcode]
            # Determine the correct process based on product description
            proses = row.get('PROSES', '')
            if not proses or str(proses).lower() == 'nan':
                # Try to extract from description
                desc = str(row.get('GİRİŞ ÜRÜN ACIKLAMA', ''))
                if desc and desc.lower() != 'nan':
                    first_word = desc.split()[0] if desc.split() else ''
                    proses = re.split(r'[^A-Za-z0-9]', first_word)[0] if first_word else 'Bilinmiyor'
                else:
                    proses = 'Bilinmiyor'
                
            return {
                'urun_aciklamasi': row.get('GİRİŞ ÜRÜN ACIKLAMA', 'Bilinmiyor'),
                'makine': row.get('MAKİNE NO', ''),
                'tuketim': row.get('GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg', 0) if pd.notna(row.get('GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg')) else 0,
                'olusturma_zamani': row.get('OLUŞTURMA ZAMANI', ''),
                'proses': proses,
                'row': row
            }
        
        # If not found anywhere, return default values
        else:
            return {
                'urun_aciklamasi': 'Bilinmiyor',
                'makine': 'Bilinmiyor',
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
            if pd.notna(value) and str(value).lower() != 'nan':
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

    # Start tracking for each target barcode
    print(f"İzleme başlatılıyor, {len(target_barcodes)} barkod işlenecek...")
    for i, barcode in enumerate(target_barcodes, 1):
        print(f"  {i}/{len(target_barcodes)}: {barcode} işleniyor...")
        find_path([barcode], [], set())
        
    # Convert results to DataFrame and sort by the length of the process cycle
    result_df = pd.DataFrame(results)
    if not result_df.empty:
        result_df['Döngü_Uzunluğu'] = result_df['İşlem Döngüsü'].str.count(' -> ')
        result_df.sort_values(by=['Döngü_Uzunluğu', 'Barkod'], ascending=[False, True], inplace=True)
        result_df.drop('Döngü_Uzunluğu', axis=1, inplace=True)
        result_df.reset_index(drop=True, inplace=True)
        
    return result_df

def save_results_to_excel(df_sonuc, search_term):
    """Save results to Excel with proper formatting"""
    if df_sonuc.empty:
        print("Kaydedilecek veri yok.")
        return
        
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
    safe_search_term = re.sub(r'[^\w\-_\. ]', '_', search_term)
    output_file = f"uretim_izleme_sonucu_{safe_search_term}_{timestamp}.xlsx"

    try:
        # Excel writer ile sütun genişliklerini ayarla
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_sonuc_ters.to_excel(writer, index=False, sheet_name='İzleme Sonuçları')
            
            # Sütun genişliklerini ayarla
            worksheet = writer.sheets['İzleme Sonuçları']
            column_widths = {
                'A': 18,  # Barkod
                'B': 40,  # Ürün Açıklaması
                'C': 12,  # Makine
                'D': 15,  # Tüketim
                'E': 20,  # Oluşturma Zamanı
                'F': 12,  # Proses
                'G': 100  # İşlem Döngüsü
            }
            
            for column, width in column_widths.items():
                worksheet.column_dimensions[column].width = width
        
        print(f"\nSonuç '{output_file}' dosyasına kaydedildi.")
        print(f"Toplam {len(df_sonuc_ters)} satır kaydedildi.")
        
    except PermissionError as e:
        print(f"Hata: Dosyaya yazma izniniz yok. Lütfen '{output_file}' dosyasının açık olmadığından emin olun.")
        print(f"Detaylı hata: {e}")
    except Exception as e:
        print(f"Beklenmeyen hata oluştu: {e}")

# --- Ana Program ---
if __name__ == "__main__":
    # Dosya yolu
    file_path = '2025.xlsx'  # Gerçek dosya yoluyla değiştirin
    
    # Kullanıcıdan giriş al
    print("=== Üretim İzlenebilirlik Sistemi ===")
    print("İzleme yapmak için aşağıdaki formatlardan birini girin:")
    print("1. Tam barkod: 79528600-33")
    print("2. Temel barkod: 79528600 (bu önekli tüm barkodlar izlenir)")
    print("Çıkmak için 'q' girin")
    
    while True:
        search_input = input("\nArama terimini girin: ").strip()
        
        if search_input.lower() == 'q':
            print("Program sonlandırılıyor.")
            break
            
        if not search_input:
            print("Lütfen geçerli bir arama terimi girin.")
            continue
            
        try:
            # Fonksiyonu çağır ve sonucu al
            df_sonuc = track_production_backwards(file_path, search_input)
            
            if df_sonuc.empty:
                print("Sonuç bulunamadı.")
                continue
                
            # İlk birkaç satırı göster
            print(f"\nİlk 10 sonuç:")
            print(df_sonuc.head(10).to_string(index=False))
            
            # Kullanıcıya dosyaya kaydetmek isteyip sorma
            save_choice = input(f"\n{len(df_sonuc)} sonuç bulundu. Excel dosyasına kaydetmek ister misiniz? (e/h): ").strip().lower()
            if save_choice in ['e', 'evet', 'y', 'yes']:
                save_results_to_excel(df_sonuc, search_input)
                
        except FileNotFoundError:
            print(f"Hata: '{file_path}' dosyası bulunamadı. Lütfen dosya yolunu kontrol edin.")
        except Exception as e:
            print(f"Beklenmeyen hata oluştu: {e}")