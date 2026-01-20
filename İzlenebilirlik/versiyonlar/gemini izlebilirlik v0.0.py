import pandas as pd

def create_traceability_report(input_file, output_file):
    """
    Bir Excel/CSV dosyasından üretim verilerini okur, izlenebilirlik zinciri oluşturur
    ve sonuçları yeni bir Excel dosyasına yazar.

    Args:
        input_file (str): Giriş verilerini içeren dosyanın adı.
        output_file (str): Oluşturulacak raporun kaydedileceği Excel dosyasının adı.
    """
    try:
        # 1. Dosyayı oku
        # HATA DÜZELTMESİ: Türkçe karakterler içeren CSV dosyalarını
        # doğru okumak için 'encoding="latin1"' parametresi eklendi.
        # Eğer bu da çalışmazsa, 'iso-8859-9' veya 'cp1254' deneyebilirsiniz.
        df = pd.read_csv(input_file, encoding="latin1")
        print(f"'{input_file}' dosyası başarıyla okundu.")

    except FileNotFoundError:
        print(f"HATA: '{input_file}' adlı dosya bulunamadı. Lütfen dosya adını kontrol edin.")
        return
    except Exception as e:
        print(f"Dosya okunurken bir hata oluştu: {e}")
        return

    # 2. İzlenebilirlik ve barkod detayları için veri yapılarını oluştur
    parent_map = {}
    barcode_details = {}

    # 3. Veri setini satır satır işle
    for _, row in df.iterrows():
        # Kolon adlarının dosyanızdaki ile tam eşleştiğinden emin olun
        child_barcode = row['GİRİŞ ÜRÜN SAP BARKODU']
        parent_barcode = row['TEYİT VERİLEN BARKOD']

        parent_map[child_barcode] = parent_barcode

        if child_barcode not in barcode_details:
            barcode_details[child_barcode] = {
                'Ürün Açıklaması': row['GİRİŞ ÜRÜN ACIKLAMA'],
                'Proses': None,
                'Makine': None,
                'Oluşturma Zamanı': None,
                'Tüketim': row['GİRİŞ ÜRÜN STOKU']
            }

        barcode_details[parent_barcode] = {
            'Ürün Açıklaması': row['ÇIKIŞ ÜRÜN ACIKLAMA'],
            'Proses': row['PROSES'],
            'Makine': row['MAKİNE NO'],
            'Oluşturma Zamanı': row['OLUŞTURMA ZAMANI'],
            'Tüketim': df[df['TEYİT VERİLEN BARKOD'] == parent_barcode]['GİRİŞ ÜRÜN TÜKETİM MİKTARI Kg'].sum()
        }
        
    all_barcodes = set(parent_map.keys()) | set(parent_map.values())
    for barcode in all_barcodes:
        if barcode not in parent_map.values():
             if barcode in barcode_details and barcode_details[barcode]['Proses'] is None:
                 barcode_details[barcode]['Proses'] = 'TF'
                 barcode_details[barcode]['Tüketim'] = 0

    # 4. İşlem döngüsünü oluşturan fonksiyon
    def get_trace_chain(barcode):
        chain = []
        current_barcode = barcode
        while current_barcode in barcode_details:
            details = barcode_details.get(current_barcode, {})
            proses = details.get('Proses', 'Bilinmiyor')
            chain.append(f"{proses} ({current_barcode})")
            if current_barcode in parent_map:
                current_barcode = parent_map[current_barcode]
            else:
                break
        return " -> ".join(chain)

    # 5. Çıktı verisini hazırla
    output_data = []
    for barcode in sorted(barcode_details.keys()):
        details = barcode_details[barcode]
        output_data.append({
            'Barkod': barcode,
            'Ürün Açıklaması': details['Ürün Açıklaması'],
            'Makine': details['Makine'],
            'Tüketim': str(details.get('Tüketim', 0)).replace('.', ','),
            'Oluşturma Zamanı': details['Oluşturma Zamanı'],
            'Proses': details['Proses'],
            'İşlem Döngüsü': get_trace_chain(barcode)
        })

    # 6. Sonuçları bir DataFrame'e dönüştür ve Excel'e yaz
    output_df = pd.DataFrame(output_data)
    desired_order = [
        'Barkod', 'Ürün Açıklaması', 'Makine', 'Tüketim', 
        'Oluşturma Zamanı', 'Proses', 'İşlem Döngüsü'
    ]
    output_df = output_df[desired_order]
    output_df.to_excel(output_file, index=False, engine='openpyxl')
    print(f"Rapor başarıyla oluşturuldu ve '{output_file}' dosyasına kaydedildi.")


# --- KODU ÇALIŞTIR ---
if __name__ == "__main__":
    # Girdi dosyasının adını tam olarak yüklediğiniz gibi güncelledim.
    input_filename = "79528600-33.xlsx"
    
    output_filename = "izlenebilirlik_raporu.xlsx"
    
    create_traceability_report(input_filename, output_filename)