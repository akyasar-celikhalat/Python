import pdfplumber
import pandas as pd
import os
import re

# PDF dosyasını açma ve veriyi okuma
def extract_data_from_pdf(pdf_path):
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"Dosya bulunamadı: {pdf_path}")
    
    data = []
    with pdfplumber.open(pdf_path) as pdf:
        print(f"PDF toplam sayfa sayısı: {len(pdf.pages)}")
        for page_num, page in enumerate(pdf.pages, 1):
            print(f"Sayfa {page_num} işleniyor...")
            
            # Önce tabloları dene
            tables = page.extract_tables()
            if tables:
                print(f"Sayfa {page_num}'da {len(tables)} tablo bulundu")
                for table in tables:
                    data.extend(table)
            
            # Eğer tablo bulunamazsa, metni satır satır oku
            else:
                text = page.extract_text()
                if text:
                    lines = text.split('\n')
                    for line in lines:
                        if line.strip():
                            # Her satırı boşluklara göre böl
                            data.append(line.split())
    
    print(f"Toplam {len(data)} satır veri çıkarıldı")
    return data

# Tonaj hesaplama fonksiyonu
def calculate_tonnage(data):
    machine_tonnage = {}
    current_machine = None
    
    for i, row in enumerate(data):
        if not row:  # Boş satırları atla
            continue
            
        # Satırı string'e çevir kontrol için
        row_str = ' '.join(str(cell) if cell else '' for cell in row)
        
        # Makine adını kontrol et - farklı formatları dene
        if "Makine" in row_str or "makine" in row_str:
            # Makine adını çıkarmaya çalış
            machine_match = re.search(r'Makine[：:]\s*(\w+)', row_str)
            if machine_match:
                current_machine = machine_match.group(1)
                machine_tonnage[current_machine] = 0
                print(f"Makine bulundu: {current_machine}")
            else:
                # Alternatif format
                if len(row) > 1:
                    current_machine = str(row[1]).strip()
                    machine_tonnage[current_machine] = 0
                    print(f"Makine bulundu (alternatif): {current_machine}")
        
        # KG değerini kontrol et
        if current_machine:
            for cell in row:
                if cell is not None:
                    cell_str = str(cell).replace(',', '.').replace(' ', '')
                    
                    # Sayısal değer kontrolü - kg olabilecek değerler
                    if re.match(r'^\d+\.?\d*$', cell_str):
                        try:
                            kg_value = float(cell_str)
                            # Mantıklı kg değeri kontrolü (0.1 kg - 10000 kg arası)
                            if 0.1 <= kg_value <= 10000:
                                machine_tonnage[current_machine] += kg_value
                                print(f"{current_machine} için {kg_value} kg eklendi")
                        except ValueError:
                            continue
    
    # Tonaja çevir (1 ton = 1000 kg)
    for machine, total_kg in machine_tonnage.items():
        machine_tonnage[machine] = round(total_kg / 1000, 3)  # 3 ondalık basamak
    
    return machine_tonnage

# Excel'e yazdırma fonksiyonu
def write_to_excel(tonnage, output_file):
    if not tonnage:
        print("Hata: Sonuçlar boş. Lütfen verileri kontrol edin.")
        return
    
    # Sonuçları göster
    print("\n=== SONUÇLAR ===")
    for machine, tons in tonnage.items():
        print(f"{machine}: {tons} ton")
    
    df = pd.DataFrame(list(tonnage.items()), columns=["Makine", "Tonaj (ton)"])
    
    # Klasör yoksa oluştur
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    
    df.to_excel(output_file, index=False)
    print(f"\nSonuçlar başarıyla '{output_file}' dosyasına kaydedildi.")

# Debug fonksiyonu - veriyi kontrol etmek için
def debug_data(data):
    print("\n=== DEBUG: İlk 10 satır ===")
    for i, row in enumerate(data[:10]):
        print(f"Satır {i}: {row}")
    
    print(f"\n=== Toplam satır sayısı: {len(data)} ===")

# Ana işlem
if __name__ == "__main__":
    pdf_path = r"C:\Users\yak\Desktop\Python\Program kg\program.pdf"
    output_file = r"C:\Users\yak\Desktop\Python\Program kg\sonuclar.xlsx"
    
    try:
        print("PDF dosyası okunuyor...")
        data = extract_data_from_pdf(pdf_path)
        
        # Debug için veriyi göster
        debug_data(data)
        
        print("\nTonaj hesaplanıyor...")
        tonnage = calculate_tonnage(data)
        
        if tonnage:
            write_to_excel(tonnage, output_file)
        else:
            print("Hiçbir makine veya tonaj verisi bulunamadı. PDF formatını kontrol edin.")
            
    except FileNotFoundError as e:
        print(e)
    except Exception as e:
        print(f"Bir hata oluştu: {e}")
        import traceback
        traceback.print_exc()