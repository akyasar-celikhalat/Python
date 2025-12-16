import pdfplumber
import pandas as pd
import os

# PDF dosyasını açma ve veriyi okuma
def extract_data_from_pdf(pdf_path):
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"Dosya bulunamadı: {pdf_path}")

    data = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            try:
                tables = page.extract_tables()
                for table_num, table in enumerate(tables):
                    # print(f"Sayfa {page_num + 1}, Tablo {table_num + 1} işleniyor...")
                    data.extend(table)
            except Exception as e:
                print(f"Sayfa {page_num + 1} işlenirken hata oluştu: {e}")
    return data

# Tonaj hesaplama fonksiyonu
def calculate_tonnage(data):
    machine_tonnage = {}
    current_machine = None
    processed_machines = set() # İşlenen makineleri takip etmek için

    for i, row in enumerate(data):
        # Hata ayıklama: Satır içeriğini yazdır (geliştirme/test aşamasında yararlıdır)
        # print(f"Satır {i}: {row}") 

        # Makine adını kontrol et (daha esnek)
        if row and row[0] and isinstance(row[0], str) and ("Makine：" in row[0] or "Makine:" in row[0]): # Hem ： hem : kontrolü
            try:
                # Makine adını al (örnek: "Makine： M179" -> "M179")
                current_machine = row[0].split("Makine")[1].split()[0].strip("：: ") 
                if current_machine not in machine_tonnage:
                    machine_tonnage[current_machine] = 0
                    processed_machines.add(current_machine)
                # print(f"Yeni makine bulundu: {current_machine}") # Hata ayıklama
            except (IndexError, AttributeError):
                # print(f"Uyarı: Satır {i} - Makine adı alınamadı: {row[0]}") # Hata ayıklama
                current_machine = None # Makine adı alınamazsa, KG değerlerini eklemeyiz
                continue

        # KG değerini kontrol et ve topla (sadece geçerli bir makine varsa)
        if current_machine and row and len(row) > 5 and row[5] is not None:
            try:
                # KG değerini temizle ve sayıya çevir
                kg_str = str(row[5]).strip().replace(',', '.').replace(' ', '') # Virgül/boşluk temizliği
                if kg_str.lower() in ('', '-', '---', 'kg'): # Boş veya geçersiz değerleri atla
                     # print(f"Uyarı: Satır {i} - Geçersiz KG değeri (boş/geçersiz): '{row[5]}'") # Hata ayıklama
                     continue
                kg = float(kg_str)
                if kg < 0:
                    # print(f"Uyarı: Satır {i} - Negatif KG değeri atlandı: {kg}") # Hata ayıklama
                    pass # Negatif değerleri de toplamaya devam edebiliriz, veya atlayabiliriz
                machine_tonnage[current_machine] += kg
                # print(f"Satır {i}: {current_machine} için {kg} kg eklendi. Toplam: {machine_tonnage[current_machine]}") # Hata ayıklama
            except ValueError:
                # print(f"Uyarı: Satır {i} - KG değeri sayıya çevrilemedi: '{row[5]}'") # Hata ayıklama
                pass # Sayıya çevrilemeyen değerleri atla

    # Tonaja çevir (1 ton = 1000 kg)
    for machine in machine_tonnage:
        machine_tonnage[machine] = machine_tonnage[machine] / 1000

    print(f"İşlenen makine sayısı: {len(processed_machines)}")
    if not machine_tonnage or all(v == 0 for v in machine_tonnage.values()):
         print("Uyarı: Hesaplanan tonajlar sıfır veya boş.")
    return machine_tonnage


# Excel'e yazdırma fonksiyonu
def write_to_excel(tonnage, output_file):
    if not tonnage:
        print("Hata: Sonuçlar boş. Lütfen verileri kontrol edin.")
        return

    # Sıralı bir şekilde yazmak için
    sorted_tonnage = dict(sorted(tonnage.items()))
    
    df = pd.DataFrame(list(sorted_tonnage.items()), columns=["Makine", "Tonaj (ton)"])
    try:
        df.to_excel(output_file, index=False)
        print(f"Sonuçlar başarıyla '{output_file}' dosyasına kaydedildi.")
    except Exception as e:
        print(f"Excel dosyasına yazılırken hata oluştu: {e}")

# Ana işlem
if __name__ == "__main__":
    pdf_path = r"C:\Users\yak\Desktop\Python\Program kg\program.pdf"  # PDF dosyasının yolu
    output_file = r"C:\Users\yak\Desktop\Python\Program kg\sonuclar.xlsx"  # Çıktı Excel dosyasının yolu

    try:
        print("PDF verileri okunuyor...")
        data = extract_data_from_pdf(pdf_path)
        print(f"PDF'den {len(data)} satır veri çekildi.")
        
        print("Tonajlar hesaplanıyor...")
        tonnage = calculate_tonnage(data)
        
        print("Sonuçlar Excel dosyasına yazılıyor...")
        write_to_excel(tonnage, output_file)
        
    except FileNotFoundError as e:
        print(e)
    except Exception as e:
        print(f"Beklenmeyen bir hata oluştu: {e}")
