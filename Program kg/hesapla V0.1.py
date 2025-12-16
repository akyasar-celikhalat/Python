import pdfplumber
import os

# PDF dosyasını açma ve veriyi okuma
def extract_data_from_pdf(pdf_path):
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"Dosya bulunamadı: {pdf_path}")
    
    data = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                data.extend(table)
    return data

# Tonaj hesaplama fonksiyonu
def calculate_tonnage(data):
    machine_tonnage = {}
    current_machine = None

    for row in data:
        # Makine adını kontrol et
        if row and isinstance(row[0], str) and "Makine：" in row[0]:
            current_machine = row[0].split("Makine：")[1].strip().split()[0]
            machine_tonnage[current_machine] = 0  # Yeni makine için tonajı sıfırla
        
        # KG değerini kontrol et ve topla
        if row and len(row) > 5 and row[5] is not None:  # KG sütununu kontrol et
            kg_value = row[5].replace(',', '.').replace(' ', '')  # Virgül veya boşlukları düzenle
            if kg_value.replace('.', '', 1).isdigit():  # Sayısal mı kontrol et
                kg = float(kg_value)
                machine_tonnage[current_machine] += kg

    # Tonaja çevir (1 ton = 1000 kg)
    for machine, total_kg in machine_tonnage.items():
        machine_tonnage[machine] = total_kg / 1000  # KG'yi tona çevir

    return machine_tonnage

# Ana işlem
if __name__ == "__main__":
    pdf_path = r"C:\Users\yak\Desktop\Python\Program kg\program.pdf"  # Tam yolu buraya yazın
    
    try:
        data = extract_data_from_pdf(pdf_path)
        tonnage = calculate_tonnage(data)
        
        # Sonuçları yazdır
        print("Makine Başına Planlanan Tonaj:")
        for machine, ton in tonnage.items():
            print(f"{machine}: {ton:.2f} ton")
    except FileNotFoundError as e:
        print(e)