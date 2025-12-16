# html_tonaj_hesapla_v2.py

import re
from bs4 import BeautifulSoup, MarkupResemblesLocatorWarning
import pandas as pd
import os
import warnings

# BeautifulSoup uyarısını bastir
warnings.filterwarnings("ignore", category=MarkupResemblesLocatorWarning)

def parse_html_for_tonnage(html_content_or_path):
    """
    HTML icerigini (string veya dosya yolu olarak) parse eder,
    her makine icin toplam KG'yi tonaj olarak hesaplar.
    """
    content = ""
    # Girdi bir dosya yolu mu?
    if isinstance(html_content_or_path, str) and os.path.exists(html_content_or_path):
        try:
            print(f"Dosya okunuyor: {html_content_or_path}")
            with open(html_content_or_path, 'r', encoding='utf-8') as f:
                content = f.read()
            print(f"Dosya okundu. Karakter sayisi: {len(content)}")
            if not content.strip():
                 print("Hata: Dosya icerigi bos.")
                 return {}
        except Exception as e:
            print(f"'{html_content_or_path}' dosyasi okunurken hata olustu: {e}")
            return {}
    elif isinstance(html_content_or_path, str):
        # Girdi bir dosya yolu degilse, dogrudan HTML icerigi olarak kabul et
        print("Girdi bir dosya yolu gibi gorunmuyor, dogrudan HTML icerigi olarak isleniyor.")
        content = html_content_or_path
    else:
        print("Hata: Gecersiz girdi turu. Bir dosya yolu (string) veya HTML icerigi (string) bekleniyordu.")
        return {}

    if not content:
        print("Hata: Isleme alinacak icerik bulunamadi.")
        return {}

    # Daha toleransli parser kullan
    print("HTML parse ediliyor...")
    soup = BeautifulSoup(content, 'html.parser') # 'lxml' yerine 'html.parser'

    # Verileri saklamak icin yapi
    machine_tonnage = {}
    order_info = {} # { (top_degeri, siparis_id) : {text: "...", machine_no: "..."} }
    kg_values = [] # [ {top: ..., value: ...} ]

    # 1. Tum span elementlerini bul
    all_spans = soup.find_all('span')
    print(f"Toplam span elementi sayisi: {len(all_spans)}")
    
    if len(all_spans) == 0:
        # Hata ayiklama: Icerigin bir kismini yazdir
        print("Uyari: Span bulunamadi. Icerigin ilk 500 karakteri:")
        print(content[:500])
        print("...")
        print("Uyari: Icerigin son 500 karakteri:")
        print("...")
        print(content[-500:])

    # 2. Span'leri isle
    for i, span in enumerate(all_spans):
        style = span.get('style', '')
        text = span.get_text(strip=True)
        
        # Hata ayiklama: Ilk birkac span'i yazdir
        # if i < 5:
        #     print(f"Span {i}: Style='{style[:100]}...', Text='{text[:50]}...'")

        if not style or not text:
            continue

        # top degerini al
        top_match = re.search(r'top:\s*(\d+)px', style, re.IGNORECASE)
        if not top_match:
            continue
        top_val = int(top_match.group(1))

        # left degerine gore islem yap
        left_match = re.search(r'left:\s*(\d+)px', style, re.IGNORECASE)
        if not left_match:
            continue
        left_val = int(left_match.group(1))

        # 1. Siparis bilgisi (genellikle left:50px)
        if 40 <= left_val <= 60:
            # Siparis numarasini cikar
            order_match = re.search(r'(\d+-\d+)', text)
            if order_match:
                order_id = order_match.group(1)
                # Urun adinda makine numarasini ara (orn: ... (179) ...)
                # Daha genis bir arama: parantez icinde rakamlar
                machine_match = re.search(r'$$(\d+)$$', text)
                machine_number = machine_match.group(1) if machine_match else None
                
                if machine_number: # Sadece makine numarasi varsa kaydet
                    key = (top_val, order_id)
                    order_info[key] = {
                        'text': text,
                        'machine_no': machine_number
                    }
                    
                    # Makine sozlugunu baslat
                    if f"M{machine_number}" not in machine_tonnage:
                        machine_tonnage[f"M{machine_number}"] = 0
                        # print(f"Makine tanimlandi: M{machine_number}")
                    
        # 2. KG bilgisi (genellikle left:1220px)
        elif 1200 <= left_val <= 1240:
             # KG degerini cikar (sadece rakamlar)
            # Metinde '+' gibi karakterler olabilir, temizleyelim
            clean_text = re.sub(r'[^\d]', '', text)
            if clean_text.isdigit():
                try:
                    kg_value = float(clean_text)
                    kg_values.append({'top': top_val, 'value': kg_value})
                    # print(f"KG degeri bulundu: {kg_value} kg (top: {top_val})")
                except ValueError:
                    pass # Gecersiz sayi

    print(f"Bulunan siparis bilgisi girisi: {len(order_info)}")
    print(f"Bulunan KG degeri girisi: {len(kg_values)}")

    if len(order_info) == 0:
        print("Uyari: Siparis bilgisi bulunamadi. HTML yapisini kontrol edin.")
    if len(kg_values) == 0:
        print("Uyari: KG degeri bulunamadi. HTML yapisini kontrol edin.")

    # 3. KG degerlerini ilgili makinelere ata
    # Basit bir yontem: KG'nin top degeri, bir siparisin top degerine esit veya buyukse,
    # ve aradaki fark minimumsa, o KG bu siparise aittir.
    sorted_orders = sorted(order_info.keys())
    
    matched_kg_count = 0
    for kg_item in kg_values:
        kg_top = kg_item['top']
        kg_val = kg_item['value']
        
        best_match_key = None
        min_diff = float('inf')
        
        # Siparisleri yineleyerek en yakin olanı bul
        for order_key in sorted_orders:
            order_top = order_key[0] # (top, order_id)
            # KG genellikle siparisle ayni satirda veya hemen sonrasinda olur
            if order_top <= kg_top:
                diff = kg_top - order_top
                if diff < min_diff:
                    min_diff = diff
                    best_match_key = order_key
        
        if best_match_key:
            order_data = order_info[best_match_key]
            machine_no = order_data.get('machine_no')
            if machine_no:
                machine_key = f"M{machine_no}"
                machine_tonnage[machine_key] += kg_val
                matched_kg_count += 1
                # print(f"M{machine_no} icin {kg_val} kg eklendi. Toplam: {machine_tonnage[machine_key]}")

    print(f"Eslesen ve kullanilan KG degeri sayisi: {matched_kg_count}")

    # 4. KG'yi tona cevir
    for machine in machine_tonnage:
        machine_tonnage[machine] = machine_tonnage[machine] / 1000.0
        
    return machine_tonnage

def write_tonnage_to_excel(tonnage_dict, output_file_path):
    """
    Hesaplanan tonaj sozlugunu bir Excel dosyasina yazar.
    """
    if not tonnage_dict:
        print("Hata: Hesaplanan tonaj verisi bos. Excel dosyasi olusturulmadi.")
        return

    # Veriyi siralayalim
    sorted_data = dict(sorted(tonnage_dict.items()))
    
    # DataFrame olustur
    df = pd.DataFrame(list(sorted_data.items()), columns=["Makine", "Tonaj (ton)"])
    
    try:
        df.to_excel(output_file_path, index=False)
        print(f"Sonuc raporu basariyla '{output_file_path}' dosyasina yazildi.")
    except Exception as e:
        print(f"Excel dosyasi olusturulurken hata olustu: {e}")


# --- Ana Program ---
if __name__ == "__main__":
    # Degistirilebilir degiskenler
    # INPUT_HTML_SOURCE degiskenine asagidaki yontemlerden birini kullanabilirsiniz:
    # 1. Dosya yolu (raw string olarak)
    INPUT_HTML_SOURCE = r"C:\Users\yak\Desktop\Python\Program kg\Pasted_Text_1753709752258.txt" # Girdi HTML dosyasi
    
    # 2. Veya dogrudan HTML icerigini bir degiskene atayip o degiskeni kullanabilirsiniz
    # INPUT_HTML_SOURCE = """...HTML icerigi...""" # Buraya HTML icerigini yapistirin
    
    OUTPUT_EXCEL_FILE = r"C:\Users\yak\Desktop\Python\Program kg\tonaj_raporu.xlsx" # Cikti Excel dosyasi

    print("HTML kaynagi isleniyor...")
    tonnage_results = parse_html_for_tonnage(INPUT_HTML_SOURCE)

    if tonnage_results:
        print("\nHesaplanan Tonajlar:")
        for machine, tonnage in tonnage_results.items():
            print(f"  {machine}: {tonnage:.2f} ton")
        
        print("\nExcel dosyasi olusturuluyor...")
        write_tonnage_to_excel(tonnage_results, OUTPUT_EXCEL_FILE)
    else:
        print("\nHata: Islem sonucunda hicbir veri elde edilemedi.")
        print("Lutfen asagidaki noktalari kontrol edin:")
        print("1. INPUT_HTML_SOURCE degiskenindeki dosya yolu dogru mu?")
        print("2. Dosya gercekten mevcut mu?")
        print("3. Dosya icerigi HTML mi ve beklenen formatta mi?")
        print("4. Dosya cok buyuk olabilir, islem suresi uzun olabilir.")
