# hesapla_V0.9.py

import re
from bs4 import BeautifulSoup, MarkupResemblesLocatorWarning
import pandas as pd
import os
import warnings
from collections import defaultdict

# BeautifulSoup uyarısını bastir
warnings.filterwarnings("ignore", category=MarkupResemblesLocatorWarning)

def parse_html_for_tonnage(html_content_or_path):
    """
    HTML icerigini (string veya dosya yolu olarak) parse eder,
    her makine icin toplam KG'yi tonaj olarak hesaplar.
    """
    content = ""
    # Girdi bir dosya yolu mu?
    if isinstance(html_content_or_path, str) and html_content_or_path == "program.html":
        # Script'in bulunduğu dizini al
        script_dir = os.path.dirname(os.path.realpath(__file__))
        full_path = os.path.join(script_dir, html_content_or_path)
        print(f"Aranan dosya yolu: {full_path}")
        if not os.path.exists(full_path):
             print(f"Hata: '{full_path}' dosyasi bulunamadi.")
             return {}
        html_content_or_path = full_path # Gercek yolu kullan

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
    soup = BeautifulSoup(content, 'html.parser')

    # Verileri saklamak icin yapi
    # top degerine gore gruplanmis veriler
    rows_by_top = defaultdict(dict) # {top: {type: {data}}}
    # type: 'order' (siparis), 'product' (urun/makine), 'kg' (kilo)

    # 1. Tum span elementlerini bul
    all_spans = soup.find_all('span')
    print(f"Toplam span elementi sayisi: {len(all_spans)}")
    
    if len(all_spans) == 0:
        print("Uyari: Span bulunamadi. Dosya icerigi uygun gorunmuyor.")
        return {}

    # 2. Span'leri isle ve top degerine gore grupla
    for span in all_spans:
        style = span.get('style', '')
        text = span.get_text(strip=True)
        span_id = span.get('id', '')
        
        if not style: # Style olmazsa konum belirlenemez
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

        # 1. Urun/Makine bilgisi (v0.8'den gelen mantik, ama araliklari daraltiyoruz)
        # Oncelik id attribute'unde 'mainmenu' ve '...' aramak
        # Yedek yontem: left degeri 310-330px (v0.7'deki gibi) ve metinde makine no varsa
        if ("mainmenu" in span_id and "'" in span_id) or (310 <= left_val <= 330):
             # id'den makine numarasini cikar
            machine_from_id = None
            if span_id:
                id_machine_match = re.search(r"'(\d+)'", span_id)
                if id_machine_match:
                    machine_from_id = id_machine_match.group(1)
            
            # Metinden makine numarasini cikar (yedek yontem)
            machine_from_text = None
            if text:
                # Parantez ici rakamlari ara (orn: HT 1.63MM-CT-KY180 (179) YAGLI)
                text_machine_match = re.search(r'$$(\d+)$$', text)
                if not text_machine_match:
                    # Eger parantez icinde yoksa, metin sonundaki rakamlari ara (orn: HT 1.63MM-CT-KY180 179)
                    # Bu daha riskli, ama bazen kullanilabilir.
                    # text_machine_match = re.search(r'(\d+)$', text.strip())
                    pass # simdilik yedek yontemi kapat
                if text_machine_match:
                    machine_from_text = text_machine_match.group(1)
            
            # Oncelik id'dekine, yoksa metindekini kullan
            final_machine_no = machine_from_id or machine_from_text
            
            if final_machine_no:
                rows_by_top[top_val]['product'] = {
                    'machine_no': final_machine_no,
                    'text': text,
                    'id': span_id
                }
                # print(f"Top {top_val}: Makine bulundu: M{final_machine_no} (ID: {span_id}, Text: {text[:30]}...)")

        # 2. KG bilgisi (v0.8'den gelen mantik, ama araliklari daraltiyoruz)
        # left degeri 1200-1250px ve metin rakam iceriyor
        elif 1200 <= left_val <= 1250 and text and re.search(r'\d', text):
             # KG degerini cikar
             # Metinde '+' gibi karakterler olabilir, temizleyelim
             # Orn: "143" gibi veya "143+" gibi
             # En sagdaki, ardışık rakamları alalım (v0.8 ile ayni)
            matches = re.findall(r'\d+', text)
            if matches:
                # Genellikle en sondaki rakam grubu KG'dir
                clean_text = matches[-1]
                try:
                    kg_value = float(clean_text)
                    # 0 KG degerlerini atlayabiliriz, cok kucuk degerler de gurultu olabilir
                    # Ama cok dusuk esik koymak bazen gercek degerleri de kaybettirebilir.
                    # Simdi tum pozitif degerleri alalim.
                    if kg_value >= 0: # Sadece negatif olmayanlari al
                         rows_by_top[top_val]['kg'] = {
                            'value': kg_value,
                            'text': text,
                            'original_left': left_val
                        }
                        # print(f"Top {top_val}: KG bulundu: {kg_value} (Text: {text}, Left: {left_val})")
                except ValueError:
                    pass # Gecersiz sayi

    print(f"Top degerine gore gruplanmis satir sayisi: {len(rows_by_top)}")

    # 3. Gruplanmis verilerden makine tonajlarini hesapla
    machine_tonnage = {}
    used_kg_entries = 0 # Eslesen KG sayisi
    
    for top_val, row_data in rows_by_top.items():
        # Bir satirda urun/makine bilgisi olmali
        product_info = row_data.get('product')
        kg_info = row_data.get('kg')
        
        if product_info and kg_info:
            machine_no = product_info['machine_no']
            kg_value = kg_info['value']
            
            machine_key = f"M{machine_no}"
            if machine_key not in machine_tonnage:
                machine_tonnage[machine_key] = 0
            machine_tonnage[machine_key] += kg_value
            used_kg_entries += 1
            # print(f"M{machine_no} icin {kg_value} kg eklendi. Toplam: {machine_tonnage[machine_key]}")
        # else:
        #     # Hata ayiklama: Hangi satirlar eslesmedi?
        #     if product_info or kg_info:
        #         print(f"Uyari: Top {top_val} satiri tam olarak islenemedi. Urun/Makine: {bool(product_info)}, KG: {bool(kg_info)}.")
        #         if product_info: print(f"  -> Urun: M{product_info['machine_no']}")
        #         if kg_info: print(f"  -> KG: {kg_info['value']}")

    print(f"Eslesen ve kullanilan satir (urun+kg) sayisi: {used_kg_entries}")
    print(f"Bulunan ve kullanilan makine sayisi: {len(machine_tonnage)}")

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
        # Excel dosyasinin yolunu da script dizinine gore ayarlayalim
        script_dir = os.path.dirname(os.path.realpath(__file__))
        full_output_path = os.path.join(script_dir, output_file_path)
        
        df.to_excel(full_output_path, index=False)
        print(f"Sonuc raporu basariyla '{full_output_path}' dosyasina yazildi.")
    except Exception as e:
        print(f"Excel dosyasi olusturulurken hata olustu: {e}")


# --- Ana Program ---
if __name__ == "__main__":
    # Degistirilebilir degiskenler
    INPUT_HTML_SOURCE = "program.html" # Girdi HTML dosyasi (sadece adi, script ile ayni dizinde olmali)
    OUTPUT_EXCEL_FILE = "tonaj_raporu.xlsx" # Cikti Excel dosyasi (sadece adi, script ile ayni dizinde olacak)

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
        print("1. 'program.html' dosyasi bu script ile ayni klasorde mi?")
        print("2. Dosya icerigi HTML mi ve beklenen formatta mi?")
        print("3. Dosya cok buyuk olabilir, islem suresi uzun olabilir.")
        print("4. HTML yapisinda beklenmedik degisiklikler olabilir.")
