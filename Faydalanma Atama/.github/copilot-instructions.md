# Copilot yönergeleri — Faydalanma Atama projesi

Aşağıdaki kısa, eyleme dönük rehber, bu repoda bir AI kod asistanının hızlıca üretken olmasını sağlar.

- **Proje özeti:** Küçük bir Python betiği (`faydalanma_atama.py`) Excel dosyalarını okuyup eşleme, tüketim ve stok hesapları yapar ve sonuçları `faydalanma_atama_sonuc_{timestamp}.xlsx` olarak kaydeder.

- **Çalıştırma (yerel, Windows):** `python faydalanma_atama.py`
- **Gerekli paketler:** `pandas`, `openpyxl` (yükleme örneği: `pip install pandas openpyxl`).

- **Giriş verisi konumu:** Script `BASE_DIR` (dosyanın olduğu klasör) ve bir üst klasörde `*.xls*` dosyalarını tarar; dosya eşlemesi dosya adı anahtar sözcüklerine göre yapılır (`KEYS` sabitindeki anahtarlar: `faydalanma`, `kilavuz`, `tuketim`). Değişiklik yaparken buna dikkat edin.

- **Önemli dosya:** [faydalanma_atama.py](faydalanma_atama.py) — en çok okunması gereken dosya; giriş/çıkış, sütun eşleme ve dönüşümler burada tanımlıdır.

## Kod ve mimari notları (hızlı referans)

- `find_files(search_dirs=None)`: Excel dosyalarını `KEYS` kullanarak tanımlar.
- `safe_read(path, sheet_name=None)`: `pandas.read_excel` sarımı; sheet name karşılaştırmaları casefold ile yapılır.
- `find_col(df, candidates)`: Sütun bulma; önce tam eşleşme, sonra casefold, sonra substring araması.
- `to_num`, `parse_diameter_mm`, `linear_kg_per_m_from_d_mm`: Veri temizleme ve fiziksel hesap yardımcıları — örnek dönüşümler için bu fonksiyonları kullanın.

## Proje-spesifik kalıplar ve kurallar

- Veri sütun isimleri çoğunlukla Türkçe ve büyük harfli varyantlarda gelir (ör. `BARKOD NUMARASI`, `İŞ EMRİ`). Yeni sütun aramaları eklerken `CANDIDATE_*` listelerine benzer aday isimleri ekleyin.
- Sütun eşlemesi toleranslıdır: küçük-büyük harf farkı ve substring eşleşmesi kullanılıyor; bu davranışı değiştirmek isterseniz `find_col` fonksiyonunu güncelleyin.
- Hata davranışı: Dosya okuma hataları `print` ile bildirilir ve işlem atlanır. Büyük değişikliklerde exception davranışını sabitleyin veya log ekleyin.

## Değişiklik yapılacak yaygın yerler

- Yeni giriş sütunu/isimleri: üstteki sabitlere (`KEYS`, `CANDIDATE_*`, `stok_desc_candidates`, vs.) eleman ekleyin.
- Yeni hesaplama veya farklı ağırlık/densite kullanımı: `linear_kg_per_m_from_d_mm` ve `parse_diameter_mm` girişleriyle entegre edin.
- Çıktı formatı değişikliği: script sonunda `pd.ExcelWriter` ile kaydediliyor — yeni sheet/summary eklemek için burayı düzenleyin.

## Örnek değişiklik senaryosu

- Senaryo: `TÜKETİM` sütunu farklı bir isimle geliyor.
  - Eylem: `C_TUKETIM_AMT` listesine yeni aday isim ekleyin ve `find_col` ile scripti yeniden çalıştırın.

## Test / doğrulama

- Bu repoda otomatik test yok. Değişiklik sonrası çalıştırıp üretilen `faydalanma_atama_sonuc_*.xlsx` dosyasını manuel kontrol edin.
- Hızlı kontrol: küçük örnek Excel dosyasıyle `python faydalanma_atama.py` çalıştırın ve `Eksik_IsEmri` / `Eksik_Barkod` sheet'lerini gözden geçirin.

## Merge talimatı (varsa mevcut `.github/copilot-instructions.md` ile)

- Eğer zaten bir `.github/copilot-instructions.md` varsa, burada verilen proje-özgü notları koruyun: özellikle `KEYS` ve `CANDIDATE_*` listelerine dair açıklamalar önemlidir. Yeni içerik eklerken mevcut tavsiyeleri silmeyin; sadece güncelleyin.

---
Eğer bu yönergede eksik veya belirsiz bir kısım varsa, hangi konuda daha fazla örnek veya ayrıntı istediğinizi belirtin; örn. örnek Excel satırı, beklenen sütun listesi veya özel hata senaryoları.

## Hızlı Başlangıç (özet)

- Çalıştırmak için: `python faydalanma_atama.py`
- Ortam: Windows; gerekli paketler `pandas` ve `openpyxl`.
- Girdi dosyaları: script çalıştığı klasör (`BASE_DIR`) ve bir üst klasörde `*.xls*` dosyalarını tarar. Dosya adı anahtar sözcüklerine göre (`KEYS`) eşleşme yapılır.

## Anahtar sabitler ve nerede değiştirilmeli

- `KEYS` — dosya adından hangi dosyanın 'faydalanma', 'kilavuz' veya 'tuketim' olduğunu belirler. Yeni bir dosya adlandırma kalıbı gelirse buraya ekleyin.
- `CANDIDATE_*` listeleri — `find_col` ile kullanılacak sütun aday isimleri. Örn: `CANDIDATE_BARKOD_COLS`, `C_TUKETIM_AMT`, `K_KOD`.
- `stok_desc_candidates` ve `stok_amount_candidates` — stokla ilgili sütun aramaları için kullanılır.

## Veri akışı (adım adım)

1. `find_files()` klasörlerde `*.xls*` arar ve `KEYS` ile eşleştirir.
2. `safe_read()` ile bulunan Excel dosyaları `pandas.read_excel` ile okunur (sheet adı eşitlemesi casefold ile yapılır).
3. `find_col(df, candidates)` ile dataframe içinden doğru sütun bulunur (tam eşleşme -> casefold -> substring).
4. `df_tuk` (tüketim) için `(işemri, barkod) -> miktar` haritası oluşturulur.
5. `df_kil` (kılavuz) için `işemri -> ürün kodu` ve (varsa) `işemri -> çap` haritaları oluşturulur.
6. `df_fayd` satır bazlı işlenir: kılavuz eşlemesi, tüketim miktarı atanır; ardından `STOK_KG`, `TUKETIM_KG`, `STOK_DURUM`, `EşMi` gibi yeni sütunlar hesaplanır.

## `parse_diameter_mm` davranışı ve örnekler

- Fonksiyonı, açıklama veya ürün kodu içinden milimetre cinsinden çapı çıkartmak için kullanın.
- Desteklenen örnekler: `"1.05mm"`, `"HT 1.05MM-CT-KY180"`, `"0,88"` gibi ondalıklı sayılar.
- Regex yaklaşımları: `...mm` son eki, `MM` büyük harf desenleri, ayrıca standalone decimal sayılar.

## Örnek Excel başlığı ve satır (CSV biçiminde örnek)

Başlık satırı (örnek):

```
İŞ EMRİ,BARKOD NUMARASI,STOK ÜRÜN TANIMI,STOK MİKTARI,KILAVUZ ÜRÜN KODU,ÜRÜN AÇIKLAMA
```

Örnek satır:

```
12345,TR-0001,TV 1.05MM-CT,120,HT 1.05MM-CT-KY180,Çelik tel 1.05mm
```

Not: gerçek giriş dosyaları Excel (.xls/.xlsx) olacaktır; yukarıdakiler CSV örneğidir.

## Değişiklik yaparken dikkat

- Yeni sütun adayları ekliyorsanız, `find_col`'un davranışını bozmayacak şekilde benzersiz, yaygın varyantlar ekleyin (büyük/küçük harf, boşluk/alt çizgi farkları).
- Eğer `safe_read` belirli bir sheet adıyla başarılı olamıyorsa, fonksiyon tüm sheet'leri okur ve casefold eşleşmesi ile doğru sheet'i seçmeye çalışır — bu davranış korunmalı.

## Hızlı doğrulama & komutlar

1. Gerekli paketleri yükleyin:

```powershell
pip install pandas openpyxl
```

2. Scripti çalıştırın:

```powershell
python faydalanma_atama.py
```

3. Oluşan dosyayı kontrol edin: `faydalanma_atama_sonuc_{timestamp}.xlsx` içinde `Faydalanma_Atama`, `Eksik_IsEmri`, `Eksik_Barkod` sheet'lerini inceleyin.

## Öneriler / İyileştirme noktaları

- `requirements.txt` ekleyerek ortam kurulumunu kolaylaştırın (ör: `pandas==1.5.3`, `openpyxl>=3.0.10`).
- Hata ve logging: `print` yerine `logging` kullanımı ve hata durumları için daha deterministik davranış önerilir.
- Birim test önerisi: `parse_diameter_mm`, `find_col` ve `linear_kg_per_m_from_d_mm` için küçük testler yazın.

---
Ekleme isteğiniz veya örnek Excel dosyası örneği isterseniz gönderin; yönergeyi buna göre daha da özelleştirip örnek satırları çoğaltırım.
