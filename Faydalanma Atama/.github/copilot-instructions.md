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
