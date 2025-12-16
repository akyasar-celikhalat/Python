import pandas as pd
import re

# Excel dosyasını okuma
file_path = 'üretim.xlsx'
df = pd.read_excel(file_path)

# Tarih sütunlarını birleştirerek yeni bir tarih sütunu oluşturma
def create_date(row):
    return pd.Timestamp(year=row['YIL_02'], month=row['AY_02'], day=row['GUN_02'])

df['TARIH'] = df.apply(create_date, axis=1)

# Malzeme sütunundan çap bilgisini çıkarma
def extract_diameter(malzeme):
    match = re.search(r'(\d+(\.\d+)?)MM', malzeme)
    if match:
        return float(match.group(1))
    return None

df['ÇAP'] = df['MALZEME ADI'].apply(extract_diameter)

# Çapın karesini hesaplama
df['ÇAP_KARE'] = df['ÇAP'] ** 2

# Çapın karesini metre değeriyle çarpma
df['ÇARPIM'] = df['ÇAP_KARE'] * df['METRE_02']

# Benzersiz kayıtları sayarak toplam vardiya sayısını hesaplama
unique_records = df.groupby(['MAKİNE_02', 'TARIH', 'VARDİYA_02']).size().reset_index(name='KAYIT_SAYISI')
daily_vardia_counts = unique_records.groupby(['MAKİNE_02', 'TARIH']).size().reset_index(name='TOPLAM_VARDİYA')
total_vardia_counts = daily_vardia_counts.groupby('MAKİNE_02')['TOPLAM_VARDİYA'].sum().reset_index()

# Üretim toplamlarını hesaplama
production_summary = df.groupby('MAKİNE_02').agg({
    'METRE_02': 'sum',
    'KG_02': 'sum',
    'ÇARPIM': 'sum'
}).reset_index()

# Sonuçları birleştirme
final_summary = pd.merge(production_summary, total_vardia_counts, on='MAKİNE_02', how='left')

# Ortalama günlük üretim hesaplama
final_summary['ORTALAMA_GÜNLÜK_METRE'] = (final_summary['METRE_02'] / final_summary['TOPLAM_VARDİYA']) * 3
final_summary['ORTALAMA_GÜNLÜK_KG'] = (final_summary['KG_02'] / final_summary['TOPLAM_VARDİYA']) * 3

# Çarpım değerlerini toplam metreye bölüp karekökünü alma
final_summary['KAREKÖK_CARPIM_METRE'] = (final_summary['ÇARPIM'] / final_summary['METRE_02']).apply(lambda x: x**0.5)

# ÇARPIM sütununu kaldırma
final_summary = final_summary.drop(columns=['ÇARPIM'])

# Tablo başlıklarını belirleme
final_summary.columns = [
    'Makine',
    'Uzunluk (m)',
    'Ağırlık (kg)',
    'Vardiya',
    'Ortalama Günlük Metre',
    'Ortalama Günlük Kilo',
    'Ortalama Çap'
]

# Sonuçları ekrana yazdırma
print(final_summary)

# Sonuçları bir Excel dosyasına yazma (isteğe bağlı)
output_file_path = 'toplam_uretim_ve_vardiya_raporu.xlsx'
final_summary.to_excel(output_file_path, index=False)
print(f"Rapor {output_file_path} dosyasına kaydedildi.")