import pandas as pd
import re

# Excel dosyasını okuma
file_path = '2025-11.xlsx'
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

# Makine sırası belirleme
makine_sirasi =  ['M301','M304','M305','M306','M307','M308','M309','M310','M311','M315','M321','M322','M323','M324','M323 - M324','M325','M326','M327','M328','M330','M331','M332','M349','M350','M351','M350 - M351','M354','M355','M354 - M355','M356','M359','M391','M394','M395','M396','M395 - M396','M397','M398','M399','M401','M424','M432','M459','M460','M494','M456','M491','M496','M497','M498','M501','M502','M503','M504','M505','M506','M507','M508','M509','M532','M101','M102','M103','M104','M105','M106','M108','M109','M116','M113','M114','M115','M117','M119','M120','M121','M122','M123','M124','M125','M126','M127','M128','M129','M130','M131','M132','M156','M160','M170','M171','M172','M173','M175','M176','M177','M178','M179','M180','M181','M183','M184','M185','M186','M187','M188','M189','M190','M191','M192','M193','M194','M195','M196','M197','M198','M199','M412','M133','M134','M511','M512','M513','M514','M515','M516','M201','M202','M200']

# Makine sırasına göre sıralama ve listelenmeyen makineleri en alta yazma
final_summary['Makine_Sira'] = final_summary['MAKİNE_02'].apply(lambda x: makine_sirasi.index(x) if x in makine_sirasi else len(makine_sirasi))
final_summary = final_summary.sort_values(by='Makine_Sira').drop(columns=['Makine_Sira'])

# Tablo başlıklarını belirleme
final_summary.columns = [
    'Makine',
    'Uzunluk (m)',
    'Ağırlık (kg)',
    'Vardiya',
    'Günlük U. (m)',
    'Günlük A. (kg)',
    'Ortalama Çap'
]

# Sonuçları ekrana yazdırma
# print(final_summary)

# Sonuçları bir Excel dosyasına yazma (isteğe bağlı)
output_file_path = 'kapabilite-2025-11.xlsx'
final_summary.to_excel(output_file_path, index=False)
print(f"Rapor {output_file_path} dosyasına kaydedildi.")