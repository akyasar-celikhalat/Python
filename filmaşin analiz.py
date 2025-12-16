import pandas as pd

# Excel dosyasını okuma
# file_path = 'ExcelReport (58).xlsx'
file_path = r"C:\Users\yak\Desktop\Python\ExcelReport (58).xlsx"
df = pd.read_excel(file_path)

# Giriş etiketine göre gruplama ve her bir giriş etiketi için kaç defa üretildiğini hesaplama
grouped = df.groupby('GİRİŞETİKET').agg(
    Üretim_Sayısı=('GİRİŞETİKET', 'size'),
    Toplam_Miktar=('KİLO', 'sum')
).reset_index()

# En çok tekrar yapan operatör bilgisini alma
opr_counts = df['OPR.'].value_counts().reset_index()
opr_counts.columns = ['OPR.', 'Tekrar_Sayısı']
most_frequent_opr = opr_counts.iloc[0]

# Raporlama için sonuçları birleştirme
result = grouped.merge(df[['GİRİŞETİKET', 'GİRİŞÜRÜNKODU', 'ÇIKANETİKET', 'ÇIKIŞÜRÜNKODU']].drop_duplicates(), on='GİRİŞETİKET')

# Sonuçları yazdırma
print("Giriş Etiketine Göre Üretim Sayısı ve Toplam Miktar:")
print(result)

print("\nEn Çok Tekrar Eden Operatör:")
print(f"Operatör No: {most_frequent_opr['OPR.']}, Tekrar Sayısı: {most_frequent_opr['Tekrar_Sayısı']}")

# İsterseniz sonucu yeni bir Excel dosyasına kaydedebilirsiniz
result.to_excel('Üretim_Raporu.xlsx', index=False)