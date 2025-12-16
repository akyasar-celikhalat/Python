import pandas as pd
import openpyxl
from datetime import datetime

def create_excel_report(data, output_file):
    """
    Excel raporu oluşturan temel fonksiyon
    
    Parameters:
    data (dict): Verilerin bulunduğu sözlük
    output_file (str): Çıktı dosyasının adı
    """
    # Örnek veri oluşturma (siz kendi verinizi kullanacaksınız)
    df = pd.DataFrame(data)
    
    # Excel yazıcı oluşturma
    writer = pd.ExcelWriter(output_file, engine='openpyxl')
    
    # Verileri Excel'e yazma
    df.to_excel(writer, sheet_name='Rapor', index=False)
    
    # Excel çalışma kitabını al
    workbook = writer.book
    worksheet = writer.sheets['Rapor']
    
    # Başlıkları biçimlendirme
    for col in worksheet.iter_cols(1, df.shape[1]):
        worksheet[f"{col[0].column_letter}1"].font = openpyxl.styles.Font(bold=True)
        worksheet.column_dimensions[col[0].column_letter].width = 15
    
    # Rapor oluşturma tarihini ekle
    worksheet['A' + str(df.shape[0] + 3)] = f"Rapor Tarihi: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    
    # Değişiklikleri kaydet
    writer.close()

# Örnek kullanım
if __name__ == "__main__":
    # Örnek veri
    sample_data = {
        'Ürün': ['Laptop', 'Mouse', 'Klavye', 'Monitor'],
        'Miktar': [10, 50, 30, 15],
        'Fiyat': [15000, 200, 500, 3000],
        'Toplam': [150000, 10000, 15000, 45000]
    }
    
    # Rapor oluştur
    create_excel_report(sample_data, 'rapor.xlsx')