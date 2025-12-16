import tkinter as tk
from tkinter import messagebox, filedialog
from barcode.codex import Code39  # Code39 sınıfını codex alt modülünden aktarın
from barcode.writer import ImageWriter
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch, mm
from io import BytesIO
from PIL import Image  # Pillow kütüphanesinden Image sınıfını içe aktarın
import tempfile  # Geçici dosya oluşturmak için
import os

# Özel sayfa boyutu (mm cinsinden)
SAYFA_GENİŞLİĞİ_MM = 120  # Örnek: 100 mm
SAYFA_YÜKSEKLİĞİ_MM = 85  # Örnek: 150 mm

def generate_barcodes():
    start_num = entry_start.get().strip()
    end_num = entry_end.get().strip()
    makine_no = entry_text1.get().strip()  # Kullanıcıdan alınan ilk metin
    merkez_besik_var = merkez_besik_checkbox_var.get()

    if not start_num or not end_num:
        messagebox.showwarning("Uyarı", "Lütfen başlangıç ve bitiş beşik numaralarını girin.")
        return
    if not makine_no:
        messagebox.showwarning("Uyarı", "Lütfen makine numarasını girin.")
        return
    try:
        start_num = int(start_num)
        end_num = int(end_num)
    except ValueError:
        messagebox.showerror("Hata", "Başlangıç ve bitiş beşik numaraları sayı olmalıdır.")
        return
    if start_num > end_num:
        messagebox.showerror("Hata", "Başlangıç beşik numarası bitiş beşik numarasından büyük olamaz.")
        return

    # PDF dosyasını oluştur
    pdf_path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
                                            title="PDF Dosyasını Kaydet")
    if not pdf_path:
        return
    try:
        # Özel sayfa boyutunu ayarla (mm -> inç dönüşümü)
        page_width = SAYFA_GENİŞLİĞİ_MM * mm
        page_height = SAYFA_YÜKSEKLİĞİ_MM * mm
        # Barkod boyutlarını sayfa boyutlarına göre ayarla
        max_barcode_width = page_width - 20 * mm  # Kenarlardan 10 mm boşluk bırak
        max_barcode_height = page_height / 1  # Sayfanın 1/4'ü kadar yükseklik
        c = canvas.Canvas(pdf_path, pagesize=(page_width, page_height))
        width, height = page_width, page_height
        y_position = height - 10 * mm  # Başlangıç y pozisyonu

        for num in range(start_num, end_num + 1):
            code = f"M{makine_no}-B{num}"
            ean = Code39(code, writer=ImageWriter(), add_checksum=False)
            # Barkodu oluştur ve belleğe kaydet
            barcode_bytes = BytesIO()
            ean.write(barcode_bytes)
            barcode_bytes.seek(0)
            # Barkod resmini aç
            barcode_img = Image.open(barcode_bytes)
            barcode_width, barcode_height = barcode_img.size
            # Barkod boyutlarını sayfa boyutlarına göre ölçekle
            scale_factor = min(max_barcode_width / barcode_width, max_barcode_height / barcode_height)
            scaled_width = barcode_width * scale_factor
            scaled_height = barcode_height * scale_factor
            # Eğer barkod sayfa dışına taşacaksa, yeni bir sayfa aç
            if y_position - scaled_height < 10 * mm:
                c.showPage()
                y_position = height - 10 * mm
            # Geçici bir dosya oluştur ve barkod resmini kaydet
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as temp_file:
                barcode_img.save(temp_file.name)
                temp_file_path = temp_file.name
            # Barkodu PDF'e ekleyin (geçici dosya yolunu kullanarak)
            c.drawImage(temp_file_path, 10 * mm, y_position - scaled_height,
                        width=scaled_width, height=scaled_height)
            # Geçici dosyayı sil
            os.unlink(temp_file_path)
            # Yeni pozisyonu güncelle
            y_position -= scaled_height + 15 * mm

        if merkez_besik_var:
            # Merkez beşik seçildiyse, son barkodu ekleyin
            code = f"M{makine_no}-M1"
            ean = Code39(code, writer=ImageWriter(), add_checksum=False)
            # Barkodu oluştur ve belleğe kaydet
            barcode_bytes = BytesIO()
            ean.write(barcode_bytes)
            barcode_bytes.seek(0)
            # Barkod resmini aç
            barcode_img = Image.open(barcode_bytes)
            barcode_width, barcode_height = barcode_img.size
            # Barkod boyutlarını sayfa boyutlarına göre ölçekle
            scale_factor = min(max_barcode_width / barcode_width, max_barcode_height / barcode_height)
            scaled_width = barcode_width * scale_factor
            scaled_height = barcode_height * scale_factor
            # Eğer barkod sayfa dışına taşacaksa, yeni bir sayfa aç
            if y_position - scaled_height < 10 * mm:
                c.showPage()
                y_position = height - 10 * mm
            # Geçici bir dosya oluştur ve barkod resmini kaydet
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as temp_file:
                barcode_img.save(temp_file.name)
                temp_file_path = temp_file.name
            # Barkodu PDF'e ekleyin (geçici dosya yolunu kullanarak)
            c.drawImage(temp_file_path, 10 * mm, y_position - scaled_height,
                        width=scaled_width, height=scaled_height)
            # Geçici dosyayı sil
            os.unlink(temp_file_path)
            # Yeni pozisyonu güncelle
            y_position -= scaled_height + 15 * mm

        c.save()
        messagebox.showinfo("Başarılı", f"Barkod(lar) başarıyla {pdf_path} PDF dosyasına kaydedildi.")
    except Exception as e:
        messagebox.showerror("Hata", str(e))
        print(e)

# Tkinter penceresini oluştur
root = tk.Tk()
root.title("Çoklu Barkod Oluşturucu")
root.geometry("400x300")

# İlk Metin Etiketi ve Giriş Alanı
label_text1 = tk.Label(root, text="Makine:")
label_text1.pack(pady=5)
entry_text1 = tk.Entry(root, width=40)
entry_text1.pack(pady=5)

# Başlangıç Numarası Etiketi ve Giriş Alanı
label_start = tk.Label(root, text="Başlangıç Beşik Numarası:")
label_start.pack(pady=5)
entry_start = tk.Entry(root, width=20)
entry_start.pack(pady=5)

# Bitiş Numarası Etiketi ve Giriş Alanı
label_end = tk.Label(root, text="Bitiş Beşik Numarası:")
label_end.pack(pady=5)
entry_end = tk.Entry(root, width=20)
entry_end.pack(pady=5)

# Merkez Beşik Checkbox
merkez_besik_checkbox_var = tk.BooleanVar()
merkez_besik_checkbox = tk.Checkbutton(root, text="Merkez Beşik", variable=merkez_besik_checkbox_var)
merkez_besik_checkbox.pack(pady=5)

# Barkod Oluştur Butonu
generate_button = tk.Button(root, text="Barkodları Oluştur", command=generate_barcodes)
generate_button.pack(pady=10)

# Uygulamayı başlat
root.mainloop()