import pandas as pd

# Excel dosyalarını oku
duruslar_df = pd.read_excel("duruslar.xlsx")
uretim_df = pd.read_excel("uretim.xlsx")

# Zaman sütunlarını datetime tipine çevir
duruslar_df["DURUŞ BAŞLANGIÇ"] = pd.to_datetime(duruslar_df["DURUŞ BAŞLANGIÇ"])
duruslar_df["DURUŞ BİTİŞ"] = pd.to_datetime(duruslar_df["DURUŞ BİTİŞ"])
uretim_df["OLUŞTURMA ZAMANI"] = pd.to_datetime(uretim_df["OLUŞTURMA ZAMANI"])

# Duruş süresini hesapla (saat cinsinden)
duruslar_df["DURUŞ SÜRESİ_SAAT"] = (duruslar_df["DURUŞ BİTİŞ"] - duruslar_df["DURUŞ BAŞLANGIÇ"]).dt.total_seconds() / 3600

# Makine bazlı toplam duruş süresi
toplam_durus = duruslar_df.groupby("MAKİNE_02")["DURUŞ SÜRESİ_SAAT"].sum().reset_index()
toplam_durus.rename(columns={"MAKİNE_02": "MAKİNE"}, inplace=True)

# Üretim verilerini makine bazında grupla
uretim_grup = uretim_df.groupby("MAKİNE NO").agg(
    TOPLAM_URETIM_KG=("TEYİT MİKTARI Kg", "sum"),
    TUKETIM_FARK_KG=("TÜKETİM FARK MİKTARI", "sum")
).reset_index()
uretim_grup.rename(columns={"MAKİNE NO": "MAKİNE"}, inplace=True)

# Veri setlerini birleştir
sonuc_df = pd.merge(toplam_durus, uretim_grup, on="MAKİNE", how="outer").fillna(0)

# Performans metrikleri hesapla (önceki kod ile aynı)
TOPLAM_CALISMA_SURESI = 8  # Saat (varsayılan)
IDEAL_URETIM_HIZI = 100    # Kg/saat

sonuc_df["KULLANILABILIRLIK"] = (TOPLAM_CALISMA_SURESI - sonuc_df["DURUŞ SÜRESİ_SAAT"]) / TOPLAM_CALISMA_SURESI * 100
sonuc_df["GERCEK_URETIM_HIZI"] = sonuc_df["TOPLAM_URETIM_KG"] / (TOPLAM_CALISMA_SURESI - sonuc_df["DURUŞ SÜRESİ_SAAT"])
sonuc_df["PERFORMANS"] = (sonuc_df["GERCEK_URETIM_HIZI"] / IDEAL_URETIM_HIZI) * 100
sonuc_df["KALITE"] = (sonuc_df["TOPLAM_URETIM_KG"] / (sonuc_df["TOPLAM_URETIM_KG"] + sonuc_df["TUKETIM_FARK_KG"])) * 100
sonuc_df["OEE"] = (sonuc_df["KULLANILABILIRLIK"] * sonuc_df["PERFORMANS"] * sonuc_df["KALITE"]) / (100**2)

# Sonucu Excel'e kaydet
sonuc_df.to_excel("sonuc.xlsx", index=False, engine="openpyxl")
print("Sonuçlar 'sonuc.xlsx' dosyasına kaydedildi.")