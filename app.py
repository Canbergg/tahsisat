import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

def process_excel(file):
    # Dosyayı yükle
    df = pd.read_excel(file)

    # Adım 1: 32. sütunun sağına iki yeni sütun ekle
    unique_count_column_index = 33  # Yeni Unique Count sütunu
    relation_column_index = 34  # Yeni İlişki sütunu
    df.insert(unique_count_column_index, "Unique Count", 0)
    df.insert(relation_column_index, "İlişki", "")

    # Unique Count hesapla (ItAtt48 -> 31. sütun, Mağaza Kodu -> 7. sütun)
    df.iloc[:, unique_count_column_index] = (
        df.groupby([df.iloc[:, 6], df.iloc[:, 30]])
        .transform("count")
        .iloc[:, 30]
    )

    # İlişki sütununu doldur (AF -> 31. sütun)
    df.iloc[:, relation_column_index] = np.select(
        [df.iloc[:, 31] == 11, df.iloc[:, 31] == 10, df.iloc[:, 31].isna()],
        ["Muadil", "Muadil stoksuz", "İlişki yok"],
        default=""
    )

    # Adım 2: Tekli sayfasını oluştur
    tekli = df[df.iloc[:, unique_count_column_index] == 1].copy()

    # Tekli için "İhtiyaç" hesapla
    def calculate_ihitiyac(row):
        try:
            # Sütun değerlerini sayısal veri tipine çevir
            donem_satis = pd.to_numeric(row.iloc[11], errors="coerce") or 0
            envanter_gun_sayisi = pd.to_numeric(row.iloc[20], errors="coerce") or 0
            zt_stock = pd.to_numeric(row.iloc[28], errors="coerce") or 0
            stock_be_completed = pd.to_numeric(row.iloc[30], errors="coerce") or 0
            minimum_miktar = pd.to_numeric(row.iloc[18], errors="coerce") or 0
            zt_qty = pd.to_numeric(row.iloc[27], errors="coerce") or 0
            kanal_envanter = pd.to_numeric(row.iloc[15], errors="coerce") or 0

            # Hesaplama
            result = max(
                0,
                round(
                    (donem_satis / envanter_gun_sayisi if envanter_gun_sayisi > 0 else 0) *
                    (zt_stock if zt_stock > 0 else stock_be_completed),
                    0
                ) + minimum_miktar + zt_qty - kanal_envanter
            )
            return result
        except Exception as e:
            st.write(f"Hata oluştu: {e}")
            return 0  # Hata durumunda varsayılan değer

    tekli["İhtiyaç"] = tekli.apply(calculate_ihitiyac, axis=1)

    # Çift sayfasını oluştur
    cift = df[df.iloc[:, unique_count_column_index] == 2].copy()

    # Çift için "İhtiyaç" hesapla
    def calculate_ihitiyac_cift(row1, row2):
        try:
            # Sütun değerlerini sayısal veri tipine çevir
            donem_satis1 = pd.to_numeric(row1.iloc[11], errors="coerce") or 0
            donem_satis2 = pd.to_numeric(row2.iloc[11], errors="coerce") or 0
            envanter_gun_sayisi1 = pd.to_numeric(row1.iloc[20], errors="coerce") or 0
            envanter_gun_sayisi2 = pd.to_numeric(row2.iloc[20], errors="coerce") or 0
            zt_stock2 = pd.to_numeric(row2.iloc[28], errors="coerce") or 0
            stock_be_completed2 = pd.to_numeric(row2.iloc[30], errors="coerce") or 0
            minimum_miktar2 = pd.to_numeric(row2.iloc[18], errors="coerce") or 0
            zt_qty1 = pd.to_numeric(row1.iloc[27], errors="coerce") or 0
            zt_qty2 = pd.to_numeric(row2.iloc[27], errors="coerce") or 0
            kanal_envanter1 = pd.to_numeric(row1.iloc[15], errors="coerce") or 0
            kanal_envanter2 = pd.to_numeric(row2.iloc[15], errors="coerce") or 0

            # Hesaplama
            result = max(
                0,
                round(
                    (donem_satis1 + donem_satis2) / max(envanter_gun_sayisi1, envanter_gun_sayisi2) *
                    (zt_stock2 if zt_stock2 > 0 else stock_be_completed2),
                    0
                ) + minimum_miktar2 + zt_qty1 + zt_qty2 - kanal_envanter1 - kanal_envanter2
            )
            return result
        except Exception as e:
            st.write(f"Hata oluştu (çift hesaplama): {e}")
            return 0  # Hata durumunda varsayılan değer

    # Çift sayfasında hesaplama
    cift_sorted = cift.sort_values(by=[df.columns[6], df.columns[30], df.columns[35]], ascending=[True, True, True])
    cift_sorted["İhtiyaç"] = ""
    for i in range(0, len(cift_sorted) - 1, 2):
        row1 = cift_sorted.iloc[i]
        row2 = cift_sorted.iloc[i + 1]
        value = calculate_ihitiyac_cift(row1, row2)
        cift_sorted.at[cift_sorted.index[i], "İhtiyaç"] = value
        cift_sorted.at[cift_sorted.index[i + 1], "İhtiyaç"] = ""

    # Çift sayfasına "İhtiyaç Çoklama" sütunu ekle
    cift_sorted["İhtiyaç Çoklama"] = cift_sorted["İhtiyaç"].fillna(method="ffill")

    # Excel dosyasını oluştur
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Ana Sayfa", index=False)
        tekli.to_excel(writer, sheet_name="Tekli", index=False)
        cift_sorted.to_excel(writer, sheet_name="Çift", index=False)
    output.seek(0)
    return output

# Streamlit UI
st.title("Excel İşleme Aracı")
uploaded_file = st.file_uploader("Excel dosyanızı yükleyin", type=["xlsx"])

if uploaded_file:
    st.success("Dosya başarıyla yüklendi, işlemler yapılıyor...")
    processed_file = process_excel(uploaded_file)
    st.download_button("İşlenmiş Dosyayı İndir", processed_file, file_name="Processed_File.xlsx")
