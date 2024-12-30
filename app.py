import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

def process_excel(file):
    # Dosyayı yükle
    df = pd.read_excel(file)

    # Adım 1: AF'nin sağına iki sütun ekle
    unique_count_column_index = 40  # Yeni Unique Count sütunu
    relation_column_index = 41  # Yeni İlişki sütunu
    df.insert(unique_count_column_index, "Unique Count", 0)
    df.insert(relation_column_index, "İlişki", "")

    # Unique Count hesapla (AE -> 30. sütun, Mağaza Kodu -> 1. sütun)
    df.iloc[:, unique_count_column_index] = (
        df.groupby([df.iloc[:, 1], df.iloc[:, 30]])
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
            return max(
                max(
                    (row.iloc[18] > 0) * round(
                        (row.iloc[11] / row.iloc[35] if row.iloc[35] != 0 else 0) * (row.iloc[28] if row.iloc[28] > 0 else row.iloc[30]), 0
                    ) + row.iloc[18] + row.iloc[26] - row.iloc[15],
                    0
                ), 0
            )
        except:
            return 0
    tekli["İhtiyaç"] = tekli.apply(calculate_ihitiyac, axis=1)

    # Adım 3: Çift sayfasını oluştur
    cift = df[df.iloc[:, unique_count_column_index] == 2].copy()

    # Çift sayfasını Mağaza Adı (7), ItAtt48 (31), ve Ürün Brüt Ağırlık (36) sırasına göre sırala
    cift_sorted = cift.sort_values(by=[df.columns[6], df.columns[30], df.columns[35]], ascending=[True, True, True])

    # Çiftli satırlar için "İhtiyaç" hesapla
    cift_sorted["İhtiyaç"] = ""
    for i in range(0, len(cift_sorted) - 1, 2):
        row1 = cift_sorted.iloc[i]
        row2 = cift_sorted.iloc[i + 1]
        value = max(
            max(
                sum([row1.iloc[18], row2.iloc[18]]) > 0 * round(
                    sum([row1.iloc[11], row2.iloc[11]]) / max(row1.iloc[35], row2.iloc[35]) * (row2.iloc[28] if row2.iloc[28] > 0 else row2.iloc[30]),
                    0
                ) + row2.iloc[18] + sum([row1.iloc[26], row2.iloc[26]]) - sum([row1.iloc[15], row2.iloc[15]]),
                0
            ), 0
        )
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
