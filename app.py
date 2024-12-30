import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

def process_excel(file):
    # Dosyayı yükle
    df = pd.read_excel(file)

    # Adım 1: Unique Count sütununu ekle
    # AE -> 30. sütun, Mağaza Kodu -> 1. sütun
    df["Unique Count"] = df.iloc[:, 30].map(df.iloc[:, 30].value_counts()) / df.iloc[:, 1].nunique()

    # Adım 2: İlişki sütununu ekle
    # AF -> 31. sütun
    df["İlişki"] = np.select(
        [df.iloc[:, 31] == 11, df.iloc[:, 31] == 10, df.iloc[:, 31].isna()],
        ["Muadil", "Muadil stoksuz", "İlişki yok"],
        default=""
    )

    # Adım 3: Tekli sayfasını oluştur
    tekli = df[df["Unique Count"] == 1].copy()

    # Tekli için "İhtiyaç" hesapla
    def calculate_ihitiyac(row):
        try:
            return max(
                max(
                    (row["S"] > 0) * round(
                        (row["L"] / row["U"] if row["U"] != 0 else 0) * (row["AC"] if row["AC"] > 0 else row["AK"]), 0
                    ) + row["S"] + row["AB"] - row["P"],
                    0
                ), 0
            )
        except:
            return 0
    tekli["İhtiyaç"] = tekli.apply(calculate_ihitiyac, axis=1)

    # Adım 4: Çift sayfasını oluştur
    cift = df[df["Unique Count"] == 2].copy()
    cift_sorted = cift.sort_values(by=["Mağaza Adı", "ItAtt48", "Ürün Brüt Ağırlık"], ascending=[True, True, True])

    # Çiftli satırlar için "İhtiyaç" hesapla
    cift_sorted["İhtiyaç"] = ""
    for i in range(0, len(cift_sorted) - 1, 2):
        row1 = cift_sorted.iloc[i]
        row2 = cift_sorted.iloc[i + 1]
        value = max(
            max(
                sum([row1["S"], row2["S"]]) > 0 * round(
                    sum([row1["L"], row2["L"]]) / max(row1["U"], row2["U"]) * (row2["AC"] if row2["AC"] > 0 else row2["AK"]),
                    0
                ) + row2["S"] + sum([row1["AB"], row2["AB"]]) - sum([row1["P"], row2["P"]]),
                0
            ), 0
        )
        cift_sorted.loc[cift_sorted.index[i], "İhtiyaç"] = value
        cift_sorted.loc[cift_sorted.index[i + 1], "İhtiyaç"] = ""

    # Çift için İhtiyaç Çoklama
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
