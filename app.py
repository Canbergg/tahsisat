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

    # Hata ayıklama: Satır ve sütun değerlerini kontrol et
    def debug_calculate_ihitiyac(row):
        try:
            # Kullanılan sütun değerlerini yazdır
            st.write({
                "Minimum Miktar (S)": row.iloc[18],
                "Dönem Satış (L)": row.iloc[11],
                "Envanter Gün Sayısı (U)": row.iloc[20],
                "ZtStockBeCompletedQty (AC)": row.iloc[28],
                "StockBeCompleted (AK)": row.iloc[30],
                "ZtQty (AB)": row.iloc[27],
                "Kanalda Envanter+Rezerv (P)": row.iloc[15]
            })

            # Formülü uygula
            return max(
                0,
                round(
                    (row.iloc[11] / row.iloc[20] if row.iloc[20] and row.iloc[20] != 0 else 0) *
                    (row.iloc[28] if row.iloc[28] and row.iloc[28] > 0 else row.iloc[30]),
                    0
                ) + (row.iloc[18] if row.iloc[18] else 0) + (row.iloc[27] if row.iloc[27] else 0) - (row.iloc[15] if row.iloc[15] else 0)
            )
        except Exception as e:
            # Hata oluşursa satır bilgilerini yazdır
            st.write("HATA!", {
                "Minimum Miktar (S)": row.iloc[18],
                "Dönem Satış (L)": row.iloc[11],
                "Envanter Gün Sayısı (U)": row.iloc[20],
                "ZtStockBeCompletedQty (AC)": row.iloc[28],
                "StockBeCompleted (AK)": row.iloc[30],
                "ZtQty (AB)": row.iloc[27],
                "Kanalda Envanter+Rezerv (P)": row.iloc[15]
            })
            st.write("Hata mesajı:", e)
            return 0

    # Tekli için "İhtiyaç" hesapla
    tekli["İhtiyaç"] = tekli.apply(debug_calculate_ihitiyac, axis=1)

    # Çift sayfasını oluştur (bu adımı daha sonra tamamlayacağız)
    cift = df[df.iloc[:, unique_count_column_index] == 2].copy()

    # Excel dosyasını oluştur
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Ana Sayfa", index=False)
        tekli.to_excel(writer, sheet_name="Tekli", index=False)
        cift.to_excel(writer, sheet_name="Çift", index=False)
    output.seek(0)
    return output

# Streamlit UI
st.title("Excel İşleme Aracı")
uploaded_file = st.file_uploader("Excel dosyanızı yükleyin", type=["xlsx"])

if uploaded_file:
    st.success("Dosya başarıyla yüklendi, işlemler yapılıyor...")
    processed_file = process_excel(uploaded_file)
    st.download_button("İşlenmiş Dosyayı İndir", processed_file, file_name="Processed_File.xlsx")
