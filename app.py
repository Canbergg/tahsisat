pip install streamlit openpyxl pandas numpy
streamlit run app.py
import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import Workbook, load_workbook
from io import BytesIO

# Define the processing function
def process_excel(file):
    # Load the file into pandas
    data = pd.read_excel(file, sheet_name=None)  # Load all sheets
    main_sheet = list(data.keys())[0]  # Assume first sheet is the main one
    df = data[main_sheet]

    # Step 1: Add "Unique Count" and "İlişki" columns
    df["Unique Count"] = df.groupby("AE")["AE"].transform("count") / df["Mağaza Kodu"].nunique()
    df["İlişki"] = np.select(
        [df["AF"] == 11, df["AF"] == 10, df["AF"].isna()],
        ["Muadil", "Muadil stoksuz", "İlişki yok"],
        default=""
    )

    # Step 2: Create "Tekli" sheet
    tekli = df[df["Unique Count"] == 1]
    tekli["İhtiyaç"] = tekli.apply(lambda row: max(
        max(
            (row["S"] > 0) * round(
                (row["L"] / row["U"] if row["U"] != 0 else 0) * (row["AC"] if row["AC"] > 0 else row["AK"]), 0
            ) + row["S"] + row["AB"] - row["P"],
            0
        ), 0
    ), axis=1)

    # Step 3: Create "Çift" sheet
    cift = df[df["Unique Count"] == 2]
    cift_sorted = cift.sort_values(by=["Mağaza Adı", "ItAtt48", "Ürün Brüt Ağırlık"], ascending=[True, True, True])

    # Step 4: Add "İhtiyaç" column and calculate for merged rows
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

    cift_sorted["İhtiyaç Çoklama"] = cift_sorted["İhtiyaç"].fillna(method="ffill")

    # Write back to Excel
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
    st.success("Dosya yüklendi. İşlemler yapılıyor...")
    processed_file = process_excel(uploaded_file)
    st.download_button("İşlenmiş Dosyayı İndir", processed_file, file_name="Processed_File.xlsx")