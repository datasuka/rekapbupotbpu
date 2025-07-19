import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Rekap Bukti Potong PDF", layout="wide")
st.title("📄 Rekap Bukti Potong PPh dari PDF ke Excel")

def extract_data_from_pdf(file):
    with pdfplumber.open(file) as pdf:
        text = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())

    try:
        data = {}
        data["NOMOR"] = re.search(r"BPPU\s+(\S+)", text).group(1)
        data["MASA PAJAK"] = re.search(rf"{data['NOMOR']}\s+(\d{{2}}-\d{{4}})", text).group(1)
        data["SIFAT PEMOTONGAN"] = re.search(r"(\bTIDAK FINAL\b|\bFINAL\b)", text).group(1)
        data["STATUS BUKTI"] = re.search(r"(NORMAL|PEMBETULAN)", text).group(1)

        data["NPWP / NIK"] = re.search(r"A\.1 NPWP / NIK\s*:\s*(\d+)", text).group(1)
        data["NAMA"] = re.search(r"A\.2 NAMA\s*:\s*(.+)", text).group(1).strip()
        data["NOMOR IDENTITAS TEMPAT USAHA"] = re.search(r"A\.3 NOMOR IDENTITAS.*?:\s*(\d+)", text).group(1)

        data["JENIS PPH"] = re.search(r"B\.2 Jenis PPh\s*:\s*(Pasal \d+)", text).group(1)
        data["KODE OBJEK"] = re.search(r"(\d{2}-\d{3}-\d{2})", text).group(1)
        data["OBJEK PAJAK"] = re.search(r"\d{2}-\d{3}-\d{2}\s+(.+)", text).group(1).split()[0]
        data["DPP"] = int(re.search(r"DPP\s*\(Rp\)\s*(\d[\d\.]*)", text).group(1).replace(".", ""))
        data["TARIF %"] = int(re.search(r"TARIF\s*\(%\)\s*(\d+)", text).group(1))
        data["PAJAK PENGHASILAN"] = int(re.search(r"PENGHASILAN\s*\(Rp\)\s*(\d[\d\.]*)", text).group(1).replace(".", ""))

        data["JENIS DOKUMEN"] = re.search(r"Jenis Dokumen\s*:\s*(.+)", text).group(1).strip()
        data["TANGGAL DOKUMEN"] = re.search(r"Tanggal\s*:\s*(\d{2} .+ \d{4})", text).group(1)

        data["NOMOR DOKUMEN"] = re.search(r"Nomor Dokumen\s*:\s*(.+)", text).group(1).strip()

        data["NPWP / NIK PEMOTONG"] = re.search(r"C\.1 NPWP / NIK\s*:\s*(\d+)", text).group(1)
        data["NOMOR IDENTITAS TEMPAT USAHA PEMOTONG"] = re.search(r"C\.2.*?:\s*(\d+)", text).group(1)
        data["NAMA PEMOTONG"] = re.search(r"C\.3.*?:\s*(.+)", text).group(1).strip()
        data["TANGGAL PEMOTONGAN"] = re.search(r"C\.4 TANGGAL\s*:\s*(\d{2} .+ \d{4})", text).group(1)
        data["NAMA PENANDATANGAN"] = re.search(r"C\.5 NAMA PENANDATANGAN\s*:\s*(.+)", text).group(1).strip()
        return data
    except Exception as e:
        st.warning(f"Gagal ekstrak data: {e}")
        return None

uploaded_files = st.file_uploader("Upload satu atau lebih file PDF", type="pdf", accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for file in uploaded_files:
        with st.spinner(f"Memproses {file.name}..."):
            result = extract_data_from_pdf(file)
            if result:
                result["FILE"] = file.name
                all_data.append(result)

    if all_data:
        df = pd.DataFrame(all_data)
        st.success(f"Berhasil mengekstrak {len(df)} bukti potong")
        st.dataframe(df)

        output = BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)
        st.download_button("⬇️ Unduh Excel", output, file_name="Rekap_Bukti_Potong.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.error("Tidak ada data berhasil diproses.")
