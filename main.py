import streamlit as st
import pandas as pd
import docx
from utils.converter import convertDocx
from utils.function import analyze_word_document, sanitize_sheet_name
import re
from io import BytesIO


# Konfigurasi halaman
st.set_page_config(page_title="Konverter Laporan", layout="wide")

# --- Judul Aplikasi ---
st.title("🚀 Konverter Laporan Word ke Excel")
st.markdown("Aplikasi ini membantu Anda mengonversi laporan terstruktur dari file `.docx` menjadi file Excel yang rapi dengan beberapa sheet.")

# --- Langkah 1: Upload File ---
st.header("Langkah 1: Upload Dokumen Anda")
uploaded_file = st.file_uploader("Pilih file Word (.docx) yang ingin Anda proses", type="docx")

# Gunakan st.session_state untuk menyimpan status
if 'analysis_done' not in st.session_state:
    st.session_state.analysis_done = False

if uploaded_file is not None:
    # Lakukan analisis hanya jika file baru di-upload atau belum dianalisis
    if not st.session_state.analysis_done:
        with st.spinner('Menganalisis struktur dokumen... Mohon tunggu sebentar.'):
            # Buka file dari memori
            doc = docx.Document(uploaded_file)
            analysis_results = analyze_word_document(doc)
            
            # Simpan hasil analisis ke session state
            st.session_state.analysis_results = analysis_results
            st.session_state.analysis_done = True
        st.success("Analisis selesai!")

    # --- Langkah 2: Halaman Analisis & Konfirmasi ---
    if st.session_state.analysis_done:
        st.header("Langkah 2: Periksa Hasil Analisis & Konfirmasi")
        st.info("Berikut adalah ringkasan dari dokumen yang Anda upload. Pastikan informasi di bawah ini sudah sesuai sebelum melanjutkan.")

        results = st.session_state.analysis_results
        
        # Tampilkan metrik utama
        col1, col2, col3 = st.columns(3)
        col1.metric("Jumlah Laporan Terdeteksi", f"{results['report_count']} Laporan")
        col2.metric("Total Tabel Ditemukan", f"{results['total_table_count']} Tabel")
        col3.metric("Jumlah Judul/Sheet", f"{len([h for h, d in results['headings'].items() if d['table_count'] > 0])} Sheet")
        
        st.markdown("---")
        
        # Tampilkan rincian sheet yang akan dibuat
        st.subheader("Rincian Sheet yang Akan Dibuat di Excel")
        
        sheet_data = []
        for heading, data in results['headings'].items():
            if data['table_count'] > 0:
                sheet_name = sanitize_sheet_name(heading)
                sheet_data.append([heading, sheet_name, data['table_count']])
        
        if sheet_data:
            df_sheets = pd.DataFrame(sheet_data, columns=["Judul Asli (Heading 1)", "Nama Sheet di Excel", "Jumlah Tabel"])
            st.dataframe(df_sheets, use_container_width=True)
            st.warning("Jika ada bagian yang tidak terbaca, silahkan rubah format judul bagian itu menjadi 'Heading 1'!")
        else:
            st.warning("Tidak ada tabel yang ditemukan di bawah 'Heading 1' manapun.")

        st.markdown("---")

        # Tombol Konfirmasi
        st.subheader("Langkah 3: Atur Opsi & Mulai Proses Konversi")

        # --- TAMBAHKAN INI: Widget Input Angka ---
        # Ambil jumlah laporan yang terdeteksi dari hasil analisis
        jumlah_laporan_terdeteksi = results['report_count']

        # Buat number input dengan nilai default sesuai hasil analisis
        laporan_untuk_diproses = st.number_input(
            label="Jumlah laporan yang ingin dikonversi:",
            min_value=1,
            max_value=jumlah_laporan_terdeteksi,
            value=jumlah_laporan_terdeteksi, # Defaultnya adalah semua laporan
            step=1,
            help=f"Dokumen ini terdeteksi memiliki {jumlah_laporan_terdeteksi} laporan. Anda bisa memilih untuk memproses lebih sedikit dari awal."
        )
        # ---------------------------------------------

        # Tombol Konfirmasi
        if st.button("✅ Ya, Lanjutkan dan Konversi ke Excel", type="primary"):
            with st.spinner(f"Mohon tunggu, sedang memproses {laporan_untuk_diproses} laporan dari dokumen..."):
                try:
                    # PENTING: Reset posisi file sebelum membacanya lagi
                    uploaded_file.seek(0) 
                    doc = docx.Document(uploaded_file)

                    # --- PERBAIKAN PADA PEMANGGILAN FUNGSI ---
                    # 1. Hapus parameter 'output_filename' yang sudah tidak dipakai
                    # 2. Tambahkan parameter baru (misal: 'num_laporan') dengan nilai dari st.number_input
                    excel_data = convertDocx(doc, num_converted=laporan_untuk_diproses)
                    # ---------------------------------------------

                    # Cek tipe data dari hasil yang dikembalikan (sudah mendukung error handling)
                    if isinstance(excel_data, bytes):
                        st.success("✅ Konversi berhasil! File Excel Anda siap diunduh.")
                        st.download_button(
                            label="📥 Download File Excel",
                            data=excel_data,
                            file_name=f"hasil_konversi_{uploaded_file.name.split('.')[0]}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    elif isinstance(excel_data, str):
                        st.error("Gagal Melakukan Konversi!", icon="🚨")
                        st.markdown(excel_data) # Tampilkan pesan error detail
                    else: # Kasus None
                        st.warning("⚠️ Tidak ada data yang bisa diekstrak dari dokumen yang diunggah.")

                except Exception as e:
                    st.error(f"Terjadi kesalahan tak terduga: {e}")

else:       
    # Reset status jika tidak ada file
    st.session_state.analysis_done = False