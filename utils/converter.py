import io
from docx import Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.text.paragraph import Paragraph
from docx.table import Table
from utils.function import process_table
import pandas as pd

# --- Ganti dengan nama file Word Anda ---
all_heading_found = set()
# -----------------------------------------

# Gunakan dictionary untuk memetakan heading ke nama prosesornya
# Ini jauh lebih rapi daripada if/elif yang panjang!
HEADING_PROCESSORS = {
    "I. DATA UMUM": "Tipe Data Umum",
    "GEOLOGI UMUM": "Tipe Geologi Umum",
    "INFORMASI LOKASI": "Tipe Informasi Lokasi",
    "KOORDINAT WILAYAH": "Tipe Koordinat Wilayah",
    "INFORMASI LEMBAR PETA DAN CITRA": "Tipe Informasi Lembar Peta dan Citra",
    "PENYELIDIK TERDAHULU": "Tipe Penyelidik Terdahulu",
    "INFORMASI PENGISI DATA / LAPORAN": "Tipe Informasi Pengisi Data / Laporan",
    # Batubara
    "JENIS DAN TAHAPAN EKSPLORASI BATUBARA": "Tipe Eksplorasi Batubara",
    "FORMASI PEMBAWA LAPISAN BATUBARA": "Tipe Formasi Pembawa Batubara",
    "III.  INFORMASI LAPISAN BATUBARA (Umum)": "Tipe Informasi Lapisan Batubara",
    "KOORDINAT BLOK WILAYAH": "Tipe Koordinat Blok Wilayah",
    "INFORMASI KUANTITAS BLOK LAPISAN BATUBARA": "Tipe Kuantitas Blok Batubara",
    "INFORMASI KUALITAS BLOK LAPISAN BATUBARA (Analisa Proksimat)": "Tipe Kualitas Batubara (Proksimat)",
    "INFORMASI KUALITAS BLOK LAPISAN BATUBARA (Analisa Ultimat)": "Tipe Kualitas Batubara (Ultimat)",
    "INFORMASI KUALITAS BLOK BATUBARA (Analisa Petrografi)": "Tipe Kualitas Batubara (Petrografi)",
    # Bitumen
    "JENIS DAN TAHAPAN EKSPLORASI BITUMEN PADAT": "Tipe Eksplorasi Bitumen",
    "FORMASI PEMBAWA LAPISAN BITUMEN PADAT": "Tipe Formasi Pembawa Bitumen",
    "III.  INFORMASI LAPISAN BITUMEN PADAT (Umum)": "Tipe Informasi Lapisan Bitumen",
    "INFORMASI KUANTITAS BLOK LAPISAN BITUMEN PADAT": "Tipe Kuantitas Blok Bitumen",
    "INFORMASI KUALITAS BLOK BITUMEN PADAT (Analisa Retorting)": "Tipe Kualitas Bitumen (Retorting)",
    # CBM
    "JENIS DAN TAHAPAN EKSPLORASI CBM": "Tipe Eksplorasi CBM",
    # Lainnya
    "INFORMASI TITIK": "Tipe Informasi Titik"
}

def convertDocx(file,output_filename='output.xlsx',num_converted=None):
    try:
        doc = file
        tabel_ke = 1
        print(f"--- Menganalisis Dokumen Berdasarkan 'Heading 1' sebagai Section ---")
        print("="*60)
        active_heading_text = None
        table_counter_in_section = 0 # <-- PENGHITUNG TABEL
        formulir_id_counter = 0
        current_formulir_id = 0
        data_sheets = {}

        for block in doc.element.body:
            if isinstance(block, CT_P):
                p = Paragraph(block, doc)
                if p.style.name.startswith('Heading 1'):
                    active_heading_text = p.text.strip()
                    table_counter_in_section = 0 # <-- PENGHITUNG TABEL
                    print(f"\n--- Memasuki Bagian: '{active_heading_text}' ---")
                    if(active_heading_text == 'I. DATA UMUM'):
                        formulir_id_counter += 1
                        current_formulir_id = formulir_id_counter
                        print(f"\n--- Memproses Laporan Baru (ID: {current_formulir_id}) ---")
                        if num_converted is not None and formulir_id_counter > num_converted:
                            break
                        
            elif isinstance(block, CT_Tbl):
                if active_heading_text:
                    all_heading_found.add(active_heading_text)
                    t = Table(block, doc)

                    df_hasil = None
                    print(f"  > Tabel ditemukan (Ukuran: {len(t.rows)}x{len(t.columns)})")
                    print(f"  > Tabel ini milik judul: '{active_heading_text}'")
                    print(f"  > Urutan table : {tabel_ke}")

                    processor_name = "Jenis tabel tidak dikenali dari judulnya."
                    # Loop melalui kamus untuk menemukan prosesor yang cocok
                    for heading_key, name in HEADING_PROCESSORS.items():
                        if heading_key in active_heading_text:
                            processor_name = name
                            # Jika sudah ketemu, hentikan loop
                            break 

                    try :
                        params_for_log = {}
                        print(f"  --> Memproses sebagai {processor_name}")
                        if processor_name == 'Tipe Data Umum':
                            df_hasil  = process_table(t, onerow=True)
                        elif processor_name == 'Tipe Geologi Umum':
                            df_hasil  = process_table(t, onerow=True)
                        elif processor_name == 'Tipe Formasi Pembawa Batubara':
                            df_hasil  = process_table(t, onerow=True)
                        elif processor_name == 'Tipe Formasi Pembawa Bitumen':
                            df_hasil  = process_table(t, onerow=True)
                        elif processor_name == 'Tipe Informasi Lokasi':
                            df_hasil  = process_table(t, onerow=True)
                        elif processor_name == 'Tipe Koordinat Wilayah':
                            df_hasil  = process_table(t)
                        elif processor_name == 'Tipe Informasi Lembar Peta dan Citra':
                            df_hasil  = process_table(t,onerow=True)
                        elif processor_name == 'Tipe Eksplorasi Batubara':
                            df_hasil  = process_table(t,category=True)
                        elif processor_name == 'Tipe Eksplorasi Bitumen':
                            df_hasil  = process_table(t,category=True)
                        elif processor_name == 'Tipe Penyelidik Terdahulu':
                            df_hasil  = process_table(t)
                        elif processor_name == 'Tipe Informasi Lapisan Batubara':
                            df_hasil  = process_table(t,onerow=True)
                        elif processor_name == 'Tipe Informasi Lapisan Bitumen':
                            df_hasil  = process_table(t,onerow=True)
                        elif processor_name == 'Tipe Koordinat Blok Wilayah':
                            if table_counter_in_section == 0:
                                processor_name = 'Koordinat Wilayah'
                                df_hasil  = process_table(t)
                            else:
                                processor_name ='Tipe Blok Wilayah'
                                df_hasil  = process_table(t,onerow=True,paired=True)
                        elif processor_name == 'Tipe Kuantitas Blok Batubara':
                            df_hasil  = process_table(t,quantity_mode=True)
                        elif processor_name == 'Tipe Kualitas Batubara (Proksimat)':
                            df_hasil  = process_table(t,quantity_mode=True)
                        elif processor_name == 'Tipe Kualitas Batubara (Petrografi)':
                            df_hasil  = process_table(t)
                        elif processor_name == 'Tipe Informasi Titik':
                            df_hasil  = process_table(t,quantity_mode=True)
                        elif processor_name == 'Tipe Kualitas Batubara (Ultimat)':
                            df_hasil  = process_table(t,quantity_mode=True)
                        elif processor_name == 'Tipe Kuantitas Blok Bitumen':
                            df_hasil  = process_table(t,quantity_mode=True)
                        elif processor_name == 'Tipe Kualitas Bitumen (Retorting)':
                            df_hasil  = process_table(t,quantity_mode=True)
                        elif processor_name == 'Tipe Informasi Pengisi Data / Laporan':
                            df_hasil  = process_table(t)
                    except Exception as e:
                        # --- PERUBAHAN 3: Jika terjadi error, buat pesan dan return ---
                        error_message = (
                            f"--- 🔴 TERJADI ERROR ---\n"
                            f"Operasi dihentikan karena ada kesalahan saat pemrosesan.\n\n"
                            f"**Detail Kesalahan:**\n"
                            f" - **Tipe Error:** `{type(e).__name__}`\n"
                            f" - **Pesan:** `{e}`\n\n"
                            f"**Konteks Proses:**\n"
                            f" - **Lokasi:** Saat memproses **tabel ke-{tabel_ke}** (secara keseluruhan).\n"
                            f" - **Judul Bagian:** `{active_heading_text}`\n"
                            f" - **Tipe Prosesor:** `{processor_name}`\n"
                            f" - **Parameter untuk `process_table`:** `{params_for_log if params_for_log else 'Tidak ada (default)'}`"
                        )
                        print(error_message) # Cetak juga di konsol server untuk logging
                        return error_message # Hentikan fungsi dan kembalikan pesan error

                    if df_hasil is not None and not df_hasil.empty:
                        print(f"  --> Tabel diproses, {df_hasil.shape[0]} baris data dihasilkan.")
                        # 1. Tambahkan Foreign Key
                        df_hasil['formulir_id'] = current_formulir_id

                        # 2. Tentukan Nama Sheet
                        sheet_name = processor_name.replace('Tipe ', '').strip()
                        sheet_name = sheet_name[:31] 
                        sheet_name = sheet_name.replace('/', '_').replace('\\', '_')

                        # 3. Simpan/Gabungkan ke "Wadah" data_sheets
                        if sheet_name not in data_sheets:
                            data_sheets[sheet_name] = df_hasil
                        else:
                            # Gabungkan DataFrame yang baru dengan yang sudah ada
                            df_hasil = df_hasil.loc[:, ~df_hasil.columns.duplicated()]
                            data_sheets[sheet_name] = pd.concat([data_sheets[sheet_name], df_hasil], ignore_index=True)

                    table_counter_in_section += 1
                    tabel_ke += 1
                else:
                    print("\n  > Peringatan: Tabel ditemukan sebelum ada 'Heading 1'. Dilewati.")

        print("\n" + "="*60)
        print("--- Analisis Selesai ---")
        print("Export to Excel...")
        if data_sheets:
            output_buffer = io.BytesIO()
            print(f"\n--- Menyimpan data ke file Excel: {output_filename} ---")
            with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
                        for sheet_name, final_df in data_sheets.items():
                            print(f"  > Menulis sheet: '{sheet_name}' ({final_df.shape[0]} baris)")

                            # ... (kode penataan ulang kolom 'formulir_id' yang sudah kita perbaiki) ...
                            has_formulir_id = any('formulir_id' in str(col) for col in final_df.columns)
                            if has_formulir_id:
                                formulir_id_col = [col for col in final_df.columns if 'formulir_id' in str(col)][0]
                                other_cols = [col for col in final_df.columns if col != formulir_id_col]
                                final_df = final_df[[formulir_id_col] + other_cols]

                            # --- BLOK PERBAIKAN UNTUK MultiIndex ---
                            # Cek apakah kolomnya adalah MultiIndex
                            if isinstance(final_df.columns, pd.MultiIndex):
                                # Ratakan header: ('SUMBERDAYA', 'Hipotetik') -> 'SUMBERDAYA Hipotetik'
                                final_df.columns = [' '.join(col).strip() for col in final_df.columns.values]
                            # --- AKHIR BLOK PERBAIKAN ---

                            final_df.to_excel(writer, sheet_name=sheet_name, index=False)
            print("\n✅ Proses Selesai! File Excel berhasil dibuat.")
            return output_buffer.getvalue()
        else:
            print("\n⚠️ Tidak ada data yang diekstrak untuk disimpan ke Excel.")

    except FileNotFoundError:
        print(f"Error: File tidak ditemukan.")


