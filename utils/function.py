import docx
import pandas as pd
import docx
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.text.paragraph import Paragraph
import re
import numpy as np
# from docx.table import Table # Tambahkan ini untuk type hinting jika suka

# Daftar konstanta bisa diletakkan di atas agar mudah diubah
PAIR_COLUMNS = ['Tebal Lapisan BB (m)', 'Jurus (o)', 'Kemiringan (o)']
PAIR_KEYS = ['maks', 'min', 'rata-rata']

#=================================================================#
# HELPER FUNCTIONS (Tidak perlu diubah)
#=================================================================#

def get_cell_content(cell):
    """Mengekstrak konten dari sel, termasuk deteksi checkbox."""
    text = cell.text.strip()
    if text:
        return text
    try:
        xml_str = str(cell._element.xml)
        if "<w:sym" in xml_str or "<w:checked" in xml_str:
            return '✓'
    except Exception as e:
        print(f"Peringatan: Gagal membaca XML untuk sebuah sel. Error: {e}")
    return ''

def parse_paired_data_row(row):
    """Mem-parsing satu baris yang berisi beberapa pasangan key-value."""
    paired_data = {}
    cells = list(row)
    for i in range(len(cells)):
        key = cells[i].strip().lower().replace('.', '')
        if key in PAIR_KEYS:
            for j in range(i + 1, len(cells)):
                value = cells[j].strip()
                if value and value not in PAIR_KEYS:
                    paired_data[key] = value
                    break
    return paired_data

#=================================================================#
# FUNGSI-FUNGSI UTAMA YANG SUDAH DISESUAIKAN
#=================================================================#

def table_to_dataframe(table_obj):
    """
    Mengubah objek tabel docx menjadi DataFrame mentah.
    Ini adalah pengganti dari 'extract_word_table'.
    """
    if not table_obj:
        return None
    data = [[get_cell_content(cell) for cell in row.cells] for row in table_obj.rows]
    return pd.DataFrame(data)

def process_cleaned_data(df_messy, paired=False):
    """
    Membersihkan DataFrame dengan logika dua kondisi.
    Nama diubah dari 'clean_extracted_data' agar lebih jelas.
    """
    if df_messy is None:
        return None
        
    df = df_messy.copy()
    df.columns = df.iloc[0]
    df_data = df.iloc[2:].reset_index(drop=True)
    
    if df_data.empty:
        return pd.DataFrame(columns=df.columns)

    final_data = []
    
    for col_name in df_data.columns:
        initial_value = df_data.iloc[0][col_name]
        column_data = df_data[col_name]
        if isinstance(column_data, pd.DataFrame):
            # Jika ini DataFrame, ambil kolom pertamanya saja sebagai Series
            column_data = column_data.iloc[:, 0]
        else:
            # Jika sudah Series, langsung gunakan
            column_data = column_data

        if paired and col_name in PAIR_COLUMNS:
            paired_dict = parse_paired_data_row(df_data[col_name]) 
            final_data.append(paired_dict)
            continue
        
        try:
            # Logika pencarian '✓' tetap sama persis
            column_data.tolist().index('✓') # Cek cepat apakah ada centang
            looking_for_value = False
            all_values = []
            for value in column_data:
                if looking_for_value:
                    if pd.notna(value) and value != '✓' and value != '':
                        all_values.append(value)
                        looking_for_value = False
                if value == '✓':
                    looking_for_value = True
            
            if len(all_values) > 1:
                final_data.append(all_values) 
            elif all_values:
                final_data.append(all_values[0])
            else:
                final_data.append(None) # Jika ada centang tapi tidak ada nilai

        except ValueError:
            final_data.append(initial_value)
            
    df_clean = pd.DataFrame([final_data], columns=df_data.columns)
    return df_clean

def process_grid_table(table_obj):
    """
    Memproses tabel dengan struktur grid (kategori di kolom pertama).
    Sekarang menerima 'table_obj' langsung.
    """
    if not table_obj:
        return None
    
    clean_data = {}
    last_category = ""

    for row in table_obj.rows:
        current_category_text = get_cell_content(row.cells[0]).strip()
        if current_category_text and current_category_text != ':':
            last_category = current_category_text.replace('*', '').strip()
            if last_category not in clean_data:
                clean_data[last_category] = []
        
        if not last_category:
            continue

        row_contents = [get_cell_content(c) for c in row.cells]
        
        if '✓' in row_contents:
            # ... Logika untuk checkbox tetap sama persis ...
            checked_values_in_row = []
            for i, cell_content in enumerate(row_contents):
                if cell_content == '✓':
                    found_value = ""
                    if i + 1 < len(row.cells):
                        value_on_right = get_cell_content(row.cells[i+1]).strip()
                        if value_on_right: found_value = value_on_right
                    if not found_value and i > 0:
                        value_on_left = get_cell_content(row.cells[i-1]).strip()
                        if value_on_left and value_on_left != ':': found_value = value_on_left
                    if found_value: checked_values_in_row.append(found_value)
            
            if checked_values_in_row:
                clean_data[last_category].extend(checked_values_in_row)
        else:
            # ... Logika untuk baris teks biasa tetap sama persis ...
            if len(row.cells) > 2:
                direct_value = get_cell_content(row.cells[2]).strip()
                if direct_value and not clean_data[last_category]:
                    clean_data[last_category].append(direct_value)
            
    return pd.DataFrame(dict([(k, pd.Series(v)) for k, v in clean_data.items()]))

def process_resource_table(df_messy):
    """
    Membersihkan tabel sumberdaya (multi-header).
    Fungsi ini sudah bagus karena menerima DataFrame, jadi tidak ada perubahan.
    Nama diubah dari 'bersihkan_tabel_sumberdaya'.
    """
    if df_messy is None or df_messy.shape[0] < 3:
        return None
    
    # ... Logika internalnya tetap sama persis ...
    header_level1 = df_messy.iloc[0].ffill()
    header_level2 = df_messy.iloc[1].str.replace('\n', ' ', regex=False)
    multi_index = pd.MultiIndex.from_arrays([header_level1, header_level2])

    df_data = df_messy.iloc[2:].reset_index(drop=True)
    df_data.columns = multi_index
    df_clean = df_data.copy()

    summary_col_name = ('', 'Nama Lapisan')
    if summary_col_name in df_clean.columns:
        df_clean = df_clean[df_clean[summary_col_name] != 'JUMLAH']

    df_clean.replace('-', np.nan, inplace=True)
    cols_to_drop = [col for col in [('', 'No'), ('', 'Metoda Estimasi')] if col in df_clean.columns]
    df_clean.drop(columns=cols_to_drop, inplace=True)

    info_cols = [col for col in [('', 'NAMA BLOK'), ('', 'Nama Lapisan')] if col in df_clean.columns]
    numeric_cols = df_clean.columns.drop(info_cols)

    for col in numeric_cols:
        df_clean[col] = df_clean[col].str.replace(',', '', regex=False)
        df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce')

    df_clean.reset_index(drop=True, inplace=True)
    return df_clean

#=================================================================#
# FUNGSI DISPATCHER UTAMA YANG SUDAH DI-REFACTOR
#=================================================================#

def process_table(table_obj, onerow=False, category=False, paired=False, quantity_mode=False):
    """
    Fungsi utama untuk memproses sebuah objek tabel berdasarkan mode yang dipilih.
    Ini adalah pengganti dari 'extract'.
    """
    # Langkah 1: Ubah objek tabel mentah menjadi DataFrame
    # Tidak ada lagi pembacaan file di sini
    df_raw = table_to_dataframe(table_obj)

    # Langkah 2: Pilih fungsi pembersih yang sesuai berdasarkan flag
    if onerow:
        return process_cleaned_data(df_raw.T, paired=paired)
    elif category:
        # Panggil fungsi yang sudah diubah, kirim objek tabelnya
        return process_grid_table(table_obj)
    elif quantity_mode:
        return process_resource_table(df_raw)
    else:
        # Jika tidak ada mode, kembalikan DataFrame mentah
        if df_raw is not None and not df_raw.empty:
            df_raw.columns = df_raw.iloc[0]
            df_raw = df_raw.iloc[1:].reset_index(drop=True)
        return df_raw
    
# =============================================================
# BAGIAN 1: FUNGSI ANALISIS DOKUMEN
# (Kita pindahkan fungsi analisis yang sudah kita buat ke sini)
# =============================================================

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
    # CBM
    "JENIS DAN TAHAPAN EKSPLORASI CBM": "Tipe Eksplorasi CBM",
    # Lainnya
    "INFORMASI TITIK": "Tipe Informasi Titik"
}

def sanitize_sheet_name(name):
    """Membersihkan teks agar menjadi nama sheet Excel yang valid."""
    name = HEADING_PROCESSORS.get(name.strip(), name.strip())
    name = re.sub(r'[\\/*?:"<>|]', "", name)
    name = name.replace('Tipe ', '').strip()
    name = name[:31] 
    name = name.replace('/', '_').replace('\\', '_')
    return name[:31].strip()

def analyze_word_document(doc_object):
    """
    Menganalisis objek dokumen Word yang di-upload untuk memberikan ringkasan.
    """
    report_count = 0
    total_table_count = 0
    heading_analysis = {}
    active_heading = "Tanpa Heading Awal"
    heading_analysis[active_heading] = {'table_count': 0}

    for block in doc_object.element.body:
        if isinstance(block, CT_P):
            p = Paragraph(block, doc_object)
            if p.style.name.startswith('Heading 1'):
                active_heading = p.text.strip()
                if not active_heading:
                    active_heading = "(Heading Kosong)"
                
                if active_heading not in heading_analysis:
                    heading_analysis[active_heading] = {'table_count': 0}

                if 'I. DATA UMUM' in active_heading:
                    report_count += 1
        
        elif isinstance(block, CT_Tbl):
            total_table_count += 1
            if active_heading in heading_analysis:
                heading_analysis[active_heading]['table_count'] += 1
            else:
                heading_analysis[active_heading] = {'table_count': 1}
    
    # Kumpulkan hasil ke dalam satu dictionary
    results = {
        "report_count": report_count,
        "total_table_count": total_table_count,
        "headings": heading_analysis
    }
    # Hapus kunci 'Tanpa Heading Awal' dari dalam kamus 'headings'
    if 'Tanpa Heading Awal' in results["headings"]:
        del results["headings"]["Tanpa Heading Awal"]
    return results

# =============================================================
# BAGIAN 2: TAMPILAN ANTARMUKA STREAMLIT
# =============================================================