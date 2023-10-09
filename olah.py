import csv
import pandas as pd
import openpyxl
import json
from fuzzywuzzy import fuzz

# Nama file CSV
nama_file = "mentah.csv"

# Headers yang ingin diambil
headers_yang_diinginkan = ['kdptimsmh', 'kdpstmsmh', 'nim_mahasiswa', 'nama_mahasiswa', 'telp_mahasiswa', 'email_mahasiswa', 'tahun_lulus', 'nik', 'npwp', 'f8', 'f504', 'f502', 'f505', 'f506', 'f5a1', 'f5a2', 'f1101', 'f1102', 'f5b', 'f5c', 'f5d', 'f18a', 'f18b', 'f18c', 'f18d', 'f1201', 'f1202', 'f14', 'f15', 'f1761', 'f1762', 'f1763', 'f1764', 'f1765', 'f1766', 'f1767', 'f1768', 'f1769', 'f1770', 'f1771', 'f1772', 'f1773', 'f1774', 'f21', 'f22', 'f23', 'f24', 'f25', 'f26', 'f27', 'f301', 'f302', 'f303', 'f401', 'f402', 'f403', 'f404', 'f405', 'f406', 'f407', 'f408', 'f409', 'f410', 'f411', 'f412', 'f413', 'f414', 'f415', 'f416', 'f6', 'f7', 'f7a', 'f1001', 'f1002', 'f1601', 'f1602', 'f1603', 'f1604', 'f1605', 'f1606', 'f1607', 'f1608', 'f1609', 'f1610', 'f1611', 'f1612', 'f1613', 'f1614']

# Inisialisasi dictionary untuk menyimpan data yang akan diambil
data_yang_diambil = {header: [] for header in headers_yang_diinginkan}

# Membuka file CSV dan membaca datanya
with open(nama_file, mode='r', newline='') as file_csv:
    reader = csv.DictReader(file_csv)

    # Loop melalui setiap baris data dalam file CSV
    for row in reader:
        if row['kdpstmsmh'] != '0':  # Hanya jika kdpstmsmh bukan 0
            for header in headers_yang_diinginkan:
                if header in row:
                    if row[header] == '0':
                        data_yang_diambil[header].append('')  # Ubah 0 menjadi string kosong
                    elif row[header] == '-':
                        data_yang_diambil[header].append('')  # Ubah - menjadi string kosong
                    elif row[header] == '061012':
                        data_yang_diambil[header].append('61012')  # Ubah - menjadi string kosong
                    else:
                        data_yang_diambil[header].append(row[header])
                else:
                    data_yang_diambil[header].append(None)  # Handle missing values with None

# Konversi data menjadi DataFrame pandas
df = pd.DataFrame(data_yang_diambil)

# Set nilai 'kdptimsmh' ke '061012' untuk semua baris
df['kdptimsmh'] = '061012'

# Simpan DataFrame sebagai file Excel (XLSX)
nama_file_xlsx = "output.xlsx"
df.to_excel(nama_file_xlsx, index=False, engine='openpyxl')

# Baca data JSON dari file 'provinsi.json'
with open('provinsi.json', 'r') as json_file:
    provinsi_data = json.load(json_file)
    
# Baca data JSON dari file 'kota.json'
with open('kota.json', 'r') as json_file:
    kota_data = json.load(json_file)

# Buat fungsi untuk mencari kesamaan terdekat dalam data JSON
def find_closest_match(data, name_key, code_key, name):
    best_match = None
    best_ratio = 0

    for item in data:
        ratio = fuzz.ratio(name.lower(), item[name_key].lower())
        if ratio > best_ratio:
            best_ratio = ratio
            best_match = item

    return best_match[code_key] if best_match else None

# Mengganti nilai 'f5a1' dan 'f5a2' dalam DataFrame dengan kode terdekat dari data JSON
df['f5a1'] = df['f5a1'].apply(lambda x: find_closest_match(provinsi_data, 'name', 'code', x))
df['f5a2'] = df['f5a2'].apply(lambda x: find_closest_match(kota_data, 'name', 'code', x))

# Menyimpan DataFrame yang telah diperbarui sebagai file Excel
df.to_excel(nama_file_xlsx, index=False, engine='openpyxl')

# Membaca kembali file Excel dan mengganti nama header
wb = openpyxl.load_workbook(nama_file_xlsx)
ws = wb.active

# Membuat dictionary untuk mengganti nama header
header_mapping = {
    'nim_mahasiswa': 'nimhsmsmh',
    'nama_mahasiswa': 'nmmhsmsmh',
    'telp_mahasiswa': 'telpomsmh',
    'email_mahasiswa': 'emailmsmh'
}

# Mengganti nama header di worksheet
for cell in ws[1]:  # Mengambil baris pertama (header)
    if cell.value in header_mapping:
        cell.value = header_mapping[cell.value]

# Menyimpan perubahan ke file Excel
wb.save(nama_file_xlsx)

