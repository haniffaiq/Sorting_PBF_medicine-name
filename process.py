import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

# Fungsi untuk mencari produk yang cocok
def find_matching_product(df_supplier, product_name):
    matching_rows = df_supplier[df_supplier['Nama Barang'].str.contains(product_name, case=False, regex=False, na=False)]
    if not matching_rows.empty:
        tempat_order_list = matching_rows['Supplier Termurah'].dropna().str.strip().tolist()
        return tempat_order_list
    else:
        return None

# Membaca data supplier
def read_supplier_data(file_path):
    df_supplier = pd.read_excel(file_path)
    df_supplier = df_supplier.loc[:, ~df_supplier.columns.str.contains('^Unnamed')]
    cols_to_check = ['Nama Barang', 'Satuan', 'Supplier Termurah', 'HPP Min', 'HPP Max', 'HPP Average', 'Harga Jual Min', 'Harga Jual Max']
    df_supplier = df_supplier.dropna(subset=cols_to_check, how='all')
    df_supplier['Nama Barang'] = df_supplier['Nama Barang'].str.lower()
    df_supplier = df_supplier.sort_values(by='Supplier Termurah', key=lambda x: x.str[0])
    df_supplier = df_supplier.fillna("PBF Tidak Ada")
    return df_supplier

# Membaca data lapstok
def read_lapstok_data(file_path):
    df_lapstok = pd.read_excel(file_path, header=6)
    df_lapstok = df_lapstok.drop(df_lapstok.columns[df_lapstok.columns.str.contains('unnamed', case=False)], axis=1)
    cols_to_check_lapstok = ['PLU', 'Nama Produk', 'Satuan', 'Qty', 'Stok Min', 'Stok Max', 'Pesan']
    df_lapstok = df_lapstok.dropna(subset=cols_to_check_lapstok)
    df_lapstok['Nama Produk'] = df_lapstok['Nama Produk'].str.lower()
    return df_lapstok

# Membaca data obat tidak dijual
def read_obat_tidak_dijual_data(file_path):
    df_obat_tidak_dijual = pd.read_excel(file_path)
    df_obat_tidak_dijual = df_obat_tidak_dijual.fillna(0)
    df_obat_tidak_dijual = df_obat_tidak_dijual.loc[df_obat_tidak_dijual['Dijual'] == 1].copy()
    df_obat_tidak_dijual = df_obat_tidak_dijual.drop(columns=df_obat_tidak_dijual.columns[df_obat_tidak_dijual.columns.str.contains('Unnamed')])
    df_obat_tidak_dijual['Nama Produk'] = df_obat_tidak_dijual['Nama Produk'].str.lower()
    return df_obat_tidak_dijual

# Menggabungkan data dan melakukan pemrosesan
def process_data(supplier_file, lapstok_file, obat_tidak_dijual_file,  hasil_file, hasil2_file):
    # Baca data dari file supplier
    df_supplier = read_supplier_data(supplier_file)

    # Baca data dari file lapstok
    df_lapstok = read_lapstok_data(lapstok_file)

    # Baca data dari file obat tidak dijual
    df_obat_tidak_dijual = read_obat_tidak_dijual_data(obat_tidak_dijual_file)

    # Inisialisasi list untuk menampung hasil
    result_rows = []
    for index, row in df_lapstok.iterrows():
        nama_produk = row['Nama Produk']
        pesan = row['Pesan']

        # Cari tempat order menggunakan fungsi find_matching_product
        tempat_order = find_matching_product(df_supplier, nama_produk)

        # Jika ada tempat order yang cocok
        if tempat_order:
            jumlah_pesan = pesan

            # Tambahkan baris ke list hasil
            result_rows.append({
                'Nama Produk': nama_produk,
                'Tempat Order': tempat_order,
                'Jumlah Pesan': jumlah_pesan
            })

    # Buat DataFrame hasil
    df_tempat_order = pd.DataFrame(result_rows, columns=['Nama Produk', 'Tempat Order', 'Jumlah Pesan'])

    # Hapus produk yang tidak dijual
    df_tempat_order = df_tempat_order[~df_tempat_order['Nama Produk'].isin(df_obat_tidak_dijual['Nama Produk'])]
    df_tempat_order['Tempat Order'] = df_tempat_order['Tempat Order'].astype(str)


    df_grouped = df_tempat_order.groupby('Tempat Order').apply(lambda x: x.apply(lambda y: pd.Series(y[y.notnull()]), axis=1)).reset_index(drop=True)

    print(df_grouped)
    # Export DataFrame ke file Excel
    export_to_excel(df_tempat_order, hasil_file)
    export_to_excel(df_grouped, hasil2_file)

    # Cetak DataFrame obat tidak dijual
    print(df_obat_tidak_dijual)

# Fungsi untuk mengekspor DataFrame ke file Excel
def export_to_excel(df, file_name):
    wb = Workbook()
    ws = wb.active

    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)

    for column in ws.columns:
        max_length = 0
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

    wb.save(file_name)
    print("DataFrame exported to Excel successfully!")

# Jalankan proses data
process_data('Database PBF KWB.xlsx', 'Lapstok KWB.xls', 'DataObatTidakDijual KWB.xlsx', "KWB_1.xlsx", "KWB_2.xlsx")
# process_data('Database PBF KWM.xlsx', 'Lapstok KWM.xls', 'DataObatTidakDijual KWM.xlsx', "KWM_1.xlsx", "KWM_2.xlsx")


