import pandas as pd

# Baca data dari file CSV
df = pd.read_csv('Guide Supplier KWM - DATABASE KWM.csv')
df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
cols_to_check = ['Nama Produk', 'Tempat order (opsi 1)', 'Tempat order (opsi 2)', 'Tempat order (opsi 3)', 'Satuan', 'Jumlah', 'Status', 'BSW']
df = df.dropna(subset=cols_to_check, how='all')
df['Nama Produk'] = df['Nama Produk'].str.lower()
df = df.sort_values(by='Tempat order (opsi 1)', key=lambda x: x.str[0])

# Tampilkan data ke dalam tabel

print(df)
# df.to_csv('Database Supplier.csv', index=False)
grouped_df = df.groupby('Tempat order (opsi 1)')

for name, group in grouped_df:
    print('Tempat order (opsi 1):', name)
    print(group)
    print()