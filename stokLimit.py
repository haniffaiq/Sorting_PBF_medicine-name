import pandas as pd

df = pd.read_excel('Lapstok MUS.xls', header=6)
df = df.drop(df.columns[df.columns.str.contains('unnamed', case=False)], axis=1)
df = df.assign(Dijual='')
# df.to_excel('DataObatTidakDijual.xlsx', index=False)

print(df)

df.to_excel("DataObatTidakDijual MUS.xlsx")