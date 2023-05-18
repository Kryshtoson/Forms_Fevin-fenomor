import pandas as pd
import itertools

writer = pd.ExcelWriter(r'C:\Users\krystof\OneDrive - MUNI\Devin\Fenomor\digitalization\Header2021_2022.xlsx', mode='w')
plots = [''.join(pair) for pair in itertools.product([f"{i}" for i in range(1, 10)], [f"-{i}" for i in range(1, 7)])]

for y in plots:
    print(y)
    df = pd.read_excel(r'C:\Users\krystof\OneDrive - MUNI\Devin\Fenomor\digitalization\Data2021_2022_raw.xlsx', y)
    df2 = df.iloc[:7]

    df3 = df2.iloc[:,[0,1,4,7,10,13,16,19,22]]
    new_col = ['id', '2021_05', '2021_07', '2021_09', '2021_12', '2022_3', '2022_5', '2022_07', '2022_09']
    for i in range(0,9):
        df3=df3.rename(columns={df3.columns[i]: new_col[i]})

    df4 = df3.melt(id_vars='id')
    df4[['year', 'month']] = df4['variable'].str.split('_', expand=True)
    df5 = df4.pivot(index='variable', columns='id', values='value')
    df5.rename(columns={'index': 'row_name'}, inplace=True)

    df5.iloc[:, 1:]
    df5.to_excel(writer, sheet_name=y, index=True)
writer.save()