import pandas as pd

path1 = 'C:\\Users\\krystof\\OneDrive - MUNI\\Devin\\Fevin\\Fevin-2019-2023.xlsx'
path2 = 'C:\\Users\\krystof\\OneDrive - MUNI\\Devin\\Fenomor\\Fenomor_2019-2022.xlsx'

step2 = pd.read_excel(path1).groupby(['plot', 'month', 'year', 'species']).size().reset_index(name='noent')
step2[step2['noent'] > 1]
