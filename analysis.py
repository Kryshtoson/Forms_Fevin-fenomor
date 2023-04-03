import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import seaborn as sns
# import sys

path = 'C:\\Users\\krystof\\OneDrive - MUNI\\Devin\\Fevin\\Fevin-2019-2023.xlsx'
fevin = pd.read_excel(path, 'data').merge(pd.read_excel(path, 'traits'))
fevin2 = fevin.copy()
fevin2['history'] = 'All'
fevin_merged = pd.concat([fevin, fevin2], ignore_index=True)


fevin_merged = fevin_merged.groupby(by=['plot', 'history', 'year', 'month']).size().reset_index(name='nospe')
fevin_merged['date'] = pd.to_datetime(fevin_merged['year'].astype(str) + fevin_merged['month'].astype(str).str.zfill(2), format='%Y%B')
fevin_merged['history'] = pd.Categorical(fevin_merged['history'],
                                         categories=['All', 'Annuals', 'Short-lived perennials', 'Perennials'],
                                         ordered=True)
fevin_merged = fevin_merged.sort_values('date')

breaks = pd.DataFrame({'date': fevin_merged['date'].unique()})
breaks = breaks.loc[breaks['date'].dt.month.isin([3, 5, 7, 9, 12])]

fig, ax = plt.subplots(figsize=(10, 6))
sns.lineplot(x='date', y='nospe', data=fevin_merged, hue='history')
ax.legend(loc='upper left')
ax.xaxis.set_major_formatter(mdates.DateFormatter('%y\n%b'))
ax.set(xlabel='', ylabel='Number of species')
ax.set_xticks(breaks['date'])
plt.savefig('plot2.png', dpi=300, bbox_inches='tight', pad_inches=0.5)
