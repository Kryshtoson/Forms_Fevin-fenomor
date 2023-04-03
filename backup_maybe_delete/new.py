import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns

path = 'C:\\Users\\krystof\\OneDrive - MUNI\\Děvín\\Fevin\\Fevin-2019-2023.xlsx'
df = pd.read_excel(path, 'data').merge(pd.read_excel(path, 'traits'))
df = df.groupby(by=['plot', 'month', 'year']).size().reset_index(name = 'nospe')
df['date'] = pd.to_datetime(df['year'].astype(str) + df['month'].astype(str).str.zfill(2), format='%Y%B')

sns.set_style("ticks")
ax2 = sns.lineplot(x = 'date',
             y = 'nospe',
             data = df,
             hue = df['plot'] == 'F1P1',
             estimator = None,
             units = 'plot',
            legend = None,
             palette=['#AAAAAA', '#ff0000'])
#ax2.set(title = 'F1P1',  fontdict={'fontsize': 8, 'fontweight': 'medium'})
plt.savefig('relplot.png', dpi=300, bbox_inches='tight', pad_inches=0.5)
