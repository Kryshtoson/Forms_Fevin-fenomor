import os
import pandas as pd
import numpy as np
import sys

# x = filename_out
for x in ['Fevin', 'Fenomor']:
    print(x)

    if x == 'Fevin':
        path = 'C:\\Users\\krystof\\OneDrive - MUNI\\Děvín\\Fevin\\Fevin-2019-2023.xlsx'
    elif x == 'Fenomor':
        path = 'C:\\Users\\krystof\\OneDrive - MUNI\\Děvín\\Fenomor\\Fenomor_2019-2022.xlsx'
    else:
        print('specify source')
        sys.exit()

    os.startfile('C:\\Users\\krystof\\OneDrive - MUNI\\Děvín\\Fenomor\\')

    # I call it fevin but it may be fenomor too in the script
    fevin = pd.read_excel(path, 'data')

    ## duplicated entries if any, sys.exit()
    dupl_seek = fevin.groupby(by=['species', 'plot', 'year', 'month']).size().reset_index(name='noent')
    dupl_seek = dupl_seek[dupl_seek['noent'] != 1]
    if len(dupl_seek) != 0:
        print('Duplicats in ' + x + '!\n')
        print(dupl_seek)
        sys.exit()

    fevin['height'] = [str(i) for i in fevin['height']]
    fevin['date'] = pd.to_datetime(fevin['year'].astype(str) + fevin['month'].astype(str), format='%Y%B')
    fevin['data'] = fevin['cover'].astype(str) + ' ' + fevin['height'].astype(str) + ' ' + fevin['phenology'].astype(str)

    fevin = fevin.sort_values('date')
    last_two = fevin[['month', 'year']].drop_duplicates().tail(3)

    fevin_last_two = pd.merge(fevin, last_two, how = 'inner')

    writer = pd.ExcelWriter('out\\' + x + '.xlsx', mode='w')

    for i in np.sort(fevin['plot'].unique().astype(str)):
        print(i)
        step = fevin_last_two[fevin_last_two['plot'] == i].copy().sort_values('species')
        step['cols'] = step['year'].astype(str) + '\n' + step['month'].astype(str)
        step = step.sort_values('date')
        original_order = step["cols"].unique()
        out = step[['species', 'cols', 'data']].pivot(columns = 'cols', values = 'data', index = 'species')
        out = out.reindex(columns = original_order)
        #out2 = pd.DataFrame({'species': out.index})
        #out2
        clname0 = out.columns.str.split('\n')[0][1][0:3] + '\n' + out.columns.str.split('\n')[0][0][2:4]
        clname1 = out.columns.str.split('\n')[1][1][0:3] + '\n' + out.columns.str.split('\n')[1][0][2:4]
        clname2 = out.columns.str.split('\n')[2][1][0:3] + '\n' + out.columns.str.split('\n')[2][0][2:4]
        clname0 = [clname0 + '\n' + str(i) for i in ['C', 'H', 'P']]
        clname1 = [clname1 + '\n' + str(i) for i in ['C', 'H', 'P']]
        clname2 = [clname2 + '\n' + str(i) for i in ['C', 'H', 'P']]
        stuff = pd.concat([pd.concat([out[out.columns[0]].str.split(' ', expand = True),
                                     out[out.columns[1]].str.split(' ', expand = True)], axis = 1),
                          out[out.columns[2]].str.split(' ', expand=True)], axis = 1)
        stuff.reset_index(inplace = True)
        stuff.columns = ['species'] + clname0 + clname1 + clname2

        stuff.to_excel(writer, sheet_name= i, index=False)

    writer.save()
