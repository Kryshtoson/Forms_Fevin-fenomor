#!/usr/bin/env python
# -*- coding: utf8 -*-
import datetime
import os
import pandas as pd
import itertools
from fpdf import FPDF
import sys


def switcher(input):
    match input:
        case "1":
            output = "Skalni step, Svaty kopecek"
        case "2":
            output = "Pechavovy travnik, Svaty kopecek"
        case "3":
            output = "Sprasova step, Milovicka stran"
        case "4":
            output = "Skalni step, Devin"
        case "5":
            output = "Pechavovy travnik, Devin"
        case "6":
            output = "Sprasova step, Vysoky roh"
        case "7":
            output = "Skalni step, Kotel"
        case "8":
            output = "Pechavovy travnik, Stolova hora"
        case "9":
            output = "Sprasova step, Lisci vrch"
    return output

for x in ['Fevin', 'Fenomor']:
    print(x)

    if x == 'Fevin':
        path = 'C:\\Users\\krystof\\OneDrive - MUNI\\Devin\\Fevin\\Fevin-2019-2023.xlsx'
        letters_f = [f"F{i}" for i in range(1, 8)]
        letters_p = [f"P{i}" for i in range(1, 7)]
    elif x == 'Fenomor':
        path = 'C:\\Users\\krystof\\OneDrive - MUNI\\Devin\\Fenomor\\Fenomor_2019-2022.xlsx'
        letters_f = [f"F{i}" for i in range(1, 10)]
        letters_p = [f"-{i}" for i in range(1, 7)]
    else:
        print('specify source')
        sys.exit()
    plots = [''.join(pair) for pair in itertools.product(letters_f, letters_p)]

    # header data
    hlava = pd.read_excel(path, 'hlava')
    hlava['author'] = hlava['author'].str.replace(', ', ' & ')
    hlava['month_abb'] = [y[0:3] for y in hlava['month']]

    pdf = FPDF('P', 'mm', 'A5')

    # months = month_abbreviations = [calendar.month_abbr[i] for i in [1,3,5,7,9,12]]

    # species_row width: 45
    cw = 6.7  # cell width
    ch = 4.8  # cell height
    left_margin = 5 # remove left margin
    tot_rows = 33


    for i in plots:
        print(i)
        # reads prepared forms (new_forms.py)
        #df = pd.read_excel('out\\' + x + '.xlsx', sheet_name='F1P1', dtype=str)
        df = pd.read_excel('out\\' + x + '.xlsx', sheet_name=i, dtype=str)
        # read headers and subtract relevant months [df.columns]
        m = [a + b for a, b in zip([str(u)[0:3] for u in df.columns[[1, 4, 7]]], [str(u)[4:6] for u in df.columns[[1, 4, 7]]])]

        pdf.add_page() # new page for new plot
        pdf.set_auto_page_break(False, margin = 0.0)
        # --------------------------------------------
        # header (static)
        # --------------------------------------------

        # mesice v zahlavi
        pdf.set_left_margin(10)
        pdf.set_font('helvetica', 'B', 10)
        pdf.cell(45, ch, '', border=0)
        pdf.cell(cw * 3, ch, m[0], align='C', border=0)
        pdf.cell(cw * 3, ch, m[1], align='C', border=0)
        pdf.cell(cw * 3, ch, m[2], align='C', border=0)
        pdf.set_font('helvetica', '', 10)
        pdf.cell(cw * 3, ch, '', align='L', border=0)
        pdf.ln()

        # zahlavi vcetne dat z hlavicek
        pdf.set_font('helvetica', 'B', 8)
        filler = '' # jakmile E1, davaji se za hodnoty procenta
        for z in ['A: ', 'F: ', 'E1: ', 'E0: ', 'R: ', 'S: ']:
            if z == 'A: ':
                pdf.set_font('helvetica', 'B', 32)
                pdf.cell(45, ch-1, i, align='L')
                pdf.set_font('helvetica', '', 7)
            elif z == 'E0: ':
                pdf.set_font('helvetica', 'B', 15)
                pdf.cell(45, ch - 1, x, align='L')
                pdf.set_font('helvetica', '', 7)
            else:
                pdf.cell(45, ch-1, '')

            column_map = {'A: ': 'author', 'F: ': 'foto', 'E1: ': 'e1', 'E0: ': 'e0', 'R: ': 'r', 'S: ': 's'}

            for m_step in m:
                df_out = hlava[(hlava['month_abb'] == m_step[0:3]) & (hlava['year'] == int('20' + str(m_step[3:5]))) & (
                            hlava['plot'] == i)]
                if z == 'E1: ':
                    filler = '%'
                z_print = "{}{}{}".format(z, df_out[column_map[z]].iloc[0], filler)
                pdf.cell(cw * 3, ch - 1, z_print, align='L', border=0)

            pdf.cell(cw * 3, ch-1, z, align='L', border=0)
            pdf.ln()

        pdf.ln()
        pdf.set_font('helvetica', 'B', 9)
        pdf.cell(45, ch, 'species', align = 'C',border=0)
        pdf.set_font('helvetica', 'B', 10)
        pdf.cell(cw, ch, 'C', align='C', border=0)
        pdf.cell(cw, ch, 'H', align='C', border=0)
        pdf.cell(cw, ch, 'P', align='C', border=0)
        pdf.cell(cw, ch, 'C', align='C', border=0)
        pdf.cell(cw, ch, 'H', align='C', border=0)
        pdf.cell(cw, ch, 'P', align='C', border=0)
        pdf.cell(cw, ch, 'C', align='C', border=0)
        pdf.cell(cw, ch, 'H', align='C', border=0)
        pdf.cell(cw, ch, 'P', align='C', border=0)
        pdf.cell(cw, ch, 'C', align='C', border=0)
        pdf.cell(cw, ch, 'H', align='C', border=0)
        pdf.cell(cw, ch, 'P', align='C', border=0)
        pdf.ln()

        # --------------------------------------------
        # body (dynamic) from source
        # --------------------------------------------

        for r in range(0,len(df)):
            r_data=df.loc[r].to_numpy()
            for c in range(0, 11):
                if c == 0:
                    pdf.set_font("Helvetica", 'I', 8)
                    pdf.cell(45, ch, str(r+1) + '. ' + str(r_data[c]), border=0)
                elif c == 10:
                    for z in [0,1]:
                        pdf.cell(cw, ch, '', border=0)
                    pdf.cell(cw, ch, '', border=0)
                    pdf.ln()
                else:
                    if str(r_data[c]) == 'nan':
                        r_data[c]='.'
                    pdf.set_font("Helvetica", '', 9)
                    pdf.cell(cw, ch, str(r_data[c]), align='C', border=0)

        pdf.set_font("Helvetica", 'I', 8)
        for r in range(len(df), tot_rows-1):
            pdf.cell(45, ch, str(r + 1) + '. ', border=0)
            pdf.ln()
    #    pdf.set_font("Helvetica", 'I', 10)
    #    for r in range(0, tot_rows-len(df)):
    #        for c in range(0, 13):
    #            if c == 0:
    #                pdf.cell(45, ch, '', border=0)
    #            elif c == 12:
    #                pdf.cell(cw, ch, '', ln=True, align='C', border=0)
    #            else:
    #                pdf.cell(cw, ch, '', align='C', border=0)

        # ------------------------------------------------------------------------
        # fancy lines
        # ------------------------------------------------------------------------

        #cw = 6  # cell width
        #ch = 4.8  # cell height

        # horizontal ones

        cord_y = 42.2 + ch-1 #- 5
        cord_x1 = 11.1 #- 5
        cord_x2 = 10 + 45 + (cw*3)*4 #- 5

        for w in range(0, tot_rows):
            pdf.line(x1=cord_x1, x2=cord_x2,
                     y1=cord_y, y2=cord_y)
            cord_y = cord_y + ch

        # vertical ones

        cord_y1 = 10 + ch-1 #- 5
        cord_y2 = 171.8 + (4.8*6) + ch-1 - 4.8 #- 5
        cord_x = 10 + 45 #- 5
        for w in range(0, 4):
            pdf.line(x1=cord_x, x2=cord_x,
                     y1=cord_y1, y2=cord_y2)
            cord_x=cord_x + (cw*3)

        #pdf.rect(x=20, y=20, w=50, h=55, style='')

        if x == 'Fevin':
            pdf.set_xy(y=206.5-4.8, x=0)
            pdf.set_font("Helvetica", 'B', 7)
            pdf.cell(10 + 45 + (cw*3)*4, 1, 'Thank you for your help!', align='R', border=0)
        else:
            pdf.set_xy(y=206.5-4.8, x=10)
            pdf.set_font("Helvetica", 'B', 7)
            pdf.cell(50, 1, switcher(i[1]), align='L', border=0)
            pdf.set_xy(y=206.5 - 4.8, x=85)
            pdf.cell(50, 1, 'Thank you for your help!', align='R', border=0)
        pdf.ln()
    pdf.output('C:\\Users\\krystof\\OneDrive - MUNI\\Devin\\' + str(datetime.datetime.now().date()) + '_forms_' + x + '.pdf')

os.startfile('C:\\Users\\krystof\\OneDrive - MUNI\\Devin\\')