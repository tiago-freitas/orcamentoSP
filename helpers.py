import sqlite3
import pickle
import os.path
import collections

import numpy as np
import pandas as pd

table_name_template = 'despesa%d'

IPCA = {2010: 1.4024986905,
        2011: 1.3151605332,
        2012: 1.2453234237,
        2013: 1.1773251765,
        2014: 1.104884986,
        2015: 1}

def sql2pickle(column_name, sum_what, DB):

    if os.path.exists('pickles/' + column_name + ' - ' + sum_what + '.p'):
        return False

    sql = '''SELECT despesa{0}."CODIGO {1}",
                    despesa{0}."NOME {1}",
                    SUM(despesa{0}."{2}")
             FROM despesa{0}
             GROUP BY despesa{0}."CODIGO {1}",
                      despesa{0}."NOME {1}"'''

    query_dict_years = { }

    for year in range(2010, 2016):
        with sqlite3.connect(DB) as conn:
            conn.execute(sql.format(year, table_name, sum_what))
            query_dict_years[year] = conn.fetchall()

    with open('pickles/' + column_name + ' - ' + sum_what + '.p', 'wb') as f:
        pickle.dump(query_dict_years, f)

    return True

def pickle2csv(pickle_name, sum_what):

    if os.path.exists('tables/' + pickle_name + ' - ' + sum_what + '.csv'):
        return False

    with open('pickles/' + pickle_name + ' - ' + sum_what + '.p', 'rb') as f:
        data_dict = pickle.load(f)

    data = collections.defaultdict(dict)
    for ano in data_dict:
        for row in data_dict[ano]:
            data[row[:2]][ano] = row[-1]

    rows = []
    grupos = sorted(list(data.keys()), key=lambda x: x[0] if x[0] else 0)

    for grupo in grupos:
        if      (grupo[0] or grupo[1] or
                (2010 in data[grupo] and data[grupo][2010]) or
                (2011 in data[grupo] and data[grupo][2011]) or
                (2012 in data[grupo] and data[grupo][2012]) or
                (2013 in data[grupo] and data[grupo][2013]) or
                (2014 in data[grupo] and data[grupo][2014]) or
                (2015 in data[grupo] and data[grupo][2015])):

            rows.append(str(grupo[0]) + ';' +
                        grupo[1].replace(';', ',') + ';' +
                        ';'.join(str(data[grupo][ano]).replace('.', ',') if ano in data[grupo] else '0'
                         for ano in range(2010, 2016)))

    with open('tables/' + pickle_name + ' - ' + sum_what + '.csv', 'w', encoding='windows-1252') as f:
        f.write('CODIGO ' + pickle_name + ';')
        f.write('NOME ' + pickle_name + ';')
        f.write('2010;2011;2012;2013;2014;2015\n')
        f.write('\n'.join(rows).encode('windows-1252').decode('windows-1252'))

    return True

def csv2xls(csv_name, sum_what, writer, abs_total=True,
            rel_dist=True, rel_increase=True, abs_corr_inflacao=True):
    table_name = 'tables/' + csv_name + ' - ' + sum_what + '.csv'
    df = pd.read_csv(table_name, sep=';', decimal=',',  encoding='windows-1252')
    df.rename(columns = {df.columns[0]: 'CODIGO'}, inplace = True)
    df.to_excel(writer, sheet_name=csv_name, index=False, float_format='%.2f')

    worksheet = writer.sheets[csv_name]
    worksheet.freeze_panes(1, 0)

    workbook = writer.book

    number = workbook.add_format({'num_format': '#,##0.00'})
    bottom_number = workbook.add_format({'num_format': '#,##0.00', 'bottom': True})
    percentage = workbook.add_format({'num_format': '0.00%'})
    bottom_perc = workbook.add_format({'num_format': '0.00%', 'bottom': True})
    bottom_border = workbook.add_format({'bottom': True})

    worksheet.set_column('A:A', len(df.columns[0]) + 3)
    worksheet.set_column('B:B', max(len(name) for name in df[df.columns[1]]) + 10)
    worksheet.set_column('C:H', 20, number)

    total_sum = df.sum()
    total_rows = len(df)

    if abs_total == True:
        # Total absoluto
        last_empty_row = total_rows + 1
        worksheet.write_row(last_empty_row, 0, ['', 'Total'], bottom_border)
        worksheet.write_row(last_empty_row, 2, total_sum[2:], bottom_number)

    if abs_corr_inflacao == True:
        # valores corrigidos pelo IPCA. Valores de dez 2015.
        last_empty_row += 2
        worksheet.write(last_empty_row, 0, 'Valores absolutos corrigidos pelo IPCA')
        last_empty_row += 1
        year = 2010
        total_sum_corr = []
        for col_index in range(8):
            if col_index >= 2:
                abs_corr_ipca = df[df.columns[col_index]] * IPCA[year]
                total_sum_corr.append(total_sum[col_index] * IPCA[year])
                worksheet.write_column(last_empty_row, col_index, abs_corr_ipca, number)
                year += 1
            else:
                worksheet.write_column(last_empty_row, col_index, df[df.columns[col_index]])
        last_empty_row += total_rows

        worksheet.write_row(last_empty_row, 0, ['', 'Total'], bottom_border)
        worksheet.write_row(last_empty_row, 2, total_sum_corr, bottom_number)



    if rel_dist == True:
        # relative distribution
        last_empty_row += 2
        worksheet.write(last_empty_row, 0, 'Distribuição relativa')
        last_empty_row += 1
        for col_index in range(8):
            if col_index >= 2:
                relative_dist = df[df.columns[col_index]] / total_sum[col_index]
                worksheet.write_column(last_empty_row, col_index, relative_dist, percentage)
            else:
                worksheet.write_column(last_empty_row, col_index, df[df.columns[col_index]])
        last_empty_row += total_rows

        worksheet.write_row(last_empty_row, 0, ['', 'Total'], bottom_border)
        worksheet.write_row(last_empty_row, 2, [1] * 6, bottom_perc)


    if rel_increase == True:
        # relative increase. year 1 = 100
        last_empty_row += 2
        worksheet.write(last_empty_row, 0, 'Crescimento relativo')
        last_empty_row += 1
        for row_index in range(total_rows):
            index = 0
            first_year = 0
            while not first_year and index < 6:
                first_year = df.iloc[row_index, 2:][index]
                index += 1
            if first_year:
                padding = index - 1
                increase_data = (df.iloc[row_index, 2 + padding:] / first_year) * 100
                worksheet.write_row(last_empty_row + row_index, 2 + padding, increase_data, number)
        for col_index in range(2):
            worksheet.write_column(last_empty_row, col_index, df[df.columns[col_index]])
        last_empty_row += total_rows
        first_year = total_sum[2:][0]

        worksheet.write_row(last_empty_row, 0, ['', 'Total'], bottom_border)
        worksheet.write_row(last_empty_row, 2, (total_sum[2:] / first_year) * 100, bottom_number)

