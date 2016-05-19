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

with open('columns2plural.txt') as f:
    columns2plural = {line.split(';')[0]:line.split(';')[1].strip() for line in f}

with open('sum2plural.txt') as f:
    sum2plural = {line.split(';')[0]:line.split(';')[1].strip() for line in f}

with open('tab_colors.txt') as f:
    tab_colors = [line.strip() for line in f]

def clean(elem):
        if isinstance(elem, str):
            return elem.strip().upper().replace(';', ',')
        elif isinstance(elem, int) or isinstance(elem, float):
            return str(elem)
        else:
            raise TypeError('Formato de {} não válido: {}'.format(elem, type(elem)))

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
            cursor = conn.cursor()
            cursor.execute(sql.format(year, column_name, sum_what))
            query_dict_years[year] = cursor.fetchall()
            cursor.close()

    with open('pickles/' + column_name + ' - ' + sum_what + '.p', 'wb') as f:
        pickle.dump(query_dict_years, f)

    return True

def pickle2csv(pickle_name, sum_what):

    if os.path.exists('tables/' + pickle_name + ' - ' + sum_what + '.csv'):
        return False

    with open('pickles/' + pickle_name + ' - ' + sum_what + '.p', 'rb') as f:
        data_dict = pickle.load(f)

    datas = collections.defaultdict(dict)
    for ano in data_dict:
        for row in data_dict[ano]:
            code = clean(row[0])
            name = clean(row[1])
            data = row[2]
            datas[(code, name)][ano] = data

    rows = []

    for grupo in datas:
        if      (grupo[0] or grupo[1] or
                (2010 in datas[grupo] and datas[grupo][2010]) or
                (2011 in datas[grupo] and datas[grupo][2011]) or
                (2012 in datas[grupo] and datas[grupo][2012]) or
                (2013 in datas[grupo] and datas[grupo][2013]) or
                (2014 in datas[grupo] and datas[grupo][2014]) or
                (2015 in datas[grupo] and datas[grupo][2015])):
            code = grupo[0] if grupo[0] else '-999'
            name = grupo[1] if grupo[1] else '-999'

            rows.append(code + ';' +
                        name + ';' +
                        ';'.join(str(datas[grupo][ano]).replace('.', ',') if ano in datas[grupo] else '0'
                        for ano in range(2010, 2016)))

    with open('tables/' + pickle_name + ' - ' + sum_what + '.csv', 'w') as f:
        f.write('CODIGO ' + pickle_name + ';')
        f.write('NOME ' + pickle_name + ';')
        f.write('2010;2011;2012;2013;2014;2015\n')
        f.write('\n'.join(rows))

    return True

def csv2xls(csv_name, sum_what, writer, sheet_number):

    def make_sheet(df, footer, title, subtitle, sheet_name):
        df.to_excel(writer,
                    header=False,
                    index=False,
                    sheet_name=sheet_name,
                    startrow=6,
                    float_format='%.2f')
        worksheet = writer.sheets[sheet_name]

        worksheet.hide_gridlines(2)
        worksheet.set_tab_color(color)

        worksheet.set_column('A:A', len(df.columns[0]) + 3, string_body)
        worksheet.set_column('B:B', max_length_col_names, string_body)
        worksheet.set_column('C:H', 20, number)

        worksheet.write_row(0, 0, [title[0]], string_title)
        worksheet.write_row(1, 0, [title[1]], string_title)
        worksheet.write_row(2, 0, ['Estado de São Paulo – 2010-2015'], string_title)
        worksheet.write_row(4, len_columns - 1, [subtitle], string_sub_header)

        for col in range(len_columns - 1):
            worksheet.write_row(5, col, [df.columns[col]], string_header)
        worksheet.write_row(5, len_columns - 1, [df.columns[len_columns - 1]], string_header_last)

        worksheet.write_row(total_rows + 6, 0, ['', 'Total'], string_footer)
        worksheet.write_row(total_rows + 6, 2, footer, number_footer)
        worksheet.write_rich_string(total_rows + 7, 0, string_footer, 'Fonte: ',
                                            string_body, source[:source_newline])
        worksheet.write_row(total_rows + 8, 0, [source[source_newline:]], string_body)

    table_path = 'tables/' + csv_name + ' - ' + sum_what + '.csv'

    df = pd.read_csv(table_path, sep=';', decimal=',')
    df.rename(columns = {df.columns[0]: 'CODIGO',
                         df.columns[1]: df.columns[1][5:]}, inplace = True)

    df.sort_values('CODIGO', inplace=True)

    # try:
    #     df.set_index('CODIGO', inplace=True, verify_integrity=True)
    # except ValueError as err:
    #     print(csv_name, sum_what)
    #     print(err)
    # df.set_index('CODIGO', inplace=True)

    workbook = writer.book

    ################################ Formats ##################################
    font_type = 'Courier New' # 'Roboto Mono'

    number = workbook.add_format({'num_format': '#,##0.00',
                                  'font_name': font_type,
                                  'font_size': 8})

    number_footer = workbook.add_format({'num_format': '#,##0.00',
                                         'font_name': font_type,
                                         'font_size': 8,
                                         'bold': True,
                                         'bottom': True})

    bottom_number = workbook.add_format({'num_format': '#,##0.00',
                                         'bottom': True,
                                         'font_name': font_type})

    percentage = workbook.add_format({'num_format': '0.00%',
                                      'font_name': font_type})

    bottom_perc = workbook.add_format({'num_format': '0.00%',
                                       'bottom': True,
                                       'font_name': font_type})

    bottom_border = workbook.add_format({'bottom': True,
                                         'font_name': font_type})

    string_header = workbook.add_format({'bottom': True,
                                         'right': True,
                                         'top': True,
                                         'font_name': font_type,
                                         'font_size': 10,
                                         'align': 'center'})

    string_header_last = workbook.add_format({'bottom': True,
                                         'top': True,
                                         'font_name': font_type,
                                         'font_size': 10,
                                         'align': 'center'})

    string_sub_header = workbook.add_format({'font_name': font_type,
                                             'font_size': 8,
                                             'align': 'right'})

    string_title = workbook.add_format({'font_name': font_type,
                                        'font_size': 10})

    string_body = workbook.add_format({'font_name': font_type,
                                       'font_size': 8})

    string_footer = workbook.add_format({'font_name': font_type,
                                         'font_size': 8,
                                         'bold': True,
                                         'bottom': True})

    ###########################################################################


    #################################### globals ##############################

    sheet_name = columns2plural[csv_name]
    sum_type = sum2plural[sum_what]
    total_rows = len(df)
    len_columns = len(df.columns)
    total_sum = df.sum()
    last_empty_row = 0
    color = tab_colors[sheet_number - 1]
    sub_num = ord('a')
    max_length_col_names = max(max(len(str(name))
                           for name in df[df.columns[1]]), 25)
    source = 'Elaborado pelo Observatório do Orçamento do Estado de ' \
             'São Paulo da Associação dos Especialistas em Políticas ' \
             'Públicas do Estado de São Paulo a partir dos dados da ' \
             'Secretaria da Fazenda do Estado de São Paulo'

    source_newline = source.find('a partir')
    ############################ Absolute values ##############################

    title = ['Tabela %d%c' % (sheet_number, chr(sub_num)),
             'Evolução da despesa em %s do orçamento por %s' % (sum_type, sheet_name)]
    subtitle = 'Valores em R$ 1,00'
    sheet = '%d%c - ' % (sheet_number, chr(sub_num)) + sheet_name
    footer = total_sum[2:]
    make_sheet(df, footer, title, subtitle, sheet)

    ######### valores corrigidos - deflator IPCA. Valores de dez 2015 #########

    sub_num += 1

    ipca = [IPCA[ano] for ano in range(2010, 2016)]
    inflacao_df = df.copy()
    inflacao_df.iloc[:, 2:] = (df.iloc[:, 2:] * ipca)

    footer = total_sum[2:] * ipca

    title = ['Tabela %d%c' % (sheet_number, chr(sub_num)),
             'Evolução da despesa em %s do orçamento ' \
             'por %s em valores de dezembro de 2015' % (sum_type, sheet_name)]
    subtitle = 'Valores em R$ 1,00'
    sheet = '%d%c - ' % (sheet_number, chr(sub_num)) + sheet_name
    make_sheet(inflacao_df, footer, title, subtitle, sheet)



    ################### relative increase. year 1 = 100 #######################

    sub_num += 1

    inc_df = df.copy()
    for i_row in range(total_rows):
        i_col = 2
        first_year = 0
        while not first_year and i_col < 7:
            first_year = df.iloc[i_row, i_col]
            i_col += 1
        i_col -= 1
        inc_df.iloc[i_row, 2:i_col] = np.nan
        inc_df.iloc[i_row, i_col:] = 100 * (df.iloc[i_row, i_col:] / first_year)


    first_year = total_sum[2:][0]
    footer = (total_sum[2:] / first_year) * 100
    title = ['Tabela %d%c' % (sheet_number, chr(sub_num)),
             'Crescimento relativo da despesa em %s do orçamento do ' \
             'por %s' % (sum_type, sheet_name)]
    subtitle = 'Ano 2010 = 100'
    sheet = '%d%c - ' % (sheet_number, chr(sub_num)) + sheet_name
    make_sheet(inc_df, footer, title, subtitle, sheet)

    ################### relative increase. year 1 = 100 #######################

    sub_num += 1

    inc_infla_df = inflacao_df.copy()
    for i_row in range(total_rows):
        i_col = 2
        first_year = 0
        while not first_year and i_col < 7:
            first_year = inflacao_df.iloc[i_row, i_col]
            i_col += 1
        i_col -= 1
        inc_infla_df.iloc[i_row, 2:i_col] = np.nan
        inc_infla_df.iloc[i_row, i_col:] = 100 * (inflacao_df.iloc[i_row, i_col:] / first_year)


    total_sum_inf = total_sum[2:] * ipca
    first_year = total_sum_inf[0]
    footer = (total_sum_inf / first_year) * 100
    title = ['Tabela %d%c' % (sheet_number, chr(sub_num)),
            'Crescimento relativo da despesa em %s do orçamento por %s em ' \
            'valores de dezembro de 2015' % (sum_type, sheet_name)]
    subtitle = 'Ano 2010 = 100'
    sheet = '%d%c - ' % (sheet_number, chr(sub_num)) + sheet_name
    make_sheet(inc_infla_df, footer, title, subtitle, sheet)


    ########################### relative distribution #########################
    sub_num += 1
    relative_df = df.copy()
    relative_df.iloc[:, 2:] = (df.iloc[:, 2:] / total_sum[2:]) * 100
    title = ['Tabela %d%c' % (sheet_number, chr(sub_num)),
             'Evolução da distribuição relativa da despesa em %s do orçamento ' \
             'por %s' % (sum_type, sheet_name)]
    subtitle = 'Valores em percentagem'
    sheet = '%d%c - ' % (sheet_number, chr(sub_num)) + sheet_name
    make_sheet(relative_df, [100] * 6, title, subtitle, sheet)
