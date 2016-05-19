import pandas as pd
from helpers import sql2pickle, pickle2csv, csv2xls

columns_names_path = 'columns_names.txt'
sum_what_path = 'sum_what.txt'

with open(columns_names_path) as f:
    columns_names = [filename.strip() for filename in f.readlines()]

with open(sum_what_path) as f:
    sum_whats = [filename.strip() for filename in f.readlines()]


for sum_what in sum_whats:
    xlsx_path = 'xlsx/despesaSP-' + sum_what.lower().replace(' ', '-') + '.xlsx'
    sheet_number = 1
    with pd.ExcelWriter(xlsx_path, engine='xlsxwriter',
                        options={'nan_inf_to_errors': True}) as writer:
        for column_name in columns_names:
            if column_name == 'CREDOR':
                continue
            sql2pickle(column_name, sum_what, '../db/orcamento_sp.sqlite')
            pickle2csv(column_name, sum_what)
            csv2xls(column_name, sum_what, writer, sheet_number)
            sheet_number += 1

        writer.book.set_properties({
                    'title':    'DespesaSP-2010a2016-' + sum_what.lower().replace(' ', '_'),
                    'subject':  'Execução Orçamentária e Financeira - Despesas entre 2010 a 2015',
                    'author':   'Tiago Barreiros de Freitas & Leandro Salvador',
                    'category': 'Orçamentária',
                    'keywords': 'Orçamento, Execução, Despesa, Estado de São Paulo',
                    'comments': 'Criado com Python, Pandas, Numpy e XlsxWriter'})

        writer.save()

