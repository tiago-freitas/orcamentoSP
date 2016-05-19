###############################################################################
# Conversor de CSV para SQLite das tabelas do Orçamento de Despesas de SP
#
# Contato dos Desenvolvedores:
# @Tiago Freitas <tiago4680@gmail.com>
# @Leandro Salvador <leandrosalvador@gmail.com>
#
# Versão: 2.0 (14.abr.2016)
#
# O que faz: transforma os arquivos CSV do Orçamento de Despesas do Governo
# do Estado de São Paulo em tabelas num banco de dados SQLite.
#
# Pré-requisitos: ter o Python 3 instalado no computador. Mais informações em:
# https://www.python.org/downloads/
#
# Modo de usar: (1) baixar os arquivos do sítio da Secretaria da Fazenda
# do Estado de São Paulo (SEFAZ), (2) colocar todos os arquivos .csv na mesma
# pasta que este arquivo Python (csv2sqlite.py), e (3) rodar o arquivo no
# prompt de comando (python csv2sqlite.py). A conversão demorará alguns minutos
# e o arquivo com o banco de dados (orcamento_sp.sqlite) poderá ficar com mais
# do que 5 GB, isso é normal.
#
# Link dos arquivos CSV no sítio da SEFAZ:
# http://www.fazenda.sp.gov.br/download/
# Seção 'Execução Orçamentária e Financeira - Atualização Diária'
#
# Para navegar pelas tabelas, além de conhecer um pouco a linguagem SQL, você
# pode utilizar a extensão do Firefox 'SQL Manager' ou o DbVisualizer:
# https://addons.mozilla.org/pt-BR/firefox/addon/sqlite-manager/
# http://dbvis.com/
###############################################################################

import sqlite3

outputDB = "orcamento_sp.sqlite"

def readCSV(inputFile):

    with open(inputFile, "r", encoding='windows-1252') as file:

        file.readline() # lê o cabeçalho e o ignora

        for l in file:

            line = l.strip('\n').split('","') # split retorna List
            line[0] = line[0][1:]
            line[-1] = line[-1][:-1]
            line = [elem.strip() for elem in line]
            for n in range(-4, 0, 1):
                if line[n]:
                    line[n] = line[n].replace(',', '.')

            yield line    # semelhante a lista.append(line), mas ad hoc

def updateDB(data, outputDB, outputTable):

    conn = sqlite3.connect(outputDB)
    c = conn.cursor()

    for line in data:
        c.execute('INSERT INTO ' + outputTable + ' VALUES (?'  + ',?' * 37 + ')' , line)

    conn.commit()
    conn.close()

def createDB(outputDB, outputTable):

    conn = sqlite3.connect(outputDB)
    c = conn.cursor()

    c.execute('''CREATE TABLE IF NOT EXISTS ''' + outputTable + ''' (
                 "ANO DE REFERENCIA" TEXT,
                 "CODIGO ORGAO" TEXT,
                 "NOME ORGAO" TEXT,
                 "CODIGO UNIDADE ORCAMENTARIA" TEXT,
                 "NOME UNIDADE ORCAMENTARIA" TEXT,
                 "CODIGO UNIDADE GESTORA" TEXT,
                 "NOME UNIDADE GESTORA" TEXT,
                 "CODIGO CATEGORIA DE DESPESA" TEXT,
                 "NOME CATEGORIA DE DESPESA" TEXT,
                 "CODIGO GRUPO DE DESPESA" TEXT,
                 "NOME GRUPO DE DESPESA" TEXT,
                 "CODIGO MODALIDADE" TEXT,
                 "NOME MODALIDADE" TEXT,
                 "CODIGO ELEMENTO DE DESPESA" TEXT,
                 "NOME ELEMENTO DE DESPESA" TEXT,
                 "CODIGO ITEM DE DESPESA" TEXT,
                 "NOME ITEM DE DESPESA" TEXT,
                 "CODIGO FUNCAO" TEXT,
                 "NOME FUNCAO" TEXT,
                 "CODIGO SUBFUNCAO" TEXT,
                 "NOME SUBFUNCAO" TEXT,
                 "CODIGO PROGRAMA" TEXT,
                 "NOME PROGRAMA" TEXT,
                 "CODIGO PROGRAMA DE TRABALHO" TEXT,
                 "NOME PROGRAMA DE TRABALHO" TEXT,
                 "CODIGO FONTE DE RECURSOS" TEXT,
                 "NOME FONTE DE RECURSOS" TEXT,
                 "NUMERO PROCESSO" TEXT,
                 "NUMERO NOTA DE EMPENHO" TEXT,
                 "CODIGO CREDOR" TEXT,
                 "NOME CREDOR" TEXT,
                 "CODIGO ACAO" TEXT,
                 "NOME ACAO" TEXT,
                 "TIPO LICITACAO" TEXT,
                 "VALOR EMPENHADO" REAL,
                 "VALOR LIQUIDADO" REAL,
                 "VALOR PAGO" REAL,
                 "VALOR PAGO DE ANOS ANTERIORES" REAL) ''')

    conn.commit()
    conn.close()

if __name__ == '__main__':
    for ano in range(2010, 2016):   # tabelas de 2010 a 2015, inclusive
        outputTable = 'despesa%d' % ano
        inputFile = 'despesa%d.csv' % ano
        createDB(outputDB, outputTable)
        data = readCSV(inputFile)
        updateDB(data, outputDB, outputTable)
