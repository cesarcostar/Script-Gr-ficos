from scipy import stats
from colour import Color
from constantes_coluna import *
import string
import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt
plt.rcdefaults()

FILENAME = "consolidado(Atualizado).xlsx"
CAMINHO = "./AutoAva e Ava por Curso/"


def func(pct, allvals):
    # print(pct)
    if pct < 0.1:
        return ""
    else:
        return "{:.1f}%".format(pct)


def autolabel(rects, ax):
    for rect in rects:
        height = round(rect.get_height(), 2)
        # print(height)
        str_height = str(height) + '%'
        ax.annotate('{}'.format(str_height),
                    xy=(rect.get_x() + rect.get_width() / 2, height),
                    xytext=(0, 1),
                    textcoords="offset points",
                    ha='center', va='bottom')


def criarGraficoDeNotas(df, caminho, value, titulo, tipo=None):
    try:
        value = value.astype(float)
    except:
        pass
    if not value.isnull().all() and value.dtype in ['int', 'float']:
        red = Color("red")
        colors = list(red.range_to(Color("green"), 5))
        colors = [color.rgb for color in colors]

        labels_notas = dicionario_labels[tipo]
        width = dicionario_tamanho_pizza[tipo]
        intervalo_notas = dicionario_escalas[tipo]

        filtro_notas = pd.cut(value, intervalo_notas, labels=labels_notas)
        labels_notas = filtro_notas.value_counts(
            sort=False, normalize=True).index.tolist()
        lista_notas = filtro_notas.value_counts(
            sort=False, normalize=True).tolist()
        lista_notas_porcentagem = [i * 100 for i in lista_notas]
        y_pos = np.arange(len(labels_notas))
        fig, ax = plt.subplots()

        rects = ax.bar(y_pos, lista_notas_porcentagem, align='center',
                       alpha=0.7, color=colors, edgecolor='black')
        ax.grid(color='gray', linestyle='--',
                axis='y', linewidth=0.4, zorder=0)
        ax.set_axisbelow(True)
        ax.set_xticks(y_pos)
        ax.set_xticklabels(labels_notas)
        ax.margins(0.09)
        # ax.xticks(y_pos, labels_notas)
        # ax.ylabel('Porcentagem(%)')
        # ax.set_title(titulo)
        autolabel(rects, ax)
        fig.savefig(caminho + '/' + titulo + '.png', bbox_inches='tight')
        fig.clf()


def criarGraficoDeCount(df, caminho, value, titulo):
    try:
        value = value.astype(float)
    except:
        pass
    if not value.isnull().all() and value.dtype in ['int', 'float']:
        labels_notas = value.index
        lista_notas = value.value_counts(sort=False, normalize=True).tolist()
        lista_notas_porcentagem = [i * 100 for i in lista_notas]

        fig, ax = plt.subplots(figsize=(6, 3), subplot_kw=dict(aspect="equal"))
        wedges, texts, autotexts = ax.pie(value, autopct=lambda pct: func(
            pct, value), textprops=dict(color="black"))

        ax.legend(wedges, labels_notas,
                  loc="center left",
                  bbox_to_anchor=(1, 0, 0.5, 1))

        plt.setp(autotexts, size=5, weight="bold")

        ax.set_title(titulo)

        plt.savefig(caminho + '/' + titulo + '.pdf', bbox_inches='tight')
        plt.clf()


def analisarDados(lista_valores, coluna, sheet, curso, disciplina, pergunta=None):
    print("------------------------------------------")
    print(curso, disciplina)
    print(sheet)
    print(coluna)
    try:
        print(pergunta)
    except:
        pass
    print("Mínimo: ", np.nanmin(lista_valores))
    print("Máximo: ", np.nanmax(lista_valores))
    print("Média: ", np.average(lista_valores))
    print("Moda: ", stats.mode(lista_valores)[0][0])
    print("Mediana: ", np.nanmedian(lista_valores))
    print("Variância: ", np.nanvar(lista_valores))
    print("Desvio Padrão: ", np.nanstd(lista_valores))
    print("Coeficiente de variação: ",
          stats.variation(lista_valores), "%", "\n\n")


def criarDiretorio(caminho1="", caminho2=None, caminho3=None):
    try:
        os.mkdir(CAMINHO + caminho1)
    except FileExistsError:
        pass
    try:
        os.mkdir(CAMINHO + caminho1 + "/" + caminho2)
    except FileExistsError:
        pass
    except:
        return None
    try:
        os.mkdir(CAMINHO + caminho1 + "/" + caminho2 + "/" + caminho3)
    except FileExistsError:
        pass
    except:
        return None


criarDiretorio()
sheets = pd.ExcelFile(FILENAME).sheet_names

for sheet in sheets:
    if sheet == 'AutoAva e Ava Disciplina Aluno ':
        df = pd.read_excel(FILENAME, sheet_name=sheet,
                           header=[0, 1], encoding="UTF-8")
        df = df.replace(r'^\s*$', np.nan, regex=True)
        for item in df.keys():
            if item[1] == 'Curso':
                curso = item

        filtro = df.groupby(curso)
        for conjunto in filtro.groups.keys():
            colunasFiltradas = filtro.get_group(conjunto)
            criarDiretorio(colunasFiltradas[curso].values[0].title().rstrip(
                string.whitespace + '.'))
            lista_valores = {}
            lista_chaves = {}
            for coluna, value in colunasFiltradas.iteritems():
                try:
                    value = value.astype(float)
                except:
                    pass
                if not value.isnull().all() and value.dtype in ['int', 'float']:
                    # analisarDados(value.tolist(), coluna[0], sheet, colunasFiltradas[curso].values[0], conjunto, coluna[1])
                    if coluna[1].title() not in lista_valores.keys():
                        lista_valores[coluna[1].title()] = value.tolist()
                    else:
                        lista_valores[coluna[1].title()] += value.tolist()
            # print(lista_valores)

            for i in lista_valores:
                # print(colunasFiltradas[curso].values[0])
                # print(conjunto)

                # print(stats.describe(lista_valores[i]))
                caminho_topico = CAMINHO + colunasFiltradas[curso].values[0].title().rstrip(
                    string.whitespace + '.')
                criarDiretorio(colunasFiltradas[curso].values[0].title().rstrip(
                    string.whitespace + '.'))
                if i.title() in PERGUNTAS_COM_MARCAÇAO_1:
                    criarGraficoDeCount(
                        df, caminho_topico, df[i].loc[df[sheet.upper()]['Curso'] == conjunto].sum(), i.title())
                else:
                    try:
                        criarGraficoDeNotas(df, caminho_topico, pd.Series(
                            lista_valores[i]), i.title(), dicionario_perguntas[i])
                    except:
                        # analisarDados(
                        #     lista_valores[i], i, sheet, colunasFiltradas[curso].values[0], conjunto)
                        criarGraficoDeNotas(df, caminho_topico, pd.Series(
                            lista_valores[i]), i.title(), 'PERGUNTAS_0_a_10')
