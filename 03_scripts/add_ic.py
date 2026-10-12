# -*- coding: utf-8 -*-
"""Acrescenta ao livro de gráficos as folhas de estatística descritiva com
intervalos de confiança a 95 %, sobre a base auditada (138 pares na revisão de 2026-10-08). Os testes de
Kruskal-Wallis e de Mann-Whitney calculados na folha «Dados planeadores» ficam
como registo da versão anterior: o desenho é emparelhado e os testes aplicáveis
são os da folha «Testes emparelhados». Preserva os gráficos nativos existentes."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.error_bar import ErrorBars
from openpyxl.chart.data_source import NumDataSource, NumRef
from openpyxl.utils import get_column_letter
from graficos_nativos import C1, C2, C3, HDR, HDR2, BR, titulo

PICK = 'Teste_pick_TM12_final_CORRIGIDO.xlsx'
AUD = 'Testes_de_tempo_e_estados_TM12_AUDITADO.xlsx'
OBST = 'Testes_Obstaculos_TM12_final_CORRIGIDO.xlsx'
PL = ['RRTConnect', 'RRTstar', 'PRMstar']

wb = openpyxl.load_workbook('G.xlsx')
TE = "'Testes estatísticos'"

def hdr(ws, r, c, txts, larg=None):
    for j, t in enumerate(txts):
        cl = ws.cell(r, c + j, t); cl.font = Font(bold=True, size=9.5); cl.fill = HDR2
        cl.alignment = Alignment(horizontal='center', wrap_text=True); cl.border = BR
    if larg:
        for j, w in enumerate(larg): ws.column_dimensions[get_column_letter(c + j)].width = w

def sub(ws, cel, txt):
    ws[cel] = txt; ws[cel].font = Font(bold=True, size=10.5, color='1F3864')

def nota(ws, cel, txt):
    ws[cel] = txt; ws[cel].font = Font(italic=True, size=9, color='595959')

# ===================================================================== 1
ws = wb.create_sheet('Estatística descritiva', 1)
ws.sheet_view.showGridLines = False
titulo(ws, 'A1', 'ESTATÍSTICA DESCRITIVA E INTERVALOS DE CONFIANÇA A 95 %')
nota(ws, 'A2', 'Tudo o que está nesta folha é fórmula. Altere um valor nos dados de origem e todos os '
               'resumos, margens de erro e intervalos se atualizam. O IC 95 % é o intervalo em que se '
               'situa a média verdadeira da população, com 95 % de confiança.')
ws.column_dimensions['A'].width = 34

COLS = ['Condição', 'n', 'Média', 'Mediana', 'Desvio-padrão', 'Mínimo', 'Máximo',
        'Q1', 'Q3', 'IQR', 'Erro-padrão', 't crítico', 'Margem de erro',
        'IC 95 % inferior', 'IC 95 % superior']
LARG = [34, 6, 10, 10, 13, 10, 10, 10, 10, 10, 12, 10, 13, 15, 15]

def bloco(ws, r0, tit, linhas, unidade='0.000'):
    """linhas = [(rótulo, intervalo)] com intervalo já qualificado com a folha."""
    sub(ws, f'A{r0}', tit)
    hdr(ws, r0 + 1, 1, COLS, LARG)
    for i, (rot, rng) in enumerate(linhas):
        r = r0 + 2 + i
        ws.cell(r, 1, rot).font = Font(size=10)
        f = {
            2: f'=COUNT({rng})',
            3: f'=AVERAGE({rng})',
            4: f'=MEDIAN({rng})',
            5: f'=_xlfn.STDEV.S({rng})',
            6: f'=MIN({rng})',
            7: f'=MAX({rng})',
            8: f'=QUARTILE({rng},1)',
            9: f'=QUARTILE({rng},3)',
            10: f'=I{r}-H{r}',
            11: f'=E{r}/SQRT(B{r})',
            12: f'=TINV(0.05,B{r}-1)',
            13: f'=L{r}*K{r}',
            14: f'=C{r}-M{r}',
            15: f'=C{r}+M{r}',
        }
        for c, v in f.items():
            cl = ws.cell(r, c, v); cl.font = Font(size=10); cl.border = BR
            cl.number_format = '0' if c == 2 else unidade
        ws.cell(r, 1).border = BR
    return r0 + 2 + len(linhas)


def faixas_da_folha(wb_):
    """Lê da folha «Testes estatísticos» os blocos contíguos de cada cenário (coluna A, a partir da
    linha 6). Evita escrever no código as linhas de cada cenário, que mudam com a base auditada."""
    t_ = wb_['Testes estatísticos']; blocos = []; r = 6
    while t_.cell(r, 1).value not in (None, ''):
        c = t_.cell(r, 1).value; a = r
        while t_.cell(r + 1, 1).value == c: r += 1
        blocos.append((c, a, r)); r += 1
    return blocos

CEN = faixas_da_folha(wb)
FIM = CEN[-1][2]   # última linha de dados na folha «Testes estatísticos»
N_PARES = FIM - 5

r = bloco(ws, 4, f'1. TEMPO DE CICLO — os {N_PARES} pares concluídos com sucesso pelas duas estratégias (Secção 5.4)', [
    ('Arquitetura temporal', f'{TE}!B6:B{FIM}'),
    ('Arquitetura por respostas', f'{TE}!C6:C{FIM}'),
    ('Diferença (respostas − temporal)', f'{TE}!D6:D{FIM}'),
])
nota(ws, f'A{r+1}', 'O IC 95 % da diferença não contém o zero, o que confirma que a redução do tempo de '
                    'ciclo não é atribuível ao acaso. É a leitura equivalente à do valor de p do teste de Wilcoxon.')

r += 3
r = bloco(ws, r, '2. TEMPO DE CICLO POR CENÁRIO — arquitetura temporal',
          [(c, f'{TE}!B{a}:B{b}') for c, a, b in CEN])
r += 2
r = bloco(ws, r, '3. TEMPO DE CICLO POR CENÁRIO — arquitetura por respostas',
          [(c, f'{TE}!C{a}:C{b}') for c, a, b in CEN])
r += 2
r = bloco(ws, r, '4. DIFERENÇA EMPARELHADA POR CENÁRIO (respostas − temporal)',
          [(c, f'{TE}!D{a}:D{b}') for c, a, b in CEN])
nota(ws, f'A{r+1}', 'Um IC 95 % que atravessa o zero significa que, nesse cenário, a diferença observada '
                    'é compatível com a ausência de efeito. É o que sucede nos cenários em que o teste de '
                    'Wilcoxon por cenário também não é significativo.')
LIN_DESC = r + 3

# ===================================================================== 2 dados auxiliares
def ler(f, folha='Dados Análise', nc=27):
    w = openpyxl.load_workbook(f, data_only=True)[folha]
    return [[w.cell(rr, cc).value for cc in range(1, nc + 1)]
            for rr in range(2, w.max_row + 1)]

lp = ler(PICK)
erro = [x[16] for x in lp[:200] if isinstance(x[16], (int, float))]
wo = openpyxl.load_workbook(OBST, data_only=True)['Dados Análise']
dist = [wo.cell(rr, 15).value for rr in range(2, 22)]
bloqueio = [float(v) for v in dist[:5]]
contorno = [float(v) for v in dist[5:]]
flange = [[wo.cell(rr, cc).value for cc in range(19, 23)] for rr in range(2, 19)]
corrig = [[wo.cell(rr, cc).value for cc in range(23, 27)] for rr in range(2, 19)]

aux = wb.create_sheet('Dados auxiliares')
aux.sheet_view.showGridLines = False
titulo(aux, 'A1', 'DADOS EM BRUTO PARA AS ESTATÍSTICAS DESCRITIVAS')
nota(aux, 'A2', 'Valores lidos dos ficheiros de ensaios. A folha «Estatística descritiva» calcula '
                'sobre estas colunas.')
hdr(aux, 3, 1, ['Erro euclidiano por movimento (mm)', 'Distância de contorno (mm)',
                'Distância de bloqueio (mm)', 'Erro euclidiano na flange (mm)',
                'Erro euclidiano com TCP corrigido (mm)'], [30, 22, 22, 26, 30])
for i, v in enumerate(erro):
    aux.cell(4 + i, 1, float(v)).number_format = '0.000'
for i, v in enumerate(contorno):
    aux.cell(4 + i, 2, float(v)).number_format = '0.000'
for i, v in enumerate(bloqueio):
    aux.cell(4 + i, 3, float(v)).number_format = '0.000'
for i, v in enumerate(flange):
    aux.cell(4 + i, 4, float(v[3])).number_format = '0.000'
for i, v in enumerate(corrig):
    aux.cell(4 + i, 5, float(v[3])).number_format = '0.000'

AX = "'Dados auxiliares'"
r = bloco(ws, LIN_DESC, '5. OUTRAS GRANDEZAS COM IC 95 % (Secções 5.5 e 5.6)', [
    ('Erro euclidiano por movimento (mm)', f'{AX}!A4:A{3+len(erro)}'),
    ('Erro euclidiano na flange (mm)', f'{AX}!D4:D{3+len(flange)}'),
    ('Erro euclidiano com TCP corrigido (mm)', f'{AX}!E4:E{3+len(corrig)}'),
    ('Distância de contorno (mm)', f'{AX}!B4:B{3+len(contorno)}'),
    ('Distância de bloqueio (mm)', f'{AX}!C4:C{3+len(bloqueio)}'),
])
nota(ws, f'A{r+1}', 'O IC 95 % da distância de bloqueio desce abaixo de zero, o que é impossível para uma '
                    'distância. É o sinal de que a média não descreve este grupo — daí a apresentação da '
                    'mediana e da amplitude no texto.')
GRAF = r + 3

# ===================================================================== 3 planeadores
def num(v): return isinstance(v, (int, float))
lp2t = [x for x in ler(PICK, nc=11) if x[2] in PL]
lo2t = [x for x in ler(OBST, nc=11) if x[2] in PL]
lp2 = [x for x in lp2t if all(num(x[i]) for i in (3, 5, 6))]
lo2 = [x for x in lo2t if all(num(x[i]) for i in (3, 5, 6))]
TOT = {('Sem obstáculos', p): sum(1 for x in lp2t if x[2] == p) for p in PL}
TOT.update({('Com obstáculos', p): sum(1 for x in lo2t if x[2] == p) for p in PL})
dp = wb.create_sheet('Dados planeadores')
dp.sheet_view.showGridLines = False
titulo(dp, 'A1', 'DADOS EM BRUTO DOS PLANEADORES E TESTES NÃO PARAMÉTRICOS')
hdr(dp, 2, 1, ['Campanha', 'Planeador', 'Tempo total (s)', 'Comprimento (m)',
               'Tempo de planeamento (s)'], [18, 14, 15, 15, 18])
nota(dp, 'G2', 'Só constam as execuções concluídas; as recusadas pelo controlador não têm métricas '
               'associadas e são contabilizadas à parte, na taxa de sucesso.')
linha = 3
faixas = {}
for camp, dados in [('Sem obstáculos', lp2), ('Com obstáculos', lo2)]:
    ini_c = linha
    for p in PL:
        ini = linha
        for x in [y for y in dados if y[2] == p]:
            dp.cell(linha, 1, camp); dp.cell(linha, 2, p)
            for k, idx in [(3, 5), (4, 6), (5, 3)]:
                cl = dp.cell(linha, k, float(x[idx])); cl.number_format = '0.0000'
            linha += 1
        faixas[(camp, p)] = (ini, linha - 1)
    faixas[(camp, 'TODOS')] = (ini_c, linha - 1)

MET = [('Tempo total (s)', 'C'), ('Comprimento (m)', 'D'), ('Tempo de planeamento (s)', 'E')]
DP = "'Dados planeadores'"
rr = 3
dp.column_dimensions['H'].width = 40
for cc in 'IJKLMN':
    dp.column_dimensions[cc].width = 13

def R_grupo(col, gi, gf, ti, tf):
    """Soma dos postos médios do grupo dentro do conjunto total."""
    g = f'${col}${gi}:${col}${gf}'; t = f'${col}${ti}:${col}${tf}'
    return (f'=SUMPRODUCT(COUNTIF({t},"<"&{g}))'
            f'+SUMPRODUCT((COUNTIF({t},"="&{g})+1)/2)')

for camp in ['Sem obstáculos', 'Com obstáculos']:
    ti, tf = faixas[(camp, 'TODOS')]
    sub(dp, f'H{rr}', f'KRUSKAL-WALLIS — campanha «{camp}»')
    rr += 1
    hdr(dp, rr, 8, ['Métrica', 'n RRTConnect', 'n RRTstar', 'n PRMstar',
                    'H', 'Graus de liberdade', 'p'], [40, 13, 13, 13, 11, 16, 12])
    rr += 1
    for nome, col in MET:
        base = rr
        dp.cell(rr, 8, nome).font = Font(size=10)
        somas = []
        for j, p in enumerate(PL):
            gi, gf = faixas[(camp, p)]
            dp.cell(rr, 9 + j, gf - gi + 1).number_format = '0'
            somas.append((R_grupo(col, gi, gf, ti, tf), gf - gi + 1))
        N = tf - ti + 1
        termo = '+'.join(f'(({f[1:]})^2/{n})' for f, n in somas)
        dp.cell(rr, 12, f'=12/({N}*({N}+1))*({termo})-3*({N}+1)').number_format = '0.000'
        dp.cell(rr, 13, 2).number_format = '0'
        dp.cell(rr, 14, f'=CHIDIST(L{base},M{base})').number_format = '0.0000'
        for c in range(8, 15): dp.cell(rr, c).border = BR
        rr += 1
    rr += 1

sub(dp, f'H{rr}', 'MANN-WHITNEY PAR A PAR COM CORREÇÃO DE BONFERRONI — campanha «Sem obstáculos»')
rr += 1
hdr(dp, rr, 8, ['Métrica e par comparado', 'n1', 'n2', 'U', 'z', 'p', 'p × 3 (Bonferroni)'],
    [40, 8, 8, 12, 11, 12, 18])
rr += 1
for nome, col in MET:
    for a in range(3):
        for b in range(a + 1, 3):
            ai, af = faixas[('Sem obstáculos', PL[a])]
            bi, bf = faixas[('Sem obstáculos', PL[b])]
            n1, n2 = af - ai + 1, bf - bi + 1
            A = f'${col}${ai}:${col}${af}'; B = f'${col}${bi}:${col}${bf}'
            R1 = (f'SUMPRODUCT(COUNTIF({A},"<"&{A}))+SUMPRODUCT(COUNTIF({B},"<"&{A}))'
                  f'+SUMPRODUCT((COUNTIF({A},"="&{A})+COUNTIF({B},"="&{A})+1)/2)')
            base = rr
            dp.cell(rr, 8, f'{nome} — {PL[a]} vs {PL[b]}').font = Font(size=10)
            dp.cell(rr, 9, n1).number_format = '0'
            dp.cell(rr, 10, n2).number_format = '0'
            dp.cell(rr, 11, f'={n1}*{n2}+{n1}*({n1}+1)/2-({R1})').number_format = '0.0'
            dp.cell(rr, 12, f'=(K{base}-{n1}*{n2}/2)/SQRT({n1}*{n2}*({n1}+{n2}+1)/12)').number_format = '0.000'
            dp.cell(rr, 13, f'=2*_xlfn.NORM.S.DIST(-ABS(L{base}),TRUE)').number_format = '0.00000'
            dp.cell(rr, 14, f'=MIN(1,M{base}*3)').number_format = '0.00000'
            for c in range(8, 15): dp.cell(rr, c).border = BR
            rr += 1
nota(dp, f'H{rr+1}', 'A correção de Bonferroni multiplica o valor de p pelo número de comparações (3). '
                     'Sem ela, comparar três grupos dois a dois inflaciona a probabilidade de encontrar '
                     'uma diferença que não existe.')

# descritiva por planeador
linhas_pl = []
for camp in ['Sem obstáculos', 'Com obstáculos']:
    for p in PL:
        gi, gf = faixas[(camp, p)]
        linhas_pl.append((f'{p} — {camp} (tempo total, s)', f'{DP}!C{gi}:C{gf}'))
r = bloco(ws, GRAF, '6. TEMPO TOTAL POR PLANEADOR, COM IC 95 % (Secção 5.7)', linhas_pl)

# ===================================================================== 4 gráfico com IC
gr = r + 2
sub(ws, f'A{gr}', '7. LEITURA GRÁFICA DOS INTERVALOS DE CONFIANÇA')
lin0 = 6  # 1ª linha do bloco 1
ch = BarChart(); ch.type = 'col'
ch.title = 'Tempo médio de ciclo por arquitetura, com intervalo de confiança a 95 %'
ch.x_axis.title = 'Arquitetura de sincronização'
ch.y_axis.title = 'Tempo de ciclo (s)'
ch.add_data(Reference(ws, min_col=3, min_row=lin0, max_row=lin0 + 1), titles_from_data=False)
ch.set_categories(Reference(ws, min_col=1, min_row=lin0, max_row=lin0 + 1))
ch.width, ch.height = 15, 9
ch.gapWidth = 90
s = ch.series[0]
s.graphicalProperties.solidFill = C1
s.errBars = ErrorBars(errDir='y', errValType='cust', errBarType='both',
                      plus=NumDataSource(NumRef(f=f"'Estatística descritiva'!$M${lin0}:$M${lin0+1}")),
                      minus=NumDataSource(NumRef(f=f"'Estatística descritiva'!$M${lin0}:$M${lin0+1}")))
ch.legend = None
ws.add_chart(ch, f'A{gr+2}')

ch2 = BarChart(); ch2.type = 'col'
ch2.title = 'Diferença emparelhada por cenário, com intervalo de confiança a 95 %'
ch2.x_axis.title = 'Cenário de ensaio'
ch2.y_axis.title = 'Diferença de tempo de ciclo (s)'
b4 = None
# bloco 4 começa depois dos blocos 1-3
b1 = 6; b2 = b1 + 3 + 4; 
# recalcular com precisão: reconstruir posições
ws_rows = {}
for rrx in range(1, ws.max_row + 1):
    v = ws.cell(rrx, 1).value
    if isinstance(v, str) and v.startswith('4. DIFERENÇA'): b4 = rrx + 2
ch2.add_data(Reference(ws, min_col=3, min_row=b4, max_row=b4 + 7), titles_from_data=False)
ch2.set_categories(Reference(ws, min_col=1, min_row=b4, max_row=b4 + 7))
ch2.width, ch2.height = 17, 9
ch2.gapWidth = 70
s2 = ch2.series[0]
s2.graphicalProperties.solidFill = C2
s2.errBars = ErrorBars(errDir='y', errValType='cust', errBarType='both',
                       plus=NumDataSource(NumRef(f=f"'Estatística descritiva'!$M${b4}:$M${b4+7}")),
                       minus=NumDataSource(NumRef(f=f"'Estatística descritiva'!$M${b4}:$M${b4+7}")))
ch2.legend = None
ws.add_chart(ch2, f'I{gr+2}')

wb.calculation.fullCalcOnLoad = True
wb.save('Graficos_Capitulo5_EDITAVEIS.xlsx')
print('gravado. folhas:', wb.sheetnames)
print('gráficos por folha:', {w.title: len(w._charts) for w in wb.worksheets})
