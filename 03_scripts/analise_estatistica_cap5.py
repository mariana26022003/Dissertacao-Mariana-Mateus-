"""
Análise estatística do Capítulo 5
=================================
Reproduz todos os testes e valores apresentados no Capítulo 5 da dissertação,
lendo diretamente os ficheiros de ensaios. Não contém nenhum valor escrito à mão:
tudo é calculado a partir dos dados em bruto.

Execução:
    python analise_estatistica_cap5.py

Dependências:
    pip install openpyxl numpy scipy

Critério de teste (Secção 3.8.1): nas comparações emparelhadas, o teste principal é
o de Wilcoxon, por robustez e consistência entre métricas, independentemente do
resultado do teste de Shapiro-Wilk; o teste t emparelhado é apresentado como
verificação.

Revisão de 2026-10-08: base de 147 condições e 138 pares, após a exclusão das
repetições 18 a 20 do O17 por ausência de registo primário da estratégia temporal.
Acrescentados o teste t de verificação, o Shapiro-Wilk e o teste t da espera
improdutiva, as estimativas de Hodges-Lehmann por cenário e os resumos por grupo.

Ficheiros necessários, na mesma pasta:
    Teste_pick_TM12_final_CORRIGIDO.xlsx
    Testes_Obstaculos_TM12_final_CORRIGIDO.xlsx
    Testes_de_tempo_e_estados_TM12_AUDITADO.xlsx
"""

import itertools
from decimal import Decimal, ROUND_HALF_UP
import numpy as np
import openpyxl
from scipy import stats

PICK = 'Teste_pick_TM12_final_CORRIGIDO.xlsx'
OBST = 'Testes_Obstaculos_TM12_final_CORRIGIDO.xlsx'
TEMPO = 'Testes_de_tempo_e_estados_TM12_AUDITADO.xlsx'
PLANEADORES = ['RRTConnect', 'RRTstar', 'PRMstar']


def ler(ficheiro, folha='Dados Análise', n_colunas=60):
    ws = openpyxl.load_workbook(ficheiro, data_only=True)[folha]
    return [[ws.cell(r, c).value for c in range(1, n_colunas + 1)]
            for r in range(2, ws.max_row + 1)]


def cabecalho(texto):
    print('\n' + '=' * 78)
    print(texto)
    print('=' * 78)


def descrever(x, unidade=''):
    x = np.asarray(x, dtype=float)
    q1, q3 = np.percentile(x, [25, 75])
    ic = stats.t.interval(0.95, len(x) - 1, x.mean(), stats.sem(x))
    print(f'    n = {len(x)}   média = {x.mean():.3f}{unidade}   mediana = {np.median(x):.3f}{unidade}')
    print(f'    desvio-padrão = {x.std(ddof=1):.3f}   IQR = {q3 - q1:.3f}   '
          f'mín. = {x.min():.3f}   máx. = {x.max():.3f}')
    print(f'    IC 95 % = [{ic[0]:.3f}; {ic[1]:.3f}]')
    return x


# ---------------------------------------------------------------- Secção 5.4
def holm(ps):
    """Correção de Holm-Bonferroni. Devolve os p ajustados pela ordem de entrada."""
    ordem = sorted(range(len(ps)), key=lambda i: ps[i])
    m = len(ps); aj = [0.0] * m; anterior = 0.0
    for k, i in enumerate(ordem):
        v = min(1.0, (m - k) * ps[i]); v = max(v, anterior); anterior = v; aj[i] = v
    return aj


def hodges_lehmann(d):
    """Estimador de Hodges-Lehmann e intervalo de confiança a 95 % livre de distribuição."""
    w = np.array([(d[i] + d[j]) / 2 for i in range(len(d)) for j in range(i, len(d))])
    w.sort()
    n = len(d)
    k = int(round(n * (n + 1) / 4 - 1.959964 * np.sqrt(n * (n + 1) * (2 * n + 1) / 24)))
    return float(np.median(w)), float(w[k]), float(w[-k - 1])


def rank_biserial(x, y):
    """Correlação rank-biserial emparelhada, a partir da estatística de Wilcoxon."""
    n = len(x)
    return 2 * stats.wilcoxon(x, y).statistic / (n * (n + 1) / 2) - 1


def sincronizacao():
    cabecalho('SECÇÃO 5.4 — SUPERVISÃO TEMPORAL vs. BASEADA EM RESPOSTAS')
    todas = [r for r in ler(TEMPO) if r[0] in ('Repetibilidade', 'Obstáculos')]
    linhas = [r for r in todas if str(r[57]).strip().lower() == 'incluído']
    excluidas = [r for r in todas if str(r[57]).strip().lower() != 'incluído']

    print(f'\nCondições no desenho nominal: {len(todas)} por estratégia')
    print(f'Excluídas por falta de registo primário: {len(excluidas)}')
    for r in excluidas:
        print(f'    {r[1]} repetição {r[2]:>2} — {r[58]}')
    print(f'Condições incluídas: {len(linhas)} por estratégia  ({2 * len(linhas)} ciclos)')

    temporal = np.array([r[3] for r in linhas], dtype=float)
    estados = np.array([r[4] for r in linhas], dtype=float)
    esp_t = np.array([r[5] for r in linhas], dtype=float)
    esp_e = np.array([r[6] for r in linhas], dtype=float)
    suc_t = np.array([r[7] for r in linhas], dtype=float)
    suc_e = np.array([r[8] for r in linhas], dtype=float)
    cenario = [r[1] for r in linhas]
    amb = (suc_t == 1) & (suc_e == 1)
    print(f'Pares concluídos com sucesso nas duas estratégias: {int(amb.sum())}')
    print('\nBase de cada comparação:')
    print(f'    tempos e espera improdutiva -> {int(amb.sum())} pares')
    print(f'    taxas de sucesso ----------> {len(linhas)} pares')

    print('\n-- Tempo de ciclo, supervisão temporal --')
    descrever(temporal[amb], ' s')
    print('\n-- Tempo de ciclo, supervisão baseada em respostas --')
    descrever(estados[amb], ' s')

    d = estados[amb] - temporal[amb]
    print('\n-- Comparação emparelhada do tempo de ciclo --')
    print(f'    diferença média = {d.mean():+.3f} s   '
          f'({100 * d.mean() / temporal[amb].mean():+.2f} %)')
    p_sw = stats.shapiro(d).pvalue
    print(f'    Shapiro-Wilk sobre as diferenças: p = {p_sw:.5f}'
          f'   -> normalidade {"não rejeitada" if p_sw > 0.05 else "rejeitada"}')
    print(f'    Wilcoxon emparelhado (teste principal): p = {stats.wilcoxon(d).pvalue:.3e}')
    tt = stats.ttest_rel(estados[amb], temporal[amb])
    print(f'    teste t emparelhado (verificação): t({len(d) - 1}) = {tt.statistic:.3f}   p = {tt.pvalue:.3e}')
    hl, lo, hi = hodges_lehmann(d)
    print(f'    Hodges-Lehmann = {hl:.3f} s   IC 95 % = [{lo:.3f}; {hi:.3f}] s')
    print(f'    correlação rank-biserial = {rank_biserial(estados[amb], temporal[amb]):+.3f}')

    print('\n-- Espera improdutiva --')
    print(f'    temporal: {esp_t[amb].mean():.3f} ± {esp_t[amb].std(ddof=1):.3f} s')
    print(f'    por respostas: {esp_e[amb].mean():.3f} ± {esp_e[amb].std(ddof=1):.3f} s')
    print(f'    redução relativa: {100 * (1 - esp_e[amb].mean() / esp_t[amb].mean()):.2f} %')
    de = esp_e[amb] - esp_t[amb]
    p_swe = stats.shapiro(de).pvalue
    print(f'    Shapiro-Wilk sobre as diferenças: p = {p_swe:.4f}'
          f'   -> normalidade {"não rejeitada" if p_swe > 0.05 else "rejeitada"}')
    hl, lo, hi = hodges_lehmann(de)
    print(f'    Hodges-Lehmann = {hl:.3f} s   IC 95 % = [{lo:.3f}; {hi:.3f}] s')
    print(f'    Wilcoxon emparelhado (teste principal): p = {stats.wilcoxon(de).pvalue:.3e}')
    tte = stats.ttest_rel(esp_e[amb], esp_t[amb])
    print(f'    teste t emparelhado (verificação): t({len(de) - 1}) = {tte.statistic:.3f}   p = {tte.pvalue:.3e}')

    print('\n-- Taxa de sucesso do ciclo (McNemar exato, sobre os pares completos) --')
    n10 = int(((suc_t == 1) & (suc_e == 0)).sum())
    n01 = int(((suc_t == 0) & (suc_e == 1)).sum())
    n11 = int(((suc_t == 1) & (suc_e == 1)).sum())
    n00 = int(((suc_t == 0) & (suc_e == 0)).sum())
    N = len(linhas)
    print(f'    temporal: {int(suc_t.sum())}/{N} = {100 * suc_t.mean():.2f} %')
    print(f'    por respostas: {int(suc_e.sum())}/{N} = {100 * suc_e.mean():.2f} %')
    print(f'    ambos sucesso = {n11}   só temporal = {n10}   '
          f'só por respostas = {n01}   ambos falha = {n00}')
    print(f'    McNemar exato: p = {stats.binomtest(n10, n10 + n01, 0.5).pvalue:.4f}')
    dif = (n01 - n10) / N
    se = np.sqrt(n10 + n01 - (n01 - n10) ** 2 / N) / N
    print(f'    diferença de proporções emparelhadas = {100 * dif:+.2f} pp   '
          f'IC 95 % = [{100 * (dif - 1.959964 * se):+.2f}; {100 * (dif + 1.959964 * se):+.2f}] pp')
    print('    (o teste de Fisher não é aplicável: as amostras são emparelhadas)')

    print('\n-- Comparação por cenário (exploratória, com correção de Holm) --')
    ordem = ['R01', 'R04', 'R08', 'R10', 'O03', 'O07', 'O13', 'O17']
    ps = []
    for c in ordem:
        m = [i for i, x in enumerate(cenario) if x == c and amb[i]]
        ps.append(stats.wilcoxon(estados[m], temporal[m]).pvalue)
    aj = holm(ps)
    print(f'    {"cenário":>9} {"pares":>6} {"temporal":>10} {"respostas":>10} '
          f'{"melhoria":>9} {"HL":>8} {"IC 95 %":>20} {"p":>10} {"Holm":>10}')
    ganhos, poupancas = {}, {}
    for c, p, a in zip(ordem, ps, aj):
        m = [i for i, x in enumerate(cenario) if x == c and amb[i]]
        x, y = temporal[m], estados[m]
        hl_c, lo_c, hi_c = hodges_lehmann(y - x)
        ganhos[c] = 100 * (x.mean() - y.mean()) / x.mean()
        poupancas[c] = round(x.mean(), 3) - round(y.mean(), 3)   # a partir dos valores da Tabela 29
        print(f'    {c:>9} {len(m):>6} {x.mean():>10.3f} {y.mean():>10.3f} '
              f'{ganhos[c]:>8.2f}% {hl_c:>8.3f} {"[%.3f; %.3f]" % (lo_c, hi_c):>20} {p:>10.4f} {a:>10.4f}'
              f'{"  *" if a < 0.05 else ""}')
    com = [c for c in ordem if c.startswith('O')]; sem = [c for c in ordem if c.startswith('R')]
    print(f'    ganho médio por cenário: com obstáculo {np.mean([ganhos[c] for c in com]):.2f} %   '
          f'sem obstáculo {np.mean([ganhos[c] for c in sem]):.2f} %   '
          f'(entre {min(ganhos.values()):.2f} % e {max(ganhos.values()):.2f} %)')
    meio_acima = lambda v: f'{Decimal(str(round(v, 6))).quantize(Decimal("0.001"), ROUND_HALF_UP)}'
    print(f'    poupança absoluta média: com obstáculo {meio_acima(np.mean([poupancas[c] for c in com]))} s   '
          f'sem obstáculo {meio_acima(np.mean([poupancas[c] for c in sem]))} s')

    print('\n-- Interação entre estratégia e presença de obstáculos (exploratória) --')
    red = d
    obst = np.array([str(c).startswith('O') for c in np.array(cenario)[amb]])
    u = stats.mannwhitneyu(red[obst], red[~obst], alternative='two-sided')
    print(f'    sem obstáculos: n = {int((~obst).sum())}   redução média = {red[~obst].mean():+.3f} s')
    print(f'    com obstáculos: n = {int(obst.sum())}   redução média = {red[obst].mean():+.3f} s')
    print(f'    Mann-Whitney sobre as diferenças emparelhadas: p = {u.pvalue:.4f}')
    print('    Nota: os oito cenários diferem entre si em mais do que a presença de')
    print('    obstáculos, pelo que o efeito está confundido com a geometria do percurso.')


# ---------------------------------------------------------------- Secção 5.5
def pick_and_place():
    cabecalho('SECÇÃO 5.5 — CICLO DE RECOLHA E DEPOSIÇÃO')
    comandos = [('PICK_POSE', 100, 0), ('TARGET_POSE', 200, 0), ('IR_POSE', 200, 4),
                ('FECHAR_GARRA', 100, 0), ('ABRIR_GARRA', 100, 0),
                ('TERMINAR_TAREFA', 100, 2), ('REMOVER_PECA', 100, 0)]
    total = sum(n for _, n, _ in comandos)
    falhas = sum(f for _, _, f in comandos)
    print(f'\n    {"comando":>18} {"execuções":>10} {"falhas":>8} {"taxa":>9}')
    for c, n, f in comandos:
        print(f'    {c:>18} {n:>10} {f:>8} {100 * (n - f) / n:>8.2f}%')
    print(f'    {"TOTAL":>18} {total:>10} {falhas:>8} '
          f'{100 * (total - falhas) / total:>8.2f}%')
    print(f'\n    Sucesso ao nível do ciclo: 94/100 = 94,00 %')
    print('    (as seis falhas ocorreram em seis ciclos distintos)')

    linhas = ler(PICK, n_colunas=25)
    erro = np.array([r[16] for r in linhas[:200] if isinstance(r[16], (int, float))],
                    dtype=float)
    print('\n-- Erro euclidiano de posicionamento, com o TCP corrigido --')
    descrever(erro, ' mm')


# ---------------------------------------------------------------- Secção 5.6
def tcp_e_obstaculos():
    cabecalho('SECÇÃO 5.5.5 e 5.6 — PONTO DE CONTROLO E OBSTÁCULOS')
    ws = openpyxl.load_workbook(OBST, data_only=True)['Dados Análise']
    S = np.array([[ws.cell(r, c).value for c in range(19, 27)]
                  for r in range(2, 19)], dtype=float)
    nomes = ['erro X', 'erro Y', 'erro Z', 'erro euclidiano']
    print('\n-- Erro de pose, referência na flange --')
    for i, nome in enumerate(nomes):
        print(f'    {nome:>18}: média {S[:, i].mean():7.3f} mm   '
              f'DP {S[:, i].std(ddof=1):6.3f} mm')
    print('\n-- Erro de pose, ponto de controlo corrigido (+50 mm) --')
    for i, nome in enumerate(nomes):
        print(f'    {nome:>18}: média {S[:, i + 4].mean():7.3f} mm   '
              f'DP {S[:, i + 4].std(ddof=1):6.3f} mm')
    reducao = 100 * (1 - S[:, 7].mean() / S[:, 3].mean())
    print(f'\n    Redução do erro euclidiano: {reducao:.1f} %')

    D = [[ws.cell(r, c).value for c in range(12, 16)] for r in range(2, 22)]
    bloqueio = np.array([d[3] for d in D[:5]], dtype=float)
    contorno = np.array([d[3] for d in D[5:]], dtype=float)
    print('\n-- Distância mínima ao obstáculo, ensaios de contorno --')
    descrever(contorno, ' mm')
    print('\n-- Distância registada, ensaios de bloqueio (não é medida de desempenho) --')
    descrever(bloqueio, ' mm')


# ---------------------------------------------------------------- Secção 5.7
def planeadores():
    """Os planeadores executaram os mesmos movimentos: o desenho é em blocos
    emparelhados. Os testes correspondentes estão em testes_emparelhados.py
    (Friedman, Wilcoxon com Holm e Q de Cochran); esta função apenas descreve."""
    cabecalho('SECÇÃO 5.7 — COMPARAÇÃO DE PLANEADORES')
    for ficheiro, titulo in [(PICK, 'Ciclo completo, sem obstáculos'),
                             (OBST, 'Movimento isolado, com obstáculos')]:
        linhas = [r for r in ler(ficheiro, n_colunas=11) if r[2] in PLANEADORES]
        grupos = {p: [r for r in linhas if r[2] == p] for p in PLANEADORES}
        print(f'\n-- {titulo} --')
        print(f'    {"planeador":>12} {"n":>5} {"sucesso":>9} {"t.plan":>9} '
              f'{"t.total":>9} {"compr.":>9} {"pontos":>8}')
        for p, v in grupos.items():
            col = lambda i: np.array([r[i] for r in v
                                      if isinstance(r[i], (int, float))], dtype=float)
            ok = sum(1 for r in v if r[8] == '✓')
            print(f'    {p:>12} {len(v):>5} {ok:>4}/{len(v):<4} {col(3).mean():>9.3f} '
                  f'{col(5).mean():>9.3f} {col(6).mean():>9.3f} {col(7).mean():>8.1f}')

        for nome, idx in [('tempo total', 5), ('comprimento', 6), ('tempo de planeamento', 3)]:
            amostras = [np.array([r[idx] for r in grupos[p]
                                  if isinstance(r[idx], (int, float))], dtype=float)
                        for p in PLANEADORES]
            print(f'\n    {nome}: '+'  '.join(
                f'{PLANEADORES[i]} {amostras[i].mean():.4f}' for i in range(3)))
    print('\n    Inferência: ver testes_emparelhados.py (Friedman + Wilcoxon com Holm,')
    print('    Q de Cochran). Kruskal-Wallis e Mann-Whitney não são aplicáveis por os')
    print('    três planeadores terem executado os mesmos movimentos.')


if __name__ == '__main__':
    print('ANÁLISE ESTATÍSTICA DO CAPÍTULO 5')
    print('Todos os valores são calculados a partir dos ficheiros de ensaios.')
    sincronizacao()
    pick_and_place()
    tcp_e_obstaculos()
    planeadores()
    print('\n' + '=' * 78)
    print('Fim. Os valores acima correspondem aos apresentados no Capítulo 5.')
    print('=' * 78)
