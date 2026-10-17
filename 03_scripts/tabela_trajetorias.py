# -*- coding: utf-8 -*-
"""Extrai, em formato longo, os pontos da trajetória planeada e o comprimento
do percurso do ponto de controlo, por planeador, cenário e repetição."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

lin = []

# ── Campanha sem obstáculos (ciclo completo) ─────────────────────────────
ws = openpyxl.load_workbook('Teste_pick_TM12_final_CORRIGIDO.xlsx', data_only=True)['Comparação Planners']
# blocos: Repetição 1 -> col 0..8 ; Repetição 2 -> col 17..25
# ordem das colunas: id, planner, t_plan, t_exec, t_total, comprimento, n_pontos, sucesso, obs
for base, rep in [(0, 1), (17, 2)]:
    ident = None
    for r in ws.iter_rows(min_row=3, values_only=True):
        v = r[base] if base < len(r) else None
        if isinstance(v, (int, float)) or (isinstance(v, str) and v.strip().isdigit()):
            ident = str(v).strip()
        pl = r[base + 1] if base + 1 < len(r) else None
        if pl not in ('RRTConnect', 'RRTstar', 'PRMstar') or ident is None: continue
        lin.append(['Sem obstáculos (ciclo completo)', f'Repetibilidade {ident}', '', rep, pl,
                    r[base+2], r[base+3], r[base+4], r[base+6], r[base+5],
                    r[base+7], r[base+8] if base+8 < len(r) else None,
                    'Teste_pick_TM12_final.xlsx / Comparação Planners'])

# ── Campanha com obstáculos (movimento isolado) ──────────────────────────
ws = openpyxl.load_workbook('Testes_Obstaculos_TM12_final_CORRIGIDO.xlsx', data_only=True)['Planners']
# ordem: id, planner, t_plan, t_exec, t_total, n_pontos, comprimento, sucesso, obs
for base, rep in [(0, 1), (17, 2)]:
    ident = None
    for r in ws.iter_rows(min_row=3, values_only=True):
        v = r[base] if base < len(r) else None
        if isinstance(v, str) and '_' in v and v[0] in 'OR': ident = v.strip()
        pl = r[base + 1] if base + 1 < len(r) else None
        if pl not in ('RRTConnect', 'RRTstar', 'PRMstar') or ident is None: continue
        cen, _, rp = ident.partition('_')
        lin.append(['Com obstáculos (movimento isolado)', cen, rp, rep, pl,
                    r[base+2], r[base+3], r[base+4], r[base+5], r[base+6],
                    r[base+7], r[base+8] if base+8 < len(r) else None,
                    'Testes_Obstaculos_TM12_final.xlsx / Planners'])

out = openpyxl.Workbook(); out.remove(out.active)
ws = out.create_sheet('Trajetórias planeadas')
ws.append(['Campanha', 'Cenário / obstáculo', 'Repetição do cenário', 'Execução',
           'Planeador', 'Tempo de planeamento (s)', 'Tempo de execução (s)',
           'Tempo total (s)', 'N.º de pontos da trajetória planeada',
           'Comprimento do percurso do TCP (m)', 'Sucesso', 'Observações', 'Ficheiro de origem'])
for l in lin: ws.append(l)
for c in ws[1]:
    c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='305496')
    c.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
ws.freeze_panes = 'A2'
for col, w in zip('ABCDEFGHIJKLM', [30, 18, 12, 11, 14, 14, 14, 13, 16, 17, 10, 22, 40]):
    ws.column_dimensions[col].width = w

n = out.create_sheet('Nota', 0)
for i, t in enumerate([
 ['Trajetórias planeadas e comprimento do percurso do ponto de controlo'], [],
 ['Distinção entre as duas grandezas:'],
 ['  trajetória planeada  - pontos gerados pelo MoveIt 2. A coluna "N.º de pontos da trajetória'],
 ['                         planeada" regista a dimensão dessa trajetória para cada execução.'],
 ['  trajetória executada - posições articulares registadas durante o movimento.'],
 ['                         ESTA GRANDEZA NÃO FOI REGISTADA. As posições articulares eram'],
 ['                         mantidas apenas em memória durante a execução e não foram'],
 ['                         guardadas em ficheiro. Não é reconstituída nem estimada.'], [],
 ['O comprimento do percurso do ponto de controlo foi calculado durante o ensaio, pela soma'],
 ['das distâncias euclidianas entre poses sucessivas do ponto de controlo.'], [],
 ['A identificação do cenário e da repetição está disponível para a campanha com obstáculos'],
 ['(identificador do tipo O03_1). Para a campanha sem obstáculos existe o número do ensaio de'],
 ['repetibilidade e a execução, não havendo identificador de obstáculo por não existirem'],
 ['obstáculos nessa campanha.'],
], 1): n.cell(i, 1, t[0] if t else '')
n.column_dimensions['A'].width = 92
n['A1'].font = Font(bold=True, size=12)

out.save('Trajetorias_Planeadas_e_TCP.xlsx')
import collections
print('linhas:', len(lin))
print(collections.Counter((l[0], l[4]) for l in lin))
