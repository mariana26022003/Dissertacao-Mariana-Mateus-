# -*- coding: utf-8 -*-
"""Repara o livro de tempos e estados a partir do registo primário do VisionCore.

Duas correções:
  1. As linhas de CANCEL em falta nas fases dos ciclos que terminaram por
     cancelamento. Só R01E-20 as tinha; as outras cinco ocorrências mostravam
     apenas «-» nas etapas seguintes, o que não distingue «não enviado» de
     «não registado».
  2. A coluna de tempo de recuperação, que continha valores não rastreáveis ao
     registo. Passa a conter o intervalo entre a resposta de erro e a resposta
     terminal ao cancelamento, medido no registo do VisionCore.
"""
import openpyxl
from openpyxl.styles import Font

FICH = 'Testes_de_tempo_e_estados_TM12_CORRIGIDO.xlsx'
LOG = 'SentinelProcessEngine20260724.log'

# (folha, cenário, repetição, etapa em falha, resposta de erro,
#  duração do comando de cancelamento, recuperação erro->CANCEL_DONE, hora)
OCORR = [
    ('Rep_Fases', 'R01E',  4, 'FINISH_HOME',  'ERROR FINISH_FAILED', 2.390, 2.400, '09:10:24,414'),
    ('Rep_Fases', 'R01E', 20, 'MOVE_PLACE',   'ERROR MOVE_FAILED',  13.892, 14.414, '09:23:27,664'),
    ('Rep_Fases', 'R08E', 11, 'MOVE_PICK',    'ERROR MOVE_FAILED',   3.448,  3.969, '10:32:33,968'),
    ('Rep_Fases', 'R08E', 17, 'MOVE_PICK',    'ERROR MOVE_FAILED',   3.409,  3.929, '10:37:10,302'),
    ('Rep_Fases', 'R10E',  3, 'MOVE_PICK',    'ERROR MOVE_FAILED',   3.404,  3.924, '10:47:00,023'),
    ('Obs_Fases', 'O03E', 16, 'MOVE_PLACE',   'ERROR MOVE_FAILED',  13.828, 14.348, '14:06:21,518'),
]

wb = openpyxl.load_workbook(FICH)
NOTA = Font(italic=True, size=9, color='595959')

def blocos(ws):
    """Devolve {(cenário, repetição): [(linha, coluna_inicial)]} para os blocos de fases."""
    d = {}
    for c0 in (1, 14, 27, 40):
        for r in range(2, ws.max_row + 1):
            cen = ws.cell(r, c0).value
            rep = ws.cell(r, c0 + 1).value
            if cen is None: continue
            d.setdefault((str(cen).strip(), rep), []).append((r, c0))
    return d

alterados = 0
for folha, cen, rep, etapa, resp, dur_cancel, recup, hora in OCORR:
    ws = wb[folha]
    linhas = blocos(ws).get((cen, rep), [])
    # localizar a etapa em falha e a primeira linha seguinte por preencher
    idx = [i for i, (r, c0) in enumerate(linhas) if ws.cell(r, c0 + 2).value == etapa]
    if not idx: 
        print(f'  !! {cen}-{rep}: etapa {etapa} não encontrada'); continue
    i = idx[0]
    if i + 1 >= len(linhas):
        print(f'  !! {cen}-{rep}: sem linha livre a seguir a {etapa}'); continue
    r, c0 = linhas[i + 1]
    ja = ws.cell(r, c0 + 2).value
    if ja == 'CANCEL':
        print(f'  = {cen}-{rep}: linha de CANCEL já presente')
    else:
        ws.cell(r, c0 + 2, 'CANCEL')
        ws.cell(r, c0 + 3, 'EM_MOVIMENTO' if dur_cancel > 10 else 'PRONTO')
        ws.cell(r, c0 + 4, dur_cancel)
        ws.cell(r, c0 + 5, None)
        ws.cell(r, c0 + 6, 'OK CANCEL_DONE')
        for k in (7, 8, 9): ws.cell(r, c0 + k, 0)
        ws.cell(r, c0 + 10, 1)
        ws.cell(r, c0 + 11, f'Ramo de erro do fluxo por respostas. Reposto do registo '
                            f'SentinelProcessEngine20260724 ({hora}).')
        alterados += 1
        print(f'  + {cen}-{rep}: linha de CANCEL acrescentada ({dur_cancel} s)')
    # etapas seguintes: marcar explicitamente como não enviadas
    for r2, c02 in linhas[i + 2:]:
        if ws.cell(r2, c02 + 2).value not in (None, '—'):
            ws.cell(r2, c02 + 3, 'NÃO ENVIADO')

# ---- coluna de recuperação em «Dados Análise» ----
ws = wb['Dados Análise']
mapa = {('R01', 4): 2.400, ('R01', 20): 14.414, ('R08', 11): 3.969,
        ('R08', 17): 3.929, ('R10', 3): 3.924, ('O03', 16): 14.348}
ws.cell(1, 17, 'Estados recuperação (s) — registo VisionCore')
ws.cell(1, 16, 'Temporal recuperação (s) — anotação manual, não comparável')
n = 0
for r in range(2, ws.max_row + 1):
    k = (ws.cell(r, 2).value, ws.cell(r, 3).value)
    if k in mapa:
        ws.cell(r, 17, mapa[k]); n += 1
print(f'  recuperações repostas: {n}')

# nota metodológica
ws2 = wb['Pressupostos']
alvo = ws2.max_row + 2
ws2.cell(alvo, 1, 'Definição operacional do tempo de recuperação (revisão de 2026-08-18)').font = Font(bold=True)
ws2.cell(alvo + 1, 1,
    'Tempo de recuperação = intervalo entre a receção da primeira resposta de erro e a receção da '
    'resposta terminal ao cancelamento, medido no registo SentinelProcessEngine20260724 do VisionCore. '
    'A arquitetura temporal não possui ramo de erro: não lê a resposta, não deteta a falha e o fluxo '
    'não contém blocos de cancelamento. Nas condições incluídas na análise comparativa não existe nenhum '
    'CANCEL nesta arquitetura; o único CANCEL presente no ficheiro temporal_ciclos.csv (TEMP_156) pertence '
    'a uma sessão posterior, executada fora da sequência de ensaios. Os valores antes inscritos na '
    'coluna da arquitetura temporal são anotações manuais do tempo até à retoma pelo operador e não '
    'são comparáveis com os da arquitetura por respostas; não são reportados no Capítulo 5.').font = NOTA

wb.save(FICH)
print(f'\nlinhas de CANCEL acrescentadas: {alterados}')
