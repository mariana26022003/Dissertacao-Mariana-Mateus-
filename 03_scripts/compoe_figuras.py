# -*- coding: utf-8 -*-
"""Converte os SVG exportados do Excel em PNG de alta resolução e compõe as
figuras multi-painel do Capítulo 5, com as etiquetas (a), (b), (c)."""
import os, json, cairosvg
from PIL import Image, ImageDraw, ImageFont

SRC = 'graficos_excel'; DST = 'figuras_excel'
os.makedirs(DST, exist_ok=True)
ESCALA = 3.2
MARGEM, ESPACO = 14, 26
FONTE = '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'

idx = {}
for linha in open(os.path.join(SRC, 'indice.tsv'), encoding='utf-8'):
    folha, titulo, fich = linha.rstrip('\n').split('\t')
    idx[titulo] = fich

def png(titulo):
    saida = os.path.join(DST, '_' + idx[titulo].replace('.svg', '.png'))
    if not os.path.exists(saida):
        cairosvg.svg2png(url=os.path.join(SRC, idx[titulo]), write_to=saida, scale=ESCALA)
    return Image.open(saida).convert('RGB')

def acha(frag):
    for t in idx:
        if frag in t: return t
    return None

def compoe(nome, titulos, colunas=None, etiquetas=True):
    em_falta = [t for t in titulos if acha(t) is None]
    if em_falta:
        print(f'  !! {nome}: gráfico não encontrado no livro -> {em_falta}; figura não gerada')
        return None, None
    ims = [png(acha(t)) for t in titulos]
    n = len(ims)
    colunas = colunas or 1
    linhas = (n + colunas - 1) // colunas
    lc = max(i.width for i in ims)
    ims = [i.resize((lc, round(i.height * lc / i.width)), Image.LANCZOS) for i in ims]
    ac = max(i.height for i in ims)
    topo = 46 if (etiquetas and n > 1) else 0
    W = MARGEM * 2 + colunas * lc + (colunas - 1) * ESPACO
    H = MARGEM * 2 + linhas * (ac + topo) + (linhas - 1) * ESPACO
    fig = Image.new('RGB', (W, H), 'white')
    d = ImageDraw.Draw(fig)
    f = ImageFont.truetype(FONTE, 40)
    for k, im in enumerate(ims):
        cx, cy = k % colunas, k // colunas
        x = MARGEM + cx * (lc + ESPACO)
        y = MARGEM + cy * (ac + topo + ESPACO)
        if topo:
            d.text((x + 4, y + 2), f'({chr(97+k)})', fill=(26, 26, 26), font=f)
        fig.paste(im, (x, y + topo))
    cam = os.path.join(DST, nome)
    fig.save(cam, dpi=(300, 300))
    return cam, fig.size

FIGS = {
 'fig5_sinc_distribuicoes.png':   ['Tempos agregados nas duas arquiteturas'],
 'fig5_sinc_por_cenario.png':     ['Tempo médio de ciclo por cenário'],
 'fig5_sinc_cenario_detalhe.png': ['Espera improdutiva média por cenário',
                                   'Ciclos com falha por cenário',
                                   'Redução do tempo de ciclo por cenário'],
 'fig5_tempos_por_fase.png':      ['Tempo por etapa — cenário R01',
                                   'Tempo por etapa — cenário O07'],
 'fig5_sinc_sucesso.png':         ['Taxa de sucesso do ciclo ('],
 'fig5_pick_sucesso.png':         ['Taxa de sucesso por comando',
                                   'Sucesso por comando e por ciclo',
                                   'Tempo médio das etapas de movimento'],
 'fig5_tcp_correcao.png':         ['Erro de posicionamento por componente'],
 'fig5_erro_distribuicao.png':    ['Distribuição do erro euclidiano'],
 'fig5_obstaculos_distancias.png':['Distância do ponto de controlo ao obstáculo',
                                   'Distância média e mediana por tipo'],
 'fig5_planeadores_tempos.png':   ['Tempo total médio por planeador — ciclo completo',
                                   'Tempo total médio por planeador — movimento'],
 'fig5_planeadores_planeamento.png': ['Tempo médio de planeamento por planeador — ciclo completo',
                                      'Tempo médio de planeamento por planeador — movimento',
                                      'Número médio de pontos da trajetória'],
 'fig5_planeadores_comprimento.png': ['Comprimento cartesiano médio da trajetória — sem',
                                      'Comprimento cartesiano médio da trajetória — com'],
 'fig5_planeadores_sucesso.png':  ['Taxa de sucesso por planeador e por campanha'],
 'fig5_seguranca_robustez.png':   ['Ensaios de segurança e robustez'],
 'fig5_ic_diferenca.png':         ['Diferença emparelhada por cenário, com intervalo'],
}
# painéis com 3 elementos: 1 em cima, 2 em baixo fica desequilibrado -> 2 colunas
COLS = {}   # todos os painéis empilhados na vertical, à largura total

dims = {}
for nome, titulos in FIGS.items():
    cam, tam = compoe(nome, titulos, COLS.get(nome))
    if cam is None: continue
    dims[nome] = list(tam)
    print(f'{nome:38s} {len(titulos)} painel(eis)  {tam[0]}×{tam[1]}')

json.dump(dims, open(os.path.join(DST, 'figdims.json'), 'w'), indent=1)
