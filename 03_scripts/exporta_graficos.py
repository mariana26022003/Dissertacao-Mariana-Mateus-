# -*- coding: utf-8 -*-
"""Exporta cada gráfico nativo do livro Excel para SVG através do LibreOffice
(UNO) e converte para PNG a 300 DPI. Preserva a aparência dos gráficos tal como
estão no ficheiro da Mariana."""
import os, re, subprocess, time, sys, unicodedata
import uno
from com.sun.star.beans import PropertyValue

FICH = os.path.abspath('Graficos_Capitulo5_EDITAVEIS.xlsx')
DEST = os.path.abspath('graficos_excel')
os.makedirs(DEST, exist_ok=True)

def pv(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p

def liga():
    ctx_local = uno.getComponentContext()
    res = ctx_local.ServiceManager.createInstanceWithContext(
        'com.sun.star.bridge.UnoUrlResolver', ctx_local)
    for _ in range(40):
        try:
            return res.resolve('uno:socket,host=127.0.0.1,port=2002;urp;'
                               'StarOffice.ComponentContext')
        except Exception:
            time.sleep(0.5)
    raise RuntimeError('sem ligação ao LibreOffice')

srv = subprocess.Popen([
    'soffice', '--headless', '--norestore', '--nolockcheck', '--nodefault',
    '--accept=socket,host=127.0.0.1,port=2002;urp;'],
    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
try:
    ctx = liga()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
    doc = desktop.loadComponentFromURL(
        uno.systemPathToFileUrl(FICH), '_blank', 0,
        (pv('Hidden', True), pv('UpdateDocMode', 3)))
    doc.calculateAll()
    exp = smgr.createInstanceWithContext(
        'com.sun.star.drawing.GraphicExportFilter', ctx)

    def slug(s):
        s = unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode()
        s = re.sub(r'[^a-zA-Z0-9]+', '_', s).strip('_').lower()
        return s[:60]

    saidas = []
    folhas = doc.Sheets
    for i in range(folhas.Count):
        sh = folhas.getByIndex(i)
        dp = sh.DrawPage
        for j in range(dp.Count):
            shape = dp.getByIndex(j)
            try:
                modelo = shape.Model            # o modelo do gráfico
                titulo = modelo.Title.String
            except Exception:
                continue
            nome = f'{i:02d}_{j}_{slug(titulo)}'
            destino = os.path.join(DEST, nome + '.svg')
            exp.setSourceDocument(shape)
            exp.filter((pv('URL', uno.systemPathToFileUrl(destino)),
                        pv('MediaType', 'image/svg+xml'),
                        pv('FilterName', 'SVG')))
            saidas.append((sh.Name, titulo, destino))
            print(f'  {sh.Name:24s} | {titulo[:62]}')
    doc.close(False)
    with open(os.path.join(DEST, 'indice.tsv'), 'w', encoding='utf-8') as f:
        for s, t, d in saidas:
            f.write(f'{s}\t{t}\t{os.path.basename(d)}\n')
    print(f'\n{len(saidas)} gráficos exportados para SVG')
finally:
    srv.terminate()
