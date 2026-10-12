# -*- coding: utf-8 -*-
"""Testes emparelhados para a comparação de planeadores e para a taxa de sucesso
das duas arquiteturas de supervisão. Substituem os testes para grupos
independentes usados na versão anterior do Capítulo 5."""
import itertools, numpy as np, openpyxl
from scipy import stats

PICK='Teste_pick_TM12_final_CORRIGIDO.xlsx'
OBST='Testes_Obstaculos_TM12_final_CORRIGIDO.xlsx'
TEMPO='Testes_de_tempo_e_estados_TM12_AUDITADO.xlsx'
PL=['RRTConnect','RRTstar','PRMstar']

def ler(f, nc=11, folha='Dados Análise'):
    ws=openpyxl.load_workbook(f,data_only=True)[folha]
    return [[ws.cell(r,c).value for c in range(1,nc+1)] for r in range(2,ws.max_row+1)]

def holm(ps):
    """Correção de Holm-Bonferroni. Devolve os p ajustados na ordem de entrada."""
    ordem=sorted(range(len(ps)), key=lambda i: ps[i])
    m=len(ps); aj=[0.0]*m; anterior=0.0
    for k,i in enumerate(ordem):
        v=min(1.0,(m-k)*ps[i]); v=max(v,anterior); anterior=v; aj[i]=v
    return aj

def campanha(fich, nome, chave):
    linhas=[r for r in ler(fich) if r[2] in PL]
    dados={p:{} for p in PL}
    for r in linhas:
        k=chave(r)
        if all(isinstance(r[i],(int,float)) for i in (3,5,6)):
            dados[r[2]][k]={'t_plan':r[3],'t_total':r[5],'compr':r[6]}
        dados[r[2]].setdefault(k,None)
        if dados[r[2]][k] is None: dados[r[2]][k]={'falha':True}
    todas=sorted(set().union(*[set(d) for d in dados.values()]))
    completos=[k for k in todas if all(dados[p].get(k) and 'falha' not in dados[p][k] for p in PL)]
    print(f'\n{"="*78}\n{nome}\n{"="*78}')
    print(f'  blocos totais: {len(todas)}   blocos completos (os três planeadores): {len(completos)}')
    for met,rot in [('t_total','Tempo total (s)'),('compr','Comprimento (m)'),('t_plan','Tempo de planeamento (s)')]:
        M=np.array([[dados[p][k][met] for p in PL] for k in completos], float)
        F=stats.friedmanchisquare(*[M[:,j] for j in range(3)])
        print(f'\n  -- {rot} --')
        print(f'     médias: '+'  '.join(f'{p} {M[:,j].mean():.4f}' for j,p in enumerate(PL)))
        W=F.statistic/(len(M)*(3-1))
        print(f'     Friedman: chi2 = {F.statistic:.3f}, gl = 2, p = {F.pvalue:.4g}')
        print(f'     Kendall W (dimensão do efeito) = {W:.3f}')
        if F.pvalue<0.05:
            pares=list(itertools.combinations(range(3),2))
            ps=[stats.wilcoxon(M[:,a],M[:,b]).pvalue for a,b in pares]
            for (a,b),p_,pa in zip(pares,ps,holm(ps)):
                marca='  <- significativo' if pa<0.05 else ''
                print(f'       {PL[a]:>11} vs {PL[b]:<11} Wilcoxon p = {p_:.4g}   Holm = {pa:.4g}{marca}')
    # Q de Cochran sobre o sucesso
    S=[]
    for k in todas:
        linha=[]
        for p in PL:
            d=dados[p].get(k)
            linha.append(0 if (d is None or 'falha' in d) else 1)
        S.append(linha)
    S=np.array(S)
    S=S[S.sum(1)%3!=0] if False else S
    k_=3; N=len(S); Gj=S.sum(0); Li=S.sum(1); N_=S.sum()
    Q=(k_-1)*(k_*(Gj**2).sum()-N_**2)/(k_*N_-(Li**2).sum()) if (k_*N_-(Li**2).sum())!=0 else float('nan')
    pQ=stats.chi2.sf(Q,k_-1)
    print(f'\n  -- Taxa de sucesso (blocos = {N}) --')
    print('     sucessos: '+'  '.join(f'{p} {Gj[j]}/{N}' for j,p in enumerate(PL)))
    print(f'     Q de Cochran: Q = {Q:.3f}, gl = 2, p = {pQ:.4g}')
    if pQ<0.05:
        pares=list(itertools.combinations(range(3),2)); ps=[]
        for a,b in pares:
            n01=int(((S[:,a]==0)&(S[:,b]==1)).sum()); n10=int(((S[:,a]==1)&(S[:,b]==0)).sum())
            ps.append(stats.binomtest(n10,n01+n10,0.5).pvalue if n01+n10 else 1.0)
        for (a,b),p_,pa in zip(pares,ps,holm(ps)):
            print(f'       {PL[a]:>11} vs {PL[b]:<11} McNemar exato p = {p_:.4g}   Holm = {pa:.4g}')

campanha(PICK, 'CAMPANHA 1 — ciclo completo, sem obstáculos', lambda r: (r[0], r[1]))
campanha(OBST, 'CAMPANHA 2 — movimento isolado, com obstáculos', lambda r: (r[1], r[0]))

# ---------------- McNemar: temporal vs. por respostas ----------------
print(f'\n{"="*78}\nTAXA DE SUCESSO DO CICLO — temporal vs. por respostas (emparelhado)\n{"="*78}')
ws=openpyxl.load_workbook(TEMPO,data_only=True)['Dados Análise']
L=[[ws.cell(r,c).value for c in range(1,61)] for r in range(2,ws.max_row+1)]
L=[r for r in L if r[0] in ('Repetibilidade','Obstáculos')]
nominal=len(L)
L=[r for r in L if str(r[57]).strip().lower()=='incluído']
print(f'  condições no desenho nominal: {nominal} | excluídas por falta de registo primário: {nominal-len(L)}')
a=np.array([r[7] for r in L],int); b=np.array([r[8] for r in L],int)
n11=int(((a==1)&(b==1)).sum()); n10=int(((a==1)&(b==0)).sum())
n01=int(((a==0)&(b==1)).sum()); n00=int(((a==0)&(b==0)).sum())
print(f'  pares = {len(a)}')
print(f'  ambos sucesso = {n11}   só temporal = {n10}   só por respostas = {n01}   ambos falha = {n00}')
print(f'  temporal {a.sum()}/{len(a)} = {100*a.mean():.2f} %   por respostas {b.sum()}/{len(b)} = {100*b.mean():.2f} %')
print(f'  McNemar exato (binomial sobre os {n10+n01} pares discordantes): p = '
      f'{stats.binomtest(n10, n10+n01, 0.5).pvalue:.4f}')
N=len(a); dif=(n01-n10)/N; se=np.sqrt(n10+n01-(n01-n10)**2/N)/N
print(f'  diferença de proporções emparelhadas = {100*dif:+.2f} pp   '
      f'IC 95 % = [{100*(dif-1.959964*se):+.2f}; {100*(dif+1.959964*se):+.2f}] pp')
