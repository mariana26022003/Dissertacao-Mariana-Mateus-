# Pacote de dados — supervisão por respostas vs. supervisão temporal no TM12

Este repositório reúne os dados, as bases de análise e os scripts que suportam os resultados experimentais da dissertação de Mariana Mateus e do artigo *Experimental Evaluation of Response-Driven Supervision for Low-Code Industrial Robot Workflows Using ROS 2 and MoveIt 2*, submetido a *The International Journal of Advanced Manufacturing Technology*.

Revisão do pacote: 2026-10-12. Os valores aqui indicados correspondem à análise final, após a auditoria de proveniência de 2026-10-08.

## Estrutura

| Caminho | Conteúdo |
|---|---|
| `01_logs_originais/sincronizacao_csv/` | Registos em bruto produzidos pelo nó de métricas de sincronização (`metricas_sincronizacao`), um ficheiro de ciclos e um de eventos por estratégia. Não foram editados. |
| `03_scripts/` | Scripts de análise e de preparação dos livros Excel (ver secção «Scripts»). |
| `Testes_de_tempo_e_estados_TM12_AUDITADO.xlsx` | Base auditada da comparação entre as duas estratégias. A folha «Dados Análise» tem a coluna `estado_inclusao` e a folha «Auditoria de inclusão» documenta o critério e as contagens. |
| `Tabela_Correspondencia_Ciclos_AUDITADA.xlsx` | Correspondência entre cada condição cenário-repetição e o ciclo do registo primário, com o estado de inclusão. |
| `Teste_pick_TM12_final_CORRIGIDO.xlsx` | Ensaios de recolha e deposição, repetibilidade e comparação de planeadores sem obstáculos. |
| `Testes_Obstaculos_TM12_final_CORRIGIDO.xlsx` | Ensaios com obstáculos e comparação de planeadores com obstáculos. |
| `Graficos_Capitulo5_EDITAVEIS.xlsx` | Dados, testes em fórmula e gráficos editáveis do Capítulo 5. |
| `05_resultados/RESULTADOS_analise_estatistica_cap5.txt` | Saída de `analise_estatistica_cap5.py` sobre os ficheiros deste pacote. |
| `05_resultados/RESULTADOS_testes_emparelhados.txt` | Saída de `testes_emparelhados.py` sobre os ficheiros deste pacote. |
| `05_resultados/RESULTADOS_analise_sensibilidade.txt` | Saída de `analise_sensibilidade.py` sobre os ficheiros deste pacote. |

Os vídeos e as imagens dos ensaios estão na [release mais recente do repositório](https://github.com/mariana26022003/Dissertacao-Mariana-Mateus-/releases/latest).

## Base da análise comparativa

O desenho nominal tem 8 cenários com 20 repetições cada, ou seja, 160 condições executadas pelas duas estratégias. Uma condição só entra na análise quando existe registo primário das duas estratégias. A exclusão depende apenas da disponibilidade de evidência primária e nunca do resultado do ensaio.

| Etapa | Número | Utilização |
|---|---|---|
| Condições no desenho nominal | 160 | 8 cenários × 20 repetições |
| Condições excluídas por falta de registo primário | 13 | 9 da estratégia temporal e 4 da estratégia por respostas |
| Pares incluídos | 147 | Base da taxa de sucesso (294 ciclos) |
| Pares concluídos com sucesso pelas duas estratégias | 138 | Base do tempo de ciclo e da espera improdutiva |
| … dos quais com correspondência confirmada diretamente nas duas estratégias | 126 | Base da análise de sensibilidade |
| … dos quais com correspondência reconstruída em pelo menos uma estratégia | 12 | Identificados pela posição na sequência cronológica |

As condições excluídas são R10-20, O03-16 a O03-20 e O17-18 a O17-20 (sem registo da estratégia temporal) e O13-17 a O13-20 (sem registo da estratégia por respostas). A exclusão das repetições 18 a 20 do O17 foi acrescentada na revisão de 2026-10-08. As versões anteriores do pacote, com 150 pares incluídos e 141 pares temporais, ficam substituídas por esta.

## Resultados principais

Estratégia temporal (Sleep) e estratégia por respostas (response-driven). As diferenças são calculadas como «por respostas − temporal».

| Grandeza | Temporal | Por respostas | Diferença |
|---|---|---|---|
| Tempo de ciclo médio, 138 pares (s) | 49,404 | 47,011 | −2,393 (−4,84 %) |
| IC 95 % da média (s) | [48,895; 49,913] | [46,466; 47,556] | [−3,025; −1,761] |
| Estimativa de Hodges-Lehmann da diferença (s) | — | — | −2,508 [−3,001; −1,941] |
| Espera improdutiva média por ciclo, 138 pares (s) | 3,270 | 1,062 | −67,53 % |
| Taxa de sucesso do ciclo, 147 pares | 143/147 (97,28 %) | 142/147 (96,60 %) | −0,68 pp |

No tempo de ciclo, o teste de Shapiro-Wilk rejeita a normalidade das diferenças (p = 0,00256). O teste de Wilcoxon emparelhado dá p = 5,96 × 10⁻¹¹ e a correlação rank-biserial é −0,642. O teste t emparelhado, apresentado como verificação, dá t(137) = −7,487 e p = 7,75 × 10⁻¹². Na taxa de sucesso, o teste exato de McNemar dá p = 1,000.

Na análise de sensibilidade, restrita aos 126 pares com correspondência confirmada diretamente, a redução do tempo de ciclo é de 5,27 % (Hodges-Lehmann −2,708 s [−3,197; −2,128], Wilcoxon p = 1,93 × 10⁻¹¹) e a da espera improdutiva é de 67,59 %. A conclusão não se altera.

## Reprodução dos resultados

Requisitos: Python 3 com `numpy`, `scipy` e `openpyxl`.

```
pip install numpy scipy openpyxl
```

Os scripts leem os livros Excel pelo nome, por isso devem ser executados a partir da pasta raiz do repositório:

```
python 03_scripts/analise_estatistica_cap5.py
python 03_scripts/testes_emparelhados.py
python 03_scripts/analise_sensibilidade.py
```

A saída de cada um deve coincidir com o ficheiro correspondente em `05_resultados/`. Estes três ficheiros foram gerados em 2026-10-12 com Python 3.13, numpy 2.5, scipy 1.18 e openpyxl 3.1. Os scripts leem os valores guardados nos livros Excel. Se um livro for editado e gravado fora do Excel, deve ser aberto e gravado de novo no Excel antes de correr os scripts.

## Scripts

**Análise (apenas leem os livros deste pacote)**

| Script | Função |
|---|---|
| `analise_estatistica_cap5.py` | Estatística descritiva e testes da comparação entre estratégias, do ciclo de recolha e deposição, do ponto de controlo, dos obstáculos e dos planeadores. |
| `analise_sensibilidade.py` | Repete a comparação do tempo de ciclo e da espera improdutiva sobre os 126 pares com correspondência confirmada diretamente nas duas estratégias, de acordo com a coluna «Estado da identificação» da tabela de correspondência (Tabelas S8 e S9 do Online Resource 1). |
| `testes_emparelhados.py` | Testes de Friedman, Wilcoxon com correção de Holm e Q de Cochran para os planeadores, e McNemar para a taxa de sucesso das duas estratégias. |
| `tabela_trajetorias.py` | Cria `Trajetorias_Planeadas_e_TCP.xlsx`, com o número de pontos e o comprimento do percurso por planeador, cenário e repetição. |

**Preparação dos livros (registo do processamento)**

Os scripts seguintes documentam como os livros auditados e o livro de gráficos foram obtidos a partir das versões de trabalho. Partem de ficheiros intermédios que não fazem parte deste pacote (por exemplo, `Testes_de_tempo_e_estados_TM12_CORRIGIDO.xlsx`, `Tabela_Correspondencia_Ciclos.xlsx`, `Graficos_Capitulo5_EDITAVEIS_ORIGINAL.xlsx` e o módulo auxiliar `graficos_nativos.py`), pelo que não se destinam a ser executados de novo sobre o pacote. Alguns alteram os livros no próprio ficheiro.

| Script | Função |
|---|---|
| `repara_cancelamentos.py` | Acrescenta as linhas de cancelamento em falta e define o tempo de recuperação a partir do registo da plataforma de supervisão. |
| `cria_auditado.py` | Cria o livro auditado, com o estado de inclusão de cada condição e a folha «Auditoria de inclusão». |
| `atualiza_correspondencia.py` | Acrescenta o estado de inclusão à tabela de correspondência. |
| `corrige_notas.py` | Alinha as notas metodológicas dos livros com os registos. |
| `atualiza_graficos_cap5.py` | Atualiza o livro de gráficos para a base auditada. |
| `excel_testes_emparelhados.py`, `excel_hodges_lehmann.py`, `add_ic.py`, `alinha_excel.py`, `completa_graficos.py` | Acrescentam ao livro de gráficos os testes em fórmula, os intervalos de confiança e os gráficos. |
| `exporta_graficos.py`, `compoe_figuras.py` | Exportam os gráficos do livro (LibreOffice) e compõem as figuras do Capítulo 5. |

## Limitações de disponibilização

Por razões de confidencialidade exigidas pelo parceiro industrial, não são disponibilizados os registos de execução da plataforma VisionCore, os ficheiros de projeto da VisionCore, incluindo a definição dos dois fluxos low-code, nem o código-fonte do back end robótico. Os tempos de ciclo e de espera usados na análise provêm dos registos do nó de métricas em `01_logs_originais/sincronizacao_csv/`, que estão disponíveis. A correspondência entre as condições do desenho e os ciclos é estabelecida sobre esses mesmos registos.

## Autoria e contacto

Mariana Rodrigues Mateus, Escola de Engenharia, Universidade do Minho. E-mail: pg56059@uminho.pt
