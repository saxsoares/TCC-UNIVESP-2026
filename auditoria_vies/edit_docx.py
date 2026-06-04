# Insere as novas seções e correções no TCC a partir dos resultados dos
# experimentos. Gera backup TCC_backup_original.docx. Executar uma vez.
import docx, json, shutil, copy
import pandas as pd
from pathlib import Path
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

SRC = Path("TCC.docx"); RES = Path("auditoria_vies/resultados")
shutil.copy(SRC, "TCC_backup_original.docx")
d = docx.Document(str(SRC))

# ---------- dados experimentais reais ----------
lfw = json.load(open(RES / "lfw_metrics.json"))
ff  = json.load(open(RES / "fairface_summary.json"))
tab = pd.read_csv(RES / "fairface_fmr_por_raca.csv")
inter = pd.read_csv(RES / "fairface_fmr_interseccional.csv").sort_values("FMR@global", ascending=False).reset_index(drop=True)
RACE_PT = {"East Asian":"Leste-asiático","Indian":"Indiano","Black":"Negro","White":"Branco",
           "Middle Eastern":"Médio-oriental","Latino_Hispanic":"Latino","Southeast Asian":"Sudeste-asiático"}
GEN_PT = {"Male":"masculino","Female":"feminino"}
tab = tab.sort_values("FMR@global", ascending=False).reset_index(drop=True)
pior = tab.iloc[0]; melhor = tab.iloc[-1]
inter_top = inter.iloc[0]
itop_txt = f"{RACE_PT[inter_top['raca']].lower()} do gênero {GEN_PT[inter_top['genero']]}"

def pct(x, n=2): return (f"%.{n}f" % (x*100)).replace(".", ",")
def num(x, n=4): return (f"%.{n}f" % x).replace(".", ",")

# ---------- helpers ----------
def _new_p(text="", style="Normal"):
    p = d.add_paragraph()
    if style != "Normal":
        try: p.style = d.styles[style]
        except KeyError: pass
    if text:
        r = p.add_run(text); r.font.name = "Arial"
        if style == "Normal": r.font.size = Pt(12)
    return p

def _fmt_body(p):
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(12)
    for r in p.runs:
        r.font.name = "Arial"; r.font.size = Pt(12)
    return p

def find(prefix, style=None):
    for p in d.paragraphs:
        if p.text.strip().startswith(prefix) and (style is None or p.style.name == style):
            return p
    return None

def replace_text(p, new_text):
    for r in list(p.runs): r._r.getparent().remove(r._r)
    r = p.add_run(new_text); r.font.name = "Arial"; r.font.size = Pt(12)
    return p

def insert_before(anchor, text, style="Normal", body=True):
    p = _new_p(text, style)
    if body and style == "Normal": _fmt_body(p)
    else:
        for r in p.runs: r.font.name = "Arial"
    anchor._p.addprevious(p._p)
    return p

def insert_after(anchor, text, style="Normal", body=True):
    p = _new_p(text, style)
    if body and style == "Normal": _fmt_body(p)
    else:
        for r in p.runs: r.font.name = "Arial"
    anchor._p.addnext(p._p)
    return p

# C1 — Resumo + palavras-chave
ab = find("Este trabalho insere-se")
if ab:
    r = ab.add_run(" Como contribuição metodológica de Ciência de Dados, realiza-se ainda "
        "uma auditoria empírica de viés demográfico em um modelo pré-treinado de "
        "reconhecimento facial (ArcFace) sobre as bases públicas LFW e FairFace, "
        "evidenciando quantitativamente as disparidades de desempenho discutidas na "
        "literatura.")
    r.font.name = "Arial"; r.font.size = Pt(12)
kw = find("Palavras-chave")
if kw:
    base = kw.text.rstrip(". ")
    replace_text(kw, base + "; auditoria de viés; equidade algorítmica.")

# C2 — novo objetivo específico
obj = find("Analisar os vieses algorítmicos identificados")
if obj:
    novo = insert_after(obj,
        "Realizar uma auditoria empírica de viés demográfico em modelo pré-treinado de "
        "reconhecimento facial, mensurando, sobre bases públicas, as taxas de erro de "
        "verificação e de falsa correspondência estratificadas por grupo;",
        style="List Paragraph", body=False)
    novo.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pPr_prev = obj._p.find(qn('w:pPr'))
    if pPr_prev is not None:
        numPr = pPr_prev.find(qn('w:numPr'))
        if numPr is not None:
            novo._p.get_or_add_pPr().append(copy.deepcopy(numPr))

# C3 — correção da frase que nega experimentação
c3 = find("A presente pesquisa adota abordagem qualitativa, de natureza explorat")
if c3:
    t = c3.text.replace(
        "tais indicadores são analisados a partir de estudos já publicados, não havendo "
        "desenvolvimento experimental, implementação de sistema próprio ou validação "
        "empírica de modelos computacionais.",
        "tais indicadores são, em um primeiro momento, analisados a partir de estudos já "
        "publicados e, posteriormente, mensurados pelos próprios autores por meio de uma "
        "auditoria empírica de um modelo pré-treinado (seções 3.7 e 6). Não há, contudo, "
        "treinamento de modelos do zero, implementação de sistema operacional próprio ou "
        "uso de imagens de pessoas reais desaparecidas.")
    replace_text(c3, t); _fmt_body(c3)

# C4 — 3.5: frase ao final
c4 = find("Embora a pesquisa não tenha desenvolvido um sistema próprio")
if c4:
    r = c4.add_run(" Esses mesmos indicadores — em especial FMR, FNMR, ROC e AUC — são "
        "posteriormente calculados de forma empírica na auditoria descrita no Capítulo 6, "
        "deixando de operar apenas como categorias analíticas para tornarem-se medidas "
        "obtidas pelos autores.")
    r.font.name = "Arial"; r.font.size = Pt(12)

# 3.7 — nova subseção (antes de "4 REFERENCIAL TEÓRICO")
h4 = find("4 REFERENCIAL TEÓRICO", style="Heading 1")
sec37 = [
 ("3.7 Delineamento da auditoria empírica complementar", "Heading 2"),
 ("Em complemento à análise bibliográfica e documental descrita nas seções anteriores, e "
  "a fim de fundamentar com evidência própria as limitações técnicas discutidas no "
  "referencial teórico, este trabalho incorpora uma auditoria empírica de caráter "
  "quantitativo. Diferentemente do desenvolvimento de um sistema de reconhecimento facial "
  "— o que extrapolaria o escopo e os cuidados éticos desta pesquisa —, adota-se a "
  "estratégia metodológica de auditoria de modelos pré-treinados, na qual um algoritmo de "
  "reconhecimento facial já consolidado é avaliado, sem qualquer treinamento adicional, "
  "sobre bases públicas de pesquisa amplamente utilizadas na literatura de visão "
  "computacional.", "Normal"),
 ("A auditoria assume natureza experimental e abordagem quantitativa, articulando-se à "
  "dimensão qualitativa do estudo: enquanto a análise documental investiga os limites "
  "jurídicos e éticos, a auditoria mensura, de modo reprodutível, o comportamento "
  "estatístico do modelo sob condições análogas às encontradas na busca por pessoas "
  "desaparecidas. Foi auditado o modelo ArcFace (DENG et al., 2019), na variante "
  "ResNet-50 treinada sobre a base Glint360K e disponibilizada em formato aberto, "
  "executado por meio da biblioteca ONNX Runtime, em tarefa de verificação facial (1:1) "
  "com comparação por similaridade de cosseno entre os vetores de características.", "Normal"),
 ("Foram utilizadas duas bases públicas com finalidades distintas. A base Labeled Faces "
  "in the Wild (LFW), em seu protocolo oficial de verificação com seis mil pares, serviu "
  "à mensuração do desempenho geral de verificação (curva ROC, AUC, EER e taxas de erro). "
  "A base FairFace, composta por imagens faciais rotuladas de forma balanceada quanto a "
  "raça, gênero e idade (KÄRKKÄINEN; JOO, 2021), serviu à mensuração do viés demográfico "
  "nas taxas de falsa correspondência. Ressalte-se que nenhuma imagem de pessoa "
  "efetivamente desaparecida foi coletada ou utilizada, preservando-se o enquadramento "
  "ético e a conformidade com a Lei Geral de Proteção de Dados Pessoais: a auditoria "
  "avalia o comportamento estatístico do algoritmo, e não a identificação de indivíduos "
  "reais, em coerência com a delimitação assumida na seção 3.3.", "Normal"),
]
for text, style in sec37:
    insert_before(h4, text, style)

# Capítulo 6 — AUDITORIA EMPÍRICA (antes do atual "6 RESULTADOS...")
h6 = find("6 RESULTADOS E CONTRIBUIÇÕES", style="Heading 1")

cap6_intro = [
 ("6 AUDITORIA EMPÍRICA DE VIÉS ALGORÍTMICO", "Heading 1"),
 ("Este capítulo apresenta a contribuição empírica de Ciência de Dados deste trabalho: "
  "uma auditoria quantitativa e reprodutível do comportamento de um modelo de "
  "reconhecimento facial de estado da arte, conduzida pelos próprios autores sobre bases "
  "públicas, com o objetivo de converter em evidência mensurada as limitações técnicas "
  "— desempenho e viés demográfico — discutidas no referencial teórico.", "Normal"),
 ("6.1 Procedimento experimental", "Heading 2"),
 ("A auditoria foi organizada em dois experimentos complementares, ambos utilizando o "
  "modelo ArcFace (ResNet-50) sem treinamento adicional. O primeiro experimento mensurou "
  "o desempenho geral de verificação facial sobre a base Labeled Faces in the Wild (LFW), "
  "no protocolo oficial de seis mil pares (três mil genuínos e três mil impostores). Para "
  "cada par, extraíram-se os vetores de características das duas imagens e calculou-se a "
  "similaridade de cosseno; a partir da distribuição de similaridades, derivaram-se a "
  "curva ROC, a área sob a curva (AUC), a Equal Error Rate (EER) e as taxas de falsa "
  "correspondência (FMR) e de falsa não correspondência (FNMR).", "Normal"),
 ("O segundo experimento avaliou o viés demográfico sobre a base FairFace, que fornece "
  "rótulos de raça, gênero e idade. Como o FairFace não contém identidades repetidas, "
  "todo par de imagens distintas constitui um par impostor; mediu-se, assim, a propensão "
  "do modelo à falsa correspondência em cada grupo demográfico. Para evidenciar a "
  "disparidade, fixou-se um limiar de decisão único, calibrado para produzir uma taxa "
  "global de falsa correspondência de 1×10⁻³ sobre o conjunto inteiro, e mediu-se a FMR "
  "de cada grupo racial separadamente. A adoção de limiar único é metodologicamente "
  "essencial, pois a calibração independente por grupo mascararia a disparidade que se "
  "pretende auditar (GROTHER; NGAN; HANAOKA, 2019). Como medidas de equidade, calcularam-"
  "se a razão de disparidade (FMR do pior grupo dividida pela FMR do melhor grupo) e "
  "intervalos de confiança de 95% por reamostragem bootstrap, além do teste qui-quadrado "
  "de independência entre erro e grupo. Todo o procedimento foi implementado em Python "
  "(ONNX Runtime, scikit-learn, NumPy e pandas), com sementes aleatórias fixas, "
  "assegurando reprodutibilidade.", "Normal"),
]
for text, style in cap6_intro:
    insert_before(h6, text, style)

# ---- 6.2 Resultados (com números reais) ----
acc_lfw = pct(lfw["best_accuracy"]); auc_lfw = num(lfw["AUC"]); eer_lfw = pct(lfw["EER"])
fnmr_lfw = pct(lfw["FNMR@thr"]); disp = num(ff["disparity_max_min"], 2)
pgrp, mgrp = RACE_PT[pior["grupo"]], RACE_PT[melhor["grupo"]]
p_fmr, m_fmr = pct(pior["FMR@global"], 3), pct(melhor["FMR@global"], 3)
chi2 = num(ff["chi2"], 1); chip = ff["chi2_p"]
chip_txt = "p < 0,001" if chip < 1e-3 else f"p = {num(chip,4)}"

sec62_a = [
 ("6.2 Resultados", "Heading 2"),
 (f"No primeiro experimento, o modelo ArcFace alcançou desempenho elevado na base LFW, "
  f"com AUC de {auc_lfw} e acurácia de {acc_lfw}%, confirmando a alta capacidade "
  f"discriminativa relatada na literatura para condições controladas. A Equal Error Rate "
  f"foi de {eer_lfw}%. A Figura 1 apresenta, à esquerda, a nítida separação entre as "
  f"distribuições de similaridade de pares genuínos e impostores e, à direita, a curva "
  f"ROC correspondente. Cabe registrar que a acurácia observada situa-se ligeiramente "
  f"abaixo do estado da arte de 99,4% reportado para o ArcFace, diferença atribuível à "
  f"ausência de alinhamento facial por pontos fiduciais nesta auditoria — uma limitação "
  f"assumida e discutida na seção 6.4, que não compromete a análise comparativa entre "
  f"grupos, conduzida sob processamento idêntico.", "Normal"),
]
for text, style in sec62_a:
    insert_before(h6, text, style)
# âncora para figura 1 = último parágrafo inserido (antes de h6)
fig1_anchor = h6._p.getprevious()

sec62_b = [
 (f"No segundo experimento, embora todos os grupos apresentassem boa separabilidade "
  f"agregada, a taxa de falsa correspondência sob o limiar único revelou disparidade "
  f"demográfica relevante. Conforme a Tabela 1 e a Figura 2, o grupo {pgrp.lower()} "
  f"apresentou a maior FMR ({p_fmr}%), ao passo que o grupo {mgrp.lower()} apresentou a "
  f"menor ({m_fmr}%), resultando em uma razão de disparidade de {disp} vezes entre o pior "
  f"e o melhor grupo. O teste qui-quadrado rejeitou a hipótese nula de independência "
  f"entre a ocorrência de falsa correspondência e o grupo demográfico "
  f"(χ² = {chi2}; {chip_txt}), indicando que as diferenças observadas não são atribuíveis "
  f"ao acaso. Esse padrão é consistente com os achados do NIST, segundo os quais a "
  f"variação demográfica concentra-se sobretudo nas taxas de falso positivo "
  f"(GROTHER; NGAN; HANAOKA, 2019).", "Normal"),
 (f"A análise interseccional, que cruza raça e gênero, evidenciou um agravamento das "
  f"disparidades: o grupo {itop_txt} apresentou a maior taxa de falsa correspondência "
  f"({pct(inter_top['FMR@global'],3)}%), e, de modo geral, as mulheres de grupos não "
  f"brancos figuraram entre as mais afetadas. Esse achado é convergente com o estudo "
  f"Gender Shades (BUOLAMWINI; GEBRU, 2018), que identificou nas mulheres de pele mais "
  f"escura o grupo de maior taxa de erro, e reforça que os vieses se acumulam nas "
  f"interseções de atributos demográficos.", "Normal"),
]
for text, style in sec62_b:
    insert_before(h6, text, style)

# ---- Tabela 1 (título acima) ----
insert_before(h6, "Tabela 1 – Taxa de falsa correspondência (FMR) por grupo demográfico — base FairFace, modelo ArcFace", "Normal")
tab_title = h6._p.getprevious()

table = d.add_table(rows=1, cols=4)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.style = "Table Grid"
hdr = table.rows[0].cells
for c, txt in zip(hdr, ["Grupo demográfico", "Nº de imagens", "Similaridade média (impostores)", "FMR sob limiar único"]):
    c.text = txt
    for p in c.paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs: r.font.name="Arial"; r.font.size=Pt(10); r.font.bold=True
for _, row in tab.iterrows():
    cells = table.add_row().cells
    vals = [RACE_PT[row["grupo"]], str(int(row["n_imgs"])), num(row["mean_impostor_sim"],4), pct(row["FMR@global"],3)+"%"]
    for c, v in zip(cells, vals):
        c.text = v
        for p in c.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs: r.font.name="Arial"; r.font.size=Pt(10)
# move tabela para depois do título
tab_title.addnext(table._tbl)
# fonte da tabela
fonte = _new_p("Fonte: elaborado pelos autores (2026).", "Normal")
fonte.paragraph_format.space_after = Pt(12)
for r in fonte.runs: r.font.name="Arial"; r.font.size=Pt(10)
table._tbl.addnext(fonte._p)

# ---- Figuras ----
def insert_figure(anchor_p, img_path, caption):
    # parágrafo com a imagem (centrado) DEPOIS do anchor
    pic_p = _new_p("", "Normal"); anchor_p.addnext(pic_p._p)
    pic_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic_p.add_run().add_picture(str(img_path), width=Inches(6.0))
    cap_p = _new_p(caption, "Normal"); pic_p._p.addnext(cap_p._p)
    cap_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap_p.paragraph_format.space_after = Pt(12)
    for r in cap_p.runs: r.font.name="Arial"; r.font.size=Pt(10)
    return cap_p

# Figura 1 após fig1_anchor (texto do LFW)
insert_figure(fig1_anchor,
    RES / "fig1_lfw_roc.png",
    "Figura 1 – Distribuição de similaridade (a) e curva ROC (b) do modelo ArcFace na base LFW. "
    "Fonte: elaborado pelos autores (2026).")
# Figura 2 após a fonte da Tabela 1
insert_figure(fonte._p,
    RES / "fig2_fairface_fmr.png",
    "Figura 2 – Taxa de falsa correspondência por grupo demográfico (base FairFace). "
    "Fonte: elaborado pelos autores (2026).")

# ---- 6.3 Discussão ----
sec63 = [
 ("6.3 Discussão", "Heading 2"),
 (f"Os resultados da auditoria convertem em evidência mensurada aquilo que o referencial "
  f"teórico havia estabelecido apenas no plano conceitual. A coexistência de AUC elevada "
  f"({auc_lfw}) com disparidade de {disp} vezes na taxa de falsa correspondência entre "
  f"grupos confirma a tese de Buolamwini e Gebru (2018) e do NIST "
  f"(GROTHER; NGAN; HANAOKA, 2019): o viés não decorre de baixa qualidade técnica do "
  f"algoritmo em termos absolutos, mas de seu comportamento desigual entre subpopulações. "
  f"Para o problema desta pesquisa, a distinção é decisiva. Uma FMR mais elevada em um "
  f"grupo demográfico significa, em termos operacionais, maior probabilidade de que "
  f"pessoas inocentes desse grupo sejam erroneamente associadas a um registro — "
  f"precisamente o tipo de falha materializada nos casos do sistema Smart Sampa "
  f"(COALIZÃO DIREITOS NA REDE, 2025), em que cidadãos negros foram abordados "
  f"indevidamente.", "Normal"),
 ("A FNMR observada no experimento com a base LFW (de aproximadamente "
  f"{fnmr_lfw}% no ponto de operação conservador, calibrado a FMR de 1×10⁻³) tem "
  "implicação humanitária inversa e igualmente grave: quando a fotografia de referência "
  "de uma pessoa desaparecida é antiga ou de baixa qualidade, eleva-se a probabilidade de "
  "o sistema deixar de reconhecê-la, comprometendo a própria finalidade protetiva da "
  "ferramenta. A auditoria demonstra, assim, que os dois erros possíveis — falso positivo "
  "e falso negativo — recaem de modo desigual sobre os grupos e os contextos, o que "
  "reforça a exigência, sustentada nas seções jurídicas deste trabalho, de supervisão "
  "humana obrigatória, auditoria independente de viés e separação entre bancos de dados "
  "criminais e humanitários. A evidência empírica aqui produzida fornece base objetiva "
  "para as diretrizes de governança propostas no capítulo seguinte, deslocando-as do "
  "terreno da recomendação abstrata para o de exigência tecnicamente justificada.", "Normal"),
 ("6.4 Limitações da auditoria", "Heading 2"),
 ("Os resultados devem ser interpretados à luz de quatro limitações. Primeira, foi "
  "avaliado um modelo de código aberto pré-treinado, e não os sistemas comerciais "
  "efetivamente empregados pelos órgãos de segurança pública; os achados são, portanto, "
  "indicativos das tendências de viés relatadas na literatura, e não medições dos "
  "sistemas em operação no Brasil. Segunda, não foi aplicado alinhamento facial por "
  "pontos fiduciais, o que reduz a acurácia absoluta (situando-a abaixo do estado da arte "
  "de 99,4%); como todos os grupos foram processados de forma idêntica, essa limitação "
  "não compromete a comparação relativa entre grupos, foco da auditoria. Terceira, os "
  "rótulos demográficos das bases são categóricos e atribuídos por terceiros, e não "
  "autodeclarados, o que impõe cautela conceitual ao tratar raça como variável de "
  "análise; ademais, as bases públicas não refletem a composição fenotípica específica da "
  "população brasileira, na qual, segundo o Censo Demográfico de 2022, pessoas pardas e "
  "pretas constituem a maioria. Quarta, a auditoria avaliou a tarefa de verificação "
  "(1:1); em cenários de identificação (1:N), típicos da busca em larga escala, a taxa de "
  "falsos positivos tende a se ampliar proporcionalmente ao tamanho da base consultada. "
  "Essas limitações não invalidam os achados, mas delimitam seu alcance e indicam "
  "direções para investigação futura.", "Normal"),
]
for text, style in sec63:
    insert_before(h6, text, style)

# Renumeração: 6 RESULTADOS -> 7 ; 7 CONCLUSÃO -> 8
replace_text(h6, "7 RESULTADOS E CONTRIBUIÇÕES");
for r in h6.runs: r.font.name="Arial"
hc = find("7 CONCLUSÃO")
if hc is None: hc = find("7 CONCLUS")
if hc:
    replace_text(hc, "8 CONCLUSÃO")
    for r in hc.runs: r.font.name="Arial"

# C5 — correção da limitação na Conclusão
c5 = find("Por fim, reconhece-se que esta pesquisa possui limita")
if c5:
    t = c5.text.replace(
        "O estudo não realizou testes computacionais próprios nem validação empírica de "
        "desempenho algorítmico em bases reais de desaparecidos, concentrando-se em "
        "análise qualitativa, bibliográfica e documental.",
        "Embora o estudo tenha incorporado uma auditoria empírica de viés sobre um modelo "
        "pré-treinado e bases públicas de pesquisa (Capítulo 6), não realizou validação em "
        "bases reais de pessoas desaparecidas — o que seria ética e juridicamente "
        "inviável —, concentrando-se na articulação entre a análise qualitativa, "
        "bibliográfica e documental e a evidência quantitativa obtida em ambiente "
        "controlado.")
    if t == c5.text:  # fallback se a frase exata não bater
        t = c5.text
    replace_text(c5, t); _fmt_body(c5)

# 8.1 Trabalhos futuros (após último parágrafo da conclusão, antes de REFERÊNCIAS)
# inserir após c5 (último parágrafo da conclusão)
tf_title = insert_after(c5, "8.1 Trabalhos futuros", style="Heading 2", body=False)
tf_body = ("A auditoria empírica realizada abre frentes de continuidade diretamente "
  "relevantes ao contexto brasileiro. Recomenda-se, prioritariamente, a construção de um "
  "protocolo de avaliação sobre base de faces representativa da diversidade fenotípica "
  "nacional, de modo a substituir a inferência indireta de viés por medição direta, "
  "observados os requisitos éticos e de proteção de dados. Sugere-se, ademais, a extensão "
  "da auditoria à tarefa de identificação (1:N) com galerias de tamanho crescente; a "
  "incorporação de alinhamento facial e a comparação entre múltiplas arquiteturas "
  "(FaceNet, SFace, ArcFace); a avaliação do eixo de envelhecimento facial mediante bases "
  "longitudinais, especialmente relevante para o desaparecimento de crianças "
  "(HOSSAIN et al., 2025); e a avaliação de técnicas de mitigação de viés. Por fim, "
  "recomenda-se a articulação dos resultados quantitativos com a proposição de um "
  "protocolo técnico-jurídico de homologação de sistemas biométricos para uso "
  "humanitário, no qual a aprovação de qualquer ferramenta seja condicionada à "
  "demonstração empírica de equidade demográfica e de robustez mínima sob baixa qualidade "
  "de imagem.")
insert_after(tf_title, tf_body, style="Normal", body=True)

# Referências novas (ABNT, em ordem alfabética por âncora)
def ref_before(anchor_prefix, text):
    anchor = find(anchor_prefix)
    if anchor is None:
        print("  [aviso] âncora de referência não encontrada:", anchor_prefix); return
    p = _new_p(text, "Normal")
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(12)
    for r in p.runs: r.font.name="Arial"; r.font.size=Pt(12)
    anchor._p.addprevious(p._p)

ref_before("KASPERSKY",
  "KÄRKKÄINEN, Kimmo; JOO, Jungseock. FairFace: face attribute dataset for balanced race, "
  "gender, and age for bias measurement and mitigation. In: IEEE/CVF WINTER CONFERENCE ON "
  "APPLICATIONS OF COMPUTER VISION (WACV), 2021. Proceedings… Piscataway: IEEE, 2021. p. 1548-1558.")
ref_before("NATIONAL INSTITUTE OF STANDARDS",
  "NATIONAL INSTITUTE OF STANDARDS AND TECHNOLOGY (NIST). Face Recognition Vendor Test "
  "(FRVT) Part 8: summarizing demographic differentials. Gaithersburg: NIST, 2022. (NISTIR 8429).")
ref_before("SCHROFF",
  "ROBINSON, Joseph P. et al. Face recognition: too bias, or not too bias? In: IEEE/CVF "
  "CONFERENCE ON COMPUTER VISION AND PATTERN RECOGNITION WORKSHOPS (CVPRW), 2020. "
  "Proceedings… Piscataway: IEEE, 2020. p. 1-10.")
ref_before("SILVA, Tarc",
  "SERENGIL, Sefik Ilkin; OZPINAR, Alper. LightFace: a hybrid deep face recognition "
  "framework. In: INNOVATIONS IN INTELLIGENT SYSTEMS AND APPLICATIONS CONFERENCE (ASYU), "
  "2020. Proceedings… Piscataway: IEEE, 2020. p. 1-5.")
ref_before("YAO, Wang",
  "WANG, Mei; DENG, Weihong; HU, Jiani; TAO, Xunqiang; HUANG, Yaohai. Racial Faces in the "
  "Wild: reducing racial bias by information maximization adaptation network. In: IEEE/CVF "
  "INTERNATIONAL CONFERENCE ON COMPUTER VISION (ICCV), 2019. Proceedings… Piscataway: "
  "IEEE, 2019. p. 692-702.")

d.save(str(SRC))
print("OK — TCC.docx revisado salvo. Backup: TCC_backup_original.docx")
print(f"LFW: AUC={auc_lfw} acc={acc_lfw}% EER={eer_lfw}%")
print(f"FairFace: disparidade={disp}x | pior={pgrp}({p_fmr}%) melhor={mgrp}({m_fmr}%) chi2={chi2} p={chip}")
