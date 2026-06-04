# Padroniza os slides 1-7 e 19-23 no mesmo estilo dos slides 8-18
# (Calibri, azul 1F4E79, titulos no topo, circulos/cards/pilares). NAO toca em 8-18.
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

ACCENT = RGBColor(0xB0, 0x20, 0x2E)   # vermelho UNIVESP (escuro) — titulos/preenchimentos
ACCENT2 = RGBColor(0xE6, 0x33, 0x29)  # vermelho vivo — reguas/destaques
CARD = RGBColor(0xF7, 0xE6, 0xE7)     # rosa-claro — fundo de card
LIGHTTINT = RGBColor(0xF2, 0xD8, 0xDA) # tom claro para texto sobre fundo vermelho
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x26, 0x26, 0x26)
GREY = RGBColor(0x59, 0x59, 0x59)

prs = Presentation("Apresentação TCC.pptx")
S = prs.slides

def clear(slide, keep_pictures=False):
    for sh in list(slide.shapes):
        if keep_pictures and sh.shape_type == 13:
            continue
        sh._element.getparent().remove(sh._element)

def _set(tf, runs, anchor=None):
    tf.word_wrap = True
    if anchor: tf.vertical_anchor = anchor
    for i, r in enumerate(runs):
        text, size, bold, color, align = r
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run(); run.text = text
        run.font.size = Pt(size); run.font.bold = bold
        run.font.name = "Calibri"; run.font.color.rgb = color

def box(slide, l, t, w, h, runs, anchor=None):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    _set(tb.text_frame, runs, anchor); return tb

def bar(slide, l, t, w, h, fill):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = fill; sh.line.fill.background()
    sh.shadow.inherit = False; return sh

def title(slide, text, color=ACCENT):
    box(slide, 0.6, 0.3, 8.8, 0.9, [(text, 28, True, color, PP_ALIGN.LEFT)])
    bar(slide, 0.62, 1.15, 2.2, 0.06, ACCENT2)   # regua de destaque sob o titulo

def circle(slide, n, l, t, d=0.6, fill=ACCENT):
    c = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(l), Inches(t), Inches(d), Inches(d))
    c.fill.solid(); c.fill.fore_color.rgb = fill; c.line.fill.background()
    c.shadow.inherit = False
    tf = c.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = str(n); r.font.size = Pt(20); r.font.bold = True
    r.font.color.rgb = WHITE; r.font.name = "Calibri"
    return c

def num_item(slide, n, top, head, body=None, left=0.9, tw=7.7, hsize=20, bsize=15):
    circle(slide, n, left, top, 0.58)
    runs = [(head, hsize, True, ACCENT, PP_ALIGN.LEFT)]
    if body: runs.append((body, bsize, False, DARK, PP_ALIGN.LEFT))
    box(slide, left + 0.85, top - 0.12, tw, 1.1, runs)

def chip(slide, n, top, label, desc=None, left=0.8, w=8.5, h=0.95):
    rr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(w), Inches(h))
    rr.fill.solid(); rr.fill.fore_color.rgb = CARD; rr.line.color.rgb = ACCENT2
    rr.line.width = Pt(0.75); rr.shadow.inherit = False
    circle(slide, n, left + 0.22, top + (h-0.58)/2, 0.58)
    runs = [(label, 18, True, ACCENT, PP_ALIGN.LEFT)]
    if desc: runs.append((desc, 13, False, GREY, PP_ALIGN.LEFT))
    box(slide, left + 1.05, top + 0.10, w - 1.3, h - 0.2, runs, MSO_ANCHOR.MIDDLE)

def card(slide, l, t, w, h, header, lines, header_fill=ACCENT):
    rr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    rr.fill.solid(); rr.fill.fore_color.rgb = CARD; rr.line.color.rgb = ACCENT2
    rr.line.width = Pt(1); rr.shadow.inherit = False
    hb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(0.7))
    hb.fill.solid(); hb.fill.fore_color.rgb = header_fill; hb.line.fill.background()
    hb.shadow.inherit = False
    _set(hb.text_frame, [(header, 18, True, WHITE, PP_ALIGN.CENTER)], MSO_ANCHOR.MIDDLE)
    runs = [(("• " + ln), 15, False, DARK, PP_ALIGN.LEFT) for ln in lines]
    box(slide, l + 0.3, t + 0.85, w - 0.6, h - 1.0, runs)

def pillar(slide, l, t, w, h, n, label, desc):
    rr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    rr.fill.solid(); rr.fill.fore_color.rgb = ACCENT; rr.line.fill.background()
    rr.shadow.inherit = False
    box(slide, l, t + 0.25, w, 0.9, [(str(n), 40, True, WHITE, PP_ALIGN.CENTER)])
    box(slide, l + 0.15, t + 1.15, w - 0.3, h - 1.3,
        [(label, 16, True, WHITE, PP_ALIGN.CENTER), (desc, 12, False, LIGHTTINT, PP_ALIGN.CENTER)],
        MSO_ANCHOR.TOP)

# ===== SLIDE 1 — Capa =====
s = S[0]; clear(s)
bar(s, 0, 0, 10, 1.7, ACCENT)
box(s, 0.6, 0.45, 8.8, 1.0,
    [("UNIVERSIDADE VIRTUAL DO ESTADO DE SÃO PAULO", 16, True, WHITE, PP_ALIGN.CENTER),
     ("Bacharelado em Ciência de Dados", 14, False, LIGHTTINT, PP_ALIGN.CENTER)],
    MSO_ANCHOR.MIDDLE)
box(s, 0.8, 2.6, 8.4, 2.4,
    [("Análise de Viabilidade Técnica e Jurídica no Uso da Inteligência Artificial "
      "para Apoio à Identificação de Pessoas Desaparecidas", 30, True, ACCENT, PP_ALIGN.CENTER)],
    MSO_ANCHOR.MIDDLE)
bar(s, 3.4, 5.15, 3.2, 0.06, ACCENT2)
box(s, 0.8, 5.4, 8.4, 1.2,
    [("Turma 004 — Grupo 2", 16, True, DARK, PP_ALIGN.CENTER),
     ("Orientador: Prof. Gustavo Figueredo Rodrigues de Sousa", 15, False, GREY, PP_ALIGN.CENTER)],
    MSO_ANCHOR.MIDDLE)
bar(s, 0, 6.85, 10, 0.65, ACCENT)
box(s, 0.6, 6.9, 8.8, 0.55, [("São Paulo · 2026", 14, True, WHITE, PP_ALIGN.CENTER)], MSO_ANCHOR.MIDDLE)

# ===== SLIDE 2 — Integrantes =====
s = S[1]; clear(s)
title(s, "Integrantes do Grupo")
autores = ["Artemio Costa Canossa", "Bruna de Carvalho Regis", "Fernando Bandeira Soares",
           "Kelly Cristina Máximo de Almeida", "Maria Luiza Ribeiro de Araujo",
           "Mario Henrique Gimenes Santana", "Rafael Oliveira da Cruz", "Valdina Alves Cogue"]
col = [autores[:4], autores[4:]]
for ci, names in enumerate(col):
    l = 0.8 + ci * 4.7
    rr = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(l), Inches(1.7), Inches(4.3), Inches(4.3))
    rr.fill.solid(); rr.fill.fore_color.rgb = CARD; rr.line.color.rgb = ACCENT2
    rr.line.width = Pt(1); rr.shadow.inherit = False
    for ri, nm in enumerate(names):
        top = 2.1 + ri * 0.95
        circle(s, ri + 1 + ci*4, l + 0.3, top, 0.5, ACCENT)
        box(s, l + 1.0, top - 0.08, 3.1, 0.7, [(nm, 16, True, ACCENT, PP_ALIGN.LEFT)], MSO_ANCHOR.MIDDLE)

# ===== SLIDE 3 — Problema/Contexto =====
s = S[2]; clear(s)
title(s, "Problema e Contexto")
box(s, 0.9, 1.45, 8.4, 0.8,
    [("O desaparecimento de pessoas é um problema social e humanitário relevante no Brasil.",
      18, False, DARK, PP_ALIGN.LEFT)])
itens3 = [("Elevado número de casos", "Demanda social persistente e de grande impacto humano."),
          ("Fragmentação das bases de dados", "Registros dispersos entre órgãos, sem integração."),
          ("Uso crescente de Inteligência Artificial", "Reconhecimento facial cada vez mais presente."),
          ("Lacunas técnicas, jurídicas e éticas", "Ausência de regulamentação e de auditoria.")]
for i, (h, b) in enumerate(itens3):
    chip(s, i + 1, 2.45 + i * 1.05, h, b)

# ===== SLIDE 4 — Motivação (mantém imagem existente) =====
s = S[3]; clear(s, keep_pictures=True)
title(s, "Motivação do Trabalho")
mot = [("Uso crescente do reconhecimento facial", "Amplamente adotado na segurança pública."),
       ("Ausência de regulamentação humanitária", "Sem marco específico para a busca de desaparecidos."),
       ("Riscos no tratamento de dados biométricos", "Implicações técnicas, jurídicas e éticas.")]
for i, (h, b) in enumerate(mot):
    num_item(s, i + 1, 2.0 + i * 1.45, h, b, left=0.9, tw=4.5, hsize=18, bsize=13)

# ===== SLIDE 5 — Objetivo geral =====
s = S[4]; clear(s)
title(s, "Objetivo Geral")
rr = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(2.2), Inches(8.0), Inches(3.0))
rr.fill.solid(); rr.fill.fore_color.rgb = ACCENT; rr.line.fill.background(); rr.shadow.inherit = False
_set(rr.text_frame,
     [("Analisar a viabilidade técnica e jurídica do uso da Inteligência Artificial, "
       "especialmente do reconhecimento facial, como apoio à localização e identificação "
       "de pessoas desaparecidas no Brasil.", 24, True, WHITE, PP_ALIGN.CENTER)],
     MSO_ANCHOR.MIDDLE)

# ===== SLIDE 6 — Objetivos específicos =====
s = S[5]; clear(s)
title(s, "Objetivos Específicos")
esp = ["Apresentar os fundamentos técnicos do reconhecimento facial",
       "Identificar limitações técnicas e vieses algorítmicos",
       "Analisar o marco jurídico brasileiro aplicável",
       "Comparar experiências internacionais de governança",
       "Propor diretrizes de governança jurídica"]
for i, t in enumerate(esp):
    num_item(s, i + 1, 1.7 + i * 1.02, t, None, left=0.9, tw=7.8, hsize=19)

# ===== SLIDE 7 — Fundamentação =====
s = S[6]; clear(s)
title(s, "Fundamentação Teórica")
fund = ["Inteligência Artificial", "Reconhecimento facial e biometria",
        "Vieses algorítmicos", "Proteção de dados pessoais", "Direitos fundamentais"]
for i, t in enumerate(fund):
    chip(s, i + 1, 1.6 + i * 1.02, t)

# ===== SLIDE 19 — Experiências Internacionais =====
s = S[18]; clear(s)
title(s, "Experiências Internacionais")
card(s, 0.7, 1.7, 4.2, 4.6, "União Europeia — AI Act",
     ["Uso humanitário permitido", "Sujeito a salvaguardas específicas",
      "Autorização judicial e avaliação de impacto", "Limites temporais e geográficos"])
card(s, 5.1, 1.7, 4.2, 4.6, "INTERPOL — I-Familia",
     ["Banco global para desaparecidos", "Separação entre bases criminais e humanitárias",
      "Governança e controle institucional", "Supervisão técnica qualificada"])

# ===== SLIDE 20 — Síntese dos resultados =====
s = S[19]; clear(s)
title(s, "Síntese dos Resultados")
res = [("A tecnologia apresenta potencial relevante", "Apoio real à busca em larga escala."),
       ("Persistem limitações técnicas importantes", "Envelhecimento, baixa qualidade e viés."),
       ("Há riscos jurídicos e éticos significativos", "Privacidade, autodeterminação e discriminação."),
       ("O uso indiscriminado não é viável", "Exige governança, controle e proporcionalidade.")]
for i, (h, b) in enumerate(res):
    chip(s, i + 1, 1.6 + i * 1.18, h, b)

# ===== SLIDE 21 — Trabalhos futuros =====
s = S[20]; clear(s)
title(s, "Trabalhos Futuros")
fut = [("Base representativa da população brasileira", "Medir o viés diretamente no contexto nacional."),
       ("Identificação 1:N e múltiplos modelos", "Estender a auditoria a cenários de larga escala."),
       ("Envelhecimento facial", "Foco em casos de crianças desaparecidas."),
       ("Protocolo de homologação humanitário", "Condicionar o uso à equidade demonstrada.")]
for i, (h, b) in enumerate(fut):
    chip(s, i + 1, 1.6 + i * 1.18, h, b)

# ===== SLIDE 22 — Condições de viabilidade (pilares) =====
s = S[21]; clear(s)
title(s, "Viabilidade: Quatro Condições Essenciais")
box(s, 0.6, 1.25, 8.8, 0.6,
    [("O uso do reconhecimento facial para fins humanitários depende de:", 16, False, DARK, PP_ALIGN.LEFT)])
pil = [("Base legal", "Legislação federal específica"),
       ("Governança", "Estrutura institucional adequada"),
       ("Supervisão humana", "Decisão final sempre humana"),
       ("Auditoria", "Avaliação contínua de viés")]
w = 2.05; gap = 0.25; x0 = 0.65
for i, (lab, de) in enumerate(pil):
    pillar(s, x0 + i * (w + gap), 2.1, w, 3.6, i + 1, lab, de)

# ===== SLIDE 23 — Encerramento =====
s = S[22]; clear(s)
bar(s, 0, 0, 10, 2.0, ACCENT)
box(s, 0.6, 0.55, 8.8, 1.0, [("Obrigado!", 40, True, WHITE, PP_ALIGN.CENTER)], MSO_ANCHOR.MIDDLE)
box(s, 0.8, 2.6, 8.4, 1.6,
    [("Análise de Viabilidade Técnica e Jurídica no Uso da IA", 20, True, ACCENT, PP_ALIGN.CENTER),
     ("para Apoio à Identificação de Pessoas Desaparecidas", 20, True, ACCENT, PP_ALIGN.CENTER)],
    MSO_ANCHOR.MIDDLE)
bar(s, 3.9, 4.35, 2.2, 0.06, ACCENT2)
box(s, 0.8, 4.6, 8.4, 1.2,
    [("UNIVESP — Bacharelado em Ciência de Dados", 16, True, DARK, PP_ALIGN.CENTER),
     ("Turma 004 — Grupo 2 · 2026", 14, False, GREY, PP_ALIGN.CENTER)],
    MSO_ANCHOR.MIDDLE)
bar(s, 0, 6.9, 10, 0.6, ACCENT)

prs.save("Apresentação TCC.pptx")
print("OK — slides 1-7 e 19-23 repaginados. Total:", len(prs.slides._sldIdLst))
