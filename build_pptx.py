# -*- coding: utf-8 -*-
"""
Presentación de defensa del TFG:
"Más allá del documental de arte: ensayo fílmico y crítica anticolonial
 en Las estatuas también mueren" — Pedro José Díaz Bravo (UMU, Historia del Arte).

Estructura y convenciones replicadas de las presentaciones de referencia del
Departamento de Historia del Arte (UMU): epígrafes canónicos en mayúsculas,
título del TFG repetido en cada diapositiva, pie "Departamento de Historia
del Arte", metodología como "fuentes más destacadas".
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ----------------------------------------------------------------------------
# Paleta (tono bronce / cine / arte africano)
# ----------------------------------------------------------------------------
DARK   = RGBColor(0x1C, 0x1B, 0x19)
INK    = RGBColor(0x2A, 0x27, 0x24)
BRONZE = RGBColor(0xC0, 0x8A, 0x3E)
BROWN  = RGBColor(0x7A, 0x52, 0x36)
CREAM  = RGBColor(0xF4, 0xF0, 0xE9)
CREAM2 = RGBColor(0xEA, 0xE3, 0xD6)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
GREY   = RGBColor(0x6B, 0x64, 0x5C)

FONT_H = "Georgia"
FONT_B = "Calibri"

FULL_TITLE = ("Más allá del documental de arte: ensayo fílmico y crítica "
              "anticolonial en Las estatuas también mueren")
DEPT = "Departamento de Historia del Arte · Universidad de Murcia"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
def add_slide():
    return prs.slides.add_slide(BLANK)


def bg(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def rect(slide, x, y, w, h, color, line=None, shape=MSO_SHAPE.RECTANGLE):
    sp = slide.shapes.add_shape(shape, x, y, w, h)
    sp.fill.solid()
    sp.fill.fore_color.rgb = color
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(1)
    sp.shadow.inherit = False
    return sp


def txt(slide, x, y, w, h, lines, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
        wrap=True):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = ln.get("align", align)
        if ln.get("space_after") is not None:
            p.space_after = Pt(ln["space_after"])
        if ln.get("space_before") is not None:
            p.space_before = Pt(ln["space_before"])
        if ln.get("line_spacing") is not None:
            p.line_spacing = ln["line_spacing"]
        runs = ln["text"] if isinstance(ln["text"], list) else [ln]
        for rr in runs:
            r = p.add_run()
            r.text = rr["text"]
            f = r.font
            f.size = Pt(rr.get("size", ln.get("size", 18)))
            f.bold = rr.get("bold", ln.get("bold", False))
            f.italic = rr.get("italic", ln.get("italic", False))
            f.name = rr.get("font", ln.get("font", FONT_B))
            f.color.rgb = rr.get("color", ln.get("color", INK))
    return tb


def bullets(slide, x, y, w, h, items, size=18, color=INK, gap=10,
            lead_color=BRONZE, line_spacing=1.05):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap)
        p.line_spacing = line_spacing
        rd = p.add_run()
        rd.text = "—  "
        rd.font.size = Pt(size); rd.font.bold = True
        rd.font.name = FONT_B; rd.font.color.rgb = lead_color
        if isinstance(it, tuple):
            lead, rest = it
            r1 = p.add_run(); r1.text = lead
            r1.font.size = Pt(size); r1.font.bold = True
            r1.font.name = FONT_B; r1.font.color.rgb = INK
            r2 = p.add_run(); r2.text = rest
            r2.font.size = Pt(size); r2.font.name = FONT_B
            r2.font.color.rgb = color
        else:
            r = p.add_run(); r.text = it
            r.font.size = Pt(size); r.font.name = FONT_B
            r.font.color.rgb = color
    return tb


def running_title(slide):
    """Título completo del TFG, repetido arriba a la derecha (como en la
    presentación de referencia de los íberos)."""
    txt(slide, Inches(4.6), Inches(0.24), Inches(8.18), Inches(0.55),
        [{"text": FULL_TITLE, "size": 8.5, "color": GREY, "italic": True,
          "align": PP_ALIGN.RIGHT, "font": FONT_H, "line_spacing": 1.0}],
        anchor=MSO_ANCHOR.TOP)


def footer(slide, page):
    rect(slide, Inches(0.55), Inches(6.95), Inches(12.23), Pt(1.2), BRONZE)
    txt(slide, Inches(0.55), Inches(7.03), Inches(10.5), Inches(0.35),
        [{"text": DEPT, "size": 9, "color": GREY, "font": FONT_B}],
        anchor=MSO_ANCHOR.MIDDLE)
    txt(slide, Inches(11.5), Inches(7.03), Inches(1.28), Inches(0.35),
        [{"text": str(page), "size": 10, "color": BRONZE, "bold": True,
          "align": PP_ALIGN.RIGHT}], anchor=MSO_ANCHOR.MIDDLE)


def header(slide, section, label=None, sub=None, size=27):
    """section: epígrafe canónico (mayúsculas). label: p.ej. 'CONTENIDO · 1'."""
    running_title(slide)
    rect(slide, Inches(0.55), Inches(0.62), Inches(0.14), Inches(1.05), BRONZE)
    top = 0.58
    if label:
        txt(slide, Inches(0.85), Inches(0.55), Inches(7.0), Inches(0.32),
            [{"text": label.upper(), "size": 12, "color": BROWN, "bold": True,
              "font": FONT_B}])
        top = 0.86
    txt(slide, Inches(0.85), Inches(top), Inches(11.9), Inches(0.9),
        [{"text": section, "size": size, "color": INK, "bold": True,
          "font": FONT_H, "line_spacing": 0.98}])
    if sub:
        txt(slide, Inches(0.85), Inches(top + 0.62), Inches(11.9), Inches(0.4),
            [{"text": sub, "size": 14, "color": BROWN, "italic": True,
              "font": FONT_B}])


# ============================================================================
# 1 — PORTADA
# ============================================================================
s = add_slide()
bg(s, DARK)
rect(s, Inches(0), Inches(0), SW, Inches(0.14), BRONZE)
rect(s, Inches(0), Inches(7.36), SW, Inches(0.14), BRONZE)
rect(s, Inches(0.9), Inches(2.0), Inches(2.2), Pt(2.5), BRONZE)

txt(s, Inches(0.9), Inches(0.95), Inches(11.5), Inches(0.5),
    [{"text": "TRABAJO FIN DE GRADO  ·  GRADO EN HISTORIA DEL ARTE",
      "size": 14, "color": BRONZE, "bold": True, "font": FONT_B}])
txt(s, Inches(0.9), Inches(2.3), Inches(11.6), Inches(2.2),
    [
        {"text": "Más allá del documental de arte:", "size": 40, "color": WHITE,
         "font": FONT_H, "space_after": 6, "line_spacing": 1.02},
        {"text": "ensayo fílmico y crítica anticolonial en", "size": 25,
         "color": CREAM, "font": FONT_H, "italic": True, "space_after": 2,
         "line_spacing": 1.05},
        {"text": "Las estatuas también mueren", "size": 30, "color": BRONZE,
         "font": FONT_H, "italic": True, "bold": True},
    ])
txt(s, Inches(0.9), Inches(5.15), Inches(11.6), Inches(1.9),
    [
        {"text": "Pedro José Díaz Bravo", "size": 20, "color": WHITE,
         "bold": True, "font": FONT_B, "space_after": 9},
        {"text": [
            {"text": "Bajo la dirección de   ", "size": 13, "color": GREY},
            {"text": "Dr. José Javier Aliaga Cárceles   ·   D. José Ramón García Cañizares",
             "size": 13, "color": CREAM},
         ], "space_after": 10},
        {"text": [
            {"text": "Universidad de Murcia · Facultad de Letras", "size": 12,
             "color": CREAM},
            {"text": "      Curso 2025-2026 · Convocatoria de junio-julio",
             "size": 12, "color": GREY},
         ], "space_after": 4},
        {"text": "Departamento de Historia del Arte", "size": 11,
         "color": BRONZE, "font": FONT_B},
    ])

# ============================================================================
# 2 — ÍNDICE
# ============================================================================
s = add_slide()
bg(s, CREAM)
header(s, "Índice")

idx = [
    ("INTRODUCCIÓN", None),
    ("ESTADO DE LA CUESTIÓN Y JUSTIFICACIÓN", None),
    ("OBJETIVOS", None),
    ("METODOLOGÍA APLICADA", None),
    ("CONTENIDO", [
        "1.  El documental de arte y su institucionalización",
        "2.  El colonialismo africano y su manifiesto en el arte",
        "3.  Las estatuas también mueren (1953)",
    ]),
    ("CONCLUSIONES", None),
    ("REFERENCIAS BIBLIOGRÁFICAS Y DOCUMENTALES", None),
]
tb = s.shapes.add_textbox(Inches(1.1), Inches(1.85), Inches(11.0), Inches(5.0))
tf = tb.text_frame; tf.word_wrap = True
first = True
for head, subs in idx:
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    first = False
    p.space_after = Pt(6 if subs else 13)
    p.space_before = Pt(0)
    rb = p.add_run(); rb.text = "▪  "
    rb.font.color.rgb = BRONZE; rb.font.size = Pt(17); rb.font.bold = True
    rh = p.add_run(); rh.text = head
    rh.font.size = Pt(19); rh.font.bold = True; rh.font.name = FONT_B
    rh.font.color.rgb = INK
    if subs:
        for sline in subs:
            ps = tf.add_paragraph()
            ps.space_after = Pt(3); ps.level = 1
            rs = ps.add_run(); rs.text = "        " + sline
            rs.font.size = Pt(15); rs.font.name = FONT_B; rs.font.color.rgb = GREY
footer(s, 2)

# ============================================================================
# 3 — INTRODUCCIÓN
# ============================================================================
s = add_slide()
bg(s, CREAM)
header(s, "Introducción", label="1")
bullets(s, Inches(0.85), Inches(1.75), Inches(7.0), Inches(4.7), [
    ("El documental de arte ", "trasciende la mera documentación, catalogación y exposición de obras: es una herramienta crítica que aborda lo geográfico, lo social y lo político."),
    ("«Edad de oro» de posguerra: ", "tras la 2.ª Guerra Mundial el género apuesta por el conocimiento, la conservación y la sensibilización patrimonial."),
    ("Función educativa y ética: ", "divulgación de las obras y formación cultural de los públicos."),
    ("Caso de estudio: ", "Las estatuas también mueren (1953), de Alain Resnais y Chris Marker."),
], size=16, gap=14)

card = rect(s, Inches(8.15), Inches(1.75), Inches(4.6), Inches(4.65), DARK)
rect(s, Inches(8.15), Inches(1.75), Inches(0.14), Inches(4.65), BRONZE)
txt(s, Inches(8.5), Inches(2.0), Inches(4.0), Inches(0.5),
    [{"text": "HIPÓTESIS", "size": 14, "color": BRONZE, "bold": True,
      "font": FONT_B}])
txt(s, Inches(8.5), Inches(2.55), Inches(4.0), Inches(3.7),
    [{"text": "Las estatuas también mueren no se limita a las convenciones "
              "del documental de arte de su contexto: las trasciende para "
              "configurarse como un ensayo fílmico de carácter crítico, que "
              "articula una denuncia de la descontextualización y apropiación "
              "del arte africano en el marco del colonialismo.",
      "size": 16.5, "color": CREAM, "font": FONT_B, "italic": True,
      "line_spacing": 1.18}])
footer(s, 3)

# ============================================================================
# 4 — ESTADO DE LA CUESTIÓN Y JUSTIFICACIÓN
# ============================================================================
s = add_slide()
bg(s, CREAM)
header(s, "Estado de la cuestión y justificación", label="2",
       sub="Un objeto de estudio interdisciplinar: historia del arte y cultura visual, sociología, etnografía, historia y geopolítica.",
       size=26)

cols = [
    ("El documental de arte", [
        "C. Fernández Cuenca — 30 años de documental de arte en España (1967).",
        "J. J. Aliaga Cárceles — La revista Imágenes y la visión documental de NO-DO (2022).",
        "G. García Fernández (Peydró) — El cine sobre arte (cine-ensayo).",
    ]),
    ("África y colonialismo", [
        "R. Hamery — Les films sur l'art d'Afrique noire (1945-1961).",
        "Ceamanos & Kabunda — El Reparto de África.",
        "Sarr & Savoy — The Restitution of African Cultural Heritage (2018).",
    ]),
    ("El film y sus directores", [
        "G. García Fernández — La piel bajo las máscaras (Les Statues...).",
        "N. M. Alter — Chris Marker (2006).",
        "A. Cortijo Talavera — los documentales de Alain Resnais.",
    ]),
]
cx = 0.85; cw = Inches(3.85); gx = 0.2
for i, (h, items) in enumerate(cols):
    x = Inches(cx + i * (3.85 + gx))
    rect(s, x, Inches(2.95), cw, Inches(0.55), DARK)
    txt(s, x, Inches(2.95), cw, Inches(0.55),
        [{"text": h, "size": 14.5, "color": WHITE, "bold": True,
          "align": PP_ALIGN.CENTER, "font": FONT_B}], anchor=MSO_ANCHOR.MIDDLE)
    card = rect(s, x, Inches(3.5), cw, Inches(2.95), WHITE)
    card.line.color.rgb = CREAM2; card.line.width = Pt(1)
    bullets(s, x + Inches(0.2), Inches(3.72), cw - Inches(0.4), Inches(2.6),
            items, size=12.5, gap=11, line_spacing=1.0)
footer(s, 4)

# ============================================================================
# 5 — OBJETIVOS
# ============================================================================
s = add_slide()
bg(s, CREAM)
header(s, "Objetivos", label="3")
objs = [
    ("Analizar", " el papel de Las estatuas también mueren como posible punto de inflexión en el desarrollo histórico del documental de arte."),
    ("Examinar", " el tratamiento del patrimonio cultural africano en el cine documental y etnográfico previo, para contextualizar la propuesta del film."),
    ("Estudiar", " los recursos formales, narrativos y discursivos empleados, para determinar cómo se construye su dimensión crítica."),
    ("Reflexionar", " sobre cómo el film articula la crítica a la descontextualización y apropiación del arte africano, atendiendo a la función y el significado original de las obras."),
]
y = 1.95
for i, (lead, rest) in enumerate(objs):
    yy = Inches(y)
    rect(s, Inches(0.85), yy, Inches(0.62), Inches(0.62), BRONZE,
         shape=MSO_SHAPE.OVAL)
    txt(s, Inches(0.85), yy, Inches(0.62), Inches(0.62),
        [{"text": str(i + 1), "size": 22, "color": WHITE, "bold": True,
          "align": PP_ALIGN.CENTER, "font": FONT_H}], anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(1.7), yy - Inches(0.05), Inches(11.0), Inches(1.05),
        [{"text": [
            {"text": lead, "size": 18, "color": BROWN, "bold": True},
            {"text": rest, "size": 18, "color": INK},
         ], "line_spacing": 1.05}], anchor=MSO_ANCHOR.MIDDLE)
    y += 1.18
footer(s, 5)

# ============================================================================
# 6 — METODOLOGÍA APLICADA
# ============================================================================
s = add_slide()
bg(s, CREAM)
header(s, "Metodología aplicada", label="4",
       sub="Estudio de caso con enfoque cualitativo e interdisciplinar y análisis histórico-artístico y formal del cortometraje.",
       size=27)
txt(s, Inches(0.85), Inches(2.42), Inches(11.9), Inches(0.35),
    [{"text": "FUENTES MÁS DESTACADAS", "size": 13, "color": BROWN,
      "bold": True, "font": FONT_B}])

sources = [
    "Fernández Cuenca\n30 años de documental de arte en España (1967)",
    "Aliaga Cárceles\nLa revista Imágenes y la visión documental de NO-DO (2022)",
    "García Fernández / Peydró\nEl cine sobre arte · La piel bajo las máscaras",
    "Hamery\nLes films sur l'art d'Afrique noire (1945-1961)",
    "Ceamanos & Kabunda\nEl Reparto de África (2016)",
    "Sarr & Savoy\nThe Restitution of African Cultural Heritage (2018)",
]
cw = Inches(3.85); ch = Inches(1.62)
x0 = 0.85; y0 = 2.85; gx = 0.2; gy = 0.22
for i, sname in enumerate(sources):
    r = i // 3; c = i % 3
    x = Inches(x0 + c * (3.85 + gx))
    yy = Inches(y0 + r * (1.62 + gy))
    card = rect(s, x, yy, cw, ch, WHITE)
    card.line.color.rgb = CREAM2; card.line.width = Pt(1)
    rect(s, x, yy, Inches(0.1), ch, BRONZE)
    title, sub = sname.split("\n", 1)
    txt(s, x + Inches(0.28), yy + Inches(0.16), cw - Inches(0.45), ch - Inches(0.3),
        [
            {"text": title, "size": 14.5, "color": INK, "bold": True,
             "font": FONT_B, "space_after": 5},
            {"text": sub, "size": 12, "color": GREY, "italic": True,
             "font": FONT_B, "line_spacing": 1.0},
        ])
footer(s, 6)

# ============================================================================
# 7 — CONTENIDO 1: documental de arte e institucionalización
# ============================================================================
s = add_slide()
bg(s, CREAM)
header(s, "El documental de arte y su institucionalización",
       label="Contenido · 1", size=25)
bullets(s, Inches(0.85), Inches(1.85), Inches(7.1), Inches(4.6), [
    ("Orígenes: ", "el cine bebe del arte desde sus inicios; en los años 30 surge el sincretismo entre lo pictórico y lo filmado."),
    ("«Edad de oro» (posguerra): ", "valor multifuncional —expositivo, político y unificador— frente al cine como propaganda (NO-DO, Goebbels)."),
    ("Caso francés: ", "del cinéma de qualité hacia un cine más poético y personal; germen del documental ensayístico."),
    ("Nueva gramática: ", "travellings, variaciones de escala y montajes asociativos para crear nuevas relaciones de significado."),
], size=15.5, gap=12)

card = rect(s, Inches(8.25), Inches(1.85), Inches(4.5), Inches(4.6), DARK)
txt(s, Inches(8.55), Inches(2.05), Inches(4.0), Inches(0.5),
    [{"text": "INSTITUCIONALIZACIÓN", "size": 13, "color": BRONZE,
      "bold": True, "font": FONT_B}])
tl = [
    ("UNESCO · FIAF (1938) · ICOM (1946)", "marco institucional de la cultura."),
    ("26 jun. 1948", "I Conferencia Internacional sobre el Cine de Arte (Louvre y Musée de l'Homme) → nace la FIFA."),
    ("1950", "Francis Bolen define la «película de arte» (congreso de Bruselas)."),
    ("1952-53", "la UNESCO reconoce la FIFA, que pasa a recibir subvención."),
]
ty = 2.62
for k, v in tl:
    rect(s, Inches(8.55), Inches(ty + 0.07), Inches(0.1), Inches(0.1), BRONZE,
         shape=MSO_SHAPE.OVAL)
    txt(s, Inches(8.78), Inches(ty), Inches(3.85), Inches(1.0),
        [{"text": [
            {"text": k + "  ", "size": 13, "color": WHITE, "bold": True},
            {"text": v, "size": 12, "color": CREAM},
        ], "line_spacing": 1.02}])
    ty += 0.92
footer(s, 7)

# ============================================================================
# 8 — CONTENIDO 2: colonialismo, museos y mercado
# ============================================================================
s = add_slide()
bg(s, CREAM)
header(s, "El colonialismo africano y su manifiesto en el arte",
       label="Contenido · 2", size=24)
blocks = [
    ("Reparto de África", [
        "Conferencia de Berlín (1884-1885): reparto del continente con fines comerciales.",
        "Expolio material, artístico, ritual e identitario de innumerables pueblos.",
    ]),
    ("Del Trocadero al Quai Branly", [
        "Museo del Trocadero (1892) y los «zoos humanos».",
        "Paul Rivet inaugura el Musée de l'Homme (1938): «museo-laboratorio».",
        "Quai Branly (2006); restituciones (Sarr & Savoy 2018; Macron).",
    ]),
    ("Mercado del arte africano", [
        "Tourist art / airport art (Jules-Rosette): obras adaptadas al gusto europeo.",
        "La «autenticidad» como valor de mercado.",
        "El ego-system (Forni & Steiner): herencia del colonialismo.",
    ]),
]
cx = 0.85; cw = Inches(3.85); gx = 0.2
for i, (h, items) in enumerate(blocks):
    x = Inches(cx + i * (3.85 + gx))
    rect(s, x, Inches(1.95), cw, Inches(0.55), BROWN)
    txt(s, x, Inches(1.95), cw, Inches(0.55),
        [{"text": h, "size": 14, "color": WHITE, "bold": True,
          "align": PP_ALIGN.CENTER, "font": FONT_B}], anchor=MSO_ANCHOR.MIDDLE)
    card = rect(s, x, Inches(2.5), cw, Inches(3.95), WHITE)
    card.line.color.rgb = CREAM2; card.line.width = Pt(1)
    bullets(s, x + Inches(0.22), Inches(2.72), cw - Inches(0.44), Inches(3.6),
            items, size=12.5, gap=12, line_spacing=1.02)
footer(s, 8)

# ============================================================================
# 9 — CONTENIDO 2 (cont.): tratamiento cinematográfico de África
# ============================================================================
s = add_slide()
bg(s, CREAM)
header(s, "Tratamiento cinematográfico de África",
       label="Contenido · 2", sub="Tres tipologías de documental sobre el arte africano.",
       size=26)
types = [
    ("Documental\npropagandístico colonial", BROWN, [
        "Discurso reaccionario y paternalista.",
        "Exalta la «benevolencia» colonial y lo «exótico».",
        "El creador africano queda reducido a artesano.",
    ]),
    ("Documental\netnográfico", BRONZE, [
        "Busca conciliar la mirada occidental con el otro.",
        "Pone en valor danza y rituales.",
        "Jean Rouch: «cine directo» → cinéma vérité.",
    ]),
    ("Documental\nde arte (puro)", DARK, [
        "Espacio intermedio, enfoque museístico.",
        "Centrado en las artes visuales.",
        "Obras descontextualizadas de su función ritual.",
    ]),
]
cx = 0.85; cw = Inches(3.85); gx = 0.2
for i, (h, col, items) in enumerate(types):
    x = Inches(cx + i * (3.85 + gx))
    rect(s, x, Inches(2.65), cw, Inches(1.05), col)
    txt(s, x + Inches(0.15), Inches(2.65), cw - Inches(0.3), Inches(1.05),
        [{"text": h, "size": 16, "color": WHITE, "bold": True,
          "align": PP_ALIGN.CENTER, "font": FONT_H, "line_spacing": 1.0}],
        anchor=MSO_ANCHOR.MIDDLE)
    card = rect(s, x, Inches(3.7), cw, Inches(2.75), WHITE)
    card.line.color.rgb = CREAM2; card.line.width = Pt(1)
    bullets(s, x + Inches(0.22), Inches(3.92), cw - Inches(0.44), Inches(2.4),
            items, size=13, gap=11, line_spacing=1.02)
footer(s, 9)

# ============================================================================
# 10 — CONTENIDO 3: Las estatuas también mueren (directores y contexto)
# ============================================================================
s = add_slide()
bg(s, CREAM)
header(s, "Las estatuas también mueren (1953)", label="Contenido · 3",
       size=27)
bullets(s, Inches(0.85), Inches(1.75), Inches(7.0), Inches(4.7), [
    ("Directores: ", "Chris Marker y Alain Resnais, del grupo de la Rive Gauche (junto a Agnès Varda)."),
    ("Encargo: ", "del intelectual senegalés Alioune Diop, a través de la revista Présence Africaine (1950)."),
    ("Ensayo cinematográfico: ", "C. Lupton lo describe como «un ejemplo magistral de ensayo cinematográfico»."),
    ("Censura: ", "prohibido en Francia durante 12 años por su lectura anticolonial."),
    ("Vigencia: ", "su huella llega a Dahomey (Mati Diop) y al debate sobre la restitución (solo 26 de +7000 piezas)."),
], size=15, gap=11)

def dir_card(y, name, sub, body):
    card = rect(s, Inches(8.2), Inches(y), Inches(4.55), Inches(2.15), WHITE)
    card.line.color.rgb = CREAM2; card.line.width = Pt(1)
    rect(s, Inches(8.2), Inches(y), Inches(0.1), Inches(2.15), BRONZE)
    txt(s, Inches(8.48), Inches(y + 0.16), Inches(4.1), Inches(1.9),
        [
            {"text": name, "size": 17, "color": INK, "bold": True,
             "font": FONT_B, "space_after": 1},
            {"text": sub, "size": 11.5, "color": BROWN, "italic": True,
             "font": FONT_B, "space_after": 6},
            {"text": body, "size": 12.5, "color": GREY, "font": FONT_B,
             "line_spacing": 1.05},
        ])

dir_card(1.75, "Chris Marker", "Documental subjetivo · film-essay",
         "«Soy un ensayista... y a mí hacer ensayos.» Comprometido con los ideales de Présence Africaine.")
dir_card(4.3, "Alain Resnais", "Rive Gauche · documental de arte",
         "Óscar por Van Gogh (1950). Aporta composiciones dinámicas, travellings y montaje para explorar las obras.")
footer(s, 10)

# ============================================================================
# 11 — CONTENIDO 3: contenido del film (dos bloques)
# ============================================================================
s = add_slide()
bg(s, CREAM)
header(s, "Contenido del film: dos bloques de obras", label="Contenido · 3",
       size=26)
rect(s, Inches(0.85), Inches(2.0), Inches(11.9), Inches(0.68), DARK)
txt(s, Inches(1.1), Inches(2.0), Inches(11.4), Inches(0.68),
    [{"text": "«Cuando los hombres mueren, se vuelven historia. "
              "Cuando las estatuas mueren, se vuelven arte.»",
      "size": 16, "color": CREAM, "italic": True, "font": FONT_H,
      "align": PP_ALIGN.CENTER}], anchor=MSO_ANCHOR.MIDDLE)

rect(s, Inches(0.85), Inches(2.92), Inches(5.85), Inches(0.55), BRONZE)
txt(s, Inches(0.85), Inches(2.92), Inches(5.85), Inches(0.55),
    [{"text": "Bloque I · Realeza y poder", "size": 15, "color": WHITE,
      "bold": True, "align": PP_ALIGN.CENTER, "font": FONT_B}],
    anchor=MSO_ANCHOR.MIDDLE)
bullets(s, Inches(1.05), Inches(3.62), Inches(5.5), Inches(2.8), [
    "Cabezas de oba (Reino de Benín) y de oni (Reino de Ife), Nigeria.",
    "Guerreros y leopardos de bronce de Benín (British Museum).",
    "Montaje rítmico y fundidos: ¿colonizó Francia un pueblo «sin cultura»?",
], size=13.5, gap=12)

rect(s, Inches(7.0), Inches(2.92), Inches(5.75), Inches(0.55), BROWN)
txt(s, Inches(7.0), Inches(2.92), Inches(5.75), Inches(0.55),
    [{"text": "Bloque II · Culto y ritual", "size": 15, "color": WHITE,
      "bold": True, "align": PP_ALIGN.CENTER, "font": FONT_B}],
    anchor=MSO_ANCHOR.MIDDLE)
bullets(s, Inches(7.2), Inches(3.62), Inches(5.4), Inches(2.8), [
    "Relicarios Kota (Gabón); tableros de adivinación Ifá, Yoruba (Nigeria).",
    "Reposacabezas Pende (R. D. del Congo); máscaras Dan (Costa de Marfil).",
    "Culto animista y a los ancestros; piezas deliberadamente sin contexto.",
], size=13.5, gap=12)
footer(s, 11)

# ============================================================================
# 12 — CONTENIDO 3: gramática fílmica y problemática de género
# ============================================================================
s = add_slide()
bg(s, CREAM)
header(s, "Forma fílmica y problemática de género", label="Contenido · 3",
       size=26)
cols = [
    ("Fotografía", "Ghislain Cloquet expone las piezas sobre fondos neutros: detalle y, a la vez, refuerzo de su descontextualización."),
    ("Música", "Guy Bernard alterna lo familiar (europeo-americano) y lo «exótico-ritual», con momentos sincréticos."),
    ("Montaje", "Material de archivo (influencia de Esfir Shub): contrasta las «bondades» coloniales con su hipocresía."),
]
cx = 0.85; cw = Inches(3.85); gx = 0.2
for i, (h, body) in enumerate(cols):
    x = Inches(cx + i * (3.85 + gx))
    card = rect(s, x, Inches(1.9), cw, Inches(2.5), WHITE)
    card.line.color.rgb = CREAM2; card.line.width = Pt(1)
    rect(s, x, Inches(1.9), cw, Inches(0.5), DARK)
    txt(s, x, Inches(1.9), cw, Inches(0.5),
        [{"text": h, "size": 15, "color": BRONZE, "bold": True,
          "align": PP_ALIGN.CENTER, "font": FONT_B}], anchor=MSO_ANCHOR.MIDDLE)
    txt(s, x + Inches(0.25), Inches(2.6), cw - Inches(0.5), Inches(1.7),
        [{"text": body, "size": 13, "color": INK, "font": FONT_B,
          "line_spacing": 1.08}])

rect(s, Inches(0.85), Inches(4.7), Inches(11.9), Inches(1.65), BROWN)
rect(s, Inches(0.85), Inches(4.7), Inches(0.14), Inches(1.65), BRONZE)
txt(s, Inches(1.2), Inches(4.85), Inches(11.3), Inches(1.4),
    [
        {"text": "¿Documental de arte o ensayo fílmico?", "size": 16,
         "color": WHITE, "bold": True, "font": FONT_B, "space_after": 4},
        {"text": "Aunque por su forma parezca un documental de arte, Resnais "
                 "y Marker lo exceden: la Rive Gauche desarrolla el ciné-essai, "
                 "que aborda con ironía y firmeza la problemática colonial y "
                 "la mirada racista aún imperante.",
         "size": 13.5, "color": CREAM, "font": FONT_B, "line_spacing": 1.1},
    ])
footer(s, 12)

# ============================================================================
# 13 — CONCLUSIONES
# ============================================================================
s = add_slide()
bg(s, CREAM)
header(s, "Conclusiones", label="5")
bullets(s, Inches(0.85), Inches(1.75), Inches(7.2), Inches(4.7), [
    ("Punto de inflexión: ", "el film supera los estándares del documental de arte y abre debate sobre identidad, poder colonial y pérdida de significados."),
    ("Tratamiento ontológico: ", "frente al cine etnográfico-expositivo (p. ej. Mangbetu, 1954), da voz a las historias reprimidas tras cada pieza."),
    ("Ensayo fílmico-crítico: ", "se confirma la hipótesis; la obra se alinea con la ensayística más que con la mera exposición."),
    ("Preguntas abiertas: ", "¿devolverá Occidente el patrimonio expoliado? ¿en qué condiciones? ¿está África preparada para conservarlo?"),
], size=15, gap=12)

card = rect(s, Inches(8.3), Inches(1.75), Inches(4.45), Inches(4.65), DARK)
rect(s, Inches(8.3), Inches(1.75), Inches(4.45), Inches(0.14), BRONZE)
txt(s, Inches(8.65), Inches(2.2), Inches(3.8), Inches(4.0),
    [{"text": "«Cuando el documental de arte abandona la mera descripción "
              "para convertirse en una herramienta de pensamiento crítico, "
              "deja de explicar el mundo para empezar a cuestionarlo.»",
      "size": 17, "color": CREAM, "italic": True, "font": FONT_H,
      "line_spacing": 1.2}], anchor=MSO_ANCHOR.MIDDLE)
footer(s, 13)

# ============================================================================
# 14 — REFERENCIAS BIBLIOGRÁFICAS Y DOCUMENTALES
# ============================================================================
s = add_slide()
bg(s, CREAM)
header(s, "Referencias bibliográficas y documentales", label="6", size=26)
refs_l = [
    "Aliaga Cárceles, J. J. La revista Imágenes y la visión documental de NO-DO. Tesis doctoral, Univ. di Bologna, 2022.",
    "Alter, N. M. Chris Marker. Urbana: University of Illinois Press, 2006.",
    "Ceamanos, R. y Kabunda, M. El Reparto de África. Madrid: Los libros de la Catarata, 2016.",
    "Cortijo Talavera, A. «Los documentales restitutivos y poéticos de Alain Resnais». IC, 21, 2024.",
    "Fernández Cuenca, C. 30 años de documental de arte en España. Madrid: EOC, 1967.",
    "Forni, S. y Steiner, C. B. «The African Art Market as Ego-System». Critical Interventions, 12, 2018.",
    "Fraiture, P.-Ph. «Statues Also Die». Journal of French and Francophone Philosophy, XXIV, 1, 2016.",
]
refs_r = [
    "García Fernández, G. «La piel bajo las máscaras... Les Statues meurent aussi». Mirando a Clío, 2010.",
    "García Peydró, G. El cine sobre arte. Santander: Shangrila, 2019.",
    "Hamery, R. «Les films sur l'art d'Afrique noire (1945-1961)», en Le film sur l'art. Rennes: PUR, 2015.",
    "Lupton, C. Chris Marker, memories of the future. Londres: Reaktion Books, 2005.",
    "Sarr, F. y Savoy, B. The Restitution of African Cultural Heritage. París, 2018.",
    "Schäuble, M. «The Adventure of the Real. Jean Rouch...». Zeitschrift für Ethnologie, 136, 2011.",
    "Robert, V., Le Forestier, L. y Albera, F. Le film sur l'art. Rennes: PUR, 2015.",
]


def ref_col(x, items):
    tb = s.shapes.add_textbox(x, Inches(1.95), Inches(5.85), Inches(4.9))
    tf = tb.text_frame; tf.word_wrap = True
    for i, t in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(9); p.line_spacing = 1.0
        rd = p.add_run(); rd.text = "•  "
        rd.font.color.rgb = BRONZE; rd.font.bold = True; rd.font.size = Pt(11)
        r = p.add_run(); r.text = t
        r.font.size = Pt(10.5); r.font.name = FONT_B; r.font.color.rgb = INK


ref_col(Inches(0.85), refs_l)
ref_col(Inches(7.0), refs_r)
footer(s, 14)

# ============================================================================
# 15 — CIERRE
# ============================================================================
s = add_slide()
bg(s, DARK)
rect(s, Inches(0), Inches(0), SW, Inches(0.14), BRONZE)
rect(s, Inches(0), Inches(7.36), SW, Inches(0.14), BRONZE)
txt(s, Inches(1.0), Inches(2.65), Inches(11.3), Inches(1.2),
    [{"text": "Gracias por su atención", "size": 40, "color": WHITE,
      "font": FONT_H, "align": PP_ALIGN.CENTER}])
rect(s, Inches(5.66), Inches(4.0), Inches(2.0), Pt(2.5), BRONZE)
txt(s, Inches(1.0), Inches(4.25), Inches(11.3), Inches(1.5),
    [
        {"text": "Pedro José Díaz Bravo", "size": 18, "color": BRONZE,
         "bold": True, "align": PP_ALIGN.CENTER, "font": FONT_B,
         "space_after": 4},
        {"text": FULL_TITLE, "size": 13, "color": CREAM, "italic": True,
         "align": PP_ALIGN.CENTER, "font": FONT_H, "line_spacing": 1.1,
         "space_after": 6},
        {"text": "Grado en Historia del Arte · Universidad de Murcia · 2025-2026",
         "size": 11, "color": GREY, "align": PP_ALIGN.CENTER, "font": FONT_B},
    ])

out = "Defensa_TFG_Las_estatuas_tambien_mueren.pptx"
prs.save(out)
print("Guardado:", out, "·", len(prs.slides._sldIdLst), "diapositivas")
