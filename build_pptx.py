# -*- coding: utf-8 -*-
"""
Genera la presentación de defensa del TFG:
"Más allá del documental de arte: ensayo fílmico y crítica anticolonial
 en Las estatuas también mueren" — Pedro José Díaz Bravo (UMU, Historia del Arte).
Estructura y estilo basados en las presentaciones de referencia del
Departamento de Historia del Arte (UMU).
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ----------------------------------------------------------------------------
# Paleta de color (tono bronce / cine / arte africano)
# ----------------------------------------------------------------------------
DARK     = RGBColor(0x1C, 0x1B, 0x19)   # casi negro cálido
INK      = RGBColor(0x2A, 0x27, 0x24)   # texto principal
BRONZE   = RGBColor(0xC0, 0x8A, 0x3E)   # ocre / bronce (acento)
BROWN    = RGBColor(0x7A, 0x52, 0x36)   # marrón
CREAM    = RGBColor(0xF4, 0xF0, 0xE9)   # fondo cálido claro
CREAM2   = RGBColor(0xEA, 0xE3, 0xD6)   # crema más saturada
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
GREY     = RGBColor(0x6B, 0x64, 0x5C)   # texto secundario

FONT_H = "Georgia"          # titulares (serif, tono académico)
FONT_B = "Calibri"          # cuerpo

SHORT_TITLE = "Más allá del documental de arte · Las estatuas también mueren"
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
    """lines: list of dicts {text, size, color, bold, italic, font, space_after,
    space_before, level, bullet}"""
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
    """items: list of (text) or (lead, text) tuples. Custom dash bullet."""
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
        # dash marker
        rd = p.add_run()
        rd.text = "—  "
        rd.font.size = Pt(size)
        rd.font.bold = True
        rd.font.name = FONT_B
        rd.font.color.rgb = lead_color
        if isinstance(it, tuple):
            lead, rest = it
            r1 = p.add_run()
            r1.text = lead
            r1.font.size = Pt(size)
            r1.font.bold = True
            r1.font.name = FONT_B
            r1.font.color.rgb = INK
            r2 = p.add_run()
            r2.text = rest
            r2.font.size = Pt(size)
            r2.font.name = FONT_B
            r2.font.color.rgb = color
        else:
            r = p.add_run()
            r.text = it
            r.font.size = Pt(size)
            r.font.name = FONT_B
            r.font.color.rgb = color
    return tb


def footer(slide, page):
    # thin bronze rule + footer text + page number
    rect(slide, Inches(0.55), Inches(6.92), Inches(12.23), Pt(1.4), BRONZE)
    txt(slide, Inches(0.55), Inches(7.0), Inches(10.5), Inches(0.4),
        [{"text": DEPT, "size": 9, "color": GREY, "font": FONT_B}],
        anchor=MSO_ANCHOR.MIDDLE)
    txt(slide, Inches(11.5), Inches(7.0), Inches(1.28), Inches(0.4),
        [{"text": str(page), "size": 10, "color": BRONZE, "bold": True,
          "align": PP_ALIGN.RIGHT}],
        anchor=MSO_ANCHOR.MIDDLE)


def header(slide, kicker, title, title_size=30):
    """Standard content-slide header with bronze sidebar + kicker + title."""
    # left accent bar
    rect(slide, Inches(0.55), Inches(0.6), Inches(0.14), Inches(1.15), BRONZE)
    txt(slide, Inches(0.85), Inches(0.55), Inches(11.7), Inches(0.35),
        [{"text": kicker.upper(), "size": 12, "color": BROWN, "bold": True,
          "font": FONT_B}])
    txt(slide, Inches(0.85), Inches(0.86), Inches(11.9), Inches(0.95),
        [{"text": title, "size": title_size, "color": INK, "bold": False,
          "font": FONT_H, "line_spacing": 1.0}])


def image_placeholder(slide, x, y, w, h, caption):
    """A tasteful labelled frame so the student knows where to drop figures."""
    card = rect(slide, x, y, w, h, CREAM2)
    card.line.color.rgb = BRONZE
    card.line.width = Pt(0.75)
    txt(slide, x, y + Emu(int(h*0.30)), w, Inches(0.4),
        [{"text": "▣  IMAGEN", "size": 12, "color": BROWN, "bold": True,
          "align": PP_ALIGN.CENTER, "font": FONT_B}])
    txt(slide, x + Inches(0.1), y + Emu(int(h*0.52)), w - Inches(0.2),
        Emu(int(h*0.42)),
        [{"text": caption, "size": 10.5, "color": GREY, "italic": True,
          "align": PP_ALIGN.CENTER, "font": FONT_B, "line_spacing": 0.95}],
        anchor=MSO_ANCHOR.TOP)
    return card


# ============================================================================
# SLIDE 1 — PORTADA
# ============================================================================
s = add_slide()
bg(s, DARK)
# bronze framing lines
rect(s, Inches(0), Inches(0), SW, Inches(0.14), BRONZE)
rect(s, Inches(0), Inches(7.36), SW, Inches(0.14), BRONZE)
rect(s, Inches(0.9), Inches(2.05), Inches(2.2), Pt(2.5), BRONZE)

txt(s, Inches(0.9), Inches(1.0), Inches(11.5), Inches(0.5),
    [{"text": "TRABAJO FIN DE GRADO  ·  GRADO EN HISTORIA DEL ARTE",
      "size": 14, "color": BRONZE, "bold": True, "font": FONT_B}])

txt(s, Inches(0.9), Inches(2.35), Inches(11.6), Inches(2.2),
    [
        {"text": "Más allá del documental de arte:", "size": 40,
         "color": WHITE, "font": FONT_H, "bold": False, "space_after": 6,
         "line_spacing": 1.02},
        {"text": "ensayo fílmico y crítica anticolonial en", "size": 26,
         "color": CREAM, "font": FONT_H, "italic": True, "space_after": 2,
         "line_spacing": 1.05},
        {"text": "Las estatuas también mueren", "size": 30,
         "color": BRONZE, "font": FONT_H, "italic": True, "bold": True},
    ])

txt(s, Inches(0.9), Inches(5.25), Inches(11.6), Inches(1.7),
    [
        {"text": "Pedro José Díaz Bravo", "size": 20, "color": WHITE,
         "bold": True, "font": FONT_B, "space_after": 8},
        {"text": [
            {"text": "Bajo la dirección de  ", "size": 13, "color": GREY},
            {"text": "Dr. José Javier Aliaga Cárceles  ·  D. José Ramón García Cañizares",
             "size": 13, "color": CREAM},
         ], "space_after": 10},
        {"text": [
            {"text": "Universidad de Murcia · Facultad de Letras", "size": 12,
             "color": CREAM},
            {"text": "      Curso 2025-2026 · Convocatoria de junio-julio",
             "size": 12, "color": GREY},
         ]},
    ])

# ============================================================================
# SLIDE 2 — ÍNDICE
# ============================================================================
s = add_slide()
bg(s, CREAM)
header(s, "Estructura de la presentación", "Índice")
left = [
    ("1.", "Introducción, justificación e hipótesis"),
    ("2.", "Estado de la cuestión"),
    ("3.", "Objetivos"),
    ("4.", "Metodología y fuentes"),
]
right = [
    ("5.", "El documental de arte y su institucionalización"),
    ("6.", "Colonialismo, museos y mercado del arte africano"),
    ("7.", "Las estatuas también mueren (1953)"),
    ("8.", "Conclusiones"),
]


def index_col(x, items):
    tb = s.shapes.add_textbox(x, Inches(2.25), Inches(5.7), Inches(4.2))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, (n, t) in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(20)
        rn = p.add_run(); rn.text = n + "  "
        rn.font.size = Pt(22); rn.font.bold = True; rn.font.name = FONT_H
        rn.font.color.rgb = BRONZE
        rt = p.add_run(); rt.text = t
        rt.font.size = Pt(18); rt.font.name = FONT_B; rt.font.color.rgb = INK


index_col(Inches(0.85), left)
index_col(Inches(7.0), right)
footer(s, 2)

# ============================================================================
# SLIDE 3 — INTRODUCCIÓN / JUSTIFICACIÓN / HIPÓTESIS
# ============================================================================
s = add_slide()
bg(s, CREAM)
header(s, "Introducción y justificación", "Punto de partida e hipótesis")

bullets(s, Inches(0.85), Inches(2.05), Inches(7.0), Inches(4.4), [
    ("El documental de arte ", "trasciende la mera documentación, catalogación y exposición de obras."),
    ("Tras la Segunda Guerra Mundial ", "vive una auténtica «edad de oro»: conocimiento, conservación y sensibilización patrimonial."),
    ("Función educativa y ética ", "del género: divulgación y formación cultural de los públicos."),
    ("Caso de estudio: ", "Las estatuas también mueren (1953), de Alain Resnais y Chris Marker."),
], size=17, gap=14)

# hypothesis card
card = rect(s, Inches(8.15), Inches(2.05), Inches(4.6), Inches(4.35), DARK)
rect(s, Inches(8.15), Inches(2.05), Inches(0.14), Inches(4.35), BRONZE)
txt(s, Inches(8.5), Inches(2.3), Inches(4.0), Inches(0.5),
    [{"text": "HIPÓTESIS", "size": 14, "color": BRONZE, "bold": True,
      "font": FONT_B}])
txt(s, Inches(8.5), Inches(2.85), Inches(4.0), Inches(3.4),
    [{"text": "El film no se limita a las convenciones del documental de "
              "arte de su tiempo: las trasciende para configurarse como un "
              "ensayo fílmico de carácter crítico, que denuncia la "
              "descontextualización y apropiación del arte africano en el "
              "marco del colonialismo.",
      "size": 16, "color": CREAM, "font": FONT_B, "italic": True,
      "line_spacing": 1.15}])
footer(s, 3)

# ============================================================================
# SLIDE 4 — ESTADO DE LA CUESTIÓN
# ============================================================================
s = add_slide()
bg(s, CREAM)
header(s, "Estado de la cuestión", "Un tema interdisciplinar")

txt(s, Inches(0.85), Inches(1.95), Inches(11.9), Inches(0.6),
    [{"text": "Historia del arte y cultura visual en diálogo con la sociología, "
              "la etnografía, la historia y la geopolítica.",
      "size": 16, "color": GREY, "italic": True, "font": FONT_B}])

cols = [
    ("Documental de arte", [
        "C. Fernández Cuenca — 30 años de documental de arte en España (1967)",
        "J. J. Aliaga Cárceles — tesis sobre Imágenes y NO-DO (2022)",
        "G. García Fernández / Peydró — El cine sobre arte y el cine-ensayo",
    ]),
    ("África y colonialismo", [
        "R. Hamery — Les films sur l'art d'Afrique noire (1945-1961)",
        "Sarr & Savoy — The Restitution of African Cultural Heritage (2018)",
        "Ceamanos & Kabunda — El Reparto de África",
    ]),
    ("El film y sus autores", [
        "G. García Fernández — La piel bajo las máscaras (Les Statues...)",
        "N. M. Alter — Chris Marker (2006)",
        "A. Cortijo Talavera — los documentales de Alain Resnais",
    ]),
]
cx = Inches(0.85)
cw = Inches(3.85)
gap = Inches(0.2)
for i, (h, items) in enumerate(cols):
    x = Emu(cx + i * (cw + gap))
    rect(s, x, Inches(2.65), cw, Inches(0.55), DARK)
    txt(s, x, Inches(2.65), cw, Inches(0.55),
        [{"text": h, "size": 15, "color": WHITE, "bold": True,
          "align": PP_ALIGN.CENTER, "font": FONT_B}], anchor=MSO_ANCHOR.MIDDLE)
    card = rect(s, x, Inches(3.2), cw, Inches(3.25), WHITE)
    card.line.color.rgb = CREAM2; card.line.width = Pt(1)
    bullets(s, x + Inches(0.2), Inches(3.45), cw - Inches(0.4), Inches(2.8),
            items, size=12.5, gap=12, line_spacing=1.0)
footer(s, 4)

# ============================================================================
# SLIDE 5 — OBJETIVOS
# ============================================================================
s = add_slide()
bg(s, CREAM)
header(s, "Objetivos de la investigación", "Objetivos")

objs = [
    ("Analizar", " el papel de Las estatuas también mueren como posible punto de inflexión en el desarrollo histórico del documental de arte."),
    ("Examinar", " el tratamiento del patrimonio cultural africano en el cine documental y etnográfico previo, para contextualizar la propuesta del film."),
    ("Estudiar", " los recursos formales, narrativos y discursivos con los que se construye su dimensión crítica."),
    ("Reflexionar", " sobre la descontextualización y apropiación del arte africano: función original de las obras frente a su exhibición en museos occidentales."),
]
y = 2.1
for i, (lead, rest) in enumerate(objs):
    yy = Inches(y)
    circle = rect(s, Inches(0.85), yy, Inches(0.62), Inches(0.62), BRONZE,
                  shape=MSO_SHAPE.OVAL)
    txt(s, Inches(0.85), yy, Inches(0.62), Inches(0.62),
        [{"text": str(i + 1), "size": 22, "color": WHITE, "bold": True,
          "align": PP_ALIGN.CENTER, "font": FONT_H}], anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(1.7), yy - Inches(0.04), Inches(11.0), Inches(1.0),
        [{"text": [
            {"text": lead, "size": 18, "color": BROWN, "bold": True},
            {"text": rest, "size": 18, "color": INK},
         ], "line_spacing": 1.05}], anchor=MSO_ANCHOR.MIDDLE)
    y += 1.15
footer(s, 5)

# ============================================================================
# SLIDE 6 — METODOLOGÍA
# ============================================================================
s = add_slide()
bg(s, CREAM)
header(s, "Metodología aplicada", "Enfoque y fuentes principales")

txt(s, Inches(0.85), Inches(1.95), Inches(11.9), Inches(0.55),
    [{"text": "Análisis histórico-artístico y fílmico de un caso de estudio, "
              "con lectura crítica de fuentes interdisciplinares y análisis "
              "formal del cortometraje.",
      "size": 16, "color": GREY, "italic": True, "font": FONT_B,
      "line_spacing": 1.05}])

sources = [
    "Fernández Cuenca\n30 años de documental de arte en España (1967)",
    "Aliaga Cárceles\nLa revista Imágenes y la visión documental de NO-DO (2022)",
    "García Fernández\nEl cine sobre arte / La piel bajo las máscaras",
    "Hamery\nLes films sur l'art d'Afrique noire (1945-1961)",
    "Sarr & Savoy\nThe Restitution of African Cultural Heritage (2018)",
    "Alter / Lupton\nChris Marker — monografías de referencia",
]
cw = Inches(3.85); ch = Inches(1.7)
x0 = 0.85; y0 = 2.75; gx = 0.2; gy = 0.25
for i, sname in enumerate(sources):
    r = i // 3; c = i % 3
    x = Inches(x0 + c * (3.85 + gx))
    yy = Inches(y0 + r * (1.7 + gy / 1))
    card = rect(s, x, yy, cw, ch, WHITE)
    card.line.color.rgb = CREAM2; card.line.width = Pt(1)
    rect(s, x, yy, Inches(0.1), ch, BRONZE)
    title, sub = sname.split("\n", 1)
    txt(s, x + Inches(0.28), yy + Inches(0.18), cw - Inches(0.45), ch - Inches(0.3),
        [
            {"text": title, "size": 15, "color": INK, "bold": True,
             "font": FONT_B, "space_after": 5},
            {"text": sub, "size": 12.5, "color": GREY, "italic": True,
             "font": FONT_B, "line_spacing": 1.0},
        ])
footer(s, 6)

# ============================================================================
# SLIDE 7 — CONTENIDO: documental de arte e institucionalización
# ============================================================================
s = add_slide()
bg(s, CREAM)
header(s, "Contenido · 1", "El documental de arte y su institucionalización")

bullets(s, Inches(0.85), Inches(2.0), Inches(7.1), Inches(4.4), [
    ("Orígenes: ", "el cine bebe del arte desde sus inicios; en los años 30 surge el sincretismo entre lo pictórico y lo filmado."),
    ("«Edad de oro» (posguerra): ", "valor multifuncional —expositivo, político y unificador— frente al cine como propaganda (NO-DO, Goebbels)."),
    ("Caso francés: ", "del cinéma de qualité hacia un cine más poético y personal; semilla del documental ensayístico."),
    ("Nueva gramática: ", "travellings, variaciones de escala y montajes asociativos para crear nuevas relaciones de significado."),
], size=16, gap=13)

# timeline card
card = rect(s, Inches(8.25), Inches(2.0), Inches(4.5), Inches(4.45), DARK)
txt(s, Inches(8.55), Inches(2.2), Inches(4.0), Inches(0.5),
    [{"text": "INSTITUCIONALIZACIÓN", "size": 13, "color": BRONZE,
      "bold": True, "font": FONT_B}])
tl = [
    ("UNESCO · FIAF · ICOM", "marco institucional de la cultura"),
    ("26 jun. 1948", "I Conferencia Internacional sobre el Cine de Arte (Louvre y Musée de l'Homme) → nace la FIFA"),
    ("1950", "Bolen define la «película de arte» (Bruselas)"),
    ("1952-53", "la UNESCO reconoce la FIFA; subvención de sus actividades"),
]
ty = 2.75
for k, v in tl:
    rect(s, Inches(8.55), Inches(ty + 0.06), Inches(0.1), Inches(0.1), BRONZE,
         shape=MSO_SHAPE.OVAL)
    txt(s, Inches(8.78), Inches(ty), Inches(3.85), Inches(0.9),
        [{"text": [
            {"text": k + "  ", "size": 13.5, "color": WHITE, "bold": True},
            {"text": v, "size": 12.5, "color": CREAM},
        ], "line_spacing": 1.0}])
    ty += 0.9
footer(s, 7)

# ============================================================================
# SLIDE 8 — CONTENIDO: colonialismo, museos y mercado
# ============================================================================
s = add_slide()
bg(s, CREAM)
header(s, "Contenido · 2", "Colonialismo, museos y mercado del arte africano")

blocks = [
    ("Reparto de África", [
        "Conferencia de Berlín (1884-1885): reparto del continente con fines comerciales.",
        "Expolio material, artístico, ritual e identitario de innumerables pueblos.",
    ]),
    ("Del Trocadero al Quai Branly", [
        "Museo del Trocadero (1892) y los «zoos humanos» etnocéntricos.",
        "Paul Rivet funda el Musée de l'Homme (1938): «museo-laboratorio».",
        "Quai Branly (2006); restituciones (Sarr & Savoy 2018; Macron).",
    ]),
    ("Mercado del arte africano", [
        "Tourist / airport art (Jules-Rosette): obras adaptadas al gusto europeo.",
        "La «autenticidad» como valor de mercado.",
        "El ego-system (Forni & Steiner): herencia desigual del colonialismo.",
    ]),
]
cx = 0.85; cw = Inches(3.85); gx = 0.2
for i, (h, items) in enumerate(blocks):
    x = Inches(cx + i * (3.85 + gx))
    rect(s, x, Inches(2.0), cw, Inches(0.55), BROWN)
    txt(s, x, Inches(2.0), cw, Inches(0.55),
        [{"text": h, "size": 14.5, "color": WHITE, "bold": True,
          "align": PP_ALIGN.CENTER, "font": FONT_B}], anchor=MSO_ANCHOR.MIDDLE)
    card = rect(s, x, Inches(2.55), cw, Inches(3.9), WHITE)
    card.line.color.rgb = CREAM2; card.line.width = Pt(1)
    bullets(s, x + Inches(0.22), Inches(2.78), cw - Inches(0.44), Inches(3.5),
            items, size=13, gap=12, line_spacing=1.02)
footer(s, 8)

# ============================================================================
# SLIDE 9 — CONTENIDO: tres tipologías de documental sobre África
# ============================================================================
s = add_slide()
bg(s, CREAM)
header(s, "Contenido · 3", "Tres miradas: el cine documental sobre África")

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
        "Obras descontextualizadas de su función.",
    ]),
]
cx = 0.85; cw = Inches(3.85); gx = 0.2
for i, (h, col, items) in enumerate(types):
    x = Inches(cx + i * (3.85 + gx))
    rect(s, x, Inches(2.05), cw, Inches(1.05), col)
    txt(s, x + Inches(0.15), Inches(2.05), cw - Inches(0.3), Inches(1.05),
        [{"text": h, "size": 16, "color": WHITE, "bold": True,
          "align": PP_ALIGN.CENTER, "font": FONT_H, "line_spacing": 1.0}],
        anchor=MSO_ANCHOR.MIDDLE)
    card = rect(s, x, Inches(3.1), cw, Inches(3.35), WHITE)
    card.line.color.rgb = CREAM2; card.line.width = Pt(1)
    bullets(s, x + Inches(0.22), Inches(3.35), cw - Inches(0.44), Inches(3.0),
            items, size=13.5, gap=12, line_spacing=1.02)
footer(s, 9)

# ============================================================================
# SLIDE 10 — LAS ESTATUAS TAMBIÉN MUEREN (1953): génesis
# ============================================================================
s = add_slide()
bg(s, CREAM)
header(s, "Contenido · 4", "Las estatuas también mueren (1953)")

# left: data
bullets(s, Inches(0.85), Inches(2.0), Inches(7.0), Inches(4.4), [
    ("Directores: ", "Chris Marker y Alain Resnais, vinculados a la Rive Gauche (junto a Agnès Varda)."),
    ("Encargo: ", "el intelectual senegalés Alioune Diop, a través de la revista Présence Africaine (1950)."),
    ("Ensayo cinematográfico: ", "C. Lupton lo describe como «un ejemplo magistral de ensayo cinematográfico»."),
    ("Recepción: ", "censurado en Francia durante 12 años por su lectura anticolonial."),
    ("Continuidad: ", "su huella llega hasta Dahomey (Mati Diop) y el debate actual sobre la restitución."),
], size=15.5, gap=12)

# right: directors mini-cards
def dir_card(x, name, sub, lines):
    card = rect(s, x, Inches(2.0), Inches(4.55), Inches(2.1), WHITE)
    card.line.color.rgb = CREAM2; card.line.width = Pt(1)
    rect(s, x, Inches(2.0), Inches(0.1), Inches(2.1), BRONZE)
    txt(s, x + Inches(0.28), Inches(2.15), Inches(4.1), Inches(1.9),
        [
            {"text": name, "size": 17, "color": INK, "bold": True,
             "font": FONT_B, "space_after": 1},
            {"text": sub, "size": 12, "color": BROWN, "italic": True,
             "font": FONT_B, "space_after": 6},
            {"text": lines, "size": 13, "color": GREY, "font": FONT_B,
             "line_spacing": 1.05},
        ])

dir_card(Inches(8.2), "Chris Marker",
         "Documental subjetivo · film-essay",
         "«Soy un ensayista... y a mí hacer ensayos.» Encargado del proyecto por su afinidad con Présence Africaine.")
dir_card2_y = None
card = rect(s, Inches(8.2), Inches(4.35), Inches(4.55), Inches(2.1), WHITE)
card.line.color.rgb = CREAM2; card.line.width = Pt(1)
rect(s, Inches(8.2), Inches(4.35), Inches(0.1), Inches(2.1), BRONZE)
txt(s, Inches(8.48), Inches(4.5), Inches(4.1), Inches(1.9),
    [
        {"text": "Alain Resnais", "size": 17, "color": INK, "bold": True,
         "font": FONT_B, "space_after": 1},
        {"text": "Rive Gauche · documental de arte", "size": 12, "color": BROWN,
         "italic": True, "font": FONT_B, "space_after": 6},
        {"text": "Óscar por Van Gogh (1950). Composiciones dinámicas, "
                 "travellings y montaje para explorar las obras.",
         "size": 13, "color": GREY, "font": FONT_B, "line_spacing": 1.05},
    ])
footer(s, 10)

# ============================================================================
# SLIDE 11 — CONTENIDO DEL FILM: dos bloques de obras
# ============================================================================
s = add_slide()
bg(s, CREAM)
header(s, "Contenido · 5", "El film: dos bloques de obras descontextualizadas")

# quote banner
rect(s, Inches(0.85), Inches(1.95), Inches(11.9), Inches(0.7), DARK)
txt(s, Inches(1.1), Inches(1.95), Inches(11.4), Inches(0.7),
    [{"text": "«Cuando los hombres mueren, se vuelven historia. "
              "Cuando las estatuas mueren, se vuelven arte.»",
      "size": 16, "color": CREAM, "italic": True, "font": FONT_H,
      "align": PP_ALIGN.CENTER}], anchor=MSO_ANCHOR.MIDDLE)

# two blocks
b1 = rect(s, Inches(0.85), Inches(2.85), Inches(5.85), Inches(0.55), BRONZE)
txt(s, Inches(0.85), Inches(2.85), Inches(5.85), Inches(0.55),
    [{"text": "Bloque I · Realeza y poder", "size": 15, "color": WHITE,
      "bold": True, "align": PP_ALIGN.CENTER, "font": FONT_B}],
    anchor=MSO_ANCHOR.MIDDLE)
bullets(s, Inches(1.05), Inches(3.55), Inches(5.5), Inches(2.9), [
    "Cabezas de oba (Reino de Benín) y de oni (Reino de Ife), Nigeria.",
    "Guerreros y leopardos de bronce de Benín.",
    "Montaje rítmico y fundidos: ¿colonizó Francia un pueblo «sin cultura»?",
], size=14, gap=12)

b2 = rect(s, Inches(7.0), Inches(2.85), Inches(5.75), Inches(0.55), BROWN)
txt(s, Inches(7.0), Inches(2.85), Inches(5.75), Inches(0.55),
    [{"text": "Bloque II · Culto y ritual", "size": 15, "color": WHITE,
      "bold": True, "align": PP_ALIGN.CENTER, "font": FONT_B}],
    anchor=MSO_ANCHOR.MIDDLE)
bullets(s, Inches(7.2), Inches(3.55), Inches(5.4), Inches(2.9), [
    "Relicarios Kota (Gabón); tableros de adivinación Ifá, Yoruba (Nigeria).",
    "Reposacabezas Pende (R. D. del Congo); máscaras Dan (Costa de Marfil).",
    "Culto animista y a los ancestros; piezas sin contexto identificativo.",
], size=14, gap=12)
footer(s, 11)

# ============================================================================
# SLIDE 12 — NUEVA GRAMÁTICA + PROBLEMÁTICA DE GÉNERO
# ============================================================================
s = add_slide()
bg(s, CREAM)
header(s, "Contenido · 6", "Forma fílmica: del documental de arte al ensayo")

cols = [
    ("Fotografía", "Ghislain Cloquet expone las piezas sobre fondos neutros: detalle y, a la vez, refuerzo de su descontextualización."),
    ("Música", "Guy Bernard alterna lo familiar (europeo-americano) y lo «exótico-ritual», con momentos sincréticos."),
    ("Montaje", "Material de archivo (influencia de Esfir Shub): contrasta las «bondades» coloniales con su hipocresía."),
]
cx = 0.85; cw = Inches(3.85); gx = 0.2
for i, (h, body) in enumerate(cols):
    x = Inches(cx + i * (3.85 + gx))
    card = rect(s, x, Inches(2.0), cw, Inches(2.5), WHITE)
    card.line.color.rgb = CREAM2; card.line.width = Pt(1)
    rect(s, x, Inches(2.0), cw, Inches(0.5), DARK)
    txt(s, x, Inches(2.0), cw, Inches(0.5),
        [{"text": h, "size": 15, "color": BRONZE, "bold": True,
          "align": PP_ALIGN.CENTER, "font": FONT_B}], anchor=MSO_ANCHOR.MIDDLE)
    txt(s, x + Inches(0.25), Inches(2.7), cw - Inches(0.5), Inches(1.7),
        [{"text": body, "size": 13.5, "color": INK, "font": FONT_B,
          "line_spacing": 1.08}])

# conclusion strip
rect(s, Inches(0.85), Inches(4.85), Inches(11.9), Inches(1.55), BROWN)
rect(s, Inches(0.85), Inches(4.85), Inches(0.14), Inches(1.55), BRONZE)
txt(s, Inches(1.2), Inches(5.0), Inches(11.3), Inches(1.3),
    [
        {"text": "¿Documental de arte o ensayo fílmico?", "size": 16,
         "color": WHITE, "bold": True, "font": FONT_B, "space_after": 4},
        {"text": "Aunque por su forma parezca un documental de arte, Resnais "
                 "y Marker lo exceden: la Rive Gauche desarrolla el ciné-essai, "
                 "que aborda con ironía y firmeza la problemática colonial y "
                 "la mirada racista aún vigente.",
         "size": 13.5, "color": CREAM, "font": FONT_B, "line_spacing": 1.08},
    ])
footer(s, 12)

# ============================================================================
# SLIDE 13 — CONCLUSIONES
# ============================================================================
s = add_slide()
bg(s, CREAM)
header(s, "Conclusiones", "Conclusiones")

bullets(s, Inches(0.85), Inches(2.0), Inches(7.2), Inches(4.4), [
    ("Punto de inflexión: ", "el film supera los estándares del documental de arte y abre debate sobre identidad, poder colonial y pérdida de significados."),
    ("Tratamiento ontológico: ", "frente al cine etnográfico-expositivo, da voz a las historias reprimidas tras cada pieza."),
    ("Ensayo fílmico-crítico: ", "se confirma la hipótesis; la obra se alinea con la ensayística más que con la exposición."),
    ("Preguntas abiertas: ", "¿devolverá Occidente el patrimonio expoliado? ¿en qué condiciones? ¿está África preparada para conservarlo?"),
], size=15.5, gap=13)

# closing quote card
card = rect(s, Inches(8.3), Inches(2.0), Inches(4.45), Inches(4.45), DARK)
rect(s, Inches(8.3), Inches(2.0), Inches(4.45), Inches(0.14), BRONZE)
txt(s, Inches(8.65), Inches(2.45), Inches(3.8), Inches(3.8),
    [{"text": "«Cuando el documental de arte abandona la mera descripción "
              "para convertirse en una herramienta de pensamiento crítico, "
              "deja de explicar el mundo para empezar a cuestionarlo.»",
      "size": 17, "color": CREAM, "italic": True, "font": FONT_H,
      "line_spacing": 1.18}], anchor=MSO_ANCHOR.MIDDLE)
footer(s, 13)

# ============================================================================
# SLIDE 14 — REFERENCIAS (selección)
# ============================================================================
s = add_slide()
bg(s, CREAM)
header(s, "Referencias bibliográficas", "Referencias (selección)")

refs_l = [
    "Aliaga Cárceles, J. J. La revista Imágenes y la visión documental de NO-DO. Tesis doctoral, 2022.",
    "Alter, N. M. Chris Marker. Urbana: University of Illinois Press, 2006.",
    "Ceamanos, R. y Kabunda, M. El Reparto de África. Madrid: Los libros de la Catarata, 2016.",
    "Fernández Cuenca, C. 30 años de documental de arte en España. Madrid: EOC, 1967.",
    "Forni, S. y Steiner, C. B. «The African Art Market as Ego-System». Critical Interventions, 2018.",
    "García Fernández, G. «La piel bajo las máscaras. Arte, política y ensayo en Les Statues meurent aussi», 2010.",
]
refs_r = [
    "García Peydró, G. El cine sobre arte. Santander: Shangrila, 2019.",
    "Hamery, R. «Les films sur l'art d'Afrique noire (1945-1961)», en Le film sur l'art, 2015.",
    "Lupton, C. Chris Marker, memories of the future. Londres: Reaktion, 2005.",
    "Sarr, F. y Savoy, B. The Restitution of African Cultural Heritage. París, 2018.",
    "Schäuble, M. «The Adventure of the Real. Jean Rouch...». Zeitschrift für Ethnologie, 2011.",
    "Robert, V., Le Forestier, L. y Albera, F. Le film sur l'art. Rennes: PUR, 2015.",
]


def ref_col(x, items):
    tb = s.shapes.add_textbox(x, Inches(2.05), Inches(5.8), Inches(4.6))
    tf = tb.text_frame; tf.word_wrap = True
    for i, t in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(11); p.line_spacing = 1.0
        rd = p.add_run(); rd.text = "•  "
        rd.font.color.rgb = BRONZE; rd.font.bold = True; rd.font.size = Pt(12)
        r = p.add_run(); r.text = t
        r.font.size = Pt(11.5); r.font.name = FONT_B; r.font.color.rgb = INK


ref_col(Inches(0.85), refs_l)
ref_col(Inches(7.0), refs_r)
footer(s, 14)

# ============================================================================
# SLIDE 15 — CIERRE
# ============================================================================
s = add_slide()
bg(s, DARK)
rect(s, Inches(0), Inches(0), SW, Inches(0.14), BRONZE)
rect(s, Inches(0), Inches(7.36), SW, Inches(0.14), BRONZE)
txt(s, Inches(1.0), Inches(2.7), Inches(11.3), Inches(1.2),
    [{"text": "Gracias por su atención", "size": 40, "color": WHITE,
      "font": FONT_H, "align": PP_ALIGN.CENTER}])
rect(s, Inches(5.66), Inches(4.05), Inches(2.0), Pt(2.5), BRONZE)
txt(s, Inches(1.0), Inches(4.3), Inches(11.3), Inches(1.2),
    [
        {"text": "Pedro José Díaz Bravo", "size": 18, "color": BRONZE,
         "bold": True, "align": PP_ALIGN.CENTER, "font": FONT_B,
         "space_after": 4},
        {"text": "Más allá del documental de arte: ensayo fílmico y crítica "
                 "anticolonial en Las estatuas también mueren",
         "size": 13, "color": CREAM, "italic": True, "align": PP_ALIGN.CENTER,
         "font": FONT_H, "line_spacing": 1.1},
        {"text": "Grado en Historia del Arte · Universidad de Murcia · 2025-2026",
         "size": 11, "color": GREY, "align": PP_ALIGN.CENTER, "font": FONT_B,
         "space_before": 6},
    ])

out = "Defensa_TFG_Las_estatuas_tambien_mueren.pptx"
prs.save(out)
print("Guardado:", out, "·", len(prs.slides._sldIdLst), "diapositivas")
