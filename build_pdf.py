from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree
import copy

ORANGE = RGBColor(0xD9, 0x53, 0x2B)
BROWN  = RGBColor(0x76, 0x48, 0x2E)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)

W = Inches(13.333)
H = Inches(7.5)

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H

blank = prs.slide_layouts[6]  # fully blank


def add_bg(slide, img_path):
    pic = slide.shapes.add_picture(img_path, 0, 0, W, H)
    slide.shapes._spTree.remove(pic._element)
    slide.shapes._spTree.insert(2, pic._element)


def add_text(slide, text, left, top, width, height,
             font_name='League Spartan', font_size=48,
             bold=False, color=BROWN, align=PP_ALIGN.CENTER,
             line_height=None):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = align
    if line_height:
        from pptx.oxml.ns import qn
        from lxml import etree
        pPr = p._pPr
        if pPr is None:
            pPr = p._p.get_or_add_pPr()
        lnSpc = etree.SubElement(pPr, qn('a:lnSpc'))
        spcPct = etree.SubElement(lnSpc, qn('a:spcPct'))
        spcPct.set('val', str(int(line_height * 1000)))
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    return txBox


def add_mixed_text(slide, parts, left, top, width, height,
                   font_name='League Spartan', font_size=48,
                   align=PP_ALIGN.CENTER):
    """parts = list of (text, bold, color)"""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = align
    for text, bold, color in parts:
        run = p.add_run()
        run.text = text
        run.font.name = font_name
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.color.rgb = color
    return txBox


def add_rounded_pill(slide, left, top, width, height, fill_color, radius_pct=50):
    from pptx.util import Pt
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        left, top, width, height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    # Round corners via XML
    sp = shape._element
    spPr = sp.find(qn('p:spPr'))
    prstGeom = spPr.find(qn('a:prstGeom'))
    if prstGeom is not None:
        spPr.remove(prstGeom)
    custGeom = etree.fromstring(f'''
      <a:prstGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        prst="roundRect">
        <a:avLst>
          <a:gd name="adj" fmla="val {radius_pct * 1000}"/>
        </a:avLst>
      </a:prstGeom>
    ''')
    spPr.insert(0, custGeom)
    return shape


# ── Slide 1: Intro ────────────────────────────────────────────
s1 = prs.slides.add_slide(blank)
add_bg(s1, 'bkg_slide.png')

# Logo
logo_w = Inches(3.6)
logo_h = Inches(1.56)
logo_l = (W - logo_w) / 2
logo_t = Inches(1.1)
s1.shapes.add_picture('logo_slide.png', logo_l, logo_t, logo_w, logo_h)

# "Weekly Check-In"
add_text(s1, 'Weekly Check-In',
         left=Inches(1), top=Inches(3.05), width=Inches(11.333), height=Inches(1.3),
         font_size=72, bold=True, color=BROWN)

# "& Sprints"
add_text(s1, '& Sprints',
         left=Inches(1), top=Inches(4.1), width=Inches(11.333), height=Inches(1.3),
         font_size=72, bold=False, color=ORANGE)

# Orange rule bar
rule = s1.shapes.add_shape(1, (W - Inches(0.7)) / 2, Inches(5.45), Inches(0.7), Pt(4))
rule.fill.solid()
rule.fill.fore_color.rgb = ORANGE
rule.line.fill.background()


# ── Slide 2: Pitches + Demo ───────────────────────────────────
s2 = prs.slides.add_slide(blank)
add_bg(s2, 'bkg_slide.png')

# "UP NEXT" eyebrow
add_text(s2, 'UP NEXT',
         left=Inches(1), top=Inches(1.1), width=Inches(11.333), height=Inches(0.4),
         font_size=13, bold=True, color=ORANGE)

# "Pitches & Demo" — mixed colour
add_mixed_text(s2,
    parts=[('Pitches & ', True, BROWN), ('Demo', True, ORANGE)],
    left=Inches(1), top=Inches(1.6), width=Inches(11.333), height=Inches(1.2),
    font_size=66)

# Instruction
add_text(s2, 'Add your project or product to Lightning Ladder',
         left=Inches(1), top=Inches(2.95), width=Inches(11.333), height=Inches(0.55),
         font_size=22, bold=False, color=BROWN)

# URL pill background
pill_w = Inches(8.2)
pill_h = Inches(0.72)
pill_l = (W - pill_w) / 2
pill_t = Inches(3.65)
add_rounded_pill(s2, pill_l, pill_t, pill_w, pill_h, BROWN, radius_pct=50)

# URL text on pill
add_text(s2, 'lightning-ladder.netlify.app/#/room/SPRINT',
         left=pill_l, top=pill_t, width=pill_w, height=pill_h,
         font_size=20, bold=True, color=WHITE)

# Note
add_mixed_text(s2,
    parts=[('Enter your ', False, RGBColor(0xA0, 0x7B, 0x5C)),
           ('project or product name', True, BROWN),
           (' — not your personal name', False, RGBColor(0xA0, 0x7B, 0x5C))],
    left=Inches(1), top=Inches(4.6), width=Inches(11.333), height=Inches(0.45),
    font_size=16)


# ── Slide 3: Roadmap ──────────────────────────────────────────
s3 = prs.slides.add_slide(blank)
add_bg(s3, 'bkg_slide.png')

# "LOOKING AHEAD" eyebrow
add_text(s3, 'LOOKING AHEAD',
         left=Inches(1), top=Inches(1.3), width=Inches(11.333), height=Inches(0.4),
         font_size=13, bold=True, color=ORANGE)

# "Roadmap"
add_text(s3, 'Roadmap',
         left=Inches(1), top=Inches(1.85), width=Inches(11.333), height=Inches(2.2),
         font_size=108, bold=True, color=BROWN)

# Subtitle
add_text(s3, "What's next for Sprints & Weekly Check-Ins",
         left=Inches(1), top=Inches(4.15), width=Inches(11.333), height=Inches(0.6),
         font_size=26, bold=False, color=BROWN)

# Gradient divider line (solid orange, centred)
div = s3.shapes.add_shape(1, (W - Inches(4)) / 2, Inches(5.0), Inches(4), Pt(2))
div.fill.solid()
div.fill.fore_color.rgb = ORANGE
div.line.fill.background()
div.fill.fore_color.theme_color  # noop, keep solid

prs.save('slides.pptx')
print('PPTX saved.')
