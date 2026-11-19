"""Low-level, fully editable DrawingML primitives (shapes, rich text, pictures)."""
import re
from pathlib import Path

from lxml import etree
from PIL import Image
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

from .theme import Theme, hex_rgb

ALIGN = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
ANCHOR = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}

# **bold accent**   _{sub}   ^{sup}   ==primary bold==
TOKEN = re.compile(r"(\*\*.+?\*\*|==.+?==|_\{.+?\}|\^\{.+?\})")


def set_run_font(run, theme: Theme, size=None, bold=None, color=None, font=None, italic=None):
    f = run.font
    if size:
        f.size = Pt(size)
    if bold is not None:
        f.bold = bold
    if italic is not None:
        f.italic = italic
    if color:
        f.color.rgb = hex_rgb(color)
    latin = theme.font_en if font is None else font
    f.name = latin
    rpr = run._r.get_or_add_rPr()
    for tag, face in (("a:ea", font or theme.font_cn), ("a:cs", font or theme.font_cn)):
        el = rpr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rpr, qn(tag))
        el.set("typeface", face)


def _segments(text: str):
    for part in TOKEN.split(text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            yield part[2:-2], "accent", 0
        elif part.startswith("==") and part.endswith("=="):
            yield part[2:-2], "primary", 0
        elif part.startswith("_{"):
            yield part[2:-1], None, -25000
        elif part.startswith("^{"):
            yield part[2:-1], None, 30000
        else:
            yield part, None, 0


def write_rich(tf, text, theme: Theme, size=16, color=None, bold=False, align="left",
               anchor="middle", line_spacing=1.1, font=None, wrap=True, autofit=False):
    """Fill a text frame. Lines split on '\n'. Supports **accent**, ==primary==, _{sub}, ^{sup}."""
    tf.clear()
    tf.word_wrap = wrap
    tf.vertical_anchor = ANCHOR.get(anchor, MSO_ANCHOR.MIDDLE)
    for side in ("margin_left", "margin_right"):
        setattr(tf, side, Emu(45720))
    tf.margin_top = tf.margin_bottom = Emu(22860)
    color = color or theme.text
    for i, line in enumerate(str(text).split("\n")):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = ALIGN.get(align, PP_ALIGN.LEFT)
        para.line_spacing = line_spacing
        for chunk, style, baseline in _segments(line):
            run = para.add_run()
            run.text = chunk
            c = {"accent": theme.accent, "primary": theme.primary}.get(style, color)
            set_run_font(run, theme, size, bold or style is not None, c, font)
            if baseline:
                run._r.get_or_add_rPr().set("baseline", str(baseline))
    if autofit:
        body = tf._txBody.find(qn("a:bodyPr"))
        for child in list(body):
            if child.tag in (qn("a:spAutoFit"), qn("a:noAutofit"), qn("a:normAutofit")):
                body.remove(child)
        etree.SubElement(body, qn("a:normAutofit"))
    return tf


def no_line(shape):
    shape.line.fill.background()


def shape_box(slide, x, y, w, h, fill=None, line=None, kind=MSO_SHAPE.RECTANGLE,
              line_w=1.0, radius=None, name=None, shadow=False):
    shp = slide.shapes.add_shape(kind, Emu(int(x)), Emu(int(y)), Emu(int(w)), Emu(int(h)))
    if fill:
        shp.fill.solid()
        shp.fill.fore_color.rgb = hex_rgb(fill)
    else:
        shp.fill.background()
    if line:
        shp.line.color.rgb = hex_rgb(line)
        shp.line.width = Pt(line_w)
    else:
        no_line(shp)
    if radius is not None and shp.adjustments and len(shp.adjustments) > 0:
        shp.adjustments[0] = radius
    if not shadow:
        sppr = shp._element.spPr
        eff = etree.SubElement(sppr, qn("a:effectLst"))
        del eff  # empty effect list => no theme shadow
    if name:
        shp.name = name
    if shp.has_text_frame:
        shp.text_frame.text = ""
    return shp


def text_box(slide, x, y, w, h, text, theme, size=16, color=None, bold=False, align="left",
             anchor="middle", fill=None, line=None, kind=None, name=None, font=None,
             line_spacing=1.1, autofit=False, radius=None):
    if kind is None and fill is None and line is None:
        shp = slide.shapes.add_textbox(Emu(int(x)), Emu(int(y)), Emu(int(w)), Emu(int(h)))
    else:
        shp = shape_box(slide, x, y, w, h, fill, line, kind or MSO_SHAPE.RECTANGLE, radius=radius)
    write_rich(shp.text_frame, text, theme, size, color, bold, align, anchor, line_spacing,
               font, autofit=autofit)
    if name:
        shp.name = name
    return shp


def gradient_fill(shape, stops, angle=0):
    """Replace fill with a linear gradient. stops: list of hex colours, evenly spaced."""
    sppr = shape._element.spPr
    for tag in ("a:solidFill", "a:noFill", "a:gradFill"):
        for el in sppr.findall(qn(tag)):
            sppr.remove(el)
    grad = etree.Element(qn("a:gradFill"), rotWithShape="1")
    lst = etree.SubElement(grad, qn("a:gsLst"))
    n = max(1, len(stops) - 1)
    for i, col in enumerate(stops):
        gs = etree.SubElement(lst, qn("a:gs"), pos=str(int(100000 * i / n)))
        etree.SubElement(gs, qn("a:srgbClr"), val=col.lstrip("#"))
    etree.SubElement(grad, qn("a:lin"), ang=str(int(angle * 60000)), scaled="0")
    geom = sppr.find(qn("a:prstGeom"))
    (geom.addnext if geom is not None else sppr.append)(grad)
    return shape


def picture_fit(slide, path, x, y, w, h, align="center"):
    """Insert a picture scaled to fit (contain) inside the box without cropping."""
    path = Path(path)
    with Image.open(path) as im:
        iw, ih = im.size
    scale = min(w / iw, h / ih)
    pw, ph = iw * scale, ih * scale
    px = x + (w - pw) / 2 if align == "center" else x
    py = y + (h - ph) / 2
    return slide.shapes.add_picture(str(path), Emu(int(px)), Emu(int(py)), Emu(int(pw)), Emu(int(ph)))


def connector(slide, x1, y1, x2, y2, color, width=1.25, arrow=True):
    from pptx.enum.shapes import MSO_CONNECTOR
    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Emu(int(x1)), Emu(int(y1)),
                                      Emu(int(x2)), Emu(int(y2)))
    line.line.color.rgb = hex_rgb(color)
    line.line.width = Pt(width)
    if arrow:
        ln = line.line._get_or_add_ln()
        etree.SubElement(ln, qn("a:tailEnd"), type="triangle", w="med", len="med")
    return line
