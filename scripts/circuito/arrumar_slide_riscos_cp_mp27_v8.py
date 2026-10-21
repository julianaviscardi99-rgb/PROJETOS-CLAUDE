"""CP MP'27 v8: reorganiza a ESTRUTURA do slide Riscos e Oportunidades (slide 6).

Textos intocados. Muda so posicao/tamanho: duas colunas simetricas, cards de mesma
altura alinhados linha a linha, valores alinhados a direita, totais na mesma linha,
barra de acento colorida em cada card.

Uso: python arrumar_slide_riscos_cp_mp27_v8.py <entrada.pptx> <saida.pptx>
"""
import copy
import sys

from lxml import etree
from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

LEFT = 457200
COL_W = 5676900
GAP = 228600
RIGHT_X = LEFT + COL_W + GAP

PAD = 180000          # recuo interno do card
VAL_W = 900000        # caixa dos valores
CARD_H = 1600000
CARD_GAP = 140000
CARDS_Y = 1760000
TOTAL_LINE_Y = CARDS_Y + 2 * CARD_H + CARD_GAP + 220000
TOTAL_H = 330000
ACCENT_W = 50800


def by_id(slide, sid):
    return next(s for s in slide.shapes if s.shape_id == sid)


def no_autofit(shape):
    bp = shape.text_frame._txBody.bodyPr
    for tag in ("spAutoFit", "normAutofit", "noAutofit"):
        for el in bp.findall(f"{{{NS_A}}}{tag}"):
            bp.remove(el)
    etree.SubElement(bp, f"{{{NS_A}}}noAutofit")


def square_card(card):
    geom = card._element.spPr.find(f"{{{NS_A}}}prstGeom")
    if geom is not None:
        geom.set("prst", "rect")
        av = geom.find(f"{{{NS_A}}}avLst")
        if av is not None:
            for g in list(av):
                av.remove(g)


def accent(slide, card, color_shape):
    """Barra fina na lateral esquerda do card, na cor do cabecalho da coluna."""
    bar = copy.deepcopy(card._element)
    card._element.addnext(bar)
    nid = max(s.shape_id for s in slide.shapes) + 1
    bar.xpath(".//p:cNvPr")[0].set("id", str(nid))
    bar.xpath(".//p:cNvPr")[0].set("name", "Acento " + str(nid))
    shp = next(s for s in slide.shapes if s.shape_id == nid)
    shp.left, shp.top, shp.width, shp.height = card.left, card.top, Emu(ACCENT_W), card.height
    shp.fill.solid()
    shp.fill.fore_color.rgb = color_shape.fill.fore_color.rgb
    shp.line.fill.background()
    if shp.has_text_frame:
        for p in shp.text_frame.paragraphs:
            for r in p.runs:
                r.text = ""
    return shp


def place_item(slide, col_x, y, card_id, text_id, val_id, hdr):
    card, text, val = by_id(slide, card_id), by_id(slide, text_id), by_id(slide, val_id)
    square_card(card)
    card.left, card.top, card.width, card.height = Emu(col_x), Emu(y), Emu(COL_W), Emu(CARD_H)
    accent(slide, card, hdr)

    inner_y, inner_h = y + 60000, CARD_H - 120000
    text.left = Emu(col_x + PAD + ACCENT_W)
    text.width = Emu(COL_W - 2 * PAD - ACCENT_W - VAL_W)
    text.top, text.height = Emu(inner_y), Emu(inner_h)
    no_autofit(text)
    text.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    for p in text.text_frame.paragraphs:
        p.alignment = PP_ALIGN.LEFT

    val.left = Emu(col_x + COL_W - PAD - VAL_W)
    val.width = Emu(VAL_W)
    val.top, val.height = Emu(inner_y), Emu(inner_h)
    no_autofit(val)
    val.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    for p in val.text_frame.paragraphs:
        p.alignment = PP_ALIGN.RIGHT


def place_total(slide, col_x, line_id, label_id, val_id):
    line, label, val = by_id(slide, line_id), by_id(slide, label_id), by_id(slide, val_id)
    line.left, line.top, line.width = Emu(col_x), Emu(TOTAL_LINE_Y), Emu(COL_W)
    y = TOTAL_LINE_Y + 90000
    label.left, label.top = Emu(col_x + PAD), Emu(y)
    label.width, label.height = Emu(COL_W - 2 * PAD - VAL_W), Emu(TOTAL_H)
    val.left, val.top = Emu(col_x + COL_W - PAD - VAL_W), Emu(y)
    val.width, val.height = Emu(VAL_W), Emu(TOTAL_H)
    for sh, al in ((label, PP_ALIGN.LEFT), (val, PP_ALIGN.RIGHT)):
        no_autofit(sh)
        sh.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        for p in sh.text_frame.paragraphs:
            p.alignment = al


def main(src, dst):
    prs = Presentation(src)
    s = prs.slides[5]
    assert "Riscos e Oportunidades" in s.shapes.title.text_frame.text

    hdr_r, hdr_o = by_id(s, 63), by_id(s, 83)
    for hdr, unit, x in ((hdr_r, by_id(s, 64), LEFT), (hdr_o, by_id(s, 84), RIGHT_X)):
        hdr.left, hdr.width = Emu(x), Emu(COL_W)
        unit.width = Emu(1750000)
        unit.left = Emu(x + COL_W - PAD - 1750000)
        for p in unit.text_frame.paragraphs:
            p.alignment = PP_ALIGN.RIGHT

    msg = by_id(s, 99)
    msg.left, msg.width = Emu(LEFT), Emu(COL_W * 2 + GAP)
    runs = [r for p in msg.text_frame.paragraphs for r in p.runs]
    runs[0].text = ("Riscos de −926 superam as oportunidades de +105; "
                    "upside com montadoras chinesas em avaliação")
    for r in runs[1:]:
        r._r.getparent().remove(r._r)

    y1, y2 = CARDS_Y, CARDS_Y + CARD_H + CARD_GAP
    place_item(s, LEFT, y1, 65, 66, 67, hdr_r)
    place_item(s, LEFT, y2, 11, 12, 13, hdr_r)
    place_item(s, RIGHT_X, y1, 6, 7, 87, hdr_o)
    place_item(s, RIGHT_X, y2, 8, 9, 10, hdr_o)

    place_total(s, LEFT, 80, 81, 82)
    place_total(s, RIGHT_X, 94, 95, 96)

    prs.save(dst)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
