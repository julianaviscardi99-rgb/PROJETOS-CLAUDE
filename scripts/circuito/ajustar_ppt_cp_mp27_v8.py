"""CP MP'27 v8: ajusta o bridge do EBIT (slide 2) sobre a v7 dela.

- Redução de despesas +150 -> +200; Projeto portaria segue -40; nova barra "Others" -50.
  Bridge continua fechando em 2.845 (2.486 +817 +494 -411 -300 -138 -213 +200 -40 -50).
- Comentário do EBITDA (painel Resumo executivo) -> comentário de inflação.
- Demais slides, tabela e comentários intocados.

Uso: python ajustar_ppt_cp_mp27_v8.py <v7.pptx> <v8.pptx>
"""
import copy
import sys

from pptx import Presentation
from pptx.util import Emu

K = 1046.8  # EMU por unidade (BRL '000) - calibrado nas barras originais
BASE_LEVEL = 6019800  # topo da barra R9 (nivel 2.486)
AXIS_Y = 6248400
X0, WIDTH = 838200, 8051800
NCOLS = 12
STEPS = [817, 494, -611, 200, -300, -138, -213, -40, -50, 200]  # EBIT: + ganho, - perda
FINAL_TOP = 5644196


def by_id(slide, sid):
    return next(s for s in slide.shapes if s.shape_id == sid)


def set_text(shape, lines, size=None):
    """Troca o texto mantendo o formato do primeiro run de cada paragrafo."""
    for para, text in zip(shape.text_frame.paragraphs, lines):
        runs = para.runs
        runs[0].text = text
        if size:
            runs[0]._r.get_or_add_rPr().set("sz", str(size))
        for r in runs[1:]:
            r._r.getparent().remove(r._r)


def set_paras(shape, lines):
    """Ajusta o numero de paragrafos (clonando o ultimo) e o texto de cada um."""
    paras = shape.text_frame.paragraphs
    while len(shape.text_frame.paragraphs) < len(lines):
        paras[-1]._p.addnext(copy.deepcopy(paras[-1]._p))
        paras = shape.text_frame.paragraphs
    for p in paras[len(lines):]:
        p._p.getparent().remove(p._p)
    for p, text in zip(shape.text_frame.paragraphs, lines):
        p.runs[0].text = text
        for r in p.runs[1:]:
            r._r.getparent().remove(r._r)


def clone_col(slide, ids):
    shapes = [by_id(slide, i) for i in ids]
    next_id = max(sh.shape_id for sh in slide.shapes) + 1
    new = []
    for sh in shapes:
        el = copy.deepcopy(sh._element)
        el.xpath(".//p:cNvPr")[0].set("id", str(next_id))
        next_id += 1
        sh._element.addnext(el)
        new.append(el)
    return [next(sh for sh in slide.shapes if sh._element is el) for el in new]


def main(src, dst):
    prs = Presentation(src)
    s = prs.slides[1]

    # ---- 1) novas colunas: "Others" (clone de portaria) e "Redução 2027" (clone de redução)
    rect9, conn9, lab9, cat9 = clone_col(s, (175, 176, 177, 189))
    set_text(lab9, ["-50"])
    set_paras(cat9, ["Others"])
    rect10, conn10, lab10, cat10 = clone_col(s, (172, 173, 174, 188))
    set_text(lab10, ["+200"])
    set_paras(cat10, ["Redução de", "despesas", "2027"])

    # ---- 2) colunas na ordem: R9, 8 deltas..., MP27
    cols = [
        (by_id(s, 152), None, by_id(s, 153), by_id(s, 181)),  # R9
    ]
    ids = [(154, 155, 156, 182), (157, 158, 159, 183), (160, 161, 162, 184),
           (163, 164, 165, 185), (166, 167, 168, 186), (169, 170, 171, 187),
           (172, 173, 174, 188), (175, 176, 177, 189)]
    for r, c, l, t in ids:
        cols.append((by_id(s, r), by_id(s, c), by_id(s, l), by_id(s, t)))
    cols.append((rect9, conn9, lab9, cat9))
    cols.append((rect10, conn10, lab10, cat10))
    cols.append((by_id(s, 178), by_id(s, 179), by_id(s, 180), by_id(s, 190)))  # MP27
    assert len(cols) == NCOLS

    pitch = WIDTH // NCOLS
    bw = int(pitch * 0.66)
    off = (pitch - bw) // 2

    # textos
    set_text(cols[7][2], ["+200"])
    # Redução de despesas logo depois das Postergações (-411)
    cols = cols[:4] + [cols[7]] + cols[4:7] + cols[8:]
    set_text(cols[3][2], ["-611"])
    set_paras(cols[4][3], ["Redução de", "despesas", "2026"])

    level = BASE_LEVEL
    centers = []
    for i, (rect, conn, lab, cat) in enumerate(cols):
        left = X0 + i * pitch + off
        cx = left + bw // 2
        centers.append(cx)
        rect.left, rect.width = Emu(left), Emu(bw)
        lab.left = Emu(cx - lab.width // 2)
        cat.width = Emu(pitch)
        cat.left = Emu(cx - pitch // 2)

        if i == 0:
            continue
        if i == NCOLS - 1:
            top = FINAL_TOP
            rect.top, rect.height = Emu(top), Emu(AXIS_Y - top)
            lab.top = Emu(top - 241300)
        else:
            d = STEPS[i - 1]
            h = int(round(abs(d) * K))
            top = level - h if d > 0 else level
            if i == NCOLS - 2:  # ultima barra fecha exatamente no nivel final
                h = abs(FINAL_TOP - level)
            rect.top, rect.height = Emu(top), Emu(h)
            lab.top = Emu(top - 177800)
            level = level - h if d > 0 else level + h
        # conector horizontal entre a barra anterior e esta
        pleft = X0 + (i - 1) * pitch + off + bw
        conn.left, conn.width = Emu(pleft), Emu(left - pleft)
    # conectores: y = nivel ao fim da barra anterior
    lvl = BASE_LEVEL
    for i in range(1, NCOLS):
        conn = cols[i][1]
        conn.top = Emu(lvl)
        if i < NCOLS - 1:
            d = STEPS[i - 1]
            h = int(round(abs(d) * K))
            lvl = lvl - h if d > 0 else lvl + h
    assert abs(lvl - FINAL_TOP) < 200, lvl

    # ---- 3) quebras de eixo (freeforms ids 2 e 3) e bracket + oval
    for fid, ci in ((2, 0), (3, NCOLS - 1)):
        ff = by_id(s, fid)
        ff.left, ff.width = Emu(cols[ci][0].left - 38100), Emu(bw + 76200)
    c0, cN = centers[0], centers[-1]
    by_id(s, 191).left = Emu(c0)
    ln = by_id(s, 192)
    ln.left, ln.width = Emu(c0), Emu(cN - c0)
    by_id(s, 193).left = Emu(cN)
    ov = by_id(s, 194)
    ov.left = Emu((c0 + cN) // 2 - ov.width // 2)

    # ---- 4) comentario EBITDA -> inflacao (mesmo lugar e formato)
    box = by_id(s, 207)
    set_text(box, [
        "Inflação 4,4%  |  −213",
        "Reajuste de contratos pressiona o custo fixo; a redução de despesas de 200 absorve 94%.",
    ], size=1100)

    prs.save(dst)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
