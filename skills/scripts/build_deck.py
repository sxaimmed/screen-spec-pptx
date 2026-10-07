#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""화면설계서 PPTX 생성기 (작성규칙 v0.45)

screen_spec.json → 화면설계서 PPTX (+ report.json, 화면목록 xlsx 선택)
  python3 build_deck.py spec.json [-o OUT_DIR] [--only 13,14] [--render all|13,14] [--xlsx]

레이아웃 코드는 모두 이 파일에 고정되어 있으므로, 작성자는 JSON만 쓰면 된다.
"""
import argparse, copy, json, math, os, re, subprocess, sys, datetime
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

VERSION = "0.45"
FONT = "맑은 고딕"
SW, SH = 13.333, 7.5
OX, OY = 0.30, 1.07            # 와이어프레임 원점
WX_MAX, WY_MAX = 10.45, 7.10    # 와이어프레임 영역 한계
S_ADM = 10.15 / 1280
S_MO = 0.0066
C = dict(text="1A1A1A", sub="666666", dis="BBBBBB", link="0066CC", marker="CC2200", tbd="FF8C00",
         line="BFBFBF", head="D9D9D9", lab="F2F2F2", btnE="BFBFBF", btnN="FFFFFF", btnB="7F7F7F",
         memo="FFF2CC", memoB="BF9000", cid="004FA3", hl="FCE4D6", grey="888888")
STATUS = {"신규": "0066CC", "수정": "00B86E", "확인필요": "FF8C00", "정책변경": "8E44AD", "기존": "7F7F7F"}
REF_RE = re.compile(r"\b([CNTQ]-\d+)\b")
EMOJI_RE = re.compile("[\U0001F300-\U0001FAFF✅❌⭐]")

WARN = []
def warn(msg):
    WARN.append(msg)

# ───────────────────────── 텍스트 폭 추정 ─────────────────────────
def is_wide(ch):
    o = ord(ch)
    return o >= 0x1100 and not (0x2000 <= o <= 0x206F)

def text_w(s, size):
    """in 단위 폭 추정 (규칙 §8-4)"""
    return sum(size * (0.98 if is_wide(c) else 0.53) / 72 for c in str(s))

def n_lines(s, size, avail):
    total = 0
    for part in str(s).split("\n"):
        total += max(1, math.ceil(text_w(part, size) / avail - 1e-9))
    return total

def line_h(size):
    return size * 1.2 * 1.2 / 72

# ───────────────────────── 기본 도형 ─────────────────────────
def rgb(h):
    return RGBColor.from_string(h.lstrip("#").upper())

def style_run(run, size=8, bold=False, color=None, underline=False, italic=False):
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = rgb(color or C["text"])
    if underline:
        f.underline = True
    f.name = FONT
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", FONT)

ALIGN = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}
VALIGN = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}

def norm_paras(content, size, bold, color, align):
    """content: str | list[para]. para: str | {"runs":[[text,{opts}]...], "bullet":0/1/2, "align"} | list-of-runs"""
    if content is None:
        content = ""
    if isinstance(content, str):
        content = content.split("\n")
    out = []
    for p in content:
        if isinstance(p, str):
            out.append({"runs": [[p, {}]], "align": align})
        elif isinstance(p, list):
            out.append({"runs": p, "align": align})
        else:
            q = dict(p)
            q.setdefault("align", align)
            if "text" in q:
                q["runs"] = [[q.pop("text"), {k: q[k] for k in ("size", "bold", "color", "underline") if k in q}]]
            out.append(q)
    for q in out:
        q["runs"] = [[r[0], {"size": size, "bold": bold, "color": color, **(r[1] if len(r) > 1 else {})}] for r in q["runs"]]
    return out

def write_tf(tf, content, size=8, bold=False, color=None, align="l", lnspc=None):
    paras = norm_paras(content, size, bold, color, align)
    tf.text = ""
    first = True
    for q in paras:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = ALIGN.get(q.get("align", "l"), PP_ALIGN.LEFT)
        pPr = p._p.get_or_add_pPr()
        ls = q.get("lnspc", lnspc)
        if ls:
            lnS = etree.SubElement(pPr, qn("a:lnSpc"))
            etree.SubElement(lnS, qn("a:spcPct")).set("val", str(int(ls * 1000)))
        b = q.get("bullet", 0)
        if b:
            mar = 101600 if b == 1 else 203200
            pPr.set("marL", str(mar)); pPr.set("indent", "-101600")
            etree.SubElement(pPr, qn("a:buSzPct")).set("val", "100000")
            etree.SubElement(pPr, qn("a:buChar")).set("char", "•" if b == 1 else "-")
        else:
            pPr.set("marL", "0"); pPr.set("indent", "0")
            etree.SubElement(pPr, qn("a:buNone"))
        last_sz = size
        for text, o in q["runs"]:
            if text == "" and len(q["runs"]) == 1:
                last_sz = o.get("size", size)
                continue
            r = p.add_run()
            r.text = T(text)
            last_sz = o.get("size", size)
            style_run(r, last_sz, o.get("bold", False), o.get("color") or C["text"], o.get("underline", False))
        end = p._p.find(qn("a:endParaRPr"))
        if end is None:
            end = etree.SubElement(p._p, qn("a:endParaRPr"))
        end.set("lang", "ko-KR"); end.set("sz", str(int(last_sz * 100)))
        for tag in ("a:latin", "a:ea", "a:cs"):
            if end.find(qn(tag)) is None:
                etree.SubElement(end, qn(tag)).set("typeface", FONT)

def txt(sl, x, y, w, h, content, size=8, bold=False, color=None, align="l", valign="t", name=None, wrap=True, lnspc=None):
    tb = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(max(w, 0.01)), Inches(max(h, 0.01)))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.auto_size = None
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = VALIGN[valign]
    write_tf(tf, content, size, bold, color, align, lnspc)
    if name:
        tb.name = name
    return tb

def set_alpha(fill_el_parent, alpha):
    sf = fill_el_parent.find(qn("a:solidFill"))
    if sf is not None and len(sf):
        a = etree.SubElement(sf[0], qn("a:alpha"))
        a.set("val", str(int(alpha * 1000)))

def box(sl, x, y, w, h, fill=None, line=None, lw=0.75, shape="rect", dash=False, name=None, text=None, size=8,
        bold=False, color=None, align="c", valign="m", ml=0.0, radius=0.04, alpha=None):
    st = {"rect": MSO_SHAPE.RECTANGLE, "round": MSO_SHAPE.ROUNDED_RECTANGLE, "oval": MSO_SHAPE.OVAL,
          "top_round": MSO_SHAPE.ROUND_2_SAME_RECTANGLE}[shape]
    sp = sl.shapes.add_shape(st, Inches(x), Inches(y), Inches(max(w, 0.005)), Inches(max(h, 0.005)))
    if shape in ("round", "top_round"):
        try:
            sp.adjustments[0] = min(0.5, radius / max(min(w, h), 0.01))
        except Exception:
            pass
    if fill:
        sp.fill.solid(); sp.fill.fore_color.rgb = rgb(fill)
        if alpha is not None:
            set_alpha(sp._element.spPr, alpha)
    else:
        sp.fill.background()
    if line:
        sp.line.color.rgb = rgb(line); sp.line.width = Pt(lw)
        if dash:
            ln = sp._element.spPr.find(qn("a:ln"))
            etree.SubElement(ln, qn("a:prstDash")).set("val", "dash")
    else:
        sp.line.fill.background()
    st_el = sp._element.find(qn("p:style"))
    if st_el is not None:
        sp._element.remove(st_el)
    tf = sp.text_frame
    tf.margin_left = Inches(ml); tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    tf.word_wrap = True
    tf.vertical_anchor = VALIGN[valign]
    if text is not None:
        write_tf(tf, text, size, bold, color, align)
    if name:
        sp.name = name
    return sp

def hline(sl, x1, y1, x2, y2, color="BFBFBF", lw=0.75, arrow=False, dash=False, name=None):
    cn = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    st_el = cn._element.find(qn("p:style"))
    if st_el is not None:
        cn._element.remove(st_el)
    cn.line.color.rgb = rgb(color); cn.line.width = Pt(lw)
    ln = cn.line._get_or_add_ln()
    if dash:
        etree.SubElement(ln, qn("a:prstDash")).set("val", "dash")
    if arrow:
        etree.SubElement(ln, qn("a:tailEnd")).set("type", "triangle")
    if name:
        cn.name = name
    return cn

def cell_border(cell, color, w_pt):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        old = tcPr.find(qn(tag))
        if old is not None:
            tcPr.remove(old)
    for tag in reversed(("a:lnL", "a:lnR", "a:lnT", "a:lnB")):
        ln = etree.Element(qn(tag), w=str(int(w_pt * 12700)), cap="flat", cmpd="sng", algn="ctr")
        sf = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr")).set("val", color)
        etree.SubElement(ln, qn("a:prstDash")).set("val", "solid")
        tcPr.insert(0, ln)

def table(sl, x, y, colw, rowh, cells, name=None, border="BFBFBF", bw=0.5, size=8):
    """cells: 2D list; 각 원소 dict | str | None(병합으로 덮인 칸).
    dict 키: text, fill, size, bold, color, align(l/c/r), valign(t/m/b), ml, mr, mt, mb (in), cs(colspan), rs(rowspan), lnspc"""
    nr, nc = len(rowh), len(colw)
    gf = sl.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(sum(colw)), Inches(sum(rowh)))
    tbl = gf.table
    tbl.first_row = False; tbl.horz_banding = False
    for i, w in enumerate(colw):
        tbl.columns[i].width = Inches(w)
    for i, h in enumerate(rowh):
        tbl.rows[i].height = Inches(h)
    merges = []
    for r in range(nr):
        for c in range(nc):
            spec = cells[r][c] if c < len(cells[r]) else None
            cell = tbl.cell(r, c)
            if spec is None:
                spec = {}
            elif isinstance(spec, str):
                spec = {"text": spec}
            if spec.get("fill"):
                cell.fill.solid(); cell.fill.fore_color.rgb = rgb(spec["fill"])
            else:
                cell.fill.background()
            cell_border(cell, border, bw)
            cell.margin_left = Inches(spec.get("ml", 0.05)); cell.margin_right = Inches(spec.get("mr", 0.05))
            cell.margin_top = Inches(spec.get("mt", 0.01)); cell.margin_bottom = Inches(spec.get("mb", 0.01))
            cell.vertical_anchor = VALIGN[spec.get("valign", "m")]
            tf = cell.text_frame
            tf.word_wrap = True
            write_tf(tf, spec.get("text", ""), spec.get("size", size), spec.get("bold", False), spec.get("color"),
                     spec.get("align", "l"), spec.get("lnspc"))
            cs, rs = spec.get("cs", 1), spec.get("rs", 1)
            if cs > 1 or rs > 1:
                merges.append((r, c, r + rs - 1, c + cs - 1))
    for r1, c1, r2, c2 in merges:
        tbl.cell(r1, c1).merge(tbl.cell(r2, c2))
    if name:
        gf.name = name
    return gf

# ───────────────────────── 페이지 토큰·참조ID ─────────────────────────
PAGES = {}          # key -> [pages]
CUR = {"page": None, "collect": False, "scope": "full"}
REFS = {}           # refId -> set(pages)

def T(s):
    s = "" if s is None else str(s)
    def rp(m):
        kind, key = m.group(1), m.group(2)
        pg = PAGES.get(key)
        if not pg:
            if CUR.get("scope") == "partial":
                return "전체본 참조"   # 해당화면 덱에 없는 화면 → 경고 없이 표기
            warn(f"페이지 토큰 대상 없음: {m.group(0)}")
            return "p.?"
        if kind == "p":
            return str(pg[0])
        return str(pg[0]) if len(pg) == 1 else f"{pg[0]}~{pg[-1]}"
    s = re.sub(r"\{\{(pr?):([^}]+)\}\}", rp, s)
    if CUR["collect"] and CUR["page"]:
        for rid in REF_RE.findall(s):
            REFS.setdefault(rid, set()).add(CUR["page"])
    return s

# ───────────────────────── 공통 장표 틀 ─────────────────────────
DOC = {}

def footer(sl, page):
    d = DOC
    txt(sl, 0.40, 7.18, 9.0, 0.2, f"{d.get('project','')} {d.get('title','')} {d.get('version','')} ({d.get('date','')})",
        size=8, color=C["grey"])
    txt(sl, 12.33, 7.18, 0.6, 0.2, str(page), size=8, color=C["grey"], align="r")

def content_slide(prs, page, title, sub=None):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    txt(sl, 0.40, 0.28, 8.5, 0.45, title, size=18, bold=True, valign="m")
    if sub:
        txt(sl, 0.40, 0.72, 12.5, 0.2, sub, size=10, color=C["sub"])
    hline(sl, 0.40, 0.98, 12.93, 0.98, "BFBFBF", 0.75)
    footer(sl, page)
    return sl

def plain_slide(prs, page):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    footer(sl, page)
    return sl

def hdr_cells(heads, rows, widths_n, head_fill=C["lab"], size=8, align=None):
    cells = [[{"text": h, "fill": head_fill, "bold": True, "align": "c", "size": size} for h in heads]]
    for r in rows:
        line = []
        for i, v in enumerate(r):
            a = (align[i] if align else "c")
            line.append({"text": "" if v is None else str(v), "align": a, "size": size})
        cells.append(line)
    return cells

# ───────────────────────── 표지·이력·요구사항·개요 ─────────────────────────
def slide_cover(prs, page):
    d = DOC
    sl = plain_slide(prs, page)
    txt(sl, 1.0, 1.7, 11.33, 0.6, d.get("project", ""), size=24, bold=True, align="c", valign="m")
    txt(sl, 1.0, 2.35, 11.33, 0.5, d.get("title", ""), size=20, align="c", valign="m", color="333333")
    rows = [["문서ID", d.get("docId") or "-"], ["버 전", d.get("version", "")], ["작 성 일", d.get("date", "")],
            ["작 성 자", d.get("author") or "-"], ["기준 문서", d.get("source") or "-"]]
    cells = [[{"text": a, "fill": C["lab"], "bold": True, "align": "c", "size": 10}, {"text": b, "align": "l", "size": 10}] for a, b in rows]
    table(sl, 3.67, 3.6, [1.4, 4.6], [0.34] * len(rows), cells, name="cover-info", border="BFBFBF")

def slide_revision(prs, page, rows, part=""):
    sl = content_slide(prs, page, "Revision History" + part)
    cells = hdr_cells(["Version", "Date", "페이지", "Notes", "Writer"], rows, None, align="cccl" "c")
    table(sl, 0.40, 1.2, [1.1, 1.3, 1.4, 7.43, 1.3], [0.3] * len(cells), cells, name="revision")

def slide_qna(prs, page, rows, part=""):
    sl = content_slide(prs, page, "요구사항 및 질문답변" + part)
    rows = rows or [[""] * 8 for _ in range(3)]
    cells = hdr_cells(["Date/V", "중요도", "페이지", "항목 및 기능", "Notes", "요청자", "의견 및 수용 여부", "답변자"], rows, None,
                      align="cccllcl" "c")
    table(sl, 0.40, 1.2, [1.1, 0.7, 0.7, 1.9, 3.9, 0.9, 2.53, 0.8], [0.3] * len(cells), cells, name="qna")

def slide_overview(prs, page, spec, stats):
    ov = spec.get("overview", {})
    sl = content_slide(prs, page, "문서 개요 · 작성 기준", "목적 / 근거 문서 / 범위 / 표기 규칙")
    txt(sl, 0.40, 1.15, 6.2, 0.25, "작성 목적", size=11, bold=True)
    txt(sl, 0.40, 1.42, 6.2, 0.9, ov.get("purpose", "[TBD]"), size=9, lnspc=120)
    txt(sl, 0.40, 2.40, 6.2, 0.25, "근거 · 기준 문서", size=11, bold=True)
    basis = ov.get("basis") or [["기능정의서", DOC.get("source", "[TBD]")]]
    cells = [[{"text": a, "fill": C["lab"], "bold": True, "align": "c"}, {"text": b}] for a, b in basis]
    table(sl, 0.40, 2.68, [1.3, 4.9], [0.28] * len(cells), cells, name="basis")
    # 설계 범위
    txt(sl, 6.95, 1.15, 6.0, 0.25, "설계 범위", size=11, bold=True)
    tiles = [(stats["screens"], "화면"), (stats["components"], "컴포넌트"), (stats["detail"], "상세 장표"), (stats["excluded"], "제외")]
    for i, (n, lab) in enumerate(tiles):
        x = 6.95 + i * 1.5
        box(sl, x, 1.45, 1.38, 0.8, fill="F7F7F7", line="D9D9D9", text=[{"text": str(n), "size": 20, "bold": True},
            {"text": lab, "size": 9, "color": C["sub"]}], align="c")
    mode = {"ADM": "관리자(Admin, 1280px 데스크톱)", "MO": "모바일(393×852)"}.get(DOC.get("mode"), DOC.get("mode"))
    txt(sl, 6.95, 2.35, 6.0, 0.4, [f"모드: {mode}"] + ov.get("scope", []), size=9, color=C["sub"])
    # 표기 규칙
    txt(sl, 6.95, 3.05, 6.0, 0.25, "표기 규칙", size=11, bold=True)
    y = 3.35
    for k, col in STATUS.items():
        box(sl, 6.95, y, 0.62, 0.2, fill=col, shape="round", text=k, size=8, bold=True, color="FFFFFF")
        desc = {"신규": "신규 화면 / 신규 영역·컬럼·버튼", "수정": "기존 화면의 변경 영역", "확인필요": "구현 범위 확인 후 개발 여부 확정",
                "정책변경": "기존 업무 정책이 바뀌는 항목", "기존": "원본 화면을 그대로 옮김"}[k]
        txt(sl, 7.7, y, 5.2, 0.2, desc, size=8.5, valign="m")
        y += 0.27
    box(sl, 6.95, y + 0.02, 0.2, 0.2, fill=C["marker"], line="FFFFFF", lw=1, shape="oval", text="1", size=8, bold=True, color="FFFFFF")
    txt(sl, 7.7, y + 0.02, 5.2, 0.2, "번호 마커 = 오른쪽 Description 번호와 대응", size=8.5, valign="m")
    y += 0.29
    box(sl, 6.95, y, 0.62, 0.2, fill=C["memo"], line=C["memoB"], text="메모", size=7)
    txt(sl, 7.7, y, 5.2, 0.2, "주황 마커 · 노란 메모 · [TBD 참조ID] = 미확정, 선행 결정사항에서 관리", size=8.5, valign="m")
    y += 0.29
    txt(sl, 6.95, y, 0.8, 0.2, "v2.88 문구", size=7, color=C["marker"], valign="m")
    txt(sl, 7.7, y, 5.2, 0.2, "버전 변경 메모(빨간 작은 글자)", size=8.5, valign="m")
    y += 0.4
    txt(sl, 6.95, y, 6.0, 0.25, "화면ID · 번호 규칙", size=11, bold=True)
    rules = ov.get("idRules") or ["화면ID·Component ID는 기능정의서 값을 그대로 사용",
                                   "번호는 Description 기준, 한 화면이 여러 장표로 나뉘면 이어서 부여"]
    txt(sl, 6.95, y + 0.28, 6.0, 1.2, [{"text": r, "bullet": 1} for r in rules], size=9, lnspc=120)

# ───────────────────────── 화면 목록 ─────────────────────────
TYPE_MARK = {"팝업": "팝업 [P]", "탭": "탭 [T]", "바텀시트": "바텀시트 [B]"}

def screen_rows(all_screens):
    rows = []
    for i, s in enumerate(all_screens, 1):
        path = (s.get("path") or []) + ["", "", ""]
        pg = PAGES.get(s["key"], [])
        pr = "-" if not pg else (str(pg[0]) if len(pg) == 1 else f"{pg[0]}~{pg[-1]}")
        typ = TYPE_MARK.get(s.get("type", "페이지"), s.get("type", "페이지"))
        rows.append([i, path[0], path[1], path[2], s["name"], s.get("screenId") or "-", typ, s.get("status", "기존"),
                     s.get("req") or "-", pr])
    return rows

def slide_screen_list(prs, page, rows, part, excluded=None):
    sl = content_slide(prs, page, "화면 목록" + part)
    heads = ["No", "1Depth", "2Depth", "3Depth", "화면명", "화면ID", "유형", "구분", "요구사항ID", "페이지"]
    cells = hdr_cells(heads, rows, None, align="cccclcccc" "c")
    for r, row in enumerate(rows, 1):
        st = row[7]
        if st in STATUS:
            cells[r][7]["color"] = STATUS[st]; cells[r][7]["bold"] = True
    table(sl, 0.40, 1.15, [0.45, 1.25, 1.35, 1.35, 2.95, 2.2, 0.95, 0.75, 0.75, 0.53], [0.25] * len(cells), cells, name="screen-list")
    if excluded:
        y = 1.15 + 0.25 * len(cells) + 0.3
        txt(sl, 0.40, y, 4, 0.25, "제외 화면", size=10, bold=True)
        ec = hdr_cells(["화면명", "화면ID", "제외 사유"], excluded, None, align="lcl")
        table(sl, 0.40, y + 0.3, [3.0, 2.2, 7.33], [0.25] * len(ec), ec, name="excluded")

# ───────────────────────── 플로우 ─────────────────────────
def slide_flow(prs, page, flow, by_key):
    sl = content_slide(prs, page, flow.get("title", "플로우"), flow.get("sub"))
    CUR["collect"] = True
    nodes = flow.get("nodes", [])
    bw, bh = 1.75, 0.72
    has_table = bool(flow.get("triggers"))
    area_w = 8.6 if has_table else 12.5
    per_row = flow.get("perRow") or max(1, int((area_w + 0.5) // (bw + 0.55)))
    pos = {}
    for i, n in enumerate(nodes):
        if isinstance(n, str):
            n = {"key": n}
        col, row = n.get("grid", [i % per_row, i // per_row])
        x = 0.5 + col * (bw + 0.55); y = 1.35 + row * (bh + 0.6)
        key = n["key"]
        s = by_key.get(n.get("screen", key))
        name = n.get("label") or (s["name"] if s else key)
        sid = n.get("id") or ((s.get("screenId") or "-") if s else "")
        pg = PAGES.get(n.get("screen", key))
        ptxt = "" if not pg else f"(p.{pg[0]})"
        st = s.get("status") if s else None
        box(sl, x, y, bw, bh, fill="FFFFFF", line=STATUS.get(st, "7F7F7F") if st in ("신규", "수정", "확인필요", "정책변경") else "7F7F7F",
            lw=1, shape="round", radius=0.06, name=f"flow-{key}",
            text=[{"text": name, "size": 9, "bold": True}, {"text": sid, "size": 8, "color": C["cid"]}, {"text": ptxt, "size": 8, "color": C["sub"]}])
        pos[key] = (x, y)
    for e in flow.get("edges", []):
        a, b = e[0], e[1]
        lab = e[2] if len(e) > 2 else ""
        if a not in pos or b not in pos:
            warn(f"플로우 연결 대상 없음: {a}->{b}"); continue
        (ax, ay), (bx, by) = pos[a], pos[b]
        if abs(ay - by) < 0.01:
            x1, y1 = (ax + bw, ay + bh / 2) if bx > ax else (ax, ay + bh / 2)
            x2, y2 = (bx, by + bh / 2) if bx > ax else (bx + bw, by + bh / 2)
        else:
            x1, y1 = ax + bw / 2, (ay + bh if by > ay else ay)
            x2, y2 = bx + bw / 2, (by if by > ay else by + bh)
        tbd = bool(REF_RE.search(lab)) or "TBD" in lab
        hline(sl, x1, y1, x2, y2, C["tbd"] if tbd else "595959", 1, arrow=True, dash=tbd)
        if lab:
            txt(sl, (x1 + x2) / 2 - 0.5, (y1 + y2) / 2 - 0.19, 1.0, 0.16, lab, size=7.5, align="c", color=C["marker"] if tbd else C["sub"])
    if has_table:
        rows = flow["triggers"]
        cells = hdr_cells(["ID", "트리거", "이동"], rows, None, align="cll")
        table(sl, 9.3, 1.3, [0.55, 1.85, 1.23], [0.26] * len(cells), cells, name="flow-trigger", size=7.5)
    for i, note in enumerate(flow.get("notes", [])):
        txt(sl, 0.5, 6.55 + i * 0.2, 8.6, 0.2, note, size=8, color=C["sub"])
    CUR["collect"] = False

def slide_divider(prs, page, title, sub=None):
    sl = plain_slide(prs, page)
    txt(sl, 1.0, 3.0, 11.33, 0.7, title, size=28, bold=True, align="c", valign="m")
    if sub:
        txt(sl, 1.0, 3.75, 11.33, 0.4, sub, size=12, align="c", color=C["sub"])

def slide_policy(prs, page, pol):
    sl = content_slide(prs, page, pol.get("title", "정책 · 상태값"), pol.get("sub"))
    CUR["collect"] = True
    y = 1.2
    for tb in pol.get("tables", []):
        if tb.get("caption"):
            txt(sl, 0.40, y, 12, 0.25, tb["caption"], size=10, bold=True); y += 0.3
        widths = tb.get("widths") or [12.53 / len(tb["cols"])] * len(tb["cols"])
        cells = hdr_cells(tb["cols"], tb["rows"], None, align=tb.get("align") or "l" * len(tb["cols"]))
        heights = [0.27] + [max(0.27, 0.06 + max(n_lines(v, 8, widths[i] - 0.1) for i, v in enumerate(r)) * line_h(8)) for r in tb["rows"]]
        table(sl, 0.40, y, widths, heights, cells, name="policy")
        y += sum(heights) + 0.3
    for n in pol.get("notes", []):
        txt(sl, 0.40, y, 12.5, 0.2, n, size=8.5, color=C["sub"]); y += 0.22
    CUR["collect"] = False

def slide_decisions(prs, page, rows, part):
    sl = content_slide(prs, page, "선행 결정사항" + part, "미확정 항목 — 참조ID별 확인 주체와 결정 필요 내용")
    cells = hdr_cells(["ID", "구분", "화면ID", "페이지", "내용", "확인 주체"], rows, None, align="cccclc")
    for r, row in enumerate(rows, 1):
        cells[r][1]["color"] = C["tbd"]; cells[r][1]["bold"] = True
    heights = [0.27] + [max(0.3, 0.06 + n_lines(r[4], 8, 6.9) * line_h(8)) for r in rows]
    table(sl, 0.40, 1.15, [0.7, 0.95, 1.9, 0.9, 7.08, 1.0], heights, cells, name="decisions")

# ───────────────────────── Description 추정·분할 ─────────────────────────
def status_tag(item):
    text = " ".join(item.get("lines", []))
    for t in ("결정필요", "개발확인", "TBD"):
        if f"[{t}" in text:
            return f"[{t}]"
    return "[TBD]" if item.get("tag") else ""

def item_head(item):
    h = f"[{item['type']}] {item['name']}" if item.get("type") else item.get("name", "")
    return h

def item_h(item):
    head = item_head(item) + (" " + status_tag(item) if status_tag(item) else "") + (f" ({item['status']})" if item.get("status") else "")
    h = 0.075 + n_lines(head, 8, 2.28) * line_h(8)
    if item.get("cid"):
        h += n_lines(item["cid"], 7.5, 2.28) * line_h(7.5)
    for ln in item.get("lines", []):
        sub = ln.startswith("  ")
        h += n_lines(ln.strip(), 8, 2.00 if sub else 2.14) * line_h(8)
    return h * 1.12

def split_items(items, first_h, cont_h, limit=6.05):
    hs = [item_h(i) for i in items]
    def greedy(L):
        pages, cur, used = [], [], first_h
        for it, h in zip(items, hs):
            if cur and used + h > L:
                pages.append(cur); cur, used = [], cont_h
            cur.append(it); used += h
        pages.append(cur)
        return pages
    base = greedy(limit)
    n = len(base)
    if n == 1:
        return base
    lo, hi = max(hs) + cont_h, limit
    best = base
    for _ in range(25):
        mid = (lo + hi) / 2
        p = greedy(mid)
        if len(p) <= n:
            best, hi = p, mid
        else:
            lo = mid
    return best

def wire_nos(obj, acc):
    if isinstance(obj, dict):
        for k, v in obj.items():
            if k == "no" and v not in (None, ""):
                acc.add(str(v))
            elif k == "mk" and isinstance(v, dict):
                acc.update(str(x) for x in v.keys())
            else:
                wire_nos(v, acc)
    elif isinstance(obj, list):
        for v in obj:
            wire_nos(v, acc)
    return acc

def plan_screen(s):
    """→ list of page dicts {screen, part, items, first}"""
    wires = s.get("wire")
    parts = wires if isinstance(wires, list) else [wires or {}]
    items = [i for i in s.get("items", []) if str(i.get("no")) != "0"]
    for it in items:
        it["no"] = str(it["no"])
    groups = [[] for _ in parts]
    if len(parts) > 1:
        sets = [wire_nos(p, set()) for p in parts]
        for it in items:
            idx = next((k for k, st in enumerate(sets) if it["no"] in st), 0)
            groups[idx].append(it)
    else:
        groups[0] = items
    ov = s.get("overview", [])
    first_h = 1.12 * (0.075 + line_h(8) + sum(n_lines(x, 8, 2.14) * line_h(8) for x in ov))
    cont_h = 1.12 * (0.075 + line_h(8) + 2 * line_h(8))
    pages = []
    for k, (part, grp) in enumerate(zip(parts, groups)):
        chunks = split_items(grp, first_h if not pages else cont_h, cont_h) if grp else [[]]
        for ch in chunks:
            pages.append({"screen": s, "part": part, "partIdx": k, "items": ch, "first": not pages})
    return pages

# ───────────────────────── 화면 상세: 헤더·Description ─────────────────────────
def fit_size(text, width, sizes=(10.5, 9.5, 9)):
    for sz in sizes:
        if text_w(text, sz) <= width - 0.12:
            return sz, text
    return sizes[-1], None

def header_table(sl, s, page, k, n, suffix=""):
    path = [p for p in (s.get("path") or []) if p]
    name = s["name"] + suffix + (f" ({k}/{n})" if n > 1 else "")
    full = " > ".join(path + [name])
    sz, ok = fit_size(full, 5.60)
    if ok is None:
        mid = path[:1] + ["…"] + path[-1:] if len(path) > 2 else path
        full = " > ".join(mid + [name]); sz = 9
    sid = s.get("screenId") or "-"
    sid_sz, _ = fit_size(sid, 1.60)
    req = s.get("req") or "-"
    if text_w(req, 10.5) > 1.55 - 0.12:
        parts = [p.strip() for p in req.split(",")]
        req = f"{parts[0]} 외 {len(parts)-1}건"
    H = lambda t: {"text": t, "fill": C["head"], "bold": True, "align": "c", "size": 10.5}
    V = lambda t, a="c", z=10.5: {"text": t, "align": a, "size": z}
    upd = s.get("updated") or DOC.get("date", "")
    cells = [[H("프로젝트명"), V(DOC.get("project", ""), "l"), H("화면ID"), V(sid, "c", sid_sz), H("요구사항"), V(req), H("마지막 업데이트일")],
             [H("화면명"), V(full, "l", sz), H("작성자"), V(s.get("author") or DOC.get("author") or "-"), H("페이지"), V(str(page)), V(upd)]]
    table(sl, 0, 0, [0.90, 5.60, 0.85, 1.60, 0.85, 1.55, 1.983], [0.33, 0.33], cells, name="page-header", border="000000", bw=0.5)

def desc_table(sl, s, pg_items, first, first_page):
    rows = []
    if first:
        paras = [{"text": "화면설명", "bold": True, "lnspc": 120}] + [{"text": x, "bullet": 1, "lnspc": 120} for x in s.get("overview", [])]
    else:
        paras = [{"text": "화면설명", "bold": True, "lnspc": 120},
                 {"text": f"p.{first_page}에서 이어짐 — 화면 목적·진입 경로는 p.{first_page} 참조", "bullet": 1, "lnspc": 120}]
    rows.append(("0", False, paras))
    for it in pg_items:
        tag = status_tag(it)
        runs = [[item_head(it), {"bold": True}]]
        if tag:
            runs.append([" " + tag, {"bold": True, "color": C["tbd"]}])
        if it.get("status"):
            runs.append([f" ({it['status']})", {"bold": True, "color": STATUS.get(it["status"], C["text"])}])
        paras = [{"runs": runs, "lnspc": 120}]
        if it.get("cid"):
            paras.append({"text": it["cid"], "size": 7.5, "color": C["cid"], "lnspc": 120})
        for ln in it.get("lines", []):
            sub = ln.startswith("  ")
            paras.append({"text": ln.strip(), "bullet": 2 if sub else 1, "lnspc": 120})
        rows.append((it["no"], bool(it.get("tag")) or bool(tag), paras))
    cells, hs = [], []
    for no, orange, paras in rows:
        cells.append([{"text": [{"text": no, "bold": True, "color": C["tbd"] if orange else C["marker"]}], "align": "c", "valign": "t",
                       "ml": 0.05, "mr": 0.05, "mt": 0.03, "mb": 0.03},
                      {"text": paras, "valign": "t", "ml": 0.05, "mr": 0.05, "mt": 0.03, "mb": 0.03}])
        hs.append(0.2)
    table(sl, 10.60, 0.66, [2.73], [0.29], [[{"text": "Description", "fill": "BFBFBF", "align": "c", "size": 10}]], name="desc-head",
          border="BFBFBF", bw=0.25)
    table(sl, 10.60, 0.946, [0.33, 2.40], hs, cells, name="desc")

# ───────────────────────── 와이어프레임 공통 ─────────────────────────
class Wire:
    def __init__(self, sl, scale, ox, oy, has, tagged):
        self.sl, self.S, self.ox, self.oy = sl, scale, ox, oy
        self.has, self.tagged = has, tagged
        self.anchors = []   # (no, x_in, y_in, kind)
        self.memos = []
        self.n = 0
    def X(self, px): return self.ox + px * self.S
    def Y(self, px): return self.oy + px * self.S
    def W(self, px): return px * self.S
    def nm(self, kind):
        self.n += 1
        return f"wf-{kind}-{self.n}"
    def anchor(self, no, xpx, ypx, kind="ui"):
        if no not in (None, ""):
            self.anchors.append((str(no), self.X(xpx), self.Y(ypx), kind))
    def mark_box(self, xpx, ypx, wpx, hpx, mark):
        col = STATUS.get(mark, C["tbd"]) if mark != True else C["tbd"]
        box(self.sl, self.X(xpx) - 0.02, self.Y(ypx) - 0.02, self.W(wpx) + 0.04, self.W(hpx) + 0.04, line=col, lw=0.75, dash=True,
            name=self.nm("mark"))

def ctrl_group(c):
    if "sel" in c: return 1
    if "btn" in c: return 2
    return 3

def ctrl_w(c, S, font=8):
    if "w" in c: return c["w"]
    if "sel" in c: return max(90, text_w(c["sel"], font) / S + 40)
    if "btn" in c: return max(56, text_w(c["btn"], font) / S + 24)
    if "date" in c: return 120
    if "inp" in c: return 220
    if "area" in c: return 760
    return 100

def draw_ctrl(w: Wire, c, xpx, ypx, hpx=24, font=8):
    sl = w.sl
    wpx = ctrl_w(c, w.S, font)
    x, y, W_, H_ = w.X(xpx), w.Y(ypx), w.W(wpx), w.W(hpx)
    if "btn" in c:
        if c.get("disabled"):
            fill, col, bold = "DDE3EE", C["dis"], False
        elif c.get("primary") or c.get("on"):
            fill, col, bold = C["btnE"], "000000", True
        else:
            fill, col, bold = C["btnN"], "000000", False
        box(sl, x, y, W_, H_, fill=fill, line=C["btnB"], lw=0.75, shape="round", radius=0.04, text=c["btn"], size=font, bold=bold,
            color=col, name=w.nm("btn").replace("wf-btn", "btn"))
    elif "sel" in c:
        box(sl, x, y, W_, H_, fill="FFFFFF", line=C["line"], name=w.nm("sel"), text=[[[c["sel"], {"color": C["dis"] if c.get("ph") else C["text"]}]]],
            size=font, align="l", ml=w.W(6))
        txt(sl, x + W_ - w.W(16), y, w.W(12), H_, "▼", size=font - 1, align="c", valign="m", name=w.nm("t"))
    elif "date" in c:
        vals = c["date"] if isinstance(c["date"], list) else [c["date"]]
        cx = x
        for i, v in enumerate(vals):
            box(sl, cx, y, W_, H_, fill="FFFFFF", line=C["line"], name=w.nm("date"), text=v, size=font, align="l", ml=w.W(6))
            box(sl, cx + W_ - w.W(16), y + (H_ - w.W(10)) / 2, w.W(10), w.W(10), fill="D9D9D9", name=w.nm("cal"))
            cx += W_
            if i < len(vals) - 1:
                txt(sl, cx, y, w.W(16), H_, "~", size=font, align="c", valign="m", name=w.nm("t"))
                cx += w.W(16)
        wpx = (cx - x) / w.S
    elif "area" in c:
        box(sl, x, y, W_, H_, fill="FFFFFF", line=C["line"], name=w.nm("area"), text=[[[c["area"], {"color": C["dis"] if not c.get("filled") else C["text"]}]]],
            size=font, align="l", valign="t", ml=w.W(6))
    else:
        v = c.get("inp", "")
        box(sl, x, y, W_, H_, fill="F2F2F2" if c.get("disabled") else "FFFFFF", line=C["line"], name=w.nm("inp"),
            text=[[[v, {"color": C["text"] if c.get("filled") else C["dis"]}]]], size=font, align="l", ml=w.W(6))
    if c.get("no"):
        w.anchor(c["no"], xpx, ypx)
    if c.get("mark") or c.get("tbd"):
        w.mark_box(xpx, ypx, wpx, hpx, c.get("mark") or True)
    return wpx

def layout_parts(w: Wire, parts, cx0, row_y, row_h, font=8, draw=True, pad=10):
    """셀 한 칸의 내용 배치. 반환 (cell_text, ml_px)."""
    if isinstance(parts, str) or parts is None:
        return parts or "", pad
    lead, ctrls, trail = [], [], []
    for p in parts:
        if isinstance(p, str):
            (trail if ctrls else lead).append(p)
        else:
            if trail:
                warn(f"셀 안 '텍스트 → 요소 → 텍스트 → 요소' 배치: 행을 나누세요 ({parts})")
            ctrls.append(p)
    lead_t = "   ".join(lead)
    if not ctrls:
        return lead_t, pad
    cx = pad + (text_w(lead_t, font) / w.S + 40 if lead_t else 0)
    prev_g = None
    last_r = cx
    gs = [ctrl_group(c) for c in ctrls]
    if any(gs[i] > gs[i + 1] for i in range(len(gs) - 1)) and not any(c.get("free") for c in ctrls):
        warn(f"묶음 폼 요소 순서(선택→컨트롤러→폼요소) 확인: {[list(c.keys())[0] + ':' + str(list(c.values())[0]) for c in ctrls]}")
    for c, g in zip(ctrls, gs):
        if prev_g is not None:
            cx += c.get("gap", 6 if g == prev_g else 16)
        h = c.get("h", 24)
        y = row_y + (row_h - h) / 2
        wpx = draw_ctrl(w, c, cx0 + cx, y, h, font) if draw else ctrl_w(c, w.S)
        cx += wpx
        last_r = cx
        prev_g = g
    trail_t = "   ".join(trail)
    if trail_t:
        if lead_t:
            txt(w.sl, w.X(cx0 + last_r + 20), w.Y(row_y), w.W(400), w.W(row_h), trail_t, size=font, valign="m", name=w.nm("t"))
            return lead_t, pad
        return trail_t, last_r + 20
    return lead_t, pad

# ───────────────────────── 관리자 와이어프레임 ─────────────────────────
def adm_rows(block):
    """단순 행 표기를 generic 행으로 변환"""
    cols = block.get("cols") or ([130, 1150] if block["t"] == "search" else [160, 100, 1020])
    tot = sum(cols)
    cols = [c * 1280 / tot for c in cols]
    dh = block.get("rh", 38 if block["t"] == "search" else 34)
    out = []   # list of (h, [cellspec...], rowmeta)
    ncol = len(cols)
    for r in block["rows"]:
        if "cells" in r:
            out.append((r.get("h", dh), r["cells"], r)); continue
        lab = {"label": r.get("label", ""), "no": r.get("no"), "req": r.get("req")}
        if "rows" in r:
            subs = r["rows"]
            for j, sr in enumerate(subs):
                h = sr.get("h", dh if j == 0 and ncol == 2 else (32 if len(subs) > 1 else dh))
                cells = [dict(lab, rs=len(subs)) if j == 0 else None]
                if ncol >= 3 and "aux" in sr:
                    cells += [{"aux": sr["aux"]}, {"v": sr.get("v"), "cs": ncol - 2}]
                else:
                    cells += [{"v": sr.get("v"), "cs": ncol - 1}]
                meta = dict(sr); meta.setdefault("no", None)
                if j == 0:
                    meta["_rowno"] = r.get("no"); meta["_mark"] = r.get("mark"); meta["_n"] = len(subs)
                    meta.setdefault("memo", r.get("memo")); meta.setdefault("vmemo", r.get("vmemo"))
                out.append((h, cells, meta))
        else:
            h = r.get("h", dh)
            if ncol >= 3 and "aux" in r:
                cells = [lab, {"aux": r["aux"]}, {"v": r.get("v"), "cs": ncol - 2}]
            else:
                cells = [lab, {"v": r.get("v"), "cs": ncol - 1}]
            meta = dict(r); meta["_rowno"] = r.get("no"); meta["_mark"] = r.get("mark"); meta["_n"] = 1
            out.append((h, cells, meta))
    return cols, out

def adm_form(w: Wire, block, ypx):
    cols, rows = adm_rows(block)
    xs = [0]
    for c in cols:
        xs.append(xs[-1] + c)
    rys = [ypx]
    for h, _, _ in rows:
        rys.append(rys[-1] + h)
    grid = []
    for ri, (h, cells, meta) in enumerate(rows):
        line = [None] * len(cols)
        ci = 0
        for cs in cells:
            if cs is None:
                ci += 1; continue
            span = cs.get("cs", 1)
            if "label" in cs:
                lab = cs["label"]
                runs = [[lab.rstrip("*").rstrip(), {"bold": True}]]
                if lab.endswith("*"):
                    runs.append([" *", {"bold": True, "color": C["marker"]}])
                spec = {"text": [runs], "fill": C["lab"], "ml": w.W(34), "cs": span, "rs": cs.get("rs", 1)}
                if cs.get("no"):
                    w.anchor(cs["no"], xs[ci], rys[ri], "label")
            elif "aux" in cs:
                spec = {"text": cs["aux"], "ml": w.W(10), "cs": span}
            else:
                cw = xs[ci + span] - xs[ci]
                text, ml = layout_parts(w, cs.get("v"), xs[ci], rys[ri], h)
                spec = {"text": text, "ml": w.W(ml), "cs": span}
                if cs.get("hl"):
                    spec["fill"] = C["hl"]
                if cs.get("grey"):
                    spec["color"] = C["grey"]
            line[ci] = spec
            ci += span
        grid.append(line)
    # 표를 먼저 그리고 오버레이를 위로 올린다
    gf = table(w.sl, w.X(0), w.Y(ypx), [w.W(c) for c in cols], [w.W(h) for h, _, _ in rows], grid, name=w.nm("form").replace("wf-form", "wf-form"), size=8)
    sp = gf._element
    tree = sp.getparent()
    tree.remove(sp)
    # 표를 오버레이보다 아래(먼저)로: 현재 도형 목록에서 첫 오버레이 위치 앞에 삽입
    tree.insert(w.block_start, sp)
    # 메모·표시 박스
    for ri, (h, cells, meta) in enumerate(rows):
        if meta.get("_mark"):
            n = meta.get("_n", 1)
            w.mark_box(0, rys[ri], 1280, sum(rows[ri + k][0] for k in range(n)), meta["_mark"])
        rowno = meta.get("_rowno") or meta.get("no")
        if meta.get("vmemo"):
            txt(w.sl, w.X(900), w.Y(rys[ri]), w.W(370), w.W(h), meta["vmemo"], size=7, color=C["marker"], valign="m", name=w.nm("vmemo"))
        if meta.get("memo") and (not rowno or w.has(str(rowno))):
            w.memos.append((900, rys[ri] + (h - 20) / 2, 370, 20, meta["memo"]))
    return rys[-1]

def adm_list(w: Wire, b, ypx):
    cols = b["cols"]
    names = [c[0] for c in cols]
    tot = sum(c[1] for c in cols)
    ws = [c[1] * 1280 / tot for c in cols]
    xs = [0]
    for c in ws:
        xs.append(xs[-1] + c)
    mk = b.get("mk") or {}
    if mk:
        ypx += 26
    links = set()
    for l in b.get("link", []):
        links.add(names.index(l) if isinstance(l, str) else l)
    hl = set(names.index(l) if isinstance(l, str) else l for l in b.get("hl", []))
    hh, rh = b.get("hh", 30), b.get("rh", 48)
    cells = [[{"text": n, "fill": C["hl"] if i in hl else C["lab"], "bold": True, "align": "c", "ml": 0.02, "mr": 0.02} for i, n in enumerate(names)]]
    for r in b.get("rows", []):
        line = []
        for i, v in enumerate(r):
            v = "" if v is None else str(v)
            if i in links and v not in ("", "-"):
                t = [[[x, {"color": C["link"], "underline": True}]] for x in v.split("\n")]
            else:
                t = v
            line.append({"text": t, "align": "c", "ml": 0.02, "mr": 0.02, "fill": C["hl"] if i in hl else None})
        cells.append(line)
    rows_h = [hh] + [rh] * len(b.get("rows", []))
    gf = table(w.sl, w.X(0), w.Y(ypx), [w.W(x) for x in ws], [w.W(h) for h in rows_h], cells, name=w.nm("list"), size=8)
    for no, col in mk.items():
        ci = names.index(col) if isinstance(col, str) else col
        w.anchors.append((str(no), w.X(xs[ci] + 16), w.Y(ypx) - 0.10, "col"))
    if b.get("no"):
        w.anchor(b["no"], 0, ypx)
    if b.get("mark"):
        w.mark_box(0, ypx, 1280, sum(rows_h), b["mark"])
    return ypx + sum(rows_h)

def adm_row_ctrls(w: Wire, items, ypx, align="c", h=30, xstart=0, xend=1280):
    widths = [ctrl_w(c, w.S) for c in items]
    gap = 8
    total = sum(widths) + gap * (len(items) - 1)
    x = {"c": xstart + (xend - xstart - total) / 2, "l": xstart, "r": xend - total}[align]
    prim = [c for c in items if c.get("primary")]
    if len(prim) > 1:
        warn(f"버튼 묶음에 강조 버튼 {len(prim)}개: {[c['btn'] for c in prim]}")
    for c, wd in zip(items, widths):
        c2 = dict(c); c2["w"] = wd
        draw_ctrl(w, c2, x, ypx + (h - c.get("h", h)) / 2, c.get("h", h))
        x += wd + gap

def draw_adm(w: Wire, part, s, has):
    sl = w.sl
    title = part.get("title", s["name"].split("(")[0])
    crumb = part.get("crumb", " > ".join([p for p in (s.get("path") or []) if p]))
    txt(sl, w.X(0), w.Y(0), w.W(700), w.W(30), title, size=10, bold=True, valign="m", name=w.nm("title"))
    txt(sl, w.X(700), w.Y(0), w.W(580), w.W(30), crumb, size=8, color=C["sub"], align="r", valign="m", name=w.nm("crumb"))
    hline(sl, w.X(0), w.Y(36), w.X(1280), w.Y(36), "BFBFBF", 0.75, name=w.nm("line"))
    y = 48
    for b in part.get("blocks", []):
        t = b["t"]
        w.block_start = len(sl.shapes._spTree)  # 표를 넣을 위치(이 블록의 오버레이 앞)
        if t in ("search", "form"):
            y = adm_form(w, b, y) + b.get("after", 16)
        elif t == "list":
            y = adm_list(w, b, y) + b.get("after", 12)
        elif t == "buttons":
            adm_row_ctrls(w, b["items"], y, b.get("align", "c"), b.get("h", 30))
            if b.get("vmemo"):
                txt(sl, w.X(900), w.Y(y), w.W(370), w.W(30), b["vmemo"], size=7, color=C["marker"], valign="m", name=w.nm("vmemo"))
            if b.get("memo") and (not b.get("no") or has(str(b["no"]))):
                w.memos.append((900, y + 5, 370, 20, b["memo"]))
            y += b.get("h", 30) + b.get("after", 14)
        elif t == "caption":
            txt(sl, w.X(0), w.Y(y), w.W(600), w.W(24), b.get("text", ""), size=8, bold=b.get("bold", False), valign="m", name=w.nm("cap"))
            if b.get("no"):
                w.anchor(b["no"], 0, y)
            if b.get("right"):
                adm_row_ctrls(w, b["right"], y, "r", 24)
            y += 24 + b.get("after", 4)
        elif t == "paging":
            txt(sl, w.X(0), w.Y(y), w.W(1280), w.W(22), b.get("text", "◀  1  |  2  |  3  |  4  |  5  ▶"), size=8, align="c", valign="m", name=w.nm("paging"))
            if b.get("no"):
                w.anchor(b["no"], 560, y)
            if b.get("right"):
                adm_row_ctrls(w, b["right"], y, "r", 22)
            y += 22 + b.get("after", 10)
        elif t == "text":
            h = b.get("h", 20)
            txt(sl, w.X(b.get("x", 0)), w.Y(y), w.W(b.get("w", 1280)), w.W(h), b["text"], size=b.get("size", 8), bold=b.get("bold", False),
                color=b.get("color", C["text"]).lstrip("#"), align=b.get("align", "l"), valign="m", name=w.nm("text"))
            if b.get("no"):
                w.anchor(b["no"], b.get("x", 0), y)
            y += h + b.get("after", 6)
        elif t == "tabs":
            x = 0
            for i, it in enumerate(b["items"]):
                wd = b.get("w", max(100, text_w(it, 8) / w.S + 30))
                on = i == b.get("sel", 0)
                box(sl, w.X(x), w.Y(y), w.W(wd), w.W(30), fill="FFFFFF", line=C["line"], text=it, size=8, bold=on,
                    color=C["link"] if on else C["text"], name=w.nm("tab"))
                if on:
                    hline(sl, w.X(x), w.Y(30 + y), w.X(x + wd), w.Y(30 + y), C["link"], 2, name=w.nm("tabon"))
                x += wd
            if b.get("no"):
                w.anchor(b["no"], 0, y)
            y += 30 + b.get("after", 12)
        elif t == "memo":
            if not b.get("no") or has(str(b["no"])):
                w.memos.append((b.get("x", 0), y, b.get("w", 600), b.get("h", 20), b["text"]))
            y += b.get("h", 20) + 6
        elif t == "gap":
            y += b.get("h", 10)
        elif t == "omit":
            txt(sl, w.X(0), w.Y(y), w.W(1280), w.W(20), "~ 생략 ~", size=8, align="c", color=C["grey"], name=w.nm("omit"))
            y += 26
        else:
            warn(f"알 수 없는 관리자 블록: {t}")
    if y > 760:
        warn(f"{s.get('screenId')} 관리자 와이어프레임 높이 {int(y)}px > 760px — wire를 상단/하단 part로 나누세요")
    if part.get("popup"):
        adm_popup(w, part["popup"], y)
    return y

def adm_popup(w: Wire, p, y_end):
    sl = w.sl
    area_h = min((WY_MAX - OY) / w.S, max(y_end, 500))
    box(sl, w.X(0), w.Y(0), w.W(1280), w.W(area_h), fill="000000", alpha=60, name=w.nm("dim"))
    pw = p.get("w", 560)
    body = p.get("body", [])
    bh = 0
    for ln in body:
        bh += 30 if isinstance(ln, list) else 18 * n_lines(ln, 8, w.W(pw - 60))
    ph = 44 + 20 + bh + 20 + (46 if p.get("buttons") else 0)
    px_ = (1280 - pw) / 2
    py = p.get("y", max(40, (area_h - ph) / 2))
    box(sl, w.X(px_), w.Y(py), w.W(pw), w.W(ph), fill="FFFFFF", line=C["btnB"], shape="round", radius=0.06, name=w.nm("pop"))
    txt(sl, w.X(px_ + 20), w.Y(py + 8), w.W(pw - 60), w.W(28), p.get("title", ""), size=10, bold=True, valign="m", name=w.nm("poptitle"))
    txt(sl, w.X(px_ + pw - 30), w.Y(py + 8), w.W(16), w.W(28), "x", size=10, align="c", valign="m", name=w.nm("popx"))
    hline(sl, w.X(px_), w.Y(py + 44), w.X(px_ + pw), w.Y(py + 44), "D9D9D9", 0.75, name=w.nm("popline"))
    if p.get("no"):
        w.anchor(p["no"], px_, py)
    yy = py + 64
    for ln in body:
        if isinstance(ln, list):
            t_, ml_ = layout_parts(w, ln, px_ + 10, yy, 30)
            if t_:
                txt(sl, w.X(px_ + 10 + ml_), w.Y(yy), w.W(pw - 20 - ml_), w.W(30), t_, size=8, valign="m", name=w.nm("popt"))
            yy += 30
        else:
            hh = 18 * n_lines(ln, 8, w.W(pw - 60))
            txt(sl, w.X(px_ + 30), w.Y(yy), w.W(pw - 60), w.W(hh), ln, size=8, align=p.get("align", "c"), valign="m", name=w.nm("popbody"))
            yy += hh
    if p.get("buttons"):
        adm_row_ctrls(w, p["buttons"], yy + 20, "c", 30, px_, px_ + pw)

# ───────────────────────── 모바일 와이어프레임 ─────────────────────────
MO_W, MO_H = 393, 852

def mo_item_h(it, S):
    t = it.get("t")
    if t == "text":
        sz = it.get("size", 9)
        return n_lines(it["text"], sz, (it.get("w", MO_W - 40)) * S) * sz * 1.25 / 72 / S + 2
    if t == "title":
        return n_lines(it["text"], 14, (MO_W - 40) * S) * 14 * 1.25 / 72 / S + 4
    if t in ("input", "select"):
        return 44 + (20 if it.get("label") else 0)
    if t == "btn": return it.get("h", 52)
    if t == "btns": return it.get("h", 48)
    if t in ("radio", "check"): return 22
    if t == "card":
        return 24 + sum(n_lines(x, 8.5, (MO_W - 72) * S) for x in it.get("lines", [])) * 8.5 * 1.3 / 72 / S
    if t == "list": return 52 * len(it.get("items", []))
    if t == "kv": return 26 * len(it.get("rows", []))
    if t == "image": return it.get("h", 140)
    if t == "tabs": return 44
    if t == "divider": return 1
    if t == "space": return it.get("h", 12)
    if t == "memo": return it.get("h", 24)
    if t == "toast": return 44
    return 30

def mo_item(w: Wire, it, x, y, wd):
    sl, S = w.sl, w.S
    t = it.get("t")
    h = mo_item_h(it, S)
    X, Y, W = w.X, w.Y, w.W
    if t in ("text", "title"):
        sz = it.get("size", 14 if t == "title" else 9)
        txt(sl, X(x), Y(y), W(it.get("w", wd)), W(h), it["text"], size=sz, bold=it.get("bold", t == "title"),
            color=(it.get("color") or C["text"]).lstrip("#"), align=it.get("align", "l"), name=w.nm("text"))
    elif t in ("input", "select"):
        yy = y
        if it.get("label"):
            txt(sl, X(x), Y(y), W(wd), W(18), it["label"] + (" *" if it.get("req") else ""), size=8.5, bold=True, name=w.nm("lab"))
            yy += 20
        v = it.get("val") or it.get("ph", "")
        box(sl, X(x), Y(yy), W(wd), W(44), fill="FFFFFF", line=C["line"], text=[[[v, {"color": C["text"] if it.get("val") else C["dis"]}]]],
            size=9, align="l", ml=W(12), name=w.nm(t))
        if t == "select":
            txt(sl, X(x + wd - 28), Y(yy), W(16), W(44), "▼", size=8, align="c", valign="m", name=w.nm("t"))
    elif t == "btn":
        c = {"btn": it["text"], "primary": it.get("primary"), "disabled": it.get("disabled"), "w": it.get("w", wd)}
        draw_ctrl(w, c, x, y, h, font=10)
    elif t == "btns":
        n = len(it["items"]); gw = (wd - 8 * (n - 1)) / n
        for i, c in enumerate(it["items"]):
            c2 = dict(c); c2["w"] = gw
            draw_ctrl(w, c2, x + i * (gw + 8), y, h, font=10)
    elif t in ("radio", "check"):
        txt(sl, X(x), Y(y), W(wd), W(h), it["text"], size=9, valign="m", name=w.nm(t))
    elif t == "card":
        box(sl, X(x), Y(y), W(wd), W(h), fill="FFFFFF", line=C["line"], shape="round", radius=0.05, name=w.nm("card"))
        txt(sl, X(x + 16), Y(y + 12), W(wd - 32), W(h - 24), it.get("lines", []), size=8.5, name=w.nm("cardt"))
    elif t == "list":
        for i, row in enumerate(it["items"]):
            yy = y + i * 52
            txt(sl, X(x), Y(yy), W(wd - 20), W(52), row, size=9, valign="m", name=w.nm("li"))
            txt(sl, X(x + wd - 16), Y(yy), W(16), W(52), ">", size=9, align="c", valign="m", color=C["grey"], name=w.nm("t"))
            hline(sl, X(x), Y(yy + 52), X(x + wd), Y(yy + 52), "E6E6E6", 0.5, name=w.nm("ln"))
    elif t == "kv":
        for i, (k, v) in enumerate(it["rows"]):
            yy = y + i * 26
            txt(sl, X(x), Y(yy), W(110), W(26), k, size=8.5, color=C["sub"], valign="m", name=w.nm("k"))
            txt(sl, X(x + 110), Y(yy), W(wd - 110), W(26), v, size=8.5, valign="m", name=w.nm("v"))
    elif t == "image":
        box(sl, X(x), Y(y), W(wd), W(h), fill="F2F2F2", line=C["line"], text="Image", size=8, color=C["grey"], name=w.nm("img"))
        hline(sl, X(x), Y(y), X(x + wd), Y(y + h), "D9D9D9", 0.5, name=w.nm("x1"))
        hline(sl, X(x), Y(y + h), X(x + wd), Y(y), "D9D9D9", 0.5, name=w.nm("x2"))
    elif t == "tabs":
        n = len(it["items"]); tw = wd / n
        for i, lab in enumerate(it["items"]):
            on = i == it.get("sel", 0)
            txt(sl, X(x + i * tw), Y(y), W(tw), W(42), lab, size=9, bold=on, color=C["link"] if on else C["text"], align="c", valign="m", name=w.nm("tab"))
            if on:
                hline(sl, X(x + i * tw), Y(y + 43), X(x + (i + 1) * tw), Y(y + 43), C["link"], 2, name=w.nm("tabon"))
        hline(sl, X(x), Y(y + 44), X(x + wd), Y(y + 44), "D9D9D9", 0.5, name=w.nm("tabl"))
    elif t == "divider":
        hline(sl, X(x), Y(y), X(x + wd), Y(y), "D9D9D9", 0.75, name=w.nm("div"))
    elif t == "memo":
        if not it.get("no") or w.has(str(it["no"])):
            w.memos.append((x, y, wd, h, it["text"]))
    elif t == "toast":
        box(sl, X(x), Y(y), W(wd), W(h), fill="333333", shape="round", radius=0.06, text=it["text"], size=9, color="FFFFFF", name=w.nm("toast"))
    elif t != "space":
        warn(f"알 수 없는 모바일 요소: {t}")
    if it.get("no") and t != "memo":
        w.anchor(it["no"], x, y)
    if it.get("mark") or it.get("tbd"):
        w.mark_box(x, y, wd, h, it.get("mark") or True)
    return h

def mo_stack(w: Wire, items, x0, y, wd, ymax):
    for it in items:
        h = mo_item_h(it, w.S)
        if y + h > ymax:
            txt(w.sl, w.X(x0), w.Y(ymax - 16), w.W(wd), w.W(14), "scroll", size=7, align="c", color=C["grey"], name=w.nm("scroll"))
            warn("모바일 프레임 본문 넘침 — 일부 요소 생략(scroll)")
            break
        mo_item(w, it, x0 + it.get("x", 0), y, it.get("w", wd) if it.get("t") not in ("text",) else wd)
        y += h + it.get("after", 12)
    return y

def draw_mo(w_base, part, s, has, tagged, sl):
    frames = part.get("frames", [])
    if len(frames) > 3:
        warn(f"{s.get('screenId')} 모바일 프레임 {len(frames)}개 > 3")
    all_anchors, memos = [], []
    for fi, f in enumerate(frames[:3]):
        fx = OX + fi * (MO_W * S_MO + 0.30)
        w = Wire(sl, S_MO, fx, OY, has, tagged)
        w.n = w_base.n
        X, Y, W = w.X, w.Y, w.W
        if f.get("label"):
            txt(sl, fx, 0.88, 2.6, 0.16, f["label"], size=8, color=C["grey"], name=f"label-{fi}")
        box(sl, X(0), Y(0), W(MO_W), W(MO_H), fill="FFFFFF", line="7F7F7F", lw=1, shape="round", radius=0.20, name=w.nm("frame"))
        txt(sl, X(22), Y(12), W(60), W(20), "9:41", size=8, bold=True, valign="m", wrap=False, name=w.nm("sb"))
        box(sl, X(340), Y(17), W(26), W(10), fill="7F7F7F", shape="round", radius=0.02, name=w.nm("sbb"))
        y = 44
        if f.get("nav", True):
            if f.get("back", True):
                txt(sl, X(12), Y(44), W(24), W(56), "<", size=12, align="c", valign="m", name=w.nm("back"))
            txt(sl, X(50), Y(44), W(MO_W - 100), W(56), f.get("title", ""), size=11, bold=True, align="c", valign="m", name=w.nm("navt"))
            if f.get("right"):
                txt(sl, X(MO_W - 70), Y(44), W(58), W(56), f["right"], size=9, align="r", valign="m", name=w.nm("navr"))
            if f.get("navNo"):
                w.anchor(f["navNo"], 12, 44)
            y = 100
        bottom = MO_H - 34
        if f.get("tabbar"):
            bottom -= 82
            tb = f["tabbar"]
            hline(sl, X(0), Y(bottom), X(MO_W), Y(bottom), "D9D9D9", 0.75, name=w.nm("tabbar"))
            n = len(tb["items"]); tw = MO_W / n
            for i, lab in enumerate(tb["items"]):
                on = i == tb.get("sel", 0)
                txt(sl, X(i * tw), Y(bottom + 8), W(tw), W(40), lab, size=8, bold=on, color=C["link"] if on else C["sub"], align="c", valign="m", name=w.nm("tbi"))
            if tb.get("no"):
                w.anchor(tb["no"], 0, bottom)
        fixed = f.get("bottom", [])
        fixed_h = sum(mo_item_h(it, S_MO) for it in fixed) + 12 * max(0, len(fixed) - 1)
        body_max = bottom - (fixed_h + 16 if fixed else 0)
        mo_stack(w, f.get("body", []), 20, y + f.get("top", 16), MO_W - 40, body_max)
        if fixed:
            mo_stack(w, fixed, 20, bottom - fixed_h - 8, MO_W - 40, bottom)
        box(sl, X((MO_W - 134) / 2), Y(MO_H - 15), W(134), W(5), fill="222222", shape="round", radius=0.02, name=w.nm("home"))
        if f.get("popup"):
            p = f["popup"]
            box(sl, X(0), Y(0), W(MO_W), W(MO_H), fill="000000", alpha=60, shape="round", radius=0.20, name=w.nm("dim"))
            pw = 311
            body = p.get("body", [])
            bh = sum(n_lines(x, 9, W(pw - 48)) for x in body) * 9 * 1.35 / 72 / S_MO
            ph = 28 + (26 if p.get("title") else 0) + bh + 24 + (48 if p.get("buttons") else 0)
            px_, py = (MO_W - pw) / 2, (MO_H - ph) / 2
            box(sl, X(px_), Y(py), W(pw), W(ph), fill="FFFFFF", shape="round", radius=0.08, name=w.nm("pop"))
            yy = py + 24
            if p.get("title"):
                txt(sl, X(px_ + 24), Y(yy), W(pw - 48), W(22), p["title"], size=10, bold=True, align="c", name=w.nm("popt")); yy += 26
            txt(sl, X(px_ + 24), Y(yy), W(pw - 48), W(bh), body, size=9, align="c", name=w.nm("popb")); yy += bh + 20
            if p.get("no"):
                w.anchor(p["no"], px_, py)
            if p.get("buttons"):
                n = len(p["buttons"]); gw = (pw - 40 - 8 * (n - 1)) / n
                for i, c in enumerate(p["buttons"]):
                    c2 = dict(c); c2["w"] = gw
                    draw_ctrl(w, c2, px_ + 20 + i * (gw + 8), yy, 40, font=9)
        if f.get("sheet"):
            sh = f["sheet"]
            box(sl, X(0), Y(0), W(MO_W), W(MO_H), fill="000000", alpha=60, shape="round", radius=0.20, name=w.nm("dim"))
            hh = sh.get("h", 360)
            box(sl, X(0), Y(MO_H - hh), W(MO_W), W(hh), fill="FFFFFF", shape="top_round", radius=0.12, name=w.nm("sheet"))
            box(sl, X((MO_W - 40) / 2), Y(MO_H - hh + 8), W(40), W(4), fill="BFBFBF", shape="round", radius=0.01, name=w.nm("handle"))
            if sh.get("title"):
                txt(sl, X(20), Y(MO_H - hh + 20), W(MO_W - 40), W(24), sh["title"], size=10, bold=True, name=w.nm("sht"))
            if sh.get("no"):
                w.anchor(sh["no"], 0, MO_H - hh)
            mo_stack(w, sh.get("body", []), 20, MO_H - hh + 52, MO_W - 40, MO_H - 20)
        all_anchors += w.anchors
        memos += [(w, m) for m in w.memos]
        w_base.n = w.n
    w_base.anchors = all_anchors
    return memos

# ───────────────────────── 화면 상세 장표 ─────────────────────────
def draw_markers(sl, anchors, page_nos, tagged, frame_x0=OX, sid=""):
    seen = set()
    for no, x, y, kind in anchors:
        if no not in page_nos or no in seen:
            continue
        seen.add(no)
        if "-" in no:
            w, h = 0.33, 0.25
        elif len(no) >= 2:
            w, h = 0.26, 0.20
        else:
            w, h = 0.20, 0.20
        left, top = x - w / 2, y - h / 2
        if left < frame_x0:
            left = frame_x0 + 0.04
        if top < 0.95:
            top = 0.95
        col = C["tbd"] if no in tagged else C["marker"]
        box(sl, left, top, w, h, fill=col, line="FFFFFF", lw=1, shape="oval" if w == h else "round", radius=0.1, text=no, size=8,
            bold=True, color="FFFFFF", name=f"marker-{no}")
    missing = page_nos - seen
    return missing

def slide_detail(prs, page, pg, k, n, first_page):
    s = pg["screen"]; part = pg["part"] or {}
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    CUR["collect"] = True
    suffix = part.get("suffix", "")
    header_table(sl, s, page, k, n, suffix)
    st = s.get("status", "기존")
    box(sl, 9.83, 0.69, 0.62, 0.20, fill=STATUS.get(st, "7F7F7F"), shape="round", radius=0.05, text=st, size=8, bold=True, color="FFFFFF", name="badge")
    page_nos = set(it["no"] for it in pg["items"])
    tagged = set(str(it["no"]) for it in s.get("items", []) if it.get("tag") or status_tag(it))
    has = lambda no: str(no) in page_nos
    mode = s.get("mode") or DOC.get("mode", "ADM")
    w = Wire(sl, S_ADM if mode == "ADM" else S_MO, OX, OY, has, tagged)
    memos = []
    if mode == "ADM":
        lab = part.get("frameLabel", s.get("frameLabel", ""))
        if lab:
            txt(sl, 0.30, 0.88, 6.0, 0.16, lab, size=8, color=C["grey"], name="frame-label")
        if part.get("blocks") or part.get("popup"):
            draw_adm(w, part, s, has)
        memos = [(w, m) for m in w.memos]
    else:
        memos = draw_mo(w, part, s, has, tagged, sl)
    for ww, (x, y, wd, h, t) in memos:
        box(sl, ww.X(x), ww.Y(y), ww.W(wd), ww.W(h), fill=C["memo"], line=C["memoB"], lw=0.75, text=t, size=7, align="l", ml=0.04,
            name=ww.nm("memo"))
    missing = draw_markers(sl, w.anchors, page_nos, tagged, OX, s.get("screenId"))
    for m in sorted(missing):
        warn(f"p.{page} {s.get('screenId')} #{m}: 와이어프레임에 마커 위치(no)가 없음")
    anchor_nos = set(a[0] for a in w.anchors)
    all_nos = set(str(i["no"]) for i in s.get("items", []))
    for a in sorted(anchor_nos - all_nos):
        warn(f"{s.get('screenId')}: 와이어프레임 no '{a}'에 해당하는 Description 항목 없음")
    desc_table(sl, s, pg["items"], pg["first"], first_page)
    hline(sl, 0, 7.215, SW, 7.215, "D9D9D9", 0.75, name="bottom-line")
    CUR["collect"] = False

# ───────────────────────── 덱 계획 ─────────────────────────
def chunk(lst, n):
    return [lst[i:i + n] for i in range(0, len(lst), n)] or [[]]

def norm_scope(v):
    v = str(v or "full").strip().lower()
    if v in ("partial", "해당화면", "일부화면", "일부", "해당"):
        return "partial"
    if v in ("full", "전체화면", "전체"):
        return "full"
    sys.exit(f"scope 값 오류: {v} (full|partial)")

def filter_targets(spec, targets):
    """targets(화면 key 또는 screenId)에 해당하는 화면만 남김. 빈 섹션은 간지째 제외."""
    want = set(map(str, targets))
    hit = set()
    def keep(s):
        ok = s["key"] in want or str(s.get("screenId")) in want
        if ok:
            hit.update({s["key"], str(s.get("screenId"))})
        return ok
    spec["common"] = [s for s in spec.get("common", []) if keep(s)]
    secs = []
    for sec in spec.get("sections", []):
        scr = [s for s in sec.get("screens", []) if keep(s)]
        if scr:
            secs.append({**sec, "screens": scr})
    spec["sections"] = secs
    for t in sorted(want - hit):
        warn(f"targets 대상 화면 없음: {t}")
    if not spec["common"] and not secs:
        sys.exit("targets에 해당하는 화면이 없습니다")

def build(spec, out_dir, only=None, xlsx=False, scope=None, targets=None):
    global DOC
    DOC = spec["doc"]
    scope = norm_scope(scope or DOC.get("scope"))
    CUR["scope"] = scope
    targets = targets or DOC.get("targets")
    if targets:
        filter_targets(spec, targets)
    DOC.setdefault("date", datetime.date.today().isoformat())
    DOC.setdefault("mode", "ADM")
    DOC.setdefault("title", "화면설계서(관리자, 웹)" if DOC["mode"] == "ADM" else "화면설계서(모바일)")
    common = spec.get("common", [])
    sections = spec.get("sections", [])
    all_screens = common + [s for sec in sections for s in sec.get("screens", [])]
    keys = set()
    for s in all_screens:
        if s["key"] in keys:
            warn(f"화면 key 중복: {s['key']}")
        keys.add(s["key"])
    by_key = {s["key"]: s for s in all_screens}

    # ── 장표 계획
    # scope: full = 전체 장표 / partial = 화면목록 + 간지 + 와이어프레임 상세만
    full = scope == "full"
    plan = []   # (kind, payload)
    if full:
        plan.append(("cover", None))
        for i, ch in enumerate(chunk(spec.get("revisions", []), 16)):
            plan.append(("rev", ch))
        for ch in chunk(spec.get("qna", []), 10):
            plan.append(("qna", ch))
        plan.append(("overview", None))
    list_chunks = chunk(all_screens, 20)
    for i, ch in enumerate(list_chunks):
        plan.append(("list", (i, len(list_chunks), ch)))
    excl = spec.get("excluded", []) if full else []
    fit_excl = excl and len(list_chunks[-1]) + len(excl) <= 16
    if excl and not fit_excl:
        plan.append(("excl", excl))
    if full:
        for fl in (spec.get("flows") or ([spec["flow"]] if spec.get("flow") else [])):
            plan.append(("flow", fl))
    detail_pages = {}
    if common:
        plan.append(("divider", ("공통 화면", "입력 문서에 근거가 있는 공통 요소")))
        for s in common:
            pgs = plan_screen(s)
            for k, pg in enumerate(pgs, 1):
                plan.append(("detail", (pg, k, len(pgs))))
    for sec in sections:
        if sec.get("title"):
            plan.append(("divider", (sec["title"], sec.get("sub"))))
        for s in sec.get("screens", []):
            pgs = plan_screen(s)
            for k, pg in enumerate(pgs, 1):
                plan.append(("detail", (pg, k, len(pgs))))
    decisions = spec.get("decisions", [])
    if full:
        for pol in spec.get("policies", []):
            plan.append(("policy", pol))
        for ch in chunk(decisions, 12) if decisions else []:
            plan.append(("dec", ch))
    # 페이지 확정
    for i, (kind, pl) in enumerate(plan, 1):
        if kind == "detail":
            PAGES.setdefault(pl[0]["screen"]["key"], []).append(i)
    # ── 그리기 (선행 결정사항 페이지 수집을 위해 결정사항은 마지막에)
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(SW), Inches(SH)
    stats = {"screens": len(all_screens), "components": sum(len([i for i in s.get("items", []) if str(i.get("no")) != "0"]) for s in all_screens),
             "detail": sum(1 for k, _ in plan if k == "detail"), "excluded": len(excl)}
    nrev = sum(1 for k, _ in plan if k == "rev"); nqna = sum(1 for k, _ in plan if k == "qna"); ndec = sum(1 for k, _ in plan if k == "dec")
    ci = {"rev": 0, "qna": 0, "dec": 0}
    dec_slots = []
    for page, (kind, pl) in enumerate(plan, 1):
        CUR["page"] = page
        if kind == "cover":
            CUR["collect"] = True; slide_cover(prs, page); CUR["collect"] = False
        elif kind == "rev":
            ci["rev"] += 1; slide_revision(prs, page, pl, f" ({ci['rev']}/{nrev})" if nrev > 1 else "")
        elif kind == "qna":
            ci["qna"] += 1; slide_qna(prs, page, pl, f" ({ci['qna']}/{nqna})" if nqna > 1 else "")
        elif kind == "overview":
            slide_overview(prs, page, spec, stats)
        elif kind == "list":
            i, n, ch = pl
            rows = screen_rows(all_screens)[i * 20:(i + 1) * 20]
            slide_screen_list(prs, page, rows, f" ({i+1}/{n})" if n > 1 else "", excl if (fit_excl and i == n - 1) else None)
        elif kind == "excl":
            sl = content_slide(prs, page, "화면 목록 — 제외 화면")
            ec = hdr_cells(["화면명", "화면ID", "제외 사유"], pl, None, align="lcl")
            table(sl, 0.40, 1.15, [3.0, 2.2, 7.33], [0.25] * len(ec), ec, name="excluded")
        elif kind == "flow":
            slide_flow(prs, page, pl, by_key)
        elif kind == "divider":
            slide_divider(prs, page, *pl)
        elif kind == "detail":
            pg, k, n = pl
            slide_detail(prs, page, pg, k, n, PAGES[pg["screen"]["key"]][0])
        elif kind == "policy":
            slide_policy(prs, page, pl)
        elif kind == "dec":
            ci["dec"] += 1
            dec_slots.append((page, pl, f" ({ci['dec']}/{ndec})" if ndec > 1 else ""))
            prs.slides.add_slide(prs.slide_layouts[6])  # 자리만 확보, 뒤에서 채움
    # 선행 결정사항 채우기
    for page, rows, part in dec_slots:
        sl = prs.slides[page - 1]
        full = []
        for r in rows:
            r = list(r) + [""] * (5 - len(r))
            rid, kind, sid, content, owner = r[:5]
            pages = sorted(REFS.get(rid, []))
            full.append([rid, f"[{kind.strip('[]')}]", sid, ", ".join(map(str, pages)) or "-", content, owner])
        # 제목·푸터 포함해서 그리기
        tmp_title = "선행 결정사항" + part
        txt(sl, 0.40, 0.28, 8.5, 0.45, tmp_title, size=18, bold=True, valign="m")
        txt(sl, 0.40, 0.72, 12.5, 0.2, "미확정 항목 — 참조ID별 확인 주체와 결정 필요 내용", size=10, color=C["sub"])
        hline(sl, 0.40, 0.98, 12.93, 0.98, "BFBFBF", 0.75)
        footer(sl, page)
        cells = hdr_cells(["ID", "구분", "화면ID", "페이지", "내용", "확인 주체"], full, None, align="cccclc")
        for ri in range(1, len(cells)):
            cells[ri][1]["color"] = C["tbd"]; cells[ri][1]["bold"] = True
        heights = [0.27] + [max(0.3, 0.06 + n_lines(r[4], 8, 6.9) * line_h(8)) for r in full]
        table(sl, 0.40, 1.15, [0.7, 0.95, 1.9, 0.9, 7.08, 1.0], heights, cells, name="decisions")
    # 참조ID 정합 (§13-10)
    dec_ids = set(r[0] for r in decisions)
    body_ids = set(REFS.keys())
    for rid in sorted(body_ids - dec_ids):
        warn(f"참조ID {rid}: 본문에 있으나 선행 결정사항에 없음 (p.{sorted(REFS[rid])})")
    if full:   # 해당화면 모드는 결정사항 장표가 없고 범위 밖 ID가 정상이라 역방향 점검 생략
        for rid in sorted(dec_ids - body_ids):
            warn(f"참조ID {rid}: 선행 결정사항에만 있고 본문에 없음")

    # ── 저장
    ymd = DOC["date"].replace("-", "")
    fname = DOC.get("fileName") or f"{DOC.get('deliverable', '화면설계서-관리자' if DOC['mode']=='ADM' else '화면설계서-모바일')}_{DOC.get('service','{SERVICE_NAME}')}_{DOC.get('version','v01')}_{ymd}.pptx"
    os.makedirs(out_dir, exist_ok=True)
    total = len(prs.slides)
    if only:
        keep = set(only)
        sldIdLst = prs.slides._sldIdLst
        for idx in reversed(range(total)):
            if idx + 1 not in keep:
                rId = sldIdLst[idx].rId
                prs.part.drop_rel(rId)
                del sldIdLst[idx]
        fname = fname.replace(".pptx", f"_p{'-'.join(map(str, sorted(keep)))}.pptx")
    path = os.path.join(out_dir, fname)
    prs.save(path)
    report = {"file": path, "scope": scope, "slides": total, "screens": len(all_screens), "detailSlides": stats["detail"],
              "refIds": {k: sorted(v) for k, v in REFS.items()}, "pages": PAGES, "warnings": WARN}
    if xlsx:
        report["xlsx"] = write_xlsx(all_screens, out_dir, ymd)
    return path, report, (sorted(only) if only else list(range(1, total + 1)))

def write_xlsx(all_screens, out_dir, ymd):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    wb = Workbook(); ws = wb.active; ws.title = "화면목록"
    heads = ["No", "1Depth", "2Depth", "3Depth", "화면명", "화면ID", "유형", "구분", "요구사항ID", "페이지"]
    ws.append(heads)
    for r in screen_rows(all_screens):
        ws.append(r)
    thin = Side(style="thin", color="BFBFBF")
    for row in ws.iter_rows():
        for c in row:
            c.font = Font(name=FONT, size=10, bold=c.row == 1)
            c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            c.alignment = Alignment(vertical="center", horizontal="left" if c.column == 5 else "center")
            if c.row == 1:
                c.fill = PatternFill("solid", fgColor="F2F2F2")
    for col, wdt in zip("ABCDEFGHIJ", [6, 14, 16, 16, 34, 26, 12, 10, 12, 9]):
        ws.column_dimensions[col].width = wdt
    p = os.path.join(out_dir, f"화면목록_{DOC.get('service','{SERVICE_NAME}')}_{DOC.get('version','v01')}_{ymd}.xlsx")
    wb.save(p)
    return p

# ───────────────────────── 검증 (§13) ─────────────────────────
def verify(path, page_numbers):
    prs = Presentation(path)
    errs = []
    E = 914400
    def runs_of(shape):
        if shape.has_text_frame:
            for p in shape.text_frame.paragraphs:
                for r in p.runs:
                    yield r
        if getattr(shape, "has_table", False) and shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    for p in cell.text_frame.paragraphs:
                        for r in p.runs:
                            yield r
    for sl, pno in zip(prs.slides, page_numbers):
        names = {sh.name: sh for sh in sl.shapes}
        hdr = names.get("page-header")
        if hdr is not None:
            tb = hdr.table
            if len(tb.columns) != 7:
                errs.append(f"p.{pno} 헤더 표 {len(tb.columns)}열 (7열이어야 함)")
            row1 = [tb.cell(0, i).text for i in (0, 2, 4, 6)]
            if row1 != ["프로젝트명", "화면ID", "요구사항", "마지막 업데이트일"]:
                errs.append(f"p.{pno} 헤더 1행 순서 오류 {row1}")
            if tb.cell(1, 5).text.strip() != str(pno):
                errs.append(f"p.{pno} 헤더 페이지 값 {tb.cell(1,5).text} ≠ 실제 {pno}")
            for r in range(2):
                for c in range(7):
                    tc = tb.cell(r, c)._tc
                    if tc.get("gridSpan") or tc.get("rowSpan") or tc.get("hMerge") or tc.get("vMerge"):
                        errs.append(f"p.{pno} 헤더 표 병합 있음"); break
            markers = set(sh.name[7:] for sh in sl.shapes if sh.name.startswith("marker-"))
            desc = names.get("desc")
            dnos = set()
            if desc is not None:
                for row in desc.table.rows:
                    t = row.cells[0].text.strip()
                    if t and t not in ("0", "기타"):
                        dnos.add(t)
            if markers != dnos:
                errs.append(f"p.{pno} 마커 {sorted(markers - dnos)} / Description {sorted(dnos - markers)} 불일치")
        for sh in sl.shapes:
            x, y = sh.left / E, sh.top / E
            w, h = sh.width / E, sh.height / E
            if x < -0.01 or y < -0.01 or x + w > SW + 0.01 or y + h > SH + 0.01:
                errs.append(f"p.{pno} 슬라이드 밖 도형: {sh.name}")
            if sh.name.startswith(("wf-", "btn-")) and hdr is not None:
                if x + w > WX_MAX + 0.02 or y + h > WY_MAX + 0.02 or y < OY - 0.02:
                    errs.append(f"p.{pno} 와이어프레임 영역 이탈: {sh.name} ({x:.2f},{y:.2f},{x+w:.2f},{y+h:.2f})")
            if sh.name.startswith("btn-"):
                try:
                    fc = str(sh.fill.fore_color.rgb)
                except Exception:
                    fc = "none"
                if fc not in ("BFBFBF", "FFFFFF", "DDE3EE"):
                    errs.append(f"p.{pno} 버튼 채움색 {fc}: {sh.name}")
            for r in runs_of(sh):
                if r.font.name != FONT:
                    errs.append(f"p.{pno} 글꼴 {r.font.name}: {r.text[:10]}"); break
                if r.font.size is not None and r.font.size.pt < 6:
                    errs.append(f"p.{pno} 6pt 미만: {r.text[:10]}")
                if sh.name == "desc" and r.font.size is not None and r.font.size.pt < 7.5:
                    errs.append(f"p.{pno} Description 8pt 미만: {r.text[:10]}")
                if EMOJI_RE.search(r.text or ""):
                    errs.append(f"p.{pno} 그림 이모지: {r.text[:10]}")
    # 중복 제거
    seen, out = set(), []
    for e in errs:
        if e not in seen:
            seen.add(e); out.append(e)
    return out

def desc_overflow(path, page_numbers):
    """렌더 없이 Description 높이 재추정 (§13-8 보조)"""
    return []

def render(path, pages, out_dir, dpi=80):
    pdf_dir = os.path.join(out_dir, "_render")
    os.makedirs(pdf_dir, exist_ok=True)
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", pdf_dir, path], capture_output=True, timeout=240)
    pdf = os.path.join(pdf_dir, os.path.splitext(os.path.basename(path))[0] + ".pdf")
    outs = []
    for p in pages:
        prefix = os.path.join(pdf_dir, f"p{p:02d}")
        subprocess.run(["pdftoppm", "-r", str(dpi), "-png", "-f", str(p), "-l", str(p), "-singlefile", pdf, prefix], capture_output=True)
        outs.append(prefix + ".png")
    return outs

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    ap.add_argument("-o", "--out", default=".")
    ap.add_argument("--only", help="예: 13,14 — 해당 장표만 저장(페이지 번호 유지)")
    ap.add_argument("--render", help="all 또는 13,14 — PNG 렌더 (육안 검수용)")
    ap.add_argument("--xlsx", action="store_true", help="화면목록 xlsx 함께 생성")
    ap.add_argument("--scope", help="full(전체화면, 기본) | partial(해당화면: 화면목록·간지·와이어프레임만). doc.scope보다 우선")
    ap.add_argument("--targets", help="해당화면 모드 대상 화면 key 또는 screenId, 쉼표 구분. doc.targets보다 우선")
    a = ap.parse_args()
    spec = json.load(open(a.spec, encoding="utf-8"))
    only = [int(x) for x in a.only.split(",")] if a.only else None
    targets = [t.strip() for t in a.targets.split(",") if t.strip()] if a.targets else None
    path, report, pnums = build(spec, a.out, only, a.xlsx, a.scope, targets)
    errs = verify(path, pnums)
    report["verify"] = errs
    if a.render:
        pages = pnums if a.render == "all" else [int(x) for x in a.render.split(",")]
        idx = [pnums.index(p) + 1 for p in pages if p in pnums]
        imgs = render(path, idx, a.out)
        report["render"] = imgs
    rp = os.path.join(a.out, "report.json")
    json.dump(report, open(rp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"OK {path}")
    print(f"범위 {'전체화면' if report['scope']=='full' else '해당화면'} · 장표 {report['slides']} · 화면 {report['screens']} · 상세 {report['detailSlides']} · 참조ID {len(report['refIds'])}")
    print(f"경고 {len(WARN)}건 · 검증 실패 {len(errs)}건")
    for m in WARN[:40]:
        print("  [경고]", m)
    for m in errs[:40]:
        print("  [검증]", m)
    if report.get("render"):
        print("렌더:", " ".join(report["render"]))

if __name__ == "__main__":
    main()
