# -*- coding: utf-8 -*-
"""paperlib —— 中文论文格式核心库。

依据：
  * 《实践教学论文格式范文.doc》  —— 主参考（字号、字体、参考文献著录格式）
  * 《论文格式模板.docx》          —— 标题样式与多级列表结构
  * 《设计说明书.docx》            —— 多级列表编号层级参照

核心能力：
  1. 生成符合规范���论文模板 .docx（页面、样式、多级列表、页脚页码）
  2. 对既有 .docx 套用规范（样式映射、三线表、字体清理）
  3. 插入图/表题注（SEQ 域，自动编号）
  4. 插入交叉引用（REF 域）
  5. 插入参考文献上标引用（REF 域 + 上标）
  6. 参考文献 / 附录 处分节分页
  7. 中文双引号强制使用中文字体
"""
from __future__ import annotations

import copy
import re
from typing import Iterable, List, Optional, Sequence

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import nsdecls, qn
from docx.oxml.parser import parse_xml
from docx.shared import Cm, Pt
from docx.text.paragraph import Paragraph
from lxml import etree

# --------------------------------------------------------------------------
# 常量
# --------------------------------------------------------------------------

FONT_EN = "Times New Roman"   # 英文/数字西文字体（参考文献除外）
FONT_SONG = "宋体"            # 正文中文
FONT_HEI = "黑体"             # 标题、摘要标签
FONT_KAI = "楷体"             # 摘要/参考文献中的双引号与参考文献条目
FONT_FANGSONG = "仿宋"

SIZE_2HAO = 24    # 小四 12pt
SIZE_3HAO = 32    # 三号 16pt
SIZE_4HAO = 28    # 四号 14pt
SIZE_5HAO = 21    # 五号 10.5pt
SIZE_XIAOWU = 18  # 小五 9pt

CN_QUOTES = ("“", "”")        # 中文左右双引号
CN_QUOTE_SINGLE = ("‘", "’")

STYLE_TITLE = "PaperTitle"
STYLE_SUBTITLE = "PaperSubtitle"
STYLE_ABSTRACT = "PaperAbstract"
STYLE_KEYWORDS = "PaperKeywords"
STYLE_BODY = "PaperBody"
STYLE_REF_HEADING = "PaperRefHeading"
STYLE_REF_CATEGORY = "PaperRefCategory"
STYLE_REF_ITEM = "PaperRefItem"
STYLE_APPENDIX = "PaperAppendix"
STYLE_APPENDIX_HEADING = "PaperAppendixHeading"
STYLE_TABLE_TEXT = "PaperTableText"
STYLE_CAPTION = "Caption"
STYLE_CITE = "PaperCite"       # 字符样式：上标引用序号
STYLE_LABEL = "PaperLabel"     # 字符样式：摘要/关键词等标签（黑体五号）
STYLE_ABSTRACT_EN = "PaperAbstractEN"

ABSTRACT_NUM_ID = 90
NUM_ID = 90

# 页面可用区域（cm）
TEXT_WIDTH_CM = 21.0 - 2 * 3.17     # 14.66
TEXT_HEIGHT_CM = 29.7 - 2 * 2.54    # 24.62

# 图片/图表尺寸上限：不超过版心宽，且高度不超过版心的三分之一
MAX_IMAGE_WIDTH_CM = 14.0
MAX_IMAGE_HEIGHT_CM = 8.0
DEFAULT_CHART_WIDTH_CM = 12.0
DEFAULT_CHART_HEIGHT_CM = 7.0

# 表格高度上限：不超过一版心的 55%，超出则自动缩排
MAX_TABLE_PAGE_FRACTION = 0.55

# 参考文献列表编号（独立于标题多级列表）
REF_ABSTRACT_NUM_ID = 91
REF_NUM_ID = 91


# --------------------------------------------------------------------------
# 底层 XML 工具
# --------------------------------------------------------------------------

def _new(tag: str, **attrs):
    el = OxmlElement(tag)
    for k, v in attrs.items():
        el.set(qn("w:" + k), str(v))
    return el


def _replace_child(parent, tag: str, new_el):
    old = parent.find(qn(tag))
    if old is not None:
        parent.remove(old)
    parent.append(new_el)
    return new_el


def _rpr_first(rpr):
    """rPr 中 rFonts 必须排在 rStyle 之后、其余元素之前。"""
    if rpr is None:
        return None
    style_el = rpr.find(qn("w:rStyle"))
    return style_el


def set_style_rpr(style, xml: str):
    """用原始 XML 整体替换样式/字符样式的 rPr（保证 w:rPr 子元素顺序合法）。"""
    element = style.element
    new = parse_xml("<w:rPr %s>%s</w:rPr>" % (nsdecls("w"), xml))
    old = element.find(qn("w:rPr"))
    if old is not None:
        element.remove(old)
    element.append(new)
    return new


def set_style_ppr(style, xml: str):
    """用原始 XML 整体替换段落样式的 pPr（保证 w:pPr 子元素顺序合法）。"""
    element = style.element
    new = parse_xml("<w:pPr %s>%s</w:pPr>" % (nsdecls("w"), xml))
    old = element.find(qn("w:pPr"))
    if old is not None:
        element.remove(old)
    # pPr 必须排在 rPr 之前
    rpr = element.find(qn("w:rPr"))
    if rpr is not None:
        rpr.addprevious(new)
    else:
        element.append(new)
    return new


def fonts_xml(east_asia: str, ascii_font: Optional[str] = FONT_EN,
              hint: str = "eastAsia") -> str:
    """生成 w:rFonts 片段。ascii/hAnsi 恒为 Times New Roman（参考文献除外）。"""
    a = ascii_font or FONT_EN
    return (
        '<w:rFonts w:ascii="{a}" w:hAnsi="{a}" w:cs="{a}" '
        'w:eastAsia="{e}" w:hint="{h}"/>'.format(a=a, e=east_asia, h=hint)
    )


def sz_xml(size_half_points: int) -> str:
    return '<w:sz w:val="{v}"/><w:szCs w:val="{v}"/>'.format(v=size_half_points)


# --------------------------------------------------------------------------
# 页面与文档默认值
# --------------------------------------------------------------------------

def setup_page(doc):
    """A4 纸张；上下 2.54cm，左右 3.17cm（与三份模板一致）。"""
    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(3.17)
        section.right_margin = Cm(3.17)
        section.header_distance = Cm(1.5)
        section.footer_distance = Cm(1.75)


NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def setup_theme(doc):
    """把主题字体改成 Times New Roman / 宋体。

    Word 在保存时会删除与继承值重复的字体属性；若主题字体是 Calibri/Cambria，
    某些界面（如样式库、插入图表的默认字体）仍会用主题字体。改主题可避免这一点。
    """
    for part in doc.part.package.iter_parts():
        if not str(part.partname).endswith("theme1.xml"):
            continue
        try:
            root = etree.fromstring(part.blob)
        except Exception:
            return
        ns = {"a": NS_A}
        for scheme, latin, ea in (("majorFont", FONT_EN, FONT_HEI),
                                  ("minorFont", FONT_EN, FONT_SONG)):
            f = root.find(".//a:fontScheme/a:" + scheme, ns)
            if f is None:
                continue
            for tag, face in (("latin", latin), ("ea", ea), ("cs", latin)):
                el = f.find("a:" + tag, ns)
                if el is not None:
                    el.set("typeface", face)
        part._blob = etree.tostring(root, xml_declaration=True,
                                    encoding="UTF-8", standalone=True)
    return doc


def setup_doc_defaults(doc):
    """文档默认：西文 Times New Roman，中文宋体，五号，行距 1.5 倍。"""
    styles_el = doc.styles.element
    dd = styles_el.find(qn("w:docDefaults"))
    if dd is None:
        dd = OxmlElement("w:docDefaults")
        styles_el.insert(0, dd)
    rpr_default = dd.find(qn("w:rPrDefault"))
    if rpr_default is None:
        rpr_default = OxmlElement("w:rPrDefault")
        dd.insert(0, rpr_default)
    rpr = rpr_default.find(qn("w:rPr"))
    if rpr is not None:
        rpr_default.remove(rpr)
    rpr = parse_xml(
        "<w:rPr %s>%s%s<w:lang w:val=\"en-US\" w:eastAsia=\"zh-CN\" w:bidi=\"ar-SA\"/>"
        "</w:rPr>" % (nsdecls("w"), fonts_xml(FONT_SONG), sz_xml(SIZE_5HAO))
    )
    rpr_default.append(rpr)

    ppr_default = dd.find(qn("w:pPrDefault"))
    if ppr_default is None:
        ppr_default = OxmlElement("w:pPrDefault")
        dd.append(ppr_default)
    ppr = ppr_default.find(qn("w:pPr"))
    if ppr is not None:
        ppr_default.remove(ppr)
    ppr_default.append(parse_xml(
        "<w:pPr %s><w:widowControl w:val=\"0\"/>"
        "<w:spacing w:line=\"360\" w:lineRule=\"auto\"/>"
        "<w:autoSpaceDE w:val=\"0\"/><w:autoSpaceDN w:val=\"0\"/>"
        "<w:jc w:val=\"both\"/>"
        "</w:pPr>" % nsdecls("w")
    ))


# --------------------------------------------------------------------------
# 样式表
# --------------------------------------------------------------------------

def _ensure_style(doc, name: str, style_type, base: Optional[str] = None):
    from docx.enum.style import WD_STYLE_TYPE
    try:
        return doc.styles[name]
    except KeyError:
        st = doc.styles.add_style(name, style_type)
        if base:
            st.base_style = doc.styles[base]
        return st


def _ppr(align: str = "both", line: int = 360, before: int = 0, after: int = 0,
         first_line_chars: Optional[int] = None, first_line: Optional[int] = None,
         left: int = 0, hanging: Optional[int] = None,
         keep_next: bool = False, keep_lines: bool = False,
         outline: Optional[int] = None, numid: Optional[int] = None,
         ilvl: Optional[int] = None, indent_chars: Optional[int] = None) -> str:
    bits = []
    if keep_next:
        bits.append("<w:keepNext/>")
    if keep_lines:
        bits.append("<w:keepLines/>")
    if numid is not None:
        bits.append('<w:numPr><w:ilvl w:val="%d"/><w:numId w:val="%d"/></w:numPr>'
                    % (ilvl or 0, numid))
    bits.append('<w:spacing w:before="%d" w:after="%d" w:line="%d" w:lineRule="auto"/>'
                % (before, after, line))
    ind_bits = []
    if left:
        ind_bits.append('w:left="%d"' % left)
    if hanging is not None:
        ind_bits.append('w:hanging="%d"' % hanging)
    if first_line_chars is not None:
        ind_bits.append('w:firstLineChars="%d"' % first_line_chars)
        if first_line is not None:
            ind_bits.append('w:firstLine="%d"' % first_line)
    elif first_line is not None:
        ind_bits.append('w:firstLine="%d"' % first_line)
    if indent_chars is not None:
        ind_bits.append('w:firstLineChars="%d"' % indent_chars)
    if ind_bits:
        bits.append("<w:ind %s/>" % " ".join(ind_bits))
    if align:
        bits.append('<w:jc w:val="%s"/>' % align)
    if outline is not None:
        bits.append('<w:outlineLvl w:val="%d"/>' % outline)
    return "".join(bits)


def _builtin(doc, name, base="Normal"):
    """取得内置样式；文档中缺失时按同名创建（styleId 与内置一致）。"""
    from docx.enum.style import WD_STYLE_TYPE
    try:
        return doc.styles[name]
    except KeyError:
        return _ensure_style(doc, name, WD_STYLE_TYPE.PARAGRAPH, base)


def build_styles(doc):
    """按规范写入全部段落样式与字符样式。"""
    from docx.enum.style import WD_STYLE_TYPE
    S = WD_STYLE_TYPE.PARAGRAPH

    # ---------- 正文 ----------
    st = _ensure_style(doc, STYLE_BODY, S, "Normal")
    st.quick_style = True
    set_style_ppr(st, _ppr(align="both", line=360, first_line_chars=200, first_line=420))
    set_style_rpr(st, fonts_xml(FONT_SONG) + sz_xml(SIZE_5HAO))

    # ---------- 论文标题 ----------
    st = _ensure_style(doc, STYLE_TITLE, S, "Normal")
    st.quick_style = True
    set_style_ppr(st, _ppr(align="center", line=360, before=0, after=0))
    set_style_rpr(st, fonts_xml(FONT_HEI) + sz_xml(SIZE_3HAO))

    # ---------- 副标题 ----------
    st = _ensure_style(doc, STYLE_SUBTITLE, S, "Normal")
    st.quick_style = True
    set_style_ppr(st, _ppr(align="center", line=360, after=0))
    set_style_rpr(st, fonts_xml(FONT_SONG) + sz_xml(SIZE_2HAO))

    # ---------- 摘要（正文楷体五号，首行缩进 2 字符；标签“摘  要：”用黑体） ----------
    st = _ensure_style(doc, STYLE_ABSTRACT, S, "Normal")
    st.quick_style = True
    set_style_ppr(st, _ppr(align="both", line=360, first_line_chars=200, first_line=420))
    set_style_rpr(st, fonts_xml(FONT_KAI) + sz_xml(SIZE_5HAO))

    # ---------- 关键词（同摘要：内容楷体五号，标签黑体） ----------
    st = _ensure_style(doc, STYLE_KEYWORDS, S, "Normal")
    st.quick_style = True
    set_style_ppr(st, _ppr(align="both", line=360, first_line_chars=200, first_line=420))
    set_style_rpr(st, fonts_xml(FONT_KAI) + sz_xml(SIZE_5HAO))

    # ---------- 英文摘要 ----------
    st = _ensure_style(doc, STYLE_ABSTRACT_EN, S, "Normal")
    set_style_ppr(st, _ppr(align="both", line=360, first_line_chars=200, first_line=420))
    set_style_rpr(st, fonts_xml(FONT_SONG) + sz_xml(SIZE_5HAO))

    # ---------- 标题 1 / 2 / 3（内置样式 + 多级列表） ----------
    h1 = _builtin(doc, "Heading 1")
    h1.quick_style = True
    set_style_ppr(h1, _ppr(align="left", line=360, before=120, after=120,
                           first_line_chars=200, first_line=480,
                           keep_next=True, keep_lines=True, outline=0,
                           numid=NUM_ID, ilvl=0))
    set_style_rpr(h1, fonts_xml(FONT_HEI) + sz_xml(SIZE_2HAO))

    h2 = _builtin(doc, "Heading 2")
    h2.quick_style = True
    set_style_ppr(h2, _ppr(align="left", line=360, before=60, after=60,
                           first_line_chars=200, first_line=420,
                           keep_next=True, keep_lines=True, outline=1,
                           numid=NUM_ID, ilvl=1))
    set_style_rpr(h2, fonts_xml(FONT_SONG) + sz_xml(SIZE_5HAO) + "<w:b/>")

    h3 = _builtin(doc, "Heading 3")
    h3.quick_style = True
    set_style_ppr(h3, _ppr(align="left", line=360, before=60, after=60,
                           first_line_chars=200, first_line=420,
                           keep_next=True, keep_lines=True, outline=2,
                           numid=NUM_ID, ilvl=2))
    set_style_rpr(h3, fonts_xml(FONT_SONG) + sz_xml(SIZE_5HAO))

    # 标题 4 及以下（不使用编号，仅占位）
    for idx in (4, 5, 6, 7, 8, 9):
        h = _builtin(doc, "Heading %d" % idx)
        set_style_ppr(h, _ppr(align="left", line=360, before=60, after=60,
                               first_line_chars=200, first_line=420,
                               keep_next=True, keep_lines=True, outline=idx - 1))
        set_style_rpr(h, fonts_xml(FONT_SONG) + sz_xml(SIZE_5HAO))

    # ---------- 参考文献 ----------
    st = _ensure_style(doc, STYLE_REF_HEADING, S, "Normal")
    st.quick_style = True
    set_style_ppr(st, _ppr(align="left", line=360, before=0, after=0))
    set_style_rpr(st, fonts_xml(FONT_HEI) + sz_xml(SIZE_2HAO))

    st = _ensure_style(doc, STYLE_REF_CATEGORY, S, "Normal")
    set_style_ppr(st, _ppr(align="left", line=360, before=0, after=0))
    set_style_rpr(st, fonts_xml(FONT_KAI) + sz_xml(SIZE_5HAO) + "<w:b/>")

    st = _ensure_style(doc, STYLE_REF_ITEM, S, "Normal")
    set_style_ppr(st, _ppr(align="left", line=360, before=0, after=0))
    # 参考文献例外：西文也用楷体
    set_style_rpr(st, fonts_xml(FONT_KAI, ascii_font=FONT_KAI) + sz_xml(SIZE_5HAO))

    # ---------- 附录 ----------
    st = _ensure_style(doc, STYLE_APPENDIX, S, "Normal")
    st.quick_style = True
    set_style_ppr(st, _ppr(align="left", line=360, before=0, after=0))
    set_style_rpr(st, fonts_xml(FONT_HEI) + sz_xml(SIZE_2HAO))

    # 附录内的小节标题：与正文多级列表脱钩，不参与「一、二、三」编号
    st = _ensure_style(doc, STYLE_APPENDIX_HEADING, S, "Normal")
    st.quick_style = True
    set_style_ppr(st, _ppr(align="left", line=360, before=60, after=60,
                           first_line_chars=200, first_line=480,
                           keep_next=True, keep_lines=True, outline=1))
    set_style_rpr(st, fonts_xml(FONT_HEI) + sz_xml(SIZE_2HAO))

    # ---------- 题注（图题在图下、表题在表上，均居中） ----------
    cap = _builtin(doc, "Caption")
    cap.quick_style = True
    set_style_ppr(cap, _ppr(align="center", line=360, before=60, after=60)
                  + '<w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/>')
    set_style_rpr(cap, fonts_xml(FONT_HEI) + sz_xml(SIZE_5HAO))

    # ---------- 表格内文字 ----------
    st = _ensure_style(doc, STYLE_TABLE_TEXT, S, "Normal")
    set_style_ppr(st, _ppr(align="center", line=360, before=30, after=30))
    set_style_rpr(st, fonts_xml(FONT_SONG) + sz_xml(SIZE_5HAO))

    # ---------- 上标引用序号（字符样式） ----------
    st = _ensure_style(doc, STYLE_LABEL, WD_STYLE_TYPE.CHARACTER)
    st.quick_style = True
    set_style_rpr(st, fonts_xml(FONT_HEI) + sz_xml(SIZE_5HAO))

    st = _ensure_style(doc, STYLE_CITE, WD_STYLE_TYPE.CHARACTER)
    st.quick_style = True
    set_style_rpr(st, fonts_xml(FONT_SONG) + sz_xml(SIZE_5HAO) + "<w:vertAlign w:val=\"superscript\"/>")

    build_reference_numbering(doc)
    setup_theme(doc)
    return doc


# --------------------------------------------------------------------------
# 多级列表：一、 /（一）/ 1.
# --------------------------------------------------------------------------

def build_multilevel(doc):
    """建立与 标题1/2/3 绑定的多级列表。"""
    numbering = doc.part.numbering_part.element
    for num in numbering.findall(qn("w:num")):
        if num.get(qn("w:numId")) == str(NUM_ID):
            numbering.remove(num)
    for an in numbering.findall(qn("w:abstractNum")):
        if an.get(qn("w:abstractNumId")) == str(ABSTRACT_NUM_ID):
            numbering.remove(an)

    lvls = []
    specs = [
        # (ilvl, numFmt, lvlText, pStyle, firstLine)
        (0, "chineseCountingThousand", "%1、", "Heading1", 480),
        (1, "chineseCountingThousand", "（%2）", "Heading2", 420),
        (2, "decimal", "%3.", "Heading3", 420),
    ]
    for ilvl, fmt, text, pstyle, fl in specs:
        lvls.append(
            '<w:lvl w:ilvl="{i}">'
            '<w:start w:val="1"/>'
            '<w:numFmt w:val="{fmt}"/>'
            '<w:pStyle w:val="{ps}"/>'
            '<w:suff w:val="nothing"/>'
            '<w:lvlText w:val="{tx}"/>'
            '<w:lvlJc w:val="left"/>'
            '<w:pPr><w:ind w:left="0" w:firstLineChars="200" w:firstLine="{fl}"/></w:pPr>'
            '<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
            '</w:lvl>'.format(i=ilvl, fmt=fmt, ps=pstyle, tx=text, fl=fl)
        )
    for ilvl in range(3, 9):
        lvls.append(
            '<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
            '<w:suff w:val="nothing"/><w:lvlText w:val="%{n}."/><w:lvlJc w:val="left"/>'
            '<w:pPr><w:ind w:left="0" w:firstLineChars="200" w:firstLine="420"/></w:pPr>'
            '</w:lvl>'.format(i=ilvl, n=ilvl + 1)
        )

    abstract = parse_xml(
        '<w:abstractNum %s w:abstractNumId="%d">'
        '<w:multiLevelType w:val="multilevel"/>'
        '<w:name w:val="论文标题多级列表"/>'
        '%s</w:abstractNum>' % (nsdecls("w"), ABSTRACT_NUM_ID, "".join(lvls))
    )
    num = parse_xml(
        '<w:num %s w:numId="%d"><w:abstractNumId w:val="%d"/></w:num>'
        % (nsdecls("w"), NUM_ID, ABSTRACT_NUM_ID)
    )
    existing = numbering.findall(qn("w:abstractNum"))
    if existing:
        existing[-1].addnext(abstract)
    else:
        numbering.insert(0, abstract)
    numbering.append(num)
    return NUM_ID


# --------------------------------------------------------------------------
# 参考文献列表编号：[1] [2] ...
# --------------------------------------------------------------------------

def build_reference_numbering(doc):
    """参考文献条目使用独立的自动编号列表（lvlText = "[%1]"）。"""
    numbering = doc.part.numbering_part.element
    for num in numbering.findall(qn("w:num")):
        if num.get(qn("w:numId")) == str(REF_NUM_ID):
            numbering.remove(num)
    for an in numbering.findall(qn("w:abstractNum")):
        if an.get(qn("w:abstractNumId")) == str(REF_ABSTRACT_NUM_ID):
            numbering.remove(an)

    lvls = []
    for ilvl in range(9):
        if ilvl == 0:
            text, fmt = "[%1]", "decimal"
            rpr = ('<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s" w:eastAsia="%s" '
                   'w:hint="eastAsia"/></w:rPr>' % (FONT_KAI, FONT_KAI, FONT_KAI, FONT_KAI))
        else:
            text, fmt = "%%%d." % (ilvl + 1), "decimal"
            rpr = ""
        lvls.append(
            '<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/>'
            '<w:suff w:val="nothing"/><w:lvlText w:val="{tx}"/><w:lvlJc w:val="left"/>'
            '<w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr>{rpr}</w:lvl>'
            .format(i=ilvl, fmt=fmt, tx=text, rpr=rpr)
        )
    abstract = parse_xml(
        '<w:abstractNum %s w:abstractNumId="%d">'
        '<w:multiLevelType w:val="multilevel"/>'
        '<w:name w:val="参考文献编号"/>%s</w:abstractNum>'
        % (nsdecls("w"), REF_ABSTRACT_NUM_ID, "".join(lvls))
    )
    num = parse_xml(
        '<w:num %s w:numId="%d"><w:abstractNumId w:val="%d"/></w:num>'
        % (nsdecls("w"), REF_NUM_ID, REF_ABSTRACT_NUM_ID)
    )
    existing = numbering.findall(qn("w:abstractNum"))
    if existing:
        existing[-1].addnext(abstract)
    else:
        numbering.insert(0, abstract)
    numbering.append(num)
    return REF_NUM_ID


def _ref_numpr_xml():
    return parse_xml('<w:numPr %s><w:ilvl w:val="0"/><w:numId w:val="%d"/></w:numPr>'
                     % (nsdecls("w"), REF_NUM_ID))


# --------------------------------------------------------------------------
# 页脚页码
# --------------------------------------------------------------------------

def build_footer(doc, force=False):
    """页脚居中插入 PAGE 域页码；页脚已有内容时默认不覆盖。"""
    for section in doc.sections:
        section.footer.is_linked_to_previous = False
        footer = section.footer
        existing = footer.paragraphs[0] if footer.paragraphs else None
        if not force and existing is not None and existing.text.strip():
            continue
        p = existing if existing is not None else footer.add_paragraph()
        for r in list(p.runs):
            r._element.getparent().remove(r._element)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run_el in field_runs("PAGE  \\* MERGEFORMAT", "1"):
            p._p.append(run_el)
    return doc


# --------------------------------------------------------------------------
# 域（field）基础
# --------------------------------------------------------------------------

def field_runs(instr: str, result_text: str = "", rpr_xml: str = "",
               field_name: Optional[str] = None):
    """生成一组构成一个 Word 域的 run 元素。

    instr      域指令，例如 ' SEQ Figure \\* ARABIC '
    result     域的缓存结果（Word 更新域前显示的内容）
    rpr_xml    施加到各 run 的 rPr 片段
    field_name 可选：w:fldChar 的 w:fldLock / 名称
    """
    runs = []

    def _run(children):
        r = OxmlElement("w:r")
        if rpr_xml:
            r.append(parse_xml("<w:rPr %s>%s</w:rPr>" % (nsdecls("w"), rpr_xml)))
        for c in children:
            r.append(c)
        runs.append(r)
        return r

    begin = _new("w:fldChar", fldCharType="begin")
    if field_name:
        begin.set(qn("w:fldLock"), "true")
    _run([begin])

    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = instr
    _run([it])

    _run([_new("w:fldChar", fldCharType="separate")])

    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = result_text
    _run([t])

    _run([_new("w:fldChar", fldCharType="end")])
    return runs


def _max_bookmark_id(doc) -> int:
    ids = []
    for part in _all_xml_parts(doc):
        for el in part.iter(qn("w:bookmarkStart")):
            try:
                ids.append(int(el.get(qn("w:id"))))
            except (TypeError, ValueError):
                pass
    return max(ids) if ids else 0


def _all_xml_parts(doc):
    parts = [doc.element]
    try:
        parts.append(doc.part.numbering_part.element)
    except Exception:
        pass
    for section in doc.sections:
        for hf in (section.header, section.footer,
                   section.even_page_header, section.even_page_footer,
                   section.first_page_header, section.first_page_footer):
            try:
                parts.append(hf._element)
            except Exception:
                pass
    return parts


def _bookmark(doc, name: str, bid: int):
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bid))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bid))
    return start, end


# --------------------------------------------------------------------------
# 题注（图/表）：SEQ 域自动编号 + 书签，供交叉引用
# --------------------------------------------------------------------------

def _count_seq(doc, seq_name: str) -> int:
    n = 0
    for part in _all_xml_parts(doc):
        for it in part.iter(qn("w:instrText")):
            if it.text and ("SEQ %s" % seq_name) in it.text:
                n += 1
    return n


def make_caption_paragraph(doc, label: str, text: str, seq_name: str,
                           bookmark_base: Optional[str] = None):
    """构造题注段落：<label> {SEQ} <text>

    生成两个书签：
      <base>        覆盖「标签 + 序号」，交叉引用结果形如「图1」
      <base>_num    仅覆盖序号，交叉引用结果形如「1」（用于「见图1」写法）
    """
    seq_id = _count_seq(doc, seq_name) + 1
    base = bookmark_base or ("_Ref_%s_%d" % (seq_name, seq_id))
    num_name = base + "_num"
    bid = _max_bookmark_id(doc) + 1

    p = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    ppr.append(_new("w:pStyle", val=STYLE_CAPTION))
    ppr.append(_new("w:autoSpaceDE", val="0"))
    ppr.append(_new("w:autoSpaceDN", val="0"))
    p.append(ppr)

    s1, e1 = _bookmark(doc, base, bid)
    s2, e2 = _bookmark(doc, num_name, bid + 1)
    p.append(s1)

    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = label
    r.append(t)
    p.append(r)

    p.append(s2)
    for el in field_runs(" SEQ %s \\* ARABIC " % seq_name, str(seq_id)):
        p.append(el)
    p.append(e2)
    p.append(e1)

    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = (" " + text) if text else ""
    r.append(t)
    p.append(r)

    if label == "表":
        # 表题在表格上方：与表格「同页」，避免表题被单独留在页尾
        ppr.append(_new("w:keepNext"))
    return p, base, num_name


def insert_caption(doc, label: str, text: str, anchor: Optional[Paragraph] = None,
                   where: str = "after", seq_name: Optional[str] = None):
    """在 anchor 段落的前/后插入题注。

    label   '图' 或 '表'
    where   'after'（图题在图下方） / 'before'（表题在表上方）
    返回 (段落, 书签名, 序号书签名)
    """
    seq_name = seq_name or ("Figure" if label == "图" else "Table")
    p_el, base, num_name = make_caption_paragraph(doc, label, text, seq_name)
    if anchor is None:
        body = doc.element.body
        sect = body.find(qn("w:sectPr"))
        if sect is not None:
            sect.addprevious(p_el)
        else:
            body.append(p_el)
    elif where == "before":
        anchor._p.addprevious(p_el)
    else:
        anchor._p.addnext(p_el)
    return Paragraph(p_el, anchor._parent if anchor is not None else doc), base, num_name


# --------------------------------------------------------------------------
# 交叉引用
# --------------------------------------------------------------------------

def add_crossref(paragraph, bookmark: str, prefix: str = "", suffix: str = "",
                 rpr_xml: str = ""):
    """在段落末尾插入 REF 交叉引用域。

    prefix '见图' + 书签 '_Ref_Figure_1_num'  → 见图1
    prefix '见'   + 书签 '_Ref_Figure_1'      → 见图1
    """
    if prefix:
        r = OxmlElement("w:r")
        if rpr_xml:
            r.append(parse_xml("<w:rPr %s>%s</w:rPr>" % (nsdecls("w"), rpr_xml)))
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = prefix
        r.append(t)
        paragraph._p.append(r)
    for el in field_runs(" REF %s \\h " % bookmark, "0", rpr_xml=rpr_xml):
        paragraph._p.append(el)
    if suffix:
        r = OxmlElement("w:r")
        if rpr_xml:
            r.append(parse_xml("<w:rPr %s>%s</w:rPr>" % (nsdecls("w"), rpr_xml)))
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = suffix
        r.append(t)
        paragraph._p.append(r)
    return paragraph


# --------------------------------------------------------------------------
# 参考文献：条目 + 上标引用
# --------------------------------------------------------------------------

REF_SLOT_PREFIX = "_Ref_Ref"
CITE_RPR = ('<w:rStyle w:val="%s"/><w:rFonts w:ascii="%s" w:hAnsi="%s" '
            'w:cs="%s" w:eastAsia="%s" w:hint="eastAsia"/>'
            '<w:vertAlign w:val="superscript"/>'
            % (STYLE_CITE, FONT_EN, FONT_EN, FONT_EN, FONT_SONG))


def _reference_paragraphs(doc):
    return [p for p in doc.paragraphs if p.style.name == STYLE_REF_ITEM]


def add_reference(doc, text: str, anchor: Optional[Paragraph] = None):
    """追加一条参考文献。

    条目采用自动编号列表（显示为 [1]、[2]…），并在段落起始处放置空书签
    _Ref_RefN。正文引用使用 { REF _Ref_RefN \\r \\h } 取得段落编号，
    序号即随增删条目自动更新。

    返回 (段落, 书签名, 显示序号)
    """
    slot = len(_reference_paragraphs(doc)) + 1
    bookmark = "%s%d" % (REF_SLOT_PREFIX, slot)
    bid = _max_bookmark_id(doc) + 1

    p = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    ppr.append(_new("w:pStyle", val=STYLE_REF_ITEM))
    ppr.append(_ref_numpr_xml())
    p.append(ppr)

    s, e = _bookmark(doc, bookmark, bid)
    p.append(s)
    p.append(e)

    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    p.append(r)

    if anchor is None:
        body = doc.element.body
        sect = body.find(qn("w:sectPr"))
        if sect is not None:
            sect.addprevious(p)
        else:
            body.append(p)
    else:
        anchor._p.addnext(p)
    return Paragraph(p, doc), bookmark, slot


CN_NUM = "一二三四五六七八九十百"
HEADING1_RE = re.compile(r"^([%s]+)、\s*" % CN_NUM)
HEADING2_RE = re.compile(r"^（([%s]+)）\s*" % CN_NUM)
HEADING3_RE = re.compile(r"^(\d{1,2})[.、]\s*")


def detect_heading_level(text: str):
    """按中文/数字前缀判断标题层级；带「级标题」批注时优先采信。

    返回 (level, 去掉编号后的标题文字) 或 None。
    """
    annotated = bool(re.search(r"[一二三]级标题", text))
    for level, rx in ((1, HEADING1_RE), (2, HEADING2_RE), (3, HEADING3_RE)):
        m = rx.match(text)
        if not m:
            continue
        if annotated:
            return level, text[m.end():]
        if len(text) <= 30 and not re.search(r"[。；，]", text):
            return level, text[m.end():]
    return None


REF_PREFIX_RE = re.compile(r"^[\s\[［]*\d+\s*[\]］]\s*")


def convert_reference_list(doc):
    """把手写的「[1] xxx」参考文献条目转换为自动编号列表，并补齐空书签。

    转换后条目由列表自动编号，正文 { REF _Ref_RefN \\r \\h } 交叉引用会自动跟随。
    """
    changed = []
    items = [p for p in doc.paragraphs if p.style.name == STYLE_REF_ITEM]
    for i, p in enumerate(items, start=1):
        ppr = p._p.get_or_add_pPr()
        if ppr.find(qn("w:numPr")) is None:
            ps = ppr.find(qn("w:pStyle"))
            npr = _ref_numpr_xml()
            if ps is not None:
                ps.addnext(npr)
            else:
                ppr.insert(0, npr)
            changed.append(i)
        # 去掉手写编号前缀
        for r in list(p._p.findall(qn("w:r"))):
            t = r.find(qn("w:t"))
            if t is None or not t.text:
                continue
            m = REF_PREFIX_RE.match(t.text)
            if m:
                t.text = t.text[m.end():]
                t.set(qn("xml:space"), "preserve")
                break
        # 补齐空书签
        names = [b.get(qn("w:name")) for b in p._p.iter(qn("w:bookmarkStart"))]
        if not any(n and n.startswith(REF_SLOT_PREFIX) for n in names):
            bid = _max_bookmark_id(doc) + 1
            bs, be = _bookmark(doc, "%s%d" % (REF_SLOT_PREFIX, i), bid)
            ppr.addnext(bs)
            bs.addnext(be)
    return changed


def renumber_references(doc):
    """兼容接口：参考文献改为自动编号列表后无需手工重排，
    此处仅调用 convert_reference_list 补齐编号与书签。"""
    return convert_reference_list(doc)


def add_citation(paragraph, slots: Sequence[int], prefix: str = ""):
    """在段落末尾追加上标引用，如 [1][2]。

    slots 为参考文献条目序号（与 add_reference 返回的 slot 一致）。
    """
    if prefix:
        r = OxmlElement("w:r")
        r.append(parse_xml("<w:rPr %s>%s</w:rPr>" % (nsdecls("w"), CITE_RPR)))
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = prefix
        r.append(t)
        paragraph._p.append(r)
    for slot in slots:
        bookmark = "%s%d" % (REF_SLOT_PREFIX, slot)
        # \\r = 取所引用段落的编号（即参考文献条目序号），保留上标格式
        for el in field_runs(" REF %s \\r \\h " % bookmark, "[%d]" % slot, rpr_xml=CITE_RPR):
            paragraph._p.append(el)
    return paragraph


# --------------------------------------------------------------------------
# 图片插入：等比缩放到版心内
# --------------------------------------------------------------------------

def _image_native_cm(path):
    """读取图片原始尺寸（cm）。"""
    from docx.image.image import Image as DocxImage
    img = DocxImage.from_file(path)
    w = img.px_width / float(img.horz_dpi or 96) * 2.54
    h = img.px_height / float(img.vert_dpi or 96) * 2.54
    return w, h


def compute_image_size(path, max_width_cm=MAX_IMAGE_WIDTH_CM,
                       max_height_cm=MAX_IMAGE_HEIGHT_CM):
    """等比缩放：不超过版心宽度，也不超过版心高度的三分之一。"""
    w, h = _image_native_cm(path)
    if w <= 0 or h <= 0:
        return max_width_cm, max_height_cm
    scale = min(1.0, max_width_cm / w, max_height_cm / h)
    return w * scale, h * scale


def insert_picture(doc, path, anchor=None, where="after",
                   width_cm=None, height_cm=None,
                   max_width_cm=MAX_IMAGE_WIDTH_CM,
                   max_height_cm=MAX_IMAGE_HEIGHT_CM,
                   keep_with_caption=True):
    """插入图片并自动等比缩放，返回 (段落, 宽cm, 高cm)。

    图片段落居中，且与随后的题注「同页」（keepNext），避免图与图题被分到两页。
    """
    if width_cm is None or height_cm is None:
        w, h = compute_image_size(path, max_width_cm, max_height_cm)
        if width_cm is not None:
            w = width_cm
            h = h * (width_cm / max(w, 1e-6))
        if height_cm is not None:
            h = height_cm
            w = w * (height_cm / max(h, 1e-6))
    else:
        w, h = compute_image_size(path, max_width_cm, max_height_cm)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(path, width=Cm(w), height=Cm(h))
    ppr = p._p.get_or_add_pPr()
    ppr.append(_new("w:keepNext"))
    ppr.append(_new("w:spacing", before="60", after="60", line="240", lineRule="auto"))

    if anchor is not None:
        if where == "before":
            anchor._p.addprevious(p._p)
        else:
            anchor._p.addnext(p._p)
    return p, w, h


def set_keep_next(paragraph):
    """段落「与下段同页」。"""
    ppr = paragraph._p.get_or_add_pPr()
    if ppr.find(qn("w:keepNext")) is None:
        ppr.insert(0, _new("w:keepNext"))
    return paragraph


# --------------------------------------------------------------------------
# 三线表
# --------------------------------------------------------------------------

def doc_styles_of(table):
    """取得表格所属文档的 styles 集合。"""
    parent = table._parent
    while parent is not None and not hasattr(parent, "styles"):
        parent = getattr(parent, "_parent", None)
    return parent.styles


def _set_tbl_borders(tbl, top_sz=12, bottom_sz=12, header_sz=6):
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    borders = parse_xml(
        '<w:tblBorders %s>'
        '<w:top w:val="single" w:sz="%d" w:space="0" w:color="000000"/>'
        '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="single" w:sz="%d" w:space="0" w:color="000000"/>'
        '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '</w:tblBorders>' % (nsdecls("w"), top_sz, bottom_sz)
    )
    old = tblPr.find(qn("w:tblBorders"))
    if old is not None:
        tblPr.remove(old)
    # tblBorders 在 tblPr 中应排在 tblLayout/tblCellMar 之前，简化处理：插到 tblW 之后
    tblPr.insert(0, borders)


def _set_row_cant_split(row):
    """禁止该行的内容跨页断开。"""
    trPr = row._tr.get_or_add_trPr()
    if trPr.find(qn("w:cantSplit")) is None:
        trPr.insert(0, _new("w:cantSplit"))


def _set_row_repeat_header(row):
    """该行为表头行，跨页时自动重复。"""
    trPr = row._tr.get_or_add_trPr()
    if trPr.find(qn("w:tblHeader")) is None:
        trPr.append(_new("w:tblHeader", val="true"))


def estimate_table_height_cm(table, font_pt=10.5):
    """估算表格高度（cm）：每行约 字号*1.5 倍行距 + 上下内边距。"""
    line = font_pt * 1.5 / 72.0 * 2.54
    per_row = line + 0.20
    return len(table.rows) * per_row


def fit_table_to_page(table, doc=None, max_fraction=MAX_TABLE_PAGE_FRACTION):
    """表格过高时自动缩排（五号→小五、行距 1.5 倍→单倍），避免占满大半页。

    返回是否触发了缩排。
    """
    limit = TEXT_HEIGHT_CM * max_fraction
    if estimate_table_height_cm(table, 10.5) <= limit:
        return False
    shrunk = False
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(9)
                pf = p.paragraph_format
                pf.line_spacing = 1.0
                pf.space_before = Pt(0)
                pf.space_after = Pt(0)
                shrunk = True
    return shrunk


def _set_cell_bottom_border(cell, sz=6):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        '<w:tcBorders %s><w:bottom w:val="single" w:sz="%d" w:space="0" w:color="000000"/>'
        '</w:tcBorders>' % (nsdecls("w"), sz)
    )
    old = tcPr.find(qn("w:tcBorders"))
    if old is not None:
        tcPr.remove(old)
    tcPr.append(borders)


def format_three_line_table(table, header_rows: int = 1, align: str = "center"):
    """把任意表格改造成标准三线表：顶线（1.5pt）、表头下线（0.75pt）、底线（1.5pt），无竖线。"""
    _set_tbl_borders(table._tbl)
    rows = table.rows
    for i, row in enumerate(rows):
        _set_row_cant_split(row)          # 单行内容不跨页断开
        if i < header_rows:
            _set_row_repeat_header(row)   # 表头行跨页重复
            for cell in row.cells:
                _set_cell_bottom_border(cell, sz=6)
        for cell in row.cells:
            for p in cell.paragraphs:
                try:
                    p.style = doc_styles_of(table)[STYLE_TABLE_TEXT]
                except Exception:
                    pass
    fit_table_to_page(table)
    return table


def add_three_line_table(doc, data: Sequence[Sequence[str]], header_rows: int = 1,
                         caption: Optional[str] = None, anchor: Optional[Paragraph] = None):
    """插入标准三线表；若给出 caption，自动在表格上方插入「表N 名称」题注。"""
    rows = len(data)
    cols = max(len(r) for r in data)
    table = doc.add_table(rows=rows, cols=cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    for i, row in enumerate(data):
        for j, val in enumerate(row):
            cell = table.cell(i, j)
            cell.text = ""
            run = cell.paragraphs[0].add_run(val)
            run.font.size = Pt(10.5)
            cell.paragraphs[0].style = doc.styles[STYLE_TABLE_TEXT]
    format_three_line_table(table, header_rows=header_rows)

    if anchor is not None:
        anchor._p.addprevious(table._tbl)
    if caption:
        cap_p, _base, _num = make_caption_paragraph(doc, "表", caption, "Table")
        table._tbl.addprevious(cap_p)
    return table


# --------------------------------------------------------------------------
# 标签（“摘  要：”“关键词：”）单独套用黑体字符样式
# --------------------------------------------------------------------------

LABEL_PATTERNS = (
    re.compile(r"^摘\s*要\s*[:：]"),
    re.compile(r"^关\s*键\s*词\s*[:：]"),
    re.compile(r"^Abstract\s*[:：]"),
)


def _set_run_style(run, style_name: str):
    rpr = run.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        run.insert(0, rpr)
    rs = rpr.find(qn("w:rStyle"))
    if rs is None:
        rs = OxmlElement("w:rStyle")
        rpr.insert(0, rs)
    rs.set(qn("w:val"), style_name)


def mark_leading_label(paragraph, style_name: str = STYLE_LABEL):
    """把段落开头的标签（摘  要：/关键词：/Abstract:）切成独立 run 并套标签样式。

    返回是否找到标签。
    """
    runs = list(paragraph._p.findall(qn("w:r")))
    if not runs:
        return False
    texts = []
    for r in runs:
        t = r.find(qn("w:t"))
        texts.append(t.text if t is not None and t.text else "")
    full = "".join(texts)
    end = None
    for rx in LABEL_PATTERNS:
        m = rx.match(full)
        if m:
            end = m.end()
            break
    if end is None:
        return False
    pos = 0
    for r, txt in zip(runs, texts):
        s0, e0 = pos, pos + len(txt)
        pos = e0
        if s0 >= end:
            break
        if e0 <= end:
            _set_run_style(r, style_name)
        else:
            t = r.find(qn("w:t"))
            t.text = txt[end - s0:]
            t.set(qn("xml:space"), "preserve")
            new_r = copy.deepcopy(r)
            nt = new_r.find(qn("w:t"))
            nt.text = txt[:end - s0]
            nt.set(qn("xml:space"), "preserve")
            _set_run_style(new_r, style_name)
            r.addprevious(new_r)
            break
    return True


# --------------------------------------------------------------------------
# 中文双引号：强制使用中文字体
# --------------------------------------------------------------------------

QUOTE_CHARS = set(CN_QUOTES) | set(CN_QUOTE_SINGLE)

# 摘要 / 关键词 / 参考文献中，双引号使用楷体
KAI_ZONE_STYLES = {
    STYLE_ABSTRACT, STYLE_ABSTRACT_EN, STYLE_KEYWORDS,
    STYLE_REF_ITEM, STYLE_REF_CATEGORY,
}


def _ensure_rfonts(rpr):
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rstyle = rpr.find(qn("w:rStyle"))
        if rstyle is not None:
            rstyle.addnext(rf)
        else:
            rpr.insert(0, rf)
    return rf


def _paragraph_style_id(p_el):
    ppr = p_el.find(qn("w:pPr"))
    if ppr is None:
        return None
    ps = ppr.find(qn("w:pStyle"))
    return ps.get(qn("w:val")) if ps is not None else None


def _split_run_for_quotes(run, target_font: str):
    ts = run.findall(qn("w:t"))
    if len(ts) != 1:
        return False
    text = ts[0].text or ""
    if not any(c in text for c in QUOTE_CHARS):
        return False
    parent = run.getparent()
    idx = parent.index(run)
    rpr = run.find(qn("w:rPr"))
    segments = [s for s in re.split(r"([“”‘’])", text) if s != ""]
    new_elems = []
    for seg in segments:
        nr = OxmlElement("w:r")
        if rpr is not None:
            nr.append(copy.deepcopy(rpr))
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = seg
        nr.append(t)
        if seg in QUOTE_CHARS:
            nrpr = nr.find(qn("w:rPr"))
            if nrpr is None:
                nrpr = OxmlElement("w:rPr")
                nr.insert(0, nrpr)
            rf = _ensure_rfonts(nrpr)
            rf.set(qn("w:ascii"), target_font)
            rf.set(qn("w:hAnsi"), target_font)
            rf.set(qn("w:cs"), target_font)
            rf.set(qn("w:eastAsia"), target_font)
            rf.set(qn("w:hint"), "eastAsia")
        new_elems.append(nr)
    for i, e in enumerate(new_elems):
        parent.insert(idx + i, e)
    parent.remove(run)
    return True


def fix_quotes(doc, abstract_font: str = FONT_KAI, ref_font: str = FONT_KAI,
               body_font: str = FONT_SONG, verbose: bool = False):
    """把文中所有中文双引号（含单引号）单独切分为独立 run，
    并强制使用中文字体：
        摘要 / 关键词 / 参考文献区 → 楷体
        正文及标题                → 该 run 原有中文字体（默认宋体）
    """
    changed = 0
    roots = [doc.element]
    for section in doc.sections:
        for hf in (section.header, section.footer,
                   section.even_page_header, section.even_page_footer,
                   section.first_page_header, section.first_page_footer):
            try:
                roots.append(hf._element)
            except Exception:
                pass

    for root in roots:
        for p_el in list(root.iter(qn("w:p"))):
            sid = _paragraph_style_id(p_el)
            if sid in KAI_ZONE_STYLES:
                zone_font = abstract_font if sid in (
                    STYLE_ABSTRACT, STYLE_ABSTRACT_EN, STYLE_KEYWORDS) else ref_font
            else:
                zone_font = None
            for run in list(p_el.iter(qn("w:r"))):
                if run.find(qn("w:fldChar")) is not None or \
                   run.find(qn("w:instrText")) is not None:
                    continue
                rpr = run.find(qn("w:rPr"))
                font = zone_font
                if font is None:
                    rf = rpr.find(qn("w:rFonts")) if rpr is not None else None
                    font = (rf.get(qn("w:eastAsia")) if rf is not None else None) or body_font
                if _split_run_for_quotes(run, font):
                    changed += 1
    if verbose:
        print("引号修正段落/文本段数：%d" % changed)
    return changed


# --------------------------------------------------------------------------
# 分节分页
# --------------------------------------------------------------------------

def insert_section_break_before(doc, paragraph, break_type: str = "nextPage"):
    """在指定段落之前插入分节符，使该段落另起一页。

    Word 的分节符保存在「上一节最后一段」的 pPr/sectPr 中，
    因此这里把当前节的 sectPr 复制一份并挂到目标段落的前一段。
    """
    body = doc.element.body
    body_sect = body.find(qn("w:sectPr"))
    if body_sect is None:
        return None
    new_sect = copy.deepcopy(body_sect)
    t = new_sect.find(qn("w:type"))
    if t is None:
        t = OxmlElement("w:type")
        new_sect.insert(0, t)
    t.set(qn("w:val"), break_type)

    p_el = paragraph._p
    prev = p_el.getprevious()
    if prev is None or prev.tag != qn("w:p"):
        holder = OxmlElement("w:p")
        ppr = OxmlElement("w:pPr")
        holder.append(ppr)
        p_el.addprevious(holder)
        prev = holder
    ppr = prev.find(qn("w:pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr")
        prev.insert(0, ppr)
    old = ppr.find(qn("w:sectPr"))
    if old is not None:
        ppr.remove(old)
    ppr.append(new_sect)
    return new_sect


def preceding_section_break_paragraph(paragraph):
    """返回 paragraph 之前最近一个「携带分节符」的段落（没有则返回 None）。

    插入内容到「参考文献/附录」标题之前时，必须插到这个段落的**前面**，
    否则新内容会落到分节符之后的下一节里，与标题同页。
    """
    prev = paragraph._p.getprevious()
    while prev is not None:
        if prev.tag == qn("w:p"):
            ppr = prev.find(qn("w:pPr"))
            if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
                return Paragraph(prev, paragraph._parent)
        prev = prev.getprevious()
    return None


def has_section_break_before(paragraph) -> bool:
    prev = paragraph._p.getprevious()
    if prev is None or prev.tag != qn("w:p"):
        return False
    ppr = prev.find(qn("w:pPr"))
    return ppr is not None and ppr.find(qn("w:sectPr")) is not None


# --------------------------------------------------------------------------
# 模板生成
# --------------------------------------------------------------------------

def _add_paragraph(doc, text: str = "", style: Optional[str] = None):
    p = doc.add_paragraph()
    if style:
        p.style = doc.styles[style]
    if text:
        p.add_run(text)
        if style in (STYLE_ABSTRACT, STYLE_KEYWORDS):
            mark_leading_label(p)
    return p


def build_template(path: str, with_examples: bool = True):
    """生成符合规范的论文模板 .docx。"""
    doc = Document()
    setup_page(doc)
    setup_doc_defaults(doc)
    build_multilevel(doc)
    build_styles(doc)
    build_footer(doc)

    # 清除默认空段落
    body = doc.element.body
    for p in body.findall(qn("w:p")):
        body.remove(p)

    if with_examples:
        _add_paragraph(doc, "【论文标题】", STYLE_TITLE)
        _add_paragraph(doc, "【副标题】", STYLE_SUBTITLE)
        _add_paragraph(doc, "")
        _add_paragraph(doc, "摘  要：【摘要正文，楷体五号，首行缩进 2 字符】", STYLE_ABSTRACT)
        _add_paragraph(doc, "关键词：【关键词一】  【关键词二】  【关键词三】（楷体五号，之间空 2 字符）", STYLE_KEYWORDS)
        _add_paragraph(doc, "")
        _add_paragraph(doc, "")

        _add_paragraph(doc, "【一级标题】", "Heading 1")
        _add_paragraph(doc, "【正文段落，宋体五号，1.5 倍行距，首行缩进 2 字符。】", STYLE_BODY)
        _add_paragraph(doc, "【二级标题】", "Heading 2")
        _add_paragraph(doc, "【正文段落】", STYLE_BODY)
        _add_paragraph(doc, "【三级标题】", "Heading 3")
        _add_paragraph(doc, "【正文段落】", STYLE_BODY)

        _add_paragraph(doc, "【插入图片后，在此段落之后插入图题】", STYLE_BODY)
        cap, base, num = insert_caption(doc, "图", "【图名】", anchor=doc.paragraphs[-1], where="after")

        # Word 原生图表示例（双击图表可在 Excel 中编辑数据）
        _add_paragraph(doc, "【Word 原生图表示例：柱状图/饼图/折线图一律用「插入→图表」，"
                            "不要用 Python 画成图片】", STYLE_BODY)
        try:
            from chartlib import insert_chart
            _p = doc.paragraphs[-1]
            insert_chart(doc, {
                "type": "column",
                "categories": ["一季度", "二季度", "三季度", "四季度"],
                "series": [{"name": "系列一", "values": [120, 150, 180, 210]}],
                "title": "【图表标题】",
                "data_labels": True,
            }, anchor=_p, where="after",
                width_cm=DEFAULT_CHART_WIDTH_CM, height_cm=DEFAULT_CHART_HEIGHT_CM)
            insert_caption(doc, "图", "【图名】", anchor=doc.paragraphs[-1], where="after")
        except Exception as exc:      # pragma: no cover
            print("（跳过图表示例：%s）" % exc)

        _add_paragraph(doc, "【插入三线表】", STYLE_BODY)

        # 三线表示例
        tbl = doc.add_table(rows=3, cols=3)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        data = [["表头 1", "表头 2", "表头 3"],
                ["数据 1", "数据 2", "数据 3"],
                ["数据 4", "数据 5", "数据 6"]]
        for i, row in enumerate(data):
            for j, val in enumerate(row):
                cell = tbl.cell(i, j)
                cell.text = ""
                cell.paragraphs[0].style = doc.styles[STYLE_TABLE_TEXT]
                cell.paragraphs[0].add_run(val)
        format_three_line_table(tbl, header_rows=1)
        insert_caption(doc, "表", "【表名】", anchor=Paragraph(tbl._tbl, doc), where="before")

        # 参考文献（分节分页）
        ref_head = _add_paragraph(doc, "参考文献", STYLE_REF_HEADING)
        insert_section_break_before(doc, ref_head)
        _add_paragraph(doc, "（1）学术著作、论文集、学位论文、报告", STYLE_REF_CATEGORY)
        add_reference(doc, "【主要责任者】．【文献题名】[M]．【出版地】：【出版者】，【出版年】．")
        _add_paragraph(doc, "（2）期刊文章", STYLE_REF_CATEGORY)
        add_reference(doc, "【主要责任者】．【文献题名】[J]．【刊名】，【出版年】，【卷号】（【期号】）．")

        # 附录（分节分页）
        app_head = _add_paragraph(doc, "附录", STYLE_APPENDIX)
        insert_section_break_before(doc, app_head)
        _add_paragraph(doc, "【附录正文】", STYLE_BODY)

    doc.save(path)
    return path


# --------------------------------------------------------------------------
# 对既有文档套用规范
# --------------------------------------------------------------------------

RE_HEADING_RE = re.compile(r"^参\s*考\s*文\s*献\s*[:：]?\s*$")
APPENDIX_RE = re.compile(r"^附\s*录\s*[:：]?\s*$")
APPENDIX_SUB_RE = re.compile(r"^附\s*录\s*[A-ZＡ-Ｚ\d一二三四五六七八九十]")
ANNOTATION_RE = re.compile(r"[（(][^）)]*[）)]\s*$")


def _strip_annotation(text: str) -> str:
    """去掉行尾的「（……）」批注，便于识别标题行。"""
    prev = None
    t = text
    while prev != t:
        prev = t
        t = ANNOTATION_RE.sub("", t).strip()
    return t
ABSTRACT_RE = re.compile(r"^摘\s*要\s*[:：]")
KEYWORDS_RE = re.compile(r"^关\s*键\s*词\s*[:：]?")
REF_ITEM_RE = re.compile(r"^\[\s*\d+\s*\]")
REF_CATEGORY_RE = re.compile(r"^[（(]\s*\d+\s*[）)]")


STRIP_PROPS = ("w:rFonts", "w:sz", "w:szCs", "w:color", "w:highlight",
               "w:shd", "w:position", "w:spacing", "w:kern", "w:lang")


def strip_direct_formatting(doc, keep_emphasis=True, keep_superscript=True):
    """清除 run 上的直接字符格式（字体、字号、颜色等），使样式成为唯一来源。

    保留：加粗/斜体/下划线（强调）、rStyle（引用序号等字符样式）、上下标。
    """
    roots = [doc.element]
    for section in doc.sections:
        for hf in (section.header, section.footer,
                   section.even_page_header, section.even_page_footer,
                   section.first_page_header, section.first_page_footer):
            try:
                roots.append(hf._element)
            except Exception:
                pass
    n = 0
    for root in roots:
        for run in root.iter(qn("w:r")):
            rpr = run.find(qn("w:rPr"))
            if rpr is None:
                continue
            for tag in STRIP_PROPS:
                el = rpr.find(qn(tag))
                if el is not None:
                    rpr.remove(el)
                    n += 1
            if not keep_emphasis:
                for tag in ("w:b", "w:bCs", "w:i", "w:iCs", "w:u"):
                    el = rpr.find(qn(tag))
                    if el is not None:
                        rpr.remove(el)
            if not keep_superscript:
                el = rpr.find(qn("w:vertAlign"))
                if el is not None:
                    rpr.remove(el)
            if len(rpr) == 0:
                run.remove(rpr)
    return n


def ensure_format_environment(doc, footer=False):
    """为既有文档补齐/覆盖规范所需的环境：页面、文档默认值、多级列表、样式。"""
    setup_page(doc)
    setup_doc_defaults(doc)
    build_multilevel(doc)
    build_styles(doc)
    if footer:
        build_footer(doc)


def apply_format(doc, fix_quote: bool = True, three_line: bool = True,
                 section_breaks: bool = True, verbose: bool = False,
                 strip_direct: bool = True, keep_emphasis: bool = True,
                 auto_heading: bool = False):
    """把既有文档的段落映射到规范样式，并做表格/引号/分节处理。"""
    ensure_format_environment(doc)
    report = []
    if strip_direct:
        n = strip_direct_formatting(doc, keep_emphasis=keep_emphasis)
        report.append("已清除直接字符格式 %d 处（字体/字号/颜色改由样式控制）" % n)
    in_ref_block = False
    title_used = False
    for p in list(doc.paragraphs):
        text = p.text.strip()
        sid = _paragraph_style_id(p._p)
        if not text:
            continue
        if sid in ("Heading1", "Heading2", "Heading3", "Title", "Caption",
                   STYLE_BODY, STYLE_ABSTRACT, STYLE_KEYWORDS,
                   STYLE_REF_ITEM, STYLE_REF_CATEGORY, STYLE_REF_HEADING,
                   STYLE_APPENDIX, STYLE_APPENDIX_HEADING, STYLE_SUBTITLE,
                   STYLE_TABLE_TEXT):
            if sid == STYLE_REF_HEADING:
                in_ref_block = True
            continue

        if RE_HEADING_RE.match(_strip_annotation(text)):
            p.style = doc.styles[STYLE_REF_HEADING]
            in_ref_block = True
            if section_breaks and not has_section_break_before(p):
                insert_section_break_before(doc, p)
                report.append("参考文献：已插入分节符（另起一页）")
            continue
        if APPENDIX_RE.match(_strip_annotation(text)):
            p.style = doc.styles[STYLE_APPENDIX]
            in_ref_block = False
            if section_breaks and not has_section_break_before(p):
                insert_section_break_before(doc, p)
                report.append("附录：已插入分节符（另起一页）")
            continue
        if APPENDIX_SUB_RE.match(text):
            # 「附录 A xxx」是小节标题：挂附录专用的标题样式，不分节
            p.style = doc.styles[STYLE_APPENDIX_HEADING]
            continue
        if in_ref_block:
            if REF_CATEGORY_RE.match(text):
                p.style = doc.styles[STYLE_REF_CATEGORY]
            elif REF_ITEM_RE.match(text):
                p.style = doc.styles[STYLE_REF_ITEM]
            else:
                p.style = doc.styles[STYLE_REF_ITEM]
            continue
        if ABSTRACT_RE.match(text):
            p.style = doc.styles[STYLE_ABSTRACT]
            mark_leading_label(p)
            continue
        if KEYWORDS_RE.match(text):
            p.style = doc.styles[STYLE_KEYWORDS]
            mark_leading_label(p)
            continue
        if auto_heading:
            hit = detect_heading_level(text)
            if hit:
                level, rest = hit
                p.style = doc.styles["Heading %d" % level]
                # 删除手写编号，改由多级列表自动编号
                for r in list(p._p.findall(qn("w:r"))):
                    t = r.find(qn("w:t"))
                    if t is None or not t.text:
                        continue
                    m = (HEADING1_RE if level == 1 else
                         HEADING2_RE if level == 2 else HEADING3_RE).match(t.text)
                    if m:
                        t.text = t.text[m.end():]
                        t.set(qn("xml:space"), "preserve")
                        break
                report.append("标题识别：%s → 标题 %d" % (text[:18], level))
                continue
        if not title_used:
            p.style = doc.styles[STYLE_TITLE]
            title_used = True
            continue
        p.style = doc.styles[STYLE_BODY]

    refs = convert_reference_list(doc)
    if refs:
        report.append("参考文献已改为自动编号列表（[1] [2] …），共 %d 条" % len(refs))

    if three_line:
        for t in doc.tables:
            format_three_line_table(t, header_rows=1)

    if fix_quote:
        n = fix_quotes(doc, verbose=verbose)
        report.append("中文双引号已统一为中文字体（修正 %d 处）" % n)

    return report


# --------------------------------------------------------------------------
# 合规自检
# --------------------------------------------------------------------------

def _resolve_style_attr(doc, style, tag: str, attr: str):
    """沿 basedOn 链解析样式的 rPr 属性，最后回落到 docDefaults。

    Word 保存时会删除与继承值重复的属性，所以必须解析"有效值"而不是字面值。
    """
    seen = set()
    st = style
    while st is not None and st.style_id not in seen:
        seen.add(st.style_id)
        rpr = st.element.find(qn("w:rPr"))
        if rpr is not None:
            el = rpr.find(qn(tag))
            if el is not None:
                v = el.get(qn(attr))
                if v is not None:
                    return v
        base = st.element.find(qn("w:basedOn"))
        if base is None:
            break
        bid = base.get(qn("w:val"))
        st = next((x for x in doc.styles if x.style_id == bid), None)

    dd = doc.styles.element.find(qn("w:docDefaults"))
    if dd is not None:
        rpr_default = dd.find(qn("w:rPrDefault"))
        rpr = rpr_default.find(qn("w:rPr")) if rpr_default is not None else None
        if rpr is not None:
            el = rpr.find(qn(tag))
            if el is not None:
                v = el.get(qn(attr))
                if v is not None:
                    return v
    return None


def verify_format(doc) -> List[str]:
    """返回问题清单；空列表表示完全合规。"""
    issues = []

    # 文档默认值：西文 Times New Roman、中文宋体
    dd = doc.styles.element.find(qn("w:docDefaults"))
    rpr_default = dd.find(qn("w:rPrDefault")) if dd is not None else None
    rpr = rpr_default.find(qn("w:rPr")) if rpr_default is not None else None
    rf = rpr.find(qn("w:rFonts")) if rpr is not None else None
    if rf is None or rf.get(qn("w:ascii")) != FONT_EN:
        issues.append("文档默认西文字体应为 %s" % FONT_EN)

    # 页面
    for i, s in enumerate(doc.sections):
        if abs(s.page_width.cm - 21.0) > 0.1 or abs(s.page_height.cm - 29.7) > 0.1:
            issues.append("第 %d 节纸张不是 A4" % (i + 1))

    # 样式字体
    expected = {
        STYLE_BODY: (FONT_SONG, FONT_EN, SIZE_5HAO),
        STYLE_TITLE: (FONT_HEI, FONT_EN, SIZE_3HAO),
        STYLE_ABSTRACT: (FONT_KAI, FONT_EN, SIZE_5HAO),
        STYLE_KEYWORDS: (FONT_KAI, FONT_EN, SIZE_5HAO),
        STYLE_LABEL: (FONT_HEI, FONT_EN, SIZE_5HAO),
        STYLE_REF_HEADING: (FONT_HEI, FONT_EN, SIZE_2HAO),
        STYLE_REF_ITEM: (FONT_KAI, FONT_KAI, SIZE_5HAO),
        "Heading 1": (FONT_HEI, FONT_EN, SIZE_2HAO),
        "Heading 2": (FONT_SONG, FONT_EN, SIZE_5HAO),
        "Heading 3": (FONT_SONG, FONT_EN, SIZE_5HAO),
        STYLE_CAPTION: (FONT_HEI, FONT_EN, SIZE_5HAO),
    }
    for name, (ea, asc, size) in expected.items():
        try:
            st = doc.styles[name]
        except KeyError:
            issues.append("缺少样式：%s" % name)
            continue
        got_ea = _resolve_style_attr(doc, st, "w:rFonts", "w:eastAsia")
        got_asc = _resolve_style_attr(doc, st, "w:rFonts", "w:ascii")
        got_sz = _resolve_style_attr(doc, st, "w:sz", "w:val")
        if got_ea != ea:
            issues.append("样式 %s 中文字体应为 %s，实为 %s" % (name, ea, got_ea))
        if got_asc != asc:
            issues.append("样式 %s 西文字体应为 %s，实为 %s" % (name, asc, got_asc))
        if got_sz != str(size):
            issues.append("样式 %s 字号不符（应为 %s 磅值，实为 %s）"
                          % (name, size / 2, got_sz))

    # 正文 run：西文必须 Times New Roman；双引号必须中文字体
    for pi, p in enumerate(doc.paragraphs):
        sid = _paragraph_style_id(p._p) or ""
        for run in p._p.iter(qn("w:r")):
            t = run.find(qn("w:t"))
            if t is None or not t.text:
                continue
            rpr = run.find(qn("w:rPr"))
            rf = rpr.find(qn("w:rFonts")) if rpr is not None else None
            if rf is None:
                continue
            only_quote = all(ch in QUOTE_CHARS for ch in t.text)
            if (not only_quote and sid != STYLE_REF_ITEM
                    and rf.get(qn("w:ascii")) not in (None, FONT_EN)):
                issues.append("第 %d 段存在非 Times New Roman 西文设置：%r"
                              % (pi + 1, t.text[:20]))
            if any(c in t.text for c in QUOTE_CHARS):
                if rf.get(qn("w:hint")) != "eastAsia":
                    issues.append("双引号未使用中文字体：%r" % t.text[:20])
                want = FONT_KAI if sid in KAI_ZONE_STYLES else None
                got = rf.get(qn("w:eastAsia"))
                if want and got != want:
                    issues.append("第 %d 段双引号应使用 %s，实为 %s"
                                  % (pi + 1, want, got))
                elif not want and got not in (FONT_SONG, FONT_KAI, FONT_HEI):
                    issues.append("双引号中文字体异常：%r" % t.text[:20])

    # 交叉引用（REF 域）指向的书签必须存在
    bookmarks = {b.get(qn("w:name")) for b in doc.element.iter(qn("w:bookmarkStart"))}
    for it in doc.element.iter(qn("w:instrText")):
        txt = it.text or ""
        m = re.search(r"REF\s+(\S+)", txt)
        if m and m.group(1) not in bookmarks:
            issues.append("交叉引用指向不存在的书签：%s" % m.group(1))

    # 图片 / 图表尺寸：不得超过版心宽，且高度不超过版心的三分之一
    EMU_CM = 360000.0
    for d in doc.element.iter(qn("w:drawing")):
        ext = d.find(".//" + qn("wp:extent"))
        if ext is None:
            continue
        cx = int(ext.get("cx") or 0) / EMU_CM
        cy = int(ext.get("cy") or 0) / EMU_CM
        if cx > MAX_IMAGE_WIDTH_CM + 0.05:
            issues.append("图片/图表宽 %.1f cm 超过上限 %.1f cm"
                          % (cx, MAX_IMAGE_WIDTH_CM))
        if cy > MAX_IMAGE_HEIGHT_CM + 0.05:
            issues.append("图片/图表高 %.1f cm 超过上限 %.1f cm（版心三分之一）"
                          % (cy, MAX_IMAGE_HEIGHT_CM))

    # 表格跨页处理：表头行重复 + 单行不跨页断开
    for ti, t in enumerate(doc.tables):
        rows = t._tbl.findall(qn("w:tr"))
        for ri, tr in enumerate(rows):
            trPr = tr.find(qn("w:trPr"))
            if trPr is None or trPr.find(qn("w:cantSplit")) is None:
                issues.append("表 %d 第 %d 行未设置「禁止跨页断行」" % (ti + 1, ri + 1))
            if ri == 0:
                if trPr is None or trPr.find(qn("w:tblHeader")) is None:
                    issues.append("表 %d 首行未设置「跨页重复标题行」" % (ti + 1))

    # 表格：应为三线表
    for ti, t in enumerate(doc.tables):
        borders = t._tbl.find(qn("w:tblPr") + "/" + qn("w:tblBorders"))
        if borders is None:
            issues.append("表 %d 未设置三线表边框" % (ti + 1))
            continue
        if borders.find(qn("w:left")) is not None and \
           borders.find(qn("w:left")).get(qn("w:val")) != "none":
            issues.append("表 %d 存在竖线，不是标准三线表" % (ti + 1))

    return issues
