# -*- coding: utf-8 -*-
"""chartlib —— 生成 Word 原生图表（DrawingML Chart），不使用 Python 绘图。

生成的图表与 Word「插入 → 图表」得到的对象完全一致：
  * document.xml 里是 <w:drawing> 引用 c:chart 部件
  * 图表部件 /word/charts/chartN.xml
  * 内嵌数据工作簿 /word/embeddings/Microsoft_Excel_WorksheetN.xlsx
在 Word 中双击图表即可进入「编辑数据」，用 Excel 修改数据。

支持的图表类型：柱状图(column)、条形图(bar)、饼图(pie)、折线图(line)、
面积图(area)、环形图(doughnut)、散点图(scatter)。
"""
from __future__ import annotations

import io
from typing import List, Optional, Sequence

from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml.ns import qn
from docx.oxml.parser import parse_xml
from docx.shared import Cm

CT_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
CT_SML_SHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
RT_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
RT_PACKAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"

NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

FONT_CN = "宋体"
FONT_CN_TITLE = "黑体"
FONT_EN = "Times New Roman"

# 类型 -> (绘图元素, 附加参数)
CHART_TYPES = {
    "column":    ("barChart", {"barDir": "col", "grouping": "clustered", "axes": "xy"}),
    "bar":       ("barChart", {"barDir": "bar", "grouping": "clustered", "axes": "xy"}),
    "stacked":   ("barChart", {"barDir": "col", "grouping": "stacked", "axes": "xy"}),
    "pie":       ("pieChart", {"axes": "none"}),
    "doughnut":  ("doughnutChart", {"axes": "none", "holeSize": 50}),
    "line":      ("lineChart", {"grouping": "standard", "axes": "xy"}),
    "area":      ("areaChart", {"grouping": "standard", "axes": "xy"}),
    "scatter":   ("scatterChart", {"scatterStyle": "lineMarker", "axes": "xy2"}),
}


def _esc(v) -> str:
    s = str(v)
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
             .replace('"', "&quot;"))


def _txpr(size=1000, bold=False, cn=FONT_CN) -> str:
    return (
        "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>"
        "<a:defRPr sz=\"{sz}\" b=\"{b}\"><a:latin typeface=\"{en}\"/>"
        "<a:ea typeface=\"{cn}\"/></a:defRPr></a:pPr>"
        "<a:endParaRPr lang=\"zh-CN\"/></a:p></c:txPr>"
        .format(sz=size, b="1" if bold else "0", en=FONT_EN, cn=cn)
    )


def _str_ref(col_letter: str, values: Sequence[str]) -> str:
    n = len(values)
    pts = "".join('<c:pt idx="%d"><c:v>%s</c:v></c:pt>' % (i, _esc(v))
                   for i, v in enumerate(values))
    return (
        "<c:strRef><c:f>Sheet1!${col}$2:${col}${last}</c:f>"
        "<c:strCache><c:ptCount val=\"{n}\"/>{pts}</c:strCache></c:strRef>"
        .format(col=col_letter, last=n + 1, n=n, pts=pts)
    )


def _num_ref(col_letter: str, values: Sequence) -> str:
    n = len(values)
    pts = "".join('<c:pt idx="%d"><c:v>%s</c:v></c:pt>' % (i, _esc(v))
                   for i, v in enumerate(values))
    return (
        "<c:numRef><c:f>Sheet1!${col}$2:${col}${last}</c:f>"
        "<c:numCache><c:formatCode>General</c:formatCode>"
        "<c:ptCount val=\"{n}\"/>{pts}</c:numCache></c:numRef>"
        .format(col=col_letter, last=n + 1, n=n, pts=pts)
    )


def _col_letter(idx: int) -> str:
    """0 -> A, 1 -> B, ... 25 -> Z"""
    s = ""
    idx += 1
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def build_chart_xml(spec: dict) -> bytes:
    """生成 chartN.xml 的内容。"""
    ctype = spec.get("type", "column")
    if ctype not in CHART_TYPES:
        raise ValueError("不支持的图表类型：%s（可选 %s）"
                         % (ctype, ", ".join(CHART_TYPES)))
    element, opts = CHART_TYPES[ctype]
    categories = [str(c) for c in spec.get("categories", [])]
    series = spec.get("series", []) or []
    if not series:
        raise ValueError("图表至少需要一个数据系列")

    # ---- 数据系列 ----
    sers = []
    for j, s in enumerate(series):
        col = _col_letter(j + 1)
        name = s.get("name", "系列%d" % (j + 1))
        values = s.get("values", [])
        tx = ('<c:tx><c:strRef><c:f>Sheet1!$%s$1</c:f><c:strCache>'
              '<c:ptCount val="1"/><c:pt idx="0"><c:v>%s</c:v></c:pt>'
              '</c:strCache></c:strRef></c:tx>' % (col, _esc(name)))
        cat = "<c:cat>%s</c:cat>" % _str_ref("A", categories)
        val = "<c:val>%s</c:val>" % _num_ref(col, values)
        dlbls = ""
        if spec.get("data_labels"):
            dlbls = ('<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/>'
                     '<c:showCatName val="0"/><c:showSerName val="0"/>'
                     '<c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>')
        if element == "barChart":
            body = (tx + '<c:invertIfNegative val="0"/>' + dlbls + cat + val)
        elif element == "lineChart":
            body = (tx + '<c:marker><c:symbol val="circle"/><c:size val="5"/>'
                    '</c:marker>' + dlbls + cat + val + '<c:smooth val="0"/>')
        elif element == "areaChart":
            body = tx + dlbls + cat + val
        elif element == "scatterChart":
            xval = "<c:xVal>%s</c:xVal>" % _num_ref("A", categories)
            yval = "<c:yVal>%s</c:yVal>" % _num_ref(col, values)
            body = (tx + '<c:marker><c:symbol val="circle"/><c:size val="5"/>'
                    '</c:marker>' + dlbls + xval + yval + '<c:smooth val="0"/>')
        else:  # pie / doughnut
            body = tx + dlbls + cat + val
        sers.append(
            '<c:ser><c:idx val="%d"/><c:order val="%d"/>%s</c:ser>'
            % (j, j, body)
        )

    # ---- 绘图区 ----
    ax_cat = "111111111"
    ax_val = "222222222"
    if element == "barChart":
        plot = ('<c:barChart><c:barDir val="%s"/><c:grouping val="%s"/>'
                '<c:varyColors val="0"/>%s<c:gapWidth val="80"/>'
                '<c:overlap val="%d"/><c:axId val="%s"/><c:axId val="%s"/></c:barChart>'
                % (opts["barDir"], opts["grouping"], "".join(sers),
                   100 if opts["grouping"] == "stacked" else -27, ax_cat, ax_val))
    elif element == "lineChart":
        plot = ('<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>%s'
                '<c:marker val="1"/><c:axId val="%s"/><c:axId val="%s"/></c:lineChart>'
                % ("".join(sers), ax_cat, ax_val))
    elif element == "areaChart":
        plot = ('<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>%s'
                '<c:axId val="%s"/><c:axId val="%s"/></c:areaChart>'
                % ("".join(sers), ax_cat, ax_val))
    elif element == "scatterChart":
        plot = ('<c:scatterChart><c:scatterStyle val="lineMarker"/>'
                '<c:varyColors val="0"/>%s<c:axId val="%s"/><c:axId val="%s"/>'
                '</c:scatterChart>' % ("".join(sers), ax_cat, ax_val))
    elif element == "pieChart":
        plot = ('<c:pieChart><c:varyColors val="1"/>%s<c:firstSliceAng val="0"/>'
                '</c:pieChart>' % "".join(sers))
    else:  # doughnutChart
        plot = ('<c:doughnutChart><c:varyColors val="1"/>%s'
                '<c:firstSliceAng val="0"/><c:holeSize val="%d"/></c:doughnutChart>'
                % ("".join(sers), opts.get("holeSize", 50)))

    # ---- 坐标轴 ----
    cat_ax = (
        '<c:catAx><c:axId val="%s"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
        '<c:delete val="0"/><c:axPos val="b"/>'
        '<c:numFmt formatCode="General" sourceLinked="1"/>'
        '<c:majorTickMark val="none"/><c:minorTickMark val="none"/>'
        '<c:tickLblPos val="nextTo"/>%s'
        '<c:crossAx val="%s"/><c:crosses val="autoZero"/><c:auto val="1"/>'
        '<c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/>'
        '</c:catAx>' % (ax_cat, _txpr(), ax_val)
    )
    val_ax = (
        '<c:valAx><c:axId val="%s"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
        '<c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/>'
        '<c:numFmt formatCode="General" sourceLinked="1"/>'
        '<c:majorTickMark val="none"/><c:minorTickMark val="none"/>'
        '<c:tickLblPos val="nextTo"/>%s'
        '<c:crossAx val="%s"/><c:crosses val="autoZero"/>'
        '<c:crossBetween val="between"/></c:valAx>' % (ax_val, _txpr(), ax_cat)
    )
    if opts.get("axes") == "xy":
        axes = cat_ax + val_ax
    elif opts.get("axes") == "xy2":
        axes = (cat_ax.replace("<c:catAx>", "<c:valAx>")
                .replace("</c:catAx>", "</c:valAx>")
                .replace('<c:axPos val="b"/>', '<c:axPos val="b"/>')
                + val_ax.replace("<c:valAx>", "<c:valAx>"))
    else:
        axes = ""

    # ---- 标题 ----
    title_xml = ""
    if spec.get("title"):
        title_xml = (
            "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr>"
            '<a:defRPr sz="1100" b="1"><a:latin typeface="%s"/><a:ea typeface="%s"/></a:defRPr>'
            "</a:pPr><a:r><a:rPr lang=\"zh-CN\" sz=\"1100\" b=\"1\"/>"
            "<a:t>%s</a:t></a:r></a:p></c:rich></c:tx>"
            '<c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>'
            % (FONT_EN, FONT_CN_TITLE, _esc(spec["title"]))
        )
    else:
        title_xml = '<c:autoTitleDeleted val="1"/>'

    legend = ""
    if spec.get("legend", len(series) > 1 or ctype in ("pie", "doughnut")):
        legend = ('<c:legend><c:legendPos val="r"/><c:overlay val="0"/>%s</c:legend>'
                  % _txpr())

    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<c:chartSpace xmlns:c="%(c)s" xmlns:a="%(a)s" xmlns:r="%(r)s">'
        '<c:roundedCorners val="0"/>'
        '<c:chart>%(title)s'
        '<c:plotArea><c:layout/>%(plot)s%(axes)s<c:spPr><a:noFill/>'
        '<a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>'
        '%(legend)s'
        '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>'
        '</c:chart>'
        '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
        '%(txpr)s'
        '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>'
        '</c:chartSpace>'
        % {"c": NS_C, "a": NS_A, "r": NS_R, "title": title_xml, "plot": plot,
           "axes": axes, "legend": legend, "txpr": _txpr()}
    )
    return xml.encode("utf-8")


def build_embedded_workbook(spec: dict) -> bytes:
    """生成内嵌的 Excel 工作簿（图表的「编辑数据」源）。"""
    from openpyxl import Workbook

    categories = [str(c) for c in spec.get("categories", [])]
    series = spec.get("series", []) or []

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for j, s in enumerate(series):
        ws.cell(row=1, column=2 + j, value=s.get("name", "系列%d" % (j + 1)))
    for i, cat in enumerate(categories):
        ws.cell(row=2 + i, column=1, value=cat)
    for j, s in enumerate(series):
        for i, v in enumerate(s.get("values", [])):
            ws.cell(row=2 + i, column=2 + j, value=v)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _next_partname(package, template: str):
    n = 1
    while True:
        uri = PackURI(template % n)
        if not any(p.partname == uri for p in package.iter_parts()):
            return uri
        n += 1


def insert_chart(doc, spec: dict, anchor=None, where: str = "after",
                 width_cm: float = 12.0, height_cm: float = 7.0):
    """把原生图表插入文档，返回图表段落。

    doc     python-docx Document
    spec    {"type","categories","series","title","legend","data_labels"}
    anchor  锚点段落；None 表示追加到文档末尾
    """
    package = doc.part.package
    chart_uri = _next_partname(package, "/word/charts/chart%d.xml")
    xlsx_uri = _next_partname(package,
                              "/word/embeddings/Microsoft_Excel_Worksheet%d.xlsx")

    chart_part = Part(chart_uri, CT_CHART, build_chart_xml(spec), package)
    xlsx_part = Part(xlsx_uri, CT_SML_SHEET, build_embedded_workbook(spec), package)
    chart_part.relate_to(xlsx_part, RT_PACKAGE)
    rId = doc.part.relate_to(chart_part, RT_CHART)

    cx = int(Cm(width_cm))
    cy = int(Cm(height_cm))
    docpr_id = 1000 + len(list(doc.element.iter(qn("w:drawing"))))

    p_xml = (
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:pPr><w:keepNext/><w:spacing w:before="60" w:after="60" w:line="240" '
        'w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>'
        '<w:r><w:drawing xmlns:wp="%(wp)s">'
        '<wp:inline distT="0" distB="0" distL="0" distR="0">'
        '<wp:extent cx="%(cx)d" cy="%(cy)d"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:docPr id="%(id)d" name="图表 %(id)d"/>'
        '<wp:cNvGraphicFramePr/>'
        '<a:graphic xmlns:a="%(a)s">'
        '<a:graphicData uri="%(c)s">'
        '<c:chart xmlns:c="%(c)s" xmlns:r="%(r)s" r:id="%(rid)s"/>'
        '</a:graphicData></a:graphic>'
        '</wp:inline></w:drawing></w:r></w:p>'
        % {"wp": NS_WP, "a": NS_A, "c": NS_C, "r": NS_R,
           "cx": cx, "cy": cy, "id": docpr_id, "rid": rId}
    )
    p_el = parse_xml(p_xml)

    if anchor is not None:
        if where == "before":
            anchor._p.addprevious(p_el)
        else:
            anchor._p.addnext(p_el)
    else:
        body = doc.element.body
        sect = body.find(qn("w:sectPr"))
        if sect is not None:
            sect.addprevious(p_el)
        else:
            body.append(p_el)
    return p_el


def parse_csv(text: str):
    """解析 CSV：第一列为类别，其余每列一个系列。

    返回 (categories, series)
    """
    import csv
    rows = list(csv.reader(io.StringIO(text)))
    rows = [r for r in rows if any(c.strip() for c in r)]
    if len(rows) < 2:
        raise ValueError("CSV 至少需要表头和一行数据")
    header = [c.strip() for c in rows[0]]
    categories = [r[0].strip() for r in rows[1:]]
    series = []
    for j in range(1, len(header)):
        name = header[j] or "系列%d" % j
        values = []
        for r in rows[1:]:
            raw = r[j].strip() if j < len(r) else ""
            try:
                values.append(float(raw))
            except ValueError:
                values.append(0)
        series.append({"name": name, "values": values})
    return categories, series
