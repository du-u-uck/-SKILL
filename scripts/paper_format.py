#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""paper_format —— 中文论文格式命令行工具。

用法：
  python paper_format.py build   输出.docx [--no-examples]
  python paper_format.py format  输入.docx [-o 输出.docx] [--no-quote] [--no-table]
  python paper_format.py quote   输入.docx [-o 输出.docx]
  python paper_format.py compose 规格.json [-o 输出.docx]
  python paper_format.py verify  输入.docx
  python paper_format.py renumber 输入.docx [-o 输出.docx]
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt

import chartlib as C
import paperlib as L


# --------------------------------------------------------------------------
# 占位符渲染：{{cite:1,2}}  {{fig:1}}  {{tab:1}}
# --------------------------------------------------------------------------

PLACEHOLDER_RE = re.compile(r"\{\{\s*(cite:[0-9,\s]+|fig:\d+|tab:\d+)\s*\}\}")


def _add_text_run(paragraph, text, bold=False):
    if not text:
        return
    r = paragraph.add_run(text)
    if bold:
        r.bold = True
    return r


def render_text(paragraph, text, ctx):
    """把带占位符的文本渲染进段落，替换为域（交叉引用 / 上标引用）。"""
    pos = 0
    for m in PLACEHOLDER_RE.finditer(text):
        _add_text_run(paragraph, text[pos:m.start()])
        pos = m.end()
        spec = m.group(1)
        if spec.startswith("cite:"):
            slots = [int(x) for x in spec[5:].replace(" ", "").split(",") if x]
            limit = ctx.get("ref_count", 0)
            bad = [n for n in slots if n > limit]
            if bad:
                print("  ! 引用序号 %s 超出参考文献条目数（%d），已忽略"
                      % (",".join(str(b) for b in bad), limit))
            slots = [n for n in slots if n <= limit]
            if slots:
                L.add_citation(paragraph, slots)
        elif spec.startswith("fig:"):
            n = int(spec[4:])
            bk = ctx["fig_bookmarks"].get(n)
            if bk:
                L.add_crossref(paragraph, bk + "_num", prefix="图")
            else:
                print("  ! 交叉引用 {{fig:%d}} 未找到对应图/图表" % n)
        elif spec.startswith("tab:"):
            n = int(spec[4:])
            bk = ctx["tab_bookmarks"].get(n)
            if bk:
                L.add_crossref(paragraph, bk + "_num", prefix="表")
            else:
                print("  ! 交叉引用 {{tab:%d}} 未找到对应表格" % n)
    _add_text_run(paragraph, text[pos:])


# --------------------------------------------------------------------------
# compose：由 JSON 规格生成完整论文
# --------------------------------------------------------------------------

def compose(spec: dict, out_path: str):
    doc = Document()
    L.setup_page(doc)
    L.setup_doc_defaults(doc)
    L.build_multilevel(doc)
    L.build_styles(doc)
    L.build_footer(doc)

    body = doc.element.body
    for p in body.findall(L.qn("w:p")):
        body.remove(p)

    # 按文档出现顺序收集所有图与表，建立 题注序号 -> 书签 映射
    seq_figs, seq_tabs = [], []

    def collect(blocks):
        for blk in blocks or []:
            kind = blk.get("type", "p")
            if kind in ("figure", "chart"):     # 图与原生图表共用「图N」编号
                seq_figs.append(blk)
            elif kind == "table":
                seq_tabs.append(blk)

    def collect_sections(sections):
        for sec in sections or []:
            collect(sec.get("content"))
            collect_sections(sec.get("subsections"))

    # 文档实际顺序：各节内容 -> 顶层 figures -> 顶层 tables
    collect_sections(spec.get("sections"))
    collect(spec.get("figures"))
    collect(spec.get("tables"))

    n_refs = 0
    for grp in (spec.get("references") or []):
        if isinstance(grp, dict):
            n_refs += len(grp.get("items", []) or [])
        else:
            n_refs += 1

    ctx = {"fig_bookmarks": {}, "tab_bookmarks": {}, "ref_count": n_refs}
    for i in range(len(seq_figs)):
        ctx["fig_bookmarks"][i + 1] = "_Ref_Figure_%d" % (i + 1)
    for i in range(len(seq_tabs)):
        ctx["tab_bookmarks"][i + 1] = "_Ref_Table_%d" % (i + 1)

    def add_p(text="", style=None):
        p = doc.add_paragraph()
        if style:
            p.style = doc.styles[style]
        if text:
            if style == L.STYLE_BODY or style is None:
                render_text(p, text, ctx)
            else:
                p.add_run(text)
                if style in (L.STYLE_ABSTRACT, L.STYLE_KEYWORDS):
                    L.mark_leading_label(p)
        return p

    if spec.get("title"):
        add_p(spec["title"], L.STYLE_TITLE)
    if spec.get("subtitle"):
        add_p(spec["subtitle"], L.STYLE_SUBTITLE)
    add_p("")
    if spec.get("abstract"):
        add_p("摘  要：" + spec["abstract"], L.STYLE_ABSTRACT)
    if spec.get("keywords"):
        add_p("关键词：" + "  ".join(spec["keywords"]), L.STYLE_KEYWORDS)
    add_p("")
    add_p("")

    fig_seq = 0
    tab_seq = 0

    def emit_blocks(blocks):
        nonlocal fig_seq, tab_seq
        for blk in blocks:
            kind = blk.get("type", "p")
            if kind == "p":
                add_p(blk.get("text", ""), L.STYLE_BODY)
            elif kind == "figure":
                fig_seq += 1
                if blk.get("file"):
                    if os.path.isfile(blk["file"]):
                        # 自动等比缩放到版心内（宽 ≤ 14cm，高 ≤ 8cm）
                        L.insert_picture(doc, blk["file"],
                                         max_width_cm=blk.get("max_width_cm",
                                                              L.MAX_IMAGE_WIDTH_CM),
                                         max_height_cm=blk.get("max_height_cm",
                                                               L.MAX_IMAGE_HEIGHT_CM))
                    else:
                        print("  ! 图片不存在，已跳过：%s" % blk["file"])
                L.insert_caption(doc, "图", blk.get("caption", ""),
                                 anchor=doc.paragraphs[-1], where="after")
            elif kind == "chart":
                fig_seq += 1
                # Word 原生图表（DrawingML Chart），不是 Python 画的位图
                cspec = {
                    "type": blk.get("chart_type", "column"),
                    "categories": blk.get("categories", []),
                    "series": blk.get("series", []),
                    "title": blk.get("title", ""),
                    "legend": blk.get("legend", None),
                    "data_labels": blk.get("data_labels", False),
                }
                C.insert_chart(doc, cspec, anchor=doc.paragraphs[-1], where="after",
                               width_cm=blk.get("width_cm", L.DEFAULT_CHART_WIDTH_CM),
                               height_cm=blk.get("height_cm", L.DEFAULT_CHART_HEIGHT_CM))
                L.insert_caption(doc, "图", blk.get("caption", ""),
                                 anchor=doc.paragraphs[-1], where="after")
            elif kind == "table":
                tab_seq += 1
                data = blk["data"]
                tbl = doc.add_table(rows=len(data), cols=max(len(r) for r in data))
                tbl.alignment = 1
                for i, row in enumerate(data):
                    for j, val in enumerate(row):
                        cell = tbl.cell(i, j)
                        cell.text = ""
                        cell.paragraphs[0].style = doc.styles[L.STYLE_TABLE_TEXT]
                        cell.paragraphs[0].add_run(str(val))
                L.format_three_line_table(tbl, header_rows=blk.get("header_rows", 1))
                cap_p, _b, _n = L.make_caption_paragraph(doc, "表", blk.get("caption", ""), "Table")
                tbl._tbl.addprevious(cap_p)

    def emit_section(sec, level=1):
        style = "Heading %d" % min(level, 3)
        add_p(sec.get("heading", ""), style)
        if sec.get("content"):
            emit_blocks(sec["content"])
        for para in sec.get("paragraphs", []) or []:
            add_p(para, L.STYLE_BODY)
        for sub in sec.get("subsections", []) or []:
            emit_section(sub, min(level + 1, 3))

    for sec in spec.get("sections", []) or []:
        emit_section(sec, sec.get("level", 1))

    if spec.get("figures"):
        emit_blocks([{"type": "figure", **f} for f in spec["figures"]])
    if spec.get("tables"):
        emit_blocks([{"type": "table", **t} for t in spec["tables"]])

    # 参考文献（分节分页）
    refs = spec.get("references") or []
    if refs:
        head = add_p("参考文献", L.STYLE_REF_HEADING)
        L.insert_section_break_before(doc, head)
        if isinstance(refs[0], dict):
            # 分组形式：{"category": "（1）期刊文章", "items": ["条目1", "条目2"]}
            for group in refs:
                if group.get("category"):
                    add_p(group["category"], L.STYLE_REF_CATEGORY)
                for item in group.get("items", []) or []:
                    L.add_reference(doc, item)
        else:
            for r in refs:
                L.add_reference(doc, r)

    # 附录（分节分页）
    appendix = spec.get("appendix") or []
    if appendix:
        head = add_p("附录", L.STYLE_APPENDIX)
        L.insert_section_break_before(doc, head)
        for item in appendix:
            if isinstance(item, str):
                add_p(item, L.STYLE_BODY)
            else:
                add_p(item.get("heading", ""), L.STYLE_APPENDIX_HEADING)
                for para in item.get("paragraphs", []) or []:
                    add_p(para, L.STYLE_BODY)

    L.fix_quotes(doc)
    doc.save(out_path)
    return out_path


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

def cmd_build(args):
    L.build_template(args.out, with_examples=not args.no_examples)
    print("已生成模板：%s" % args.out)


def cmd_format(args):
    doc = Document(args.input)
    report = L.apply_format(doc,
                            fix_quote=not args.no_quote,
                            three_line=not args.no_table,
                            section_breaks=not args.no_section,
                            strip_direct=not args.keep_direct,
                            keep_emphasis=not args.no_emphasis,
                            auto_heading=args.auto_heading)
    out = args.output or args.input
    doc.save(out)
    for line in report:
        print(" - " + line)
    print("已保存：%s" % out)


def cmd_quote(args):
    doc = Document(args.input)
    n = L.fix_quotes(doc, verbose=True)
    out = args.output or args.input
    doc.save(out)
    print("已保存：%s（修正 %d 处）" % (out, n))


def cmd_compose(args):
    with open(args.spec, "r", encoding="utf-8") as f:
        spec = json.load(f)
    out = args.output or "论文.docx"
    compose(spec, out)
    print("已生成：%s" % out)


def cmd_chart(args):
    """在既有文档中插入 Word 原生图表。"""
    doc = Document(args.input)
    if args.data:
        with open(args.data, "r", encoding="utf-8-sig") as f:
            categories, series = C.parse_csv(f.read())
    elif args.data_json:
        payload = json.loads(args.data_json)
        categories = payload.get("categories", [])
        series = payload.get("series", [])
    else:
        raise SystemExit("需要 --data <csv 文件> 或 --data-json <JSON>")

    spec = {
        "type": args.type,
        "categories": categories,
        "series": series,
        "title": args.title or "",
        "data_labels": args.data_labels,
    }

    anchor = None
    where = "after"
    if args.after is not None:
        paras = doc.paragraphs
        if not (0 <= args.after < len(paras)):
            raise SystemExit("--after 超出段落范围（共 %d 段）" % len(paras))
        anchor = paras[args.after]
    else:
        for p in doc.paragraphs:
            if p.text.strip() in ("参考文献", "附录"):
                # 必须插到分节符之前，否则图表会落到下一节、与标题同页
                brk = L.preceding_section_break_paragraph(p)
                if brk is not None:
                    anchor = brk
                    where = "before"
                else:
                    anchor = p
                    where = "before"
                break

    p_el = C.insert_chart(doc, spec, anchor=anchor, where=where,
                          width_cm=args.width, height_cm=args.height)
    from docx.text.paragraph import Paragraph
    chart_p = Paragraph(p_el, doc)
    if args.caption:
        L.insert_caption(doc, "图", args.caption, anchor=chart_p, where="after")
    L.fix_quotes(doc)
    out = args.output or args.input
    doc.save(out)
    print("已插入 %s 图表%s" % (args.type, ("：「%s」" % args.caption) if args.caption else ""))
    print("已保存：%s" % out)


def cmd_verify(args):
    doc = Document(args.input)
    issues = L.verify_format(doc)
    if not issues:
        print("✔ 合规检查通过，未发现问题。")
        return 0
    print("发现 %d 个问题：" % len(issues))
    for i in issues:
        print("  ✗ " + i)
    return 1


def cmd_renumber(args):
    doc = Document(args.input)
    changed = L.renumber_references(doc)
    out = args.output or args.input
    doc.save(out)
    print("已重排 %d 条参考文献序号。" % len(changed))
    for old, new in changed:
        print("  %s → %s" % (old, new))


def main(argv=None):
    ap = argparse.ArgumentParser(description="中文论文格式工具")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("build", help="生成论文模板")
    p.add_argument("out")
    p.add_argument("--no-examples", action="store_true")
    p.set_defaults(func=cmd_build)

    p = sub.add_parser("format", help="对既有文档套用格式规范")
    p.add_argument("input")
    p.add_argument("-o", "--output")
    p.add_argument("--no-quote", action="store_true")
    p.add_argument("--no-table", action="store_true")
    p.add_argument("--no-section", action="store_true")
    p.add_argument("--keep-direct", action="store_true",
                   help="保留 run 上的直接字体/字号设置")
    p.add_argument("--no-emphasis", action="store_true",
                   help="同时清除加粗/斜体等强调格式")
    p.add_argument("--auto-heading", action="store_true",
                   help="按「一、/（一）/1.」前缀自动识别标题并改用多级列表编号")
    p.set_defaults(func=cmd_format)

    p = sub.add_parser("quote", help="仅修正中文双引号字体")
    p.add_argument("input")
    p.add_argument("-o", "--output")
    p.set_defaults(func=cmd_quote)

    p = sub.add_parser("compose", help="依据 JSON 规格生成论文")
    p.add_argument("spec")
    p.add_argument("-o", "--output")
    p.set_defaults(func=cmd_compose)

    p = sub.add_parser("chart", help="插入 Word 原生图表（柱状图/饼图/折线图等）")
    p.add_argument("input")
    p.add_argument("-o", "--output")
    p.add_argument("--type", default="column",
                   choices=list(C.CHART_TYPES.keys()),
                   help="column 柱状图 / bar 条形图 / pie 饼图 / line 折线图 / "
                        "area 面积图 / doughnut 环形图 / scatter 散点图 / stacked 堆积柱状图")
    p.add_argument("--data", help="CSV：第一列为类别，其余为各系列")
    p.add_argument("--data-json", help="JSON：{categories:[], series:[{name,values}]}")
    p.add_argument("--title", default="")
    p.add_argument("--caption", default="", help="图题（不含“图N”，序号自动生成）")
    p.add_argument("--width", type=float, default=12.0, help="图宽 cm，默认 12")
    p.add_argument("--height", type=float, default=7.0, help="图高 cm，默认 7")
    p.add_argument("--after", type=int, default=None,
                   help="插到第 N 个正文段之后（0 起）")
    p.add_argument("--data-labels", action="store_true", help="显示数据标签")
    p.set_defaults(func=cmd_chart)

    p = sub.add_parser("verify", help="合规自检")
    p.add_argument("input")
    p.set_defaults(func=cmd_verify)

    p = sub.add_parser("renumber", help="重排参考文献序号")
    p.add_argument("input")
    p.add_argument("-o", "--output")
    p.set_defaults(func=cmd_renumber)

    args = ap.parse_args(argv)
    return args.func(args) or 0


if __name__ == "__main__":
    sys.exit(main())
