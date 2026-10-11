# -*- coding: utf-8 -*-
"""element_order.py — 从本地 OOXML XSD 里提取「子元素顺序」，生成可查的顺序表。

为什么需要它：OOXML 用 `xsd:sequence` 定义子元素顺序，顺序即语义。插错位置
的后果不是「警告」而是 Word 直接报「内容有问题」且拒绝打开。凭记忆写顺序表
一定会出错（settings 有 98 个子元素，pPr/sectPr/图表各有几十个），所以这张表
必须从权威 XSD 生成。

用法：
    python pro/scripts/element_order.py --list                  # 可查询的类型清单
    python pro/scripts/element_order.py --type CT_SectPr        # 单个类型的顺序
    python pro/scripts/element_order.py --type CT_PPr CT_DLbls  # 多个类型
    python pro/scripts/element_order.py --doc                   # 生成 references/element-order.md
    python pro/scripts/element_order.py --type CT_Settings --json

输出记号：
    name      必需且恰好一次
    name?     可选（minOccurs="0"）
    name*     可重复（maxOccurs="unbounded"）
    name?*    可选且可重复
    (a | b)   xsd:choice —— 这一组里只能选一个成员
    <xsd:any> 通配符槽位（扩展点）
    [Group↻]  group 自引用，已就地展开过，避免无限递归
"""
import argparse
import json
import os
import re
import sys

try:
    from lxml import etree
except ImportError:
    print("需要 lxml：请用 envs/default 的 python 运行本脚本", file=sys.stderr)
    raise

XSD = "http://www.w3.org/2001/XMLSchema"
HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_ROOT = os.path.dirname(os.path.dirname(HERE))
SCHEMA_DIR = os.path.join(SKILL_ROOT, "pro", "schemas")
DOC_PATH = os.path.join(SKILL_ROOT, "references", "element-order.md")

# 生成文档时按此分组顺序输出；每一项是 (章节名, [类型名...], 备注)
PRESETS = [
    ("正文与段落", [
        "CT_Body", "CT_P", "CT_PPr", "CT_PPrBase", "CT_RPr", "CT_ParaRPr",
        "CT_R", "CT_TrackChange", "CT_PPrChange", "CT_RPrChange",
    ], "段落是最常被改写的地方：pPr 与 rPr 里的任何新增子元素都必须落进上面的槽位。"),
    ("节与页面", [
        "CT_SectPr", "CT_PageMar", "CT_PageSz", "CT_Columns", "CT_PageBorders",
        "CT_PageNumber", "CT_LineNumber", "CT_DocGrid", "CT_HdrFtr",
    ], "sectPr 的 headerReference/footerReference 必须排在最前，type 决定分节方式。"),
    ("表格", [
        "CT_Tbl", "CT_TblPr", "CT_TblPrBase", "CT_TblPrEx", "CT_Tr", "CT_TrPr",
        "CT_Tc", "CT_TcPr", "CT_TblGrid", "CT_TblWidth", "CT_TblBorders",
        "CT_TblCellMar", "CT_TblLayoutType", "CT_TblLook", "CT_TcBorders",
        "CT_Shd", "CT_VMerge", "CT_HMerge",
    ], "tcPr 里边框/底纹/宽度有固定次序；tblPr 的 tblStyle 必须第一个出现。"),
    ("修订与批注", [
        "CT_Markup", "CT_TrackChangesView", "CT_Comment", "CT_Comments",
        "CT_CellIns", "CT_CellDel",
    ], "修订类元素（ins/del/moveFrom/moveTo）包住的是内容，不是兄弟节点，插错会破坏树。"),
    ("设置部件 w:settings", [
        "CT_Settings", "CT_MailMerge", "CT_DocVars", "CT_Compat", "CT_ProofState",
        "CT_WriteProtection", "CT_DocProtect",
    ], "全库最长的 sequence：98 个子元素固定顺序，`apply` 的 `settings.elements` 会自动按此排序插入。CT_MailMerge 里的 `dataType` 是**必需项**——只写 mainDocumentType + destination 会直接校验失败。"),
    ("样式与编号", [
        "CT_Styles", "CT_Style", "CT_LatentStyles", "CT_LsdException",
        "CT_Numbering", "CT_Num", "CT_AbstractNum", "CT_Lvl", "CT_NumPr",
        "CT_DecimalNumber", "CT_MultiLevelType",
    ], "styles.xml / numbering.xml 的部件级顺序；lvl 内的 pPr/rPr 各自再走上面的段落表。"),
    ("构建基块与内容控件", [
        "CT_GlossaryDocument", "CT_DocParts", "CT_DocPart", "CT_DocPartPr",
        "CT_DocPartName", "CT_DocPartCategory", "CT_DocPartTypes",
        "CT_DocPartBehavior", "CT_DocPartGallery",
        "CT_SdtBlock", "CT_SdtRun", "CT_SdtPr", "CT_SdtContentRun",
    ], "gallery 只能挂在 category 里，且取值受 CT_DocPartGallery 枚举约束。"),
    ("图表 chart", [
        "CT_ChartSpace", "CT_Chart", "CT_PlotArea", "CT_BarChart", "CT_BarSer",
        "CT_LineChart", "CT_LineSer", "CT_PieChart", "CT_PieSer",
        "CT_CatAx", "CT_ValAx", "CT_DLbls", "CT_DLbl", "CT_Legend",
        "CT_Title", "CT_Scaling", "CT_NumFmt", "CT_Boolean",
    ], "dLbls 的 Group_DLblShared 是 maxOccurs=1 的 choice —— 标了 showVal 就不能再跟 txPr。"),
    ("绘图与图形", [
        "CT_Inline", "CT_Anchor", "CT_GraphicalObjectFrame", "CT_GraphicalObject",
        "CT_ShapeProperties", "CT_TextBody", "CT_TextParagraph", "CT_TextBodyProperties",
        "CT_Transform2D", "CT_Picture", "CT_BlipFillProperties", "CT_SrcRect",
        "CT_Blip", "CT_GroupShape", "CT_Connector", "CT_Canvas",
    ], "wp:inline / wp:anchor 内部的顺序错一位，Word 会整张图丢掉。"),
    ("WordprocessingML 扩展（2010/2012/2018）", [
        "CT_SdtContentBlock", "CT_SdtContentCell",
    ], "扩展命名空间的元素同样受各自 XSD 约束；它们不能被丢进 mc:Ignorable。"),
]


def q(tag):
    return "{%s}%s" % (XSD, tag)


class SchemaIndex(object):
    """把多个 XSD 里的 complexType / group 按名字索引起来。"""

    def __init__(self, schema_dir):
        self.types = {}
        self.groups = {}
        self.files = []
        if not os.path.isdir(schema_dir):
            raise SystemExit("找不到 schema 目录：%s" % schema_dir)
        for fn in sorted(os.listdir(schema_dir)):
            if not fn.endswith(".xsd"):
                continue
            path = os.path.join(schema_dir, fn)
            try:
                root = etree.parse(path).getroot()
            except Exception:
                continue
            self.files.append(fn)
            for ct in root.iter(q("complexType")):
                name = ct.get("name")
                if name:
                    self.types.setdefault(name, (fn, ct))
            for g in root.iter(q("group")):
                name = g.get("name")
                if name:
                    self.groups.setdefault(name, (fn, g))

    # ── 展开 ───────────────────────────────────────────────────────────
    def _occurs(self, node):
        """返回 (记号, 是否可重复多次)。"""
        try:
            mn = int(node.get("minOccurs", "1") or "1")
        except ValueError:
            mn = 1
        mx = node.get("maxOccurs", "1") or "1"
        if mx == "unbounded":
            return ("?" if mn == 0 else "+"), True
        try:
            mxi = int(mx)
        except ValueError:
            mxi = 1
        if mxi <= 1:
            return ("?" if mn == 0 else ""), False
        return "{%d,%d}" % (mn, mxi), True

    @staticmethod
    def _all_repeatable(items):
        return bool(items) and all(
            any(c in m for c in "*+{") for _t, m, _k in items
        )

    def expand(self, node, stack=()):
        """返回 [(text, mark, kind)] 的扁平列表。"""
        out = []
        for ch in node:
            tag = ch.tag
            if not isinstance(tag, str):
                continue
            if tag == q("element"):
                name = ch.get("name") or (ch.get("ref") or "").split(":")[-1]
                mark, _rep = self._occurs(ch)
                out.append((name, mark, "element"))
            elif tag in (q("sequence"), q("all")):
                out.extend(self.expand(ch, stack))
            elif tag in (q("complexContent"), q("simpleContent")):
                out.extend(self.expand(ch, stack))
            elif tag in (q("extension"), q("restriction")):
                # CT_PPr extends CT_PPrBase 之类的继承链：base 的内容排在前，
                # 自身追加的槽位排在后 —— 这正是查顺序表时最需要的信息。
                base = (ch.get("base") or "").split(":")[-1]
                if tag == q("extension") and base in self.types and base not in stack:
                    out.extend(self.expand(self.types[base][1], stack + (base,)))
                out.extend(self.expand(ch, stack))
            elif tag == q("choice"):
                inner = self.expand(ch, stack)
                joined = " | ".join(i[0] + i[1] for i in inner)
                mark, _rep = self._occurs(ch)
                out.append(("(" + joined + ")", mark, "choice"))
            elif tag == q("group"):
                ref = (ch.get("ref") or "").split(":")[-1]
                mark, rep = self._occurs(ch)
                if ref in stack:
                    out.append(("[%s]" % ref, mark, "cycle"))
                elif ref in self.groups:
                    inner = self.expand(self.groups[ref][1], stack + (ref,))
                    # group ref 自带的 occurs 要保留，否则像
                    # EG_HdrFtrReferences(maxOccurs="6") 会被渲染成"恰好一次"。
                    # 但 minOccurs=0/maxOccurs=1 不必包一层：成员自己就带 `?`。
                    if rep and len(inner) == 1:
                        t, m, k = inner[0]
                        out.append((t, m + mark, k))
                    elif rep and not self._all_repeatable(inner):
                        joined = " ".join(i[0] + i[1] for i in inner)
                        out.append(("(" + joined + ")", mark, "group"))
                    else:
                        out.extend(inner)
                else:
                    out.append(("[%s?]" % ref, mark, "missing"))
            elif tag == q("any"):
                mark, _rep = self._occurs(ch)
                out.append(("<xsd:any>", mark, "any"))
        return out

    def order(self, name):
        if name in self.types:
            fn, node = self.types[name]
            return fn, self.expand(node)
        if name in self.groups:
            fn, node = self.groups[name]
            return fn, self.expand(node)
        return None, None


def render_line(items, width=92, indent="    "):
    """把 [(text, mark, kind)] 排成定宽多行；换行只是排版，顺序是连续的。"""
    pieces = []
    for text, mark, kind in items:
        pieces.append(text + mark)
    lines = []
    cur = ""
    for p in pieces:
        if cur and len(cur) + 1 + len(p) > width:
            lines.append(cur)
            cur = p
        else:
            cur = (cur + " " + p).strip()
    if cur:
        lines.append(cur)
    return "\n".join(indent + l for l in lines)


def cmd_list(index):
    print("complexType（%d 个）：" % len(index.types))
    for n in sorted(index.types):
        print("  %-42s %s" % (n, index.types[n][0]))
    print()
    print("group（%d 个）：" % len(index.groups))
    for n in sorted(index.groups):
        print("  %-42s %s" % (n, index.groups[n][0]))


def cmd_types(index, names, as_json):
    out = {}
    for n in names:
        fn, items = index.order(n)
        if items is None:
            print("!! 未找到类型：%s" % n, file=sys.stderr)
            continue
        out[n] = {
            "source": fn,
            "order": [{"name": t, "occurs": m, "kind": k} for t, m, k in items],
        }
        if not as_json:
            print("== %s  (%s)" % (n, fn))
            print(render_line(items))
            print()
    if as_json:
        print(json.dumps(out, ensure_ascii=False, indent=2))


def flatten_names(items):
    """把 (text, mark, kind) 列表摊平成去重的子元素名清单。"""
    names = []
    for text, _mark, kind in items:
        if kind == "element":
            cands = [text]
        elif kind in ("choice", "group"):
            # 渲染过的 choice/group 文本里只含元素名 + |?*+{}() 这些记号，
            # 直接按 XML 名抽取，避免手工剥括号时漏掉嵌套层。
            cands = re.findall(r"[A-Za-z_][A-Za-z0-9_.\-]*", text)
        else:
            continue
        for c in cands:
            if c in ("xsd", "any") or c in names:
                continue
            names.append(c)
    return names


def cmd_doc(index):
    parts = []
    parts.append("""# OOXML 元素顺序表（自动生成，勿手改）

> 由 `pro/scripts/element_order.py --doc` 从 `pro/schemas/*.xsd` 生成。
> XSD 更新后重跑该命令即可刷新，**不要手工编辑本文件**。

## 为什么顺序是硬约束

OOXML 用 `xsd:sequence` 定义子元素：顺序不是风格问题，而是 schema 的一部分。
顺序错的后果不是「Word 给个警告」，而是 `Schemas validity error` /
「文件内容有问题，Word 无法打开」。因此**任何插入操作都要先决定落点**：

1. 在当前子元素序列里找到「应该排在自己前面的那个元素」，插到它后面；
2. 若前面没有更早的元素，插到第一个位置；
3. 插到末尾前先确认没有「必须收尾」的元素（如 `w:sectPr` 必须留在 `w:body` 最后）。

`apply` 引擎与 `extend` 扩展包都已按本表自动排序（`settings.elements`、
`sectPr` 子元素、`docProps` 子元素等），只有走 `edit` 的自由 `xml` 片段
或自写脚本时才需要人工对照本表。

## 记号

| 记号 | 含义 |
|---|---|
| `name` | 必需，恰好一次 |
| `name?` | 可选（`minOccurs="0"`） |
| `name+` | 一次或多次（`maxOccurs="unbounded"`） |
| `name?*` / `name*` | 可选且可重复 |
| `name{0,6}` | 有上限的重复（如 `sectPr` 里页眉页脚引用最多 6 次） |
| `(a \\| b)` | `xsd:choice`：这一组里**只能选一个** |
| `<xsd:any>` | 通配符槽位（扩展点，见 `ooxml-extension-points.md`） |
| `[GroupName]` | group 自引用；内容已在前面就地展开 |

代码块里的换行只是排版，**顺序沿列连续**。

## 速查：最容易插错的三处

1. **`w:sectPr`** —— `headerReference` / `footerReference` 必须排在最前；
   改页边距（`pgMar`）时要落在 `pgSz` 之后、`pgBorders` 之前。
2. **`w:body`** —— 段落的 `sectPr` 永远是 body 的**最后一个**子元素。
3. **`c:dLbls`** —— `showXxx` 系列与 `txPr` 同在 `maxOccurs=1` 的 choice 里：
   二者**不能并存**，数据标签的字号请改放 `c:catAx` / `c:valAx` 的 `c:txPr`。

## 怎么读必填项

**没有 `?` 的就是必需**（`minOccurs=1`）。造元素之前先扫一遍必需项，能省掉一轮
「not expected / missing required element」的往返。典型例子：

- `CT_MailMerge` → `dataType` 必需（只写 `mainDocumentType` + `destination` 会失败）。
- `CT_DocPartPr` → `name` 必需。
- `CT_Chart` → `plotArea` 必需（`autoTitleDeleted`、`pivotFmts` 等反而可选）。
- `CT_SectPr` → 一个必需项都没有：整节属性全部可选。

遇到 `(a | b | c)` 时先确认「这一组里到底要选哪个」，再确认它排在哪几个兄弟之间 ——
choice 内部成员的**相对位置**同样受前后 sequence 约束。
""")
    for section, names, note in PRESETS:
        parts.append("## %s\n" % section)
        if note:
            parts.append("%s\n" % note)
        for n in names:
            fn, items = index.order(n)
            if items is None:
                continue
            parts.append("### `%s`\n" % n)
            parts.append("来源：`pro/schemas/%s`\n" % fn)
            parts.append("```")
            parts.append(render_line(items))
            parts.append("```\n")
            # 嵌套 choice 的类型（w:p / w:body / w:hdr 等）光看长行很难读，
            # 额外给一份"能放哪些元素"的去重清单。
            if any(k in ("choice", "group") for _t, _m, k in items):
                flat = flatten_names(items)
                parts.append("可出现的子元素（去重）：%s\n"
                             % " ".join("`%s`" % f for f in flat))
    with open(DOC_PATH, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(parts))
    n_types = sum(1 for _s, ns, _n in PRESETS for _ in ns if index.order(_)[1] is not None)
    print("已生成 %s（%d 个类型）" % (DOC_PATH, n_types))


JS_ORDER_TABLES = [
    ("RPR_CHILD_ORDER", "CT_RPr"),
    ("TBL_PR_CHILD_ORDER", "CT_TblPr"),
    ("TC_PR_CHILD_ORDER", "CT_TcPr"),
    ("TR_PR_CHILD_ORDER", "CT_TrPr"),
    ("SECT_PR_CHILD_ORDER", "CT_SectPr"),
    ("PPR_CHILD_ORDER", "CT_PPr"),
    ("SETTINGS_CHILD_ORDER", "CT_Settings"),
]


def cmd_verify(index, js_path):
    """把脚本里手写的顺序表与 XSD 对账。

    手工表只要参与「按顺序插入」就必须与 schema 一致：共同项的相对顺序错了，
    插入就会落到错误位置；这比漏项更危险，所以顺序分歧按 ERROR 报。
    表里故意只列「可能插入的项」是允许的（缺项只做提示）。
    """
    with open(js_path, encoding="utf-8") as f:
        src = f.read()
    errors = 0
    for const, type_name in JS_ORDER_TABLES:
        m = re.search(r"const %s\s*=\s*\[(.*?)\];" % const, src, re.S)
        if not m:
            print("!! %-22s 脚本里找不到该常量" % const)
            errors += 1
            continue
        js_list = re.findall(r'"([^"]+)"', m.group(1))
        fn, items = index.order(type_name)
        if items is None:
            print("!! %-22s XSD 里找不到 %s" % (const, type_name))
            errors += 1
            continue
        xsd = flatten_names(items)
        xsd_set, js_set = set(xsd), set(js_list)
        missing = [x for x in xsd if x not in js_set]
        extra = [x for x in js_list if x not in xsd_set]
        common_js = [x for x in js_list if x in xsd_set]
        common_xsd = [x for x in xsd if x in js_set]
        ordered = common_js == common_xsd
        flag = "OK  " if ordered and not extra else "FAIL"
        if not ordered or extra:
            errors += 1
        print("%s %-22s ↔ %-14s js=%d xsd=%d 未覆盖=%d"
              % (flag, const, type_name, len(js_list), len(xsd), len(missing)))
        if extra:
            print("      表里有、schema 里没有（应为扩展项或笔误）：%s" % ", ".join(extra))
        if not ordered:
            for i, (a, b) in enumerate(zip(common_js, common_xsd)):
                if a != b:
                    print("      顺序分歧 @%d：表里 %s，schema 是 %s" % (i, a, b))
                    break
    print()
    print("顺序表对账：%s" % ("全部一致" if errors == 0 else "%d 处需要复核" % errors))
    return 0 if errors == 0 else 1


def main():
    ap = argparse.ArgumentParser(description="OOXML 子元素顺序查询 / 文档生成")
    ap.add_argument("--type", nargs="+", help="要查询的 complexType / group 名")
    ap.add_argument("--list", action="store_true", help="列出所有可查询的名字")
    ap.add_argument("--doc", action="store_true", help="生成 references/element-order.md")
    ap.add_argument("--verify", metavar="APPLY_JS",
                    help="把 apply.js 里手写的顺序表与 XSD 对账")
    ap.add_argument("--json", action="store_true", help="以 JSON 输出")
    ap.add_argument("--schema-dir", default=SCHEMA_DIR, help="XSD 目录")
    args = ap.parse_args()

    index = SchemaIndex(args.schema_dir)
    if args.list:
        cmd_list(index)
    elif args.doc:
        cmd_doc(index)
    elif args.verify:
        sys.exit(cmd_verify(index, args.verify))
    elif args.type:
        cmd_types(index, args.type, args.json)
    else:
        ap.print_help()


if __name__ == "__main__":
    main()
