# OOXML 扩展点：命名空间、部件注册与 `mc:Ignorable`

引擎（`apply`）覆盖「样式 / 编号 / 页面 / 编辑」这些常规改写。当需求落到引擎之外
——域、脚注、水印、图形、图表、OLE、内容控件、构建基块、自定义 XML —— 就要直接
操作 OOXML 的**扩展点**。这份文档是那部分的知识底座，配合
[`element-order.md`](element-order.md)（顺序）与 `scripts/extend/`（已封装的构件库）
一起用。

## 1. 根元素上的命名空间声明

**症状**：手工拼出来的 `w:pict` / `wp:inline` / `w14:*` 片段让 Word 报
「文件内容有问题」。原因是根元素上根本没有声明这些前缀——模板（`_assets/blank.docx`）
的 `w:document` 通常只声明 `xmlns:w` 与 `xmlns:r`，任何新前缀都必须补上声明。

**做法**：写入部件前先跑一遍「根声明补齐」：

```
wp       http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing
wp14     http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing
wps      http://schemas.microsoft.com/office/word/2010/wordprocessingShape
wpg      http://schemas.microsoft.com/office/word/2010/wordprocessingGroup
wpc      http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas
w14      http://schemas.microsoft.com/office/word/2010/wordml
w15      http://schemas.microsoft.com/office/word/2012/wordml
v        urn:schemas-microsoft-com:vml
o        urn:schemas-microsoft-com:office:office
r        http://schemas.openxmlformats.org/officeDocument/2006/relationships
mc       http://schemas.openxmlformats.org/markup-compatibility/2006
```

已实现：`apply` 的 `ensureRootDeclarations()`（JS）与 `extend/_common.py` 的
`ensure_root_declarations()`（Python）——幂等，可重复调用。

### 1.1 `mc:Ignorable` 只能列「可安全丢弃」的前缀

这是最容易踩、后果最隐蔽的一条：

- `mc:Ignorable="w14 w15 wp14"` —— 正确。这些是纯增强标记，老版本 Word 丢掉它们
  仍然能正常渲染。
- `mc:Ignorable="wps wpg wpc"` —— **错误，会掉内容**。`wps`/`wpg`/`wpc` 承载的是
  真实的图形（文本框、形状、组合、画布）。把它们标成 ignorable，等于告诉不认识的
  消费者「这些随便扔」，Word 或转换器可能整块丢弃图形。

规则：**渲染型命名空间永不进 Ignorable**；只有「丢了也不影响语义」的扩展标记才进。
`ensure_root_declarations(..., ignorable=[...])` 会主动把 `wps/wpg/wpc` 从传入的
ignorable 列表里剔除，防止调用方写错。

## 2. 新增一个部件 = 四件事，缺一即损坏

| 步骤 | 落点 | 漏掉的后果 |
|---|---|---|
| 1. 内容类型 | `[Content_Types].xml` 的 `<Override PartName="/word/x.xml" ContentType="..."/>`，或按扩展名给 `<Default>` | Word 报「内容有问题」，包校验直接失败 |
| 2. 关系 | 源部件的 `.rels` 里加 `<Relationship Id Target Type/>` | 部件存在但无人引用 → 被当作垃圾/校验报 `Unreferenced file` |
| 3. 根声明 | 新部件自己的根元素上补齐命名空间（见 §1） | XML 解析失败 |
| 4. 正文引用 | 文档里放指向该部件的元素（`r:id` / `r:embed`） | 部件是孤儿，视觉效果不出现 |

### 2.1 关系 Target 的相对基准（曾经的真实 bug）

OPC 规定：`Target` 相对**源部件所在目录**解析，**不是**相对 `.rels` 文件自身。

| `.rels` 文件 | 源部件 | Target 基准 |
|---|---|---|
| `_rels/.rels` | 整个包 | 包根（`""`） |
| `word/_rels/document.xml.rels` | `word/document.xml` | `word/` |
| `customXml/item1.xml.rels` | `customXml/item1.xml` | `customXml/` |
| `customXml/_rels/item1.xml.rels` | `customXml/item1.xml` | `customXml/` |

最常见的错法是把第 3 行按「相对 rels 文件自己」算，得出
`customXml/item1.xml.rels/itemProps1.xml` 这种不存在的路径。`docx-master` 两套校验器
（`scripts/_shared/docx-validate.js` 与 `pro/scripts/validators_base.py`）都曾犯过这个错，
已在 2026-10-10 修正为按源部件目录解析。

自检：`customXml/item1.xml.rels` 引用同目录的 `itemProps1.xml` 时，Target 就写
`itemProps1.xml`（不带目录、不带 `../`）。

## 3. 文档属性部件（`docProps/`）

`docProps/core.xml`（Dublin Core + `cp:`）与 `docProps/app.xml`（`ExtendedProperties`）
不只是元数据——`NUMPAGES` / `NUMWORDS` / `AUTHOR` 域的**取值来源**就在这里。缺了它们，
域更新后会变空。

要点：

- 两个部件都要在 `_rels/.rels` 里登记关系（`core-properties` / `extended-properties`），
  并在 `[Content_Types].xml` 里给 Override。
- 子元素顺序固定（见 `element_order.md` 的 `CT_Settings` 之后相关条目 / 直接查
  `pro/schemas/opc-coreProperties.xsd` 与 `shared-documentPropertiesExtended.xsd`）：
  `core.xml` 是 `dc:title → dc:subject → dc:creator → cp:keywords → dc:description →
  cp:lastModifiedBy → cp:revision → cp:lastPrinted → dcterms:created → dcterms:modified →
  cp:category → cp:contentStatus`；`app.xml` 里 `Application` 之后才是 `Words` /
  `Paragraphs` 等统计项，`TotalTime` / `Pages` / `Characters*` 各有槽位。
- 字数口径：中文按**字**计、西文按**词**计（对齐 Word 的中文计数），
  `extend` 侧由 `countBodyStats()` 统计并同时写入 `Words` / `Characters` /
  `CharactersWithSpaces` / `Paragraphs`。
- 用 `revision` 递增而不是时间戳，避免每次生成都产生无意义的 diff。

## 4. 扩展点速查（按落点分类）

### 4.1 段落级 / 运行级

| 能力 | 元素 | 归属 |
|---|---|---|
| 域（PAGE / TOC / SEQ / REF / STYLEREF…） | `w:fldSimple` 或 `w:r/w:fldChar + w:instrText` | `w:p` 的块级序列 / 运行序列 |
| 脚注 / 尾注引用 | `w:footnoteReference` / `w:endnoteReference` | `w:r` 内 |
| 书签（跨引用锚点） | `w:bookmarkStart` / `w:bookmarkEnd` | `w:p` 的块级序列 |
| 内容控件（SDT） | `w:sdt`（块级）/ `w:sdt`（运行级，在 `w:p` 内） | 见 `element-order.md` 的 `CT_SdtBlock` |
| VML 水印 | `w:pict/v:shape/v:textpath` | `w:r` 内 |
| 图片 / 形状 / 画布 / 组合 | `w:drawing/wp:inline|wp:anchor`（`wps` / `wpc` / `wpg`） | `w:r` 内 |
| 图表 | `w:drawing` → `a:graphicData` 里的 `c:chart r:id` | `w:r` 内，图表部件在 `word/charts/` |
| OLE 对象 | `w:object` + `v:shape` + `o:OLEObject`，配 `word/embeddings/*.bin` | `w:r` 内 |
| 符号 | `w:sym w:font w:char` | `w:r` 内 |

### 4.2 部件级

| 部件 | 用途 | 关系类型后缀 |
|---|---|---|
| `word/settings.xml` | 文档级开关（`updateFields` / `trackRevisions` / `evenAndOddHeaders`…） | `settings` |
| `word/footnotes.xml` / `endnotes.xml` | 脚注 / 尾注内容（分隔符 id 固定为 -1 与 0） | `footnotes` / `endnotes` |
| `word/glossary/document.xml` | 构建基块（自动图文集） | `glossaryDocument` |
| `word/comments*.xml` | 批注（现代多部件形态：`comments` + `commentsExtended` + `commentsIds` + `people`） | `comments` 等 |
| `customXml/itemN.xml` + `itemPropsN.xml` | 自定义 XML 数据岛 | `customXml` / `customXmlProps` |
| `word/charts/chartN.xml` | 图表定义 | `chart` |
| `word/embeddings/*.xlsx` / `*.bin` | 图表数据源 / OLE 复合文档 | `package` / `oleObject` |
| `word/theme/theme1.xml` | 主题（字体 + 配色 + 效果） | `theme` |
| `word/headerN.xml` / `footerN.xml` | 页眉 / 页脚 | `header` / `footer` |

### 4.3 `mc:AlternateContent`（兼容回退）

同一个位置想给新版与老版 Word 两套渲染时用：

```xml
<mc:AlternateContent>
  <mc:Choice Requires="wps">  <!-- 新式：DrawingML 形状 -->
    ...
  </mc:Choice>
  <mc:Fallback>               <!-- 老式：VML 回退 -->
    ...
  </mc:Fallback>
</mc:AlternateContent>
```

`Requires` 里的前缀**必须**已在根元素声明，否则 `mc:AlternateContent` 的约束会被
内容校验器判为无效。

## 5. 命名空间 → 能力对照

| 前缀 | 命名空间 | 承载什么 | 可否进 `mc:Ignorable` |
|---|---|---|---|
| `w` | `.../wordprocessingml/2006/main` | 正文主命名空间 | 否 |
| `r` | `.../officeDocument/2006/relationships` | 关系引用 | 否 |
| `wp` / `wp14` | `.../drawingml/2006/wordprocessingDrawing` / `.../2010/wordprocessingDrawing` | 图片定位（inline/anchor） | `wp14` 可，`wp` 不可 |
| `wps` | `.../2010/wordprocessingShape` | 文本框 / 形状本体 | **否（渲染内容）** |
| `wpg` | `.../2010/wordprocessingGroup` | 图形组合 | **否（渲染内容）** |
| `wpc` | `.../2010/wordprocessingCanvas` | 画布 | **否（渲染内容）** |
| `w14` / `w15` | `.../2010/wordml` / `.../2012/wordml` | 增强标记（`w14:paraId`、`docId` 等） | 可 |
| `v` / `o` | `urn:schemas-microsoft-com:vml` / `...:office:office` | 兼容层：水印、OLE 外壳、旧图形 | 视内容而定，水印不可 |
| `mc` | `.../markup-compatibility/2006` | 兼容性声明本身 | 否 |
| `c` / `a` | `.../drawingml/2006/chart` / `.../drawingml/2006/main` | 图表 / DrawingML | 否 |
| `m` / `wps` | `.../officeDocument/2006/math` | OMML 公式 | 否 |
| `dc` / `dcterms` / `cp` | Dublin Core / `.../package/2006/metadata/core-properties` | `docProps/core.xml` | 否 |

## 6. 本次实测踩过的坑（真实案例，均已修）

1. **`w:gallery` 放错层**：只能作为 `w:category` 的子元素，不能直接挂在 `w:docPartPr`
   下；且取值受 `CT_DocPartGallery` 枚举约束，`"gallery"` 不是合法值 —— 词库类
   构建基块写 `autoTxt`。（属于「值域」错误，和顺序错误一样是硬失败。）
2. **`c:dLbls` 里并存 `c:txPr`**：`showVal` 系列与 `txPr` 同处一个
   `maxOccurs=1` 的 choice，二者只能取一个。数据标签字号要改 `c:catAx` / `c:valAx`
   的 `c:txPr`（位置：`tickLblPos` 之后、`crossAx` 之前）。
3. **`mc:Ignorable` 混入 `wps wpg wpc`**：图形可能被整块丢弃。
4. **`customXml/itemN.xml.rels` 的 Target 基准错**：写成了相对 rels 文件自身，
   校验器报 `Broken reference` + `Unreferenced file`。

4. **`w:object`（OLE）的对象模型陷阱**：ISO 的 `CT_Object` 只接受
   `vml* / office* / drawing? / (control | objectLink | objectEmbed | movie)`，
   **没有 `w:r`**。Word 自己写的「`w:r` 包 EMBED 域 + `w:object`」形式依赖微软扩展 schema，
   按 ISO 校验会失败。跨版本稳妥的写法是 `w:pict/v:shape`（预览图）+
   `v:shape/o:OLEObject`（真实对象），即 `capabilities.py` 的 `ole_object_paragraph()`。
   若在 Word 里双击无法激活，第一处要回头检查的就是这里。
5. **标签之间的全角空格会要命**：XML 里元素之间只允许 tab / LF / CR / space。
   手写片段时若在标签之间留了全角空格（U+3000），它会被解析成父元素的**文本内容**，
   XSD 立刻报 `element-only` 错误。需要空格时把它放进独立 run 的
   `<w:t xml:space="preserve"> </w:t>` 里，不要留在标签外面。

## 7. 怎么用已封装的构件库

不必从零手写这些元素。`scripts/extend/` 把上面每一种扩展点都做成了可复用构件：

```
python scripts/extend/extend.py <config.json>
```

config 里一个操作一段，覆盖 `sectionBreak` / `toc` / `field` / `footnote` /
`endnote` / `watermark` / `pageProperty` / `pageBackground` / `drawing` / `canvas` /
`image` / `chart` / `ole` / `customXml`。构件库里已经处理好根声明、内容类型、
关系登记与子元素顺序——只有构件库**没有**覆盖的新扩展点，才需要回到这份文档手工下手。
