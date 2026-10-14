# extend：OOXML 扩展通道

`extend` 是 `apply` 之外的第二个写入器，专门负责**引擎刻意不写的那些特性**：
分节符、命名书签、域驱动的清单（目录 / 图表目录 / 索引）、脚注尾注部件、水印、页面装饰、
图形画布、图表、嵌入对象、自定义 XML、索引条目（XE）、内容控件（SDT）、
原生引文与书目（sources.xml）、首字下沉、嵌套表格、自定义文档属性。

分工是硬的，不要混用：

| 需求 | 用谁 |
|---|---|
| 样式、编号、页面设置、页眉页脚、按位置改段落 / 表格单元格、插图 | **`apply`**（Node 引擎） |
| 目录 / 图表目录 / 索引（含 XE 条目）、命名书签、脚注尾注、水印、页面边框 / 背景、形状 / 连接线 / 画布 / 艺术字、图表、OLE、自定义 XML、SDT 内容控件、原生引文与书目、首字下沉、嵌套表格、自定义文档属性、任意 XML 片段定位插入 | **`extend`**（本 CLI） |

`extend` 是**后处理**：先 `apply` 得到一份结构正确的文档，再 `extend` 往上面挂扩展内容。

## 运行

```bash
python scripts/extend/extend.py config.json
python scripts/extend/extend.py config.json --dry-run   # 跑完不落盘
```

依赖：Python + `lxml`（本机已装在 `envs/default`）。命令用绝对路径调用该解释器：

```bash
"C:/Users/PC/.workbuddy/binaries/python/envs/default/Scripts/python.exe" \
  scripts/extend/extend.py config.json
```

写完**必须**跑两套校验器，缺一不可：

```bash
node scripts/validate.js out.docx
python pro/scripts/validate.py out.docx
```

## Config 形状

```json
{
  "source": "in.docx",
  "output": "out.docx",
  "operations": [
    { "op": "bookmark", "at": { "afterText": "第二章" }, "name": "chap2" },
    { "op": "watermark", "header": 1, "text": "内部资料", "color": "C8C8C8" },
    { "op": "toc", "at": { "paragraphIndex": 1 }, "levels": "1-3" },
    { "op": "footnote", "at": { "text": "脚注" }, "text": "脚注正文" },
    { "op": "customXml", "fields": { "项目名": "示例", "阶段": "回归" } }
  ]
}
```

`operations` 按数组顺序执行；每个操作都是「先解析锚点 → 造片段 → 按 XSD 顺序插入 →
登记部件 / 关系 / 内容类型」。

### 锚点（`at`）

| 写法 | 含义 |
|---|---|
| `{"paragraphIndex": 12}` | 正文第 12 个顶层段落，**0 基**；负数从尾部数（`-1` 是最后一段） |
| `{"bookmark": "锚点名"}` | 含该 `w:bookmarkStart` 的段落 |
| `{"afterText": "第二章"}` | 文本**包含**该子串的段落 |
| `{"beforeText": "..."}` / `{"text": "..."}` | 同上（`text` 是 `afterText` 的简写） |
| `{..., "nth": 1}` | 子串命中多段时取第 2 个（0 基）；不写就取第 1 个，命中数为 0 会直接报错 |

锚点全部按**操作执行前**的文档状态解析，与 `apply` 的 `edits[]` 契约一致。

## 操作参考

### 定位类

| op | 必填 | 可选 | 产出 |
|---|---|---|---|
| `sectionBreak` | `at` | `kind`（`nextPage` / `continuous` / `evenPage` / `oddPage`）、`pgSz`、`pgMar`、`pgNum: [fmt, start]`、`columns: [num, space, sep]`、`textDirection`、`inheritHeaders`（默认 `true`） | 段落内嵌 `sectPr`，结束当前节并按声明开新节。**页眉/页脚引用、`titlePg`、`pgSz`/`pgMar`、`cols`、`docGrid` 默认从参考节继承**——`pgSz` 与 `pgMar` 各自独立继承，只声明其一不会让另一个掉回 Word 默认值；真要断开继承（如信封节不要页眉）写 `inheritHeaders: false` |
| `bookmark` | `name` +（`at` 或 `from`/`to`） | `nth`、`mark: "point"` | 给段落或段落区间加**命名书签**（`w:bookmarkStart` / `w:bookmarkEnd` 包裹），`w:id` 自动取未占用的最小值。默认包裹锚点整段；写 `from` + `to` 则跨段包裹；`mark: "point"` 插零长度书签（只标位置）。这是 `refTo`、`{ "link": "#名字" }`、`PAGEREF`、`TOC \b 名字` 的**落点**——先建书签，再让别的东西引用它。名字须合法且唯一（字母 / 中文 / 数字 / 下划线，不以数字开头，≤40 字符）；重名直接报错，因为已有引用会变歧义 |
| `toc` | `at` | `levels`（默认 `"1-3"`）、`hyperlinks`、`title`、`style`、`where`；`identifier` + `label` → 图表目录；`index: true` + `columns` → 索引 | TOC / TOC \o / INDEX 域 |
| `field` | `at` + `instruction` | `placeholder`、`inline`、`where` | 任意域的 `fldChar` 序列（PAGE / NUMPAGES / REF / STYLEREF / DATE…）。`inline: true` 把域运行**追加进锚点段落内部**（邮件合并行内域、`MERGEFIELD` 用这个），不写则单独成段 |
| `drawing` / `canvas` / `image` | `at` | 见下 | `w:drawing` 图形 |

### 内容类

| op | 必填 | 可选 | 产出 |
|---|---|---|---|
| `footnote` / `endnote` | `at`、`text` | `kind` | 引用运行 + `footnotes.xml` / `endnotes.xml` 条目（自动补齐 id 与分隔符） |
| `watermark` | — | `header`（默认 1）、`text`、`color`、`font`、`rotation`（默认 315）、`opacity` | 页眉里的 VML 水印（`w:pict` + `v:textpath`） |
| `pageProperty` | — | `sections`（`"all"` / `"1"` / `"1,3"`）、`borders`、`pgNum: [fmt, start]`、`columns`、`textDirection` | 就地更新目标节的 `sectPr` 子元素（同名替换，不追加） |
| `pageBackground` | `color` | `drawing: true` | `w:background` 或背景图形 |
| `customXml` | — | `fields: {}`、`root`（默认 `project`） | `customXml/itemN.xml` + `itemPropsN.xml` + 关系与内容类型 |
| `xe` | `term` | `at`（不填则追加到最后一段） | 隐藏 `XE "term"` 域标记（`w:vanish`），配合 `index: true` 的 INDEX 域，F9 生成真索引 |
| `sdt` | `text` | `alias`、`tag`、`style`（默认 `BodyText`）、`at`、`dropdown: [{displayText, value}]`（写了就出下拉列表，`text` 作为当前值） | 块级富文本内容控件（`w:sdt` + `sdtPr` 的 alias/tag/id，可选 `w:dropDownList`）。**必须最后插**——SDT 会阻塞 `apply` 的编辑阶段 |
| `source` | `tag`、`type`、`fields`、`authors`、`cite: {at, text}` | — | 原生引文/书目：登记 `word/sources.xml`（shared-bibliography.xsd）+ 就地插 `CITATION tag \m` 域。配合 `field` op 插 `BIBLIOGRAPHY` 域，F9 生成完整书目。字段名按 XSD：`Issue`（非 Number）、`StateProvince`、`CountryRegion`；人名 `{"last": "张", "first": "三"}`（字符串 ≤4 字自动按中文姓名拆分，更长视为机构名） |

`source` 的 `type` 取 ST_SourceType 枚举：`Book` / `JournalArticle` / `ArticleInAPeriodical` / `BookSection` / `ConferenceProceedings` / `Report` / `Patent` / `Case` / `Film` / `Interview` / `InternetSite` / `DocumentFromInternetSite` 等 19 种。`fields` 键为 `b:` 元素局部名（`Title` / `Year` / `Publisher` / `City` / `JournalName` / `Pages` / `Volume` / `URL` …），白名单外的键直接报错。

| `dropCap` | `at` | `lines`（默认 2）、`kind`（`drop` / `margin`）、`size`（首字字号，如 `"21pt"`） | 首字下沉：段落 `pPr` 写入 `framePr`，并把首字符拆成放大的独立 run（默认 21pt / 42 half-points） |
| `nestedTable` | `at: {cellText}`（可加 `nth`） | `rows`（二维数组，必填）、`widths`（每列 twips）、`borderColor`、`headerFill`、`headerBold` | 向匹配文本的**表格单元格**内嵌子表。插在单元格最后一个段落之前，保证 `w:tc` 以 `w:p` 结尾的规范不变量。注意 `at` 用 `cellText`（搜全部单元格），不是普通段落锚点 |
| `customProps` | `properties: {}` | — | `docProps/custom.xml` 自定义文档属性（文件 → 信息 → 属性 → 高级属性可见）。值为 str/bool/int/float 自动映射 `vt:lpwstr/bool/i4/r8`；登记根 rels 与 content type，重复运行 upsert 幂等 |
| `glossary` | `entries: [{name, paragraphs\|xml}]` | 每条目 `gallery`（默认 `autoTxt`）、`category`（默认 `General`）、`autoInsert`（打字自动插入） | `word/glossary/document.xml` 构建基块部件（自动图文集）。条目出现在 Word「插入 → 部件 → 构建基块管理器」对应 gallery 下。`paragraphs` 项为纯文本或 `{text, style}`；重复运行追加条目。gallery 取值：`autoTxt` / `coverPg` / `tbls` / `pgNum` / `hdrs` / `ftrs` / `watermarks` / `txtBox` / `docParts` 等 |
| `cover` | `title` | `subtitle: []`、`titleStyle`（默认 `Title`）、`subtitleStyle`（默认 `Subtitle`）、`firstPageDifferent`（默认 true） | 封面页：把标题 + 副标题段落插到正文最前，并追加一个 `vAlign=center` + `titlePg` 的分节符——封面节垂直居中、首页页眉页脚自动留空（未声明首页页眉时）。样式需已存在（apply 默认样式集含 Title / Subtitle） |

`borders` 的键直接透传给 `pg_borders()`：`val` / `sz` / `space` / `color` / `offset_from`。

### 图形类字段

| op | 关键字段 |
|---|---|
| `drawing` | `shape`（预设几何名，默认 `roundRect`）、`text`、`fill`、`line`、`warp`（艺术字变形）、`width` / `height`（厘米）、`anchor`（`t`/`ctr`/`b`）、`insets: [l,t,r,b]`（EMU）、`paragraphs: [{text,size,bold,italic,color,font,align}]`（多段落富文本）、`name`；`symbols` → 符号段落、`ink` → 手写墨迹；`connector: true` → 直线连接线 `straightConnector1`（配合 `flipH` / `flipV` / `line` / `arrow: true`）；**`textBox: true` → 真文本框**（`wps:cNvSpPr txBox="1"`，默认无填充/无边框，配 `paragraphs` 多段落、`insets` 内边距、`anchor` 垂直锚点、`wrapText` 自动换行、`rotation`（度）） |
| `canvas` | `children: [{"text": "...", "x": 0.2, "y": 0.9, "width": 3.0, "height": 1.3, "shape": "roundRect", "fill": "...", "textColor": "...", "size": 18}]`（**`x`/`y` 为画布内厘米坐标**，缺省即左上角）、画布 `width` / `height`、`name` |
| `smartArt` | `kind`（`process` 水平流程 / `cycle` 环形 / `list` 垂直列表 / `hierarchy` 树形）、`nodes: [str \| {text}]`、`width` / `height`、`fill` / `line` / `textColor` / `font` / `size` / `arrow`、`name`。自动计算节点尺寸与连线并置于 `wpc` 画布内（可整体缩放/移动，文字可编辑） |
| `image` | `part`（包内已有部件，如 `word/media/image1.png`）**或** `file`（本地图片路径，自动写入 `word/media/` 并登记内容类型）、`rid`（不填则自动登记关系）、`crop: [l,t,r,b]`、`lum: [bright, contrast]`、`effects`（`a:effectLst` 内容，可含外层包装；子元素会按 `blur → fillOverlay → glow → innerShdw → outerShdw → prstShdw → reflection → softEdge` 的 XSD 顺序自动重排）、`floating: true` + `wrapText`（`square` 四周型 / `tight` 默认 / `bothSides` / `left` / `right` / `largest`） |
| `chart` | `rows`（工作表数据）、`categories`、`series`、`title`、`kind`（默认 `bar`）、`name`、`width` / `height` |
| `ole` | `rows`（载荷工作簿）、`stream`（默认 `Package`）、`progId`、`name`、`previewPart`（包内已有媒体部件）**或** `previewFile`（本地图片，自动入库）；两者都缺或指向不存在的部件会直接报错，不会留下悬空关系 |

## 它替你处理了什么

这些是手写 OOXML 最容易翻车的地方，`extend` 已内置：

1. **根命名空间补齐** —— 写回前对每个 `word/*.xml` 补 `xmlns:wp/wps/wpg/wpc/w14/w15/mc…`，
   并按「渲染型前缀不进 `mc:Ignorable`」的规则整理 `mc:Ignorable`。详见
   [`ooxml-extension-points.md`](ooxml-extension-points.md) §1。
2. **元素顺序** —— 所有插入都按 XSD 顺序落位（`insert_ordered` + `SECTPR_ORDER` /
   `PPR_ORDER`）。顺序表和自查方式见 [`element-order.md`](element-order.md)。
3. **部件注册四件套** —— 内容类型 Override/Default、关系（Target 按**源部件目录**解析）、
   关系部件、正文引用，一次做全。
4. **同名去重** —— `pageProperty` 走 upsert（`pgBorders` / `pgNumType` / `cols` /
   `textDirection` 同名先删后插），不会叠出两个。
5. **幂等可重跑** —— `ensure_root_declarations` 幂等；但业务型操作（域、脚注、图形）
   是**追加**语义，同一份 config 不要对同一文档跑两遍。

## 全角 / 半角与文本细节

- 域结果、脚注正文等文本里**不要**在 CJK 与拉丁 / 数字交界处手打空格，交给 Word 的
  autoSpace 处理（`apply` 的 Pangu 检查同样适用于这里写入的文本）。
- 需要保留首尾空格时用 `xml:space="preserve"`（构件库已处理）。
- 水印的 `size_pt` 故意保持极小：VML `textpath` 的字号由形状缩放决定，直接写大字号
  会溢出形状。

## 边界

- 只处理**顶层正文段落**的锚点；表格单元格里的段落不做锚点匹配（要精确定位请先用
  `apply` 的 `cell` 定位改写，或先导出结构再指定 `paragraphIndex`）。
- 不重建已有 TOC 的内容 —— 写入的是域，页码需在 Word 里 Ctrl+A → F9 更新。
- 图表是静态缓存 + 内嵌工作簿：改数据要重新跑 `chart`，或让用户在 Word 里编辑数据源。
- **SmartArt（dgm 语义对象）不生成，但 `smartArt` op 提供等效图示**：真 SmartArt 需
  data / layout / quickStyle / colors 四个部件，其中 layout 是 Word 内置布局算法定义
  （`dgm:layoutDef` 的 forEach / alg / constrLst 指令），无法离线可靠构造，硬拼会让
  Word 弹修复提示。`smartArt` op 改为用 `wpc` 画布 + 预设几何形状 + `straightConnector1`
  连线自动布局（process / cycle / list / hierarchy），节点文字可编辑、整组可缩放移动，
  视觉与交互达到 SmartArt 的等效水平；差别只在于它不是 Word 的「SmartArt 对象」
  （没有版式切换下拉、无法用 SmartArt 工具增删节点）。
- **ActiveX 控件不支持**：控件属性是二进制 OLE 流（`word/activeX/activeX*.bin`），
  格式无公开完整规范，静态构造不可靠。交互控件请用 `sdt` 内容控件替代。
- **真 InkML 手写笔迹不支持**：InkML 部件 + `w14:contentPart` 引用的注册方式缺少可验证
  的 XSD 覆盖，写入失败会导致文档无法打开；改用 VML 折线（`drawing` 的 `ink`）呈现
  笔迹外观，全版本可渲染，代价是没有笔顺回放与压感。
