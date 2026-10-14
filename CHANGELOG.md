# Changelog — docx-master

只记录**能力与行为变更**：新增字段 / 修掉的缺陷 / 明确的边界。文档润色不记。
每条都可在 `apply.js`、`scripts/extend/*.py` 里检索到对应实现。

---

## 2026-10-14（第十三轮）—— run / 段落级格式扩展 + 拼音注音正式纳入引擎

用户对照 Word 字体 / 段落对话框逐项盘点后的能力补齐。

### `runFormat` 新增 17 字段（`buildRPrChildren` 全部落地，顺序由 `RPR_CHILD_ORDER` 排序保证）

- `highlight` — 荧光笔，`w:highlight` 16 色 + `none`
- `characterScale` — 字符缩放，数字 1–600（150 = 150%）；**线上格式是 `ST_TextScale`，
  必须写带 `%` 的字符串**（`"150%"`）——裸整数 union 分支会被 XSD 拒绝，实测踩过
- `position` — 字符位置提升/降低（Length，可负；`w:position` 半磅）
- `shadow` / `outlineText` / `emboss` / `imprint` / `smallCaps` / `allCaps` / `hidden` — 字体效果七开关
  （`hidden` → `w:vanish`；`outlineText` → `w:outline`，避开与段落 `outlineLevel` 撞名）
- `kern` — 为字体调整字间距的阈值（Length，半磅；Word 默认 `"1pt"`）
- `fontCs` / `sizeCs` — 复杂文种字体 / 字号（`w:rFonts@w:cs` / 独立 `w:szCs`）
- `underlineStyle`（18 种线型）+ `underlineColor` — 下划线从布尔扩展；style 优先，`underline: false` 仍显式关
- `shading` — 字符底纹（`w:shd`，fill 必填）
- `border` — 字符边框（`w:bdr`，size 八分之一磅 / space 磅）
- `RPR_MANAGED_LOCAL_NAMES` 同步扩容（`format` op 清理时认得新元素）

### `paraFormat` 新增 14 字段（`buildPPrChildren` 重写 push 序列）

- 换行和分页：`keepNext` / `keepLines` / `pageBreakBefore` / `widowControl`
- 中文版式：`kinsoku`（避头尾）/ `wordWrap`（`false` = 允许西文断词）/ `overflowPunct` /
  `topLinePunct` / `autoSpaceDE` / `autoSpaceDN`
- `snapToGrid`（与文档网格对齐）、`textAlignment`（文本对齐方式 5 值）、`shading`（段落底纹，含 pattern）
- **`buildPPrChildren` 没有 sort，push 序列即 CT_PPrBase 顺序契约**——新元素按
  keepNext→…→shd→kinsoku→…→autoSpaceDN→bidi→snapToGrid→spacing→ind→jc→textAlignment→outlineLvl 插入

### 拼音注音（`w:ruby`）正式纳入引擎

- `text` 数组新增 inline piece：shorthand `{ ruby: "拼音注音", pinyin: ["pīn","yīn","zhù","yīn"] }`
  （逐字对应，长度必须一致，`""` 跳过该字）或 pair 形式
  `{ ruby: [{ base: "行", pinyin: "xíng" }, { base: "、" }] }`（多音字精确控制）
- 可选 `align` / `hps` / `raise` / `baseSize`；`CT_RubyPr` 五个子元素全部 required，emitter 填默认值
  （拼音字号=基文一半、`hpsRaise`=基文字号、`lid=zh-CN`）；基文 run 继承 inline format
- **P2-20 修复**：`find_text.js` 投影与 `_shared/locator.js` 的 `paragraphText()` 现在把
  `w:rubyBase` 里的基文计入正文、`w:rt` 的拼音排除——带拼音的段落重新可查找、可断言、可锚定；
  `verify.js` 的 `textOf()` 先剥 `<w:rt>` 再跑 token 正则（线性扫描里拼音与基文交错，不剥必假阴）
- 实测：`find_text "拼音注音"` → 1 match（cross-run）；`verify contains` PASS；
  多音字两个独立 ruby（行 xíng / 行 háng）连续可匹配

### 回归

`node tests/run.mjs --smoke` → PASS 95/95；能力验证文档 28 项新元素逐项 XML 核验全部命中，
Node `validate.js` PASS，PRO `validate.py`（XSD + OPC）PASS。

---

## 2026-10-14（第十四轮）—— 缺陷修复：批注锚点丢失 / extend 自闭合 `<w:t/>` / 批注断言缺位

对照第十三轮的效果文档回跑主交付链时暴露出的三个缺陷。

### 🔴 P1-13 `replace` / `delete` 静默丢弃批注锚点

- **现象**：对含批注的段落做 `replace` 或 `delete`，正文变了但 `commentRangeStart` /
  `commentRangeEnd` / `commentReference` 三个锚点随被删段落一起消失 → `word/comments.xml`
  里留下**永不显示的孤儿批注**。实测修复前 `start/end/ref = 3/3/3`，`comments.xml` 定义 4 条 → 孤儿 `[1]`。
- **修复**：新增 `commentAnchorsIn()` / `nearestParagraphSibling()` / `migrateCommentAnchors()`。
  删除或替换前先把锚点**搬到相邻段落**（先找 previous，无则 next）：
  `commentRangeStart` 插到目标段 `pPr` 之后（保持 start→end 的先后次序），
  `commentRangeEnd` + 含 `commentReference` 的 run 追加到段末。
  当落地段是「邻居」而非替换段时（替换内容为空），全部锚点按段落边界追加（`atEnd = true`）。
- **兜底**：若无任何可搬迁的邻居段落，`replace` 直接抛错而不是继续丢锚点。
- **可观测**：迁移发生时 stderr 打 `[apply] warning: replace — N comment anchor(s) moved ...`。
- **实测**：`replace` 与 `delete` 两条路径均回到 `start/end/ref = 4/4/4`、孤儿 `[]`。

### 🔴 P1-14 `extend` 序列化未展开自闭合 `<w:t/>`

- **现象**：Python 通道写出的 part 里残留 `<w:t/>` / `<w:t xml:space="preserve"/>`。
  这类空文本节点在 `verify.js` 的行式 token 正则里会**整段吞掉**，导致断言假阴；
  下游按 token 定位的工具也会错位。实测 `build/stage2-extended.docx` 里 12 处。
- **修复**：`_common.py` 新增 `expand_self_closed_text()`，在 `DocxPackage.save()` 的**包边界**
  统一跑一遍——覆盖所有 op 写入，并连带覆盖后续新造的 part。
  快速检测（`b"<w:t/" in data`）会漏掉带属性的形态，改为**无条件跑正则**。
- **实测**：`stage2-extended.docx` 自闭合 `<w:t/>` 12 → 0；主文档规范化阶段同样从 12 → 0。

### 🔴 P3-19 `verify` 没有批注断言

- **现象**：批注部件存在性有断言（`partExists`），但「有没有孤儿批注」「锚点全不全」无断言，
  P1-13 这类缺陷在交付门上是**看不见的**。
- **修复**：`verify.js` 新增批注扫描与两种断言：
  - `{ "type": "commentCount", "eq": 4 }` — `comments.xml` 里 `w:comment` 的定义数
  - `{ "type": "commentsIntact" }` — 每条批注都必须**同时**被锚定（`commentRangeStart`）
    且被引用（`commentReference`），且无悬空引用、无未闭合区间；零容差
- **自证**：故意剥离 id=1 锚点造出 `build/orphan-comment.docx` 对照样本 →
  `[FAIL] commentsIntact ... 1 orphan comment(s) [1]`，断言确实会 fire。
- 主交付 `config/verify.json` 已加入这两条，`commentCount = 4` PASS、
  `commentsIntact` → `4 comment(s), all anchored and referenced` PASS。

### 回归

`node tests/run.mjs --smoke` → **PASS 95/95**；主交付文档断言 **64/64 passed**
（含新增两条批注断言）；Node `validate.js` PASS；PRO `verify_elements.py` 落地 76 项。

---

## 2026-10-13（第十轮）—— 新建 `<w:pgMar/>` 漏写 XSD 必填属性

- **触发**：给一个**没有 pgMar** 的节声明页边距——`pageSetup.sections."1".margins`，封面节是典型场景
  （封面节由 `cover` op 建出，sectPr 里只有 `type` / `vAlign` / `titlePg`）。
- **根因**：`mutatePgMar()` 新建 `<w:pgMar/>` 后只写入配置声明的字段，而 **`CT_PageMar` 的七个属性
  （top / right / bottom / left / header / footer / gutter）全部是 required**。声明了六个、漏掉 `gutter`，
  文档就在它刚创建的那个节点上失败：
  `w:pgMar: The attribute 'w:gutter' is required but missing`。
- **修法**：新增 `seedPgMarAttrs()`——新建 pgMar 时先把七个属性补齐：优先从文档里已有的 `pgMar` 继承
  （新节沿用文档自身的页面几何），文档里没有 donor 时才用 Word 默认值
  （1440 / 1800 / 1440 / 1800 / 851 / 992 / 0 twips）。
- **实测**：同一份配置（**未**声明 `gutter`）修复前 `Validation FAILED — 2 new error(s)`，修复后
  `Validation: PASS`；成品的 11 个节全部 7/7 必填齐全，新建节继承到文档自身的 `gutter=567` twips。
- **同类排查**：`mutatePgSz()` 不受影响——它新建后立刻写 `w:w` / `w:h`（两者均为 required），
  `orient` 本身可选。

---

## 2026-10-13（第九轮）—— 缺陷修复：交付门误报、XSD 漏网、静默失效

这一轮的每一条都来自一次**全能力演示文档**的实测（86 项元素逐项落地 + 双校验 + 43 条断言），
不是代码审读猜出来的。按「回归测试跑得出来的才叫修好」标准，每条都附实测证据。

### P0 —— 交付门误报：`verify` 的文本提取在自闭合 `<w:t/>` 上崩掉

- **根因**：`TOKEN_RE` 第一个分支 `<w:t(?:\s[^>]*)?>([\s\S]*?)</w:t>` 的属性类 `[^>]*` 会吞掉自闭合写法
  `<w:t xml:space="preserve"/>` 结尾的 `/`，把它读成开标签，于是「一路匹配到下一个 `</w:t>`」，
  中间整段正文被当成域结果丢弃。实测：77136 字符的 `document.xml` 只提取出 251 字符，第 5 个域之后正文全丢。
- **触发条件**：文档含自闭合文本节点。而 `apply` 为 `captions[].chapterPrefix` 写隐藏章节计数器时正好产出这种节点
  ——**凡用了按章题注编号的文档都会中招**，属高频路径。
- **修法（两侧同时）**：
  1. `scripts/verify.js`：`TOKEN_RE` 增加自闭合分支并置于配对分支**之前**（`<w:t …/>` / `<w:instrText …/>`），
     `textOf()` 对自闭合节点既不出文本也不出换行。
  2. `scripts/_shared/xml-utils.js`：`serializeXml()` 统一把 `w:t` / `w:instrText` 的自闭合写法展开成显式配对
     （`expandSelfClosedTextNodes()`）。语义逐字节等价，从源头消除隐患——所有部件都受益。
- **实测**：`apply` 输出中自闭合节点 **6 → 0**；把 1 处自闭合节点注入成品文档后，
  新版 `verify` **43/43 通过**，旧版同一文档 **2 条断言失败**（对照实验）。
  外部规避脚本 `normalize_empty_text.py` 从此不再需要。

### P0 —— Node 校验器放过了 XSD 级错误

同一份文档：`node scripts/validate.js` 全 PASS，PRO `validate.py` 报 9 个 XSD 错误。两处都是真缺陷：

1. **`<w:sym @w:char>` 必须是 `ST_ShortHexNumber`**，传实际字符（`→`）直接非法。
   修法：新增 `symbol_char_code()`，接受**字符本身**（`"→"`）、`U+2192`、`0x2192`、十六进制 `"2192"`
   或十进制码位（整数 `8594`），统一规范化；超出 BMP 或多余字符给出可读报错。
   歧义规则：≤4 位的纯十六进制字符按码位解释（这正是 `ST_ShortHexNumber` 的形态，也是符号表的写法）。
2. **`<b:Corporate>` 不在 `CT_PersonType` 里**。它其实是 `CT_AuthorType` 的子元素——机构作者不该包在
   `<b:Person>` 内。原 `_person()` 对长于 4 字符的姓名一律输出 `b:Corporate`，**任何机构作者都会产出非法 XML**。
   修法：拆成 `is_corporate()` / `corporate_text()` / `_person()` / `author_block()`；人名进
   `<b:NameList><b:Person>`，机构名进 CT_AuthorType 直系 `<b:Corporate>`，两者可混写。
   实测输出：`<b:Author><b:Author><b:NameList><b:Person><b:Last>张</b:Last>…</b:Person></b:NameList><b:Corporate>ISO/IEC JTC 1</b:Corporate></b:Author></b:Author>`，
   PRO 校验 PASS。

### P1 —— 文档与实现字段名不一致（其中一条是静默失效）

- `source` op：文档写 `type` / `cite: {at, text}`，实现读 `sourceType` / `citeAt`。现实现**两者都收**，
  两个都写且不一致时直接报错（别名分歧是配置错误，不是偏好）。`cite.text` 用作域占位文本。
- `drawing` 连接线：文档写 `flipV` / `flipH`，实现读 `flipVertical`。这是**静默失效**——不报错，形状就是不翻转，
  比报错更难发现。现四个拼写都认，并新增 `spec_flag()`：同一开关的多个拼写都出现且互相矛盾时报错。

### P1 —— 浮动图片定位硬编码，与宽 inline 图混排必然错乱

- **现象**：`floating: true` 的图片在 Word 里把同页其他图形挤得重叠、越出页边距。
- **根因（两层）**：① `picture_run()` 浮动分支把定位写死为 `column/0 + line/0 + allowOverlap=1`，
  配置里没有任何参数能改；② 环绕（`wrapSquare`）只对**文字行**生效，同页后续比环绕剩余宽度更宽的
  **嵌入式**图形会被挤向右侧、越界叠压。
- **修法**：`picture_run()` 暴露 `offset_x_cm` / `offset_y_cm` / `allow_overlap` / `behind_doc`，
  `image` op 透传 `offsetX` / `offsetY` / `allowOverlap` / `behindDoc`；
  `extend.md` 明确警告该组合风险，并建议版式关键图形一律走嵌入式。
- **实测**：修复前 1 个 `wp:anchor`(9×5.4cm) + 10 个 inline；改为嵌入式后 **anchor=0 / inline=11**，
  双校验与 43 条断言复跑全过。

### P1 —— `strictObject` union 的报错归因指向错误位置

`numbering` 里塞了非法键（`indentLeft`）时，报错是 `numbering: expected object, got [...]`——
因为 union 的两个分支里，对象分支只报 1 条笼统的浅层类型错，数组分支报 N 条精确错误，
而原策略是「**取问题最少的分支**」，恰好选中最没用的那个。

修法：改为按**具体性**打分（`unionVariantScore()`）——`unrecognized_keys` 说明形状已被理解、只错在键名，
得最高分；深层 path 加分；**根层 `invalid_type` 扣分**（意味着该变体根本不适用）。
实测：同一份配置现在报 `numbering[0].levels[0]: unknown key "indentLeft"`，精确指向真实位置。
同时 `config-schema.md` 写明：缩进只能落在绑定样式上，编号层级只管 `numFmt` / `lvlText` / `suff` / `numRPr`。

### P2 —— 锚点命中多处时静默取首个

按子串匹配的文本锚点命中多段时，原实现静默取第一个。「表格」这类中文短词实测命中 7 段——
落在标题、正文或生成的目录条目上都是可能的。现在命中多于一处时往 stderr 打一条 warning
（同一改法只提示一次，不阻断），并说明更好的做法：在 `apply` 阶段落命名书签，`extend` 用
`{"bookmark": "..."}` 精确指过去。`extend.md` 的锚点章节同步加上「优先用 bookmark 锚点」。

### P2 —— `doctor` 探不到已装好的 venv

原实现遇到第一个**能启动**的解释器就返回：PATH 上的 `python` 通常是没有 lxml 的裸解释器，
于是一台机器上明明躺着配好的 venv，报告仍写 "MISS Python + lxml"，把人引去装已经有的包。

修法：优先试托管 venv（含 `envs/default`），**找到带 lxml 的才收工**；否则记住可用的那个继续找，
最后在 detail 里列出发现的 venv 路径。实测：现在报
`OK Python + lxml — …\envs\default\Scripts\python.exe — Python 3.13.14`。

### 新增 `references/out-of-scope.md`

把「明确不做」的能力集中成一份可引用的清单，每项都写清**为什么不做**和**替代路径**：
宏（`.vbaProject.bin`）· ActiveX · 真 SmartArt（`dgm`）· InkML · `.dotx` 导出 ·
PDF/XPS 导出 · 翻译 / 听写 · 信封 / 标签 · 以及属于 Word 应用层而非文档层的（自动更正词条、
词典、双面打印、快速样式集）。SKILL.md 的 Out of scope 段给一行索引 + 指针，
避免每次都要重新论证一遍。主控体积 19981 / 20000 B。

### 更正

- 第八轮清单里「PRO `validate.py` 失败时退出码仍为 0」**不成立**：`validate.py` 早已是
  `sys.exit(0 if success else 1)`，`validate_against_xsd()` 在 `new_errors` 非空时返回 `False`。
  当时观测到 0 是管道（`| tail`）吞掉了上游退出码，属测量误差，非 skill 缺陷。

### 验证

- `node tests/run.mjs --smoke` → **PASS 95/95**
- 演示工程全链路回归：`apply → extend → 双校验 → verify` 43 条断言全过；书签 0 悬空
- 对照实验：同一份含自闭合节点的文档，旧 `verify` 2 条失败 / 新 `verify` 全过

---

## 2026-10-13（第八轮）—— skill 结构重构：主控瘦身，规则可测

这一轮不改能力，改**技能本身的可维护性**：让主控每次触发少烧 token，同时保证一条规则都不丢。

### 主控瘦身（SKILL.md 35413 B → 19908 B，↓43.8%）

原先 SKILL.md 单文件承载了主流程 + 硬规则 + 完整工具表 + 交付门细节 + 引擎行为说明，每次触发全量注入。
现按「必须当场知道 vs. 用到才查」重新分工：

- **留在主控**：触发条件、设计心智模型、目标态、主流程、硬规则（Agent 必须主动遵守的那些）、路由索引、侦察工具清单。
- **下移到子文件**：
  - `references/tools.md`（新）—— 19 个 CLI 的完整调用与参数、`batch_fill` 语义，外加 **输出裁剪**（哪些 `--paras` / `--limit` / `--problems-only` / `--json` / `--quiet` 参数能让侦察结果只送必要片段进上下文）。
  - `references/delivery-gate.md`（新）—— 交付门 5 步、断言类型清单、变更清单读法、最后一道人工自检。
  - `references/engine-behavior.md`（新）—— 引擎替你做的事：命名空间修复、`settings` 幂等、`docProps` 刷新、样式名预检、restyle 语义、基线差分校验。这类「引擎已经做了、你只需知道」的事实不再挤占主控。
- **消除重复**：书签「三扇门」细节、PRO 工具路由、`--allow-validation-warnings` 等原先在主控与子文件各有一份，现只保留**一处**定义，另一处指向。

### 新增：回归测试（`tests/`）

`node tests/run.mjs` 零依赖检查四类事实 —— 结构契约（frontmatter / 体积 / description 长度）、
**规则契约**（`tests/contract.json` 里每条「不许删的硬规则」是否还在其声明位置）、
资源契约（references / scripts 存在性）、链接完整性（相对链接与 `#锚点` 按 GitHub slug 解析）。

`--smoke` 额外跑 `tests/cases/*.json`（当前 1 例：最小 apply config 通过 dry-run）。加规则 / 加用例见 `tests/README.md`。

### 保真核验

重构按「旧主控的每个逻辑块都必须有新归宿」逐条比对，而非凭感觉压缩：旧文件引用的**全部脚本名、CLI 标志、反引号标识符**在新文件集中均能命中，差异只有措辞压缩与位置下移。核验中发现并补回一处真实丢失 —— `element_order.py` 的调用形式（`--type` / `--list` / `--doc`），已落到 [pro-toolchain.md §4.7](references/pro-toolchain.md)。`tests/contract.json` 把这批硬规则固化成断言，防止后续再被误删。

### 边界（本轮未动）

引擎行为、CLI 参数、config schema、产出质量均无变化；`scripts/` 下的脚本一个字节未改。

---

## 2026-10-13（第七轮）—— 书签与交叉引用：补上"看不见的对象"这一层

上一轮补的是"产出是否忠于用户原话"，这一轮补的是**同一性质空白的第二个实例**：书签没有任何渲染、
不占字数、不带样式，丢一个之后文件照样打开、照样过 XSD，只有 Word 会在某句话中间印出
"错误！未定义书签。"，或者目录条目点不动。整个工具链此前没有任何一处会去看它 ——
`inspect_*` 里没有书签工具，`verify` 的断言类型里也没有书签项。本轮四处补齐。

顺带核实了两件以前只能靠猜的事：**目录是域**（`TOC \o "1-3" \h \z`，结果为空，交给 Word 填），
而目录条目的"可点击跳转"来自 Word 更新域时**自动生成**的 `_Toc*` 书签 + PAGEREF + HYPERLINK 三元组
（实测一份 11 章文档生成 70 组）—— 不需要、也不应该手工埋这些书签。

### 新增能力

- **`scripts/_shared/bookmarks.js` —— 书签与引用分析核心（纯文本、零依赖）。**
  以 `fldChar` 状态机（带栈）扫描部件的 XML，因此**能正确处理域嵌套**：Word 更新后的目录里每个条目都是一个
  PAGEREF 域，嵌在 TOC 域内部，平坦正则会把这些指令全部算到外层域头上。识别 REF / PAGEREF / NOTEREF、
  `HYPERLINK \l`、`TOC \b`、结构化 `w:hyperlink w:anchor`（STYLEREF 引的是样式名，不算引用）。
  报告书签（名字 / id / 所在部件 / 段落序号 / 来源）、引用及其解析状态、以及六类问题：
  `dangling`（引用不存在的名字）、`unclosed` / `unopened`（书签区间没闭合 / 游离的 bookmarkEnd）、
  `duplicate-id` / `duplicate-name`、`empty-name` / `empty-target`。跨部件解析：页眉里的 REF 可以指向正文书签。
  来源分类把 Word 自产的 `_Toc*` / `_Ref*` / `_GoBack` / `_Hlt` / `_OLE_LINK` 与**自己写的**锚点区分开。
- **`scripts/inspect_bookmarks.js` —— 第 12 个侦察工具。**
  `node scripts/inspect_bookmarks.js <file> [--json] [--problems-only] [--exit-zero] [--limit N]`；
  输出书签清单、引用清单（按 kind + target 聚合，70 条 TOC 引用压成 70 行 ×2 而非 140 行）、问题清单；
  **发现断开引用或畸形书签即 exit 1**，可直接接进 CI / 交付门。`--exit-zero` 只报告不改退出码。
- **`scripts/_shared/opc.js` —— 共享 OPC 读取。**
  原先 `verify.js` 内联了一份 ZIP 读取；两份实现迟早会漂移，抽成共享模块（`openZip` +
  `referenceBearingParts` 部件枚举，含 rels 解析与常规命名兜底）。`verify.js` 改为引用它。

### 新增写入能力（书签此前只能"读"不能"写到位"）

- **`extend`：`{"op": "bookmark", "name": "chap1", ...}`。**
  `at` 包裹锚点整段；`from` + `to` 跨段包裹（这也是把 `TOC \b 名字` 限定到多段的唯一办法）；
  `mark: "point"` 插零长度书签。`w:id` 自动取未占用的最小值（同时扫描 bookmarkStart 与 bookmarkEnd，
  复用游离 end 的 id 会被 Word 重新编号）。
- **`apply`：`format` op 新增 `anchor` 字段。**
  此前给段落挂锚点只有 Block 级 `anchor` 一条路，而它只能覆盖**本次运行新建**的段落 ——
  文档里已经存在的段落永远拿不到锚点，也就永远不能被 REF / 跳转链接指向。现在
  `{ "op": "format", "at": {...}, "anchor": "name" }` 可以只打锚点（可作为唯一字段），
  多段目标直接报错（一个书签只能命名一个段落）。
- 硬约束（本轮实测确认）：`InlineRef` 的 `display: "number"` 要求目标绑定过 `numbering`
  （或属题注类目标），普通正文段落没有编号可引用，apply 拒绝而非写出渲染为空的域；改用 `display: "full"`。

### 交付门扩展

- **`verify` 新增六种断言**：`refsResolve`（所有引用都能解析，失败时逐条列出引用类型 + 名字 + 部件 + 段落）、
  `bookmarkProblems`（默认要求 0 个畸形书签；`unclosed-field` 降级为提示，因为域跨部件开闭合法）、
  `bookmarkExists` / `bookmarkMissing`（名字拼错时提示包里有多少书签、其中多少是自己写的）、
  `bookmarkCount` / `refCount`（`eq`/`ne`/`min`/`max`）。
- **变更清单新增**：`bookmarks` / `references` 增减行，以及**出现悬空引用时的显式警示行**
  （"a dead reference renders as text in Word"）。
- 默认报告新增一行：`bookmarks N · references N · dangling N`。

### 修掉的缺陷

- `w:id` 被复用时，先前那个尚未闭合的书签会从待闭合集合里被挤掉，导致 `unclosed` 不报。
  现在重复 id 时把旧的转入 orphan 集合，收尾时同样报 unclosed（边界用例实测：注入 9 类缺陷全部命中）。

### 边界（明确不做）

- 不给**已有段落**跨段建书签（一次 `bookmark` op 只能覆盖一个连续区间，`from`/`to` 已覆盖该需求）。
- 不管理 Word 自产的 `_Toc*` / `_Ref*`：它们由 Word 在每次更新域时重建，手工改动没有意义。
- `PAGEREF`（页码交叉引用）不在 `apply` 的 `refTo` 范围内；需要页码引用时用 `extend` 的 `field` op 手写指令。



缘起是一次真实的翻车。给一份 Word 笔记补页脚时，我把书名一起写进了页脚（用户只要求"页脚只保留页码"），
而 **Node `validate.js` 与 PRO `validate.py` 全部 PASS** —— 因为 XML 完全合规。同一轮还有第二处失误：
改了配置源 `build_config.js` 却只重跑 `apply` + `extend`，没重跑配置生成步骤，于是写入器拿着**上一版
config.json** 写出了文件，后面每一道校验照样 PASS。

两个缺陷指向同一处空白：现有校验全在**形式层**（schema / XSD / 结构完整性），没有一层在**意图层**
（产出是否忠于用户原话、链路是否真的跑完整）。本轮补上这一层。

### 新增能力

- **`node scripts/pipeline.js <build.json>` —— 单入口链路编排。**
  步骤类型 `run` / `apply` / `extend` / `validate` / `proValidate`；按声明顺序执行，**首个失败立即中止**
  （没有"继续执行"模式）；未声明 `validate` 步骤时**自动为交付物补上双校验**，构建不可能"悄悄"以未校验状态收尾；
  声明 `verify` 后断言失败即构建失败。`--plan` 只打印解析后的命令链；`--dry-run` 时 apply / extend 走各自干跑。
  解释器自动探测：优先选**能 `import lxml, defusedxml`** 的 python（裸 `python` 可能指向无依赖的解释器，实测踩到）。
- **`--emit-script` —— 无子进程环境的等价方案。**
  部分沙盒禁止进程创建（`spawnSync` 一律 `EBUSY`），此时 pipeline 自动降级：生成 `set -e` 的 bash 脚本并
  以退出码 3 返回。`bash <script>` 一条命令跑完整条链，**中止语义与直接执行完全相同**；脚本内 Windows 路径
  已转正斜杠以适配 Git Bash。
- **`node scripts/verify.js <verify.json>` —— 交付门。**
  ① **断言引擎**：counts（`sectionCount` / `paragraphCount` / `tableCount` / `imageCount`，支持 `eq`/`ne`/`min`/`max`）、
  presence（`contains` / `notContains` / `textMatches`，作用域 `body`|`header`|`footer`|`all`）、parts（`headerText` / `footerText`）、
  fields（`hasField` / `noField`）、setup（`pageSize` / `partExists`）、styles（`styleExists` / `styleMissing`）、
  metadata（`author` / `title`）。任一失败 → **exit 1，交付阻断**。
  ② **变更清单**（`--diff <baseline>`）：输出与基线的结构化差异（节数 / 段落数 / 部件 / 页眉页脚文本 / 域 / 样式 / 新增段落样本），
  供逐条对照用户原话核对。
  零依赖实现：自写 ZIP 读取（EOCD + 中央目录 + `inflateRaw`）与 XML 文本提取，直读 OPC 包，不引入任何第三方包。
- **`node scripts/doctor.js` —— 环境能力探测。**
  报告 Node 引擎、**子进程创建能力**、Python+lxml、LibreOffice、pdftoppm、pandoc、npm docx 的可用性，
  并说明每个缺失项会让哪条路径降级。恒退出 0（是报告，不是门）。

### 关键实现决策（均为实测踩坑后确定）

- **域结果不计入文本。** `<w:instrText>` 与 `separate`…`end` 之间的文本是**缓存值**（PAGE 背后的 `1`、TOC 背后的目录条目）。
  文本提取按 fldChar 状态机跳过它，把域渲染为 `<PAGE>` 占位符——于是断言可写成
  `^\s*第\s*<PAGE>\s*页\s*\/\s*共\s*<NUMPAGES>\s*页\s*$`：既不依赖上次分页结果，也能一眼看出多出来的书名。
- **样式按 `name` 比较，不按 `styleId`。** 写入器会把 styleId 序列化成不透明数字（`2`、`10`…，本技能的 `apply` 即如此），
  拿它做差集会让每次重建都报"全部样式增删"——纯噪音，违背"报告要能读"的初衷。
- **`headerText` / `footerText` 默认"任一部件匹配即通过"。** 一份文档可能持有多个部件（封面空部件、分节覆盖），
  而意图是"页脚显示页码"；只有显式给 `index` 才收紧到某个具体部件。

### 回归验证

把已知真实缺陷重新注入产出物（在页脚加回书名），跑同一套断言：

| 断言 | 结果 |
|---|---|
| `notContains "Linux 基础命令笔记" [footer]` | **FAIL** —— 该断言存在的唯一理由被验证 |
| `footerText` 精确匹配 | **FAIL**，并打印实际值 `"Linux 基础命令笔记　|　第 <PAGE> 页 / 共 <NUMPAGES> 页"` |
| 退出码 | **1**（交付阻断） |

同一缺陷在 Node `validate.js` + PRO `validate.py` 下**全部 PASS**。

端到端：`bash pipeline-run.sh` 一条命令跑完 6 步（生成配置 → apply → extend → 双校验 → 12 条断言），退出码 0。

改动位置：`scripts/pipeline.js`、`scripts/verify.js`、`scripts/doctor.js`（新增）；
`SKILL.md`（新增 § Delivery gate，Commands 表 / Tool Reference 表 / Cross-command invariants / frontmatter description 同步）；`README.md`。

---

## 2026-10-10（第五轮）—— 修正「脚注 / 尾注」的边界表述矛盾

`SKILL.md` 把 `footnotes` 列进 Out of scope，与 `extend` 的实际能力冲突。实测
`{"op":"footnote"}` / `{"op":"endnote"}` 均可写入（创建 `word/footnotes.xml` / `endnotes.xml`、
登记关系与内容类型、回填 id），Node `validate.js` 与 PRO `validate.py` 双校验 PASS。

**边界澄清**（非能力变更）：真正不在范围内的只是**索引** —— 脚注 / 尾注正文存放在
`word/footnotes.xml` / `endnotes.xml`，不参与 `apply` 对 `word/document.xml` 的段落编号，
所以 `#NNN` 定位器够不到它们；**插入与填充是支持的**，走 `extend`。

改动位置：`SKILL.md`（Asking-the-user 段 + Out of scope 段）、`references/standardize.md` §4。

---

## 2026-10-10（第四轮）—— 新增 `batch_fill`：模板 × 数据表 → N 份成品

以「对比市场里所有 docx 技能」为契机做的缺口补齐。21 个市场技能里，本地唯一真正没有的批处理入口
是 `docx-batch-filler-hckp` 独占的**模板套打**（一模板 + 一数据表 → N 份文档）。本轮把它内化进
零依赖 Node 链路，复用已有的 `_shared/merge-runs.js` 与 `DocxReader.copyAndModify`。

### 新增能力

- **`node scripts/batch_fill.js <template> <data.csv|data.json> --out-dir <dir>`** —— 逐行生成文档；
  命名可用 `--name "{{编号}}-{{客户名称}}.docx"` 或 `--name-field <列>`，默认 `output-{{_index}}.docx`。
  占位符语法 `{{field}}` / `{{field|fallback}}` / 内置 `{{_index}}` / `{{_row}}`；`--delims` 可换定界符。
- **跨 run 缝合** —— 页面上看得到的 `{{客户名称}}` 在 XML 里常被拆成 `{{客` / `户` / `名}}`（rsid、
  proofErr、编辑历史）。工具先对每个参与部件合并一次 run，再对仍跨边界的占位符做缝合：值写进**起始
  字符所在的 run**，被跨过的 run 清空但**不删除**（书签 / 域 / 批注锚点位置不变）。
- **多部件** —— `--parts`（默认 `all`）覆盖 document + headers + footers + notes：合同编号、页眉客户名
  这类"重复出现在页眉"的字段一并填充。
- **值形态** —— `\n` → `<w:br/>`、`\t` → `<w:tab/>`（多行地址不再被压成一行）；含边缘空白的 `<w:t>`
  自动补 `xml:space="preserve"`。
- **数据源** —— CSV（自实现 RFC 4180 解析：引号字段、`""` 转义、CRLF、BOM、引号内换行）与 JSON
  数组。行为 `--limit N`、`--dry-run`、`--strict`、`--validate`（逐份跑 schema 校验）。
- **投放前体检** —— 列出模板占位符清单、模板需要而数据缺的列（默认保留原样并警告，`--strict` 致命）、
  数据有而模板未用的列、以及即将生成的文件名。

### 修掉的缺陷（实现期暴露）

| # | 位置 | 症状 | 修复 |
|---|---|---|---|
| 1 | 占位符清单扫描 | 在**原始 XML 字符串**上正则扫描 → 跨 run 的占位符被读成 `公</w:t></w:r><w:r><w:rPr>…<w:t>司名`，凭空造出字段名、并把真实字段（`公司名称`）误报为"未使用列" | 改为按**段落的可见文本**扫描（DOM 遍历 `w:t` 拼接后匹配）；同时把内置变量 `_index` / `_row` 从"缺列"里排除 |

### 语义保证（有测试背书）

| 保证 | 验证方式 |
|---|---|
| 跨 run 缝合正确 | 敌对模板把 `{{公司名称}}` 拆成 3 个 run（中间那个加粗，merge 合不掉），产出为完整值且无残留 `{{` |
| 页眉等同处理 | 页眉部件里的 `{{编号}}` / `{{客户名称}}` 与正文用同一套逻辑替换 |
| 渲染语义不劣化 | `\n` 产出 `<w:br/>`（实测 1 个 br）；跨过 run 被清空而非删除，`w:r` 数量守恒 |
| 未闭合 / 未知字段安全 | `{{未闭合` 原样保留；无对应列且无 fallback → 原样保留 + 汇总警告；`--strict` 退出码 1 |
| 数据表解析健壮 | 含逗号的引号字段（`"12,000"`）、引号内换行、BOM、CRLF 全部正确 |
| 交付即可用 | 3 份产物 Node `validate.js` 与 PRO `validate.py` 双校验 PASS（9 段 / 件） |

### 边界

- 值继承**起始 run** 的 `rPr`：占位符跨了不同格式的 run 时，以第一个 run 的格式为准（其他 run 被清空）。
- 不做图片字段（`{{头像}}` 不能塞图片）、不做表格行级循环/条件区块（需要逐行重复整段表格时请用 `apply` 的循环语义或分次生成）。
- `.doc` 输入仍需先转 `.docx`（`pro/scripts/soffice.py`）；域代码（`instrText`）内的占位符不替换。
- 单份文档若还需要 TOC / 水印 / 图表等，正确的顺序是**先 `apply`+`extend` 造出模板，再用 `batch_fill` 套打**。

### 验证

测试基准与脚本：`Desktop/1111111/docx-master-tidy/test/batch/`（`make_template.py` 造敌对模板，
`verify.py` 12 项断言全 PASS）。回归：CSV / JSON 两种数据源、`--name` / `--name-field` / `--limit` /
`--strict` / `--dry-run` / `--validate` 六条路径均实测；重复运行覆盖既有产物正常。

---

## 2026-10-10（第三轮）—— 对标复核：pro 层 Windows 兼容性 4 处补丁

以「再次对标 docx-pro」为契机做**运行期实测**（不再只是文件 diff）：逐项跑 pro 的可执行路径。
本机实测依赖状态：`lxml` ✅、LibreOffice ❌、Poppler ❌、pandoc ❌、npm `docx` ❌（与
`pro-toolchain.md` 的声明一致）。据此修掉 4 处本地真实缺陷：

| # | 位置 | 症状 | 修复 |
|---|---|---|---|
| 1 | `pro/scripts/soffice.py` `_needs_shim` | `socket.AF_UNIX` 在 Windows **不存在** → `AttributeError`，脚本在寻找二进制之前就崩（并连带 `accept_changes.py` —— 它 import `get_soffice_env`） | 改用 `getattr(socket, "AF_UNIX", None)` 探测；无该常量直接判定「不需要 shim」（shim 本就是 Linux 沙盒专用） |
| 2 | `pro/scripts/soffice.py` `run_soffice` | 硬编码 `["soffice"]`，依赖缺失时抛 `FileNotFoundError` 栈回溯 | 新增 `find_soffice()`（PATH → Windows/macOS/Linux 常见安装根），未装则抛带安装指引的 `RuntimeError` |
| 3 | `pro/scripts/soffice.py` CLI 入口 | 无参数或依赖缺失时报栈回溯 | 新增 `--check` 自检子命令（exit 0/1）；转换路径捕获 `RuntimeError` → 可读提示 + exit 2 |
| 4 | `pro/scripts/accept_changes.py` | `LIBREOFFICE_PROFILE = "/tmp/..."` 硬编码（Windows 无 `/tmp`，`Path()` 会静默落到当前盘根） | 改用 `tempfile.gettempdir()`；入口前置 `find_soffice()` 检查，缺依赖返回可读 error |

实测：`soffice.py --check` → 友好提示 exit 1；转换 → 可读报错 exit 2；`accept_changes` → 可读 error。
回归：`py_compile` 全脚本通过；`validate.py`（419 段 PASS）、`merge_runs.py`（337 run）、`element_order.py`（1366 类型）、`comment.py`（4 批注部件）全部正常。
文档：`pro-toolchain.md` 补运行时自检命令；`SKILL.md` 的 PRO 能力表补 `--check`。

---

## 2026-10-10（第二轮）—— `merge_runs` 归入 Node 引擎

以市场技能 `docx-pro`（display_name「Word 文档处理」v1.0.0，skillId `skill_2099517733028458496`）
为参照做能力对账。该技能此前已整体 vendored 到 `pro/`；本轮把它**唯一还被 Python 独占的核心能力**
——合并碎片化 run——内化到零依赖 Node 链路，使「编辑外部文档」不再必须切解释器。

### 新增能力

- **`node scripts/merge_runs.js`** —— 合并相邻同 `<w:rPr>` 的 run，剥离 `rsid*` 属性与 `<w:proofErr>` 标记。
  与 `pro/scripts/merge_runs.py` **语义等价**，并有两处增强：
  - `--parts headers,footers,notes,all` —— Python 版只处理 `word/document.xml`；Node 版可覆盖页眉 / 页脚 / 脚尾注（实测在同一基准上多合并 12 个页脚 run）。
  - 原地改写走「同目录暂存文件 → rename」，不会留下半写状态的 zip。
- **`apply` 的 `mergeRuns: true` / `--merge-runs`** —— 在**解析段落、解析任何 locator 之前**执行合并，
  因此配置里写的索引与合并后的 run 布局一致。两种入口（配置键 / CLI 标志）产出完全相同。
- 核心逻辑抽到 `_shared/merge-runs.js`，CLI 与 `apply` 共用一份实现。

### 语义保证（有测试背书）

| 保证 | 验证方式 |
|---|---|
| 与 Python 版等价 | 同一碎片化基准（791 run / 180 rsid / 20 proofErr）下合并数均为 **337**，产出 canonical XML 完全一致 |
| 渲染不变 | 419 个段落的可见文本（按 `xml:space` 语义计算）合并前后逐一相等 |
| 修订结构不被破坏 | `<w:ins>` 计数与其中的 run 数不变；`<w:delText>` 不并入 `<w:t>` |
| 幂等 | 对已合并文档再跑一次，合并 **0** run |
| 既有行为不变 | 不带该开关的既有配置，改动前后产出逐字节一致（document / styles / numbering） |

### 边界

- 合并会**重排 run 索引**：写在碎片化原文上的 `set-run` / `runIndex` 定位会失效。
  该开关要与「按合并后布局寻址」二者择一，不能混用。
- 输入必须是 `.docx` / `.dotx` 整包；需要直接处理**已解包目录**时仍走 `pro/scripts/merge_runs.py`。

---

## 2026-10-10

以 `C:\Users\PC\Desktop\doxc-master\out\docx-master-demo.docx` 为靶子做全量复刻，
**只用 `apply` + `extend` 两个写入器**（不借助原 demo 的 5 个 Python 补丁脚本）复现全部 31 个包内部件。
复刻过程中暴露并修掉了下列缺陷。交付件与靶子在部件数、段落数、节数上完全一致，Node + PRO 双校验 PASS。

### 修复 —— `extend`

| # | 位置 | 症状 | 修复 |
|---|---|---|---|
| 1 | `op_field` 的 `inline: true` | 域的 run 被当**兄弟节点**插到锚点段落之外，`<w:r>` 直接落在 `w:body` 层 → XSD 报 `Element '{…}r': This element is not expected. Expected is ( {@…}sectPr )`，Word 打开会弹修复提示。邮件合并的行内 `MERGEFIELD` 全走这条路。 | 改为**追加进锚点段落内部**（返回 "inline field appended into anchor paragraph"）。 |
| 2 | `op_section_break` | 新建的节不带页眉/页脚引用与 `titlePg`，与参考节割裂。 | 新增**继承**：页眉/页脚引用、`titlePg`、`pgSz`/`pgMar`、`cols`、`docGrid` 默认从参考节继承。新增 `inheritHeaders`（默认 `true`）作为断开开关。 |
| 3 | `op_section_break`（同上） | 只声明 `pgSz` 时 `pgMar` 整体丢失 → 静默回落到 Word 默认页边距，正文重排；两者同时继承时还会顺序错位（CT_SectPr 要求 `pgSz` 在 `pgMar` 前）。 | `pgSz` 与 `pgMar` **各自独立**判断继承，并按 `type → pgSz → pgMar` 的顺序重建，不再整块跳过。 |
| 4 | `op_ole` | `previewPart` 指向不存在的部件时**静默**生成悬空关系（文档损坏）；且无法从本地图片引入预览图。 | 新增 `previewFile`（本地图片自动写入 `word/media/` 并登记内容类型）；`previewPart` / `previewFile` 都缺、或指向的部件不存在 → **直接报错**。 |
| 5 | `op_image` | 只能引用**包内已有**媒体部件；对 `extend` 独占的文档（无 `apply` 前置产物）根本没法插图。 | 新增 `file`（本地图片路径，自动入库 + 登记内容类型 + 建关系）。 |

### 修复 —— `apply`

| # | 位置 | 症状 | 修复 |
|---|---|---|---|
| 6 | `settings.elements[].children` | 重复子节点按**位置**覆盖：三个 `docVar` 会被压成一个（按名字一层去重）。 | 改为按**标识属性**（`name` / `id`）做键匹配；无匹配则追加。`docVars` 里的多个 `docVar`、多个 `rsid` 现在逐条落位。 |

### 修复 —— 构件库 `capabilities.py`

| # | 位置 | 症状 | 修复 |
|---|---|---|---|
| 7 | `picture_run` 的 `effects` | 传入的 `a:effectLst` 子元素顺序不合 XSD（如 `glow` 写在 `outerShdw` 之后）→ 校验失败；若把整个 `<a:effectLst>` 包装传进来，会被再包一层导致 `This element is not expected`。 | 自动剥掉外层 `effectLst` 包装，并把子元素按 `blur → fillOverlay → glow → innerShdw → outerShdw → prstShdw → reflection → softEdge` 重排。 |
| 8 | `picture_run` 的 `wrapText` | 口语值 `"square"` 被直接当 `ST_WrapText` 传 → 枚举校验失败（该集合只有 `bothSides` / `left` / `right` / `largest`）。 | `"square"` 映射为 `wp:wrapSquare` + `wrapText="bothSides"`；其余非法值兜底 `bothSides`。 |
| 9 | `ink_run` | 单位错：`cx_cm * 72` 把**厘米当英寸**算，墨迹框放大 2.54 倍；且缺 `coordsize` 被 Word 二次缩放。 | 改用 `cm * 72 / 2.54` 得 pt，并写出 `coordsize`。 |

### 新增能力

- **`paraFormat.bidi`（`<w:bidi/>`）、`runFormat.rtl`（`<w:rtl/>`）、`table.bidiVisual`（`<w:tblPr><w:bidiVisual/>`）** —— 段落 RTL 基线方向、run 级 RTL 文字（影响 cs 字体与标点镜像）、表格镜像列序。
  三张顺序表里本就有这三个元素的位置，缺的只是 schema 与写出逻辑。此前探测误判为"不支持"。
- **SDT 下拉列表** —— `sdt` op 新增 `dropdown: [{ displayText, value }]` → `<w:dropDownList>`；`text` 作为当前值。
- **`settings.elements` 的裸名约定** —— `name` / `attrs` 键不带 `w:` 前缀（写 `docVars` 而非 `w:docVars`），命名空间由引擎补。写错前缀会报 `invalid character in qualified name "w:w:…"`。

### 边界（未做，且短期内不打算做）

| 项 | 原因 |
|---|---|
| 真 SmartArt（`dgm` 语义对象） | layout 是 Word 内置版式算法定义，无法离线可靠构造；用 `wpc` 画布 + 预设几何 + 连接线做**视觉等效**（`smartArt` op）。 |
| 真 InkML 手写笔迹 | 部件注册方式缺可验证的 XSD 覆盖；用 VML 折线等效（`drawing` 的 `ink`），代价是无笔顺回放与压感。 |
| ActiveX 控件 | 属性是二进制 OLE 流，无公开完整规范；交互控件请用 `sdt`。 |
| TOC / 图目录 / 索引的**内容**生成 | 写入的是域，页码需在 Word 里 F9 更新（设计如此）。 |
| 图表动态刷新 | `chart` op 产出静态缓存 + 内嵌工作簿。 |
| OLE 预览图的 EMF/WMF 格式 | 目前用 PNG 预览。 |

### 验证要求（未变）

写完**必须**跑两套校验器，缺一不可：

```bash
node scripts/validate.js out.docx
python pro/scripts/validate.py out.docx
```
