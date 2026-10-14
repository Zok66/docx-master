# Changelog — docx-master

只记录**能力与行为变更**：新增字段 / 修掉的缺陷 / 明确的边界。文档润色不记。
每条都可在 `apply.js`、`scripts/extend/*.py` 里检索到对应实现。

---

## 2026-10-13（第六轮）—— 新增交付层：`pipeline` 编排 + `verify` 交付门 + `doctor` 探测

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
