# docx-master

> 直接操作 OOXML 的 Word 文档工具箱 —— **生成、读取、编辑、校验**，主链路零依赖。

不依赖 Microsoft Word，不依赖 LibreOffice，不依赖 npm 上的 `docx` 包。打开 `.docx` 这个 zip、
直接改里面的 XML，再原样封回去 —— 所以能做的事比"库能封装的 API"多得多：
段落级重排、多级自动编号、目录域、脚注尾注、水印、页面边框、形状、图表、OLE、内容控件、批量套打。

原文件永不被修改：每个写盘命令都产出新文件，且在写入前跑校验。

## 能力一览

| 动作 | 说明 |
|---|---|
| **生成** | 省略 `source` 即以内置空白模板脚手架一篇新文档；也可从现有模板导入样式与编号 |
| **规范化** | 按角色重塑整篇文档：段落分类 → 绑定命名样式 → 安装多级自动编号，把"手打的 `一、`/`1.1`"换成真编号 |
| **编辑** | 定位式外科手术：替换 / 插入 / 删除段落、表格单元格、嵌入图片，可选生成修订痕迹 |
| **扩展** | `extend` 补引擎"故意不写"的部分：目录 / 图表目录 / 索引域、脚注尾注、水印、页面边框与背景、形状 / 文本框 / 画布 / 艺术字、图表、OLE、构建基块、封面、自定义 XML、内容控件 |
| **批量** | 一份模板 + 一张 CSV / JSON 数据表 → 每行一份成品（Word 邮件合并的替代品） |
| **编排** | `pipeline`：把「生成配置 → apply → extend → 校验 → 断言」声明成一条链路，**一条命令跑完**，任一步失败立即中止——不留"跑了前半段"的半成品。无子进程环境（沙盒）可 `--emit-script` 生成 `set -e` 脚本，`bash` 一次执行得到同样的中止语义 |
| **读取** | 结构侦察：样式定义与继承链、编号方案与消耗数、节与页面设置、主题、题注、表格、文档骨架、任意范围的段落全文与计算样式、逐 run 格式 |
| **校验** | 三层：Node schema 检查 + XSD / OPC 包图检查（写盘前自动跑），外加 `doctor` 的环境能力探测 |
| **交付门** | `verify`：① 声明式断言（如"页脚不得出现书名"）失败即非零退出、**阻断交付**；② 变更清单（产出 vs 基线）逐项列出本次引入的改动，供对着用户原话核对 |
| **批注 / 修订** | 添加批注（含回复）、接受全部修订生成干净副本 |
| **转换** | 旧版二进制 `.doc` → `.docx`（需 LibreOffice），渲染成图做视觉复核 |

## 安装

作为技能放进你的 skills 目录即可（WorkBuddy / Claude Code 等支持 `SKILL.md` 的环境）：

```bash
git clone https://github.com/Zok66/docx-master.git ~/.workbuddy/skills/docx-master
```

## 快速开始

```bash
SK=~/.workbuddy/skills/docx-master

# 0) 开工前：探测环境能力（哪些可选项真的可用、哪些会降级）
node $SK/scripts/doctor.js

# 1) 先侦察：结构、样式、编号、骨架
node $SK/scripts/overview.js 我的文档.docx

# 2) 单步设计：一份 config，先干跑不落盘，再落盘（输出新文件，原文件不动）
node $SK/scripts/apply.js --dry-run my.config.json
node $SK/scripts/apply.js my.config.json
python $SK/scripts/extend/extend.py extend.config.json   # 目录 / 水印 / 脚注 / 图表 …

# 3) 或者：把上面几步声明成一条链路，一条命令跑完（推荐）
node $SK/scripts/pipeline.js build.json          # 任一步失败即中止，末尾自动双校验
node $SK/scripts/pipeline.js build.json --plan   # 只看将要执行什么，不落盘

# 4) 交付门：断言 + 变更清单
node $SK/scripts/verify.js verify.json                     # 断言失败 → exit 1，交付阻断
node $SK/scripts/verify.js out.docx --diff 源文档.docx      # 列出本次改了什么，逐条核对
```

断言长这样（`verify.json`）——它是用机器表达的"用户要求"：

```json
{
  "file": "out.docx",
  "baseline": "源文档.docx",
  "assert": [
    { "type": "sectionCount", "eq": 3 },
    { "type": "hasField", "field": "TOC", "scope": "body" },
    { "type": "headerText", "regex": "佐罗斯" },
    { "type": "notContains", "text": "我的书名", "scope": "footer" },
    { "type": "footerText", "regex": "^\\s*第\\s*<PAGE>\\s*页\\s*\\/\\s*共\\s*<NUMPAGES>\\s*页\\s*$" }
  ]
}
```

最后两条正是为了抓「页脚混进书名」这类**格式完全合规、但用户没要求**的改动——任何 schema 校验器都看不见它。

批量套打：

```bash
node $SK/scripts/batch_fill.js 合同模板.docx 客户.csv \
  --out-dir out/ --name "{{编号}}-{{客户名称}}.docx" --validate
```

## 设计原则

- **原文件永不被修改。** 每个写盘命令产出新副本，原文件保持原样。
- **稀疏声明。** 只声明坏掉的、缺失的、或用户明确要改的部分；没碰的样式 / 编号 / 段落原样保留。
- **先校验再交付。** 两个校验器检查不同层（Node schema 层 vs XSD + OPC 包图层），都要过。
- **多步链路必须由机器串起来。** 靠人记着"改完内容源要重跑配置生成"是最容易翻车的地方：源头没重跑，写入器会拿着**上一版配置**写出文件，而后面每一道校验都报 PASS。链路声明一次，跑一次。
- **"格式合规"不等于"可以交付"。** 用户没要求的改动（页脚多一行书名、多加一行署名）是合规的 XML，也是缺陷。断言管机械那一半，读变更清单管剩下那一半。
- **内容与结构分离。** 标题 / 列表 / 题注这类"结构信号"由样式、编号、SEQ 域渲染，不写死在文本里。

## 目录结构

```
docx-master/
├── SKILL.md              # 技能主入口（给 AI 读的工作流与规则）
├── CHANGELOG.md          # 能力与行为变更（新增 / 修复 / 边界）
├── README.md             # 本文件
├── scripts/              # 主链路：零依赖 Node + Python lxml
│   ├── apply.js          # 统一写入器（样式 / 编号 / 主题 / 模板 / 编辑）
│   ├── batch_fill.js     # 模板 × 数据表 → N 份文档
│   ├── pipeline.js       # 单入口编排：一条命令跑完整条链路，失败即停
│   ├── verify.js         # 交付门：声明式断言 + 变更清单（零依赖直读 OPC）
│   ├── doctor.js         # 环境能力探测（子进程 / Python / 渲染工具链）
│   ├── merge_runs.js     # 合并碎片化 run（外部文档编辑前预处理）
│   ├── validate.js       # schema 校验
│   ├── overview.js       # 与 11 个 inspect_* 侦察工具
│   └── extend/           # 扩展写入器（TOC / 脚注 / 水印 / 图形 / 图表 / OLE …）
├── references/           # 领域参考：配置 schema、定位器契约、编号格式、题注、公式 …
└── pro/                  # 可选的 Python 工具链（批注、接受修订、渲染、XSD 校验）
```

## 依赖

| 部分 | 依赖 |
|---|---|
| 主链路（`scripts/`） | Node（无第三方包）+ Python + `lxml` |
| 扩展写入器（`scripts/extend/`） | Python + `lxml` |
| 校验（`scripts/validate.js`） | Node，零依赖 |
| 校验（`pro/scripts/validate.py`） | Python + `lxml` + `defusedxml` |
| 可选能力：渲染成图 / `.doc` 转换 / 接受修订 | LibreOffice（`soffice`）+ Poppler（`pdftoppm`） |

## 适用边界

**不做**：PDF、表格（xlsx）、纯文本 / Markdown 源文件（除非任务就是**输出** docx）。

**明确不在范围内**：布局表格的**结构重排**（单元格内的段落仍可编辑）、目录的**正文内容**（域会生成，但需在 Word 里更新域才算出结果）、页码交叉引用、域代码内部编辑。

**已支持**：页眉 / 页脚（含分节覆盖）、页面设置、脚注 / 尾注、批注、修订痕迹。

## 许可

MIT
