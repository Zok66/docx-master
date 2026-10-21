# docx-master

<div align="center">

**简体中文** | [English](README_EN.md)

</div>

<div align="center">

**直接操作 OOXML 的 Word 文档技能**

打开 `.docx` 这个 zip、直接改里面的 XML、原样封回去 —— **生成 · 读取 · 编辑 · 校验**，主链路零依赖

![License](https://img.shields.io/badge/License-MIT-blue)
![release](https://img.shields.io/badge/release-v0.16.0-green)
![stars](https://img.shields.io/badge/stars-%F0%9F%8C%9F%20%E2%9C%A8-yellow)
![standard](https://img.shields.io/badge/标准版-日常办公%20·%20论文%20·%20报告-3182ce)
![pro](https://img.shields.io/badge/pro版-水印%20·%20形状%20·%20图表%20·%20OLE-805ad5)
![deps](https://img.shields.io/badge/主链路-零依赖%20Node%20%2B%20lxml-orange)

[文档示例 demo](#效果-demo真实产出可直接下载打开) · [常见问题](README.md#常见问题-faq) · [更新日志](CHANGELOG.md) · [开发规范](README.md#开发规范) · [安全策略](SECURITY.md) · [贡献](CONTRIBUTING.md)

</div>

> 🤖 **下一个 AI / 开发者请先读 [SKILL.md](SKILL.md)**（主流程、硬规则、路由索引），不要先全库扫描。

**作者**：佐罗斯（[Zok66](https://github.com/Zok66)） · MIT 免费开源，可自由使用 / 修改 / 商用（保留版权声明即可）

## 目录

- ① [亮点](#亮点) —— 一句话说清它凭什么存在
- ② [效果 demo](#效果-demo真实产出可直接下载打开) · [隐私与安全](#隐私与安全) · [两个版本](#两个版本) —— 真实产出先看效果；本地处理；标准版 / pro 版边界
- ③ [版本与更新日志](#版本与更新日志) · [文档索引](#文档索引) —— 当前版本、历史变更、全部文档
- ④ [能力一览](#能力一览) · [安装](#安装) · [快速开始](#快速开始) —— 能做什么、怎么装上、四条命令跑通
- ⑤ [设计原则](#设计原则) · [开发规范](#开发规范) · [常见问题 FAQ](#常见问题-faq) —— 为什么这么设计、怎么提交、常见坑
- ⑥ [目录结构](#目录结构) · [回归测试](#回归测试) · [依赖](#依赖) · [适用边界](#适用边界) · [许可](#许可)

## 亮点

- **原文件永不被修改** —— 所有写命令产出新副本，写盘前自动双校验（Node schema + XSD/OPC）
- **生成 · 读取 · 编辑 · 校验全链路，零 npm 依赖** —— 主链路只用 Node 标准库 + Python lxml
- **规范化**：把"手打的 `一、`/`1.1`"换成真多级自动编号，按角色重塑整篇排版
- **论文 / 实训报告开箱即用**：目录域、页码、页眉页脚、脚注尾注、题注、交叉引用、引文书目
- **交付门**：声明式断言（失败 exit 1 **阻断交付**）+ 变更清单，专抓"格式合规但用户没要求"的缺陷
- **看不见的对象有专门守门**：悬空书签 / 坏引用在 Word 里只会印"错误！未定义书签"，XSD 看不见它——`inspect_bookmarks` + `refsResolve` 看得见
- **批量套打**：一份模板 × 一张 CSV → N 份成品（Word 邮件合并的替代品）
- **pipeline 一条命令跑完整条链路**，任一步失败立即中止，不留"跑了前半段"的半成品
- **AI 友好**：`SKILL.md` 主控入口 + 规则唯一来源契约，回归测试守住"规则不许凭空消失"

零依赖不是口号，是四个具体的取舍 ——「不做什么 / 换成什么 / 换来什么」：

| ① 不依赖 | ② 换成什么 | ③ 于是能做什么 |
|---|---|---|
| Microsoft Word | 直接读写 OOXML（ECMA-376 / ISO 29500） | 段落级重排、多级自动编号、目录域、页面边框、内容控件、批量套打 |
| LibreOffice | 主链路纯 Node + Python `lxml` | 无 GUI、无 COM、可在 CI / 沙箱 / 服务器上跑 |
| npm `docx` 包 | 自己封 `w:p` / `w:tbl` / `w:sdt` | 库没封装的角落也能改，不被 API 表达力卡住 |
| 云转换服务 | 全程本地 | 文档不出机器，涉密 / 内部资料可用 |

## 效果 demo（真实产出，可直接下载打开）

**🌐 [在线预览（前端直接渲染，无需下载）](https://zok66.github.io/docx-master/preview/)**

下面四份 `.docx` 全部由本技能生成，不是手写的 —— [`demo/`](demo/) 目录即仓库内文件：

| 文件 | 演示重点 |
|---|---|
| [`demo/docx-master-demo.docx`](demo/docx-master-demo.docx) | **综合能力全景**（10 章 / 8 节 / 243 段）：文字排版、引用注释与目录、页面分节、表格、插图与题注、域与自动更新、批量自动化、链接导航、输出校对 |
| [`demo/docx-master能力演示-审阅版.docx`](demo/docx-master能力演示-审阅版.docx) | **审阅场景**：同一篇文档带修订痕迹与批注，演示 `extend` 修订写入、批注锚点与 `inspect_blockers` 的 blocker 识别 |
| [`demo/docx-master拼音注音演示.docx`](demo/docx-master拼音注音演示.docx) | **拼音注音（`w:ruby`）**：逐字注音的 OOXML 直写实现，适合语文教材 / 对外汉语类排版 |
| [`demo/docx-master新能力演示.docx`](demo/docx-master新能力演示.docx) | **第三十轮 31 个格式字段**：荧光笔 16 色、字符缩放 / 位置 / 间距、下划线 18 线型、阴影空心阳文阴文、边框与底纹等 |

打开目录页后按 `Ctrl+A` → `F9` 更新域，可看到目录 / 图表目录 / 交叉引用的完整联动效果（原因见 FAQ ③）。

## 隐私与安全

- **文档全部在本地处理**，不上传到任何云转换服务
- **不执行文档内的宏 / OLE 载荷**；不联网（`doctor` 仅探测本地已装工具）
- 处理**不受信任来源**的 docx 时，建议在隔离环境运行（与打开陌生 Word 文档的常识一致）
- 漏洞报告走私密渠道，详见 [SECURITY.md](SECURITY.md)

不依赖 Microsoft Word，不依赖 LibreOffice，不依赖 npm 上的 `docx` 包。打开 `.docx` 这个 zip、
直接改里面的 XML，再原样封回去 —— 所以能做的东西比"库能封装的 API"多得多。
下面按「不做什么 / 换成什么 / 换来什么」三栏列清楚：

## 两个版本

同一个仓库、两个分支。**标准版面向日常办公、论文、实训报告**——够用且轻；**pro 版是全功能**，包含所有不常用的扩展对象。

| 能力 | 标准版（`main`，默认） | pro 版（`pro`） |
|---|:---:|:---:|
| 生成 / 编辑 / 规范化 / 模板导入 | ✅ | ✅ |
| 读取侦察（12 个 `inspect_*` + `find_text`） | ✅ | ✅ |
| 页数 / 分页 / 页码 / 目录 / 图表目录 / 索引 | ✅ | ✅ |
| 页眉页脚（含分节覆盖）/ 页面设置 / 页面边框 | ✅ | ✅ |
| 脚注尾注 / 题注 / 交叉引用 / 书签 | ✅ | ✅ |
| 插图（嵌入式图片、裁剪、影效、浮动环绕） | ✅ | ✅ |
| 批注 / 修订 / 接受修订 / 批量套打 / 交付门 | ✅ | ✅ |
| 封面页 / 首字下沉 / 内容控件 / 引文书目 | ✅ | ✅ |
| 水印 | — | ✅ |
| 形状 / 文本框 / 艺术字 / 画布 / 智能图形 | — | ✅ |
| 图表（内嵌工作簿）/ OLE 嵌入 | — | ✅ |
| 构建基块（自动图文集）/ 自定义 XML | — | ✅ |
| PRO 完整 XSD 架构集（全套 schema 校验） | 精简 | 全套 |

装法与版本切换见「[安装](#安装)」。被剥离的 8 个 `extend` op 在**代码里仍在**（只是没注册进 `OPERATIONS`），调用会明确报错并提示可用的 16 项操作——不会静默降级。清单与历史实现见 [`references/pro-only-ops.md`](references/pro-only-ops.md)。

## 版本与更新日志

- **当前版本：[v0.16.0](https://github.com/Zok66/docx-master/releases/tag/v0.16.0)**（2026-10-20，标准版首发）。
- **协议**：[MIT](LICENSE) —— 自由使用、修改、分发与商业使用，保留版权与许可声明即可。
- **升级**：技能目录内 `git pull` 覆盖安装，完成后按「[回归测试](#回归测试)」跑一遍自检。
- **下载 / 历史版本**：见 [Releases](https://github.com/Zok66/docx-master/releases) 页，Latest 即当前推荐版，每个版本附能力清单与验收记录。

### 版本速览

| 版本 | 定位 | 分支 | 关键变化 |
|---|---|---|---|
| **v0.16.0** ⬅ 当前 | 标准版 · 日常办公 / 论文 / 实训报告 | `main` | 版本拆分精简、社区文件、Star History、中英双语 README |
| v0.16.0-pro | 全功能版 · 含水印 / 形状 / 图表 / OLE | `pro` | 同标准版代码 + 24 项全量 `extend` |
| v0.15.0 | 拆分前最后一版（全功能） | — | 31 字段格式扩展、拼音注音、WPS「插入」对齐 |
| 更早 | — | — | 见 [CHANGELOG.md](CHANGELOG.md) 第五~十五轮 |

### 更新日志

完整记录见 **[`CHANGELOG.md`](CHANGELOG.md)** —— 只记**能力与行为变更**（新增字段 / 修掉的缺陷 / 明确的边界），
文档润色不记，每条都可在 `apply.js`、`scripts/extend/*.py` 里检索到对应实现。

最近三轮：

| 日期 | 版本 | 摘要 |
|---|---|---|
| 2026-10-20 | — | 授权确认：维持 MIT（曾一度改 PolyForm NC，已回撤，`LICENSE` 与 v0.16.0 一致） |
| 2026-10-20 | **v0.16.0** | ① 拆标准版 / pro 版（`OPERATIONS` 16 vs 24）② 8 个 op 剥离并明确报错 ③ 中英双语 + 社区文件 + Star History |
| 2026-10-20 | v0.15.0 | ① run / 段落级 31 字段 + 拼音注音 ② WPS「插入」对齐 ③ 批注锚点等三项缺陷修复 |

发版流程（打 tag → 双平台 Release → 同步核对）见 [`Push-Sync.markdown`](Push-Sync.markdown)。

## 文档索引

| 文档 | 内容 |
|---|---|
| [`SKILL.md`](SKILL.md) | 技能主入口 —— 主流程、硬规则、路由索引（AI 触发即读） |
| [`references/pro-only-ops.md`](references/pro-only-ops.md) | 标准版不含的 8 个操作清单 + 切到 pro 版的方法 |
| [`CHANGELOG.md`](CHANGELOG.md) | 完整更新记录（能力与行为变更，逐轮可检索） |
| [`CONTRIBUTING.md`](CONTRIBUTING.md) | 怎么贡献：开发流程、硬约束、缺陷报告格式 |
| [`CODE_OF_CONDUCT.md`](CODE_OF_CONDUCT.md) | 行为准则 |
| [`SECURITY.md`](SECURITY.md) | 安全策略：支持的版本、漏洞私密报告渠道、威胁模型 |
| [`tests/README.md`](tests/README.md) | 回归测试怎么跑、怎么加规则 / 用例 |
| [`LICENSE`](LICENSE) | MIT 许可证 |
| [`THIRD-PARTY-NOTICES.md`](THIRD-PARTY-NOTICES.md) | 第三方组件声明（Temml / ECMA-376 XSD 的来源与授权） |
| [`Push-Sync.markdown`](Push-Sync.markdown) | GitHub ↔ Gitee 全量同步指南（分支 / Tag / Release / 许可证） |
| [`references/`](references/) | 领域参考：配置 schema、定位器契约、工具全表、编号格式、题注、公式 … |
| [`pro/`](pro/) | 可选 Python 工具链：批注、接受修订、渲染、XSD 校验 |

## 能力一览

| 动作 | 说明 |
|---|---|
| **生成** | 省略 `source` 即以内置空白模板脚手架一篇新文档；也可从现有模板导入样式与编号 |
| **规范化** | 按角色重塑整篇文档：段落分类 → 绑定命名样式 → 安装多级自动编号，把"手打的 `一、`/`1.1`"换成真编号 |
| **编辑** | 定位式外科手术：替换 / 插入 / 删除段落、表格单元格、嵌入图片，可选生成修订痕迹 |
| **扩展** | `extend` 补引擎"故意不写"的部分：目录 / 图表目录 / 索引域、分页分节、页码、命名书签、脚注尾注、页面边框与背景、插入图片、内容控件、封面页、引文书目。水印 / 形状 / 图表 / OLE / 智能图形等属 pro 版（见「[两个版本](#两个版本)」）|
| **书签 / 交叉引用** | 建命名书签（新建段落、已有段落、跨段区间三种入口）、`refTo` → REF 域交叉引用、`{ "link": "#名字" }` 页内跳转；并用 `inspect_bookmarks` + `refsResolve` 断言抓"引用指向不存在的书签"这类**看不见的缺陷** |
| **批量** | 一份模板 + 一张 CSV / JSON 数据表 → 每行一份成品（Word 邮件合并的替代品） |
| **编排** | `pipeline`：把「生成配置 → apply → extend → 校验 → 断言」声明成一条链路，**一条命令跑完**，任一步失败立即中止——不留"跑了前半段"的半成品。无子进程环境（沙盒）可 `--emit-script` 生成 `set -e` 脚本，`bash` 一次执行得到同样的中止语义 |
| **读取** | 结构侦察：样式定义与继承链、编号方案与消耗数、节与页面设置、主题、题注、表格、文档骨架、任意范围的段落全文与计算样式、逐 run 格式、书签与交叉引用 |
| **校验** | 三层：Node schema 检查 + XSD / OPC 包图检查（写盘前自动跑），外加 `doctor` 的环境能力探测 |
| **交付门** | `verify`：① 声明式断言（如"页脚不得出现书名"）失败即非零退出、**阻断交付**；② 变更清单（产出 vs 基线）逐项列出本次引入的改动，供对着用户原话核对 |
| **批注 / 修订** | 添加批注（含回复）、接受全部修订生成干净副本 |
| **转换** | 旧版二进制 `.doc` → `.docx`（需 LibreOffice），渲染成图做视觉复核 |

## 安装

作为技能放进你的 skills 目录即可（WorkBuddy / Claude Code 等支持 `SKILL.md` 的环境）：

```bash
# 标准版（默认，克隆即得）
git clone https://github.com/Zok66/docx-master.git ~/.workbuddy/skills/docx-master

# 要全功能版：同一个仓库切分支，不用重装
cd ~/.workbuddy/skills/docx-master && git checkout pro

# 切回标准版
git checkout main
```

## 快速开始

```bash
SK=~/.workbuddy/skills/docx-master

# ① 开工前：探测环境能力（哪些可选项真的可用、哪些会降级）
node $SK/scripts/doctor.js

# ② 先侦察：结构、样式、编号、骨架
node $SK/scripts/overview.js 我的文档.docx
node $SK/scripts/inspect_bookmarks.js 我的文档.docx   # 书签 + 所有引用 + 悬空引用（有问题则 exit 1）

# ③ 单步设计：一份 config，先干跑不落盘，再落盘（输出新文件，原文件不动）
node $SK/scripts/apply.js --dry-run my.config.json
node $SK/scripts/apply.js my.config.json
python $SK/scripts/extend/extend.py extend.config.json   # 目录 / 页码 / 脚注 / 页面边框 …（标准版 16 项；水印 / 图表属 pro 版）

# ④ 或者：把上面几步声明成一条链路，一条命令跑完（推荐）
node $SK/scripts/pipeline.js build.json          # 任一步失败即中止，末尾自动双校验
node $SK/scripts/pipeline.js build.json --plan   # 只看将要执行什么，不落盘

# ⑤ 交付门：断言 + 变更清单
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
    { "type": "footerText", "regex": "^\\s*第\\s*<PAGE>\\s*页\\s*\\/\\s*共\\s*<NUMPAGES>\\s*页\\s*$" },
    { "type": "refsResolve" },
    { "type": "bookmarkProblems" }
  ]
}
```

中间两条正是为了抓「页脚混进书名」这类**格式完全合规、但用户没要求**的改动——任何 schema 校验器都看不见它。最后两条同理：一个指向不存在书签的 REF / 跳转链接，在 Word 里只会渲染成"错误！未定义书签。"，而 XML 完全合规。

批量套打：

```bash
node $SK/scripts/batch_fill.js 合同模板.docx 客户.csv \
  --out-dir out/ --name "{{编号}}-{{客户名称}}.docx" --validate
```

## 设计原则

1. **原文件永不被修改。** 每个写盘命令产出新副本，原文件保持原样。
2. **稀疏声明。** 只声明坏掉的、缺失的、或用户明确要改的部分；没碰的样式 / 编号 / 段落原样保留。
3. **先校验再交付。** 两个校验器检查不同层（Node schema 层 vs XSD + OPC 包图层），都要过。
4. **多步链路必须由机器串起来。** 靠人记着"改完内容源要重跑配置生成"是最容易翻车的地方：源头没重跑，写入器会拿着**上一版配置**写出文件，而后面每一道校验都报 PASS。链路声明一次，跑一次。
5. **"格式合规"不等于"可以交付"。** 用户没要求的改动（页脚多一行书名、多加一行署名）是合规的 XML，也是缺陷。断言管机械那一半，读变更清单管剩下那一半。
6. **看不见的对象要专门查。** 书签没有渲染、没有样式、不加字数：丢一个，文件照样打开、照样通过 XSD，只有 Word 在某句话中间印出"错误！未定义书签。"。`inspect_bookmarks`（人工查看）与 `refsResolve`（机器守门）是唯一会去看它的两处。
7. **内容与结构分离。** 标题 / 列表 / 题注这类"结构信号"由样式、编号、SEQ 域渲染，不写死在文本里。
8. **主控只放规则，细节按需加载。** `SKILL.md` 每次触发都会全量进入上下文，所以它只承载"必须当场知道"的东西：主流程、硬规则、路由索引。长表与语义细节住在 `references/`，用到才读。一条规则只在**一处**定义，别处只做指向 —— `tests/run.mjs` 用契约清单守住这条线：规则可以换个文件放，但不能凭空消失。

## 开发规范

### 提交说明

不强制格式，写清**改了什么 / 为什么**即可。两个约定：

① **一次提交一件事**。代码 + 对应测试 + CHANGELOG 同一提交；纯文档润色单独提。

② **两个分支同步**：`main` 与 `pro` 共用代码，改动落到 `main` 后
`git checkout pro && git checkout main -- <改过的文件>` 再提一次。

### 代码与文档硬约束

| 约束 | 怎么守 |
|---|---|
| `SKILL.md` ≤ 20000 B | `node tests/run.mjs` 会卡；塞不下就下移到 `references/` 并留一行指向 |
| **一条规则只在一处定义** | 规则进 `tests/contract.json` 契约清单，删了测试就红 |
| 所有文档内链（含锚点）可解析 | `node tests/run.mjs` 的「链接」类契约 |
| 新增能力同步补契约与 CHANGELOG | 见 [CONTRIBUTING.md](CONTRIBUTING.md)「开发流程」 |

### 提交前自检

按「[回归测试](#回归测试)」跑 `--smoke` 那一档即可。贡献流程、缺陷报告格式见 [CONTRIBUTING.md](CONTRIBUTING.md)。

## 常见问题 FAQ

<details>
<summary><b>① 需要装 Microsoft Word 或 LibreOffice 吗？</b></summary>

主链路**都不需要**。哪些能力依赖什么，见「[亮点](#亮点)」末尾的取舍表与「[依赖](#依赖)」一节；
两个可选能力（`.doc` 转换、渲染成图）需要 LibreOffice。`node scripts/doctor.js` 会探测本机实际可用情况。

</details>

<details>
<summary><b>② 会改动我的原文件吗？</b></summary>

**不会。** 原文件一个字节都不动，`verify.js --diff` 也只读不写。

</details>

<details>
<summary><b>③ 目录生成了但打开是空的 / 页码没显示？</b></summary>

这是 OOXML 的固有行为，**不是缺陷**。TOC / 页码 / SEQ 都是**域（field）**，Word 打开时按缓存值显示；
域代码已写入，但结果页需要 Word 计算。

处理方式：打开后 `Ctrl+A` → `F9`（或右键目录 → 更新域）刷新一次。WPS 同理。

</details>

<details>
<summary><b>④ 「错误！未定义书签。」是什么？</b></summary>

有个 REF 域或页内跳转链接指向了**不存在的书签**。这种问题 XSD 查不出来、Word 只在正文里印一行字。

先用侦察工具确认：

```bash
node scripts/inspect_bookmarks.js 我的文档.docx   # 书签 + 所有引用 + 悬空引用，有问题 exit 1
```

再用 `verify` 的 `refsResolve` / `bookmarkProblems` 断言把它挡在交付前。

</details>

<details>
<summary><b>⑤ 标准版和 pro 版怎么选？切过去要重装吗？</b></summary>

不用重装 —— 同一个仓库两个分支，`git checkout pro` / `git checkout main` 即可，命令见「[安装](#安装)」。

日常办公、论文、实训报告用标准版就够；要水印、形状、图表再切 `pro`。完整差异见「[两个版本](#两个版本)」，8 个 op 清单见 [`references/pro-only-ops.md`](references/pro-only-ops.md)。

</details>

<details>
<summary><b>⑥ 标准版里误用了 pro 才有的操作会怎样？</b></summary>

**明确报错，不会静默降级**。被剥离的 8 个 op 函数体仍在代码里，只是没注册进 `OPERATIONS`；
调用时会 `exit` 并列出标准版可用的 16 项操作，同时提示怎么切到 pro 分支。

</details>

<details>
<summary><b>⑦ 中文排版（首行缩进、字体、间距）默认是什么？</b></summary>

正文默认两字符首行缩进（`firstLineIndent: "2char"`）、CJK 字体与字号按中文公文 / 论文习惯配置，
中英文之间是否自动加空格由 `lint.panguSpacing` 控制（默认静默）。
具体字段见 [`references/config-schema.md`](references/config-schema.md) 与 [`references/chinese-font-sizes.md`](references/chinese-font-sizes.md)。

</details>

<details>
<summary><b>⑧ 能批量生成吗（一份模板 → N 份文档）？</b></summary>

能，这是 Word 邮件合并的替代品 —— 一条命令，见「[快速开始](#快速开始)」末尾的 `batch_fill.js` 示例。
模板里放占位符，CSV / JSON 每行产出一份成品，`--validate` 逐份跑校验。

</details>

<details>
<summary><b>⑨ 处理涉密 / 内部文档安全吗？</b></summary>

全程本地：不联网、不上传云转换、不执行文档里的宏和 OLE 载荷。详见 [隐私与安全](#隐私与安全) 与 [SECURITY.md](SECURITY.md)。
处理来源不明的文档时，仍建议在隔离环境跑（和打开陌生 Word 文档的常识一致）。

</details>

<details>
<summary><b>⑩ 支持 PDF / xlsx / Markdown 吗？</b></summary>

**不做**。这是 docx 专用技能：PDF、表格（xlsx）、纯文本 / Markdown 源文件都不在范围内
（除非任务本身就是**输出** docx）。详见 [适用边界](#适用边界)。

</details>

<details>
<summary><b>⑪ 布局表格能重排结构吗？</b></summary>

不能直接重排表格结构（增删行 / 合并单元格的整体重排），但**单元格内的段落可以正常编辑**。
这是刻意划的边界：结构重排的失败模式太多，交给人在 Word 里做更稳。

</details>

<details>
<summary><b>⑫ 可以商用吗？</b></summary>

可以。[MIT](LICENSE) 许可，自由使用、修改、分发、商用，**保留版权与许可声明**即可。

</details>

## 目录结构

```
docx-master/
├── SKILL.md              # 技能主入口 —— 主流程 / 硬规则 / 路由索引（细节按需读 references/）
├── CHANGELOG.md          # 能力与行为变更（新增 / 修复 / 边界）
├── README.md             # 本文件
├── scripts/              # 主链路：零依赖 Node + Python lxml
│   ├── apply.js          # 统一写入器（样式 / 编号 / 主题 / 模板 / 编辑）
│   ├── batch_fill.js     # 模板 × 数据表 → N 份文档
│   ├── pipeline.js       # 单入口编排：一条命令跑完整条链路，失败即停
│   ├── verify.js         # 交付门：声明式断言 + 变更清单（零依赖直读 OPC）
│   ├── inspect_bookmarks.js  # 书签 / 引用 / 悬空引用侦察（有问题 exit 1）
│   ├── doctor.js         # 环境能力探测（子进程 / Python / 渲染工具链）
│   ├── merge_runs.js     # 合并碎片化 run（外部文档编辑前预处理）
│   ├── validate.js       # schema 校验
│   ├── overview.js       # 与 12 个 inspect_* 侦察工具
│   ├── _shared/          # 共享模块：OPC 读取、书签分析、定位器、校验器 …
│   └── extend/           # 扩展写入器（TOC / 书签 / 脚注 / 水印 / 图形 / 图表 / OLE …）
├── references/           # 领域参考：配置 schema、定位器契约、工具全表、编号格式、题注、公式 …
├── tests/                # 回归测试：结构 / 规则 / 资源 / 链接 四类契约 + smoke case
└── pro/                  # 可选的 Python 工具链（批注、接受修订、渲染、XSD 校验）
```

本节描述的是**标准版**（`main` 分支）的目录；`pro` 分支结构相同。

## 回归测试

零依赖，一条命令：

```bash
node tests/run.mjs            # 结构 + 规则 + 资源 + 链接 四类契约
node tests/run.mjs --quiet    # 只打印失败项
node tests/run.mjs --smoke    # 额外跑 tests/cases/*.json
```

它回答四个问题：`SKILL.md` 的 frontmatter 与体积是否合规、**契约清单里每条「不许删的硬规则」是否还在**、声明的 references / scripts 是否都存在、所有文档链接（含锚点）是否可解析。改动 skill 之后跑一遍即可。

加规则、加用例的做法见 [`tests/README.md`](tests/README.md)。

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

## Star History

<picture>
  <source media="(prefers-color-scheme: dark)" srcset="assets/star-history-dark.svg">
  <img src="assets/star-history-light.svg" alt="Zok66/docx-master 的 Star History 曲线">
</picture>

由 [`tools/gen-star-history.py`](tools/gen-star-history.py) 每 12 小时拉取 GitHub 星标时间并更新
（[workflow](.github/workflows/star-history.yml)）。浅色、深色 SVG 保存在本仓库；
曲线按当前保留的星标汇总。
