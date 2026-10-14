# PRO 工具链（docx-pro 整合层）

本文件说明 `docx-master` 与 `docx-pro` 两个技能的整合方式、能力路由、依赖状态与回退路径。

- 整合日期：2026-10-10
- 来源：市场技能 `docx-pro`（display_name「Word 文档处理」v1.0.0，skillId `skill_2099517733028458496`，author `zxh`）
- vendored 位置：本技能目录下的 `pro/`（`pro/scripts`、`pro/templates`、`pro/schemas`、`pro/zh-CN.md`、`pro/README.md`）
- 原技能目录 `~/.workbuddy/skills/docx-pro/` **已从本机移除**；`pro/` 是唯一留存副本，本技能不依赖任何外部技能即可独立运行。

## 1. 能力路由：什么时候用哪条链路

| 任务 | 走哪条链路 | 理由 |
|---|---|---|
| 整篇规范化（样式 / 多级编号 / 模板导入 / 域 / 交叉引用 / 图文块） | **Node 引擎**（`scripts/apply.js`） | 零外部依赖，能力最强，原子化配置 |
| 定位、查改、样式指纹、结构审计 | **Node 引擎**（`overview` / `inspect_*` / `find_*`） | 与 apply 的 locator 体系同源 |
| 生成带修订痕迹的编辑 | **Node 引擎**（`edits[]` 的 tracked changes） | 已有支持，且校验内建 |
| **添加批注 / 回复批注** | **PRO**（`pro/scripts/comment.py`） | Node 引擎无批注写入能力 |
| **接受全部修订、产出干净副本** | **PRO**（`pro/scripts/accept_changes.py`） | Node 引擎只产生修订，不清理修订 |
| **合并碎片化 run**（外部文档、编辑前预处理） | **Node 优先**（`scripts/merge_runs.js`；或 apply 配置里 `mergeRuns: true` / `--merge-runs`） | 让文本在 XML 中重新连续可查找，主链路不必切到 Python；Node 版还能覆盖页眉 / 页脚 / 脚尾注 |
| 同上，但文档已解包成目录 | **PRO**（`pro/scripts/merge_runs.py unpacked/`） | Node 版只接受 `.docx` / `.dotx` 整包 |
| **旧版 `.doc` 转换** | **PRO**（`pro/scripts/soffice.py --headless --convert-to docx file.doc`） | Node 引擎不认识二进制 `.doc` |
| **XSD 校验 + 修订完整性核查（`--author`）** | **PRO**（`pro/scripts/validate.py --original`） | 修订漏包 `<w:ins>`/`<w:del>` 时能报出，Node 版不做此项 |
| **渲染成图人工核对版式** | **PRO**（`soffice.py` + `pdftoppm`） | 视觉验证是最终把关手段 |
| **查元素的 XSD 子元素顺序** | **PRO**（`pro/scripts/element_order.py`） | Node 引擎不提供顺序查询；`references/element-order.md` 是它的生成产物 |
| 用 docx-js（npm `docx`）从零建文档 | **PRO**（`pro/zh-CN.md` 陷阱清单） | Node 引擎走手写 OOXML，无 docx-js 路径 |
| 纯内容读取/抽取 | **Node 引擎**（`overview` / `inspect_range`） | 无需 pandoc |

一句话原则：**整篇重塑走 Node，批注 / 修订清理 / 渲染验证走 PRO；run 合并已归 Node。**

## 2. 环境状态与依赖

当前机器（Windows）实测：

| 依赖 | 状态 |
|---|---|
| Python 3.13.12（托管） | ✅ 已装 |
| `defusedxml` 0.7.1 | ✅ 已装入 `C:\Users\PC\.workbuddy\binaries\python\envs\default` |
| `lxml` 6.1.1 | ✅ 已有（同环境） |
| ECMA / ISO-IEC29500 XSD schema | ✅ 已补齐（含 `pro/schemas/ecma/fouth-edition/opc-*.xsd`，补自本技能原有 schema 集） |
| `pandoc` | ❌ 未装 |
| LibreOffice（`soffice`） | ❌ 未装 |
| `pdftoppm`（Poppler） | ❌ 未装 |
| npm `docx` | ❌ 未装（在 `binaries/node/workspace` 下） |

缺哪个补哪个，命令示例：

```bash
# npm docx（放进隔离工作区）
cd C:/Users/PC/.workbuddy/binaries/node/workspace && npm install docx

# Python 侧（若换环境重装）
C:/Users/PC/.workbuddy/binaries/python/envs/default/Scripts/pip.exe install defusedxml lxml

# LibreOffice / Pandoc / Poppler：Windows 需手动安装（winget 或官网安装包）
winget install TheDocumentFoundation.LibreOffice
winget install JohnMacFarlane.Pandoc
winget install oschwartz10612.Poppler
```

**在 `soffice` / `pdftoppm` / `pandoc` 装好之前，渲染验证、`.doc` 转换、docx-js 创建这三条路径不可用**；此时用 Node 引擎的 `validate.js` 做结构校验即可。依赖缺失时脚本会给可读提示（含上面这组安装命令），不再抛栈回溯：

```bash
# 一条命令自检：soffice 是否可用
python pro/scripts/soffice.py --check    # exit 0 = OK（打印路径）；exit 1 = 未装（打印安装指引）
```

## 3. 调用约定

- 解释器：优先 `C:\Users\PC\.workbuddy\binaries\python\envs\default\Scripts\python.exe`（已带 `defusedxml` + `lxml`）。
- 脚本内部用相对导入（`from helpers import ...`），**必须在 `pro/scripts/` 目录下执行**，或把该目录加进 `PYTHONPATH`：

```bash
cd C:/Users/PC/.workbuddy/skills/docx-master/pro/scripts
python comment.py ../../out.docx "这条金额上限偏低"
```

- 路径解析：脚本内的 `schemas/`、`templates/` 均相对 `pro/` 定位，移动 `pro/` 目录时三个子目录须整体移动。

## 4. 常用工作流

### 4.1 添加批注

```bash
# 直接对 .docx（自动写入 comments.xml / commentsExtended.xml / commentsIds.xml / commentsExtensible.xml + 关系 + content-type）
python comment.py contract.docx "This cap is too low" -o annotated.docx

# 对已解包目录（接下来还要改 document.xml 时更省事）
python comment.py unpacked/ "Fees & expenses cap is too low"
python comment.py unpacked/ "Agreed" --parent 0      # 回复第 0 条
```

脚本会打印需要放进 `word/document.xml` 的 `<w:commentRangeStart>` / `<w:commentRangeEnd>` / `<w:commentReference>` 片段——**不放置这些锚点，批注写入了但看不见**。

### 4.2 接受修订 / 清理

```bash
python accept_changes.py in.docx out.docx
```

注意：接受「被删除的段落标记」本应让该段与下一段合并，Word 能正确做，基于 LibreOffice 的 `accept_changes.py` 在「被删段落后紧跟空段落」时会留下空段落（自动编号时表现为多余空列表项）。渲染视图中的空列表项是视图产物，务必回 XML 确认。若只做校验，用 `pandoc --track-changes=accept` 也不合并段落。

### 4.3 编辑前预处理（合并 run）

```bash
unzip -q doc.docx -d unpacked/
find unpacked -type l -delete        # 外部 docx 不可信：删掉符号链接条目
python merge_runs.py unpacked/       # 或 python merge_runs.py doc.docx -o merged.docx
```

Word 因修订 ID、拼写检查标记会把一句话拆成多个 `<w:r>`；不合并则「肉眼可见的字符串」在 XML 中并不连续，查找/替换会漏。

### 4.4 校验（含修订核查）

```bash
python validate.py out.docx --original src.docx                    # 只看被本次引入的新错误
python validate.py out.docx --original src.docx --author "佐罗斯"   # 逐处核查是否都包了 w:ins / w:del
python validate.py out.docx --auto-repair                          # 自动修常见问题
```

### 4.5 渲染验证

```bash
python soffice.py --headless --convert-to pdf output.docx
pdftoppm -jpeg -r 100 output.pdf page
ls page-*.jpg        # 页码按总页数补零：page-01.jpg … page-12.jpg，然后逐张 Read
```

### 4.6 docx-js 创建路径的陷阱

建新文档时若走 `docx`（npm），必读 `pro/zh-CN.md` 的陷阱清单：页面尺寸默认 A4、表格需 table `columnWidths` + 每格 `width` 双设且用 `WidthType.DXA`、底纹用 `ShadingType.CLEAR`、列表用 `numbering`、`ImageRun` 必带 `type`、TOC 需内置 `HeadingLevel` 或设 `outlineLevel`、同行右对齐页码用 `PositionalTab`。

### 4.7 查元素的 XSD 子元素顺序

```bash
python element_order.py --type CT_SectPr CT_PPr   # 查指定类型的子元素顺序
python element_order.py --list                    # 列出全部已知类型
python element_order.py --doc                     # 重新生成 references/element-order.md
```

`references/element-order.md` 就是 `--doc` 的产物（权威、自动生成），**手工改会被下一次重生成覆盖**。写原始片段前不确定某个元素该插在哪，先查它。

## 5. 边界与冲突

1. **不要用 PRO 做整篇规范化**：它没有样式/编号/指纹体系，会与 Node 引擎的 `apply` 产物互相覆盖。
2. **两套 validate 的取舍**：Node `validate.js`（基线差分、与 apply 同源）用于日常；PRO `validate.py` 只在需要 `--author` 修订核查或 `--auto-repair` 时用。
3. **批注属于 Node 引擎的 blocker**：`inspect_blockers` 会拒绝编辑带批注的段落；先决定批注与编辑的先后——需要两者时，先改结构（Node），再加批注（PRO）。
4. **文件不可变原则一致**：两条链路都写新文件，原文件不动。
5. **`pro/scripts/` 里的 `pptx_*.py` 与本技能无关**，属原技能自带，忽略即可。

## 6. 安全审计结论（2026-10-10，逐文件人工审计）

原始 `docx-pro` 为市场技能（官方市场来源，带 `_skillhub_meta.json`）。审计发现：

- **无网络出站**：无 `requests` / `urllib` 出网请求，出现的 `http(s)://` 全部是 OOXML 命名空间常量。
- **无凭据读取、无环境变量泄露、无 eval/exec 动态执行**。
- **外部进程调用仅限文档工具链**：`subprocess` 只用于 `soffice` / `pandoc` / `pdftoppm`；被调用的可执行名是写死的，不接受用户输入拼接。
- **文件删除仅限临时文件**：`unlink()` 只删自身创建的临时文件；解包流程中的 `find unpacked -type l -delete` 是**防护措施**（剥离外部文档里的符号链接条目）。
- **需知晓的一点（P1 级别提示，非 P0）**：`soffice.py` 在检测到沙箱禁止 Unix socket 时，会现场编译一个 C shim 并通过 `LD_PRELOAD` 注入 LibreOffice，用于绕过 socket 限制。该逻辑**仅 Linux 生效**（依赖 `AF_UNIX` 与 `.so`），Windows 上走不到这条分支。若将来在 Linux 环境使用本技能，需知悉此行为。
- 结论：**P2（可安全使用）**，附带上述 Linux shim 说明。

## 7. 回退

整合前的 `docx-master` 已完整备份：

```
C:\Users\PC\.workbuddy\skills-backups\docx-master-backup-2026-10-10          （pro 整合前）
C:\Users\PC\.workbuddy\skills-backups\docx-master-before-optimize-2026-10-10 （能力优化前）
```

需要回退时，删除本技能目录下的 `pro/` 与 `references/pro-toolchain.md`，并把 `SKILL.md` 恢复为备份版本即可（备份目录内为整合前状态）。
