# PRO 工具链（可选 Python 扩展层）

`pro/` 是 `docx-master` 的**可选**扩展，不是独立技能。默认链路始终是 Node 引擎
（`scripts/apply.js` + `inspect_*` / `find_*`，零外部依赖）；只有 Node 引擎唯一不覆盖的场景才走这里。

装不装 `pro/` 的 Python 依赖，只影响下列少数路径；核心写入与结构校验不受影响。

## 1. 能力路由：什么时候用哪条链路

| 任务 | 走哪条链路 | 理由 |
|---|---|---|
| 整篇规范化（样式 / 多级编号 / 模板导入 / 域 / 交叉引用 / 图文块） | **Node 引擎**（`scripts/apply.js`） | 零外部依赖，能力最强，原子化配置 |
| 定位、查改、样式指纹、结构审计 | **Node 引擎**（`overview` / `inspect_*` / `find_*`） | 与 apply 的 locator 体系同源 |
| 生成带修订痕迹的编辑 | **Node 引擎**（`edits[]` 的 tracked changes） | 已有支持，且校验内建 |
| **添加批注 / 回复批注** | **PRO**（`scripts/comment.py`） | Node 引擎无批注写入能力 |
| **接受全部修订、产出干净副本** | **PRO**（`scripts/accept_changes.py`） | Node 引擎只产生修订，不清理修订 |
| **合并碎片化 run** | **Node 优先**（`scripts/merge_runs.js`；或 `mergeRuns: true`） | 主链路不必切 Python；Node 版还能覆盖页眉 / 页脚 / 脚尾注 |
| 同上，但已解包成目录 | **PRO**（`scripts/merge_runs.py unpacked/`） | Node 版只接受 `.docx` / `.dotx` 整包 |
| **旧版 `.doc` 转换** | **PRO**（`scripts/soffice.py` 或 `scripts/render.py`） | Node 引擎不认识二进制 `.doc` |
| **XSD 校验 + 修订完整性核查（`--author`）** | **PRO**（`scripts/validate.py --original`） | 修订漏包 `<w:ins>`/`<w:del>` 时能报出，Node 版不做此项 |
| **渲染成图，人工核对版式** | **PRO**（`scripts/render.py`） | 视觉核对是发现版式事故的唯一手段 |
| **查元素的 XSD 子元素顺序** | **PRO**（`scripts/element_order.py`） | Node 引擎不提供顺序查询 |
| 用 docx-js（npm `docx`）从零建文档 | **PRO**（`zh-CN.md` 陷阱清单） | Node 引擎走手写 OOXML，无 docx-js 路径 |
| 纯内容读取 / 抽取 | **Node 引擎**（`overview` / `inspect_range`） | 无需 pandoc |

一句话原则：**整篇重塑走 Node；批注、修订清理、视觉核对走 PRO；run 合并已归 Node。**

## 2. 依赖

三个可选项，各自独立。缺哪个只断对应的那条路径。

| 依赖 | 用来做什么 | 缺失时受影响的路径 |
|---|---|---|
| Python 3.9+ 与 `lxml`、`defusedxml` | 全部 `pro/scripts/*.py` | 批注、修订清理、PRO 校验、渲染 |
| LibreOffice（`soffice`） | `.doc` 转换；PDF 渲染器之一 | `.doc` 转换；若同时无 Office/WPS 则视觉核对不可用 |
| Microsoft Office 或 WPS Office | PDF 渲染器之一（Windows COM / macOS AppleScript） | 无 soffice 时视觉核对不可用 |
| `pymupdf` / `pdftoppm` / ImageMagick | PDF → 图片 | 只能拿到 PDF，需自己打开看 |
| `pandoc` | 外来格式导入（md / html / odt → docx）；修订对照 | 该导入路径 |
| npm `docx` | docx-js 建文档路径 | 建文档只能走 Node 引擎（推荐路径） |

安装示例（按需取一条）：

```bash
# Python 侧
pip install lxml defusedxml pymupdf

# LibreOffice：Windows
winget install TheDocumentFoundation.LibreOffice
# LibreOffice：macOS
brew install --cask libreoffice
# LibreOffice：Debian / Ubuntu
sudo apt install libreoffice

# pandoc
winget install JohnMacFarlane.Pandoc     # Windows
brew install pandoc                       # macOS
sudo apt install pandoc                   # Debian / Ubuntu

# poppler（pdftoppm）
winget install oschwartz10612.Poppler    # Windows
brew install poppler                      # macOS
sudo apt install poppler-utils            # Debian / Ubuntu
```

先跑一次探测，确认这台机器实际有什么：

```bash
node scripts/doctor.js                                  # 全能力矩阵
python pro/scripts/render.py --probe --human             # 只看渲染相关，带解析身份
```

## 3. 调用约定

- **解释器**：用你自己环境里的 `python` / `python3`，或显式传 `node scripts/doctor.js --python <exe>`。
  本技能不假设任何固定安装路径。
- **工作目录**：脚本内部用相对导入（`from helpers import ...`），
  **必须在 `pro/scripts/` 目录下执行**，或把该目录加进 `PYTHONPATH`：

```bash
cd pro/scripts
python comment.py ../../out.docx "这条金额上限偏低"
```

- **路径解析**：脚本内的 `schemas/`、`templates/` 均相对 `pro/` 定位，移动 `pro/` 时须整体移动。

## 4. 常用工作流

### 4.1 添加批注

```bash
# 直接对 .docx（自动写入 comments*.xml + 关系 + content-type）
python comment.py contract.docx "This cap is too low" -o annotated.docx

# 对已解包目录（接下来还要改 document.xml 时更省事）
python comment.py unpacked/ "Fees & expenses cap is too low"
python comment.py unpacked/ "Agreed" --parent 0      # 回复第 0 条
```

脚本会打印需要放进 `word/document.xml` 的 `<w:commentRangeStart>` / `<w:commentRangeEnd>` /
`<w:commentReference>` 片段——**不放置这些锚点，批注写入了但看不见**。

### 4.2 接受修订 / 清理

```bash
python accept_changes.py in.docx out.docx
```

注意：接受「被删除的段落标记」本应让该段与下一段合并。Word 能正确处理；
基于 LibreOffice 的 `accept_changes.py` 在「被删段落后紧跟空段落」时会留下空段落
（自动编号时表现为多余空列表项）。渲染视图中的空列表项是视图产物，务必回 XML 确认。

### 4.3 编辑前预处理（合并 run）

外部 `.docx` 不可信，解包时先剥离符号链接条目。命令按平台选一条：

```bash
# Linux / macOS
unzip -q doc.docx -d unpacked/
find unpacked -type l -delete

# Windows PowerShell
Expand-Archive doc.docx -DestinationPath unpacked
Get-ChildItem unpacked -Recurse -Attributes ReparsePoint | Remove-Item -Force
```

```bash
python merge_runs.py unpacked/       # 或 python merge_runs.py doc.docx -o merged.docx
```

Word 会因修订 ID、拼写检查标记把一句话拆成多个 `<w:r>`；不合并则「肉眼可见的字符串」
在 XML 中并不连续，查找 / 替换会漏。

### 4.4 校验（含修订核查）

```bash
python validate.py out.docx --original src.docx                    # 只看本次引入的新错误
python validate.py out.docx --original src.docx --author "<你的署名>"   # 逐处核查是否都包了 w:ins / w:del
python validate.py out.docx --auto-repair                          # 自动修常见问题
```

### 4.5 渲染成图（视觉核对）

`render.py` 运行期探测可用渲染器，**不假设任何办公套件**：

```
PDF 渲染器降级链
├─ 1. soffice        LibreOffice，跨平台、可复现、无 GUI 副作用
├─ 2. Windows COM     Word.Application → KWPS.Application（未装 soffice 时）
├─ 3. macOS Word     AppleScript 驱动 Word.app（未装 soffice 时）
└─ 4. 全不可用        exit 3，打印安装指引；不产出「假通过」

PDF → 图片降级链
├─ 1. pymupdf（pip，跨平台）
├─ 2. pdftoppm（poppler）
├─ 3. ImageMagick（需 Ghostscript）
└─ 4. 全不可用        保留 PDF，由使用者自行打开
```

```bash
python render.py REPORT.docx                        # 全页 PNG + PDF
python render.py REPORT.docx --pages 1,6-8 --dpi 200
python render.py REPORT.docx --out ./preview --pdf-only
python render.py --probe --human                    # 只看这台机器有什么、解析出的身份是什么
```

输出 JSON 含 `renderer`、`renderer_identity`、`image_engine`、`images[]`，可直接读图核对。

> **渲染器身份必须核对，不能只看 `Name`。** WPS 把自己注册在 `Word.Application` 这个 ProgID 下，
> 并让 `Application.Name` 返回 `"Microsoft Word"`。`render.py` 因此用 `Application.Path`
> （再以 `Build` 号兜底）判定真身，输出里会如实写 `wps` 或 `word`。

### 4.6 docx-js 创建路径的陷阱

建新文档若走 npm `docx`，必读 `zh-CN.md` 的陷阱清单：页面尺寸默认 A4、表格需
table `columnWidths` + 每格 `width` 双设且用 `WidthType.DXA`、底纹用 `ShadingType.CLEAR`、
列表用 `numbering`、`ImageRun` 必带 `type`、TOC 需内置 `HeadingLevel` 或设 `outlineLevel`、
同行右对齐页码用 `PositionalTab`。

### 4.7 查元素的 XSD 子元素顺序

```bash
python element_order.py --type CT_SectPr CT_PPr   # 查指定类型的子元素顺序
python element_order.py --list                    # 列出全部已知类型
python element_order.py --doc                     # 重新生成 references/element-order.md
```

`references/element-order.md` 是 `--doc` 的产物（自动生成），**手工改会被下一次重生成覆盖**。

## 5. 边界与冲突

1. **不要用 PRO 做整篇规范化**：它没有样式 / 编号 / 指纹体系，会与 Node 引擎的 `apply` 产物互相覆盖。
2. **两套 validate 的取舍**：Node `validate.js`（基线差分、与 apply 同源）用于日常；
   PRO `validate.py` 只在需要 `--author` 修订核查或 `--auto-repair` 时用。
3. **批注属于 Node 引擎的 blocker**：`inspect_blockers` 会拒绝编辑带批注的段落；
   需要两者时先改结构（Node），再加批注（PRO）。
4. **文件不可变原则一致**：两条链路都写新文件，原文件不动。
5. **视觉核对是 best-effort，不是判决**：结论只对产出它的渲染器成立（见 §6）。
6. **`scripts/pptx_*.py` 与本技能无关**，属上游工具链残留，忽略即可。

## 6. 视觉核对的效力边界

这一节比代码重要，出问题的地方通常不是脚本而是结论被高估。

| 环节 | 验什么 | 结论强度 |
|---|---|---|
| `verify.js` 断言 + `validate.js` | 样式、编号、域、书签、XML 合法性 | **客观**，任何人跑结果一致 |
| `render.py` 视觉核对 | 分页、页眉页脚、表格观感、图文相对位置 | **主观**，随渲染器变化 |
| 使用者在自己的软件里打开 | 真实使用场景 | **唯一权威** |

- WPS 渲染通过 ≠ Word 渲染正确，反之亦然。字体度量、换行算法、浮动对象处理都不同。
- 报告必须记录渲染器身份（`wps 12.1.0.29690` / `soffice 26.8` / `Word 16.0`），否则结论无法追溯。
- 交付前，用**对方实际会用的软件**复核，不要用本机可用的那个代替。

## 7. 安全声明

- **无网络出站**：无 `requests` / `urllib` 出网请求；出现的 `http(s)://` 全部是 OOXML 命名空间常量。
- **无凭据读取、无环境变量泄露、无 `eval` / `exec` 动态执行**。
- **外部进程调用仅限文档工具链**：`soffice` / `python` / `powershell` / `osascript` / `pandoc` /
  `pdftoppm` / `magick` / `tasklist` / `taskkill`。可执行名写死，不接受用户输入拼接。
- **COM 渲染的清理边界**：`render.py` 渲染前后各取一次 Word / WPS 进程快照，
  超时只回收**本次新增**的进程，不触碰使用者已打开的文档。
- **文件删除仅限临时文件**：`unlink()` 只删自身创建的临时文件；
  解包流程中的符号链接剥离是防护措施，不是清理行为。
- **`soffice.py` 的 Linux shim**：检测到沙箱禁止 Unix socket 时，会现场编译 C shim 并通过
  `LD_PRELOAD` 注入 LibreOffice 以绕过 socket 限制。该分支**仅 Linux 生效**（依赖 `AF_UNIX`），
  Windows / macOS 走不到。在受限 Linux 环境使用前需知悉此行为。

## 8. 来源与授权

`pro/` 的早期版本衍生自技能市场中的 `docx-pro`（作者 `zxh`），此后由本仓库独立维护与扩展。
该部分代码的授权以原技能声明为准；如权利人认为再分发方式不当，请提 issue 联系移除。

`pro/schemas/` 为 ECMA-376 / ISO-IEC 29500 规范附带的 XSD 文件，按原样再分发、未做修改，
详见仓库根目录 `THIRD-PARTY-NOTICES.md`。
