# pro 层方法论：docx-js 创建与 pandoc 读取

> 本文件是 vendored 的 `docx-pro` 方法论原文，脚本路径已按 `docx-master` 的目录结构对齐为 `pro/scripts/`。
> 能力路由、依赖状态与调用约定见 [`../references/pro-toolchain.md`](../references/pro-toolchain.md)。

## 技能用途

全面处理 `.docx` 与 `.dotx` 文件的完整生命周期：创建专业文档（目录、页眉、表格、编号）、读取与抽取内容、查找替换、插入图片、处理修订（tracked changes）与批注、模板操作。凡交付 Word 报告类成果时必用。

核心认知：`.docx` 本质是一个 XML 文件的 ZIP 压缩包。按任务选方法：

| 任务 | 方法 |
|---|---|
| **创建**新文档 | 编写 `docx`（npm）脚本，见下文陷阱清单 |
| **编辑**已有文档 | `unzip` 解包 → 直接改 `word/document.xml` → `zip` 打包（docx-js 无法打开已有文件） |
| **读取**内容 | `pandoc -t markdown file.docx` |

下文脚本一律位于本技能的 `pro/scripts/` 下，示例统一写作 `python pro/scripts/X.py`；
若已 `cd` 到该目录，直接写脚本名即可（脚本内部用相对导入，需在该目录下执行）。

## 用 docx-js 创建文档——关键陷阱

`docx` 包已预装，**不要先跑 `npm install`**，直接写脚本 `require('docx')`；仅当 require 失败时才 `npm install docx`。

- **页面尺寸默认 A4**。要 US Letter 需显式设置 `page: { size: { width: 12240, height: 15840 } }`（DXA 单位；1440 = 1 英寸）。
- **横向页**：传纵向尺寸再加 `orientation: PageOrientation.LANDSCAPE`，docx-js 会在内部交换宽高。
- **表格必须双重设宽**：在 table 上设 `columnWidths`，同时在每个单元格上设 `width`，二者都用 `WidthType.DXA`（PERCENTAGE 在 Google Docs 中会显示异常）。各列宽度之和必须等于表格总宽。
- **表格底纹**：用 `ShadingType.CLEAR`，绝不用 `SOLID`（会渲染成纯黑）。
- **列表**：绝不直接插入 `•` 字符，用 `numbering` 配置 + `LevelFormat.BULLET`。
- **`ImageRun` 必须指定 `type:`**（`"png"`、`"jpg"` 等）。
- **`PageBreak` 必须放在 `Paragraph` 内部**。
- **绝不用 `\n`**：换行要用独立的 `Paragraph` 元素。
- **目录（TOC）**：标题必须使用内置 `HeadingLevel.*`；自定义标题样式必须设置 `outlineLevel`，否则不会出现在目录中。
- **不要用表格当水平分隔线**——改用段落的底边框。
- **点线引导 / 同行右对齐**（如目录页码）：在 `TextRun` 内使用 `PositionalTab`（`alignment: PositionalTabAlignment.RIGHT`，`leader: PositionalTabLeader.DOT`），不要用字面 `.` 或空格填充。

## 创建后必须验证

写完 `.docx` 后要渲染成图片实际查看：

```bash
python pro/scripts/soffice.py --headless --convert-to pdf output.docx
pdftoppm -jpeg -r 100 output.pdf page
ls page-*.jpg   # 然后逐张 Read 查看图片
```

`pdftoppm` 会按总页数宽度对页码补零（`page-01.jpg` … `page-12.jpg`）。

## 编辑已有文档

旧版 `.doc` 文件必须先转换：`python pro/scripts/soffice.py --headless --convert-to docx file.doc`。

标准编辑流程：

```bash
unzip -q doc.docx -d unpacked/
find unpacked -type l -delete   # 删除符号链接条目——外部来源的 docx 不可信任
python pro/scripts/merge_runs.py unpacked/   # 合并碎片化 run，使文本可被查找
# 直接编辑 unpacked/word/document.xml——不要重新格式化或美化打印
(cd unpacked && rm -f ../out.docx && zip -Xr ../out.docx .)
python pro/scripts/validate.py out.docx --original doc.docx   # XSD 校验；--auto-repair 可自动修复常见问题
# 做修订（redline）时：加 --author "<你批注所用的名字>" 检查每处编辑都已被追踪
```

**merge_runs.py**：Word 会因修订 ID、拼写检查标记等原因把文本拆散到许多 `<w:r>` run 中，导致在文档里可见的一句话在 XML 中不是连续字符串。该脚本把 `word/document.xml` 中相邻且格式完全相同的 run 合并，不改变内容和渲染效果；也可直接对 `.docx` 运行：`python pro/scripts/merge_runs.py doc.docx -o merged.docx`。

**修订（Tracked changes）规则**：
- 校验时加 `--author "<批注所用名字>"`（需配合 `--original`），它会报告任何你修改了但没有被 `<w:ins>`/`<w:del>` 包裹的文本——这种遗漏很容易发生，且在"接受修订"视图下不可见。
- 包裹 run 时需带 `w:id`、`w:author`、`w:date` 属性。
- `<w:del>` 内部文本元素是 `<w:delText>`，不是 `<w:t>`。
- 删除段落标记的写法：`<w:pPr><w:rPr><w:del w:id=".." w:author=".." w:date=".."/></w:rPr></w:pPr>`，含义是"把此段落并入下一段"。彻底删除一个段落 = 删除段落标记 + 每个.run 都用 `<w:del>` 包裹。
- `<w:del/>` 必须放在 rPr 其他子元素**之前**，元素顺序由 schema 强制约束。

**生成接受全部修订后的干净副本**：`python pro/scripts/accept_changes.py in.docx out.docx`。

接受"已删除的段落标记"本应把该段落与下一段合并，从而使所有 run 都被删除的段落整体消失——Word 能正确做到，但 `accept_changes.py` 和 `pandoc --track-changes=accept` 不总能做到，且失败方式相同：删掉文字却留下空段落，若该段原为自动编号，就会显示成一个多余的空列表项：

- `pandoc --track-changes=accept` 从不合并段落。
- `accept_changes.py`（基于 LibreOffice）通常能正确合并，唯独当被删段落后紧跟一个空的间隔段落时会失败。

注意：任一视图中的空列表项都只是该视图的渲染产物，不是文档本身的缺陷。段落删除务必回到 XML 中确认。

## 批注（Comments）

添加批注需要六个交叉关联的文件。使用 helper 脚本 **comment.py**——若接下来还要编辑 `document.xml`，用目录模式（省去一次解包/重打包）；否则直接对 `.docx` 操作：

```bash
# 对已解包目录操作（同时还要放锚点标记时首选）
python pro/scripts/comment.py unpacked/ "Fees & expenses cap is too low"
python pro/scripts/comment.py unpacked/ "Agreed" --parent 0

# 直接对 .docx 操作
python pro/scripts/comment.py contract.docx "This cap is too low" -o annotated.docx
```

脚本会自动写入 `comments.xml`、`commentsExtended.xml`、`commentsIds.xml`、`commentsExtensible.xml`、关系文件和 content-type 覆盖项，批注 ID 自动分配。随后它会打印出需要添加到 `word/document.xml` 的 `<w:commentRangeStart>`/`<w:commentRangeEnd>`/`<w:commentReference>` 代码片段，用于把批注锚定到具体文本——**在放置这些标记之前，批注虽已写入但不可见**。

## 依赖清单

`docx`（npm，已预装，仅 require 失败时安装）· `pandoc` · LibreOffice（`soffice`）· `pdftoppm`（Poppler）。

## 常见陷阱速查

- 直接 `npm install` 而没先尝试 require 预装包。
- 表格只设 `columnWidths` 不设单元格 `width`，或用 PERCENTAGE 导致 Google Docs 显示异常。
- 用 `\n` 或字面 `•` 伪造换行和列表。
- 编辑已有文档时先美化打印 XML 再打包，破坏文档结构。
- 外部 docx 解包后不删符号链接，带来安全风险。
- 修订时漏写 `<w:delText>`、漏 `w:id/w:author/w:date` 属性，或 `<w:del/>` 位置放错。
- 创建或编辑后不渲染成图片验证，直接交付。
