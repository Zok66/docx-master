# 工具参考（全部 CLI）

主控 [`SKILL.md`](../SKILL.md) 只保留「常用工具索引 + 何时用」；本文件是**完整清单与语义细节**，
按需查阅即可，不必随主控一起注入。

全部工具通过 `node <script> <args>` 调用，输出到 stdout。脚本位于技能根的 `scripts/`。
下面每条描述都是能力契约，措辞与主控同源——**改动本文件时同步核对 SKILL.md 的索引行**。

---

## 侦察 / 读取类

| 工具 | 调用 | 何时用 |
|------|------|--------|
| `overview` | `node scripts/overview.js <file> [--paras=A..B\|none] [--include-unused]` | 第一调用。Metadata, page setup, theme, style defs, numbering schemes (with `used by N paragraphs` count — 0 flags orphans), visual style statistics with **direct-format per fingerprint** (drives "don't redeclare" rule), document skeleton. `--paras` slices or drops the skeleton; `--include-unused` shows usage=0 styles. |
| `inspect_range` | `node scripts/inspect_range.js <file> <from> <to>` | Full text and computed styles for a paragraph range. |
| `inspect_runs` | `node scripts/inspect_runs.js <file> <para>` | Per-run rPr dump. Use for paragraphs with mixed run-level formatting OR **form-fill segments** (label + underscore-blank pattern). |
| `inspect_neighbors` | `node scripts/inspect_neighbors.js <file> <para> [--radius N]` | What surrounds a paragraph. First choice for figure-caption / table-caption / first-after-heading classification. |
| `inspect_style` | `node scripts/inspect_style.js <file> <fingerprint>` | What role a fingerprint plays. `<fingerprint>` accepts letter label or 6-char hash. |
| `inspect_style_def` | `node scripts/inspect_style_def.js <file> <styleId>` | Pre-defined styles in `styles.xml` and their `basedOn` chain. Use before reusing or overriding an existing styleId. |
| `inspect_section` | `node scripts/inspect_section.js <file> <index>` | Page setup differences between sections. |
| `inspect_table` | `node scripts/inspect_table.js <file>` | Top-level tables with `[row,col]` cell snippets and per-cell paragraph-index spans. Use before composing a `cell` locator or a `range` touching table content. |
| `inspect_blockers` | `node scripts/inspect_blockers.js <file>` | Paragraphs `apply`'s edit phase will refuse — tracked changes, complex fields, SDT controls. |
| `inspect_bookmarks` | `node scripts/inspect_bookmarks.js <file> [--json] [--problems-only] [--exit-zero] [--limit N]` | Bookmarks, every reference that names one (REF / PAGEREF / `HYPERLINK \l` / `TOC \b` / `w:anchor`), and the ones that resolve to nothing. Non-zero exit on a dangling reference or a malformed bookmark (`--exit-zero` 只报告不改退出码)。跨部件解析：页眉里的 REF 可以指向正文书签。 |
| `inspect_caption` | `node scripts/inspect_caption.js <file> [identifier]` | SEQ-based captions in the document. See [captions.md](captions.md). |
| `migrate_captions` | `node scripts/migrate_captions.js <file> [--style <styleId>]...` | Read-only detector for manually-numbered caption-shaped paragraphs (`图 2.1` typed, no SEQ). Agent uses the output to write apply config converting them to `CaptionBlock`. |
| `find_paragraphs` | `node scripts/find_paragraphs.js <file> --regex <pat> [--flags <flags>] [--limit N] [--fingerprint X]` | Cross-document paragraph regex. Validate `pattern_rules` coverage before applying. |
| `find_text` | `node scripts/find_text.js <file> <pattern> [--regex] [--paragraph N \| --range A-B] [--limit N] [--context N]` | Character-level locator. Returns paragraph index, run index, char offset, length, structural-region annotations. Pair with `inspect_runs` for rPr or with `set-run` to replace at `runIndex`. |
| `merge_runs` | `node scripts/merge_runs.js <file.docx> [-o out.docx] [--parts headers,footers,notes,all] [--dry-run]` | Coalesce adjacent identically-formatted runs (also strips `rsid*` attributes and `<w:proofErr>` markers). **Run before a surgical edit on a document of unknown provenance** — Word fragments a paragraph's text across many runs, so a phrase visible on the page is often not one contiguous string in the XML and `find_text` / `set-run` miss it. Rendering unchanged, revision wrappers preserved. |

## 写入类

| 工具 | 调用 | 何时用 |
|------|------|--------|
| `apply --dry-run` | `node scripts/apply.js --dry-run <config.json>` | Iterate on a config without writing. **Use between every config edit.** |
| `apply` | `node scripts/apply.js <config.json>` | Unified writer. Default for any write task. 顺序见 [`SKILL.md`](../SKILL.md) `apply` pipeline order。 |
| `batch_fill` | `node scripts/batch_fill.js <template.docx> <data.csv\|data.json> --out-dir <dir> [--name "{{字段}}.docx"] [--name-field <col>] [--parts all] [--delims "{{,}}"] [--limit N] [--strict] [--dry-run] [--validate]` | Mail merge without Word: one template + one data table → one document per row. 语义细节见下节。 |

## 校验 / 编排 / 交付类

| 工具 | 调用 | 何时用 |
|------|------|--------|
| `validate` | `node scripts/validate.js <file>` | Schema-aware OOXML check. `apply` runs this automatically; standalone for spot-checking arbitrary .docx files. |
| `pipeline` | `node scripts/pipeline.js <build.json> [--plan] [--emit-script [out.sh]] [--dry-run] [--no-auto-validate] [--python <exe>]` | Runs a declared multi-step chain in order with one exit code, aborting on the first failure (`set -e` semantics when emitted as a script). Auto-appends schema validation when the manifest declares none. `--plan` prints the chain; `--emit-script` writes it as a bash script — the fallback when child processes are unavailable. |
| `verify` | `node scripts/verify.js <verify.json \| file.docx> [--diff <baseline.docx>] [--json] [--quiet]` | The delivery gate: runs declarative assertions (non-zero exit on any failure) and, with a baseline, prints the change manifest. Zero-dependency; reads the OPC package directly. 完整说明见 [`delivery-gate.md`](delivery-gate.md)。 |
| `doctor` | `node scripts/doctor.js [--json] [--python <exe>]` | Reports which capabilities are actually available (Node engine, subprocess spawning, Python+lxml, LibreOffice, pdftoppm, pandoc, npm docx) and what each missing one disables. Always exits 0 — it is a report, not a gate. |

---

## `batch_fill` 语义

Placeholders are matched against the **merged** layout, not the file's run fragmentation: runs are
merged once across every participating part, then a placeholder still spanning a run boundary is
stitched (a boundary only survives merging when the two runs differ in `rPr`). The value lands in the
run that carried the placeholder's **first** character, so that run's `rPr` decides how the inserted
text looks; the runs it spanned are emptied, never deleted, so bookmarks / fields / comment anchors
around them keep their positions. A placeholder with no matching column is left exactly as written
(`--strict` makes that fatal instead); `{{field|fallback}}` supplies replacement text. `{{_index}}` /
`{{_row}}` 是内置行号占位符；`\n` → `<w:br/>`、`\t` → `<w:tab/>`。headers / footers / notes 与正文
一同填充（`--parts all`）。The template itself is never modified — every row produces a fresh copy.
Pre-flight reports template-vs-column mismatches（缺失列默认原样保留，`--strict` 变致命）与未使用列。
Pass `--validate` (or run `validate.js` afterwards) before delivering a batch.

---

## 输出裁剪（省 token 的推荐用法）

侦察类工具在真实文档上动辄输出几百行。**只取当下决策需要的片段**，不要整段吞进上下文：

| 需求 | 做法 |
|---|---|
| 只要骨架的某一段 | `overview --paras=120..180`；完全不要骨架用 `--paras=none` |
| 只要样式用量分布 | `overview --include-unused` 视情况开；默认已隐藏 usage=0 |
| 只要书签问题、不要全清单 | `inspect_bookmarks --problems-only` |
| 大文档的书签清单太长 | `inspect_bookmarks --limit 50` |
| 只要匹配计数 | `find_paragraphs --limit 20` / `find_text --limit 20` |
| 交给程序再加工 | 带 `--json`（`inspect_bookmarks` / `verify` / `doctor`），让上层过滤而非全文回灌 |
| 交付门断言批量核对 | `verify --quiet`（只报失败项） |

原则：**工具负责产出全量事实，裁剪发生在调用参数上**。多处需要同一事实时，一次取够、复用，
不要在多个步骤里反复重跑同一个侦察命令。
