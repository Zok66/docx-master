---
name: docx-master
description: "Standardize, edit, or audit a Word (.docx / .dotx) document via direct OOXML mutation. Core CLIs: `apply` (the unified writer — install styles + numbering + theme + template, restyle by pattern / fingerprint, insert content via edits, all in one config), `audit` (read-only conformance check, no file written), `extend` (JSON-driven post-processing for what the writer deliberately skips — TOC / table-of-figures / index fields, footnotes & endnotes, watermarks, page borders, shapes / text boxes / canvas / WordArt / SmartArt-equivalent diagrams, charts, OLE objects, glossary building blocks, cover pages, custom XML), plus `batch_fill` (one template + a CSV / JSON data table → one finished document per row — mail merge without Word), and the delivery layer: `pipeline` (declare a multi-step chain once; one command runs it, aborting on the first failure) + `verify` (declarative delivery assertions and a change manifest — catches the defects no schema validator can see) + `doctor` (environment capability probe). Common config shapes inside `apply`: standardize (role-based whole-doc reshape — paragraph classification, named styles, multi-level auto-numbering, template import), edit (location-based surgical changes — replace/insert/delete paragraphs, table cells, image embedding, optional tracked changes). Use whenever the user wants to format / restyle / normalize / edit / audit a Word document. Also covers adding comments (批注), accepting tracked changes, legacy `.doc` → `.docx` conversion, and rendering pages for visual verification through the vendored python toolchain under `pro/` (see references/pro-toolchain.md). Do NOT use for: PDFs, spreadsheets, or plain-text / Markdown source files (unless the task is specifically to *output* a docx)."
---

# docx-master

Mutates Word (.docx / .dotx) OOXML directly: produce well-formed documents from messy templates, install styles + numbering, convert hand-typed prefixes to auto-numbering, insert content, audit conformance. Output is a new docx; the original is never touched. Legacy binary `.doc` is not understood by the Node engine — convert first via `python pro/scripts/soffice.py --headless --convert-to docx file.doc` (see references/pro-toolchain.md).

## How to think about formatting

User prompt overrides everything below.

Otherwise, treat a document as two kinds of paragraphs:

- **Content chrome** — paragraphs with typed text already in the document (chapter markers, instruction notes, fixed labels). Three operations:
  - **Preserve typography** (font / size / spacing / indent — direct `pPr` / `rPr`). Engine's selective-strip keeps direct format as long as the new style doesn't override the same attribute.
  - **Retag structure** (bind `pStyle`, strip typed structural prefixes via `pattern_rules`, install matching `numbering`). Chapter / heading / list / caption markers are structure, not typography — they belong to auto-numbering or SEQ fields. Existing typed prefixes are the **input** to retagging, not output to preserve.
  - **Don't replace, don't duplicate.** A chrome paragraph stays as itself — never the target of `replace`, never quoted into an inserted paragraph's `text`. Content insertion lands in the empty placeholder slot beside the chrome (`set-run` on inline blanks; `replace` against the empty slot for separate-paragraph forms).
- **Empty slots** — blank paragraphs pre-allocated for content insertion. Bind inserted content to a semantic style (`BodyText` / `ListNumber` / `Heading1..N`), installing the style in `styles[]` if no existing one fits. `Normal` is fallback, not a body style.

**For `styles[]`**: one style per semantic role. If the role has any paragraph already playing it in the source, the entry MUST use `fromParagraph: N` — do not redeclare top-level typography fields on represented roles (silent template override). Full rule and rationale: [standardize.md §1](references/standardize.md).

Tools surface visible facts; classification and judgment are yours.

**Standardize-shape blocks install the canonical anchors every other op binds to** — `edits[]` `styleId` + MDF, `pattern_rules` fingerprints, `InlineRef` / STYLEREF resolution. Every paragraph the agent writes is canonical: typed structural signals (heading / list / caption prefixes, in-prose counters) never appear as literal text in the emit; install the matching anchor (`numbering` / `captions` / bookmark) and let it render the signal. Source chrome the task doesn't restructure stays as-is. User-supplied content that arrives with typed prefixes is input shape, not a directive to preserve them — strip the prefix and emit each typed-prefix line as a **separate** paragraph bound to a heading / list / caption style, unless the user explicitly requested literal typed form.

## Target state

A well-formed Word document expresses structure through **styles + numbering + sections + SEQ fields**, not typed text mimicking structure. When designing the reshape:

- Every paragraph carries a semantic styleId (Heading1..N / BodyText / ListNumber / Caption / etc.); direct paragraph format only as one-off exceptions.
- Numbering: one unified multi-level scheme bound to heading styles; separate single-level schemes per list-bound style. Auto-numbering markers come from the scheme, never typed in `text`.
- Captions (figures / tables / equations) use top-level `captions` config + `CaptionBlock` / `EquationBlock` — SEQ + STYLEREF fields, not `pattern_rules`. **Placement: figure / equation captions go below; table captions go above.** See [`references/captions.md`](references/captions.md).
- Cross-references (any cite to an auto-numbered target — figures, tables, sections, equations, reference entries) use `InlineRef` nodes in `edits[]`, never typed counters. Literal counters silently desync when a target moves; REF fields stay correct. Page-number cites out of scope. See [`references/cross-references.md`](references/cross-references.md).
- Heading levels nest without skipping.
- Match content shape to slot shape — prose to body, list to `ListNumber` / `ListBullet`, figure / table / equation to their `*Block` types with `captionId`. Per-block schemas in [`references/edit.md`](references/edit.md).
- Locale: CJK prose body and list items get a 2-char first-line indent (`firstLineIndent: "2char"`). No typed spaces at CJK ↔ Latin / digit boundaries; Word's autoSpace handles them. Chinese font sizes: [`references/chinese-font-sizes.md`](references/chinese-font-sizes.md).
- CJK/Latin spacing audit: every write reports advisory Pangu findings (missing or redundant space at CJK ↔ Latin / digit boundaries) in `lint.pangu` — in `edits[]` text and `headerFooter` text. **Advisory only, never blocks the write.** Deliberate spacing → silence it with `lint.panguSpacing: false`.

## Commands

| Command | When | Reference |
|---|---|---|
| **`apply`** | The unified writer. Install styles + numbering + theme + template, restyle by pattern / fingerprint, insert content via edits — single config, single call. | [standardize.md](references/standardize.md) + [edit.md](references/edit.md) |
| `audit` | Read-only conformance check workflow. No CLI; uses the inspect tools to produce a violation report. | [audit.md](references/audit.md) |
| **`extend`** | Post-processing for what the engine deliberately does not write: TOC / table-of-figures / index fields, footnotes & endnotes, VML watermarks, page borders & background, shapes / text boxes / canvas / WordArt / SmartArt-equivalent diagrams, charts, OLE objects, glossary (Building Blocks), cover pages, custom XML, anchored raw-fragment insertion. JSON-driven, Python + lxml. | [extend.md](references/extend.md) |
| **`pipeline`** | One-command rebuild of a multi-step chain (config source → `apply` → `extend` → validate → verify). Aborts on the first failure; auto-validates the deliverable; `--plan` / `--emit-script` for environments that cannot spawn. | [§ Delivery gate](#delivery-gate) |
| **`verify`** | The delivery gate itself: declarative assertions (exit 1 on failure) + change manifest vs. a baseline. | [§ Delivery gate](#delivery-gate) |
| `doctor` | Environment capability probe. Run first on an unfamiliar machine. | [§ Delivery gate](#delivery-gate) |

`apply` pipeline order:

```
merge runs (pre-pass — only when `mergeRuns: true` / `--merge-runs`)
  → install styles + numbering + theme + template
  → run edits (referencing just-installed styleIds)
  → re-fingerprint
  → run rules (pattern_rules / bulk_rules / assignments / exclude —
    match BOTH pre-existing chrome AND agent-inserted content uniformly)
  → validate, write
```

Need fields / notes / watermarks / drawings / charts? Run `extend` on `apply`'s output —
the two are complementary, not alternatives ([`references/extend.md`](references/extend.md)).
It re-declares root namespaces for every part it touches and keeps rendering namespaces
(`wps` / `wpg` / `wpc`) out of `mc:Ignorable`; ordering and part registration are handled
for you ([`references/element-order.md`](references/element-order.md),
[`references/ooxml-extension-points.md`](references/ooxml-extension-points.md)).

Sparse by design — only declared blocks apply; untouched styles / numbering / paragraphs / theme stay as they are. **Declare only what's wrong, missing, or what the user explicitly asks to change.**

**Creating from scratch.** Omit `source` to scaffold from the bundled blank template (one empty Normal paragraph, A4 portrait, no other styles or numbering). Declare what you need in the usual blocks (`styles` / `numbering` / `pageSetup` / `headerFooter` / `edits`); `styles[]` must use Mode B (direct fields) since the blank has no representative paragraphs to extract from. The empty paragraph at index 1 is a `replace` or `insert-after` target (`insert-before` creates a stray leading empty paragraph; `delete` would empty the body). Incompatible with `template` (transplanting from a template into a blank is conceptually inconsistent — start with a host source if you need template-import).

**Documents of unknown provenance.** Word-authored files — and anything round-tripped through another tool — fragment a paragraph's text across many runs (revision ids, spell-check markers), so a phrase visible on the page is often not one contiguous string in the XML. Text-matched anchors and `set-run` then silently miss it. Two ways to fix that before you design anything:

- `node scripts/merge_runs.js <file>` — standalone pre-pass; rewrites the file (or `-o` elsewhere).
- `mergeRuns: true` in the config (CLI: `--merge-runs`) — `apply` merges before it parses anything, so every index you write refers to the merged layout.

Rendering is unchanged either way, and runs inside different `<w:ins>` / `<w:del>` wrappers are never joined.

## Workflow

1. **Survey.** `overview` first. From the output, note:
   - **Existing structure** — Heading levels, numbering schemes (with consumption counts; 0 = orphan), fingerprints, document skeleton.
   - **Content chrome formatting** — what direct `pPr` / `rPr` content paragraphs carry. This IS the document's typographic convention; preserve it by not redeclaring those attributes on the re-tagging styles.
   - **Typed structural prefixes** — heading shapes (`第N章`, `一、`, `1.1`, `（一）`) feed `pattern_rules`; caption shapes (`图 2.1`) and in-prose counters (`如图 3.2`) feed caption / cross-ref migration via `edits[]` (different mechanism).
   - **Form-fill paragraphs** — text shaped like `label + whitespace gap` or `label + ____ underscore placeholder`. The blank is usually a separate run with `<w:u/>`. Note indices for Step 2.
   - **Source content** (fill tasks) — the structural outline of what will land: which semantic roles appear (Heading 1..N, BodyText, ListNumber/Bullet, Caption, inline cross-refs), heading depth, list / figure / table / equation presence.
2. **Design ONE config.** Plan styles first based on the content's structural outline; route in one config. Six config blocks (`styles` / `numbering` / `pattern_rules` / `bulk_rules` / `assignments` / `exclude` / `edits`) detailed in [standardize.md §1–§3](references/standardize.md) and [edit.md](references/edit.md). Reactive additions accrete debt.
3. **Dry-run.** `apply --dry-run` first. The change report's signals (Style Resolution / Sample Affected / Implicit-keep failure / unmatched stripPrefix) — see [standardize.md §5](references/standardize.md).
4. **Apply.** Output is a fresh docx; the original is never modified.

## Asking the user

Default first. Don't ask unless one of the cases below applies:
- typed sentinel mid-prose vs. heading
- bold paragraph as sub-heading vs. emphasis
- missing source content for template slots
- unsupported structures (manual TOC body, page-number cites)
- font / spacing prompt without explicit scope (theme layer vs `Normal` cascade vs per-role `styles[]` — see [standardize.md "Font scope layers"](references/standardize.md))

## Out of scope

Surface to the user when blocked by: layout-**table** restructuring, TOC body content, page-number cites. Paragraphs *inside* layout-table cells ARE indexed and fully editable; only the table holding them is off-limits.

Footnotes / endnotes are **NOT** out of scope. What is out of scope is only their *indexing*: `apply`'s paragraph index covers `word/document.xml`, so note bodies stored in `word/footnotes.xml` / `word/endnotes.xml` never appear in the skeleton and are not reachable by `#NNN` locators. Inserting and populating them is fully supported through `extend` — `{"op": "footnote"}` / `{"op": "endnote"}` creates the part, registers the relationship and backfills the id ([`references/extend.md`](references/extend.md)).

Comments (批注) are NOT off-limits any more — add them through the PRO toolchain below (`pro/scripts/comment.py`).

Page setup (paper size, orientation, margins, columns, page-number format) and headers / footers ARE supported, including per-section overrides — see [`config-schema.md`](references/config-schema.md#page-setup) and [`references/header-footer.md`](references/header-footer.md).

## Tool Reference

All tools invoked via `node <script> <args>`, output to stdout.

| Tool | Invocation | When to Use |
|------|------------|-------------|
| `overview` | `node scripts/overview.js <file> [--paras=A..B\|none] [--include-unused]` | First call on any task. Metadata, page setup, theme, style defs, numbering schemes (with `used by N paragraphs` count — 0 flags orphans), visual style statistics with **direct-format per fingerprint** (drives "don't redeclare" rule), document skeleton. `--paras` slices or drops the skeleton; `--include-unused` shows usage=0 styles. |
| `inspect_range` | `node scripts/inspect_range.js <file> <from> <to>` | Full text and computed styles for a paragraph range. |
| `inspect_runs` | `node scripts/inspect_runs.js <file> <para>` | Per-run rPr dump. Use for paragraphs with mixed run-level formatting OR **form-fill segments** (label + underscore-blank pattern). |
| `inspect_neighbors` | `node scripts/inspect_neighbors.js <file> <para> [--radius N]` | What surrounds a paragraph. First choice for figure-caption / table-caption / first-after-heading classification. |
| `inspect_style` | `node scripts/inspect_style.js <file> <fingerprint>` | What role a fingerprint plays. `<fingerprint>` accepts letter label or 6-char hash. |
| `inspect_style_def` | `node scripts/inspect_style_def.js <file> <styleId>` | Pre-defined styles in `styles.xml` and their `basedOn` chain. Use before reusing or overriding an existing styleId. |
| `inspect_section` | `node scripts/inspect_section.js <file> <index>` | Page setup differences between sections. |
| `inspect_table` | `node scripts/inspect_table.js <file>` | Top-level tables with `[row,col]` cell snippets and per-cell paragraph-index spans. Use before composing a `cell` locator or a `range` touching table content. |
| `inspect_blockers` | `node scripts/inspect_blockers.js <file>` | Paragraphs `apply`'s edit phase will refuse — tracked changes, complex fields, SDT controls. |
| `inspect_caption` | `node scripts/inspect_caption.js <file> [identifier]` | SEQ-based captions in the document. See [captions.md](references/captions.md). |
| `migrate_captions` | `node scripts/migrate_captions.js <file> [--style <styleId>]...` | Read-only detector for manually-numbered caption-shaped paragraphs (`图 2.1` typed, no SEQ). Agent uses the output to write apply config converting them to `CaptionBlock`. |
| `find_paragraphs` | `node scripts/find_paragraphs.js <file> --regex <pat> [--flags <flags>] [--limit N] [--fingerprint X]` | Cross-document paragraph regex. Validate `pattern_rules` coverage before applying. |
| `find_text` | `node scripts/find_text.js <file> <pattern> [--regex] [--paragraph N \| --range A-B] [--limit N] [--context N]` | Character-level locator. Returns paragraph index, run index, char offset, length, structural-region annotations. Pair with `inspect_runs` for rPr or with `set-run` to replace at `runIndex`. |
| `merge_runs` | `node scripts/merge_runs.js <file.docx> [-o out.docx] [--parts headers,footers,notes,all] [--dry-run]` | Coalesce adjacent identically-formatted runs (also strips `rsid*` attributes and `<w:proofErr>` markers). **Run before a surgical edit on a document of unknown provenance** — Word fragments a paragraph's text across many runs, so a phrase visible on the page is often not one contiguous string in the XML and `find_text` / `set-run` miss it. Rendering unchanged, revision wrappers preserved. |
| `batch_fill` | `node scripts/batch_fill.js <template.docx> <data.csv\|data.json> --out-dir <dir> [--name "{{字段}}.docx"] [--name-field <col>] [--parts all] [--delims "{{,}}"] [--limit N] [--strict] [--dry-run] [--validate]` | Mail merge without Word: one template + one data table → one document per row. Placeholders `{{field}}`, `{{field\|fallback}}`, `{{_index}}` / `{{_row}}`. Stitches placeholders Word split across runs, turns `\n` into `<w:br/>` and `\t` into `<w:tab/>`, and fills headers / footers / notes as well as the body. Pre-flight reports template-vs-column mismatches (missing → left as-is, or fatal with `--strict`) and unused columns. |
| `validate` | `node scripts/validate.js <file>` | Schema-aware OOXML check. `apply` runs this automatically; standalone for spot-checking arbitrary .docx files. |
| `apply --dry-run` | `node scripts/apply.js --dry-run <config.json>` | Iterate on a config without writing. **Use between every config edit.** |
| `apply` | `node scripts/apply.js <config.json>` | Unified writer. Default for any write task. |
| `pipeline` | `node scripts/pipeline.js <build.json> [--plan] [--emit-script [out.sh]] [--dry-run] [--no-auto-validate] [--python <exe>]` | Runs a declared multi-step chain in order with one exit code, aborting on the first failure (`set -e` semantics when emitted as a script). Auto-appends schema validation when the manifest declares none. `--plan` prints the chain; `--emit-script` writes it as a bash script — the fallback when child processes are unavailable. |
| `verify` | `node scripts/verify.js <verify.json \| file.docx> [--diff <baseline.docx>] [--json] [--quiet]` | The delivery gate: runs declarative assertions (non-zero exit on any failure) and, with a baseline, prints the change manifest. Zero-dependency; reads the OPC package directly. |
| `doctor` | `node scripts/doctor.js [--json] [--python <exe>]` | Reports which capabilities are actually available (Node engine, subprocess spawning, Python+lxml, LibreOffice, pdftoppm, pandoc, npm docx) and what each missing one disables. Always exits 0 — it is a report, not a gate. |

**`batch_fill` semantics.** Placeholders are matched against the **merged** layout, not the file's run fragmentation: runs are merged once across every participating part, then a placeholder still spanning a run boundary is stitched (a boundary only survives merging when the two runs differ in `rPr`). The value lands in the run that carried the placeholder's **first** character, so that run's `rPr` decides how the inserted text looks; the runs it spanned are emptied, never deleted, so bookmarks / fields / comment anchors around them keep their positions. A placeholder with no matching column is left exactly as written (`--strict` makes that fatal instead); `{{field|fallback}}` supplies replacement text. The template itself is never modified — every row produces a fresh copy. Pass `--validate` (or run `validate.js` afterwards) before delivering a batch.

## PRO toolchain (optional extension)

The Node engine above stays the default and remains zero-dependency. Reach for the vendored Python toolchain under `pro/` (integrated from the `docx-pro` skill on 2026-10-10) only for what it uniquely covers:

| Need | Tool |
|---|---|
| Add a comment (6-file comment set + anchor markers) | `python pro/scripts/comment.py <docx\|unpacked/> "text" [--parent N] [-o out.docx]` |
| Accept all tracked changes → clean copy | `python pro/scripts/accept_changes.py in.docx out.docx` |
| Coalesce fragmented runs so text becomes findable | **Node (preferred):** `node scripts/merge_runs.js <docx\|dotx> [-o out.docx] [--parts headers,footers,notes,all] [--dry-run]` — same semantics as the Python version, and unlike it can also cover headers / footers / notes. Python fallback, for an already-unpacked directory: `python pro/scripts/merge_runs.py unpacked/` |
| Convert a legacy `.doc` (or anything LibreOffice reads) to `.docx` | `python pro/scripts/soffice.py --headless --convert-to docx file.doc` · self-check with `python pro/scripts/soffice.py --check` |
| XSD check + redlining completeness (`--author`) | `python pro/scripts/validate.py out.docx --original src.docx [--author "name"] [--auto-repair]` |
| Look up the XSD-mandated child order of any element | `python pro/scripts/element_order.py --type CT_SectPr CT_PPr` / `--list` / `--doc` (regenerates [`references/element-order.md`](references/element-order.md)) |
| Render to images and eyeball the layout | `python pro/scripts/soffice.py --headless --convert-to pdf x.docx` → `pdftoppm -jpeg -r 100 x.pdf page` → Read `page-*.jpg` |
| Create from scratch with docx-js (npm `docx`) | Footgun list in [`pro/zh-CN.md`](pro/zh-CN.md) |

Dependency state on this machine: `defusedxml` + `lxml` installed in `envs/default`; `pandoc` / LibreOffice (`soffice`) / `pdftoppm` / npm `docx` are **NOT installed** — install them before using the render-verify, `.doc` conversion, or docx-js paths, otherwise stay on the Node engine (`validate.js` covers schema checking).

Routing rules, boundary conflicts, and the full workflows: [`references/pro-toolchain.md`](references/pro-toolchain.md).

## Delivery gate

Schema validation (`validate`, PRO `validate.py`) answers one question: *is this
file well-formed?* It cannot answer the other one: *is this what the user asked
for?* A footer carrying a title nobody requested validates perfectly. Both
questions get answered before handover, so delivery is a gate, not a final
`validate` call:

| # | Step | Command | Fails when |
|---|---|---|---|
| 1 | Probe the environment | `node scripts/doctor.js` | never — it reports which optional paths (render / soffice / pandoc) are live, so you learn at minute one instead of step nine |
| 2 | Dry-run each config edit | `node scripts/apply.js --dry-run <cfg>` | the config is invalid |
| 3 | Run the whole chain | `node scripts/pipeline.js build.json` | any step fails — the chain aborts and later steps never run |
| 4 | Assert the artifact | `node scripts/verify.js verify.json` | any assertion fails |
| 5 | Read the change manifest | `node scripts/verify.js <out.docx> --diff <src.docx>` | never — but every line must be defensible against the user's words |

Steps 3 and 4 wire together: declare `"verify"` in the pipeline manifest and a
failed assertion fails the build. A build may not finish unvalidated — if no
`validate` step is declared, the pipeline appends one for the deliverable.

**When the environment forbids child processes** (sandboxed agents), step 3
becomes `--plan` (print the resolved chain) or `--emit-script` (write a `set -e`
bash script); `bash <script>` gives the same abort-on-first-failure contract in
one command, and a non-zero exit from it *is* the build result. `doctor` reports
this capability explicitly — checked before you plan a long task, not during one.

### Assertions encode intent

Write assertions before the config: they are the user's requirements stated
mechanically. Their whole purpose is the defect class no schema check can see:

```json
{ "type": "notContains", "text": "My Book", "scope": "footer" },
{ "type": "footerText", "regex": "^\\s*第\\s*<PAGE>\\s*页\\s*\\/\\s*共\\s*<NUMPAGES>\\s*页\\s*$" },
{ "type": "hasField", "field": "TOC", "scope": "body" }
```

The first line exists because a footer carrying an unrequested title is
perfectly well-formed XML that passes every validator. `<PAGE>` is the field
marker: cached field *results* are ignored, so assertions survive repagination
and don't depend on what Word last computed.

Types — counts (`sectionCount` / `paragraphCount` / `tableCount` / `imageCount`
with `eq` / `ne` / `min` / `max`), presence (`contains` / `notContains` /
`textMatches`, scoped `body` | `header` | `footer` | `all`), parts (`headerText` /
`footerText` with `regex` / `equals` / `index`), fields (`hasField` / `noField`),
setup (`pageSize` / `partExists`), styles (`styleExists` / `styleMissing` by
`styleId` or `name`), metadata (`author` / `title`).

Check styles by `name`, not `styleId`: a writer may serialize ids as opaque
numbers (`2`, `10`, …) — this skill's own `apply` does — so the id is not a
stable identity across runs.

### The manifest is read by a person

`verify --diff <baseline>` prints every structural difference the run
introduced: sections, paragraph count, parts, header/footer text, fields,
styles, samples of new body text. It exists to be read, deliberately:

> Every line above is a change this run introduced. Anything here the user did
> not ask for is a defect, not a feature.

### The last check is not a program

Before handover, restate the request as a checklist in the user's own terms and
walk the artifact against it: every requested item present, and *nothing present
that was not requested*. Self-added polish — a title in the footer, an extra
cover line, a "helpful" section — is the most common way a technically perfect
file still fails, and it is invisible to every automated check above. When the
manifest shows something you cannot trace back to the user's words, delete it.

## Cross-command invariants

- Original file is never modified; every applying CLI writes a fresh copy + validates before keeping. Validation uses baseline-diff — only errors *introduced* by the run are fatal. See [`standardize.md §Validation behavior`](references/standardize.md#validation-behavior); `--allow-validation-warnings` override is documented in [`config-schema.md`](references/config-schema.md).
- Section properties sparse-by-design: untouched unless declared via `pageSetup`. Details: [`config-schema.md`](references/config-schema.md#page-setup).
- Paragraph indexing is 1-based, matching `#NNN` in skeleton. Layout-table paragraphs are indexed; data-table cell paragraphs aren't. Both are reachable by surgical edits — layout cells via global `#NNN` index, data cells via `cell` locator with optional `paragraph: K` / `to: M`. `RunLocator` accepts cell coords too for `set-run`. See [`references/edit.md` locator table](references/edit.md#locators-at).
- All `edits[]` locators resolve against the **pre-edits** document state — `#NNN` from `overview` is what locators reference, regardless of intervening ops. Resolved Element refs survive subsequent mutations.
- Paths resolve against CWD; use absolute paths if you may have changed directories.
- Restyle: run-level direct formatting uniform across all runs gets stripped; per-run differences preserved as intentional inline emphasis.
- Field codes (`STYLEREF` / `TOC` / `REF` / `DATE` / …) preserved as-is; not editable inside. TOC content is not regenerated — user must right-click → "Update Field" in Word.
- Edit blockers: the edit phase refuses paragraphs inside existing tracked changes / complex fields / SDT controls. Run `inspect_blockers` first. **Ordering rule that follows from it: content controls (SDT) are inserted LAST** — run every `edits[]` op first, then add SDTs (`extend` / raw fragments). An SDT placed before the edit phase silently blocks every later edit that targets paragraphs inside it.
- Style-name preflight: before any mutation, the engine catches `<w:name>` collisions (including en/zh-CN locale aliases) and aborts with a fix hint; `--dry-run` lists collisions as warnings instead.
- `settings.elements` is order-safe and duplicate-safe: each entry is merged into an existing element of the same name when one exists, otherwise inserted at the XSD-correct position from the authoritative 98-child order. `updateFields`, `evenAndOddHeaders`, etc. never appear twice. Repeatable children work too — declaring the same child name more than once (e.g. three `docVar`s inside `docVars`, or several `rsid`s) appends one node per declaration instead of collapsing them into a single node.
- `docProps/core.xml` + `app.xml` are created or refreshed automatically (`docProps.enabled`, default on): title / creator / revision, plus `Words` / `Characters` / `Paragraphs` counted the way Word counts Chinese (CJK per character, Latin per word). These parts are what `NUMPAGES` / `NUMWORDS` / `AUTHOR` fields read.
- Root namespace declarations are repaired on every part written — `xmlns:w14` / `w15` / `wp14` / `wps` / `wpg` / `wpc` / `v` / `o` / `mc` as needed, with rendering prefixes (`wps` / `wpg` / `wpc`) explicitly kept **out** of `mc:Ignorable`. Template files that declare only `xmlns:w` + `xmlns:r` are the normal case, not the exception.
- **Before delivering any file, run the PRO validator** — `python pro/scripts/validate.py out.docx` — not just the Node `validate.js`. The two check different layers (XSD + OPC graph vs. the Node schema pass) and the PRO run also reports paragraph / word counts and redlining completeness. Treat a non-zero exit as not-done. This is one half of the gate; the other half is `verify` — see [§ Delivery gate](#delivery-gate).
- **Multi-step chains go through `pipeline`, not by hand.** Running `apply` and `extend` as separately-remembered commands is how a stale intermediate ships: regenerate the config *source* but forget to re-emit the config JSON, and the writer happily re-writes the previous config while every downstream validator reports PASS. Declare the chain once; the runner aborts on the first failure and validates the deliverable.
- **Schema-valid is not delivery-ready.** A file can pass every validator and still be wrong — an unrequested footer title is well-formed XML. Assertions are the mechanical guard; reading the change manifest against the user's own words is the one no program can do for you.
- **`mergeRuns` renumbers runs.** It is the only pre-pass that mutates the source DOM ahead of design decisions, so a `set-run` locator written against the fragment-laden original shifts under it. Reach for it *instead of* hand-addressing `runIndex` values, not alongside them. It is idempotent (an already-merged document merges 0 runs) and `--dry-run` reports it with the rest of the change report.

Block-level config details: [`references/standardize.md`](references/standardize.md), [`references/edit.md`](references/edit.md), [`references/config-schema.md`](references/config-schema.md).

## Where things live

- **Field semantics** (any apply config field — values, constraints, defaults) → [`references/config-schema.md`](references/config-schema.md)
- **Locator + op contracts** (addressing, edit ops, block types) → [`references/edit.md`](references/edit.md)
- **Workflow** (audit → standardize → edit pipeline, design decisions) → respective workflow refs (`audit.md`, `standardize.md`, `edit.md`)
- **Domain knowledge** (Chinese counting patterns, caption pipeline, classifier reasons, table layouts) → concept refs (`numbering-formats.md`, `captions.md`, `tables.md`, `equations.md`, `cross-references.md`, `header-footer.md`, `chinese-font-sizes.md`)
- **PRO toolchain** (comments / accept-changes / run merging / render verification) → [`references/pro-toolchain.md`](references/pro-toolchain.md)
- **OOXML extension channel** (fields, notes, watermarks, shapes, charts, OLE, custom XML) → [`references/extend.md`](references/extend.md)
- **Element order** (the XSD-mandated child sequence of any element — authoritative, auto-generated) → [`references/element-order.md`](references/element-order.md)
- **Extension points** (root namespace declarations, `mc:Ignorable` rules, part registration, `docProps`) → [`references/ooxml-extension-points.md`](references/ooxml-extension-points.md)
- **What changed** (capability additions, fixed defects, deliberate boundaries — newest first) → [`CHANGELOG.md`](CHANGELOG.md)
