---
name: docx-master
description: "Standardize, restyle, edit, or audit Word (.docx / .dotx) documents by mutating OOXML directly — semantic styles + multi-level numbering, template import, content insertion (paragraphs / tables / images / equations), captions & cross-references, headers / footers, page setup, comments (批注), tracked changes, footnotes, TOC / watermarks / shapes / charts via `extend`, batch mail-merge, and a delivery gate (`pipeline` / `verify` / `doctor`). Use for any format / normalize / restyle / edit / audit / generate-a-Word-document task. Also: legacy `.doc` → `.docx` conversion and page rendering for visual checks. Do NOT use for PDFs, spreadsheets, or plain-text / Markdown sources (unless the task is specifically to *output* a docx)."
---

# docx-master

Mutates Word (.docx / .dotx) OOXML directly: produce well-formed documents from messy templates, install styles + numbering, convert hand-typed prefixes to auto-numbering, insert content, audit conformance. Output is a new docx; the original is never touched. Legacy binary `.doc` is not understood by the Node engine — convert first via `python pro/scripts/soffice.py --headless --convert-to docx file.doc` (see [pro-toolchain.md](references/pro-toolchain.md)).

**How this file works.** Contract + routing index; details live in `references/` and load **on demand** — when a rule names a reference, read it *before* designing the config that depends on it.

## How to think about formatting

User prompt overrides everything below.

Otherwise, treat a document as two kinds of paragraphs:

- **Content chrome** — paragraphs already carrying typed text (chapter markers, instruction notes, fixed labels). Three operations:
  - **Preserve typography** (font / size / spacing / indent — direct `pPr` / `rPr`); the engine's selective-strip keeps it unless the new style overrides the same attribute.
  - **Retag structure** (bind `pStyle`, strip typed structural prefixes via `pattern_rules`, install matching `numbering`). Chapter / heading / list / caption markers are structure, not typography — they belong to auto-numbering or SEQ fields. Existing typed prefixes are the **input** to retagging, not output to preserve.
  - **Don't replace, don't duplicate.** A chrome paragraph stays as itself — never the target of `replace`, never quoted into an inserted paragraph's `text`. Content insertion lands in the empty placeholder slot beside the chrome (`set-run` on inline blanks; `replace` against the empty slot for separate-paragraph forms).
- **Empty slots** — blank paragraphs pre-allocated for content insertion. Bind inserted content to a semantic style (`BodyText` / `ListNumber` / `Heading1..N`), installing the style in `styles[]` if no existing one fits. `Normal` is fallback, not a body style.

**For `styles[]`**: one style per semantic role. If the role has any paragraph already playing it in the source, the entry MUST use `fromParagraph: N` — do not redeclare top-level typography fields on represented roles (silent template override). Rationale: [standardize.md §1](references/standardize.md#1-style-roles--typography).

Tools surface visible facts; classification and judgment are yours.

Standardize-shape blocks install the canonical anchors every other op binds to (`edits[]` `styleId` + MDF, `pattern_rules` fingerprints, `InlineRef` / STYLEREF resolution). Typed structural signals (heading / list / caption prefixes, in-prose counters) never appear as literal text in the emit. User-supplied content arriving with typed prefixes is input shape, not a directive to preserve them: strip the prefix and emit each typed-prefix line as a **separate** paragraph bound to a heading / list / caption style, unless the user explicitly asked for literal typed form.

## Target state

Structure comes from **styles + numbering + sections + SEQ fields**, not typed text mimicking structure:

- Every paragraph carries a semantic styleId (Heading1..N / BodyText / ListNumber / Caption / …); direct format only as one-off exceptions.
- Numbering: one unified multi-level scheme for headings; separate single-level schemes per list-bound style. Markers come from the scheme, never typed in `text`.
- Captions (figures / tables / equations) use top-level `captions` config + `CaptionBlock` / `EquationBlock` — SEQ + STYLEREF fields, not `pattern_rules`. **Placement: figure / equation captions go below; table captions go above.** See [captions.md](references/captions.md).
- Cross-references (any cite to an auto-numbered target — figures / tables / sections / equations / reference entries) use `InlineRef` in `edits[]`, never typed counters — literal counters silently desync when a target moves, REF fields stay correct. Page-number cites out of scope ([cross-references.md](references/cross-references.md)).
- Heading levels nest without skipping.
- Match content shape to slot shape — prose → body, list → `ListNumber` / `ListBullet`, figure / table / equation → their `*Block` types with `captionId` ([edit.md](references/edit.md)).
- Locale: CJK body / list items get a 2-char first-line indent (`firstLineIndent: "2char"`); no typed spaces at CJK ↔ Latin / digit boundaries. Font sizes: [chinese-font-sizes.md](references/chinese-font-sizes.md).
- Every write reports advisory Pangu findings (CJK ↔ Latin / digit spacing) in `lint.pangu` for `edits[]` and `headerFooter` text — **never blocks the write**. Deliberate spacing → `lint.panguSpacing: false`.

## Commands

- **`apply`** — the unified writer: install styles + numbering + theme + template, restyle by pattern / fingerprint, insert content via edits. One config, one call. → [standardize.md](references/standardize.md) + [edit.md](references/edit.md)
- `audit` — read-only conformance check: no CLI, uses the inspect tools to produce a violation report. → [audit.md](references/audit.md)
- **`extend`** — post-processing for what the engine deliberately does not write: TOC / table-of-figures / index fields, footnotes & endnotes, watermarks, page borders, shapes / text boxes / canvas / WordArt / diagrams, charts, OLE, glossary, cover pages, custom XML. → [extend.md](references/extend.md)
- **`pipeline`** — one-command rebuild of a multi-step chain (config source → `apply` → `extend` → validate → verify); aborts on first failure, auto-validates the deliverable, `--plan` / `--emit-script` when child processes are forbidden. → [delivery-gate.md](references/delivery-gate.md)
- **`verify`** — the delivery gate: declarative assertions (exit 1 on failure) + change manifest vs. a baseline. → [delivery-gate.md](references/delivery-gate.md)
- `doctor` — environment capability probe; run first on an unfamiliar machine. → [delivery-gate.md](references/delivery-gate.md)

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

Need fields / notes / watermarks / drawings / charts? Run `extend` on `apply`'s output — the two are complementary, not alternatives. It re-declares root namespaces per part, keeps rendering namespaces (`wps` / `wpg` / `wpc`) out of `mc:Ignorable`, and handles ordering + part registration ([element-order.md](references/element-order.md), [ooxml-extension-points.md](references/ooxml-extension-points.md)).

Sparse by design — only declared blocks apply; untouched styles / numbering / paragraphs / theme stay as they are. **Declare only what's wrong, missing, or what the user explicitly asks to change.**

**Creating from scratch.** Omit `source` to scaffold from the bundled blank template (one empty Normal paragraph, A4 portrait). `styles[]` must use Mode B (direct fields) — there is nothing to extract from. Paragraph 1 is a `replace` / `insert-after` target (`insert-before` leaves a stray leading empty paragraph; `delete` would empty the body). Incompatible with `template`.

**Documents of unknown provenance.** Word-authored files (and anything round-tripped through another tool) fragment a paragraph's text across many runs, so a phrase visible on the page is often not one contiguous string in the XML and text-matched anchors / `set-run` silently miss it. Fix before designing anything: `node scripts/merge_runs.js <file>` as a standalone pre-pass, or `mergeRuns: true` (CLI `--merge-runs`) so `apply` merges before it parses and every index you write refers to the merged layout. Rendering is unchanged either way; runs inside different `<w:ins>` / `<w:del>` wrappers are never joined.

## Workflow

1. **Survey.** `overview` first. Note from its output: **structure** (heading levels, numbering schemes with consumption counts — 0 = orphan, fingerprints, skeleton); **chrome formatting** (the direct `pPr` / `rPr` content paragraphs carry — this IS the document's typographic convention, so don't redeclare those attributes on re-tagging styles); **typed structural prefixes** (heading shapes `第N章` / `一、` / `1.1` feed `pattern_rules`; caption shapes and in-prose counters feed caption / cross-ref migration via `edits[]`); **form-fill paragraphs** (`label + gap`, or `label + ____` with a `<w:u/>` run); **source content** for fill tasks (roles that will land, heading depth, list / figure / table / equation presence).
2. **Design ONE config.** Plan styles first from the content's structural outline; route in one config. Config blocks (`styles` / `numbering` / `pattern_rules` / `bulk_rules` / `assignments` / `exclude` / `edits`) — see [standardize.md §1–§3](references/standardize.md) and [edit.md](references/edit.md). Reactive additions accrete debt.
3. **Dry-run.** `apply --dry-run` first; read the change report's signals (Style Resolution / Sample Affected / Implicit-keep failure / unmatched stripPrefix) against intent — see [standardize.md §5](references/standardize.md).
4. **Apply.** Output is a fresh docx; the original is never modified.

**Keep reconnaissance cheap.** These tools emit hundreds of lines — pass truncation flags (`overview --paras=A..B|none`, `--limit N`, `--problems-only`, `--json`, `--quiet`) so only the needed slice enters context; fetch a fact once, reuse it. Full list: [tools.md](references/tools.md) § 输出裁剪.

## Asking the user

Default: don't ask. Ask only for — typed sentinel mid-prose vs. heading · bold paragraph as sub-heading vs. emphasis · missing source content for template slots · unsupported structures (manual TOC body, page-number cites) · font / spacing scope ambiguity (theme layer vs `Normal` cascade vs per-role `styles[]` — [standardize.md "Font scope layers"](references/standardize.md#font-scope-layers)).

## Out of scope

**Blocked — surface to the user:** layout-**table** restructuring, TOC body content, page-number cites. Paragraphs *inside* layout-table cells ARE indexed and fully editable; only the table holding them is off-limits.

**Supported, despite appearances:** footnote / endnote *bodies* via `extend` (`{"op": "footnote"}` / `"endnote"}` creates part + rel + id; only *indexing* is absent — notes live in `word/footnotes.xml` / `word/endnotes.xml`); comments via PRO `comment.py`; page setup + headers / footers incl. per-section overrides ([extend.md](references/extend.md), [config-schema.md §page setup](references/config-schema.md#page-setup), [header-footer.md](references/header-footer.md)).

**Not attempted (binary part / external app):** `.dotx` · real SmartArt (`dgm`) · ActiveX · InkML · macros · envelope / label layouts · PDF/XPS · translation / dictation → [out-of-scope.md](references/out-of-scope.md).

## Tool index

All tools run as `node <script> <args>` → stdout. Write / delivery commands are in § Commands above; **完整调用、参数与输出裁剪 → [tools.md](references/tools.md)**。

- `overview` — always first: structure, page setup, theme, style defs, numbering usage, fingerprints, skeleton
- `inspect_range` · `inspect_runs` · `inspect_neighbors` — paragraph range + computed styles · per-run rPr (also form-fill segments) · surrounding paragraphs (caption / first-after-heading)
- `inspect_style` · `inspect_style_def` — role a fingerprint plays · existing `styleId` definitions + `basedOn` chain
- `inspect_section` · `inspect_table` — page-setup differences between sections · top-level tables with cell spans
- `inspect_blockers` — paragraphs the edit phase refuses: tracked changes, complex fields, SDT
- `inspect_bookmarks` — bookmarks + references naming one + dangling ones (non-zero exit on problems)
- `inspect_caption` · `migrate_captions` — existing SEQ captions · detector for hand-numbered caption paragraphs
- `find_paragraphs` · `find_text` — paragraph regex · character-level locator (paragraph / run / offset)
- `merge_runs` — coalesce fragmented runs; before a surgical edit on unknown-provenance files

## PRO toolchain (optional extension)

The Node engine is the default and zero-dependency. The vendored Python toolchain under `pro/` covers only what Node cannot — **comments**, **accepting tracked changes**, **legacy `.doc` conversion**, **render-to-image verification**, **XSD / redlining check**, **element-order lookup**, docx-js footguns.

`doctor.js` reports which optional tools are missing — install them before using the render-verify, `.doc` conversion, or docx-js paths. Routing table, invocation conventions (run from `pro/scripts/`), boundary conflicts, fallback: [pro-toolchain.md](references/pro-toolchain.md).

## Delivery gate

Schema validation answers *is this file well-formed?* — never *is this what the user asked for?* A footer carrying an unrequested title validates perfectly. Both get answered before handover:

1. Probe — `node scripts/doctor.js`: always exits 0 (a report, not a gate)
2. Dry-run each config edit — `node scripts/apply.js --dry-run <cfg>`
3. Run the whole chain — `node scripts/pipeline.js build.json`: aborts on first failure, auto-appends validation
4. Assert the artifact — `node scripts/verify.js verify.json`: exit 1 on any failed assertion
5. Read the change manifest — `node scripts/verify.js <out.docx> --diff <src.docx>`: every line must be defensible against the user's words

**Schema-valid is not delivery-ready.** Write assertions *before* the config — the requirements stated mechanically, the only guard against the defect class no schema check can see. 全量断言类型、`--emit-script` 降级路径、变更清单读法、人工自检 → **[delivery-gate.md](references/delivery-gate.md)**。

## Hard rules

The rules **you** must actively follow. What the engine handles on your behalf — namespace
repair, `settings` idempotence, `docProps` refresh, style-name preflight, restyle semantics —
is a separate list: [engine-behavior.md](references/engine-behavior.md). Read it when designing
a config that touches those points.

- Original file is never modified; every applying CLI writes a fresh copy and validates before keeping it. Validation is baseline-diff — only errors *introduced* by the run are fatal ([standardize.md §Validation behavior](references/standardize.md#validation-behavior)).
- Paragraph indexing is 1-based, matching `#NNN` in the skeleton. Layout-table paragraphs are indexed; data-table cells aren't — reach layout cells via global `#NNN`, data cells via a `cell` locator with optional `paragraph: K` / `to: M`. `RunLocator` takes cell coords too for `set-run` ([edit.md locator table](references/edit.md#locators-at)).
- All `edits[]` locators resolve against the **pre-edits** document state — the `#NNN` from `overview` is what locators reference, regardless of intervening ops. Resolved Element refs survive subsequent mutations.
- Paths resolve against CWD; use absolute paths if you may have changed directories.
- Field codes (`STYLEREF` / `TOC` / `REF` / `DATE` / …) are preserved as-is and are not editable inside. TOC content is not regenerated — the user must right-click → "Update Field" in Word.
- Edit blockers: the edit phase refuses paragraphs inside existing tracked changes / complex fields / SDT controls — run `inspect_blockers` first. **Ordering rule that follows: SDTs are inserted LAST** — every `edits[]` op first, then `extend` / raw fragments. An SDT placed before the edit phase silently blocks every later edit targeting paragraphs inside it.
- **A bookmark has to exist before anything names it.** Three doors create one — a Block's `anchor` (content this run emits; pre-scanned, so forward refs work), `format`'s `anchor` (a paragraph already on disk), `extend`'s `bookmark` op (any paragraph, or a `from`/`to` range). Wire it via `refTo` → REF, `{ "link": "#name" }` → in-page jump, `extend field` → `PAGEREF` / `TOC \b name`. `display: "number"` needs an auto-numbered target, else `"full"`. Table: [cross-references.md](references/cross-references.md#creating-an-anchor--the-three-doors).
- `_Toc*` / `_Ref*` / `_GoBack` bookmarks are **Word's, not yours** — written when Word updates a field. The TOC field ships *empty* by design ([toc_field](references/extend.md)); the `_Toc*` + PAGEREF + HYPERLINK triplets behind clickable TOC entries appear only after Word recomputes. Never cite or hand-edit them. A `_Toc*` reference whose bookmark is gone IS a defect — `inspect_bookmarks` and `refsResolve` catch exactly that.
- **Before delivering, run the PRO validator too** — `python pro/scripts/validate.py out.docx`, not just `validate.js`; the two check different layers (XSD + OPC graph vs. Node schema pass). Non-zero exit = not done. One half of the gate; `verify` is the other.
- **Multi-step chains go through `pipeline`, not by hand.** Separately-remembered `apply` / `extend` commands are how a stale intermediate ships: regenerate the config *source* but forget to re-emit the JSON, and the writer re-writes the previous config while every validator reports PASS.
- **`mergeRuns` renumbers runs** — the only pre-pass that mutates the source DOM ahead of design decisions, so a `set-run` locator written against the fragment-laden original shifts under it. Reach for it *instead of* hand-addressing `runIndex`, not alongside. Idempotent.

Block-level config details: [standardize.md](references/standardize.md), [edit.md](references/edit.md), [config-schema.md](references/config-schema.md).

## Where things live

- **Field semantics** → [config-schema.md](references/config-schema.md) · **Tool invocations / flags / truncation** → [tools.md](references/tools.md) · **Locator + op contracts** → [edit.md](references/edit.md)
- **Workflows** (audit / standardize / edit) → [audit.md](references/audit.md), [standardize.md](references/standardize.md), [edit.md](references/edit.md)
- **Delivery gate** (assertions / manifest / handover checklist) → [delivery-gate.md](references/delivery-gate.md) · **Engine auto-behavior** → [engine-behavior.md](references/engine-behavior.md)
- **Domain knowledge** → [numbering-formats.md](references/numbering-formats.md), [captions.md](references/captions.md), [tables.md](references/tables.md), [equations.md](references/equations.md), [cross-references.md](references/cross-references.md), [header-footer.md](references/header-footer.md), [chinese-font-sizes.md](references/chinese-font-sizes.md)
- **PRO toolchain** → [pro-toolchain.md](references/pro-toolchain.md) · **OOXML extension channel** → [extend.md](references/extend.md)
- **Element order** → [element-order.md](references/element-order.md) · **Extension points** → [ooxml-extension-points.md](references/ooxml-extension-points.md) · **What changed** → [CHANGELOG.md](CHANGELOG.md)
