# docx-master

<div align="center">

[简体中文](README.md) | **English**

</div>

<div align="center">

**A Word skill that mutates OOXML directly**

Unzip the `.docx`, edit the XML inside, zip it back — **generate · read · edit · validate**, zero-dependency main chain

![License](https://img.shields.io/badge/License-MIT-blue)
![release](https://img.shields.io/badge/release-v0.16.0-green)
![stars](https://img.shields.io/badge/stars-%F0%9F%8C%9F%20%E2%9C%A8-yellow)
![standard](https://img.shields.io/badge/standard-everyday%20·%20thesis%20·%20reports-3182ce)
![pro](https://img.shields.io/badge/pro-watermarks%20·%20shapes%20·%20charts%20·%20OLE-805ad5)
![deps](https://img.shields.io/badge/main%20chain-zero%20deps%20Node%20%2B%20lxml-orange)

[FAQ](README_EN.md#frequently-asked-questions) · [Changelog](CHANGELOG.md) · [Dev conventions](README_EN.md#development-conventions) · [Security](SECURITY.md) · [Contributing](CONTRIBUTING.md)

</div>

> 🤖 **AI / developers: read [SKILL.md](SKILL.md) first** (main flow, hard rules, routing index) — don't scan the whole repo first.

**Author**: Zoros ([Zok66](https://github.com/Zok66)) · MIT-licensed, free to use / modify / sell commercially (just keep the copyright notice)

## Contents

- ① [Highlights](#highlights) — why this exists, in one pass
- ② [Privacy & Security](#privacy--security) · [Two editions](#two-editions) — local-only processing; standard vs pro boundary
- ③ [Version & Changelog](#version--changelog) · [Documentation index](#documentation-index) — current release, recent changes, every doc
- ④ [Capabilities](#capabilities) · [Install](#install) · [Quick start](#quick-start) — what it does, how to install, four commands
- ⑤ [Design principles](#design-principles) · [Development conventions](#development-conventions) · [FAQ](#frequently-asked-questions) — why, how to commit, common pitfalls
- ⑥ [Repository layout](#repository-layout) · [Regression tests](#regression-tests) · [Dependencies](#dependencies) · [Scope](#scope) · [License](#license)

## Highlights

- **Original file is never modified** — every write command produces a new copy, auto-validated twice before landing (Node schema + XSD/OPC)
- **Full generate · read · edit · validate chain, zero npm dependencies** — main chain uses only the Node standard library + Python lxml
- **Normalize**: turns hand-typed `一、`/`1.1` prefixes into real multi-level auto-numbering and reshapes typography by role
- **Ready for theses / lab reports**: TOC fields, page numbering, headers & footers, footnotes, captions, cross-references, citations & bibliography
- **Delivery gate**: declarative assertions (failure exits 1, **blocking delivery**) + change manifest — catches changes that are schema-perfect but never requested
- **Invisible objects get a dedicated guard**: dangling bookmarks / broken references print "Error! Bookmark not defined." in Word while XSD stays green — `inspect_bookmarks` + `refsResolve` see them
- **Batch mail-merge**: one template × one CSV → N documents (a replacement for Word mail-merge)
- **One command for the whole chain**: `pipeline` aborts on first failure — no half-done intermediates
- **AI-friendly**: `SKILL.md` master entry + single-source-of-truth rule contracts, guarded by regression tests

Zero-dependency is not a slogan — it is four concrete trade-offs ("drop / replace / gain"):

| ① Not depending on | ② Replaced by | ③ So you can |
|---|---|---|
| Microsoft Word | Reading/writing OOXML directly (ECMA-376 / ISO 29500) | Restyle paragraphs, multi-level auto-numbering, TOC fields, page borders, content controls, mail-merge |
| LibreOffice | Node stdlib + Python `lxml` for the main chain | Run headless in CI, sandboxes, and servers — no GUI, no COM |
| npm `docx` package | Hand-written `w:p` / `w:tbl` / `w:sdt` | Touch corners no library wraps; never blocked by API expressiveness |
| Cloud conversion services | Everything local | Documents never leave the machine — safe for internal material |

## Privacy & Security

- **Documents are processed entirely locally** — nothing is uploaded to any cloud conversion service
- **Macros / OLE payloads inside documents are never executed**; no network access (`doctor` only probes locally installed tools)
- When processing **untrusted** docx files, run in an isolated environment (same common sense as opening unknown Word documents)
- Report vulnerabilities privately — see [SECURITY.md](SECURITY.md)

No Microsoft Word, no LibreOffice, no npm `docx` package. Unzip the `.docx`, edit the XML inside,
zip it back — which is why it can do far more than "what a library API can wrap". Three columns say it plainly:

## Two editions

One repo, two branches. **Standard (for everyday office work, theses, lab reports)** — enough and light;
**pro is full-featured**, covering all the rarely-used extension objects.

| Capability | Standard (`main`, default) | pro (`pro`) |
|---|:---:|:---:|
| Generate / edit / normalize / template import | ✅ | ✅ |
| Read & recon (12 `inspect_*` tools + `find_text`) | ✅ | ✅ |
| Page count / breaks / numbering / TOC / index | ✅ | ✅ |
| Headers & footers (per-section overrides) / page setup / page borders | ✅ | ✅ |
| Footnotes & endnotes / captions / cross-references / bookmarks | ✅ | ✅ |
| Images (inline pictures, crop, effects, floating wrap) | ✅ | ✅ |
| Comments / tracked changes / accept revisions / batch merge / delivery gate | ✅ | ✅ |
| Cover page / drop cap / content controls / citations & bibliography | ✅ | ✅ |
| Watermarks | — | ✅ |
| Shapes / text boxes / WordArt / canvas / SmartArt | — | ✅ |
| Charts (embedded workbook) / OLE embedding | — | ✅ |
| Building blocks (AutoText) / custom XML | — | ✅ |
| Full XSD schema set (complete schema validation) | trimmed | full |

Install and edition switching: see [Install](#install). The 8 stripped `extend` ops **still exist in the code** (just not registered in `OPERATIONS`);
calling them fails loudly with the 16 available ops listed — never a silent downgrade.
See [`references/pro-only-ops.md`](references/pro-only-ops.md) for the list and history.

## Version & Changelog

- **Current: [v0.16.0](https://github.com/Zok66/docx-master/releases/tag/v0.16.0)** (2026-10-20, first standard-edition release).
- **License**: [MIT](LICENSE) — free to use, modify, distribute, and use commercially; keep the copyright and permission notice.
- **Upgrade**: `git pull` inside the skill directory, then self-check as described under [Regression tests](#regression-tests).
- **Download / history**: see [Releases](https://github.com/Zok66/docx-master/releases) — Latest is the recommended one; each release ships a capability list and acceptance record.

### At a glance

| Release | Positioning | Branch | Key changes |
|---|---|---|---|
| **v0.16.0** ⬅ current | Standard · everyday office / theses / lab reports | `main` | Edition split, community files, star history, bilingual README |
| v0.16.0-pro | Full-featured · watermarks / shapes / charts / OLE | `pro` | Same code + all 24 `extend` ops |
| v0.15.0 | Last pre-split release (full-featured) | — | 31 format fields, pinyin annotation, WPS "Insert" alignment |
| Earlier | — | — | See rounds 5–15 in [CHANGELOG.md](CHANGELOG.md) |

### Changelog

Full record lives in **[`CHANGELOG.md`](CHANGELOG.md)** — capability and behaviour changes only
(new fields / fixed defects / clarified boundaries), never documentation polish. Every entry maps to
code in `apply.js` or `scripts/extend/*.py`.

Last three entries:

| Date | Release | Summary |
|---|---|---|
| 2026-10-20 | — | License reverted to MIT (a brief PolyForm NC change was rolled back; `LICENSE` matches v0.16.0) |
| 2026-10-20 | **v0.16.0** | ① standard / pro split (`OPERATIONS` 16 vs 24) ② 8 ops stripped with loud errors ③ bilingual README, community files, star history |
| 2026-10-20 | v0.15.0 | ① 31 run/paragraph fields + pinyin (ruby) ② WPS "Insert" alignment ③ three comment-anchor fixes |

Release procedure (tag → dual-platform release → verification) is documented in [`Push-Sync.markdown`](Push-Sync.markdown).

## Documentation index

| Doc | Contents |
|---|---|
| [`SKILL.md`](SKILL.md) | Skill entry — main flow, hard rules, routing index (AI reads this on trigger) |
| [`references/pro-only-ops.md`](references/pro-only-ops.md) | The 8 ops not in standard + how to switch to pro |
| [`CHANGELOG.md`](CHANGELOG.md) | Full change record (capability & behavior changes, searchable per round) |
| [`CONTRIBUTING.md`](CONTRIBUTING.md) | How to contribute: workflow, hard constraints, bug report format |
| [`CODE_OF_CONDUCT.md`](CODE_OF_CONDUCT.md) | Code of conduct |
| [`SECURITY.md`](SECURITY.md) | Security policy: supported versions, private vulnerability reporting, threat model |
| [`tests/README.md`](tests/README.md) | Running regression tests, adding rules / cases |
| [`LICENSE`](LICENSE) | MIT license |
| [`THIRD-PARTY-NOTICES.md`](THIRD-PARTY-NOTICES.md) | Third-party notices (Temml / ECMA-376 XSD provenance & licensing) |
| [`Push-Sync.markdown`](Push-Sync.markdown) | GitHub ↔ Gitee full-sync guide (branches / tags / releases / license) |
| [`references/`](references/) | Domain references: config schema, locator contracts, tool tables, numbering formats, captions, formulas … |
| [`pro/`](pro/) | Optional Python toolchain: comments, accept revisions, rendering, XSD validation |

## Capabilities

| Action | What it does |
|---|---|
| **Generate** | Omit `source` to scaffold a new document from the bundled blank template; or import styles & numbering from an existing template |
| **Normalize** | Reshape the whole document by role: classify paragraphs → bind named styles → install multi-level auto-numbering, turning hand-typed `一、`/`1.1` into real numbering |
| **Edit** | Locator-based surgery: replace / insert / delete paragraphs, table cells, embedded images — optionally leaving tracked changes |
| **Extend** | `extend` fills what the engine deliberately doesn't write: TOC / table-of-figures / index fields, page & section breaks, page numbering, named bookmarks, footnotes & endnotes, page borders & background, images, content controls, cover pages, citations. Watermarks / shapes / charts / OLE / SmartArt etc. are pro-only (see [Two editions](#two-editions)) |
| **Bookmarks / cross-references** | Named bookmarks (three entry points), `refTo` → REF cross-references, `{ "link": "#name" }` in-page jumps; `inspect_bookmarks` + `refsResolve` assertions catch the **invisible defect class** "reference points to a missing bookmark" |
| **Batch** | One template + one CSV / JSON data table → one document per row (a replacement for Word mail-merge) |
| **Orchestrate** | `pipeline`: declare 「config → apply → extend → validate → assert」 as one chain, **run in a single command**, abort on first failure — no half-done intermediates. In no-subprocess environments (sandboxes) use `--emit-script` to generate a `set -e` script with identical abort semantics |
| **Read** | Structure recon: style definitions & inheritance chains, numbering schemes & consumption, sections & page setup, theme, captions, tables, document skeleton, paragraph ranges with computed styles, per-run formatting, bookmarks & cross-references |
| **Validate** | Three layers: Node schema check + XSD / OPC package-graph check (run automatically before writing), plus `doctor` environment probing |
| **Delivery gate** | `verify`: ① declarative assertions (e.g. "footer must not contain the book title") exit non-zero on failure — **blocking delivery**; ② change manifest (output vs baseline) lists every change introduced, to check against the user's words |
| **Comments / revisions** | Add comments (with replies), accept all tracked changes into a clean copy |
| **Convert** | Legacy binary `.doc` → `.docx` (needs LibreOffice), render to images for visual review |

## Install

Drop it into your skills directory (any environment that supports `SKILL.md` — WorkBuddy / Claude Code etc.):

```bash
# Standard edition (default — what you get on clone)
git clone https://github.com/Zok66/docx-master.git ~/.workbuddy/skills/docx-master

# Full-featured: switch branch in the same repo, no reinstall
cd ~/.workbuddy/skills/docx-master && git checkout pro

# back to standard
git checkout main
```

## Quick start

```bash
SK=~/.workbuddy/skills/docx-master

# ① Before anything: probe environment capabilities (what's really available, what degrades)
node $SK/scripts/doctor.js

# ② Recon first: structure, styles, numbering, skeleton
node $SK/scripts/overview.js my-doc.docx
node $SK/scripts/inspect_bookmarks.js my-doc.docx   # bookmarks + references + dangling ones (exit 1 on problems)

# ③ Single-step design: one config — dry-run first, then write (new file; original untouched)
node $SK/scripts/apply.js --dry-run my.config.json
node $SK/scripts/apply.js my.config.json
python $SK/scripts/extend/extend.py extend.config.json   # TOC / numbering / notes / page borders … (16 ops in standard; watermarks / charts are pro-only)

# ④ Or: declare the whole chain and run it in one command (recommended)
node $SK/scripts/pipeline.js build.json          # aborts on first failure, auto-validates at the end
node $SK/scripts/pipeline.js build.json --plan   # just show what would run, write nothing

# ⑤ Delivery gate: assertions + change manifest
node $SK/scripts/verify.js verify.json                     # failed assertion → exit 1, delivery blocked
node $SK/scripts/verify.js out.docx --diff source.docx     # list what changed, line by line
```

An assertion file (`verify.json`) — the user's requirements expressed mechanically:

```json
{
  "file": "out.docx",
  "baseline": "source.docx",
  "assert": [
    { "type": "sectionCount", "eq": 3 },
    { "type": "hasField", "field": "TOC", "scope": "body" },
    { "type": "headerText", "regex": "author name" },
    { "type": "notContains", "text": "book title", "scope": "footer" },
    { "type": "footerText", "regex": "^\\s*Page\\s+<PAGE>\\s+of\\s+<NUMPAGES>\\s*$" },
    { "type": "refsResolve" },
    { "type": "bookmarkProblems" }
  ]
}
```

The middle entries exist to catch changes that are **schema-perfect but never requested** — like a
book title sneaking into the footer. No schema validator can see that. The last two are the same
idea: a REF or hyperlink pointing at a missing bookmark renders in Word as "Error! Bookmark not
defined." while staying fully schema-valid.

Batch mail-merge:

```bash
node $SK/scripts/batch_fill.js contract-template.docx clients.csv \
  --out-dir out/ --name "{{id}}-{{client}}.docx" --validate
```

## Design principles

1. **The original file is never modified.** Every write command produces a new copy; the source stays untouched.
2. **Sparse by design.** Declare only what's broken, missing, or explicitly requested; untouched styles / numbering / paragraphs stay as they are.
3. **Validate before delivering.** Two validators check different layers (Node schema vs XSD + OPC package graph); both must pass.
4. **Multi-step chains must be machine-orchestrated.** Relying on a human to remember "re-run the config source after editing" is how stale intermediates ship: the writer uses the **previous config** while every later validator reports PASS. Declare the chain once, run it once.
5. **"Schema-valid" ≠ "deliverable".** Unrequested changes (an extra title line in the footer, an extra signature) are valid XML — and defects. Assertions handle the mechanical half; reading the change manifest handles the rest.
6. **Invisible objects need dedicated checks.** Bookmarks render nothing, carry no style, add no word count: lose one and the file still opens, still passes XSD — until Word prints "Error! Bookmark not defined." mid-sentence. `inspect_bookmarks` (human) and `refsResolve` (machine) are the only two places that look.
7. **Content and structure are separate.** Structural signals — headings, lists, captions — are rendered by styles, numbering and SEQ fields, never typed as literal text.
8. **The master file holds rules only; details load on demand.** `SKILL.md` enters context in full on every trigger, so it carries only what must be known immediately: main flow, hard rules, routing index. Long tables and semantic detail live in `references/`, read when needed. Every rule is defined in **exactly one** place; everything else points at it — `tests/run.mjs` guards this with a contract list: rules may move files, but must never vanish.

## Development conventions

### Commit messages

No forced format — just make **what changed and why** clear. Two conventions:

① **One commit, one thing.** Code + its tests + CHANGELOG land together; pure doc polish goes separately.

② **Keep branches in sync**: `main` and `pro` share code. After landing on `main`,
run `git checkout pro && git checkout main -- <changed files>` and commit again.

### Hard constraints

| Constraint | How it is enforced |
|---|---|
| `SKILL.md` ≤ 20000 B | `node tests/run.mjs` fails the build; overflow moves to `references/` with a pointer left behind |
| **One rule, one definition** | Rules are listed in `tests/contract.json`; delete one and tests go red |
| All internal links (anchors included) must resolve | The "links" contract class in `node tests/run.mjs` |
| New capability → contract entry + CHANGELOG | See [CONTRIBUTING.md](CONTRIBUTING.md) |

### Before you commit

Run the `--smoke` tier described under [Regression tests](#regression-tests).
Contribution flow and bug-report format: [CONTRIBUTING.md](CONTRIBUTING.md).

## Frequently asked questions

<details>
<summary><b>① Do I need Microsoft Word or LibreOffice?</b></summary>

**Neither** for the main chain. What depends on what: see the trade-off table at the end of
[Highlights](#highlights) and the [Dependencies](#dependencies) section. The two optional capabilities
(`.doc` conversion, render-to-image) need LibreOffice. `node scripts/doctor.js` probes what is actually available.

</details>

<details>
<summary><b>② Will it modify my source file?</b></summary>

**Never.** The source stays byte-for-byte untouched, and `verify.js --diff` is read-only.

</details>

<details>
<summary><b>③ The TOC came out empty / page numbers don't show?</b></summary>

That is inherent OOXML behaviour, **not a defect**. TOC, page numbers and SEQ are **fields**: Word
renders the cached value on open. The field code is written; the rendered result needs Word's calculation.

Fix: open the file, `Ctrl+A` then `F9` (or right-click the TOC → Update Field). Same in WPS.

</details>

<details>
<summary><b>④ What is "Error! Bookmark not defined."?</b></summary>

A REF field or internal link points at a **bookmark that does not exist**. XSD cannot see this and Word
only prints a line of text in the middle of a sentence.

```bash
node scripts/inspect_bookmarks.js my-doc.docx   # bookmarks + all references + dangling ones, exit 1 on problems
```

Then gate it with the `refsResolve` / `bookmarkProblems` assertions in `verify`.

</details>

<details>
<summary><b>⑤ Standard or pro — do I reinstall to switch?</b></summary>

No reinstall — one repo, two branches: `git checkout pro` / `git checkout main`, commands under [Install](#install).

Everyday office work, theses and lab reports are covered by the standard edition; switch to `pro` for
watermarks, shapes and charts. Full comparison in [Two editions](#two-editions), the 8-op list in
[`references/pro-only-ops.md`](references/pro-only-ops.md).

</details>

<details>
<summary><b>⑥ What if I call a pro-only op on the standard edition?</b></summary>

**A loud error, never a silent downgrade.** The 8 stripped ops still exist in code, just not registered in
`OPERATIONS`. Calling one exits with the 16 available ops listed plus how to switch to `pro`.

</details>

<details>
<summary><b>⑦ What are the Chinese typography defaults?</b></summary>

Body text defaults to a two-character first-line indent (`firstLineIndent: "2char"`); CJK fonts and sizes
follow Chinese document / thesis conventions. Automatic spacing between CJK and Latin text is controlled
by `lint.panguSpacing` (silent by default). Fields are documented in
[`references/config-schema.md`](references/config-schema.md) and [`references/chinese-font-sizes.md`](references/chinese-font-sizes.md).

</details>

<details>
<summary><b>⑧ Can it batch-generate documents?</b></summary>

Yes — it replaces Word mail merge in one command; see the `batch_fill.js` example at the end of
[Quick start](#quick-start). Placeholders live in the template; each CSV / JSON row produces one output,
and `--validate` checks every file.

</details>

<details>
<summary><b>⑨ Is it safe for confidential or internal documents?</b></summary>

Everything is local: no network, no cloud conversion, no execution of macros or OLE payloads.
See [Privacy & Security](#privacy--security) and [SECURITY.md](SECURITY.md). For files of unknown
origin, still prefer an isolated environment — the same common sense as opening an unknown Word document.

</details>

<details>
<summary><b>⑩ Does it handle PDF / xlsx / Markdown?</b></summary>

**No.** This is a docx-only skill: PDF, spreadsheets (xlsx) and plain-text / Markdown sources are out of
scope — unless the task is to **produce** a docx. See [Scope](#scope).

</details>

<details>
<summary><b>⑪ Can layout tables be restructured?</b></summary>

Table structure cannot be reshuffled wholesale (row insertion/deletion, cell merging as a block), but
**paragraphs inside cells are fully editable**. That boundary is deliberate: structural rewrites have too
many failure modes and are safer done by a human in Word.

</details>

<details>
<summary><b>⑫ Can I use it commercially?</b></summary>

Yes. [MIT](LICENSE) — free to use, modify, distribute and sell commercially, keeping the copyright and
permission notice.

</details>

## Repository layout

```
docx-master/
├── SKILL.md              # skill entry — main flow / hard rules / routing index (details on demand in references/)
├── CHANGELOG.md          # capability & behavior changes (added / fixed / boundaries)
├── README.md             # this file
├── scripts/              # main chain: zero-dependency Node + Python lxml
│   ├── apply.js          # unified writer (styles / numbering / theme / template / edits)
│   ├── batch_fill.js     # template × data table → N documents
│   ├── pipeline.js       # single-entry orchestration: one command for the whole chain, abort on failure
│   ├── verify.js         # delivery gate: declarative assertions + change manifest (reads OPC directly, zero deps)
│   ├── inspect_bookmarks.js  # bookmark / reference / dangling-ref recon (exit 1 on problems)
│   ├── doctor.js         # environment capability probe (subprocess / Python / render toolchain)
│   ├── merge_runs.js     # coalesce fragmented runs (pre-pass before editing foreign documents)
│   ├── validate.js       # schema validation
│   ├── overview.js       # plus 12 inspect_* recon tools
│   ├── _shared/          # shared modules: OPC reading, bookmark analysis, locators, validators …
│   └── extend/           # extension writer (TOC / bookmarks / notes / graphics / charts / OLE …)
├── references/           # domain references: config schema, locator contracts, tool tables, numbering formats, captions, formulas …
├── tests/                # regression: structure / rules / resources / links contracts + smoke cases
└── pro/                  # optional Python toolchain (comments, accept revisions, rendering, XSD validation)
```

This section describes the **standard edition** (`main`); the `pro` branch has the same layout.

## Regression tests

Zero dependencies, one command:

```bash
node tests/run.mjs            # structure + rules + resources + links contracts
node tests/run.mjs --quiet    # print failures only
node tests/run.mjs --smoke    # additionally run tests/cases/*.json
```

It answers four questions: is `SKILL.md`'s frontmatter valid and within size budget, **is every
"must-not-delete" hard rule from the contract list still present**, do all declared references /
scripts exist, and do all documentation links (including anchors) resolve. Run it after any skill change.

How to add rules and cases: [`tests/README.md`](tests/README.md).

## Dependencies

| Part | Dependencies |
|---|---|
| Main chain (`scripts/`) | Node (no third-party packages) + Python + `lxml` |
| Extension writer (`scripts/extend/`) | Python + `lxml` |
| Validation (`scripts/validate.js`) | Node, zero deps |
| Validation (`pro/scripts/validate.py`) | Python + `lxml` + `defusedxml` |
| Optional: render-to-image / `.doc` conversion / accept revisions | LibreOffice (`soffice`) + Poppler (`pdftoppm`) |

## Scope

**Not supported**: PDF, spreadsheets (xlsx), plain-text / Markdown sources (unless the task is to **output** a docx).

**Explicitly out of scope**: structural **re-layout** of layout tables (paragraphs inside cells remain editable), the **body content** of a TOC (the field is generated but must be updated in Word), page-number cross-references, editing inside field codes.

**Supported**: headers / footers (incl. per-section overrides), page setup, footnotes / endnotes, comments, tracked changes.

## License

MIT

## Star History

<picture>
  <source media="(prefers-color-scheme: dark)" srcset="assets/star-history-dark.svg">
  <img src="assets/star-history-light.svg" alt="Zok66/docx-master star history">
</picture>

Updated every 12 hours by [`tools/gen-star-history.py`](tools/gen-star-history.py)
([workflow](.github/workflows/star-history.yml)) from GitHub stargazer timestamps.
Light and dark SVGs are stored in this repo; the curve aggregates currently-retained stars.
