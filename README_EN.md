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

[CHANGELOG](CHANGELOG.md) · [Security](SECURITY.md) · [pro-only ops](references/pro-only-ops.md) · [Regression tests](tests/README.md) · [Contributing](CONTRIBUTING.md)

</div>

> 🤖 **AI / developers: read [SKILL.md](SKILL.md) first** (main flow, hard rules, routing index) — don't scan the whole repo first.

**Author**: Zoros ([Zok66](https://github.com/Zok66)) · MIT-licensed, free to use / modify / sell commercially (just keep the copyright notice)

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

## Privacy & Security

- **Documents are processed entirely locally** — nothing is uploaded to any cloud conversion service
- **Macros / OLE payloads inside documents are never executed**; no network access (`doctor` only probes locally installed tools)
- The original file stays untouched: every write command produces a new file
- When processing **untrusted** docx files, run in an isolated environment (same common sense as opening unknown Word documents)
- Report vulnerabilities privately — see [SECURITY.md](SECURITY.md)

No Microsoft Word dependency, no LibreOffice dependency, no npm `docx` package. Unzip the `.docx`,
edit the XML inside, zip it back — so it can do far more than "what a library API can wrap":
paragraph-level restyling, multi-level auto-numbering, TOC fields, footnotes & endnotes, page borders,
content controls, batch mail-merge; the `pro` branch goes further with watermarks, shapes, canvas,
SmartArt, charts and OLE.

The original file is never modified: every write command produces a new file, validated before it lands.

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

```bash
# Standard (default — what you get on clone)
git clone https://github.com/Zok66/docx-master.git ~/.workbuddy/skills/docx-master

# pro (full-featured)
cd ~/.workbuddy/skills/docx-master && git checkout pro

# back to standard
git checkout main
```

The 8 stripped `extend` ops **still exist in the code** (just not registered in `OPERATIONS`);
calling them fails loudly with the 16 available ops listed — never a silent downgrade.
See [`references/pro-only-ops.md`](references/pro-only-ops.md) for the list and history.

## Version

- **Current: [v0.16.0](https://github.com/Zok66/docx-master/releases/tag/v0.16.0)** (2026-10-20, first standard-edition release).
- **License**: [MIT](LICENSE) — free to use, modify, distribute, and use commercially; keep the copyright and permission notice.
- **Upgrade**: `git pull` inside the skill directory; run `node tests/run.mjs` afterwards to confirm environment contracts.
- **Download**: pick Latest from [Releases](https://github.com/Zok66/docx-master/releases); each release ships a capability list and acceptance record.
- **History**: [Releases](https://github.com/Zok66/docx-master/releases).

What changed vs previous releases:

| Previous release | Highlights of this one |
|---|---|
| v0.15.0 / rounds 13–15 | 31 run/paragraph format fields + pinyin annotation (`w:ruby`) engineized; WPS "Insert" tab alignment (docPartObj content controls + glossary cover layout); three comment-anchor fixes |
| v0.15.0-pro | Full-featured edition frozen into the `pro` branch: watermarks, shapes, canvas, SmartArt, charts, OLE, building blocks, custom XML |
| Round 10 and earlier | Delivery-gate false positives / XSD leaks / silent failures fixed, skill restructure, bookmarks & cross-references, `batch_fill` mail-merge |

Full history in [`CHANGELOG.md`](CHANGELOG.md).

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
| **Extend** | `extend` fills what the engine deliberately doesn't write: TOC / table-of-figures / index fields, page & section breaks, page numbering, named bookmarks, footnotes & endnotes, page borders & background, images, content controls, cover pages, citations. Watermarks / shapes / charts / OLE / SmartArt / building blocks / custom XML are pro-only |
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
git clone https://github.com/Zok66/docx-master.git ~/.workbuddy/skills/docx-master
```

## Quick start

```bash
SK=~/.workbuddy/skills/docx-master

# 0) Before anything: probe environment capabilities (what's really available, what degrades)
node $SK/scripts/doctor.js

# 1) Recon first: structure, styles, numbering, skeleton
node $SK/scripts/overview.js my-doc.docx
node $SK/scripts/inspect_bookmarks.js my-doc.docx   # bookmarks + references + dangling ones (exit 1 on problems)

# 2) Single-step design: one config — dry-run first, then write (new file; original untouched)
node $SK/scripts/apply.js --dry-run my.config.json
node $SK/scripts/apply.js my.config.json
python $SK/scripts/extend/extend.py extend.config.json   # TOC / notes / charts …

# 3) Or: declare the whole chain and run it in one command (recommended)
node $SK/scripts/pipeline.js build.json          # aborts on first failure, auto-validates at the end
node $SK/scripts/pipeline.js build.json --plan   # just show what would run, write nothing

# 4) Delivery gate: assertions + change manifest
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

- **The original file is never modified.** Every write command produces a new copy; the source stays untouched.
- **Sparse by design.** Declare only what's broken, missing, or explicitly requested; untouched styles / numbering / paragraphs stay as they are.
- **Validate before delivering.** Two validators check different layers (Node schema vs XSD + OPC package graph); both must pass.
- **Multi-step chains must be machine-orchestrated.** Relying on a human to remember "re-run the config source after editing" is how stale intermediates ship: the writer uses the **previous config** while every later validator reports PASS. Declare the chain once, run it once.
- **"Schema-valid" ≠ "deliverable".** Unrequested changes (an extra title line in the footer, an extra signature) are valid XML — and defects. Assertions handle the mechanical half; reading the change manifest handles the rest.
- **Invisible objects need dedicated checks.** Bookmarks render nothing, carry no style, add no word count: lose one and the file still opens, still passes XSD — until Word prints "Error! Bookmark not defined." mid-sentence. `inspect_bookmarks` (human) and `refsResolve` (machine) are the only two places that look.
- **Content and structure are separate.** Structural signals — headings, lists, captions — are rendered by styles, numbering and SEQ fields, never typed as literal text.
- **The master file holds rules only; details load on demand.** `SKILL.md` enters context in full on every trigger, so it carries only what must be known immediately: main flow, hard rules, routing index. Long tables and semantic detail live in `references/`, read when needed. Every rule is defined in **exactly one** place; everything else points at it — `tests/run.mjs` guards this with a contract list: rules may move files, but must never vanish.

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

Standard and pro share one codebase; the only difference is the `OPERATIONS` registry in
`scripts/extend/extend.py` (16 vs 24 ops). This section describes the **standard edition** (`main`).

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
