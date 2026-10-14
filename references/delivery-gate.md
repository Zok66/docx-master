# 交付门（Delivery gate）

主控 [`SKILL.md`](../SKILL.md) 只保留 5 步骨架；本文件是完整说明——断言类型清单、变更清单读法、
最后一道人工自检。**交付前按本文件执行**。

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

---

## Assertions encode intent

Write assertions before the config: they are the user's requirements stated
mechanically. Their whole purpose is the defect class no schema check can see:

```json
{ "type": "notContains", "text": "My Book", "scope": "footer" },
{ "type": "footerText", "regex": "^\\s*第\\s*<PAGE>\\s*页\\s*\\/\\s*共\\s*<NUMPAGES>\\s*页\\s*$" },
{ "type": "hasField", "field": "TOC", "scope": "body" },
{ "type": "refsResolve" }
```

The first line exists because a footer carrying an unrequested title is
perfectly well-formed XML that passes every validator. `<PAGE>` is the field
marker: cached field *results* are ignored, so assertions survive repagination
and don't depend on what Word last computed.

`refsResolve` covers the other invisible class: a REF / PAGEREF / `#anchor`
pointing at a bookmark that no longer exists renders as plain text in Word
("Error! Bookmark not defined." mid-sentence, a TOC entry that stops jumping)
and satisfies every schema check on the way out. Pair it with `bookmarkProblems`
(malformed bookmark ranges, duplicate ids) whenever the task creates or edits
anchors.

### 断言类型清单

- **counts** — `sectionCount` / `paragraphCount` / `tableCount` / `imageCount`，配 `eq` / `ne` / `min` / `max`
- **presence** — `contains` / `notContains` / `textMatches`，作用域 `body` | `header` | `footer` | `all`
- **parts** — `headerText` / `footerText`，配 `regex` / `equals` / `index`
- **fields** — `hasField` / `noField`
- **bookmarks** — `refsResolve` / `bookmarkProblems` / `bookmarkExists` / `bookmarkMissing` / `bookmarkCount` / `refCount`
- **setup** — `pageSize` / `partExists`
- **styles** — `styleExists` / `styleMissing`，按 `styleId` 或 `name`
- **metadata** — `author` / `title`

Check styles by `name`, not `styleId`: a writer may serialize ids as opaque
numbers (`2`, `10`, …) — this skill's own `apply` does — so the id is not a
stable identity across runs.

---

## The manifest is read by a person

`verify --diff <baseline>` prints every structural difference the run
introduced: sections, paragraph count, parts, header/footer text, fields,
styles, samples of new body text. It exists to be read, deliberately:

> Every line above is a change this run introduced. Anything here the user did
> not ask for is a defect, not a feature.

---

## The last check is not a program

Before handover, restate the request as a checklist in the user's own terms and
walk the artifact against it: every requested item present, and *nothing present
that was not requested*. Self-added polish — a title in the footer, an extra
cover line, a "helpful" section — is the most common way a technically perfect
file still fails, and it is invisible to every automated check above. When the
manifest shows something you cannot trace back to the user's words, delete it.
