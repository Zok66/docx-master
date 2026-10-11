import { _ as parseXml, a as firstChildNS, c as paragraphRuns, d as textContent, f as wAttr, g as DocxReader, h as NS, i as descendantsNS, l as paragraphStyleId, m as walkBodyParagraphs, o as getChildren, p as wVal, r as buildPlainTextRun, s as getChildrenNS, t as addVanishRPr, v as serializeXml, y as require_lib } from "./_shared/xml-utils.js";
import { n as StyleResolver, r as applyThemeFontOverrides, t as DocumentParser } from "./_shared/document-parser.js";
import { t as summarizeTable } from "./_shared/table-classifier.js";
import { t as Fingerprinter } from "./_shared/fingerprint.js";
import { a as resolveRunLocator, i as resolveLocator, l as assertNever, o as trailingBodySectPr, s as walkIndexedParagraphs, t as buildResolverContext, u as makeTrackContext } from "./_shared/locator.js";
import { n as sectionUsableWidthTwips, t as sectionForParagraph } from "./_shared/section-metrics.js";
import { n as seqFields, t as parseFieldRuns } from "./_shared/field-parse.js";
import { n as explainBlockerReason, r as summarizeBlockers, t as detectBlockers } from "./_shared/blockers.js";
import { n as validateOMath, t as validateDocxFile } from "./_shared/docx-validate.js";
import { mergeRunsInDocument } from "./_shared/merge-runs.js";
import { dirname, extname, join, resolve } from "node:path";
import { copyFileSync, existsSync, mkdirSync, readFileSync, unlinkSync } from "node:fs";
import { fileURLToPath, pathToFileURL } from "node:url";

//#region lib/xml/writable-archive.ts
/**
* Writer-bound replacement map with an enforced single-writer invariant
* for the OOXML package's coordination parts.
*
* Background: `[Content_Types].xml` and `word/_rels/document.xml.rels`
* are shared parts that multiple apply subsystems would naively want to
* mutate — numbering registration, settings fabrication, header / footer
* Override entries, image rels, hyperlink rels, etc. Letting each
* subsystem read-modify-write the part text independently produces a
* silent last-writer-wins race: subsystem A reads the source bytes,
* appends entry X, writes the result back; subsystem B reads the SAME
* source bytes, appends entry Y, writes the result back, **dropping X**.
* This actually shipped (the chineseCounting numbering scheme being
* silently dropped when the body asset registry's flushTo ran after
* `ensureNumberingContentType` — discovered during the blank-source demo).
*
* The fix is structural: both parts have exactly one writer — the
* accumulator owned by `bodyAssetRegistry` (`ContentTypes` for the
* content-types file, body `PartRels` for the doc rels file). Every
* subsystem that wants to register a new Default / Override / Relationship
* routes through those accumulators; one final `flushTo` at end of apply
* serialises the accumulated state.
*
* This wrapper enforces the invariant at the API boundary: a runtime
* `throw` on any attempt to `.set()` either of the two coordination
* paths from outside the accumulator. The accumulators themselves use
* `setFromAccumulator()` (intentionally underscore-prefixed so grep for
* `replacements.set(` doesn't false-positive on legitimate flushTo
* calls).
*/
const FORBIDDEN_DIRECT_WRITES = new Set(["[Content_Types].xml", "word/_rels/document.xml.rels"]);
var WritableArchive = class {
	inner = /* @__PURE__ */ new Map();
	/** Stage `content` at `path` in the archive. Throws on the two
	*  coordination paths — those must go through the shared accumulators
	*  owned by `bodyAssetRegistry`. */
	set(path, content) {
		if (FORBIDDEN_DIRECT_WRITES.has(path)) throw new Error(`WritableArchive.set: direct writes to "${path}" are forbidden — multiple subsystems otherwise race-overwrite each other. Route the entry through bodyAssetRegistry.getContentTypes() (for Override / Default entries) or bodyAssetRegistry.getPartRels() (for body Relationships). See lib/xml/writable-archive.ts.`);
		this.inner.set(path, content);
	}
	/** Escape hatch for the two legitimate writers of forbidden paths:
	*  `ContentTypes.flushTo` and `PartRels.flushTo` (when its target path
	*  is `word/_rels/document.xml.rels`). The underscore prefix flags
	*  this as internal so it stays grep-visible during reviews. */
	setFromAccumulator(path, content) {
		this.inner.set(path, content);
	}
	get(path) {
		return this.inner.get(path);
	}
	has(path) {
		return this.inner.has(path);
	}
	get size() {
		return this.inner.size;
	}
	/** Underlying Map handed to `DocxReader.copyAndModify`. Read-only by
	*  convention — callers should not mutate it directly. */
	toMap() {
		return this.inner;
	}
	[Symbol.iterator]() {
		return this.inner.entries();
	}
};

//#endregion
//#region lib/apply/template-import.ts
/**
* Cross-document style import: copy `<w:style>` definitions from a template
* docx into the target's styles.xml, transitively pulling in basedOn
* ancestors and migrating any numbering references so the imported styles
* actually work in the new document context.
*
* Why this lives outside apply-styles.ts: the logic is self-contained and
* has its own data flow (open template, walk styles + numbering, mutate
* target docs). Keeping it isolated keeps apply-styles' main pipeline
* readable.
*/
/**
* Imports the named styles from `templatePath` into the existing
* `targetStylesDoc`, also injecting any required abstractNum/num entries
* into `targetNumberingDoc`. Mutates both target docs in place.
*/
async function importTemplateStyles(templatePath, styleIds, targetStylesDoc, targetNumberingDoc, options = {}) {
	const importNumbering = options.importNumbering ?? true;
	const reader = await DocxReader.open(templatePath);
	const tplStylesDoc = await reader.readXml("word/styles.xml");
	if (!tplStylesDoc) throw new Error(`Template ${templatePath} has no word/styles.xml`);
	const tplNumberingDoc = importNumbering ? await reader.readXml("word/numbering.xml") : null;
	const tplStyles = collectStyles(tplStylesDoc);
	const targetStyles = collectStyles(targetStylesDoc);
	const toImport = /* @__PURE__ */ new Set();
	const pulledAncestors = /* @__PURE__ */ new Set();
	const queue = [...styleIds];
	while (queue.length > 0) {
		const id = queue.shift();
		if (toImport.has(id)) continue;
		const tpl = tplStyles.get(id);
		if (!tpl) throw new Error(`template.styles: "${id}" is not defined in ${templatePath}'s styles.xml.\n  Available: [${[...tplStyles.keys()].sort().join(", ")}]`);
		toImport.add(id);
		const basedOn = firstChildNS(tpl, NS.w, "basedOn");
		if (basedOn) {
			const parent = wAttr(basedOn, "val");
			if (parent && !targetStyles.has(parent) && !toImport.has(parent)) {
				if (!isBuiltinAlwaysPresent(parent)) {
					pulledAncestors.add(parent);
					queue.push(parent);
				}
			}
		}
	}
	const numIdRemap = /* @__PURE__ */ new Map();
	if (importNumbering && tplNumberingDoc) {
		const referencedNumIds = /* @__PURE__ */ new Set();
		for (const id of toImport) {
			const pPr = firstChildNS(tplStyles.get(id), NS.w, "pPr");
			if (!pPr) continue;
			const numPr = firstChildNS(pPr, NS.w, "numPr");
			if (!numPr) continue;
			const numIdEl = firstChildNS(numPr, NS.w, "numId");
			if (!numIdEl) continue;
			const v = wAttr(numIdEl, "val");
			if (v) referencedNumIds.add(v);
		}
		for (const oldNumId of referencedNumIds) {
			const fresh = migrateNumIdToTarget(oldNumId, tplNumberingDoc, targetNumberingDoc);
			if (fresh !== null) numIdRemap.set(oldNumId, fresh);
		}
	}
	for (const id of toImport) upsertStyleNode(targetStylesDoc, id, importStyleNode(tplStyles.get(id), targetStylesDoc, numIdRemap));
	return {
		imported: [...toImport],
		numIdRemap,
		pulledAncestors: [...pulledAncestors]
	};
}
function collectStyles(stylesDoc) {
	const out = /* @__PURE__ */ new Map();
	const root = stylesDoc.documentElement;
	if (!root) return out;
	for (const s of getChildrenNS(root, NS.w, "style")) {
		const id = wAttr(s, "styleId");
		if (id) out.set(id, s);
	}
	return out;
}
function isBuiltinAlwaysPresent(styleId) {
	return styleId === "Normal" || styleId === "DefaultParagraphFont";
}
/**
* Clone a <w:style> element from the template document into the target,
* remapping any numPr/numId references via numIdRemap.
*/
function importStyleNode(src, targetDoc, numIdRemap) {
	const cloned = deepCloneIntoDocument(src, targetDoc);
	const w = NS.w;
	const visit = (el) => {
		if (el.namespaceURI === w && el.localName === "numId") {
			const v = wAttr(el, "val");
			if (v && numIdRemap.has(v)) el.setAttributeNS(w, "w:val", numIdRemap.get(v));
		}
		for (const c of getChildren(el)) visit(c);
	};
	visit(cloned);
	return cloned;
}
function deepCloneIntoDocument(src, targetDoc) {
	const ns = src.namespaceURI;
	const tag = src.tagName;
	const out = ns ? targetDoc.createElementNS(ns, tag) : targetDoc.createElement(tag);
	if (src.attributes) for (let i = 0; i < src.attributes.length; i++) {
		const a = src.attributes.item(i);
		if (a.namespaceURI) out.setAttributeNS(a.namespaceURI, a.name, a.value);
		else out.setAttribute(a.name, a.value);
	}
	for (const child of Array.from(src.childNodes)) if (child.nodeType === 1) out.appendChild(deepCloneIntoDocument(child, targetDoc));
	else if (child.nodeType === 3 || child.nodeType === 4) out.appendChild(targetDoc.createTextNode(child.nodeValue || ""));
	return out;
}
/**
* Replace any existing <w:style> with the same styleId, otherwise append.
*/
function upsertStyleNode(stylesDoc, styleId, node) {
	const root = stylesDoc.documentElement;
	const existing = getChildrenNS(root, NS.w, "style").find((s) => wAttr(s, "styleId") === styleId);
	if (existing) root.replaceChild(node, existing);
	else root.appendChild(node);
}
/**
* Copy the abstractNum + num pair referenced by `oldNumId` from the template
* numbering doc into the target. Returns the new numId in the target doc,
* or null if the template doesn't actually have that numId.
*/
function migrateNumIdToTarget(oldNumId, tplNumberingDoc, targetNumberingDoc) {
	const w = NS.w;
	const tplRoot = tplNumberingDoc.documentElement;
	const targetRoot = targetNumberingDoc.documentElement;
	const tplNum = getChildrenNS(tplRoot, w, "num").find((n) => wAttr(n, "numId") === oldNumId);
	if (!tplNum) return null;
	const absRef = firstChildNS(tplNum, w, "abstractNumId");
	if (!absRef) return null;
	const oldAbsId = wVal(absRef);
	if (!oldAbsId) return null;
	const tplAbs = getChildrenNS(tplRoot, w, "abstractNum").find((a) => wAttr(a, "abstractNumId") === oldAbsId);
	if (!tplAbs) return null;
	const existingAbsIds = getChildrenNS(targetRoot, w, "abstractNum").map((e) => parseInt(wAttr(e, "abstractNumId") || "0", 10));
	const existingNumIds = getChildrenNS(targetRoot, w, "num").map((e) => parseInt(wAttr(e, "numId") || "0", 10));
	const newAbsId = (existingAbsIds.length ? Math.max(...existingAbsIds) : -1) + 1;
	const newNumId = (existingNumIds.length ? Math.max(...existingNumIds) : 0) + 1;
	const clonedAbs = deepCloneIntoDocument(tplAbs, targetNumberingDoc);
	clonedAbs.setAttributeNS(w, "w:abstractNumId", String(newAbsId));
	const firstNum = getChildrenNS(targetRoot, w, "num")[0];
	if (firstNum) targetRoot.insertBefore(clonedAbs, firstNum);
	else targetRoot.appendChild(clonedAbs);
	const clonedNum = deepCloneIntoDocument(tplNum, targetNumberingDoc);
	clonedNum.setAttributeNS(w, "w:numId", String(newNumId));
	const newAbsRef = firstChildNS(clonedNum, w, "abstractNumId");
	if (newAbsRef) newAbsRef.setAttributeNS(w, "w:val", String(newAbsId));
	targetRoot.appendChild(clonedNum);
	return String(newNumId);
}

//#endregion
//#region lib/parse/builtin-styles.ts
/**
* Word built-in style names that Word translates between locales.
*
* `<w:name w:val="heading 1"/>` is the *English canonical* form. zh-CN
* Word UI renders the same style as "标题 1"; German as "Überschrift 1";
* etc. Word resolves field codes (STYLEREF, REF, etc.) against the
* *locale-translated* form — so `STYLEREF "heading 1"` fails to match
* in any non-EN Word, even though the underlying styles.xml carries
* the English name.
*
* Two consumers in the apply pipeline care about this list:
*   - style-name collision detection (apply-styles.ts preflight) treats
*     locale-aliased pairs as the same built-in identity to avoid the
*     "two style entries with the same effective name" failure
*   - STYLEREF emit (`emitInlineStyleRef`) routes built-in heading
*     targets through the locale-neutral `STYLEREF N` form, and refuses
*     to emit `STYLEREF "<name>"` for built-in *non-heading* styles
*     (Title / Caption / Subtitle / etc.) — that form silently fails
*     in non-EN Word at render time
*
* Maintenance: the list of built-in names is fixed by Word's template
* (`normal.dotm` ships ~250 styles, ~30 of which are routinely
* referenced and locale-translated). zh-CN is the only locale we
* currently track since it's the only one in production use; new
* locale pairs are additive and don't require code changes.
*/
/** Built-in style English canonical name → zh-CN localized form. */
const ENGLISH_TO_ZH_CN = [
	["Normal", "正文"],
	["heading 1", "标题 1"],
	["heading 2", "标题 2"],
	["heading 3", "标题 3"],
	["heading 4", "标题 4"],
	["heading 5", "标题 5"],
	["heading 6", "标题 6"],
	["heading 7", "标题 7"],
	["heading 8", "标题 8"],
	["heading 9", "标题 9"],
	["Title", "标题"],
	["Subtitle", "副标题"],
	["Body Text", "正文文本"],
	["Caption", "题注"],
	["Quote", "引用"],
	["List Bullet", "列表项目符号"],
	["List Number", "列表编号"],
	["Header", "页眉"],
	["Footer", "页脚"],
	["TOC 1", "目录 1"],
	["TOC 2", "目录 2"],
	["TOC 3", "目录 3"],
	["Default Paragraph Font", "默认段落字体"],
	["Normal Table", "普通表格"],
	["No List", "无列表"]
];
/** All names (English canonical + every locale alias) that resolve to
*  a Word built-in style. Used to detect when an agent-emitted STYLEREF
*  argument would silently fail in a non-EN Word locale. */
const BUILT_IN_NAMES = new Set(ENGLISH_TO_ZH_CN.flatMap(([eng, zh]) => [eng, zh]));
/** Canonical-name map: any locale form → English canonical form. Used
*  by collision detection to treat "Normal" and "正文" as the same
*  built-in identity. */
const TO_CANONICAL = new Map(ENGLISH_TO_ZH_CN.flatMap(([eng, zh]) => [[eng, eng], [zh, eng]]));
/** Map a style name to its English canonical form when it's a known
*  built-in identity; otherwise return the input unchanged. Used for
*  comparing two names to see if they'd collide on Word's built-in
*  identity. */
function canonicalStyleName(name) {
	return TO_CANONICAL.get(name) ?? name;
}
/** Returns true when the name matches a built-in Word style whose
*  display form differs across Word UI languages. Locale-aliased pairs
*  match by either side (English "Title" or Chinese "标题"). */
function isBuiltInLocalizable(name) {
	return BUILT_IN_NAMES.has(name);
}

//#endregion
//#region lib/xml/docx-plumbing.ts
function blankStylesDoc() {
	return parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="${NS.w}"></w:styles>`);
}
function blankNumberingDoc() {
	return parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering xmlns:w="${NS.w}"></w:numbering>`);
}

//#endregion
//#region lib/apply/para-mutation.ts
function applyToBody(documentDoc, ctx) {
	const root = documentDoc.documentElement;
	const body = firstChildNS(root, NS.w, "body");
	if (!body) return;
	const byIdx = /* @__PURE__ */ new Map();
	for (const p of ctx.paragraphs) byIdx.set(p.index, p);
	let nextIdx = 1;
	const traverseChildren = (parentEl, _insideLayout) => {
		const children = getChildren(parentEl);
		for (const child of children) {
			if (child.namespaceURI !== NS.w) continue;
			if (child.localName === "p") {
				const idx = nextIdx++;
				const para = byIdx.get(idx);
				if (para) processOneParagraph(child, para, ctx);
			} else if (child.localName === "tbl") {
				if (summarizeTable(child).classification === "layout") for (const tr of getChildrenNS(child, NS.w, "tr")) for (const tc of getChildrenNS(tr, NS.w, "tc")) traverseChildren(tc, true);
			}
		}
	};
	traverseChildren(body, false);
}
function processOneParagraph(pEl, para, ctx) {
	if (ctx.excludeSet.has(para.index)) return;
	const a = ctx.assignmentMap.get(para.index);
	let action = "keep";
	let targetStyle;
	let reason;
	let matchedPattern = null;
	if (a) {
		action = a.action;
		targetStyle = a.style;
		reason = a.reason;
	} else {
		for (const rule of ctx.patternRules) {
			const m = para.text.match(rule.regex);
			if (m && m.index === 0) {
				action = "restyle";
				targetStyle = rule.style;
				matchedPattern = {
					rule,
					matchLen: m[0].length
				};
				break;
			}
		}
		if (!targetStyle) {
			const bulkStyle = ctx.bulkMap.get(para.fingerprint);
			if (bulkStyle) {
				action = "restyle";
				targetStyle = bulkStyle;
			}
		}
	}
	if (action === "flag" && reason) {
		ctx.flags.push({
			paraIndex: para.index,
			reason
		});
		return;
	}
	if (action !== "restyle" || !targetStyle) {
		if (!a) {
			if (ctx.editTouchedIndices?.has(para.index)) return;
			const isEmpty = para.text.trim().length === 0;
			const cur = ctx.implicitKeepByFingerprint.get(para.fingerprint) ?? {
				empty: 0,
				nonEmpty: 0,
				nonEmptySamples: []
			};
			if (isEmpty) cur.empty += 1;
			else {
				cur.nonEmpty += 1;
				if (cur.nonEmptySamples.length < 2) {
					const snippet = para.text.trim().slice(0, 30);
					cur.nonEmptySamples.push(snippet + (para.text.trim().length > 30 ? "…" : ""));
				}
			}
			ctx.implicitKeepByFingerprint.set(para.fingerprint, cur);
		}
		return;
	}
	const oldPStyle = para.styleId;
	setParagraphStyle(pEl, targetStyle);
	stripConflictingDirectFormatting(pEl, ctx.stylePPrCascade.get(targetStyle) ?? /* @__PURE__ */ new Set(), ctx.styleRPrCascade.get(targetStyle) ?? /* @__PURE__ */ new Set());
	ctx.restyleStats.set(targetStyle, (ctx.restyleStats.get(targetStyle) ?? 0) + 1);
	const existingSamples = ctx.samples.get(targetStyle) ?? [];
	let thisSample = null;
	if (existingSamples.length < ctx.samplesPerStyleCap) {
		const via = a ? "assignment" : matchedPattern ? "pattern" : "bulk";
		thisSample = {
			paraIndex: para.index,
			oldStyle: oldPStyle,
			newStyle: targetStyle,
			textPreview: para.text.slice(0, 60) + (para.text.length > 60 ? "…" : ""),
			via,
			patternSource: matchedPattern?.rule.source,
			notes: []
		};
		existingSamples.push(thisSample);
		ctx.samples.set(targetStyle, existingSamples);
	}
	if (matchedPattern) {
		const key = matchedPattern.rule.source;
		ctx.patternMatchStats.set(key, (ctx.patternMatchStats.get(key) ?? 0) + 1);
		if (matchedPattern.rule.stripMatch) {
			if (removeRegexPrefix(pEl, matchedPattern.rule.regex)) {
				ctx.patternStripStats.set(key, (ctx.patternStripStats.get(key) ?? 0) + 1);
				if (thisSample) thisSample.notes.push(`stripped pattern /${key}/`);
			}
		}
	}
	const lvlPatterns = ctx.numLvlTextByStyle.get(targetStyle);
	if (lvlPatterns && lvlPatterns.length > 0) {
		let stripped = false;
		for (const pat of lvlPatterns) if (removeManualNumberingPrefix(pEl, pat)) {
			ctx.manualNumberingRemoved.set(pat, (ctx.manualNumberingRemoved.get(pat) ?? 0) + 1);
			let perStyle = ctx.manualNumberingByStyle.get(targetStyle);
			if (!perStyle) {
				perStyle = /* @__PURE__ */ new Map();
				ctx.manualNumberingByStyle.set(targetStyle, perStyle);
			}
			perStyle.set(pat, (perStyle.get(pat) ?? 0) + 1);
			if (thisSample) thisSample.notes.push(`stripped manual prefix "${pat}"`);
			stripped = true;
			break;
		}
		const alreadyStrippedByPatternRule = !!matchedPattern?.rule.stripMatch;
		if (!stripped && !alreadyStrippedByPatternRule) {
			let perStyle = ctx.unstrippedByStyle.get(targetStyle);
			if (!perStyle) {
				perStyle = {
					count: 0,
					samples: []
				};
				ctx.unstrippedByStyle.set(targetStyle, perStyle);
			}
			perStyle.count += 1;
			if (perStyle.samples.length < 3) {
				const snippet = para.text.trim().slice(0, 40);
				perStyle.samples.push(snippet + (para.text.trim().length > 40 ? "…" : ""));
			}
		}
	}
}
function setParagraphStyle(pEl, styleId) {
	const w = NS.w;
	let pPr = firstChildNS(pEl, w, "pPr");
	if (!pPr) {
		pPr = pEl.ownerDocument.createElementNS(w, "w:pPr");
		pEl.insertBefore(pPr, pEl.firstChild);
	}
	let pStyle = firstChildNS(pPr, w, "pStyle");
	if (!pStyle) {
		pStyle = pEl.ownerDocument.createElementNS(w, "w:pStyle");
		pPr.insertBefore(pStyle, pPr.firstChild);
	}
	pStyle.setAttributeNS(w, "w:val", styleId);
}
const RPR_CONFLICT_NAMES = [
	"rFonts",
	"sz",
	"szCs",
	"b",
	"bCs",
	"i",
	"iCs",
	"color",
	"vertAlign"
];
function stripConflictingDirectFormatting(pEl, stylePPrProvides, styleRPrProvides) {
	const w = NS.w;
	const pPr = firstChildNS(pEl, w, "pPr");
	if (pPr) {
		const directConflicts = new Set([
			"jc",
			"spacing",
			"ind",
			"outlineLvl"
		]);
		for (const c of Array.from(getChildren(pPr))) if (c.namespaceURI === w && directConflicts.has(c.localName) && stylePPrProvides.has(c.localName)) pPr.removeChild(c);
		const paraRPr = firstChildNS(pPr, w, "rPr");
		if (paraRPr) {
			removeRPrConflicts(paraRPr, styleRPrProvides);
			if (getChildren(paraRPr).length === 0) pPr.removeChild(paraRPr);
		}
	}
	const runs = getChildrenNS(pEl, w, "r");
	if (runs.length === 0) return;
	const valuesByProp = /* @__PURE__ */ new Map();
	for (const name of RPR_CONFLICT_NAMES) valuesByProp.set(name, /* @__PURE__ */ new Set());
	for (const r of runs) {
		const rPr = firstChildNS(r, w, "rPr");
		for (const name of RPR_CONFLICT_NAMES) {
			const child = rPr ? firstChildNS(rPr, w, name) : null;
			valuesByProp.get(name).add(child ? rPrChildSignature(child, name) : "<absent>");
		}
	}
	const uniformToStrip = /* @__PURE__ */ new Set();
	for (const [name, vals] of valuesByProp) if (vals.size <= 1 && styleRPrProvides.has(name)) uniformToStrip.add(name);
	for (const r of runs) {
		const rPr = firstChildNS(r, w, "rPr");
		if (!rPr) continue;
		for (const c of Array.from(getChildren(rPr))) if (c.namespaceURI === w && uniformToStrip.has(c.localName)) rPr.removeChild(c);
		if (getChildren(rPr).length === 0) r.removeChild(rPr);
	}
}
function removeRPrConflicts(rPr, styleProvides) {
	const w = NS.w;
	const conflicts = new Set(RPR_CONFLICT_NAMES);
	for (const c of Array.from(getChildren(rPr))) if (c.namespaceURI === w && conflicts.has(c.localName) && styleProvides.has(c.localName)) rPr.removeChild(c);
}
function rPrChildSignature(el, name) {
	if (name === "rFonts") return [
		"ascii",
		"hAnsi",
		"eastAsia",
		"cs"
	].map((a) => `${a}=${wAttr(el, a) ?? ""}`).join("|");
	if (name === "b" || name === "bCs" || name === "i" || name === "iCs") {
		const v = wAttr(el, "val");
		return v === "0" || v === "false" ? "off" : "on";
	}
	return wAttr(el, "val") ?? "";
}
/**
* Replace a leading match in the paragraph's leading text with the regex
* stripped. Returns true on a hit. Accumulates text from consecutive w:t
* runs so that prefixes split across runs (e.g. "五" + "、Title", which
* Word emits when a paragraph was hand-edited) still match. Removed
* characters are distributed across the runs they came from; subsequent
* runs are emptied as needed and the first remaining run keeps the tail.
*/
function stripLeadingMatch(pEl, re) {
	const w = NS.w;
	const tEls = [];
	let combined = "";
	for (const run of getChildrenNS(pEl, w, "r")) {
		const tEl = firstChildNS(run, w, "t");
		if (!tEl) continue;
		tEls.push(tEl);
		combined += textContent(tEl);
		const m = combined.match(re);
		if (!m) continue;
		let remaining = m[0].length;
		for (const t of tEls) {
			const txt = textContent(t);
			while (t.firstChild) t.removeChild(t.firstChild);
			if (remaining >= txt.length) {
				remaining -= txt.length;
				t.appendChild(t.ownerDocument.createTextNode(""));
			} else {
				t.appendChild(t.ownerDocument.createTextNode(txt.slice(remaining)));
				remaining = 0;
			}
			t.setAttribute("xml:space", "preserve");
		}
		return true;
	}
	return false;
}
function removeRegexPrefix(pEl, regex) {
	return stripLeadingMatch(pEl, regex.source.startsWith("^") ? regex : new RegExp("^" + regex.source, regex.flags));
}
function removeManualNumberingPrefix(pEl, lvlText) {
	const pattern = lvlText.replace(/[.*+?^${}()|[\]\\]/g, (m) => "\\" + m).replace(/%\d/g, "(?:\\d+|[一二三四五六七八九十百千]+)");
	return stripLeadingMatch(pEl, new RegExp("^\\s*" + pattern + "\\s*"));
}

//#endregion
//#region lib/xml/xml-order.ts
/**
* Schema-correct insertion helpers for OOXML mutation paths.
*
* The engine builds elements with a per-call convenience order (e.g.
* `pPrAdditions` pushes spacing → ind → jc → outlineLvl in upsertStyle, the
* order they're constructed in code). CT_PPr's actual schema order is
* different (spacing → ind → jc → outlineLvl per ECMA-376). Inserting
* blindly with `appendChild` produces schema-violating XML that Word
* rejects with "needs repair".
*
* `insertChildInOrder` places a new child at its schema-correct position
* relative to existing children, given an explicit ordered list of localNames.
*
* Kept separate from docx-validate.ts so the engine can use it without
* importing the validator.
*/
/** EG_RPrBase child order, per ECMA-376 17.7.9.1 (children of CT_RPr inside
* style / run / paragraph mark). Truncated to the children this engine
* actually emits. Notably: `b` / `bCs` come BEFORE `sz` / `szCs` — Word's
* loader rejects mis-ordered run properties with "file needs repair" even
* though XSD validators (xmllint) often pass them via relaxed xs:all. */
const RPR_CHILD_ORDER = [
	"rStyle",
	"rFonts",
	"b",
	"bCs",
	"i",
	"iCs",
	"caps",
	"smallCaps",
	"strike",
	"dstrike",
	"outline",
	"shadow",
	"emboss",
	"imprint",
	"noProof",
	"snapToGrid",
	"vanish",
	"webHidden",
	"color",
	"spacing",
	"w",
	"kern",
	"position",
	"sz",
	"szCs",
	"highlight",
	"u",
	"effect",
	"bdr",
	"shd",
	"fitText",
	"vertAlign",
	"rtl",
	"cs",
	"em",
	"lang",
	"eastAsianLayout",
	"specVanish"
];
/** CT_TblPr child order, per ECMA-376 17.4.62. Truncated to children the
* engine emits. Out-of-order writes (e.g. `tblBorders` before `jc`)
* produce "needs repair" warnings on open. */
const TBL_PR_CHILD_ORDER = [
	"tblStyle",
	"tblpPr",
	"tblOverlap",
	"bidiVisual",
	"tblStyleRowBandSize",
	"tblStyleColBandSize",
	"tblW",
	"jc",
	"tblCellSpacing",
	"tblInd",
	"tblBorders",
	"shd",
	"tblLayout",
	"tblCellMar",
	"tblLook",
	"tblCaption",
	"tblDescription",
	"tblPrChange"
];
/** CT_TcPr child order, per ECMA-376 17.4.69. Truncated to children the
* engine emits. `gridSpan` and `vMerge` come BEFORE `tcBorders` — Word
* loaders that re-render mis-ordered cells silently break merging. */
const TC_PR_CHILD_ORDER = [
	"cnfStyle",
	"tcW",
	"gridSpan",
	"hMerge",
	"vMerge",
	"tcBorders",
	"shd",
	"noWrap",
	"tcMar",
	"textDirection",
	"tcFitText",
	"vAlign",
	"hideMark",
	"headers",
	"tcPrChange"
];
/** CT_TrPr child order, per ECMA-376 17.4.82. Truncated. */
const TR_PR_CHILD_ORDER = [
	"cnfStyle",
	"divId",
	"gridBefore",
	"gridAfter",
	"wBefore",
	"wAfter",
	"cantSplit",
	"trHeight",
	"tblHeader",
	"tblCellSpacing",
	"jc",
	"hidden",
	"ins",
	"del",
	"trPrChange"
];
/** CT_SectPr child order, per ECMA-376 17.6.18. Covers every sectPr
* child written by an apply subsystem — `page-setup-mutation` writes
* pgSz / pgMar / cols; `header-footer-mutation` writes headerReference
* / footerReference / titlePg. Other sectPr children present on source
* (type / lnNumType / pgNumType / formProt / vAlign / noEndnote /
* textDirection / bidi / rtlGutter / docGrid / printerSettings /
* sectPrChange) are preserved as-is and listed for completeness so
* `insertChildInOrder` correctly positions new children relative to
* them. */
const SECT_PR_CHILD_ORDER = [
	"headerReference",
	"footerReference",
	"footnotePr",
	"endnotePr",
	"type",
	"pgSz",
	"pgMar",
	"paperSrc",
	"pgBorders",
	"lnNumType",
	"pgNumType",
	"cols",
	"formProt",
	"vAlign",
	"noEndnote",
	"titlePg",
	"textDirection",
	"bidi",
	"rtlGutter",
	"docGrid",
	"printerSettings",
	"sectPrChange"
];
/** CT_PPr child order, per ECMA-376 17.3.1.26. Truncated to children we
* encounter when building or mutating paragraph properties. */
const PPR_CHILD_ORDER = [
	"pStyle",
	"keepNext",
	"keepLines",
	"pageBreakBefore",
	"framePr",
	"widowControl",
	"numPr",
	"suppressLineNumbers",
	"pBdr",
	"shd",
	"tabs",
	"suppressAutoHyphens",
	"kinsoku",
	"wordWrap",
	"overflowPunct",
	"topLinePunct",
	"autoSpaceDE",
	"autoSpaceDN",
	"bidi",
	"adjustRightInd",
	"snapToGrid",
	"spacing",
	"ind",
	"contextualSpacing",
	"mirrorIndents",
	"suppressOverlap",
	"jc",
	"textDirection",
	"textAlignment",
	"textboxTightWrap",
	"outlineLvl",
	"divId",
	"cnfStyle",
	"rPr"
];
/**
* Insert `newChild` at its schema-correct position per `expectedOrder`.
* Children whose localName is later in the order go after; earlier ones go
* before. Falls back to append if `newChild`'s localName isn't in the list.
*/
function insertChildInOrder(parent, newChild, expectedOrder, ns = NS.w) {
	const newName = newChild.localName;
	const newIdx = expectedOrder.indexOf(newName);
	if (newIdx < 0) {
		parent.appendChild(newChild);
		return;
	}
	for (const sibling of getChildren(parent)) {
		if (sibling.namespaceURI !== ns) continue;
		if (expectedOrder.indexOf(sibling.localName) > newIdx) {
			parent.insertBefore(newChild, sibling);
			return;
		}
	}
	parent.appendChild(newChild);
}

//#endregion
//#region lib/shared/units.ts
/**
* Single source of truth for length-unit parsing and conversion.
*
* Agent-facing API exposes pt / cm / mm / in via string suffixes
* (`"12pt" / "2.54cm" / "5mm" / "1in"`). Bare numbers are pt. Internal
* OOXML units (twips, half-pt, EMU, eighth-pt) live behind dedicated
* conversion helpers — callers never multiply by 20 / 2 / 12700 / 8 directly.
*
* Two narrower types extend the basic Length:
*   - `IndentValue` adds `"Nchar"` (round-trips Word's `w:firstLineChars` /
*     `w:hangingChars`, auto-scales with font size).
*   - `LineSpacingValue` carries three Word `lineRule` modes via type
*     discrimination: number = multiplier (auto), Length-string = exact
*     fixed height, `{ atLeast: Length }` = at-least fixed height.
*     The previous magnitude heuristic (`value >= 10 → exact`) is gone.
*/
const TWIPS_PER_PT = 20;
const HALF_PT_PER_PT = 2;
const EIGHTH_PT_PER_PT = 8;
const EMU_PER_PT = 12700;
const TWENTIETHS_PER_LINE = 240;
const PT_PER_CM = 28.3464566929;
const PT_PER_MM = 2.83464566929;
const PT_PER_IN = 72;
const SUPPORTED_LENGTH_UNITS = "pt, cm, mm, in";
/** Parse a Length input to pt. Bare number is pt; strings carry one of
*  pt / cm / mm / in. Throws with a clear message on bad input. */
function parseLengthPt(v, fieldName = "length") {
	if (typeof v === "number") {
		if (!Number.isFinite(v)) throw new Error(`${fieldName}: ${v} is not a finite number`);
		return v;
	}
	const m = v.trim().match(/^(-?\d+(?:\.\d+)?)\s*(pt|cm|mm|in)$/i);
	if (!m) throw new Error(`${fieldName} "${v}": expected a number (pt) or a string like "12pt" / "2.54cm" / "5mm" / "1in". Supported units: ${SUPPORTED_LENGTH_UNITS}.`);
	const n = parseFloat(m[1]);
	const unit = m[2].toLowerCase();
	switch (unit) {
		case "pt": return n;
		case "cm": return n * PT_PER_CM;
		case "mm": return n * PT_PER_MM;
		case "in": return n * PT_PER_IN;
	}
	throw new Error(`${fieldName} "${v}": unsupported unit "${unit}".`);
}
/** Parse an indent input into its OOXML-ready tagged unit.
*
*   number          → pt → twips (`w:firstLine` / `w:hanging`)
*   "Npt|Ncm|..."   → pt → twips
*   "Nchar"         → 1/100 char (`w:firstLineChars` / `w:hangingChars`)
*   null            → null (caller drops the attribute)
*/
function parseIndent(v, fieldName = "indent") {
	if (v === null) return null;
	if (typeof v === "number") {
		if (!Number.isFinite(v)) throw new Error(`${fieldName}: ${v} is not a finite number`);
		return {
			kind: "twip",
			value: Math.round(v * TWIPS_PER_PT)
		};
	}
	const trimmed = v.trim();
	const charMatch = trimmed.match(/^(-?\d+(?:\.\d+)?)\s*chars?$/i);
	if (charMatch) return {
		kind: "char",
		value: Math.round(parseFloat(charMatch[1]) * 100)
	};
	const pt = parseLengthPt(trimmed, fieldName);
	return {
		kind: "twip",
		value: Math.round(pt * TWIPS_PER_PT)
	};
}
/** Parse a lineSpacing input into mode + OOXML-ready value.
*
*   number          → multiplier (auto), value in 240ths-of-line
*   "Npt|Ncm|..."   → exact line height, value in twips
*   { atLeast: L }  → at-least line height, value in twips
*/
function parseLineSpacing(v, fieldName = "lineSpacing") {
	if (typeof v === "number") {
		if (!Number.isFinite(v) || v <= 0) throw new Error(`${fieldName}: ${v} is not a positive finite multiplier`);
		return {
			mode: "auto",
			value: Math.round(v * TWENTIETHS_PER_LINE)
		};
	}
	if (typeof v === "string") {
		const pt = parseLengthPt(v, fieldName);
		return {
			mode: "exact",
			value: Math.round(pt * TWIPS_PER_PT)
		};
	}
	if (v && typeof v === "object" && "atLeast" in v) {
		const pt = parseLengthPt(v.atLeast, `${fieldName}.atLeast`);
		return {
			mode: "atLeast",
			value: Math.round(pt * TWIPS_PER_PT)
		};
	}
	throw new Error(`${fieldName}: expected a number (multiplier, e.g. 1.5), a string (exact, e.g. "24pt"), or { atLeast: "<length>" }.`);
}
/** Expand CSS-shorthand padding to four edges in pt. Throws on negative
*  values — OOXML cell margins are `xsd:unsignedDecimalNumber`.
*  Error labels reflect CSS shorthand semantics: 2-value reports
*  `.vertical` / `.horizontal`, 3-value reports `.top` / `.horizontal` /
*  `.bottom`, 4-value reports `.top` / `.right` / `.bottom` / `.left`. */
function parsePadding(v, fieldName = "padding") {
	const parts = Array.isArray(v) ? v : [v];
	const at = (i, label) => parseLengthPt(parts[i], label ? `${fieldName}.${label}` : fieldName);
	let edges;
	if (parts.length === 1) {
		const all = at(0, "");
		edges = {
			top: all,
			right: all,
			bottom: all,
			left: all
		};
	} else if (parts.length === 2) {
		const vert = at(0, "vertical");
		const horiz = at(1, "horizontal");
		edges = {
			top: vert,
			right: horiz,
			bottom: vert,
			left: horiz
		};
	} else if (parts.length === 3) {
		const horiz = at(1, "horizontal");
		edges = {
			top: at(0, "top"),
			right: horiz,
			bottom: at(2, "bottom"),
			left: horiz
		};
	} else if (parts.length === 4) edges = {
		top: at(0, "top"),
		right: at(1, "right"),
		bottom: at(2, "bottom"),
		left: at(3, "left")
	};
	else throw new Error(`${fieldName}: tuple must have 1-4 entries, got ${parts.length}`);
	for (const side of [
		"top",
		"right",
		"bottom",
		"left"
	]) if (edges[side] < 0) throw new Error(`${fieldName}.${side}: ${edges[side]} is negative; OOXML cell margins must be non-negative`);
	return edges;
}
function toTwips(v, fieldName) {
	return Math.round(parseLengthPt(v, fieldName) * TWIPS_PER_PT);
}
function toHalfPt(v, fieldName) {
	return Math.round(parseLengthPt(v, fieldName) * HALF_PT_PER_PT);
}
function toEighthPt(v, fieldName) {
	return Math.round(parseLengthPt(v, fieldName) * EIGHTH_PT_PER_PT);
}
function toEmu(v, fieldName) {
	return Math.round(parseLengthPt(v, fieldName) * EMU_PER_PT);
}
/** Format twips back into a `"Npt"` string for round-tripping engine state
*  into config-shaped output (dry-run reports, fromParagraph extraction). */
function twipsToPtString(twips) {
	return `${twips / TWIPS_PER_PT}pt`;
}
/** Format twips into a `"N.NNcm"` display string for dry-run / report
*  contexts where users think in cm (margins, paper sizes). Two-decimal
*  precision matches Word's UI display granularity. */
function twipsToCmString(twips) {
	return `${(twips / (TWIPS_PER_PT * PT_PER_CM)).toFixed(2)}cm`;
}
/** Half-pt → pt (number form, since size fields take bare pt). */
function halfPtToPt(halfPt) {
	return halfPt / HALF_PT_PER_PT;
}
/** Char-units (×100, as Word stores `firstLineChars`) → "Nchar" string. */
function charUnitsToString(charUnits) {
	return `${charUnits / 100}char`;
}
/** Inverse of parseLineSpacing's value, given the mode it was parsed under.
*  Returns the config-shaped value that would round-trip back. */
function lineSpacingToConfig(parsed) {
	switch (parsed.mode) {
		case "auto": return parsed.value / TWENTIETHS_PER_LINE;
		case "exact": return twipsToPtString(parsed.value);
		case "atLeast": return { atLeast: twipsToPtString(parsed.value) };
	}
}

//#endregion
//#region lib/xml/ind-attr.ts
const w$18 = NS.w;
const FIRST_LINE_GROUP = [
	"firstLine",
	"firstLineChars",
	"hanging",
	"hangingChars"
];
/** Write a parsed indent value onto a <w:ind> element.
*  Emits even when `parsed.value === 0` — explicit zero overrides the
*  style cascade. Callers gate on `parsed !== null` themselves
*  (a `null` parse result means the source field was absent). */
function setIndentAttr(ind, slot, parsed) {
	const attr = parsed.kind === "char" ? `w:${slot}Chars` : `w:${slot}`;
	ind.setAttributeNS(w$18, attr, String(parsed.value));
}
/** Remove all attrs in the firstLine / hanging mutually-exclusive group.
*  In OOXML these four attrs are pairwise exclusive (same offset, different
*  units; opposing directions). Field-merge into an existing <w:ind> must
*  clear the group before writing the new value, otherwise stale attrs
*  from a prior mutation produce ambiguous Word behavior. */
function clearFirstLineHangingGroup(ind) {
	for (const a of FIRST_LINE_GROUP) ind.removeAttributeNS(w$18, a);
}

//#endregion
//#region lib/apply/style-mutation.ts
/**
* Insert a freshly-created `<w:pPr>` into a `<w:style>` at the schema-correct
* position. CT_Style requires `pPr` before `rPr` / `tblPr` / `trPr` / `tcPr` /
* `tblStylePr`. A naive `appendChild` would land it after an existing `rPr`,
* which Word's strict validator rejects with "file needs repair".
*/
function insertPPrIntoStyle(styleEl, pPr) {
	const w = NS.w;
	for (const name of [
		"rPr",
		"tblPr",
		"trPr",
		"tcPr",
		"tblStylePr"
	]) {
		const el = firstChildNS(styleEl, w, name);
		if (el) {
			styleEl.insertBefore(pPr, el);
			return;
		}
	}
	styleEl.appendChild(pPr);
}
function resolveStyleDef(def, paragraphs) {
	if (def.fromParagraph === void 0) return def.overrides ? {
		...def,
		...def.overrides
	} : def;
	const para = paragraphs.find((p) => p.index === def.fromParagraph);
	if (!para) {
		const indices = paragraphs.map((p) => p.index);
		const minIdx = indices[0] ?? 0;
		const maxIdx = indices[indices.length - 1] ?? 0;
		const closest = paragraphs.reduce((best, p) => Math.abs(p.index - def.fromParagraph) < Math.abs(best.index - def.fromParagraph) ? p : best, paragraphs[0]);
		throw new Error(`style "${def.id}": fromParagraph #${def.fromParagraph} not found.\n  Document has ${paragraphs.length} indexed paragraphs (range: #${minIdx}–#${maxIdx}).\n  Closest valid: #${closest.index} ("${closest.text.slice(0, 40)}${closest.text.length > 40 ? "…" : ""}")\n  Note: paragraphs inside data tables are not indexed and cannot be referenced.`);
	}
	const extracted = paragraphToStyleEntry(para);
	const { id: _id, name: _name, fromParagraph: _fp, basedOn: _bo, overrides: _ov, ...topLevel } = def;
	return {
		basedOn: "Normal",
		...extracted,
		...topLevel,
		...def.overrides ?? {},
		id: def.id,
		name: def.name,
		...def.basedOn !== void 0 ? { basedOn: def.basedOn } : {}
	};
}
/**
* Capture the cascade-resolved display fields for an existing styleId BEFORE
* this apply run mutates it. Returns null when the styleId is absent from
* source — the agent is installing it fresh, so there's no prior state to
* diff against. Shape mirrors `extractDisplayFields` (report.ts) so the two
* can be compared field-by-field for the dry-run Δ line.
*/
function extractPriorDisplayFields(resolver, styleId) {
	if (!resolver.getStyleDefinition(styleId)) return null;
	const { rPr, pPr } = resolver.resolveStyleChain(styleId);
	const out = {};
	if (rPr.fontEastAsia) out.fontCJK = rPr.fontEastAsia;
	const latin = rPr.fontAscii ?? rPr.fontHAnsi;
	if (latin) out.fontLatin = latin;
	if (rPr.size !== void 0) out.size = halfPtToPt(rPr.size);
	if (rPr.bold !== void 0) out.bold = rPr.bold;
	if (rPr.italic !== void 0) out.italic = rPr.italic;
	if (rPr.color && rPr.color !== "auto") out.color = rPr.color;
	if (rPr.vertAlign) out.vertAlign = rPr.vertAlign;
	if (pPr.alignment) out.alignment = pPr.alignment;
	if (pPr.spaceBefore !== void 0) out.spaceBefore = twipsToPtString(pPr.spaceBefore);
	if (pPr.spaceAfter !== void 0) out.spaceAfter = twipsToPtString(pPr.spaceAfter);
	if (pPr.lineSpacing !== void 0) out.lineSpacing = lineSpacingToConfig({
		mode: pPrLineRuleToMode(pPr.lineRule),
		value: pPr.lineSpacing
	});
	if (pPr.firstLineIndentChars !== void 0) out.firstLineIndent = charUnitsToString(pPr.firstLineIndentChars);
	else if (pPr.firstLineIndent !== void 0) out.firstLineIndent = twipsToPtString(pPr.firstLineIndent);
	if (pPr.outlineLevel !== void 0) out.outlineLevel = pPr.outlineLevel;
	return out;
}
/** Resolve the OOXML `lineRule` attribute string (or missing) to the
*  units.ts LineSpacingParsed mode. Defaults to "auto" matching Word's
*  ECMA-376 default when the attribute is absent. */
function pPrLineRuleToMode(lineRule) {
	if (lineRule === "exact" || lineRule === "atLeast") return lineRule;
	return "auto";
}
function paragraphToStyleEntry(p) {
	return computedStyleToEntry(p.rPr, p.pPr);
}
/** Convert a ComputedRunStyle + ComputedParaStyle pair into the field
* shape used by `StyleConfigEntry`. Shared between Mode A's `fromParagraph`
* extraction (cascade-merged values) and the dry-run vs-target-direct
* classifier (per-paragraph direct values). Both produce the same field
* names + units so they can be diffed against an agent-declared style. */
function computedStyleToEntry(r, pp) {
	const out = {};
	const latin = r.fontAscii ?? r.fontHAnsi;
	if (latin) out.fontLatin = latin;
	if (r.fontEastAsia && r.fontEastAsia !== latin) out.fontCJK = r.fontEastAsia;
	if (r.size !== void 0) out.size = halfPtToPt(r.size);
	if (r.bold) out.bold = true;
	if (r.italic) out.italic = true;
	if (r.color && r.color !== "auto") out.color = r.color;
	if (r.vertAlign) out.vertAlign = r.vertAlign;
	if (pp.alignment) out.alignment = pp.alignment;
	if (pp.spaceBefore !== void 0) out.spaceBefore = twipsToPtString(pp.spaceBefore);
	if (pp.spaceAfter !== void 0) out.spaceAfter = twipsToPtString(pp.spaceAfter);
	if (pp.lineSpacing !== void 0) out.lineSpacing = lineSpacingToConfig({
		mode: pPrLineRuleToMode(pp.lineRule),
		value: pp.lineSpacing
	});
	if (pp.firstLineIndentChars !== void 0) out.firstLineIndent = charUnitsToString(pp.firstLineIndentChars);
	else if (pp.firstLineIndent !== void 0) out.firstLineIndent = twipsToPtString(pp.firstLineIndent);
	if (pp.hangingIndentChars !== void 0) out.hangingIndent = charUnitsToString(pp.hangingIndentChars);
	else if (pp.hangingIndent !== void 0) out.hangingIndent = twipsToPtString(pp.hangingIndent);
	if (pp.indentLeft !== void 0) out.indentLeft = twipsToPtString(pp.indentLeft);
	if (pp.indentRight !== void 0) out.indentRight = twipsToPtString(pp.indentRight);
	if (pp.outlineLevel !== void 0) out.outlineLevel = pp.outlineLevel;
	return out;
}
/**
* Resolve a style cross-reference (basedOn / link / next val) to an existing
* `w:styleId`. ECMA-376 says these refs must target a styleId; agents and
* defaults often write the conventional name ("Normal", "heading 1") instead.
* Falls back to case-insensitive `w:name` match when the literal val isn't a
* styleId. Returns null when neither matches (caller drops the ref rather
* than leaving it dangling).
*
* `selfId` is the styleId of the style currently being written; self-refs
* (`basedOn === id`) collapse the cascade and are silently dropped.
*/
function resolveStyleRef(stylesDoc, val, selfId) {
	const w = NS.w;
	const styles = getChildrenNS(stylesDoc.documentElement, w, "style");
	for (const s of styles) if (wAttr(s, "styleId") === val) return val === selfId ? null : val;
	const target = val.toLowerCase();
	for (const s of styles) {
		const id = wAttr(s, "styleId");
		if (id === selfId) continue;
		const nameEl = firstChildNS(s, w, "name");
		if (nameEl && (wAttr(nameEl, "val") || "").toLowerCase() === target) return id ?? null;
	}
	return null;
}
/** Return (or create) a child element with the given local name under
* `parent`. When creating, the new element is NOT inserted — caller must
* place it via `insertChildInOrder` or equivalent so schema order is
* respected. When the element already exists it is returned as-is for
* attribute-level mutation. */
function getOrCreateNS(parent, doc, ns, localName) {
	const existing = firstChildNS(parent, ns, localName);
	if (existing) return existing;
	return doc.createElementNS(ns, `w:${localName}`);
}
/**
* Upsert an OOXML toggle element (bold, italic, etc.) in `parent`.
*
* - `on === true`:  ensure element exists with no `w:val` attribute (presence = on).
* - `on === false`: ensure element exists with `w:val="0"` (explicit off — overrides
*   any `basedOn` ancestor that declares the toggle; removing the element would mean
*   "inherit", not "force off").
*/
function upsertToggle(parent, doc, ns, localName, on, order) {
	let el = firstChildNS(parent, ns, localName);
	if (!el) {
		el = doc.createElementNS(ns, `w:${localName}`);
		insertChildInOrder(parent, el, order);
	}
	if (on) el.removeAttributeNS(ns, "val");
	else el.setAttributeNS(ns, "w:val", "0");
}
function upsertStyle(stylesDoc, def) {
	const w = NS.w;
	const root = stylesDoc.documentElement;
	const existing = getChildrenNS(root, w, "style").find((s) => wAttr(s, "styleId") === def.id);
	let target;
	let result;
	if (existing) {
		target = existing;
		result = "updated";
	} else {
		target = stylesDoc.createElementNS(w, "w:style");
		target.setAttributeNS(w, "w:type", "paragraph");
		target.setAttributeNS(w, "w:styleId", def.id);
		root.appendChild(target);
		result = "created";
	}
	let nameEl = firstChildNS(target, w, "name");
	if (result === "updated" && def.name === void 0) {} else {
		const nameVal = def.name ?? def.id;
		if (nameEl) nameEl.setAttributeNS(w, "w:val", nameVal);
		else {
			nameEl = stylesDoc.createElementNS(w, "w:name");
			nameEl.setAttributeNS(w, "w:val", nameVal);
			target.insertBefore(nameEl, target.firstChild);
		}
	}
	if (def.basedOn) {
		const resolved = resolveStyleRef(stylesDoc, def.basedOn, def.id);
		let bo = firstChildNS(target, w, "basedOn");
		if (resolved) if (bo) bo.setAttributeNS(w, "w:val", resolved);
		else {
			bo = stylesDoc.createElementNS(w, "w:basedOn");
			bo.setAttributeNS(w, "w:val", resolved);
			const afterName = firstChildNS(target, w, "name")?.nextSibling ?? null;
			if (afterName) target.insertBefore(bo, afterName);
			else target.appendChild(bo);
		}
		else if (bo) target.removeChild(bo);
	}
	const hasPPrChange = def.outlineLevel !== void 0 || def.alignment !== void 0 || def.spaceBefore !== void 0 || def.spaceAfter !== void 0 || def.lineSpacing !== void 0 || def.firstLineIndent != null || def.hangingIndent != null || def.indentLeft != null || def.indentRight != null;
	let pPr = firstChildNS(target, w, "pPr");
	if (hasPPrChange) {
		if (!pPr) {
			pPr = stylesDoc.createElementNS(w, "w:pPr");
			insertPPrIntoStyle(target, pPr);
		}
		if (def.outlineLevel !== void 0) {
			const ol = getOrCreateNS(pPr, stylesDoc, w, "outlineLvl");
			ol.setAttributeNS(w, "w:val", String(def.outlineLevel));
			if (!ol.parentNode) insertChildInOrder(pPr, ol, PPR_CHILD_ORDER);
		}
		if (def.alignment !== void 0) {
			const jc = getOrCreateNS(pPr, stylesDoc, w, "jc");
			jc.setAttributeNS(w, "w:val", def.alignment);
			if (!jc.parentNode) insertChildInOrder(pPr, jc, PPR_CHILD_ORDER);
		}
		if (def.spaceBefore !== void 0 || def.spaceAfter !== void 0 || def.lineSpacing !== void 0) {
			const spacing = getOrCreateNS(pPr, stylesDoc, w, "spacing");
			if (def.spaceBefore !== void 0) spacing.setAttributeNS(w, "w:before", String(toTwips(def.spaceBefore, "spaceBefore")));
			if (def.spaceAfter !== void 0) spacing.setAttributeNS(w, "w:after", String(toTwips(def.spaceAfter, "spaceAfter")));
			if (def.lineSpacing !== void 0) {
				const ls = parseLineSpacing(def.lineSpacing, "lineSpacing");
				spacing.setAttributeNS(w, "w:line", String(ls.value));
				spacing.setAttributeNS(w, "w:lineRule", ls.mode);
			}
			if (!spacing.parentNode) insertChildInOrder(pPr, spacing, PPR_CHILD_ORDER);
		}
		if (def.firstLineIndent != null || def.hangingIndent != null || def.indentLeft != null || def.indentRight != null) {
			const ind = getOrCreateNS(pPr, stylesDoc, w, "ind");
			if (def.firstLineIndent != null || def.hangingIndent != null) clearFirstLineHangingGroup(ind);
			if (def.firstLineIndent != null) {
				const r = parseIndent(def.firstLineIndent);
				if (r) setIndentAttr(ind, "firstLine", r);
			}
			if (def.hangingIndent != null) {
				const r = parseIndent(def.hangingIndent);
				if (r) setIndentAttr(ind, "hanging", r);
			}
			if (def.indentLeft != null) {
				const r = parseIndent(def.indentLeft);
				if (r) setIndentAttr(ind, "left", r);
			}
			if (def.indentRight != null) {
				const r = parseIndent(def.indentRight);
				if (r) setIndentAttr(ind, "right", r);
			}
			if (!ind.parentNode) insertChildInOrder(pPr, ind, PPR_CHILD_ORDER);
		}
	}
	const hasRPrChange = def.fontLatin !== void 0 || def.fontCJK !== void 0 || def.size !== void 0 || def.bold !== void 0 || def.italic !== void 0 || def.color !== void 0 || def.vertAlign !== void 0;
	let rPr = firstChildNS(target, w, "rPr");
	if (hasRPrChange) {
		if (!rPr) {
			rPr = stylesDoc.createElementNS(w, "w:rPr");
			target.appendChild(rPr);
		}
		if (def.fontLatin !== void 0 || def.fontCJK !== void 0) {
			const rFonts = getOrCreateNS(rPr, stylesDoc, w, "rFonts");
			if (def.fontLatin !== void 0) {
				rFonts.setAttributeNS(w, "w:ascii", def.fontLatin);
				rFonts.setAttributeNS(w, "w:hAnsi", def.fontLatin);
			}
			if (def.fontCJK !== void 0) rFonts.setAttributeNS(w, "w:eastAsia", def.fontCJK);
			if (!rFonts.parentNode) insertChildInOrder(rPr, rFonts, RPR_CHILD_ORDER);
		}
		if (def.size !== void 0) {
			const halfPt = toHalfPt(def.size, "size");
			const sz = getOrCreateNS(rPr, stylesDoc, w, "sz");
			sz.setAttributeNS(w, "w:val", String(halfPt));
			if (!sz.parentNode) insertChildInOrder(rPr, sz, RPR_CHILD_ORDER);
			const szCs = getOrCreateNS(rPr, stylesDoc, w, "szCs");
			szCs.setAttributeNS(w, "w:val", String(halfPt));
			if (!szCs.parentNode) insertChildInOrder(rPr, szCs, RPR_CHILD_ORDER);
		}
		if (def.bold !== void 0) {
			upsertToggle(rPr, stylesDoc, w, "b", def.bold, RPR_CHILD_ORDER);
			upsertToggle(rPr, stylesDoc, w, "bCs", def.bold, RPR_CHILD_ORDER);
		}
		if (def.italic !== void 0) {
			upsertToggle(rPr, stylesDoc, w, "i", def.italic, RPR_CHILD_ORDER);
			upsertToggle(rPr, stylesDoc, w, "iCs", def.italic, RPR_CHILD_ORDER);
		}
		if (def.color !== void 0) {
			const color = getOrCreateNS(rPr, stylesDoc, w, "color");
			color.setAttributeNS(w, "w:val", def.color);
			if (!color.parentNode) insertChildInOrder(rPr, color, RPR_CHILD_ORDER);
		}
		if (def.vertAlign !== void 0) {
			const va = getOrCreateNS(rPr, stylesDoc, w, "vertAlign");
			va.setAttributeNS(w, "w:val", def.vertAlign);
			if (!va.parentNode) insertChildInOrder(rPr, va, RPR_CHILD_ORDER);
		}
	}
	return result;
}
/**
* Move the styles named in `orderedIds` to the top of styles.xml's <w:style>
* entry list, preserving the given order. Styles not in the list keep their
* relative positions. Non-style children of the root (<w:docDefaults>,
* <w:latentStyles>) stay above all <w:style> entries — they're not touched.
*
* Each id appears at most once in `orderedIds`; duplicates are deduped (a
* style can be both template-imported and re-declared in config.styles[]).
*/
function reorderAgentTouchedStylesFirst(stylesDoc, orderedIds) {
	const w = NS.w;
	const root = stylesDoc.documentElement;
	if (!root) return;
	const seen = /* @__PURE__ */ new Set();
	const dedupedIds = orderedIds.filter((id) => {
		if (seen.has(id)) return false;
		seen.add(id);
		return true;
	});
	const targetSet = new Set(dedupedIds);
	const touchedById = /* @__PURE__ */ new Map();
	let firstUntouched = null;
	for (const el of getChildrenNS(root, w, "style")) {
		const id = wAttr(el, "styleId");
		if (id && targetSet.has(id)) touchedById.set(id, el);
		else if (!firstUntouched) firstUntouched = el;
	}
	for (const el of touchedById.values()) if (el.parentNode === root) root.removeChild(el);
	for (const id of dedupedIds) {
		const el = touchedById.get(id);
		if (!el) continue;
		if (firstUntouched) root.insertBefore(el, firstUntouched);
		else root.appendChild(el);
	}
}

//#endregion
//#region lib/shared/vs-direct.ts
/**
* "vs target direct" classifier for the dry-run Style Resolution report.
*
* The existing Δ-line compares an agent's declared style against the source
* styles.xml's prior cascade — but the sparse-by-design invariant actually
* lives at the **direct format** layer (the pPr / rPr literally on each
* paragraph). A declaration of `bold: true` may be redundant because every
* target paragraph already has `<w:b/>` direct, or it may override a direct
* `<w:b w:val="false"/>` on those same paragraphs. The styles-cascade Δ-line
* can't see either case.
*
* For each field the agent declares on a style, this classifier walks the
* target paragraphs (post-rules-pass: any paragraph whose final pStyle is
* the agent's styleId) and tallies:
*   - override  — the paragraph carries a direct value that differs from
*                 the declaration; the new style replaces it (after the
*                 engine's selective-strip removes the direct conflict).
*   - redundant — paragraph's direct value matches the declaration; the
*                 declaration adds nothing new (the strip removes the
*                 direct then the style cascade reinstates the same value).
*   - new       — paragraph has no direct value for this field; the
*                 declaration is a fresh addition via the style cascade.
*
* Mode A same-source silencing: when a style was built from a representative
* paragraph (`fromParagraph: N`), the fields extracted necessarily match N's
* direct values, and every paragraph sharing N's fingerprint will trivially
* match too. Those redundant counts are noise — pass N's fingerprint here
* and the classifier skips those paragraphs' redundant tallies (override
* and new still count, since cross-fingerprint paragraphs are informative).
*/
const STRIPPABLE_FIELDS = new Set([
	"alignment",
	"lineSpacing",
	"spaceBefore",
	"spaceAfter",
	"firstLineIndent",
	"hangingIndent",
	"indentLeft",
	"indentRight",
	"outlineLevel",
	"fontLatin",
	"fontCJK",
	"size",
	"bold",
	"italic",
	"color",
	"vertAlign"
]);
const FORMAT_FIELDS = [
	"fontLatin",
	"fontCJK",
	"size",
	"bold",
	"italic",
	"color",
	"vertAlign",
	"alignment",
	"lineSpacing",
	"spaceBefore",
	"spaceAfter",
	"firstLineIndent",
	"hangingIndent",
	"indentLeft",
	"indentRight",
	"outlineLevel"
];
function analyzeVsDirect(declared, targets, fromParagraphFingerprint) {
	const fields = [];
	for (const field of FORMAT_FIELDS) {
		const declaredRaw = declared[field];
		if (declaredRaw === void 0) continue;
		const declaredCanonical = canonicalize(field, declaredRaw);
		let override = 0;
		let redundant = 0;
		let fresh = 0;
		let overrideFrom;
		for (const p of targets) {
			const directRaw = computedStyleToEntry(p.directRPr, p.directPPr)[field];
			if (directRaw === void 0) {
				fresh++;
				continue;
			}
			const directCanonical = canonicalize(field, directRaw);
			if (sameValue(directCanonical, declaredCanonical)) {
				if (fromParagraphFingerprint && p.fingerprint === fromParagraphFingerprint) continue;
				redundant++;
			} else {
				override++;
				if (overrideFrom === void 0) overrideFrom = directCanonical;
			}
		}
		if (override === 0 && redundant === 0 && fresh === 0) continue;
		fields.push({
			field,
			override,
			redundant,
			fresh,
			overrideFrom,
			declared: declaredCanonical,
			willStrip: STRIPPABLE_FIELDS.has(field)
		});
	}
	return {
		targetCount: targets.length,
		fields
	};
}
function canonicalize(field, val) {
	if (field === "lineSpacing") {
		const ls = parseLineSpacing(val, "lineSpacing");
		return `${ls.mode}:${ls.value}`;
	}
	return val;
}
/** Shallow value equality for restyle / report comparisons. Numbers
* compare with a 1e-6 tolerance to absorb twip round-trip noise; other
* types stringify and compare. Exported so `report.ts` and other
* comparators stay aligned on what "same" means. */
function sameValue(a, b) {
	if (typeof a === "number" && typeof b === "number") return Math.abs(a - b) < 1e-6;
	return String(a) === String(b);
}

//#endregion
//#region lib/apply/numbering-mutation.ts
/**
* Decide the marker-suffix character and the lvlText that should actually be
* written to OOXML. When the user sets `suff` explicitly, that wins. Otherwise
* we read the user's hand-written `lvlText` for trailing ASCII spaces — 0 → no
* gap, 1 → one space, 2+ → tab — which captures their format intent without
* us having to detect "is this CJK or Western punctuation". Trailing spaces
* are stripped from the emitted lvlText so the gap is owned solely by suff
* (otherwise a `"%1. "` lvlText with `suff=space` would render double-spaced).
*/
function resolveSuff(lvlText, explicit) {
	const trailing = lvlText.match(/ *$/)[0].length;
	return {
		suff: explicit ?? (trailing === 0 ? "nothing" : trailing === 1 ? "space" : "tab"),
		effectiveLvlText: lvlText.replace(/ +$/, "")
	};
}
function injectNumbering(numberingDoc, config, opts = {}) {
	const w = NS.w;
	const root = numberingDoc.documentElement;
	const existingAbsIds = getChildrenNS(root, w, "abstractNum").map((e) => parseInt(wAttr(e, "abstractNumId") || "0", 10));
	const existingNumIds = getChildrenNS(root, w, "num").map((e) => parseInt(wAttr(e, "numId") || "0", 10));
	const nextAbs = (existingAbsIds.length ? Math.max(...existingAbsIds) : -1) + 1;
	let nextNum;
	if (opts.requestedNumId !== void 0) nextNum = opts.requestedNumId;
	else {
		const claimed = opts.claimedNumIds ?? /* @__PURE__ */ new Set();
		let candidate = (existingNumIds.length ? Math.max(...existingNumIds) : 0) + 1;
		while (claimed.has(candidate)) candidate++;
		nextNum = candidate;
	}
	const abs = numberingDoc.createElementNS(w, "w:abstractNum");
	abs.setAttributeNS(w, "w:abstractNumId", String(nextAbs));
	const multiLevelType = numberingDoc.createElementNS(w, "w:multiLevelType");
	multiLevelType.setAttributeNS(w, "w:val", config.levels.length > 1 ? "multilevel" : "singleLevel");
	abs.appendChild(multiLevelType);
	for (const lvl of config.levels) {
		const lvlEl = numberingDoc.createElementNS(w, "w:lvl");
		lvlEl.setAttributeNS(w, "w:ilvl", String(lvl.level));
		const start = numberingDoc.createElementNS(w, "w:start");
		start.setAttributeNS(w, "w:val", String(lvl.start ?? 1));
		lvlEl.appendChild(start);
		const numFmtEl = numberingDoc.createElementNS(w, "w:numFmt");
		numFmtEl.setAttributeNS(w, "w:val", lvl.numFmt);
		lvlEl.appendChild(numFmtEl);
		const pStyle = numberingDoc.createElementNS(w, "w:pStyle");
		pStyle.setAttributeNS(w, "w:val", lvl.styleId);
		lvlEl.appendChild(pStyle);
		if (lvl.isLgl) {
			const isLgl = numberingDoc.createElementNS(w, "w:isLgl");
			lvlEl.appendChild(isLgl);
		}
		const { suff, effectiveLvlText } = resolveSuff(lvl.lvlText, lvl.suff);
		const suffEl = numberingDoc.createElementNS(w, "w:suff");
		suffEl.setAttributeNS(w, "w:val", suff);
		lvlEl.appendChild(suffEl);
		const lvlTextEl = numberingDoc.createElementNS(w, "w:lvlText");
		lvlTextEl.setAttributeNS(w, "w:val", effectiveLvlText);
		lvlEl.appendChild(lvlTextEl);
		const lvlJc = numberingDoc.createElementNS(w, "w:lvlJc");
		lvlJc.setAttributeNS(w, "w:val", "left");
		lvlEl.appendChild(lvlJc);
		if (lvl.numRPr) {
			const rPr = buildLvlRPr(numberingDoc, lvl.numRPr);
			if (rPr) lvlEl.appendChild(rPr);
		}
		abs.appendChild(lvlEl);
	}
	const firstNum = getChildrenNS(root, w, "num")[0];
	if (firstNum) root.insertBefore(abs, firstNum);
	else root.appendChild(abs);
	const num = numberingDoc.createElementNS(w, "w:num");
	num.setAttributeNS(w, "w:numId", String(nextNum));
	const absRef = numberingDoc.createElementNS(w, "w:abstractNumId");
	absRef.setAttributeNS(w, "w:val", String(nextAbs));
	num.appendChild(absRef);
	root.appendChild(num);
	return {
		numId: String(nextNum),
		abstractNumId: String(nextAbs)
	};
}
function buildLvlRPr(doc, spec) {
	const w = NS.w;
	const rPr = doc.createElementNS(w, "w:rPr");
	const add = (el) => insertChildInOrder(rPr, el, RPR_CHILD_ORDER);
	if (spec.fontLatin || spec.fontCJK) {
		const rFonts = doc.createElementNS(w, "w:rFonts");
		if (spec.fontLatin) {
			rFonts.setAttributeNS(w, "w:ascii", spec.fontLatin);
			rFonts.setAttributeNS(w, "w:hAnsi", spec.fontLatin);
		}
		if (spec.fontCJK) rFonts.setAttributeNS(w, "w:eastAsia", spec.fontCJK);
		add(rFonts);
	}
	if (spec.size !== void 0) {
		const sz = doc.createElementNS(w, "w:sz");
		sz.setAttributeNS(w, "w:val", String(toHalfPt(spec.size, "numbering.rPr.size")));
		add(sz);
	}
	if (spec.bold !== void 0) {
		const b = doc.createElementNS(w, "w:b");
		if (!spec.bold) b.setAttributeNS(w, "w:val", "0");
		add(b);
	}
	if (spec.italic !== void 0) {
		const i = doc.createElementNS(w, "w:i");
		if (!spec.italic) i.setAttributeNS(w, "w:val", "0");
		add(i);
	}
	if (spec.color) {
		const color = doc.createElementNS(w, "w:color");
		color.setAttributeNS(w, "w:val", spec.color);
		add(color);
	}
	return getChildren(rPr).length > 0 ? rPr : null;
}
/**
* Mint a fresh `<w:num>` pointing to `abstractNumId`, with
* `<w:lvlOverride><w:startOverride val="1"/></w:lvlOverride>` so Word resets
* the counter to 1. Returns the new numId. Used by the per-instance restart
* pass for single-level (list-shaped) schemes — every "list instance" in the
* document gets its own counter via this fork, so a `1.` `2.` `3.` list in
* Chapter 1 doesn't continue as `4.` `5.` `6.` in Chapter 2.
*/
function forkNumWithStartOverride(numberingDoc, abstractNumId, level) {
	const w = NS.w;
	const root = numberingDoc.documentElement;
	const existingNumIds = getChildrenNS(root, w, "num").map((e) => parseInt(wAttr(e, "numId") || "0", 10));
	const nextNum = (existingNumIds.length ? Math.max(...existingNumIds) : 0) + 1;
	const num = numberingDoc.createElementNS(w, "w:num");
	num.setAttributeNS(w, "w:numId", String(nextNum));
	const absRef = numberingDoc.createElementNS(w, "w:abstractNumId");
	absRef.setAttributeNS(w, "w:val", abstractNumId);
	num.appendChild(absRef);
	const lvlOverride = numberingDoc.createElementNS(w, "w:lvlOverride");
	lvlOverride.setAttributeNS(w, "w:ilvl", String(level));
	const startOverride = numberingDoc.createElementNS(w, "w:startOverride");
	startOverride.setAttributeNS(w, "w:val", "1");
	lvlOverride.appendChild(startOverride);
	num.appendChild(lvlOverride);
	root.appendChild(num);
	return String(nextNum);
}
/**
* Set (or replace) paragraph-level `<w:numPr>` on a `<w:p>` element. Inserted
* at the correct position per CT_PPr schema (after `<w:pStyle>`, before the
* formatting children). Used by the per-instance restart pass to override the
* style-level numId binding with a per-run forked numId.
*/
function setParagraphNumPr(pEl, numId, level) {
	const w = NS.w;
	const ownerDoc = pEl.ownerDocument;
	let pPr = firstChildNS(pEl, w, "pPr");
	if (!pPr) {
		pPr = ownerDoc.createElementNS(w, "w:pPr");
		pEl.insertBefore(pPr, pEl.firstChild);
	}
	const existing = firstChildNS(pPr, w, "numPr");
	if (existing) pPr.removeChild(existing);
	const numPr = ownerDoc.createElementNS(w, "w:numPr");
	const ilvl = ownerDoc.createElementNS(w, "w:ilvl");
	ilvl.setAttributeNS(w, "w:val", String(level));
	numPr.appendChild(ilvl);
	const numIdEl = ownerDoc.createElementNS(w, "w:numId");
	numIdEl.setAttributeNS(w, "w:val", numId);
	numPr.appendChild(numIdEl);
	const pStyle = firstChildNS(pPr, w, "pStyle");
	if (pStyle && pStyle.nextSibling) pPr.insertBefore(numPr, pStyle.nextSibling);
	else if (pStyle) pPr.appendChild(numPr);
	else pPr.insertBefore(numPr, pPr.firstChild);
}
function attachNumberingToStyle(stylesDoc, styleId, numId, level) {
	const w = NS.w;
	const styleEl = getChildrenNS(stylesDoc.documentElement, w, "style").find((s) => wAttr(s, "styleId") === styleId);
	if (!styleEl) return;
	let pPr = firstChildNS(styleEl, w, "pPr");
	if (!pPr) {
		pPr = stylesDoc.createElementNS(w, "w:pPr");
		insertPPrIntoStyle(styleEl, pPr);
	}
	const existing = firstChildNS(pPr, w, "numPr");
	if (existing) pPr.removeChild(existing);
	const numPr = stylesDoc.createElementNS(w, "w:numPr");
	const ilvl = stylesDoc.createElementNS(w, "w:ilvl");
	ilvl.setAttributeNS(w, "w:val", String(level));
	numPr.appendChild(ilvl);
	const numIdEl = stylesDoc.createElementNS(w, "w:numId");
	numIdEl.setAttributeNS(w, "w:val", numId);
	numPr.appendChild(numIdEl);
	pPr.insertBefore(numPr, pPr.firstChild);
}

//#endregion
//#region lib/apply/list-restart.ts
/**
* List-restart pass for single-level numbering schemes that opt in via a
* scheme-level `restart` value.
*
* Word's `<w:num>` element IS the counter. All paragraphs sharing a numId
* share one counter — so a style-level binding makes the counter run
* continuously across the whole document. That's the right behavior for
* caption / reference / global-counter shapes and is the default
* (`restart: "continuous"`).
*
* The engine supports three opt-in restart modes:
*
*   "perInstance"        — restart at each contiguous run of paragraphs
*                          bound to the scheme's styleId. A non-target
*                          paragraph breaks the run and triggers a restart.
*                          Classic use: 1./2./3. lists in Chapter 1 that
*                          should start fresh in Chapter 2.
*
*   "byHeading"          — restart whenever a heading-styled paragraph
*                          (any style with <w:outlineLvl>, or a paragraph
*                          with direct <w:outlineLvl> in its own pPr) is
*                          encountered. Each heading increments a per-target
*                          epoch counter; when a list paragraph's last-seen
*                          epoch differs from the current epoch, a new fork
*                          is started. Use for "each chapter gets its own
*                          1, 2, 3, …" without caring which heading level
*                          triggered it.
*
*   { atStyleChange: S } — restart whenever a paragraph bound to styleId S
*                          precedes the current list paragraph. Useful when
*                          chapter boundaries are marked by a custom style
*                          that doesn't carry outlineLvl (e.g. "ProposalH2").
*
* All three modes share the same OOXML mechanism: fork a fresh `<w:num>`
* pointing to the same abstractNumId but carrying
* `<w:lvlOverride><w:startOverride val="1"/></w:lvlOverride>`, and write
* paragraph-level `<w:numPr>` on each affected paragraph so it overrides
* the style-level binding.
*
* Only single-level schemes with a non-"continuous" restart are forked;
* multi-level schemes and continuous single-level schemes are skipped.
*/
/**
* Build the set of styleIds that declare <w:outlineLvl> in their <w:pPr>
* in styles.xml. Used by applyListRestartPass for byHeading mode to detect
* heading paragraphs via the style cascade, not just via direct pPr overrides.
*/
function buildHeadingStyleIdSet(stylesDoc) {
	const w = NS.w;
	const result = /* @__PURE__ */ new Set();
	const root = stylesDoc.documentElement;
	if (!root) return result;
	for (const styleEl of getChildrenNS(root, w, "style")) {
		const id = wAttr(styleEl, "styleId");
		if (!id) continue;
		const pPr = firstChildNS(styleEl, w, "pPr");
		if (!pPr) continue;
		if (firstChildNS(pPr, w, "outlineLvl")) result.add(id);
	}
	return result;
}
function applyListRestartPass(documentDoc, numberingDoc, installedSchemes, headingStyleIds) {
	const targets = [];
	for (const scheme of installedSchemes) {
		if (scheme.levels.length !== 1) continue;
		const lvl = scheme.levels[0];
		if (lvl.restart === "continuous" || !lvl.restart) continue;
		targets.push({
			styleId: lvl.styleId,
			abstractNumId: scheme.abstractNumId,
			level: lvl.level,
			mode: lvl.restart,
			baseNumId: scheme.numId
		});
	}
	if (targets.length === 0) return;
	const styleIdToTarget = new Map(targets.map((t) => [t.styleId, t]));
	const runs = /* @__PURE__ */ new Map();
	const currentRunByStyle = /* @__PURE__ */ new Map();
	for (const t of targets) {
		runs.set(t.styleId, []);
		currentRunByStyle.set(t.styleId, []);
	}
	const flush = (styleId) => {
		const cur = currentRunByStyle.get(styleId);
		if (cur.length > 0) {
			runs.get(styleId).push(cur);
			currentRunByStyle.set(styleId, []);
		}
	};
	const w = NS.w;
	const headingEpoch = { value: 0 };
	const lastSeenEpochByStyle = /* @__PURE__ */ new Map();
	const atStylePending = /* @__PURE__ */ new Map();
	for (const t of targets) if (t.mode === "byHeading") lastSeenEpochByStyle.set(t.styleId, 0);
	else if (typeof t.mode === "object") atStylePending.set(t.styleId, false);
	const body = firstChildNS(documentDoc.documentElement, w, "body");
	if (!body) return;
	for (const child of walkBodyParagraphs(body)) {
		const pPr = firstChildNS(child, w, "pPr");
		const pStyleEl = pPr ? firstChildNS(pPr, w, "pStyle") : null;
		const paragraphStyleId = pStyleEl ? wAttr(pStyleEl, "val") : "";
		let isHeading = false;
		if (paragraphStyleId && headingStyleIds?.has(paragraphStyleId)) isHeading = true;
		else if (pPr && firstChildNS(pPr, w, "outlineLvl")) isHeading = true;
		const target = paragraphStyleId ? styleIdToTarget.get(paragraphStyleId) : void 0;
		if (target) {
			const { mode } = target;
			if (mode === "perInstance") {
				currentRunByStyle.get(target.styleId).push(child);
				for (const t of targets) if (t.styleId !== target.styleId && t.mode === "perInstance") flush(t.styleId);
			} else if (mode === "byHeading") {
				if (lastSeenEpochByStyle.get(target.styleId) !== headingEpoch.value) {
					flush(target.styleId);
					lastSeenEpochByStyle.set(target.styleId, headingEpoch.value);
				}
				const existingNumPr = pPr ? firstChildNS(pPr, w, "numPr") : null;
				if (existingNumPr) {
					const numIdEl = firstChildNS(existingNumPr, w, "numId");
					const existingNumId = numIdEl ? wAttr(numIdEl, "val") : null;
					if (existingNumId !== null && existingNumId !== target.baseNumId) {
						flush(target.styleId);
						continue;
					}
				}
				currentRunByStyle.get(target.styleId).push(child);
			} else {
				if (atStylePending.get(target.styleId)) {
					flush(target.styleId);
					atStylePending.set(target.styleId, false);
				}
				currentRunByStyle.get(target.styleId).push(child);
			}
		} else {
			if (isHeading) headingEpoch.value++;
			for (const t of targets) if (t.mode === "perInstance") flush(t.styleId);
			else if (typeof t.mode === "object") {
				if (paragraphStyleId === t.mode.atStyleChange) atStylePending.set(t.styleId, true);
			}
		}
	}
	for (const t of targets) flush(t.styleId);
	for (const t of targets) for (const run of runs.get(t.styleId)) {
		const newNumId = forkNumWithStartOverride(numberingDoc, t.abstractNumId, t.level);
		for (const p of run) setParagraphNumPr(p, newNumId, t.level);
	}
}
/**
* Fork a fresh `<w:num>` for a paragraph-level explicit restart
* (`numbering.restart: true` in a ParagraphBlock).
*
* Looks up the abstractNumId from the paragraph's current numId in
* numberingDoc, forks it with `<w:startOverride val="1"/>`, and rewrites the
* paragraph's `<w:numPr>` to use the new numId. Restart applies only to this
* paragraph; subsequent paragraphs continue with their own declared numbering.
*
* Throws when numId doesn't resolve in numberingDoc — schema-valid input but
* doctree-inconsistent (caller passed a numId not installed in the doc).
*/
function applyParagraphLevelRestart(pEl, numberingDoc, numId, level) {
	const w = NS.w;
	const root = numberingDoc.documentElement;
	let abstractNumId = null;
	for (const num of getChildrenNS(root, w, "num")) if (wAttr(num, "numId") === numId) {
		const absRef = firstChildNS(num, w, "abstractNumId");
		abstractNumId = absRef ? wAttr(absRef, "val") : null;
		break;
	}
	if (!abstractNumId) throw new Error(`numbering.restart: numId "${numId}" not found in numbering.xml. Fix: use a numId that exists in the document's numbering.xml, or install a numbering scheme via config.numbering first.`);
	setParagraphNumPr(pEl, forkNumWithStartOverride(numberingDoc, abstractNumId, level), level);
}

//#endregion
//#region lib/parse/manual-numbering-detect.ts
const COMMON_TYPED_PREFIX_PATTERNS = [
	/^\s*\d+(\.\d+)+\s/,
	/^\s*\d+\.\s/,
	/^\s*\d+、/,
	/^\s*第[一二三四五六七八九十百千零〇0-9]+[章节篇部条款]/,
	/^\s*[一二三四五六七八九十百千零〇]+、/,
	/^\s*[（(][一二三四五六七八九十百千零〇0-9]+[)）]/,
	/^\s*[①-⑳⒈-⒛]/,
	/^\s*Chapter\s+\d+/i
];
function detectManualNumbering(edits, numberedStyleIds) {
	const out = /* @__PURE__ */ new Map();
	if (!edits) return out;
	for (const edit of edits) {
		const blocks = blocksOf(edit);
		if (!blocks) continue;
		for (const block of blocks) {
			if (block.type !== "paragraph") continue;
			if (!block.styleId) continue;
			const text = richTextToPlain(block.text);
			if (!text) continue;
			if (!COMMON_TYPED_PREFIX_PATTERNS.some((rx) => rx.test(text))) continue;
			const kind = numberedStyleIds.has(block.styleId) ? "bound" : "unbound";
			let hit = out.get(block.styleId);
			if (!hit) {
				hit = {
					count: 0,
					samples: [],
					kind
				};
				out.set(block.styleId, hit);
			}
			hit.count += 1;
			if (hit.samples.length < 3) {
				const trimmed = text.trim().slice(0, 40);
				hit.samples.push(trimmed + (text.trim().length > 40 ? "…" : ""));
			}
		}
	}
	return out;
}
function blocksOf(edit) {
	if (edit.op === "replace") return edit.with;
	if (edit.op === "insert-before" || edit.op === "insert-after") return edit.content;
	return null;
}
function richTextToPlain(rich) {
	if (typeof rich === "string") return rich;
	return rich.map((r) => "text" in r ? r.text : "").join("");
}

//#endregion
//#region lib/shared/report.ts
/**
* Pick the user-facing typographic fields from a resolved style for
* side-by-side display. Excludes mechanical fields (id, name, basedOn,
* fromParagraph, overrides) the user wouldn't recognize.
*/
function extractDisplayFields(def) {
	const out = {};
	for (const k of [
		"fontCJK",
		"fontLatin",
		"size",
		"bold",
		"italic",
		"color",
		"vertAlign",
		"alignment",
		"lineSpacing",
		"spaceBefore",
		"spaceAfter",
		"firstLineIndent",
		"hangingIndent",
		"indentLeft",
		"indentRight",
		"outlineLevel"
	]) {
		if (def[k] === void 0) continue;
		if (k === "lineSpacing") {
			const ls = parseLineSpacing(def[k], "lineSpacing");
			out[k] = `${ls.mode}:${ls.value}`;
		} else out[k] = def[k];
	}
	return out;
}
function formatResolvedFields(fields) {
	const parts = [];
	for (const [k, v] of Object.entries(fields)) parts.push(`${k}=${JSON.stringify(v)}`);
	return parts.length === 0 ? "{}" : `{ ${parts.join(", ")} }`;
}
/** Compact Δ-line per the legend in the Style Resolution header.
*  changes/+new/~matches stay in that order so the most actionable signal
*  (changes) reads first. Returns empty string when nothing to show. */
function formatDeltaLine(resolved, prior) {
	const changes = [];
	const fresh = [];
	const matches = [];
	for (const [k, v] of Object.entries(resolved)) if (!(k in prior)) fresh.push(`+${k}=${shortVal(v)}`);
	else if (sameValue(prior[k], v)) matches.push(`~${k}`);
	else changes.push(`${k} ${shortVal(prior[k])}→${shortVal(v)}`);
	return [
		...changes,
		...fresh,
		...matches
	].join(", ");
}
function shortVal(v) {
	if (typeof v === "string") return v;
	if (typeof v === "boolean") return v ? "true" : "false";
	return String(v);
}
/** Render the vs-target-direct sub-block under a styleResolution entry.
*  Hides the whole block when there's nothing to learn:
*    - 0 target paragraphs → emit a single placeholder line.
*    - All declared fields land in "new" (no override or redundant on any
*      target) → skip — sparse declaration that adds without conflict. */
function renderVsDirect(r, lines) {
	if (r.targetCount === 0) {
		lines.push(`    vs target direct: (no chrome targets; edits[]-inserts excluded from analysis)`);
		return;
	}
	if (r.fields.filter((f) => f.override > 0 || f.redundant > 0).length === 0) return;
	lines.push(`    vs target direct (${r.targetCount} paragraphs):`);
	for (const f of r.fields) {
		if (f.override === 0 && f.redundant === 0 && f.fresh === 0) continue;
		const n = r.targetCount;
		if (f.override > 0) {
			const tag = f.willStrip ? "" : " [mixed runs — direct stays]";
			const fromTo = f.overrideFrom !== void 0 ? `${shortVal(f.overrideFrom)} → ${shortVal(f.declared)}` : `${shortVal(f.declared)}`;
			lines.push(`      override:  ${f.field} ${fromTo} (${f.override}/${n})${tag}`);
		}
		if (f.redundant > 0) lines.push(`      redundant: ${f.field}=${shortVal(f.declared)} (${f.redundant}/${n} already match direct)`);
		if (f.fresh > 0) lines.push(`      new:       ${f.field}=${shortVal(f.declared)} (${f.fresh}/${n} no direct equivalent)`);
	}
}
/**
* Compute a compact pre→post paragraph index drift map from a set of dry-run
* edit previews.  Returns an empty array when there is no body-level drift
* (all ops are format / set-run / cell-container ops that don't shift body
* indices).
*
* @param entries   Per-op previews from `previewEditOps`.
* @param totalPre  Last paragraph index (inclusive) in the pre-edit document.
*                  When > 0 the map shows trailing unchanged bands; when ≤ 0
*                  the map stops after the last affected region.
*/
function computeDriftBands(entries, totalPre) {
	const bodyOps = entries.filter((e) => {
		if (e.container !== "body") return false;
		return e.willReplaceOrDeleteIndices.length > 0 || e.willInsertCount > 0;
	});
	if (bodyOps.length === 0) return [];
	const removedKind = /* @__PURE__ */ new Map();
	const offsetEvents = [];
	for (const e of bodyOps) {
		const removed = e.willReplaceOrDeleteIndices;
		if (removed.length === 0) {
			const anchor = e.targetParaIndices[0] ?? -1;
			if (anchor < 0) continue;
			const delta = e.willInsertCount;
			if (delta === 0) continue;
			if (e.op === "insert-before") offsetEvents.push({
				afterIndex: anchor - 1,
				delta
			});
			else offsetEvents.push({
				afterIndex: anchor,
				delta
			});
			continue;
		}
		const sortedRemoved = [...removed].sort((a, b) => a - b);
		const isMerge = e.op === "merge";
		const kind = isMerge ? "merged" : "deleted";
		for (const idx of sortedRemoved) {
			if (isMerge && idx === e.survivorIndex) continue;
			removedKind.set(idx, kind);
		}
		const netDelta = (isMerge ? 1 : e.willInsertCount) - sortedRemoved.length;
		if (netDelta !== 0) offsetEvents.push({
			afterIndex: sortedRemoved[sortedRemoved.length - 1],
			delta: netDelta
		});
	}
	if (removedKind.size === 0 && offsetEvents.length === 0) return [];
	offsetEvents.sort((a, b) => a.afterIndex - b.afterIndex);
	const maxPreFromOps = removedKind.size > 0 ? Math.max(...removedKind.keys(), ...offsetEvents.map((ev) => ev.afterIndex)) : Math.max(...offsetEvents.map((ev) => ev.afterIndex));
	const maxPre = totalPre > 0 ? totalPre : maxPreFromOps + 20;
	const bands = [];
	let runningOffset = 0;
	let eventIdx = 0;
	let runStart = 0;
	let runKind = "unchanged";
	let runPostStart = 0;
	function flushRun(preEnd) {
		if (runStart > preEnd) return;
		if (runKind === "unchanged") bands.push({
			kind: "unchanged",
			preStart: runStart,
			preEnd
		});
		else if (runKind === "deleted") bands.push({
			kind: "deleted",
			preStart: runStart,
			preEnd
		});
		else if (runKind === "merged") bands.push({
			kind: "merged",
			preStart: runStart,
			preEnd
		});
		else bands.push({
			kind: "shifted",
			preStart: runStart,
			preEnd,
			postStart: runPostStart
		});
	}
	for (let pre = 0; pre <= maxPre; pre++) {
		while (eventIdx < offsetEvents.length && offsetEvents[eventIdx].afterIndex < pre) {
			runningOffset += offsetEvents[eventIdx].delta;
			eventIdx++;
		}
		const thisKind = removedKind.get(pre) ?? (runningOffset !== 0 ? "shifted" : "unchanged");
		const thisPostStart = pre + runningOffset;
		if (thisKind !== runKind || thisKind === "shifted" && thisPostStart !== runPostStart + (pre - runStart)) {
			flushRun(pre - 1);
			runStart = pre;
			runKind = thisKind;
			runPostStart = thisPostStart;
		}
	}
	flushRun(maxPre);
	return bands;
}
/** Render drift bands as report lines. Returns empty array when there's nothing
*  to show (no-drift case). */
function renderDriftLines(bands) {
	if (bands.filter((b) => b.kind !== "unchanged").length === 0) return [];
	const lines = [];
	lines.push("=== Paragraph index drift (pre-edit → post-edit) ===");
	function fmtRange(start, end) {
		const w3 = (n) => String(n).padStart(3, "0");
		return start === end ? `#${w3(start)}` : `#${w3(start)}..#${w3(end)}`;
	}
	for (const b of bands) if (b.kind === "unchanged") lines.push(`  pre ${fmtRange(b.preStart, b.preEnd)}  → unchanged`);
	else if (b.kind === "deleted") lines.push(`  pre ${fmtRange(b.preStart, b.preEnd)}  → DELETED`);
	else if (b.kind === "merged") lines.push(`  pre ${fmtRange(b.preStart, b.preEnd)}  → MERGED`);
	else {
		const postEnd = b.postStart + (b.preEnd - b.preStart);
		lines.push(`  pre ${fmtRange(b.preStart, b.preEnd)}  → post ${fmtRange(b.postStart, postEnd)}`);
	}
	lines.push("");
	return lines;
}
function printReport(args) {
	const lines = [];
	lines.push(args.dryRun ? "=== Change Report (DRY RUN — no file written) ===" : "=== Change Report ===");
	lines.push(`Source: ${args.source}`);
	lines.push(`Output: ${args.output}`);
	lines.push("");
	if (args.templateImport) {
		const ti = args.templateImport;
		const directly = ti.imported.filter((id) => !ti.pulledAncestors.includes(id));
		lines.push(`Imported from template: ${directly.length} requested + ${ti.pulledAncestors.length} basedOn ancestors`);
		lines.push(`  styles: [${ti.imported.join(", ")}]`);
		if (ti.pulledAncestors.length > 0) lines.push(`  pulled ancestors: [${ti.pulledAncestors.join(", ")}]`);
		if (ti.numIdRemap.size > 0) {
			const remaps = [...ti.numIdRemap.entries()].map(([o, n]) => `${o}→${n}`).join(", ");
			lines.push(`  numIds migrated: ${remaps}`);
		}
		lines.push("");
	}
	if (args.pageSetup && args.pageSetup.sections.length > 0) {
		const ps = args.pageSetup;
		lines.push(`Page setup: applies to ${ps.sections.length} section(s); ${ps.touchedCount} mutated.`);
		for (const sec of ps.sections) {
			if (!sec.changed) continue;
			const diffs = pageSetupDiff(sec.before, sec.after);
			if (diffs.length === 0) continue;
			lines.push(`  Section ${sec.index}: ${diffs.join("; ")}`);
		}
		lines.push("");
	}
	if (args.headerFooter && args.headerFooter.parts.length > 0) {
		const hf = args.headerFooter;
		const bind = args.headerFooterBinding;
		const flags = [];
		if (hf.groups.some((g) => g.hasFirst)) flags.push("titlePg");
		if (hf.hasEven) flags.push("evenAndOddHeaders");
		const flagSuffix = flags.length > 0 ? ` (flags: ${flags.join(", ")})` : "";
		lines.push(`Header/footer: ${hf.parts.length} part(s) generated; bound to ${bind?.sectionCount ?? 0} section(s)${flagSuffix}.`);
		for (const p of hf.parts) {
			const extras = [`${p.blockCount} block(s)`];
			if (p.hasHyperlinks) extras.push("hyperlink");
			lines.push(`  ${p.surface}.${p.variant} → ${p.partName} (rId=${p.rId}, ${extras.join(", ")})`);
		}
		for (const warning of hf.separatorWarnings) lines.push(`  warning: ${warning}`);
		lines.push("");
	}
	const annotate = (id) => {
		const src = args.derivedFrom.get(id);
		return src !== void 0 ? `${id} (from #${src})` : id;
	};
	lines.push(`Styles injected: ${args.injected.length} (${args.injected.map(annotate).join(", ")})`);
	lines.push(`Styles updated:  ${args.updated.length} (${args.updated.map(annotate).join(", ")})`);
	lines.push("");
	let totalRestyled = 0;
	for (const c of args.restyleStats.values()) totalRestyled += c;
	lines.push(`Paragraphs restyled: ${totalRestyled}`);
	for (const [styleId, count] of args.restyleStats) lines.push(`  ${styleId}: ${count} paragraphs`);
	if (args.implicitKeepByFingerprint.size > 0) {
		let totalEmpty = 0;
		let totalNonEmpty = 0;
		for (const v of args.implicitKeepByFingerprint.values()) {
			totalEmpty += v.empty;
			totalNonEmpty += v.nonEmpty;
		}
		lines.push(`Paragraphs untouched: ${totalEmpty + totalNonEmpty}`);
		if (totalEmpty > 0) lines.push(`  empty (likely spacers): ${totalEmpty}`);
		if (totalNonEmpty > 0) {
			lines.push(`  non-empty (verify coverage): ${totalNonEmpty}`);
			const sortedEntries = [...args.implicitKeepByFingerprint.entries()].filter(([, v]) => v.nonEmpty > 0).sort((a, b) => b[1].nonEmpty - a[1].nonEmpty);
			for (const [fp, v] of sortedEntries) {
				const samples = v.nonEmptySamples.length > 0 ? `  e.g. ${v.nonEmptySamples.map((s) => `"${s}"`).join(" / ")}` : "";
				lines.push(`    ${fp}×${v.nonEmpty}${samples}`);
			}
		}
	}
	lines.push("");
	if (args.excludeSamples && args.excludeSamples.length > 0) {
		lines.push(`Excluded paragraphs (${args.excludeSamples.length}; will not be touched):`);
		const cap = 15;
		for (const s of args.excludeSamples.slice(0, cap)) lines.push(`  #${s.index}  "${s.snippet}"`);
		if (args.excludeSamples.length > cap) lines.push(`  … (${args.excludeSamples.length - cap} more)`);
		lines.push("  → Verify these indices still match — if document order shifted,");
		lines.push("    exclude entries silently aim at the wrong paragraphs.");
		lines.push("");
	}
	if (args.numberingAllocation.length > 0) {
		lines.push("=== Numbering schemes ===");
		for (const entry of args.numberingAllocation) {
			const label = `scheme[${entry.schemeIndex}]`;
			const tag = entry.explicit ? "(explicit)" : "(allocated)";
			lines.push(`  ${label.padEnd(12)} → numId=${entry.numId} ${tag}`);
		}
		lines.push("");
	}
	if (args.numberingBindings.length > 0) {
		lines.push("Auto-numbering bindings (will prepend at render):");
		for (const b of args.numberingBindings) lines.push(`  ${b.styleId} → "${b.lvlText}" + suff:${b.suff} (level ${b.level})`);
		lines.push("");
	}
	if (args.manualNumberingRemoved.size > 0) {
		lines.push("Manual numbering converted:");
		for (const [pat, count] of args.manualNumberingRemoved) lines.push(`  Prefix removed: "${pat}" (${count})`);
		lines.push("");
	}
	const mixedStyles = [...args.manualNumberingByStyle.entries()].filter(([, m]) => m.size >= 2);
	if (mixedStyles.length > 0) {
		lines.push("Mixed manual numbering detected (source inconsistent):");
		for (const [styleId, patMap] of mixedStyles) {
			const breakdown = [...patMap.entries()].sort((a, b) => b[1] - a[1]).map(([pat, n]) => `"${pat}"×${n}`).join(", ");
			lines.push(`  ${styleId}: ${breakdown}`);
		}
		lines.push("  Normalization unified these to one scheme — worth confirming with the user before final write.");
		lines.push("");
	}
	if (args.unstrippedByStyle.size > 0) {
		lines.push("Numbered-style paragraphs not matched by any stripPrefixPattern:");
		for (const [styleId, info] of args.unstrippedByStyle) {
			const samples = info.samples.map((s) => `"${s}"`).join(" / ");
			lines.push(`  ${styleId}: ${info.count} paragraphs  e.g. ${samples}`);
		}
		lines.push("  → Read the samples: a typed prefix you missed (add to stripPrefixPatterns)");
		lines.push("    or a clean heading without manual numbering (no action needed).");
		lines.push("");
	}
	if (args.manualNumberingDetected && args.manualNumberingDetected.size > 0) {
		const bound = [...args.manualNumberingDetected].filter(([, v]) => v.kind === "bound");
		const unbound = [...args.manualNumberingDetected].filter(([, v]) => v.kind === "unbound");
		lines.push("Manual numbering detected in edits[] paragraphs:");
		if (bound.length > 0) {
			for (const [styleId, info] of bound) {
				const samples = info.samples.map((s) => `"${s}"`).join(" / ");
				lines.push(`  ${styleId} (bound to numbering): ${info.count} paragraphs  e.g. ${samples}`);
			}
			lines.push("    → Drop the typed prefix from `text` — the scheme's lvlText already emits the marker.");
		}
		if (unbound.length > 0) {
			for (const [styleId, info] of unbound) {
				const samples = info.samples.map((s) => `"${s}"`).join(" / ");
				lines.push(`  ${styleId} (no numbering binding): ${info.count} paragraphs  e.g. ${samples}`);
			}
			lines.push("    → Declare a list / heading style bound to a numbering scheme;");
			lines.push("      drop the typed prefix from `text` and let lvlText emit the marker.");
			lines.push("      User-supplied content with typed prefixes is input shape, not a");
			lines.push("      directive to preserve them.");
		}
		lines.push("");
	}
	if (args.patternMatchStats.size > 0) {
		lines.push("Pattern rules matched:");
		for (const [src, count] of args.patternMatchStats) {
			const stripped = args.patternStripStats.get(src) ?? 0;
			const stripNote = stripped > 0 ? ` (stripped match in ${stripped})` : "";
			lines.push(`  /${src}/: ${count} paragraphs${stripNote}`);
		}
		lines.push("");
	}
	if (args.flags.length > 0) {
		lines.push(`Flagged (not modified): ${args.flags.length}`);
		for (const f of args.flags) lines.push(`  #${f.paraIndex}: ${f.reason}`);
		lines.push("");
	}
	if (args.styleResolutions.length > 0) {
		lines.push("=== Style Resolution (verify by reading) ===");
		const hasAnySpec = args.styleResolutions.some((r) => r.userSpec !== null);
		if (hasAnySpec) {
			lines.push("  The script does not parse natural language. For styles with a");
			lines.push("  user spec, compare it to the agent-resolved fields by eye —");
			lines.push("  any mismatch means the agent's translation needs adjustment.");
			lines.push("  Styles without a spec are still listed so the resolved fields");
			lines.push("  are auditable.");
		} else lines.push("  No `requirements` entries declared — Agent Resolved fields below reflect the structured `styles[]` declarations directly.");
		lines.push("");
		if (args.styleResolutions.some((r) => r.priorState !== null)) {
			lines.push("  Δ-line vs source (when styleId existed pre-apply):");
			lines.push("    A→B    field changed; source had A, declaration sets B");
			lines.push("    +field new declaration; source didn't have this field");
			lines.push("    ~field declaration matches source's cascade value (may be redundant)");
			lines.push("");
		}
		for (const r of args.styleResolutions) {
			const freshTag = r.priorState === null ? "  [fresh]" : "";
			lines.push(`  ${r.styleId}${freshTag}`);
			if (hasAnySpec) if (r.userSpec !== null) lines.push(`    User specified: "${r.userSpec}"`);
			else lines.push(`    User specified: (none — no requirements entry)`);
			lines.push(`    Agent resolved: ${formatResolvedFields(r.resolved)}`);
			if (r.priorState !== null) {
				const delta = formatDeltaLine(r.resolved, r.priorState);
				if (delta) lines.push(`    Δ vs source:    ${delta}`);
			}
			if (r.priorUsage && r.priorUsage > 0) lines.push(`    Existing usage: ${r.priorUsage} paragraph(s) — will re-render with the new definition.`);
			if (r.warnings) for (const wMsg of r.warnings) lines.push(`    note: ${wMsg}`);
			if (r.vsDirect) renderVsDirect(r.vsDirect, lines);
		}
		lines.push("");
	}
	if (args.editsPreview.length > 0) {
		lines.push("=== Edits Preview (locator-resolved; applied in memory, not written to disk in dry-run) ===");
		let totalReplaceDelete = 0;
		let totalInsert = 0;
		for (const e of args.editsPreview) {
			const targetSpan = e.targetParaIndices.length === 0 ? "(end-of-body)" : e.targetParaIndices.length === 1 ? `#${e.targetParaIndices[0]}` : `#${e.targetParaIndices[0]}–#${e.targetParaIndices[e.targetParaIndices.length - 1]} (${e.targetParaIndices.length})`;
			const insertNote = e.willInsertCount > 0 ? `; +${e.willInsertCount} new` : "";
			const replaceNote = e.willReplaceOrDeleteIndices.length > 0 ? `; -${e.willReplaceOrDeleteIndices.length} replaced/deleted` : "";
			const containerNote = e.container === "cell" ? " [in cell]" : "";
			lines.push(`  edits[${e.index}] ${e.op} → ${targetSpan}${containerNote}${replaceNote}${insertNote}`);
			totalReplaceDelete += e.willReplaceOrDeleteIndices.length;
			totalInsert += e.willInsertCount;
		}
		lines.push(`  Total: ${args.editsPreview.length} ops; ${totalReplaceDelete} paragraphs replaced/deleted, ${totalInsert} new paragraphs inserted.`);
		lines.push("  Note: implicit-keep counts above already exclude paragraphs the edits[] pass will replace/delete.");
		lines.push("");
	}
	if (args.dryRun && args.editsPreview.length > 0) {
		const bands = computeDriftBands(args.editsPreview, args.totalParagraphs ?? -1);
		for (const line of renderDriftLines(bands)) lines.push(line);
	}
	if (args.captionsPreview) {
		const cp = args.captionsPreview;
		lines.push("=== Captions / Cross-Refs Preview (dry-run) ===");
		lines.push(`  Chapter SEQs injected: ${cp.chapterSeqsInjected}  (hidden auto-counters on outline paragraphs)`);
		lines.push(`  Caption paragraphs re-emitted: ${cp.standardizeReemitted}  (existing captions rebuilt against current config)`);
		lines.push(`  Caption paragraphs freshly emitted: ${cp.freshlyEmitted}  (new captions inserted by edits[])`);
		if (cp.samples.length > 0) {
			lines.push(`  Predicted text (first ${cp.samples.length}):`);
			for (const s of cp.samples) lines.push(`    [${s.identifier}] "${s.text}"`);
		} else lines.push(`  Predicted text: (no caption samples — config has captions but body emits none)`);
		lines.push("");
	}
	if (args.panguWarnings && args.panguWarnings.length > 0) {
		const w = args.panguWarnings;
		lines.push("=== Possible Pangu spacing in author-supplied text ===");
		lines.push("  Word's autoSpace handles CJK ↔ Latin/digit gaps; typed ASCII spaces stack on top and render too wide.");
		lines.push("  Advisory only — never blocks the write. If the spacing is deliberate, set lint.panguSpacing=false to silence.");
		const shown = w.slice(0, 5);
		for (const entry of shown) lines.push(`  ${entry.source}  ...${entry.snippet}...   (matched "${entry.hit}")`);
		if (w.length > shown.length) lines.push(`  (... ${w.length - shown.length} more not shown)`);
		lines.push("");
	}
	if (args.settingsElements && args.settingsElements.length > 0) {
		lines.push("=== word/settings.xml elements (settings.elements[]) ===");
		for (const e of args.settingsElements) lines.push(`  w:${e.name}  [${e.action}]${e.action === "updated" ? " — merged into the existing element (no duplicate)" : ""}`);
		lines.push("");
	}
	if (args.docProps) {
		const dp = args.docProps;
		lines.push("=== docProps (core.xml / app.xml) ===");
		lines.push(`  core.xml ${dp.coreCreated ? "created" : "updated"} · app.xml ${dp.appCreated ? "created" : "updated"} · revision ${dp.revision}`);
		lines.push(`  Counters written: Words=${dp.words}  Paragraphs=${dp.paragraphs}`);
		lines.push("");
	}
	if (args.samples.size > 0) {
		lines.push("=== Sample Affected Paragraphs (first per style) ===");
		for (const [styleId, samples] of args.samples) {
			lines.push(`  ${styleId}:`);
			for (const s of samples) {
				const notes = s.notes.length > 0 ? `  [${s.notes.join("; ")}]` : "";
				lines.push(`    #${s.paraIndex} via=${s.via}${s.patternSource ? ` /${s.patternSource}/` : ""}: "${s.textPreview}"${notes}`);
			}
		}
		lines.push("");
	}
	if (args.dryRun) {
		lines.push("Dry run — no file written, no validation performed.");
		lines.push("Re-run without --dry-run to commit changes.");
	} else {
		lines.push("Validation: PASS");
		lines.push(`Output: ${args.output}`);
	}
	console.log(lines.join("\n"));
}
/** Per-section before/after summary lines. Returns only fields that actually
*  changed so the report stays sparse. Twips → mm + cm for readability. */
function pageSetupDiff(before, after) {
	const out = [];
	if (before.paperSize !== after.paperSize) out.push(`paper ${before.paperSize ?? "?"} → ${after.paperSize ?? "?"}`);
	if (before.orientation !== after.orientation) out.push(`orient ${before.orientation} → ${after.orientation}`);
	const mEdges = [
		"top",
		"bottom",
		"left",
		"right"
	];
	const mDiffs = [];
	for (const e of mEdges) {
		const b = before.margins[e];
		const a = after.margins[e];
		if (b !== a) mDiffs.push(`${e} ${twipsToCmString(b)} → ${twipsToCmString(a)}`);
	}
	if (mDiffs.length > 0) out.push(`margins(${mDiffs.join(", ")})`);
	return out;
}

//#endregion
//#region lib/parse/style-names.ts
/**
* Build a `styleId → StyleInfo` resolver from a styles.xml `Document`.
*
* Used by `emitInlineStyleRef` to drive a 3-way dispatch on the
* referenced style:
*   1. style has `<w:outlineLvl>` → emit `STYLEREF N` (locale-neutral)
*   2. style is a custom name (not in the built-in localizable set) →
*      emit `STYLEREF "<name>"` (Word doesn't translate custom names)
*   3. style is a built-in localizable non-outline style (Title /
*      Caption / Subtitle / etc.) → caller should reject, since
*      `STYLEREF "<English-name>"` silently fails in non-EN Word UIs
*
* Returns `undefined` for unknown styleIds — caller decides whether
* to throw or fall back.
*/
const w$17 = NS.w;
function buildStyleResolver(stylesDoc) {
	if (!stylesDoc) return () => void 0;
	const root = stylesDoc.documentElement;
	if (!root) return () => void 0;
	const map = /* @__PURE__ */ new Map();
	for (const styleEl of getChildrenNS(root, w$17, "style")) {
		const id = wAttr(styleEl, "styleId");
		if (!id) continue;
		const nameEl = firstChildNS(styleEl, w$17, "name");
		const name = nameEl ? wAttr(nameEl, "val") ?? id : id;
		let outlineLevel;
		const pPr = firstChildNS(styleEl, w$17, "pPr");
		if (pPr) {
			const lvlEl = firstChildNS(pPr, w$17, "outlineLvl");
			if (lvlEl) {
				const v = wAttr(lvlEl, "val");
				const n = v !== null && v !== void 0 ? parseInt(v, 10) : NaN;
				if (Number.isFinite(n) && n >= 0 && n <= 8) outlineLevel = n + 1;
			}
		}
		map.set(id, {
			name,
			outlineLevel,
			isBuiltInLocalizable: isBuiltInLocalizable(name)
		});
	}
	return (id) => map.get(id);
}

//#endregion
//#region lib/edit/fields/complex-field.ts
/**
* Shared 5-run skeleton for OOXML complex fields.
*
* Every complex field (REF / SEQ / STYLEREF / ...) renders as:
*
*   <w:r>[ rPr? ]<w:fldChar fldCharType="begin"/></w:r>
*   <w:r>[ rPr? ]<w:instrText xml:space="preserve"> <code> </w:instrText></w:r>
*   <w:r>[ rPr? ]<w:fldChar fldCharType="separate"/></w:r>
*   <w:r>[ rPr? ]<w:t>[result]</w:t></w:r>
*   <w:r>[ rPr? ]<w:fldChar fldCharType="end"/></w:r>
*
* The result run carries the user-visible text Word displays BEFORE field
* update; after F9 (or auto-update on open via `updateFields=true` in
* settings.xml) Word replaces it with the resolved value. Field-type-
* specific emitters supply `instrCode` (without surrounding spaces) and an
* optional `initialResult`; the skeleton handles the run layout, the
* `xml:space="preserve"` wrapping, and the format-preservation contract.
*
* `format`: optional `RunFormat`. When non-empty, `applyFieldFormat` (1)
* replicates the same rPr onto every field run and (2) appends
* `\* MERGEFORMAT` to the field code. Both are required together — Word's
* field-update logic rewrites the result run on F9 and the new run
* inherits rPr from the surrounding field runs, not from the prior result
* run we wrote; MERGEFORMAT tells Word to preserve that rPr across
* subsequent updates. Either alone is insufficient.
*/
const w$16 = NS.w;
const XML_NS = "http://www.w3.org/XML/1998/namespace";
/** Emit a complex field's 5-run sequence. Returns the runs plus direct
* references to the result run and its text element for callers that
* need to backfill the placeholder later (e.g. REF after counter sim). */
function emitComplexField(ownerDoc, spec) {
	const begin = ownerDoc.createElementNS(w$16, "w:r");
	const beginFld = ownerDoc.createElementNS(w$16, "w:fldChar");
	beginFld.setAttributeNS(w$16, "w:fldCharType", "begin");
	begin.appendChild(beginFld);
	const instr = ownerDoc.createElementNS(w$16, "w:r");
	const instrText = ownerDoc.createElementNS(w$16, "w:instrText");
	instrText.setAttributeNS(XML_NS, "xml:space", "preserve");
	instr.appendChild(instrText);
	const separate = ownerDoc.createElementNS(w$16, "w:r");
	const sepFld = ownerDoc.createElementNS(w$16, "w:fldChar");
	sepFld.setAttributeNS(w$16, "w:fldCharType", "separate");
	separate.appendChild(sepFld);
	const resultRun = ownerDoc.createElementNS(w$16, "w:r");
	const resultTextEl = ownerDoc.createElementNS(w$16, "w:t");
	resultTextEl.setAttributeNS(XML_NS, "xml:space", "preserve");
	resultTextEl.textContent = spec.initialResult ?? "";
	resultRun.appendChild(resultTextEl);
	const end = ownerDoc.createElementNS(w$16, "w:r");
	const endFld = ownerDoc.createElementNS(w$16, "w:fldChar");
	endFld.setAttributeNS(w$16, "w:fldCharType", "end");
	end.appendChild(endFld);
	const { instrTextSuffix } = applyFieldFormat([
		begin,
		instr,
		separate,
		resultRun,
		end
	], spec.format, ownerDoc);
	instrText.textContent = ` ${spec.instrCode}${instrTextSuffix} `;
	return {
		runs: [
			begin,
			instr,
			separate,
			resultRun,
			end
		],
		resultRun,
		resultTextEl
	};
}
/** Replicate agent-declared rPr across every run + emit `\\* MERGEFORMAT`
* iff the format produces non-empty rPr. Empty format (`{}` /
* `{ color: undefined }`) treated as no-format — matches inheritance
* semantics in the edit engine.
*
* Returns the field-code suffix to append (empty when no format). */
function applyFieldFormat(fieldRuns, format, ownerDoc) {
	if (!format) return { instrTextSuffix: "" };
	const rPr = ownerDoc.createElementNS(w$16, "w:rPr");
	for (const c of buildRPrChildren(format, ownerDoc)) rPr.appendChild(c);
	if (rPr.childNodes.length === 0) return { instrTextSuffix: "" };
	for (const r of fieldRuns) r.insertBefore(rPr.cloneNode(true), r.firstChild);
	return { instrTextSuffix: " \\* MERGEFORMAT" };
}

//#endregion
//#region lib/edit/fields/inline-fields.ts
/**
* Inline field nodes — thin wrappers over `emitComplexField`.
*
*   `{ field: "page"      }` → PAGE
*   `{ field: "numPages"  }` → NUMPAGES
*   `{ field: "date"      }` → DATE
*   `{ styleRef: "..."    }` → STYLEREF "<name>"
*   `{ styleRef: "...", numberOnly: true }` → STYLEREF "<name>" \n
*
* The placeholder text shown before Word's field-update pass is a sensible
* default per field type — Word replaces it on next open thanks to the
* `updateFields` flag set in settings.xml.
*
* STYLEREF requires the referenced styleId to exist when Word resolves the
* field. Validation is the engine's job (apply reads stylesDoc and throws
* on missing styleId); zod doesn't have that context.
*/
const FIELD_DESCRIPTORS = {
	page: {
		instr: "PAGE",
		placeholder: "1"
	},
	numPages: {
		instr: "NUMPAGES",
		placeholder: "1"
	},
	date: {
		instr: "DATE",
		placeholder: ""
	}
};
function emitInlineField(ownerDoc, field, format) {
	const { instr, placeholder } = FIELD_DESCRIPTORS[field];
	const { runs } = emitComplexField(ownerDoc, {
		instrCode: instr,
		initialResult: placeholder,
		format
	});
	return runs;
}
function emitInlineStyleRef(ownerDoc, info, numberOnly, format) {
	const switches = numberOnly ? " \\n" : "";
	let argument;
	if (info.outlineLevel !== void 0) argument = String(info.outlineLevel);
	else if (!info.isBuiltInLocalizable) argument = `"${info.name}"`;
	else throw new Error(`InlineStyleRef: cannot reference built-in style "${info.name}" via STYLEREF. Word translates built-in style names per UI locale (e.g. "Title" ↔ "标题", "Caption" ↔ "题注"), so STYLEREF "${info.name}" silently fails in non-English Word. Fix by either: (a) binding this style to an outline level (set styles[].outlineLevel = 0-8) so the engine can emit STYLEREF N (locale-neutral); or (b) renaming the style to a custom name (e.g. "Doc${info.name}") — custom names aren't translated.`);
	const { runs } = emitComplexField(ownerDoc, {
		instrCode: `STYLEREF ${argument}${switches}`,
		initialResult: "",
		format
	});
	return runs;
}

//#endregion
//#region lib/edit/table-emit.ts
/**
* TableBlock → OOXML emission.
*
* Translates the JSON TableBlock into `<w:tbl>` + `<w:tblPr>` + `<w:tblGrid>`
* + `<w:tr>` + `<w:tc>` with all merge / borders / cell-property emit. The
* agent declares only "restart" cells (those with content); engine computes
* a grid occupancy map and injects vMerge continuation cells where rowspans
* claim cells in subsequent rows.
*
* Three OOXML correctness invariants enforced here:
*
*   1. CT_TblPr / CT_TrPr / CT_TcPr child order — Word loaders silently
*      mis-render or refuse mis-ordered properties. Child-order helpers
*      from `xml-order.ts` enforce it.
*
*   2. `<w:tblGrid>` matches effective column count — Word relies on this
*      to lay out columns. Engine auto-generates from row structure if
*      agent omits `cols`; if `cols` is provided, length must match.
*
*   3. Every `<w:tc>` contains at least one block-level child — Word
*      rejects an empty cell with "needs repair". Empty content (empty
*      string, empty Block[], empty InlineNode[]) is replaced with a
*      single empty `<w:p/>`.
*
* Recursive cell content (TableBlock inside cells) is NOT supported in v1
* by schema — CellBlockSchema excludes TableBlock. The grid occupancy
* algorithm and continuation logic still applies to one-level tables.
*/
const w$15 = NS.w;
function emitTableBlock(block, ownerDoc, ctx) {
	const headerRows = block.headerRows ?? 0;
	const rows = block.sort ? sortTableRows(block, headerRows) : block.rows;
	const { grid, effectiveCols } = buildGrid(rows, headerRows);
	if (block.cols !== void 0 && block.cols.length !== effectiveCols) throw new Error(`table.cols length (${block.cols.length}) does not match effective column count (${effectiveCols}). Effective columns = max(declared cells per row + ongoing rowspan claims), expanded by colspans.`);
	const tbl = ownerDoc.createElementNS(w$15, "w:tbl");
	const usableWidth = ctx.usableWidthTwips ?? DEFAULT_USABLE_WIDTH_TWIPS;
	const autoShareTwips = computeAutoShareTwips(block.cols, effectiveCols, usableWidth);
	const tblPr = buildTblPr(block, autoShareTwips, usableWidth, ownerDoc);
	tbl.appendChild(tblPr);
	tbl.appendChild(buildTblGrid(block.cols, effectiveCols, autoShareTwips, ownerDoc));
	for (let r = 0; r < grid.length; r++) {
		const rowEl = ownerDoc.createElementNS(w$15, "w:tr");
		if (r < headerRows) {
			const trPr = ownerDoc.createElementNS(w$15, "w:trPr");
			insertChildInOrder(trPr, ownerDoc.createElementNS(w$15, "w:tblHeader"), TR_PR_CHILD_ORDER);
			rowEl.appendChild(trPr);
		}
		for (const slot of grid[r]) rowEl.appendChild(buildTc(slot, block, ownerDoc, ctx));
		tbl.appendChild(rowEl);
	}
	return tbl;
}
/** Walk declared cells per row; weave in continuation slots for ongoing
* rowspans. Validates span bounds. Returns a fully resolved grid. */
function buildGrid(rows, headerRows) {
	const grid = rows.map(() => []);
	const ongoing = /* @__PURE__ */ new Map();
	let effectiveCols;
	for (let r = 0; r < rows.length; r++) {
		const row = rows[r];
		let col = 0;
		let declaredIdx = 0;
		while (declaredIdx < row.length || ongoing.has(col)) {
			if (ongoing.has(col)) {
				const active = ongoing.get(col);
				const cont = {
					kind: "vmerge-continue",
					cell: active.source.cell,
					colspan: active.source.colspan,
					rowspan: active.source.rowspan,
					isHeaderRow: r < headerRows,
					rowIndex: r
				};
				grid[r].push(cont);
				active.rowsRemaining -= 1;
				if (active.rowsRemaining <= 0) ongoing.delete(col);
				col += active.source.colspan;
				continue;
			}
			if (declaredIdx >= row.length) break;
			const cellObj = normalizeCell(row[declaredIdx]);
			const colspan = cellObj.colspan ?? 1;
			const rowspan = cellObj.rowspan ?? 1;
			if (rowspan > rows.length - r) throw new Error(`table row ${r + 1} col ${col + 1}: rowspan ${rowspan} exceeds rows remaining (${rows.length - r}).`);
			const slot = {
				kind: "restart",
				cell: cellObj,
				colspan,
				rowspan,
				isHeaderRow: r < headerRows,
				rowIndex: r
			};
			grid[r].push(slot);
			if (rowspan > 1) ongoing.set(col, {
				rowsRemaining: rowspan - 1,
				source: slot
			});
			declaredIdx += 1;
			col += colspan;
		}
		while (ongoing.has(col)) {
			const active = ongoing.get(col);
			const cont = {
				kind: "vmerge-continue",
				cell: active.source.cell,
				colspan: active.source.colspan,
				rowspan: active.source.rowspan,
				isHeaderRow: r < headerRows,
				rowIndex: r
			};
			grid[r].push(cont);
			active.rowsRemaining -= 1;
			if (active.rowsRemaining <= 0) ongoing.delete(col);
			col += active.source.colspan;
		}
		if (effectiveCols === void 0) effectiveCols = col;
		else if (col !== effectiveCols) throw new Error(`table row ${r + 1}: declared cells + ongoing rowspans span ${col} columns; row 1 spans ${effectiveCols}. All rows must total the same column count.`);
	}
	if (ongoing.size > 0) throw new Error(`table: rowspan extends past the last row at column(s) [${[...ongoing.keys()].join(", ")}].`);
	return {
		grid,
		effectiveCols: effectiveCols ?? 0
	};
}
function normalizeCell(cell) {
	if (typeof cell === "string") return { content: cell };
	if (Array.isArray(cell)) return { content: cell };
	return cell;
}
/** Recursively extract a plain-text sort key from any cell-content form:
* string → itself; InlineNode[]/Block[] → concatenated `text` fields;
* object form → its `content` (or bare `text` for an inline node). */
function cellSortText(cell) {
	if (cell === null || cell === void 0) return "";
	if (typeof cell === "string") return cell;
	if (Array.isArray(cell)) return cell.map(cellSortText).join("");
	if (typeof cell === "object") {
		if (typeof cell.text === "string") return cell.text;
		if (cell.content !== void 0) return cellSortText(cell.content);
		return "";
	}
	return String(cell);
}
/** Apply `block.sort`: pin header rows, sort body rows by the target
* column. Rejects rowspans (reordering would orphan vMerge continuations)
* and out-of-range columns. Numeric on both sides → numeric compare
* (thousands separators stripped); otherwise zh-CN collation. */
function sortTableRows(block, headerRows) {
	const sort = block.sort;
	const rows = block.rows;
	const colIdx = sort.column - 1;
	const maxLen = Math.max(...rows.map((row) => row.length));
	if (colIdx >= maxLen) throw new Error(`table.sort.column ${sort.column} exceeds the effective column count (${maxLen}).`);
	for (let r = 0; r < rows.length; r++) for (const c of rows[r]) {
		const obj = typeof c === "object" && c !== null && !Array.isArray(c) ? c : void 0;
		if (obj && obj.rowspan !== void 0 && obj.rowspan > 1) throw new Error(`table.sort: row ${r + 1} declares a rowspan-${obj.rowspan} cell; reordering rows would orphan its vMerge continuations. Remove the rowspan or sort the data before passing it to apply.`);
	}
	const dir = sort.order === "desc" ? -1 : 1;
	const head = rows.slice(0, headerRows);
	const body = rows.slice(headerRows);
	body.sort((a, b) => {
		const ka = cellSortText(a[colIdx]).trim();
		const kb = cellSortText(b[colIdx]).trim();
		const na = Number(ka.replace(/,/g, ""));
		const nb = Number(kb.replace(/,/g, ""));
		let cmp;
		if (ka !== "" && kb !== "" && Number.isFinite(na) && Number.isFinite(nb)) cmp = na < nb ? -1 : na > nb ? 1 : 0;
		else cmp = ka.localeCompare(kb, "zh-Hans-CN");
		return dir * cmp;
	});
	return [...head, ...body];
}
/** Fallback usable width when ctx doesn't supply one — covers ops targeting
* an existing `<w:tc>` via a `cell` locator (a TableBlock inserted into a
* cell can't know the cell's width without reading `<w:tcW>`; deferred
* out of v1), sections with missing pgSz/pgMar, and pre-emit synthetic
* paths. A4 portrait with 30 mm side margins gives 150 mm ≈ 8504 twips;
* 8500 stays just inside that. For an `autofit` table the fallback only
* sets a seed and Word reflows on open; for a `fixed` table inserted into
* a narrower cell the fallback overflows permanently — agent should
* declare `cols` widths to match the cell. */
const DEFAULT_USABLE_WIDTH_TWIPS = 8500;
function buildTblPr(block, autoShareTwips, usableWidth, ownerDoc) {
	const tblPr = ownerDoc.createElementNS(w$15, "w:tblPr");
	if (block.bidiVisual !== void 0) {
		// Table direction: mirrors the column order (Word's 从右向左).
		// Sits right after tblOverlap, before tblW.
		insertChildInOrder(tblPr, toggleElement(ownerDoc, "w:bidiVisual", block.bidiVisual), TBL_PR_CHILD_ORDER);
	}
	if (block.alignment) {
		const jc = ownerDoc.createElementNS(w$15, "w:jc");
		jc.setAttributeNS(w$15, "w:val", block.alignment);
		insertChildInOrder(tblPr, jc, TBL_PR_CHILD_ORDER);
	}
	const tblW = ownerDoc.createElementNS(w$15, "w:tblW");
	if (block.layout === "fixed") {
		const totalTwips = computeFixedTotalTwips(block.cols, autoShareTwips, usableWidth);
		tblW.setAttributeNS(w$15, "w:w", String(totalTwips));
		tblW.setAttributeNS(w$15, "w:type", "dxa");
	} else {
		tblW.setAttributeNS(w$15, "w:w", "0");
		tblW.setAttributeNS(w$15, "w:type", "auto");
	}
	insertChildInOrder(tblPr, tblW, TBL_PR_CHILD_ORDER);
	const borders = resolveTableBorders(block);
	if (borders) insertChildInOrder(tblPr, buildTblBorders(borders, ownerDoc), TBL_PR_CHILD_ORDER);
	if (block.layout === "fixed") {
		const tblLayout = ownerDoc.createElementNS(w$15, "w:tblLayout");
		tblLayout.setAttributeNS(w$15, "w:type", "fixed");
		insertChildInOrder(tblPr, tblLayout, TBL_PR_CHILD_ORDER);
	}
	const tableEdges = resolveTablePadding(block);
	if (tableEdges !== void 0) insertChildInOrder(tblPr, buildCellMar("tblCellMar", tableEdges, ownerDoc), TBL_PR_CHILD_ORDER);
	return tblPr;
}
/** Share allocated to each `"auto"` column: distribute the leftover budget
* (usable width minus the sum of explicit columns) evenly across auto slots.
* Same number is used for gridCol seeds AND the fixed-mode `tblW` total —
* if the two diverged, fixed-mode tables would either overflow the page or
* leave a phantom gap between declared widths and the table edge. Returns
* the even-split share when `cols` is omitted (all columns are "auto"). */
function computeAutoShareTwips(cols, effectiveCols, usableWidth) {
	if (cols === void 0) return Math.max(0, Math.round(usableWidth / effectiveCols));
	let explicitSum = 0;
	let autoCount = 0;
	for (const c of cols) if (c.width === "auto") autoCount += 1;
	else explicitSum += toTwips(c.width, "table.cols.width");
	if (autoCount === 0) return 0;
	return Math.max(0, Math.round((usableWidth - explicitSum) / autoCount));
}
function computeFixedTotalTwips(cols, autoShareTwips, usableWidth) {
	if (cols === void 0) return usableWidth;
	let sum = 0;
	for (const c of cols) if (c.width === "auto") sum += autoShareTwips;
	else sum += toTwips(c.width, "table.cols.width");
	return sum;
}
/** Resolve preset → BordersCustom + apply default. Returns null when borders
* should be omitted entirely (only when `borders: "none"` is requested AND
* we want to suppress the element entirely — but it's safer to always emit
* tblBorders to override theme defaults). */
function resolveTableBorders(block) {
	const b = block.borders ?? "all";
	if (typeof b === "string") switch (b) {
		case "all": return {
			top: "single",
			bottom: "single",
			left: "single",
			right: "single",
			insideH: "single",
			insideV: "single"
		};
		case "none": return {
			top: "none",
			bottom: "none",
			left: "none",
			right: "none",
			insideH: "none",
			insideV: "none"
		};
		case "outer": return {
			top: "single",
			bottom: "single",
			left: "single",
			right: "single",
			insideH: "none",
			insideV: "none"
		};
		case "three-line": return {
			top: "thick",
			bottom: "thick",
			left: "none",
			right: "none",
			insideH: "none",
			insideV: "none"
		};
	}
	return {
		top: b.top ?? "none",
		bottom: b.bottom ?? "none",
		left: b.left ?? "none",
		right: b.right ?? "none",
		insideH: b.insideH ?? "none",
		insideV: b.insideV ?? "none"
	};
}
function buildTblBorders(b, ownerDoc) {
	const el = ownerDoc.createElementNS(w$15, "w:tblBorders");
	for (const side of [
		"top",
		"left",
		"bottom",
		"right",
		"insideH",
		"insideV"
	]) {
		const edge = b[side];
		if (!edge) continue;
		el.appendChild(buildBorderElement("w:" + side, edge, ownerDoc));
	}
	return el;
}
function buildTcBorders(b, ownerDoc) {
	const el = ownerDoc.createElementNS(w$15, "w:tcBorders");
	for (const side of [
		"top",
		"left",
		"bottom",
		"right",
		"insideH",
		"insideV"
	]) {
		const edge = b[side];
		if (!edge) continue;
		el.appendChild(buildBorderElement("w:" + side, edge, ownerDoc));
	}
	return el;
}
/** Builds one OOXML border edge element (e.g. `<w:top>`, `<w:bottom>`) from
* a config-shape `BorderEdge`. Shared between table borders (tblBorders /
* tcBorders, space=0) and paragraph borders (pBdr, space=1 — text needs a
* gap from the line). */
function buildBorderElement(qname, edge, ownerDoc, space = 0) {
	const el = ownerDoc.createElementNS(w$15, qname);
	if (edge === "none") {
		el.setAttributeNS(w$15, "w:val", "nil");
		return el;
	}
	let style;
	let sizeEighthPt;
	let color = "auto";
	if (typeof edge === "string") {
		style = edge;
		let defaultPt = edge === "thick" ? 1.5 : .5;
		if (edge === "double") defaultPt = .75;
		sizeEighthPt = toEighthPt(defaultPt, "border.size");
	} else {
		style = edge.style;
		const defaultPt = edge.style === "thick" ? 1.5 : .5;
		sizeEighthPt = edge.size !== void 0 ? toEighthPt(edge.size, "border.size") : toEighthPt(defaultPt, "border.size");
		color = edge.color ?? "auto";
	}
	if (style === "thick") style = "single";
	el.setAttributeNS(w$15, "w:val", style);
	el.setAttributeNS(w$15, "w:sz", String(Math.max(2, sizeEighthPt)));
	el.setAttributeNS(w$15, "w:space", String(space));
	el.setAttributeNS(w$15, "w:color", color);
	return el;
}
/** Three-line preset default vertical padding. Without it, single-line cells
* + auto row height + skill-default `vAlign: "center"` would render
* indistinguishably from top-aligned (row height = content height, vAlign has
* no slack). 4pt gives ~8pt vertical room for centering to read visually,
* matching academic typesetting conventions (LaTeX `booktabs`-ish). Only
* applies when the table doesn't carry an explicit `padding`. */
const THREE_LINE_DEFAULT_VPAD_PT = 4;
/** Resolve table-level padding to the edges that should be emitted as
* `<w:tblCellMar>` children. Explicit `block.padding` always emits all four
* edges (CSS-faithful — `padding: 0` flattens everything, including
* TableNormal's left/right). The three-line preset default emits only
* top/bottom so left/right inherit TableNormal (≈5.4pt). */
function resolveTablePadding(block) {
	if (block.padding !== void 0) return parsePadding(block.padding, "table.padding");
	if (block.borders === "three-line") return {
		top: THREE_LINE_DEFAULT_VPAD_PT,
		bottom: THREE_LINE_DEFAULT_VPAD_PT
	};
}
function buildCellMar(tagName, edges, doc) {
	const el = doc.createElementNS(w$15, "w:" + tagName);
	for (const side of [
		"top",
		"left",
		"bottom",
		"right"
	]) {
		const pt = edges[side];
		if (pt === void 0) continue;
		const child = doc.createElementNS(w$15, "w:" + side);
		child.setAttributeNS(w$15, "w:w", String(Math.round(pt * 20)));
		child.setAttributeNS(w$15, "w:type", "dxa");
		el.appendChild(child);
	}
	return el;
}
function buildTblGrid(cols, effectiveCols, autoShareTwips, ownerDoc) {
	const grid = ownerDoc.createElementNS(w$15, "w:tblGrid");
	if (cols !== void 0) for (const c of cols) grid.appendChild(buildGridCol(c.width, autoShareTwips, ownerDoc));
	else for (let i = 0; i < effectiveCols; i++) grid.appendChild(buildGridCol("auto", autoShareTwips, ownerDoc));
	return grid;
}
function buildGridCol(width, autoTwips, ownerDoc) {
	const col = ownerDoc.createElementNS(w$15, "w:gridCol");
	const twips = width === "auto" ? autoTwips : toTwips(width, "table.cols.width");
	col.setAttributeNS(w$15, "w:w", String(twips));
	return col;
}
function buildTc(slot, block, ownerDoc, ctx) {
	const tc = ownerDoc.createElementNS(w$15, "w:tc");
	const tcPr = buildTcPr(slot, block, ownerDoc);
	tc.appendChild(tcPr);
	if (slot.kind === "vmerge-continue") {
		tc.appendChild(ownerDoc.createElementNS(w$15, "w:p"));
		return tc;
	}
	const blocks = emitCellContent(slot.cell, block, slot.isHeaderRow, ownerDoc, ctx);
	if (blocks.length === 0) tc.appendChild(ownerDoc.createElementNS(w$15, "w:p"));
	else {
		for (const el of blocks) tc.appendChild(el);
		const last = blocks[blocks.length - 1];
		if (last.namespaceURI === w$15 && last.localName === "tbl") tc.appendChild(ownerDoc.createElementNS(w$15, "w:p"));
	}
	return tc;
}
function buildTcPr(slot, block, ownerDoc) {
	const tcPr = ownerDoc.createElementNS(w$15, "w:tcPr");
	if (slot.colspan > 1) {
		const gridSpan = ownerDoc.createElementNS(w$15, "w:gridSpan");
		gridSpan.setAttributeNS(w$15, "w:val", String(slot.colspan));
		insertChildInOrder(tcPr, gridSpan, TC_PR_CHILD_ORDER);
	}
	if (slot.kind === "vmerge-continue") insertChildInOrder(tcPr, ownerDoc.createElementNS(w$15, "w:vMerge"), TC_PR_CHILD_ORDER);
	else if (slot.rowspan > 1) {
		const vMerge = ownerDoc.createElementNS(w$15, "w:vMerge");
		vMerge.setAttributeNS(w$15, "w:val", "restart");
		insertChildInOrder(tcPr, vMerge, TC_PR_CHILD_ORDER);
	}
	const cellBorders = slot.cell.borders;
	const presetHeaderBottom = computePresetHeaderBottom(slot, block);
	if (cellBorders !== void 0 || presetHeaderBottom !== void 0) {
		const merged = { ...cellBorders ?? {} };
		if (presetHeaderBottom !== void 0 && merged.bottom === void 0) merged.bottom = presetHeaderBottom;
		insertChildInOrder(tcPr, buildTcBorders(merged, ownerDoc), TC_PR_CHILD_ORDER);
	}
	if (slot.cell.shading) {
		const shd = ownerDoc.createElementNS(w$15, "w:shd");
		shd.setAttributeNS(w$15, "w:val", "clear");
		shd.setAttributeNS(w$15, "w:color", "auto");
		shd.setAttributeNS(w$15, "w:fill", slot.cell.shading);
		insertChildInOrder(tcPr, shd, TC_PR_CHILD_ORDER);
	}
	if (slot.cell.padding !== void 0 && slot.kind !== "vmerge-continue") insertChildInOrder(tcPr, buildCellMar("tcMar", parsePadding(slot.cell.padding, "cell.padding"), ownerDoc), TC_PR_CHILD_ORDER);
	if (slot.cell.textDirection && slot.kind !== "vmerge-continue") {
		const td = ownerDoc.createElementNS(w$15, "w:textDirection");
		td.setAttributeNS(w$15, "w:val", slot.cell.textDirection);
		insertChildInOrder(tcPr, td, TC_PR_CHILD_ORDER);
	}
	if (slot.kind !== "vmerge-continue") {
		const resolvedVAlign = slot.cell.vAlign ?? block.vAlign ?? "center";
		const vAlign = ownerDoc.createElementNS(w$15, "w:vAlign");
		vAlign.setAttributeNS(w$15, "w:val", resolvedVAlign);
		insertChildInOrder(tcPr, vAlign, TC_PR_CHILD_ORDER);
	}
	return tcPr;
}
/** For the `"three-line"` preset, inject a thin bottom border on each cell
* of the LAST header row — that's the "middle" line of the three-line look.
* Returns undefined when not applicable. */
function computePresetHeaderBottom(slot, block) {
	if (block.borders !== "three-line") return void 0;
	const headerRows = block.headerRows ?? 0;
	if (headerRows === 0) return void 0;
	if (!slot.isHeaderRow) return void 0;
	if (slot.rowIndex !== headerRows - 1) return void 0;
	return {
		style: "single",
		size: .5,
		color: "auto"
	};
}
/** Emit cell content as an array of block-level elements (paragraphs /
* images / breaks). Handles all four cell-content forms. Empty content
* returns `[]`; caller injects empty `<w:p/>`. */
function emitCellContent(cell, block, isHeaderRow, ownerDoc, ctx) {
	if (cell.formula !== void 0) {
		const instrCode = cell.formula.trim();
		if (!instrCode.startsWith("=")) throw new Error(`table cell.formula must start with "=" (got ${JSON.stringify(cell.formula)}). Examples: "=SUM(ABOVE)", "=AVERAGE(LEFT)", "=B2*C2".`);
		const cached = typeof cell.content === "string" ? cell.content : "";
		const headerStyle = isHeaderRow ? block.headerStyle : void 0;
		const p = ownerDoc.createElementNS(w$15, "w:p");
		if (headerStyle) applyDefaultPStyle(p, headerStyle, ownerDoc);
		const field = emitComplexField(ownerDoc, { instrCode, initialResult: cached });
		for (const r of field.runs) p.appendChild(r);
		return [p];
	}
	const content = cell.content;
	const headerStyle = isHeaderRow ? block.headerStyle : void 0;
	if (typeof content === "string") return [emitParagraphFromText(content, headerStyle, ownerDoc, ctx)];
	if (!Array.isArray(content) || content.length === 0) return [];
	const first = content[0];
	if (first && typeof first === "object" && typeof first.type === "string") {
		const out = [];
		for (const b of content) {
			const el = emitBlock(b, ownerDoc, ctx);
			if (headerStyle && el.namespaceURI === w$15 && el.localName === "p" && !hasPStyle(el)) applyDefaultPStyle(el, headerStyle, ownerDoc);
			out.push(el);
		}
		return out;
	}
	return [emitParagraphFromRichText(content, headerStyle, ownerDoc, ctx)];
}
function emitParagraphFromText(text, headerStyle, ownerDoc, ctx) {
	const p = ownerDoc.createElementNS(w$15, "w:p");
	if (headerStyle) applyDefaultPStyle(p, headerStyle, ownerDoc);
	for (const r of emitRichText(text, ownerDoc, ctx, void 0)) p.appendChild(r);
	return p;
}
function emitParagraphFromRichText(inline, headerStyle, ownerDoc, ctx) {
	const p = ownerDoc.createElementNS(w$15, "w:p");
	if (headerStyle) applyDefaultPStyle(p, headerStyle, ownerDoc);
	for (const r of emitRichText(inline, ownerDoc, ctx, void 0)) p.appendChild(r);
	return p;
}
function hasPStyle(p) {
	const pPr = firstChildNS(p, w$15, "pPr");
	if (!pPr) return false;
	return firstChildNS(pPr, w$15, "pStyle") !== null;
}
function applyDefaultPStyle(p, styleId, ownerDoc) {
	let pPr = firstChildNS(p, w$15, "pPr");
	if (!pPr) {
		pPr = ownerDoc.createElementNS(w$15, "w:pPr");
		p.insertBefore(pPr, p.firstChild);
	}
	if (firstChildNS(pPr, w$15, "pStyle")) return;
	const pStyle = ownerDoc.createElementNS(w$15, "w:pStyle");
	pStyle.setAttributeNS(w$15, "w:val", styleId);
	pPr.insertBefore(pStyle, pPr.firstChild);
}
/**
* Walk `container`'s direct children and fix two Word-mandatory invariants:
*
*   1. Adjacent `<w:tbl>` siblings get an empty `<w:p/>` injected between.
*   2. If the container's last meaningful child is `<w:tbl>`, append `<w:p/>`.
*      "Meaningful" excludes a trailing `<w:sectPr>` (body level); a sectPr
*      after a `<w:tbl>` still violates the rule (Word wants `<w:p>`
*      between `<w:tbl>` and `<w:sectPr>`), so it's treated as needing a
*      separator.
*
* Both rules are idempotent — repeat calls are no-ops once normalized.
* Call after inserting any fragment that may include TableBlock output,
* for both body and cell containers.
*/
function normalizeTableSequencing(container, ownerDoc) {
	const isTbl = (el) => el.namespaceURI === w$15 && el.localName === "tbl";
	const isSectPr = (el) => el.namespaceURI === w$15 && el.localName === "sectPr";
	const newP = () => ownerDoc.createElementNS(w$15, "w:p");
	const children = getChildren(container);
	for (let i = 0; i < children.length - 1; i++) {
		const cur = children[i];
		const next = children[i + 1];
		if (isTbl(cur) && (isTbl(next) || isSectPr(next))) container.insertBefore(newP(), next);
	}
	const last = container.lastElementChild;
	if (last && isTbl(last)) container.appendChild(newP());
}

//#endregion
//#region lib/edit/math/mml-to-omml/constants.ts
var import_lib = require_lib();
/**
* Shared constants for the MathML → OMML converter.
*
* Sources:
*   - n-ary chr table: ECMA-376 Part 1 §22.1.2.70 (CT_NaryPr/chr) plus
*     the operator class assignments in MathML Core Appendix B (Operator
*     Dictionary). Cross-checked against TEI Stylesheets' mml2omml.xsl
*     (BSD-2 / CC-BY-SA 3.0, Copyright 2011–2020 TEI Consortium).
*   - Cambria Math font: Word's hard-coded math font; every <m:r> wears
*     it regardless of declared family. Matches Word's own export.
*/
const MML_NS = "http://www.w3.org/1998/Math/MathML";
const M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
/** Operators that promote a surrounding <mrow> into <m:nary>. The string
*  value is what we emit for `<m:chr m:val="…">`; if the symbol matches
*  Word's default for that nary (∑/∏/∫), Word lets us omit chr — but we
*  always emit it for explicitness (matches Word's own round-tripping).
*  Keys are Unicode codepoints as decimal numbers for fast lookup. */
const NARY_OPERATORS = new Map([
	["∑", "∑"],
	["∏", "∏"],
	["∐", "∐"],
	["∫", "∫"],
	["∬", "∬"],
	["∭", "∭"],
	["∮", "∮"],
	["∯", "∯"],
	["∰", "∰"],
	["⨀", "⨀"],
	["⨁", "⨁"],
	["⨂", "⨂"],
	["⋃", "⋃"],
	["⋂", "⋂"],
	["⋁", "⋁"],
	["⋀", "⋀"]
]);
/** Infix binary operators that terminate an n-ary's operand scan. The
*  scan inside an mrow grabs siblings following the n-ary until one of
*  these (or another n-ary, or end of mrow) is hit. Excludes parens/
*  brackets — those bind tighter and belong inside the operand. */
const NARY_OPERAND_TERMINATORS = new Set([
	"+",
	"−",
	"-",
	"=",
	"≠",
	"<",
	">",
	"≤",
	"≥",
	"≈",
	"≡",
	"∼",
	"∝",
	"∈",
	"∉",
	"⊆",
	"⊂",
	"⊇",
	"⊃",
	"⊕",
	"⊖",
	"⊗",
	"⊘",
	"→",
	"←",
	"↔",
	"⇒",
	"⇐",
	"⇔",
	"↦",
	"∧",
	"∨",
	"∪",
	"∩",
	"⊢",
	"⊨",
	"±",
	"∓"
]);
/** Characters that, when wrapped by `<mover>` / `<munder>` with the
*  `accent="true"` attribute (or implied by the character class), map
*  to OMML `<m:acc>` with the given chr value. The chr attribute drops
*  when the accent is the default (combining acute = U+0301). */
const ACCENT_CHARS = new Set([
	"ˇ",
	"˘",
	"˙",
	"˚",
	"˜",
	"^",
	"`",
	"´",
	"¯",
	"⃐",
	"⃑",
	"⃖",
	"⃗",
	"⃛",
	"⃜",
	"⃡"
]);
/** Characters that as the `<mover>` child mean an over-bar (m:bar).
*  Bracket-class chars (⏞ ⏜ ⏝ ⏟) use m:groupChr instead — handled in
*  GROUP_CHR_MAP below. */
const BAR_OVER_CHARS = new Set([
	"¯",
	"̄",
	"‾"
]);
/** Characters that as the `<munder>` child mean an under-bar
*  (m:bar pos="bot"). Symmetric with BAR_OVER_CHARS. */
const BAR_UNDER_CHARS = new Set([
	"_",
	"̲",
	"‗"
]);
/** Bracket-class over/under characters that map to OMML <m:groupChr>
*  (a grouping bracket sized to the base). Value is the position the
*  bracket is rendered on. */
const GROUP_CHR_MAP = new Map([
	["⏞", "top"],
	["⏟", "bot"],
	["⏜", "top"],
	["⏝", "bot"],
	["⎴", "top"],
	["⎵", "bot"],
	["︷", "top"],
	["︸", "bot"]
]);
/** mathvariant values → OMML `<m:sty m:val="…">`. Unknown variants drop
*  back to default (Word's italic for length-1 mi, plain for the rest). */
const MATHVARIANT_STYLE = new Map([
	["normal", "p"],
	["bold", "b"],
	["italic", "i"],
	["bold-italic", "bi"]
]);

//#endregion
//#region lib/edit/math/mml-to-omml/dom.ts
/**
* DOM helpers narrowed for MathML tree walking. Skips text/comment nodes
* (MathML is element-structural — whitespace text between elements is
* not semantic and Word's OMML reader rejects extra text nodes at
* structural levels like inside <m:nary>).
*/
/** MathML grouping wrappers that carry no semantic when they hold a
*  single child — the wrapper is purely for layout grouping. Flatten
*  before fusion passes so n-ary / fence detectors see operator and
*  operand as siblings. mpadded/mstyle attributes are deliberately
*  lost; see SCOPE.md "What OMML cannot express". */
const TRANSPARENT_WRAPPERS = new Set([
	"mrow",
	"mstyle",
	"mpadded",
	"maction"
]);
/** Recursively unwrap single-child transparent wrappers from a list of
*  elements. The output preserves the order of meaningful siblings;
*  intermediate single-child mrows disappear. */
function flattenSingleChildWrappers(kids) {
	const out = [];
	for (const k of kids) if (TRANSPARENT_WRAPPERS.has(k.localName) && elementChildren(k).length === 1) out.push(...flattenSingleChildWrappers(elementChildren(k)));
	else out.push(k);
	return out;
}
function elementChildren(parent) {
	const out = [];
	for (let n = parent.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) out.push(n);
	return out;
}
function isMmlElement(n, localName) {
	if (!n || n.nodeType !== 1) return false;
	const el = n;
	if (el.namespaceURI !== "http://www.w3.org/1998/Math/MathML") return false;
	return localName === void 0 || el.localName === localName;
}
/** Concatenated text content of an Element, MathML-flavored — trims
*  surrounding whitespace because MathML leaf text is conventionally
*  exactly the symbol/identifier with no padding, but some serializers
*  emit a stray newline inside <mo>+</mo>.
*
*  Does NOT use this for `<mtext>` / `<ms>` — those preserve whitespace
*  per MathML 3 §3.2.6. Use `mmlTextLiteral` instead. */
function mmlText(el) {
	return (el.textContent ?? "").trim();
}
/** Whitespace-preserving text extraction for `<mtext>` / `<ms>` — the
*  surrounding spaces of `\text{ if and only if }` carry typographic
*  intent and must reach <m:t> verbatim. */
function mmlTextLiteral(el) {
	return el.textContent ?? "";
}
/** Single named attribute, undefined when absent. Works for MathML
*  attributes which are typically unnamespaced. */
function attr(el, name) {
	return el.hasAttribute(name) ? el.getAttribute(name) : void 0;
}
/** Create an element in the OMML math namespace using the `m:` prefix
*  so xmldom emits readable `<m:nary>` instead of `<nary xmlns="…">`. */
function mEl(doc, localName) {
	return doc.createElementNS(M_NS, `m:${localName}`);
}
/** Create an element in the wordprocessingml namespace, used only for
*  the `<w:rPr><w:rFonts/></w:rPr>` font-selection block that lives
*  inside every OMML run. */
function wEl(doc, localName) {
	return doc.createElementNS(W_NS, `w:${localName}`);
}
/** Set a `m:val` attribute on a "property" element (e.g. `<m:chr m:val="∑"/>`).
*  All OMML property attributes use this single shape. */
function setMVal(el, value) {
	el.setAttributeNS(M_NS, "m:val", value);
}

//#endregion
//#region lib/edit/math/mml-to-omml/style.ts
/**
* Math character style resolution. Determines whether to emit
* `<m:sty m:val="…"/>` inside a run, based on:
*   - the leaf kind (mi/mn/mo/mtext/ms)
*   - the text (mi length-1 vs length>1 drives MathML's italic default)
*   - an explicit `mathvariant` attribute on the source element
*
* Returns the OMML sty value or `undefined` (use Word's default).
*
* MathML default conventions (from MathML 3 §3.2.3):
*   - mi length 1 → italic
*   - mi length >1 → normal/upright (function-name convention)
*   - mn, mo, mtext, ms → normal
*
* Word's defaults in Cambria Math math context:
*   - Plain text in <m:r> renders italic for single-codepoint Latin
*     letters, plain for digits/operators/multi-char text
*   - <m:sty m:val="p"/> forces plain; "i" forces italic; "b" bold;
*     "bi" bold-italic
*
* So emit sty ONLY when our intent diverges from Word's leaf default:
*   - mi with length>1: emit "p" (Word would render italic on Latin chars)
*   - explicit mathvariant: emit corresponding mapping
*/
function resolveStyle(kind, text, mathvariant) {
	if (mathvariant !== void 0) {
		const explicit = MATHVARIANT_STYLE.get(mathvariant);
		if (explicit !== void 0) return explicit;
		return;
	}
	if (kind === "mi") return codepointLength(text) > 1 ? "p" : void 0;
}
/** Codepoint count, not UTF-16 unit count — keeps surrogate pairs like
*  ℒ (U+2112) and 𝐀 (U+1D400) counted as one symbol. */
function codepointLength(s) {
	let n = 0;
	for (const _ of s) n++;
	return n;
}

//#endregion
//#region lib/edit/math/mml-to-omml/run.ts
const CAMBRIA = "Cambria Math";
/** Build a single <m:r> with the given text and styling. Empty text
*  still emits a run so adjacent runs don't merge in serialization. */
function buildRun(doc, text, kind, mathvariant) {
	const r = mEl(doc, "r");
	const sty = resolveStyle(kind, text, mathvariant);
	const isLiteralText = kind === "mtext" || kind === "ms";
	if (sty !== void 0 || isLiteralText) {
		const mRPr = mEl(doc, "rPr");
		if (isLiteralText) mRPr.appendChild(mEl(doc, "nor"));
		if (sty !== void 0) {
			const styEl = mEl(doc, "sty");
			setMVal(styEl, sty);
			mRPr.appendChild(styEl);
		}
		r.appendChild(mRPr);
	}
	const wRPr = wEl(doc, "rPr");
	const rFonts = wEl(doc, "rFonts");
	rFonts.setAttributeNS(W_NS, "w:ascii", CAMBRIA);
	rFonts.setAttributeNS(W_NS, "w:eastAsia", CAMBRIA);
	rFonts.setAttributeNS(W_NS, "w:hAnsi", CAMBRIA);
	rFonts.setAttributeNS(W_NS, "w:cs", CAMBRIA);
	wRPr.appendChild(rFonts);
	r.appendChild(wRPr);
	const t = mEl(doc, "t");
	if (text.length > 0 && (text.startsWith(" ") || text.endsWith(" ") || isLiteralText)) t.setAttribute("xml:space", "preserve");
	t.textContent = text;
	r.appendChild(t);
	return r;
}

//#endregion
//#region lib/edit/math/mml-to-omml/nary.ts
/**
* N-ary operator detection and fusion.
*
* Why this exists: temml emits a sum/integral/product as
*
*   <mrow>
*     <munderover><mo>∑</mo><mrow>i=1</mrow><mi>n</mi></munderover>
*     <msup><mi>i</mi><mn>2</mn></msup>     <-- the operand
*   </mrow>
*
* The operator-with-limits and the operand are *sibling* children of an
* mrow. Node-by-node converters (mml2omml LGPL, TEI XSLT) translate each
* sibling independently and emit `<m:nary>` with an empty `<m:e/>` and
* the operand as a sibling outside the nary — Word renders that as a
* dashed empty box before the actual operand.
*
* Fix: at every mrow scope, scan for an n-ary-bearing element (`mo`,
* `munder`/`mover`/`munderover` or `msub`/`msup`/`msubsup` whose base is
* `mo` in the n-ary chr set) and fuse it with the following siblings as
* operand, up to an infix-binary stop or end of mrow.
*
* Returns either null (the element at index `i` is not n-ary) or a
* NaryMatch describing how many siblings to consume and what their
* roles are. The walker then emits one `<m:nary>` and skips the consumed
* indices.
*/
/** Examine sibling list `siblings` starting at `i`. If `siblings[i]` is
*  an n-ary head, fuse with following siblings into a NaryMatch.
*  Returns null when no n-ary at that position. */
function detectNary(siblings, i) {
	const head = siblings[i];
	if (!head) return null;
	let chr;
	let sub = null;
	let sup = null;
	let limitsAboveBelow = false;
	if (isMmlElement(head, "mo")) {
		chr = NARY_OPERATORS.get(mmlText(head));
		limitsAboveBelow = true;
	} else if (isMmlElement(head, "munder") || isMmlElement(head, "mover") || isMmlElement(head, "munderover")) {
		const kids = elementChildren(head);
		const base = kids[0];
		if (base && isMmlElement(base, "mo")) {
			chr = NARY_OPERATORS.get(mmlText(base));
			limitsAboveBelow = true;
			if (isMmlElement(head, "munder")) sub = [kids[1]];
			else if (isMmlElement(head, "mover")) sup = [kids[1]];
			else {
				sub = [kids[1]];
				sup = [kids[2]];
			}
		}
	} else if (isMmlElement(head, "msub") || isMmlElement(head, "msup") || isMmlElement(head, "msubsup")) {
		const kids = elementChildren(head);
		const base = kids[0];
		if (base && isMmlElement(base, "mo")) {
			chr = NARY_OPERATORS.get(mmlText(base));
			limitsAboveBelow = false;
			if (isMmlElement(head, "msub")) sub = [kids[1]];
			else if (isMmlElement(head, "msup")) sup = [kids[1]];
			else {
				sub = [kids[1]];
				sup = [kids[2]];
			}
		}
	}
	if (chr === void 0) return null;
	const operand = [];
	let j = i + 1;
	while (j < siblings.length) {
		const next = siblings[j];
		if (isMmlElement(next, "mo") && NARY_OPERAND_TERMINATORS.has(mmlText(next))) break;
		operand.push(next);
		j++;
	}
	return {
		chr,
		sub,
		sup,
		operand,
		consumed: j - i,
		limitsAboveBelow
	};
}

//#endregion
//#region lib/edit/math/mml-to-omml/walk.ts
/**
* Main MathML → OMML walker.
*
* `emitChildren` is the recursion point: it iterates over an mrow-like
* scope, applies n-ary fusion lookahead, then dispatches each remaining
* element through `emitElement`. Per-element emitters call back into
* `emitChildren` for their own sub-scopes.
*
* Element mappings ported from TEI Stylesheets' mml2omml.xsl (BSD-2 /
* CC-BY-SA, © 2011–2020 TEI Consortium). N-ary fusion and the
* transparent passthrough for mpadded/mstyle are net-new — neither TEI
* nor mml2omml gets those right.
*/
/** Emit a list of MathML elements into an OMML container, applying
*  single-child wrapper flattening and n-ary fusion. Single-child
*  `<mrow>` / `<mpadded>` / `<mstyle>` / `<maction>` are unwrapped
*  first — temml wraps every operator-with-limits in such a grouping
*  mrow that would otherwise hide the fusion candidate from its
*  sibling operand. (menclose is deliberately NOT in the flatten set:
*  its `notation` attribute carries semantic that must reach
*  `emitMenclose`.) */
function emitSequence(items, host, doc) {
	const kids = flattenSingleChildWrappers(items);
	for (let i = 0; i < kids.length; i++) {
		const nary = detectNary(kids, i);
		if (nary !== null) {
			emitNary(nary, host, doc);
			i += nary.consumed - 1;
			continue;
		}
		emitElement(kids[i], host, doc);
	}
}
/** Emit `mmlParent`'s element children into `ommlParent`. Convenience
*  wrapper around emitSequence for the common "walk this element's
*  children" pattern. */
function emitChildren(mmlParent, ommlParent, doc) {
	emitSequence(elementChildren(mmlParent), ommlParent, doc);
}
function emitNary(m, parent, doc) {
	const nary = mEl(doc, "nary");
	const naryPr = mEl(doc, "naryPr");
	const chr = mEl(doc, "chr");
	setMVal(chr, m.chr);
	naryPr.appendChild(chr);
	const limLoc = mEl(doc, "limLoc");
	setMVal(limLoc, m.limitsAboveBelow ? "undOvr" : "subSup");
	naryPr.appendChild(limLoc);
	if (m.sub === null) {
		const subHide = mEl(doc, "subHide");
		setMVal(subHide, "1");
		naryPr.appendChild(subHide);
	}
	if (m.sup === null) {
		const supHide = mEl(doc, "supHide");
		setMVal(supHide, "1");
		naryPr.appendChild(supHide);
	}
	nary.appendChild(naryPr);
	const sub = mEl(doc, "sub");
	if (m.sub !== null) emitSequence(m.sub, sub, doc);
	nary.appendChild(sub);
	const sup = mEl(doc, "sup");
	if (m.sup !== null) emitSequence(m.sup, sup, doc);
	nary.appendChild(sup);
	const e = mEl(doc, "e");
	emitSequence(m.operand, e, doc);
	nary.appendChild(e);
	parent.appendChild(nary);
}
/** Dispatch one MathML element to its OMML emitter and append into
*  `parent`. */
function emitElement(el, parent, doc) {
	const ln = el.localName;
	switch (ln) {
		case "mi":
		case "mn":
		case "mo":
		case "mtext":
		case "ms":
			parent.appendChild(emitLeaf(el, ln, doc));
			return;
		case "mspace":
			parent.appendChild(buildRun(doc, " ", "mspace", void 0));
			return;
		case "mrow":
			if (isFenceMrow(el)) {
				parent.appendChild(emitStretchyFence(el, doc));
				return;
			}
			emitChildren(el, parent, doc);
			return;
		case "mpadded":
		case "mstyle":
		case "maction":
			emitChildren(el, parent, doc);
			return;
		case "menclose":
			parent.appendChild(emitMenclose(el, doc));
			return;
		case "mphantom": {
			const phant = mEl(doc, "phant");
			const phantE = mEl(doc, "e");
			emitChildren(el, phantE, doc);
			phant.appendChild(phantE);
			parent.appendChild(phant);
			return;
		}
		case "msub":
			parent.appendChild(emitSubSup(el, doc, "sub"));
			return;
		case "msup":
			parent.appendChild(emitSubSup(el, doc, "sup"));
			return;
		case "msubsup":
			parent.appendChild(emitSubSup(el, doc, "both"));
			return;
		case "munder":
			parent.appendChild(emitUnderOver(el, doc, "under"));
			return;
		case "mover":
			parent.appendChild(emitUnderOver(el, doc, "over"));
			return;
		case "munderover":
			parent.appendChild(emitUnderOver(el, doc, "both"));
			return;
		case "mfrac":
			parent.appendChild(emitFrac(el, doc));
			return;
		case "msqrt":
			parent.appendChild(emitSqrt(el, doc));
			return;
		case "mroot":
			parent.appendChild(emitRoot(el, doc));
			return;
		case "mfenced":
			parent.appendChild(emitFenced(el, doc));
			return;
		case "mtable":
			parent.appendChild(emitTable(el, doc));
			return;
		case "merror": throw new Error(`MathML <merror> in input — temml hit an error rendering this LaTeX. Inspect the operand: ${(el.textContent ?? "").slice(0, 120)}. Fix the LaTeX, or switch this equation to the omml escape hatch on the EquationBlock.`);
		case "semantics":
			{
				const first = elementChildren(el)[0];
				if (first) emitElement(first, parent, doc);
			}
			return;
		case "annotation":
		case "annotation-xml": return;
		default: throw new Error(`Unsupported MathML element <${el.localName}> in MathML→OMML conversion. Switch this equation to the omml escape hatch on the EquationBlock; see references/equations.md "Known fragile LaTeX tokens".`);
	}
}
function emitLeaf(el, kind, doc) {
	return buildRun(doc, kind === "mtext" || kind === "ms" ? mmlTextLiteral(el) : mmlText(el), kind, attr(el, "mathvariant"));
}
/** Append a new `<m:NAME>` wrapper containing `source` (emitted) to
*  `parent`. Compresses the repeated three-line "create element,
*  recurse one source into it, append" pattern. */
function appendWrapped(parent, name, source, doc) {
	const el = mEl(doc, name);
	emitElement(source, el, doc);
	parent.appendChild(el);
}
/** msub / msup / msubsup → <m:sSub> / <m:sSup> / <m:sSubSup>. */
function emitSubSup(el, doc, which) {
	const kids = elementChildren(el);
	const out = mEl(doc, which === "sub" ? "sSub" : which === "sup" ? "sSup" : "sSubSup");
	appendWrapped(out, "e", kids[0], doc);
	if (which === "sub" || which === "both") appendWrapped(out, "sub", kids[1], doc);
	if (which === "sup" || which === "both") appendWrapped(out, "sup", kids[which === "both" ? 2 : 1], doc);
	return out;
}
/** munder / mover / munderover → <m:limLow> / <m:limUpp> / <m:limLowUpp>
*  when the over/under is an accent character → <m:acc>;
*  when over is bar-class → <m:bar>;
*  otherwise the limit/limit-low-upp shapes. */
function emitUnderOver(el, doc, which) {
	const kids = elementChildren(el);
	if (which === "over" || which === "under") {
		const decorator = kids[1];
		const decText = mmlText(decorator);
		const pos = GROUP_CHR_MAP.get(decText);
		if (pos !== void 0 && (which === "over" ? pos === "top" : pos === "bot")) return emitGroupChr(kids[0], decText, pos, doc);
	}
	if (which === "over") {
		const over = kids[1];
		const overText = mmlText(over);
		if (BAR_OVER_CHARS.has(overText)) {
			const bar = mEl(doc, "bar");
			const barPr = mEl(doc, "barPr");
			const pos = mEl(doc, "pos");
			setMVal(pos, "top");
			barPr.appendChild(pos);
			bar.appendChild(barPr);
			const e = mEl(doc, "e");
			emitElement(kids[0], e, doc);
			bar.appendChild(e);
			return bar;
		}
		const explicit = attr(el, "accent") === "true";
		const isSingleCharMo = isMmlElement(over, "mo") && [...overText].length === 1;
		if (explicit || ACCENT_CHARS.has(overText) || isSingleCharMo) {
			const acc = mEl(doc, "acc");
			const accPr = mEl(doc, "accPr");
			const chr = mEl(doc, "chr");
			setMVal(chr, overText);
			accPr.appendChild(chr);
			acc.appendChild(accPr);
			const e = mEl(doc, "e");
			emitElement(kids[0], e, doc);
			acc.appendChild(e);
			return acc;
		}
	}
	if (which === "under") {
		const under = kids[1];
		if (BAR_UNDER_CHARS.has(mmlText(under))) {
			const bar = mEl(doc, "bar");
			const barPr = mEl(doc, "barPr");
			const pos = mEl(doc, "pos");
			setMVal(pos, "bot");
			barPr.appendChild(pos);
			bar.appendChild(barPr);
			const e = mEl(doc, "e");
			emitElement(kids[0], e, doc);
			bar.appendChild(e);
			return bar;
		}
	}
	if (which === "both") {
		const outer = mEl(doc, "limUpp");
		const innerE = mEl(doc, "e");
		const inner = mEl(doc, "limLow");
		appendWrapped(inner, "e", kids[0], doc);
		appendWrapped(inner, "lim", kids[1], doc);
		innerE.appendChild(inner);
		outer.appendChild(innerE);
		appendWrapped(outer, "lim", kids[2], doc);
		return outer;
	}
	const out = mEl(doc, which === "under" ? "limLow" : "limUpp");
	appendWrapped(out, "e", kids[0], doc);
	appendWrapped(out, "lim", kids[1], doc);
	return out;
}
function emitMenclose(el, doc) {
	const notation = (attr(el, "notation") ?? "longdiv").trim().toLowerCase();
	const flags = new Set(notation.split(/\s+/));
	if ([...flags].every((f) => f === "box" || f === "roundedbox")) {
		const box = mEl(doc, "borderBox");
		const e = mEl(doc, "e");
		emitChildren(el, e, doc);
		box.appendChild(e);
		return box;
	}
	if (flags.size === 1 && (flags.has("top") || flags.has("bottom"))) {
		const bar = mEl(doc, "bar");
		const barPr = mEl(doc, "barPr");
		const pos = mEl(doc, "pos");
		setMVal(pos, flags.has("top") ? "top" : "bot");
		barPr.appendChild(pos);
		bar.appendChild(barPr);
		const e = mEl(doc, "e");
		emitChildren(el, e, doc);
		bar.appendChild(e);
		return bar;
	}
	throw new Error(`MathML <menclose notation="${notation}"> has no OMML equivalent (only "box", "roundedbox", "top", "bottom" map cleanly). Switch this equation to the omml escape hatch on the EquationBlock; see references/equations.md "Known fragile LaTeX tokens".`);
}
function emitGroupChr(base, chr, pos, doc) {
	const g = mEl(doc, "groupChr");
	const gPr = mEl(doc, "groupChrPr");
	const chrEl = mEl(doc, "chr");
	setMVal(chrEl, chr);
	gPr.appendChild(chrEl);
	const posEl = mEl(doc, "pos");
	setMVal(posEl, pos);
	gPr.appendChild(posEl);
	const vertJc = mEl(doc, "vertJc");
	setMVal(vertJc, pos);
	gPr.appendChild(vertJc);
	g.appendChild(gPr);
	const e = mEl(doc, "e");
	emitElement(base, e, doc);
	g.appendChild(e);
	return g;
}
function emitFrac(el, doc) {
	const kids = elementChildren(el);
	const f = mEl(doc, "f");
	const fPr = mEl(doc, "fPr");
	const linethickness = attr(el, "linethickness");
	if (linethickness === "0" || linethickness === "0pt") {
		const type = mEl(doc, "type");
		setMVal(type, "noBar");
		fPr.appendChild(type);
		f.appendChild(fPr);
	}
	const num = mEl(doc, "num");
	emitElement(kids[0], num, doc);
	f.appendChild(num);
	const den = mEl(doc, "den");
	emitElement(kids[1], den, doc);
	f.appendChild(den);
	return f;
}
function emitSqrt(el, doc) {
	const rad = mEl(doc, "rad");
	const radPr = mEl(doc, "radPr");
	const degHide = mEl(doc, "degHide");
	setMVal(degHide, "1");
	radPr.appendChild(degHide);
	rad.appendChild(radPr);
	rad.appendChild(mEl(doc, "deg"));
	const e = mEl(doc, "e");
	emitChildren(el, e, doc);
	rad.appendChild(e);
	return rad;
}
function emitRoot(el, doc) {
	const kids = elementChildren(el);
	const rad = mEl(doc, "rad");
	const deg = mEl(doc, "deg");
	emitElement(kids[1], deg, doc);
	rad.appendChild(deg);
	const e = mEl(doc, "e");
	emitElement(kids[0], e, doc);
	rad.appendChild(e);
	return rad;
}
function emitFenced(el, doc) {
	const open = attr(el, "open") ?? "(";
	const close = attr(el, "close") ?? ")";
	const separators = attr(el, "separators") ?? ",";
	const d = mEl(doc, "d");
	const dPr = mEl(doc, "dPr");
	const begChr = mEl(doc, "begChr");
	setMVal(begChr, open);
	dPr.appendChild(begChr);
	if (separators !== ",") {
		const sepChr = mEl(doc, "sepChr");
		setMVal(sepChr, separators.charAt(0));
		dPr.appendChild(sepChr);
	}
	const endChr = mEl(doc, "endChr");
	setMVal(endChr, close);
	dPr.appendChild(endChr);
	d.appendChild(dPr);
	for (const child of elementChildren(el)) {
		const e = mEl(doc, "e");
		emitElement(child, e, doc);
		d.appendChild(e);
	}
	return d;
}
function isFenceMrow(mrow) {
	const kids = elementChildren(mrow);
	if (kids.length < 2) return false;
	const first = kids[0];
	const last = kids[kids.length - 1];
	return isMmlElement(first, "mo") && attr(first, "fence") === "true" && isMmlElement(last, "mo") && attr(last, "fence") === "true";
}
function emitStretchyFence(mrow, doc) {
	const kids = elementChildren(mrow);
	const open = kids[0];
	const close = kids[kids.length - 1];
	const body = kids.slice(1, -1);
	const d = mEl(doc, "d");
	const dPr = mEl(doc, "dPr");
	const begChr = mEl(doc, "begChr");
	setMVal(begChr, mmlText(open));
	dPr.appendChild(begChr);
	const endChr = mEl(doc, "endChr");
	setMVal(endChr, mmlText(close));
	dPr.appendChild(endChr);
	dPr.appendChild(mEl(doc, "grow"));
	d.appendChild(dPr);
	const e = mEl(doc, "e");
	for (const child of body) emitElement(child, e, doc);
	d.appendChild(e);
	return d;
}
function emitTable(el, doc) {
	const m = mEl(doc, "m");
	for (const tr of elementChildren(el)) {
		if (!isMmlElement(tr, "mtr") && !isMmlElement(tr, "mlabeledtr")) continue;
		const mr = mEl(doc, "mr");
		for (const td of elementChildren(tr)) {
			if (!isMmlElement(td, "mtd")) continue;
			const e = mEl(doc, "e");
			emitChildren(td, e, doc);
			mr.appendChild(e);
		}
		m.appendChild(mr);
	}
	return m;
}

//#endregion
//#region lib/edit/math/mml-to-omml/index.ts
/**
* MathML → OMML converter — in-tree replacement for the LGPL
* `mathml2omml` package. Public surface is one function.
*
*   import { convertMathMLToOMML } from "@lib/edit/math/mml-to-omml/index.ts"
*
* The output is a string suitable for embedding as `<m:oMath>` inside
* a Word `<w:p>` (display) or inline run sequence. The string carries
* `xmlns:m` on its root, so it survives importNode into any document
* without further namespace bookkeeping.
*
* Module layout (all under this directory):
*   constants.ts — namespaces, n-ary chr table, accent + bar chars,
*                  mathvariant → m:sty mapping
*   style.ts     — leaf-kind → m:sty resolution (italic-by-length
*                  default for mi)
*   run.ts       — <m:r> with <w:rPr> + optional <m:rPr> + <m:t>
*   dom.ts       — DOM walking helpers + element factories
*   nary.ts      — n-ary fusion detection (this is the structural fix
*                  vs mml2omml's empty-<m:e/> bug)
*   walk.ts      — element dispatch table; recurses through emitChildren
*
* Element coverage:
*   leaves      mi mn mo mtext ms mspace
*   structural  mrow msub msup msubsup munder mover munderover
*               mfrac msqrt mroot mfenced mtable mtr mtd
*               mphantom (→ m:phant)
*   transparent mpadded mstyle menclose maction (→ children only)
*   discarded   semantics, annotation, annotation-xml
*   accents     mover with accent="true" or accent-class char → m:acc
*               mover/munder with bar-class char → m:bar
*
* Any other element triggers a throw with a pointer to the omml escape
* hatch in references/equations.md. The error contract matches the rest
* of the math pipeline (latex-to-omml.ts).
*/
const serializer = new import_lib.XMLSerializer();
function convertMathMLToOMML(mathmlXml) {
	const root = parseXml(mathmlXml).documentElement;
	if (!root) throw new Error(`MathML → OMML: input has no root element`);
	if (root.namespaceURI !== "http://www.w3.org/1998/Math/MathML" || root.localName !== "math") throw new Error(`MathML → OMML: expected <math xmlns="${MML_NS}"> root, got <${root.localName} xmlns="${root.namespaceURI ?? ""}">`);
	const outDoc = parseXml(`<m:oMath xmlns:m="${M_NS}"/>`);
	const oMath = outDoc.documentElement;
	emitChildren(root, oMath, outDoc);
	return serializer.serializeToString(outDoc);
}

//#endregion
//#region lib/edit/math/latex-to-omml.ts
/**
* LaTeX → OMML pipeline.
*
*   stage 1: Temml renders LaTeX → MathML XML string. MIT-licensed, browser
*            + Node. `xml: true` keeps the renderer DOM-free. Loaded via
*            dynamic import + fallback because tsdown corrupts the
*            surrogate-range literals in temml's tokenizer when bundling;
*            build-skill.ts copies the runtime into `_shared/temml/`. Same
*            pattern as xmllint-wasm.
*
*   stage 2: in-tree MathML → OMML converter (`@lib/edit/math/mml-to-omml`).
*            Synchronous; throws on unsupported MathML elements with a
*            pointer to the `omml` escape hatch.
*
* Synchronization model:
*   The emit chain is synchronous (one call per Block, recursive into
*   tables). Temml loads asynchronously via dynamic import. The engine
*   pre-walks every Block in `edits[]`, calls `prepareLatex` for each
*   `latex` it finds, then emit calls `getOmmlSync` from the populated
*   cache. The cache is module-scoped — keeping it stateful is acceptable
*   for a per-invocation CLI.
*/
const MODULE_DIR$1 = dirname(fileURLToPath(import.meta.url));
let cachedTemml = null;
async function loadTemml() {
	if (cachedTemml) return cachedTemml;
	const candidates = [
		"temml",
		pathToFileURL(join(MODULE_DIR$1, "temml", "dist", "temml.cjs")).href,
		pathToFileURL(join(MODULE_DIR$1, "_shared", "temml", "dist", "temml.cjs")).href
	];
	let lastErr = null;
	for (const spec of candidates) try {
		const mod = await import(spec);
		const resolved = "renderToString" in mod ? mod : mod.default;
		cachedTemml = resolved;
		return resolved;
	} catch (err) {
		lastErr = err;
	}
	throw new Error(`temml runtime not found. Looked for: ${candidates.join(", ")}. Last error: ${lastErr?.message ?? "(unknown)"}.`);
}
function cacheKey(latex, displayMode) {
	return `${displayMode ? "B" : "I"}:${latex}`;
}
const ommlCache = /* @__PURE__ */ new Map();
/** Pre-resolve a batch of LaTeX expressions into OMML strings, populating
* the module cache the synchronous emit functions read from. The engine
* calls this once before running edits; idempotent + safe to call with
* duplicate items. */
async function prepareLatex(items) {
	if (items.length === 0) return;
	const temml = await loadTemml();
	for (const { latex, displayMode } of items) {
		const key = cacheKey(latex, displayMode);
		if (ommlCache.has(key)) continue;
		let mathml;
		try {
			mathml = temml.renderToString(latex, {
				xml: true,
				displayMode,
				throwOnError: true
			});
		} catch (err) {
			throw new Error(`LaTeX parse error in ${truncateLatex(latex)}: ${err.message}`, { cause: err });
		}
		let omml;
		try {
			omml = convertMathMLToOMML(mathml);
		} catch (err) {
			throw new Error(`MathML → OMML conversion failed for ${truncateLatex(latex)}: ${err.message}`, { cause: err });
		}
		const schemaErrors = await validateOMath(omml);
		if (schemaErrors.length > 0) throw new Error(`OOXML math schema validation failed for ${truncateLatex(latex)}: ${schemaErrors[0]}` + (schemaErrors.length > 1 ? ` (+${schemaErrors.length - 1} more)` : "") + ` — switch this equation to the omml escape hatch on the EquationBlock; see references/equations.md "Known fragile LaTeX tokens".`);
		ommlCache.set(key, omml);
	}
}
/** Short, JSON-safe rendering of a LaTeX source for error messages.
* Cap at 80 chars so a long expression doesn't drown the stderr line. */
function truncateLatex(latex) {
	const trimmed = latex.length > 80 ? latex.slice(0, 80) + "…" : latex;
	return JSON.stringify(trimmed);
}
/** Read a pre-resolved OMML string. Throws when the latex wasn't passed
* through prepareLatex first — this is an engine bug (the pre-walk missed
* a code path), not user-actionable. */
function getOmmlSync(latex, displayMode) {
	const key = cacheKey(latex, displayMode);
	const omml = ommlCache.get(key);
	if (!omml) throw new Error(`internal: equation not pre-resolved (latex=${JSON.stringify(latex)}, displayMode=${displayMode}). Engine must call prepareLatex on all edits before emit.`);
	return omml;
}

//#endregion
//#region lib/edit/math/omml-build.ts
/**
* Shared `<m:oMath>` builder.
*
* Resolves a math source (LaTeX via temml, or raw OMML escape hatch)
* into an `<m:oMath>` Element imported into `ownerDoc`. The OMML
* produced by `latexToOmml` already declares `xmlns:m` (and the
* unused-but-harmless `xmlns:w`) on its root — xmldom preserves these
* on import, so the result serializes correctly without touching the
* document root's namespace declarations.
*
* `displayMode` is forwarded to temml; OMML escape-hatch sources ignore
* it (the agent-supplied XML already encodes display vs inline).
*
* Both caption-emit (numbered equation cell content) and equation-emit
* (numbered + unnumbered top-level equations) share this resolver —
* splitting it out avoids a circular import (equation-emit already
* imports `emitNumberedEquation` from caption-emit).
*/
function buildOMath(source, ownerDoc, displayMode) {
	const root = parseXml("latex" in source ? getOmmlSync(source.latex, displayMode) : source.omml).documentElement;
	if (!root) throw new Error(`Math source produced no root element: ${JSON.stringify(source)}`);
	return ownerDoc.importNode(root, true);
}

//#endregion
//#region lib/edit/fields/seq-field.ts
/**
* SEQ field emitter — caption-class auto-numbering.
*
*   { SEQ <identifier> \* <FORMAT> [\s <N>] [\r <N>] [\c] [\h] [\* MERGEFORMAT] }
*
* Each `SEQ <id>` occurrence increments the document-local counter for
* that identifier. Different identifiers track independent counters —
* `SEQ Figure` and `SEQ Equation` don't interfere. This is Word's native
* caption mechanism (References → Insert Caption produces SEQ).
*
* Switches:
*   `\* FORMAT`     numeric format for the rendered counter (ARABIC,
*                   alphabetic, ALPHABETIC, roman, ROMAN, CHINESENUM2,
*                   CHINESENUM3). Always present.
*   `\s N`          restart the counter at each occurrence of a paragraph
*                   at outline level N (1-indexed, matches Word UI's
*                   "Heading 1" = level 1). Paired with a STYLEREF chapter
*                   prefix in the caption's run sequence.
*   `\r N`          reset the counter to N at this field (used by
*                   CaptionCounterReset). Mutually exclusive with `\c`.
*   `\c`            repeat the previous value of this identifier without
*                   incrementing (used by subequation "continue" members
*                   to keep the parent number stable across (1a)(1b)).
*
* Format preservation (rPr replication + MERGEFORMAT) handled by the
* shared skeleton — see `complex-field.ts`.
*
* No `\h` (hide) switch: Word's documented quirk is that `\h` is
* silently overridden by a `\*` format switch in the same field
* (Microsoft Q&A; Office Watch). Since our emitter always emits `\*`,
* `\h` would silently fail. Callers that need a hidden counter-advance
* marker wrap each emitted run in `<w:vanish/>` rPr instead (see
* `addVanishRPr` in `lib/xml/xml-utils.ts`).
*/
/** Mapping from agent-facing format string to Word's `\*` switch token.
* Engine emits these verbatim in instrText. */
const FORMAT_SWITCH = {
	arabic: "ARABIC",
	alphabetic: "alphabetic",
	ALPHABETIC: "ALPHABETIC",
	roman: "roman",
	ROMAN: "ROMAN",
	chinese: "CHINESENUM2",
	"chinese-formal": "CHINESENUM3"
};
/** Emit the 5-run SEQ sequence. Returns the runs plus the result text
* element for counter-sim backfill. */
function emitSeqField(ownerDoc, spec) {
	if (spec.repeat && spec.resetTo !== void 0) throw new Error(`SEQ field "${spec.identifier}": \\c (repeat) and \\r (reset) are mutually exclusive; pass one or the other.`);
	const parts = [`SEQ ${spec.identifier}`, `\\* ${FORMAT_SWITCH[spec.format]}`];
	if (spec.restartAtOutlineLevel !== void 0) parts.push(`\\s ${spec.restartAtOutlineLevel}`);
	if (spec.resetTo !== void 0) parts.push(`\\r ${spec.resetTo}`);
	if (spec.repeat) parts.push("\\c");
	const { runs, resultTextEl } = emitComplexField(ownerDoc, {
		instrCode: parts.join(" "),
		initialResult: spec.initialResult,
		format: spec.formatRPr
	});
	return {
		runs,
		resultTextEl
	};
}

//#endregion
//#region lib/edit/fields/styleref-field.ts
/**
* STYLEREF field emitter — chapter-prefix lookup.
*
*   { STYLEREF "<styleName>" \n [\* MERGEFORMAT] }
*
* Resolves to the rendered content of the nearest preceding paragraph
* with the given style. Paired with SEQ in caption layouts to produce
* chapter-prefixed numbers like "2.3" — STYLEREF "Heading 1" \n returns
* the H1's paragraph number ("2"); SEQ Equation \s 1 returns the
* equation counter within that chapter ("3"); a literal "." run joins
* them.
*
* The `\n` switch (paragraph number) is the canonical caption switch —
* returns just the numbered prefix without the heading text. Other
* switches are available for non-caption uses (`\l` lower formatting,
* `\p` position, etc.) but the caption pipeline only uses `\n`.
*
* styleName: Word's display name (the `<w:name w:val="..."/>` value in
* styles.xml), NOT the styleId. The engine resolves styleId → name at
* resolution time. STYLEREF requires double-quoted style names —
* required for names with spaces ("Heading 1"), harmless for names
* without. The emitter always quotes.
*
* Format preservation (rPr replication + MERGEFORMAT) handled by the
* shared skeleton — see `complex-field.ts`.
*/
/** Emit the 5-run STYLEREF sequence. Returns the runs plus the result
* text element for counter-sim backfill. */
function emitStyleRefField(ownerDoc, spec) {
	const switchPart = spec.switches.length > 0 ? ` ${spec.switches.join(" ")}` : "";
	const { runs, resultTextEl } = emitComplexField(ownerDoc, {
		instrCode: `STYLEREF "${spec.styleName}"${switchPart}`,
		initialResult: spec.initialResult,
		format: spec.formatRPr
	});
	return {
		runs,
		resultTextEl
	};
}

//#endregion
//#region lib/edit/caption-emit.ts
/**
* Caption-class layout emitters.
*
* EquationBlock with captionId → 3-col borderless table:
*   [ left spacer | centered OMML | caption (prefix + STYLEREFs + SEQ + suffix) ]
*
* EquationBlock without captionId → single centered paragraph with OMML.
*
* CaptionBlock → single paragraph: prefix + STYLEREFs + SEQ + suffix +
* bodySeparator + text (the last two omitted when text is empty).
*
* CaptionCounterReset → single paragraph with one SEQ `\r N` field
* wrapped in `<w:vanish/>` rPr — invisible counter-advance marker.
*
* All caption-bearing emitters return both the XML root element AND a
* PendingCaptionFill / PendingCaptionReset record. The apply pipeline
* collects these, passes to the counter simulator, then backfills the
* computed values into the result text elements.
*
* Bookmark scope (per spec §4.7): the primary bookmark wraps prefix run
* through suffix run — number + decoration, NOT the bodySeparator + body
* for CaptionBlock. Caller supplies the pre-allocated `{ id, name }` via
* `bookmark` option; emit emits bookmarkStart/End around the correct
* range. When `bookmark` is omitted, no bookmark is emitted.
*/
const w$14 = NS.w;
const m$2 = NS.m;
/** SEQ identifier for the hidden chapter counter paired with a heading
* style under chapterPrefix `format` override. Engine reserves the
* `_chap_` prefix; agent-declared captionIds matching this pattern
* throw at schema validation. */
function chapterCounterIdentifier(styleId) {
	return `_chap_${styleId}`;
}
/** Emit a numbered equation as a 3-col borderless table. Middle cell
* holds the OMML; right cell holds the caption (prefix + chapter +
* counter + suffix), bookmark-wrapped per spec §4.7. */
function emitNumberedEquation(ownerDoc, opts) {
	const total = opts.usableWidthTwips ?? 8500;
	const leftW = 850;
	const rightW = 1500;
	const midW = Math.max(2e3, total - leftW - rightW);
	const leftCell = buildCell(ownerDoc, leftW, [emptyParagraph(ownerDoc)]);
	const middleCell = buildCell(ownerDoc, midW, [centeredEquationParagraph(ownerDoc, opts.mathSource, opts.equationStyleId)]);
	const captionResult = buildCaptionParagraph(ownerDoc, {
		captionConfig: opts.captionConfig,
		subGroup: opts.subGroup,
		bookmark: opts.bookmark,
		body: void 0
	});
	const rightCell = buildCell(ownerDoc, rightW, [captionResult.paragraph]);
	return {
		table: buildBorderlessTable(ownerDoc, [
			leftW,
			midW,
			rightW
		], [
			leftCell,
			middleCell,
			rightCell
		]),
		fill: {
			paragraph: captionResult.paragraph,
			identifier: opts.captionConfig.identifier,
			subGroup: opts.subGroup,
			chapterPrefixResults: captionResult.chapterPrefixResults,
			parentSeqResult: captionResult.parentSeqResult,
			subSeqResult: captionResult.subSeqResult
		}
	};
}
/** Emit a caption paragraph (FigureCaption / TableCaption / Theorem /
* ...). Single paragraph: prefix + chapter STYLEREFs + SEQ + suffix +
* bodySeparator + text. Last two omitted when text is empty. */
function emitCaptionBlock(ownerDoc, opts) {
	const built = buildCaptionParagraph(ownerDoc, {
		captionConfig: opts.captionConfig,
		subGroup: opts.subGroup,
		bookmark: opts.bookmark,
		body: opts.text === "" ? void 0 : opts.text
	});
	const fill = {
		paragraph: built.paragraph,
		identifier: opts.captionConfig.identifier,
		subGroup: opts.subGroup,
		chapterPrefixResults: built.chapterPrefixResults,
		parentSeqResult: built.parentSeqResult,
		subSeqResult: built.subSeqResult,
		bodyText: opts.text === "" ? void 0 : opts.text
	};
	return {
		paragraph: built.paragraph,
		fill
	};
}
/** Emit a SEQ-marker paragraph (vanish-hidden) that resets the
* identifier's counter so the NEXT caption emits the agent's specified
* `newValue`. Word's `SEQ \r N` resets to N AND the marker itself
* emits N; the next visible SEQ increments to N+1. To match agent
* intent ("the next caption is newValue"), emit `\r (newValue - 1)`.
* Counter sim mirrors the same convention. */
function emitCaptionReset(ownerDoc, opts) {
	const p = ownerDoc.createElementNS(w$14, "w:p");
	const pPr = ownerDoc.createElementNS(w$14, "w:pPr");
	const pStyle = ownerDoc.createElementNS(w$14, "w:pStyle");
	pStyle.setAttributeNS(w$14, "w:val", "Normal");
	pPr.appendChild(pStyle);
	p.appendChild(pPr);
	const { runs } = emitSeqField(ownerDoc, {
		identifier: opts.identifier,
		format: "arabic",
		resetTo: opts.newValue - 1
	});
	for (const r of runs) {
		addVanishRPr(r, ownerDoc);
		p.appendChild(r);
	}
	return {
		paragraph: p,
		reset: {
			paragraph: p,
			identifier: opts.identifier,
			newValue: opts.newValue
		}
	};
}
/** Build the run sequence that goes inside a caption paragraph (no
* wrapping `<w:p>`, no `<w:pPr>`). Shared by `buildCaptionParagraph`
* (which wraps in a fresh paragraph) and `standardize-captions` (which
* appends into an existing paragraph after preserving its pPr). */
function buildCaptionRunSequence(ownerDoc, opts) {
	const config = opts.captionConfig;
	const runs = [];
	if (opts.bookmark) {
		const bmStart = ownerDoc.createElementNS(w$14, "w:bookmarkStart");
		bmStart.setAttributeNS(w$14, "w:id", String(opts.bookmark.id));
		bmStart.setAttributeNS(w$14, "w:name", opts.bookmark.name);
		runs.push(bmStart);
	}
	if (config.prefix !== "") runs.push(buildPlainTextRun(ownerDoc, config.prefix));
	// labelSeparator sits between the label and the chapter counter so that
	// "图" + chapterPrefix(2) + chapterSeparator(.) + SEQ(1) can render as
	// "图 2.1". Emitted only when chapter numbering is active — otherwise it
	// would wedge a gap between the label and the bare counter.
	if (config.chapterPrefix.length > 0 && config.labelSeparator) runs.push(buildPlainTextRun(ownerDoc, config.labelSeparator));
	const chapterPrefixResults = [];
	for (const entry of config.chapterPrefix) {
		let resultTextEl;
		if (entry.format !== void 0) {
			const { runs: seqRuns, resultTextEl: t } = emitSeqField(ownerDoc, {
				identifier: chapterCounterIdentifier(entry.styleId),
				format: entry.format,
				repeat: true
			});
			for (const r of seqRuns) runs.push(r);
			resultTextEl = t;
		} else {
			const { runs: styleRefRuns, resultTextEl: t } = emitStyleRefField(ownerDoc, {
				styleName: entry.styleName,
				switches: ["\\n"]
			});
			for (const r of styleRefRuns) runs.push(r);
			resultTextEl = t;
		}
		chapterPrefixResults.push(resultTextEl);
		runs.push(buildPlainTextRun(ownerDoc, config.chapterSeparator));
	}
	const parent = emitSeqField(ownerDoc, {
		identifier: config.identifier,
		format: config.format,
		restartAtOutlineLevel: config.restartAtOutlineLevel,
		repeat: opts.subGroup === "continue"
	});
	for (const r of parent.runs) runs.push(r);
	let subSeqResult;
	if (opts.subGroup !== void 0 && config.subCounter) {
		if (config.subCounter.prefix !== "") runs.push(buildPlainTextRun(ownerDoc, config.subCounter.prefix));
		const sub = emitSeqField(ownerDoc, {
			identifier: `${config.identifier}Sub`,
			format: config.subCounter.format,
			resetTo: opts.subGroup === "start" ? 1 : void 0
		});
		for (const r of sub.runs) runs.push(r);
		subSeqResult = sub.resultTextEl;
		if (config.subCounter.suffix !== "") runs.push(buildPlainTextRun(ownerDoc, config.subCounter.suffix));
	}
	if (config.suffix !== "") runs.push(buildPlainTextRun(ownerDoc, config.suffix));
	if (opts.bookmark) {
		const bmEnd = ownerDoc.createElementNS(w$14, "w:bookmarkEnd");
		bmEnd.setAttributeNS(w$14, "w:id", String(opts.bookmark.id));
		runs.push(bmEnd);
	}
	if (opts.body !== void 0) {
		runs.push(buildPlainTextRun(ownerDoc, config.bodySeparator));
		runs.push(buildPlainTextRun(ownerDoc, opts.body));
	}
	return {
		runs,
		chapterPrefixResults,
		parentSeqResult: parent.resultTextEl,
		subSeqResult
	};
}
function buildCaptionParagraph(ownerDoc, opts) {
	const p = ownerDoc.createElementNS(w$14, "w:p");
	const pPr = ownerDoc.createElementNS(w$14, "w:pPr");
	const pStyle = ownerDoc.createElementNS(w$14, "w:pStyle");
	pStyle.setAttributeNS(w$14, "w:val", opts.captionConfig.paragraphStyleId);
	pPr.appendChild(pStyle);
	p.appendChild(pPr);
	const seq = buildCaptionRunSequence(ownerDoc, opts);
	for (const r of seq.runs) p.appendChild(r);
	return {
		paragraph: p,
		chapterPrefixResults: seq.chapterPrefixResults,
		parentSeqResult: seq.parentSeqResult,
		subSeqResult: seq.subSeqResult
	};
}
function centeredEquationParagraph(ownerDoc, source, equationStyleId) {
	const p = ownerDoc.createElementNS(w$14, "w:p");
	const pPr = ownerDoc.createElementNS(w$14, "w:pPr");
	const pStyle = ownerDoc.createElementNS(w$14, "w:pStyle");
	pStyle.setAttributeNS(w$14, "w:val", equationStyleId);
	pPr.appendChild(pStyle);
	p.appendChild(pPr);
	const oMath = buildOMath(source, ownerDoc, true);
	const oMathPara = ownerDoc.createElementNS(m$2, "m:oMathPara");
	oMathPara.appendChild(oMath);
	p.appendChild(oMathPara);
	return p;
}
function emptyParagraph(ownerDoc) {
	return ownerDoc.createElementNS(w$14, "w:p");
}
function buildCell(ownerDoc, widthTwips, content) {
	const tc = ownerDoc.createElementNS(w$14, "w:tc");
	const tcPr = ownerDoc.createElementNS(w$14, "w:tcPr");
	const tcW = ownerDoc.createElementNS(w$14, "w:tcW");
	tcW.setAttributeNS(w$14, "w:w", String(widthTwips));
	tcW.setAttributeNS(w$14, "w:type", "dxa");
	tcPr.appendChild(tcW);
	tc.appendChild(tcPr);
	for (const c of content) tc.appendChild(c);
	return tc;
}
function buildBorderlessTable(ownerDoc, colWidths, cells) {
	const tbl = ownerDoc.createElementNS(w$14, "w:tbl");
	const tblPr = ownerDoc.createElementNS(w$14, "w:tblPr");
	const tblW = ownerDoc.createElementNS(w$14, "w:tblW");
	tblW.setAttributeNS(w$14, "w:w", String(colWidths.reduce((a, b) => a + b, 0)));
	tblW.setAttributeNS(w$14, "w:type", "dxa");
	tblPr.appendChild(tblW);
	const borders = ownerDoc.createElementNS(w$14, "w:tblBorders");
	for (const edge of [
		"top",
		"left",
		"bottom",
		"right",
		"insideH",
		"insideV"
	]) {
		const b = ownerDoc.createElementNS(w$14, `w:${edge}`);
		b.setAttributeNS(w$14, "w:val", "none");
		b.setAttributeNS(w$14, "w:sz", "0");
		b.setAttributeNS(w$14, "w:color", "auto");
		borders.appendChild(b);
	}
	tblPr.appendChild(borders);
	const layout = ownerDoc.createElementNS(w$14, "w:tblLayout");
	layout.setAttributeNS(w$14, "w:type", "fixed");
	tblPr.appendChild(layout);
	tbl.appendChild(tblPr);
	const tblGrid = ownerDoc.createElementNS(w$14, "w:tblGrid");
	for (const width of colWidths) {
		const gc = ownerDoc.createElementNS(w$14, "w:gridCol");
		gc.setAttributeNS(w$14, "w:w", String(width));
		tblGrid.appendChild(gc);
	}
	tbl.appendChild(tblGrid);
	const tr = ownerDoc.createElementNS(w$14, "w:tr");
	for (const c of cells) tr.appendChild(c);
	tbl.appendChild(tr);
	return tbl;
}

//#endregion
//#region lib/edit/math/equation-emit.ts
/**
* EquationBlock + InlineEquation → OOXML.
*
* Block form emits a `<w:p>` carrying `<m:oMathPara>` (centered display
* math). Inline form returns the `<m:oMath>` element to splice into a run
* sequence inside an existing paragraph.
*
* The OMML produced by `latexToOmml` already declares `xmlns:m` (and the
* unused-but-harmless `xmlns:w`) on the `<m:oMath>` root — xmldom preserves
* these on import, so the result serializes correctly without needing to
* touch the document root's namespace declarations.
*/
const w$13 = NS.w;
const m$1 = NS.m;
function emitEquationBlock(block, ownerDoc, ctx) {
	if (block.captionId !== void 0) return emitNumberedEquationDispatch(block, block.captionId, ownerDoc, ctx);
	return emitUnnumberedEquationLegacy(block, ownerDoc, ctx);
}
function emitNumberedEquationDispatch(block, captionId, ownerDoc, ctx) {
	if (!ctx.captions) throw new Error("EquationBlock.captionId: ctx.captions callbacks not provided by the engine. Numbered equations require the captions table to be declared in the apply config.");
	const config = ctx.captions.resolve(captionId);
	if (!config) throw new Error(`EquationBlock: captionId "${captionId}" is not declared in captions table.`);
	if (block.subGroup !== void 0 && config.subCounter === void 0) throw new Error(`EquationBlock: subGroup="${block.subGroup}" requires captions["${captionId}"].subCounter to be declared.`);
	const mathSource = block.latex !== void 0 ? { latex: block.latex } : { omml: block.omml };
	const bookmark = block.anchor !== void 0 ? ctx.captions.allocateBookmark(block.anchor) : void 0;
	const { table, fill } = emitNumberedEquation(ownerDoc, {
		mathSource,
		equationStyleId: block.styleId ?? "Equation",
		captionConfig: config,
		subGroup: block.subGroup,
		bookmark,
		usableWidthTwips: ctx.usableWidthTwips
	});
	if (block.anchor !== void 0) ctx.captions.bindBookmark(block.anchor, fill.paragraph);
	ctx.captions.registerFill(fill);
	return table;
}
function emitUnnumberedEquationLegacy(block, ownerDoc, ctx) {
	const p = ownerDoc.createElementNS(w$13, "w:p");
	if (block.styleId !== void 0 || block.paraFormat !== void 0) {
		const pPr = ownerDoc.createElementNS(w$13, "w:pPr");
		if (block.styleId) {
			const ps = ownerDoc.createElementNS(w$13, "w:pStyle");
			ps.setAttributeNS(w$13, "w:val", block.styleId);
			pPr.appendChild(ps);
		}
		if (block.paraFormat) for (const c of buildPPrChildren(block.paraFormat, ownerDoc)) pPr.appendChild(c);
		p.appendChild(pPr);
	}
	const oMath = buildOMath(block.latex !== void 0 ? { latex: block.latex } : block.omml !== void 0 ? { omml: block.omml } : (() => {
		throw new Error("EquationBlock: no math source — schema validation should have caught this. This is an engine invariant violation.");
	})(), ownerDoc, true);
	const oMathPara = ownerDoc.createElementNS(m$1, "m:oMathPara");
	oMathPara.appendChild(oMath);
	p.appendChild(oMathPara);
	if (block.anchor) {
		if (!ctx.adoptAnchor) throw new Error("EquationBlock.anchor encountered but ctx.adoptAnchor was not provided by the engine");
		ctx.adoptAnchor(block.anchor, p);
	}
	return p;
}
/** Build the `<m:oMath>` element for an inline equation. Caller splices it
* between `<w:r>` siblings inside the paragraph. */
function emitInlineEquation(latex, ownerDoc) {
	return buildOMath({ latex }, ownerDoc, false);
}

//#endregion
//#region lib/edit/fragment-emit.ts
/**
* Fragment → OOXML elements.
*
* Each Block emits exactly one <w:p>. The emitter's job is mechanical
* serialization; the content model (RichText, RunFormat, ParagraphFormat)
* is what the agent declared. New block types or new format properties = a
* new case here, contained.
*
* The format helpers (`buildRPrChildren`, `buildPPrChildren`) are shared
* with edit-engine.ts — the format op clears managed children and appends
* the same elements this emitter would produce for fresh paragraphs.
*
* Image emission is wired through a callback (`ImageEmitter`) so this
* module stays free of the image-asset / zip-mutation concern. The engine
* supplies the callback when image assets are available; absent callback +
* `image` block = error (Step 7 plumbs this in via image-asset.ts).
*/
const w$12 = NS.w;
/** OOXML toggle element: presence-only when `on=true`, val="0" when false. */
function toggleElement(ownerDoc, qname, on) {
	const el = ownerDoc.createElementNS(w$12, qname);
	if (!on) el.setAttributeNS(w$12, "w:val", "0");
	return el;
}
function buildRPrChildren(fmt, ownerDoc) {
	const out = [];
	if (fmt.fontLatin || fmt.fontCJK) {
		const rFonts = ownerDoc.createElementNS(w$12, "w:rFonts");
		const ascii = fmt.fontLatin ?? fmt.fontCJK ?? "";
		const ea = fmt.fontCJK ?? fmt.fontLatin ?? "";
		if (ascii) {
			rFonts.setAttributeNS(w$12, "w:ascii", ascii);
			rFonts.setAttributeNS(w$12, "w:hAnsi", ascii);
		}
		if (ea) rFonts.setAttributeNS(w$12, "w:eastAsia", ea);
		out.push(rFonts);
	}
	if (fmt.size !== void 0) {
		const halfPt = toHalfPt(fmt.size, "size");
		const sz = ownerDoc.createElementNS(w$12, "w:sz");
		sz.setAttributeNS(w$12, "w:val", String(halfPt));
		out.push(sz);
		const szCs = ownerDoc.createElementNS(w$12, "w:szCs");
		szCs.setAttributeNS(w$12, "w:val", String(halfPt));
		out.push(szCs);
	}
	if (fmt.bold !== void 0) {
		out.push(toggleElement(ownerDoc, "w:b", fmt.bold));
		out.push(toggleElement(ownerDoc, "w:bCs", fmt.bold));
	}
	if (fmt.italic !== void 0) {
		out.push(toggleElement(ownerDoc, "w:i", fmt.italic));
		out.push(toggleElement(ownerDoc, "w:iCs", fmt.italic));
	}
	if (fmt.underline !== void 0) {
		const u = ownerDoc.createElementNS(w$12, "w:u");
		u.setAttributeNS(w$12, "w:val", fmt.underline ? "single" : "none");
		out.push(u);
	}
	if (fmt.strike !== void 0) out.push(toggleElement(ownerDoc, "w:strike", fmt.strike));
	if (fmt.doubleStrike !== void 0) out.push(toggleElement(ownerDoc, "w:dstrike", fmt.doubleStrike));
	if (fmt.charSpacing !== void 0) {
		// Character spacing: w:spacing @w:val is in twips (1/20 pt). Bare
		// numbers are pt, same convention as every other Length field.
		const sp = ownerDoc.createElementNS(w$12, "w:spacing");
		sp.setAttributeNS(w$12, "w:val", String(toTwips(fmt.charSpacing, "charSpacing")));
		out.push(sp);
	}
	if (fmt.emphasisMark) {
		const em = ownerDoc.createElementNS(w$12, "w:em");
		em.setAttributeNS(w$12, "w:val", fmt.emphasisMark);
		out.push(em);
	}
	if (fmt.color) {
		const c = ownerDoc.createElementNS(w$12, "w:color");
		c.setAttributeNS(w$12, "w:val", fmt.color);
		out.push(c);
	}
	if (fmt.vertAlign) {
		const va = ownerDoc.createElementNS(w$12, "w:vertAlign");
		va.setAttributeNS(w$12, "w:val", fmt.vertAlign);
		out.push(va);
	}
	if (fmt.rtl !== void 0) out.push(toggleElement(ownerDoc, "w:rtl", fmt.rtl));
	const orderIdx = (el) => {
		const i = RPR_CHILD_ORDER.indexOf(el.localName);
		return i < 0 ? RPR_CHILD_ORDER.length : i;
	};
	out.sort((a, b) => orderIdx(a) - orderIdx(b));
	return out;
}
/** Run-property children that `buildRPrChildren` can produce — exposed so
* edit-engine's `format` op can clear exactly these before re-applying the
* builder's output, preserving anything the user didn't ask to change (lang,
* kern, w, ...).
*
* Scope: run-level (from `RunFormat`). Wider than style-mutation's
* RPR_MANAGED_CHILDREN — RunFormat has `u` and `strike` which
* `StyleConfigEntry` does not. Don't fold the two sets together; their
* scopes are intentionally different. */
const RPR_MANAGED_LOCAL_NAMES = new Set([
	"rFonts",
	"sz",
	"szCs",
	"b",
	"bCs",
	"i",
	"iCs",
	"u",
	"strike",
	"dstrike",
	"spacing",
	"em",
	"color",
	"vertAlign",
	"rtl"
]);
function buildPPrChildren(fmt, ownerDoc) {
	const out = [];
	// `<w:bidi/>` is the first child the builder can emit (CT_PPrBase order:
	// ... autoSpaceDN, bidi, adjustRightInd, snapToGrid, spacing, ind, ...),
	// and this list is written in push order — so bidi goes on top.
	if (fmt.bidi !== void 0) out.push(toggleElement(ownerDoc, "w:bidi", fmt.bidi));
	if (fmt.spaceBefore !== void 0 || fmt.spaceAfter !== void 0 || fmt.lineSpacing !== void 0) {
		const spacing = ownerDoc.createElementNS(w$12, "w:spacing");
		if (fmt.spaceBefore !== void 0) spacing.setAttributeNS(w$12, "w:before", String(toTwips(fmt.spaceBefore, "spaceBefore")));
		if (fmt.spaceAfter !== void 0) spacing.setAttributeNS(w$12, "w:after", String(toTwips(fmt.spaceAfter, "spaceAfter")));
		if (fmt.lineSpacing !== void 0) {
			const ls = parseLineSpacing(fmt.lineSpacing, "lineSpacing");
			spacing.setAttributeNS(w$12, "w:line", String(ls.value));
			spacing.setAttributeNS(w$12, "w:lineRule", ls.mode);
		}
		out.push(spacing);
	}
	if (fmt.firstLineIndent != null || fmt.hangingIndent != null || fmt.indentLeft != null || fmt.indentRight != null) {
		const ind = ownerDoc.createElementNS(w$12, "w:ind");
		const fli = parseIndent(fmt.firstLineIndent ?? null);
		if (fli) setIndentAttr(ind, "firstLine", fli);
		const hi = parseIndent(fmt.hangingIndent ?? null);
		if (hi) setIndentAttr(ind, "hanging", hi);
		const il = parseIndent(fmt.indentLeft ?? null);
		if (il) setIndentAttr(ind, "left", il);
		const ir = parseIndent(fmt.indentRight ?? null);
		if (ir) setIndentAttr(ind, "right", ir);
		out.push(ind);
	}
	if (fmt.alignment) {
		const jc = ownerDoc.createElementNS(w$12, "w:jc");
		jc.setAttributeNS(w$12, "w:val", fmt.alignment);
		out.push(jc);
	}
	if (fmt.outlineLevel !== void 0) {
		const ol = ownerDoc.createElementNS(w$12, "w:outlineLvl");
		ol.setAttributeNS(w$12, "w:val", String(fmt.outlineLevel));
		out.push(ol);
	}
	return out;
}
/** pPr children that buildPPrChildren manages (excluding pStyle and numPr,
* which are addressed separately by the engine — pStyle through styleId
* assignment, numPr through numbering binding). */
const PPR_MANAGED_LOCAL_NAMES = new Set([
	"bidi",
	"spacing",
	"ind",
	"jc",
	"outlineLvl"
]);
function emitRun(text, format, ownerDoc) {
	const r = ownerDoc.createElementNS(w$12, "w:r");
	if (format) {
		const rPr = ownerDoc.createElementNS(w$12, "w:rPr");
		for (const c of buildRPrChildren(format, ownerDoc)) rPr.appendChild(c);
		if (rPr.childNodes.length > 0) r.appendChild(rPr);
	}
	const t = ownerDoc.createElementNS(w$12, "w:t");
	t.setAttribute("xml:space", "preserve");
	t.appendChild(ownerDoc.createTextNode(text));
	r.appendChild(t);
	return r;
}
function emitRichText(rt, ownerDoc, ctx, defaultFormat) {
	if (typeof rt === "string") return rt.length === 0 ? [] : [emitRun(rt, defaultFormat, ownerDoc)];
	const out = [];
	for (const piece of rt) {
		if ("refTo" in piece) {
			if (!ctx.emitRef) throw new Error("InlineRef encountered but ctx.emitRef was not provided by the engine");
			const fmt = piece.format ?? defaultFormat;
			for (const r of ctx.emitRef(piece, ownerDoc, fmt)) out.push(r);
			continue;
		}
		if ("math" in piece) {
			out.push(emitInlineEquation(piece.math, ownerDoc));
			continue;
		}
		if ("link" in piece) {
			if (!ctx.emitHyperlink) throw new Error("HyperlinkNode encountered but ctx.emitHyperlink was not provided by the engine");
			const fmt = piece.format ?? defaultFormat;
			out.push(ctx.emitHyperlink(piece.link, piece.text, fmt, ownerDoc));
			continue;
		}
		if ("field" in piece) {
			const fmt = piece.format ?? defaultFormat;
			for (const r of emitInlineField(ownerDoc, piece.field, fmt)) out.push(r);
			continue;
		}
		if ("styleRef" in piece) {
			const fmt = piece.format ?? defaultFormat;
			if (!ctx.resolveStyle) throw new Error("InlineStyleRef encountered but ctx.resolveStyle was not provided. The engine must wire a styleId → StyleInfo resolver so STYLEREF picks a locale-safe field code shape at emit time.");
			const info = ctx.resolveStyle(piece.styleRef);
			if (!info) throw new Error(`InlineStyleRef: styleId "${piece.styleRef}" not found in styles.xml. Declare it via styles[] in this apply, or reference an existing styleId.`);
			for (const r of emitInlineStyleRef(ownerDoc, info, piece.numberOnly ?? false, fmt)) out.push(r);
			continue;
		}
		if ("break" in piece) {
			const r = ownerDoc.createElementNS(w$12, "w:r");
			const br = ownerDoc.createElementNS(w$12, "w:br");
			if (piece.break !== "line") br.setAttributeNS(w$12, "w:type", piece.break);
			r.appendChild(br);
			out.push(r);
			continue;
		}
		out.push(emitRun(piece.text, piece.format ?? defaultFormat, ownerDoc));
	}
	return out;
}
function ensurePPr(p, ownerDoc) {
	for (const c of Array.from(p.childNodes)) if (c.nodeType === 1) {
		const el = c;
		if (el.namespaceURI === w$12 && el.localName === "pPr") return el;
	}
	const pPr = ownerDoc.createElementNS(w$12, "w:pPr");
	p.insertBefore(pPr, p.firstChild);
	return pPr;
}
function emitParagraphBlock(block, ownerDoc, ctx) {
	const p = ownerDoc.createElementNS(w$12, "w:p");
	if (block.styleId !== void 0 || block.paraFormat !== void 0 || block.numbering !== void 0) {
		const pPr = ensurePPr(p, ownerDoc);
		if (block.styleId) {
			const ps = ownerDoc.createElementNS(w$12, "w:pStyle");
			ps.setAttributeNS(w$12, "w:val", block.styleId);
			pPr.appendChild(ps);
		}
		if (block.numbering) if (block.numbering.restart) {
			if (!ctx.forkNumRestart) throw new Error("ParagraphBlock.numbering.restart: true requires numberingDoc to be available. Pass numberingDoc in RunEditOpsInput (apply-styles always does; standalone edit runs need it when using numbering.restart).");
			ctx.forkNumRestart(p, block.numbering.numId, block.numbering.level);
		} else {
			const numPr = ownerDoc.createElementNS(w$12, "w:numPr");
			const ilvl = ownerDoc.createElementNS(w$12, "w:ilvl");
			ilvl.setAttributeNS(w$12, "w:val", String(block.numbering.level));
			numPr.appendChild(ilvl);
			const numId = ownerDoc.createElementNS(w$12, "w:numId");
			numId.setAttributeNS(w$12, "w:val", block.numbering.numId);
			numPr.appendChild(numId);
			pPr.appendChild(numPr);
		}
		if (block.paraFormat) for (const c of buildPPrChildren(block.paraFormat, ownerDoc)) pPr.appendChild(c);
	}
	for (const r of emitRichText(block.text, ownerDoc, ctx, block.runFormat)) p.appendChild(r);
	if (block.anchor) {
		if (!ctx.adoptAnchor) throw new Error("ParagraphBlock.anchor encountered but ctx.adoptAnchor was not provided by the engine");
		ctx.adoptAnchor(block.anchor, p);
	}
	return p;
}
function emitImageBlock(block, ownerDoc, ctx) {
	if (!ctx.emitImage) throw new Error("image block encountered but no ImageEmitter wired (image asset path inactive)");
	const drawing = ctx.emitImage(block.src, block.width, block.height, block.alt, ownerDoc);
	const p = ownerDoc.createElementNS(w$12, "w:p");
	if (block.styleId !== void 0 || block.paraFormat !== void 0) {
		const pPr = ensurePPr(p, ownerDoc);
		if (block.styleId) {
			const ps = ownerDoc.createElementNS(w$12, "w:pStyle");
			ps.setAttributeNS(w$12, "w:val", block.styleId);
			pPr.appendChild(ps);
		}
		if (block.paraFormat) for (const c of buildPPrChildren(block.paraFormat, ownerDoc)) pPr.appendChild(c);
	}
	const r = ownerDoc.createElementNS(w$12, "w:r");
	r.appendChild(drawing);
	p.appendChild(r);
	return p;
}
function emitPageBreakBlock(ownerDoc) {
	const p = ownerDoc.createElementNS(w$12, "w:p");
	const r = ownerDoc.createElementNS(w$12, "w:r");
	const br = ownerDoc.createElementNS(w$12, "w:br");
	br.setAttributeNS(w$12, "w:type", "page");
	r.appendChild(br);
	p.appendChild(r);
	return p;
}
function emitHorizontalRuleBlock(ownerDoc) {
	const p = ownerDoc.createElementNS(w$12, "w:p");
	const pPr = ensurePPr(p, ownerDoc);
	const pBdr = ownerDoc.createElementNS(w$12, "w:pBdr");
	const bottom = ownerDoc.createElementNS(w$12, "w:bottom");
	bottom.setAttributeNS(w$12, "w:val", "single");
	bottom.setAttributeNS(w$12, "w:sz", "6");
	bottom.setAttributeNS(w$12, "w:space", "1");
	bottom.setAttributeNS(w$12, "w:color", "auto");
	pBdr.appendChild(bottom);
	pPr.appendChild(pBdr);
	return p;
}
function emitBlock(block, ownerDoc, ctx) {
	switch (block.type) {
		case "paragraph": return emitParagraphBlock(block, ownerDoc, ctx);
		case "image": return emitImageBlock(block, ownerDoc, ctx);
		case "page-break": return emitPageBreakBlock(ownerDoc);
		case "horizontal-rule": return emitHorizontalRuleBlock(ownerDoc);
		case "table": return emitTableBlock(block, ownerDoc, ctx);
		case "equation": return emitEquationBlock(block, ownerDoc, ctx);
		case "caption": return dispatchCaption(block, ownerDoc, ctx);
		case "caption-counter-reset": return dispatchCaptionReset(block, ownerDoc, ctx);
		default: return assertNever(block);
	}
}
function dispatchCaption(block, ownerDoc, ctx) {
	if (!ctx.captions) throw new Error("CaptionBlock: ctx.captions callbacks not provided by the engine. Caption blocks require the captions table to be declared in the apply config.");
	const config = ctx.captions.resolve(block.captionId);
	if (!config) throw new Error(`CaptionBlock: captionId "${block.captionId}" is not declared in captions table.`);
	const bookmark = block.anchor !== void 0 ? ctx.captions.allocateBookmark(block.anchor) : void 0;
	const { paragraph, fill } = emitCaptionBlock(ownerDoc, {
		captionConfig: config,
		text: block.text,
		bookmark
	});
	if (block.anchor !== void 0) ctx.captions.bindBookmark(block.anchor, paragraph);
	ctx.captions.registerFill(fill);
	return paragraph;
}
function dispatchCaptionReset(block, ownerDoc, ctx) {
	if (!ctx.captions) throw new Error("CaptionCounterReset: ctx.captions callbacks not provided by the engine.");
	if (!ctx.captions.resolve(block.captionId)) throw new Error(`CaptionCounterReset: captionId "${block.captionId}" is not declared in captions table.`);
	const { paragraph, reset } = emitCaptionReset(ownerDoc, {
		identifier: block.captionId,
		newValue: block.newValue ?? 1
	});
	ctx.captions.registerReset(reset);
	return paragraph;
}
function emitFragment(fragment, ownerDoc, ctx) {
	return fragment.map((b) => emitBlock(b, ownerDoc, ctx));
}

//#endregion
//#region lib/edit/track-changes.ts
/**
* Track-changes wrappers — produce <w:ins> / <w:del> / <w:rPrChange> /
* <w:pPrChange> / <w:numPrChange> elements from edit ops.
*
* The engine wires `TrackContext.enabled`; when false, every helper is a
* no-op (deletion still removes, insertion still inserts plain). When
* true, deletions get wrapped, insertions get wrapped, format ops attach
* a snapshot of the previous run / paragraph properties.
*
* Author and date defaults come from TrackContext (author empty, date
* fixed for the run). The schema (ECMA-376 §17.13) requires both
* attributes; we emit empty author since identity is not in scope.
*/
const w$11 = NS.w;
/**
* Wrap a list of paragraph children (typically <w:r> elements) inside a
* single <w:ins>. The caller passes the already-built run nodes; this just
* groups them. Returns the <w:ins> element ready to insert.
*/
function wrapInsertion(runs, ownerDoc, ctx) {
	const ins = ownerDoc.createElementNS(w$11, "w:ins");
	ins.setAttributeNS(w$11, "w:id", String(ctx.nextId()));
	ins.setAttributeNS(w$11, "w:author", ctx.author);
	ins.setAttributeNS(w$11, "w:date", ctx.date);
	for (const r of runs) ins.appendChild(r);
	return ins;
}
/**
* Wrap a list of <w:r> elements inside a <w:del> element, converting each
* descendant <w:t> to <w:delText> per ECMA-376 §17.4.16. The runs are
* detached from any current parent before being re-appended; callers are
* responsible for splicing the returned element back into the tree.
*/
function wrapDeletion(runs, ownerDoc, ctx) {
	const del = ownerDoc.createElementNS(w$11, "w:del");
	del.setAttributeNS(w$11, "w:id", String(ctx.nextId()));
	del.setAttributeNS(w$11, "w:author", ctx.author);
	del.setAttributeNS(w$11, "w:date", ctx.date);
	for (const r of runs) {
		convertTtoDelText(r, ownerDoc);
		del.appendChild(r);
	}
	return del;
}
function convertTtoDelText(node, ownerDoc) {
	for (const child of Array.from(getChildren(node))) if (child.namespaceURI === w$11) if (child.localName === "t") replaceLocalName(child, "w:delText", ownerDoc);
	else if (child.localName === "instrText") replaceLocalName(child, "w:delInstrText", ownerDoc);
	else convertTtoDelText(child, ownerDoc);
	else convertTtoDelText(child, ownerDoc);
}
function replaceLocalName(el, qname, ownerDoc) {
	const rep = ownerDoc.createElementNS(w$11, qname);
	for (let i = 0; i < el.attributes.length; i++) {
		const attr = el.attributes[i];
		if (attr.namespaceURI) rep.setAttributeNS(attr.namespaceURI, attr.name, attr.value);
		else rep.setAttribute(attr.name, attr.value);
	}
	while (el.firstChild) rep.appendChild(el.firstChild);
	el.parentNode.replaceChild(rep, el);
}
/**
* Mark the paragraph mark (<w:p>'s pilcrow) as deleted, so when the user
* accepts changes the paragraph merges with the next. Used when an entire
* paragraph is deleted under track-changes; without this, accept leaves an
* empty paragraph behind.
*
* Implementation: ensure <w:pPr><w:rPr><w:del/></w:rPr> exists.
*/
function markParagraphMarkDeleted(p, ownerDoc, ctx) {
	let pPr = firstChildNS(p, w$11, "pPr");
	if (!pPr) {
		pPr = ownerDoc.createElementNS(w$11, "w:pPr");
		p.insertBefore(pPr, p.firstChild);
	}
	let rPr = firstChildNS(pPr, w$11, "rPr");
	if (!rPr) {
		rPr = ownerDoc.createElementNS(w$11, "w:rPr");
		pPr.appendChild(rPr);
	}
	for (const c of getChildren(rPr)) if (c.namespaceURI === w$11 && c.localName === "del") return;
	const del = ownerDoc.createElementNS(w$11, "w:del");
	del.setAttributeNS(w$11, "w:id", String(ctx.nextId()));
	del.setAttributeNS(w$11, "w:author", ctx.author);
	del.setAttributeNS(w$11, "w:date", ctx.date);
	rPr.insertBefore(del, rPr.firstChild);
}
/**
* Wrap every <w:r> child of paragraph `p` inside a single <w:del>, marking
* the run content as deleted. The paragraph mark itself is NOT marked here —
* call `markParagraphMarkDeleted` separately if the whole paragraph is going.
*/
function wrapParagraphContentInDel(p, ownerDoc, ctx) {
	const runs = getChildrenNS(p, w$11, "r");
	if (runs.length === 0) return;
	const before = runs[0].previousSibling;
	const detached = [];
	for (const r of runs) {
		p.removeChild(r);
		detached.push(r);
	}
	const del = wrapDeletion(detached, ownerDoc, ctx);
	if (before && before.nextSibling) p.insertBefore(del, before.nextSibling);
	else if (before) p.appendChild(del);
	else p.insertBefore(del, p.firstChild);
}
/**
* Wrap every <w:r> child of paragraph `p` inside a single <w:ins>, and mark
* the paragraph mark as inserted (`<w:pPr><w:rPr><w:ins/>`). Used when a
* fresh paragraph is being added under track-changes — without the mark
* insertion, accept-all leaves the paragraph break as if it were original.
*/
function markParagraphAsInserted(p, ownerDoc, ctx) {
	const orderedChildren = Array.from(getChildren(p));
	let runGroup = [];
	let groupAnchor = null;
	const flushRunGroup = () => {
		if (runGroup.length === 0) return;
		const detached = [];
		for (const r of runGroup) {
			p.removeChild(r);
			detached.push(r);
		}
		const ins = wrapInsertion(detached, ownerDoc, ctx);
		if (groupAnchor && groupAnchor.nextSibling) p.insertBefore(ins, groupAnchor.nextSibling);
		else if (groupAnchor) p.appendChild(ins);
		else p.insertBefore(ins, p.firstChild);
		runGroup = [];
		groupAnchor = null;
	};
	for (const child of orderedChildren) {
		if (child.namespaceURI !== w$11) continue;
		if (child.localName === "r") {
			if (runGroup.length === 0) groupAnchor = child.previousSibling;
			runGroup.push(child);
			continue;
		}
		if (child.localName === "hyperlink") {
			flushRunGroup();
			const innerRuns = getChildrenNS(child, w$11, "r");
			if (innerRuns.length === 0) continue;
			const innerBefore = innerRuns[0].previousSibling;
			const innerDetached = [];
			for (const r of innerRuns) {
				child.removeChild(r);
				innerDetached.push(r);
			}
			const innerIns = wrapInsertion(innerDetached, ownerDoc, ctx);
			if (innerBefore && innerBefore.nextSibling) child.insertBefore(innerIns, innerBefore.nextSibling);
			else if (innerBefore) child.appendChild(innerIns);
			else child.insertBefore(innerIns, child.firstChild);
		}
	}
	flushRunGroup();
	let pPr = firstChildNS(p, w$11, "pPr");
	if (!pPr) {
		pPr = ownerDoc.createElementNS(w$11, "w:pPr");
		p.insertBefore(pPr, p.firstChild);
	}
	let rPr = firstChildNS(pPr, w$11, "rPr");
	if (!rPr) {
		rPr = ownerDoc.createElementNS(w$11, "w:rPr");
		pPr.appendChild(rPr);
	}
	for (const c of getChildren(rPr)) if (c.namespaceURI === w$11 && c.localName === "ins") return;
	const insMark = ownerDoc.createElementNS(w$11, "w:ins");
	insMark.setAttributeNS(w$11, "w:id", String(ctx.nextId()));
	insMark.setAttributeNS(w$11, "w:author", ctx.author);
	insMark.setAttributeNS(w$11, "w:date", ctx.date);
	rPr.insertBefore(insMark, rPr.firstChild);
}
/**
* Append a <w:rPrChange> snapshot to a freshly-mutated <w:rPr>. Snapshot is
* the *previous* state of the rPr (before mutation); pass null when there
* was no rPr at all (in which case <w:rPrChange> contains an empty <w:rPr/>
* to indicate "previous state was nothing").
*/
function attachRPrChange(rPr, previousRPrSnapshot, ownerDoc, ctx) {
	if (!ctx.enabled) return;
	const change = ownerDoc.createElementNS(w$11, "w:rPrChange");
	change.setAttributeNS(w$11, "w:id", String(ctx.nextId()));
	change.setAttributeNS(w$11, "w:author", ctx.author);
	change.setAttributeNS(w$11, "w:date", ctx.date);
	const inner = previousRPrSnapshot ? previousRPrSnapshot.cloneNode(true) : ownerDoc.createElementNS(w$11, "w:rPr");
	change.appendChild(inner);
	rPr.appendChild(change);
}
/**
* Same idea for paragraph properties. Snapshot of the previous <w:pPr> goes
* inside <w:pPrChange>. Schema: <w:pPrChange> is the LAST child of <w:pPr>,
* after even <w:rPr>; the wrapped snapshot may include its own <w:rPr> as
* the previous paragraph-mark formatting.
*/
function attachPPrChange(pPr, previousPPrSnapshot, ownerDoc, ctx) {
	if (!ctx.enabled) return;
	const change = ownerDoc.createElementNS(w$11, "w:pPrChange");
	change.setAttributeNS(w$11, "w:id", String(ctx.nextId()));
	change.setAttributeNS(w$11, "w:author", ctx.author);
	change.setAttributeNS(w$11, "w:date", ctx.date);
	const inner = previousPPrSnapshot ? previousPPrSnapshot.cloneNode(true) : ownerDoc.createElementNS(w$11, "w:pPr");
	change.appendChild(inner);
	pPr.appendChild(change);
}

//#endregion
//#region lib/edit/part-rels.ts
/**
* One docx part's `_rels/<part>.rels` state.
*
* Each OOXML part (document.xml, header1.xml, footer1.xml, …) has its own
* rels file with a per-part rId namespace. This class owns one such file —
* parsing, allocating fresh rIds, appending Relationship entries, and
* flushing the modified text back into the writer's replacement map.
*
* Mutation is string-level: the rels XML is small enough that DOM-mutate
* + serialize would add machinery for no benefit. `appendRel` rewrites
* `</Relationships>` with the new entry preceding it.
*/
const DEFAULT_RELS_XML = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>";
var PartRels = class PartRels {
	relsText;
	rels = [];
	nextRId = 1;
	dirty = false;
	/** Highest seen `media/imageN.<ext>` index + 1 — used by `nextImageNum()`
	*  helper so the asset-registry layer can dedup with existing media. */
	highestImageNum = 1;
	constructor(initialRelsXml) {
		this.relsText = initialRelsXml;
		this.parseRels();
	}
	/** Open from a docx reader. `relsPath` is the in-archive path
	*  (e.g. `word/_rels/document.xml.rels`). Falls back to an empty
	*  Relationships skeleton when the part has no existing rels file. */
	static async open(reader, relsPath) {
		return new PartRels(await reader.readText(relsPath) ?? DEFAULT_RELS_XML);
	}
	/** Append a Relationship, returning the freshly allocated rId. `mode`
	*  is "External" for hyperlinks / external links; omit for internal
	*  part-to-part references (images, headers, footers, …). Targets are
	*  XML-attr-escaped for agent-supplied values (hrefs); the caller is
	*  responsible for pre-encoding URLs if they want percent encoding. */
	appendRel(type, target, mode) {
		const rId = this.allocateRId();
		const modeAttr = mode ? ` TargetMode="${mode}"` : "";
		const insert = `<Relationship Id="${rId}" Type="${type}" Target="${escapeXmlAttr(target)}"${modeAttr}/>`;
		this.relsText = this.relsText.replace("</Relationships>", `${insert}</Relationships>`);
		this.rels.push({
			id: rId,
			target,
			type
		});
		this.dirty = true;
		return { rId };
	}
	/** Returns whether any rels were appended since construction. Callers
	*  gate flush on this so untouched parts don't rewrite their rels file. */
	isDirty() {
		return this.dirty;
	}
	/** Check whether a Relationship with the given Target string already
	*  exists. Used by callers that want to be idempotent about adding a
	*  named-part rel (e.g. the numbering.xml registration, which apply
	*  re-runs on a doc that already has the rel from a prior apply). */
	hasRelTo(target) {
		return this.rels.some((r) => r.target === target);
	}
	/** Stage the modified rels text at `path` in the writer's replacement
	*  map. No-op when nothing was appended. The body's rels path
	*  (`word/_rels/document.xml.rels`) is one of the forbidden direct-
	*  write targets on `WritableArchive`; uses the `_setFromAccumulator`
	*  escape hatch since this IS the legitimate accumulator path.
	*  Per-part HF rels paths (header1.xml.rels etc.) aren't forbidden
	*  and could go through `.set` too, but routing both through one
	*  method keeps the call sites uniform. */
	flushTo(replacements, path) {
		if (!this.dirty) return;
		replacements.setFromAccumulator(path, this.relsText);
	}
	/** Returns the next free `imageN` suffix for media path allocation.
	*  Used by the asset registry to avoid clashing with images already
	*  registered in this part's rels. */
	nextImageNum() {
		return this.highestImageNum;
	}
	/** Bump the image-number cursor past `n`. Called by the asset registry
	*  after it assigns `media/imageN.<ext>` for a fresh upload, so the
	*  next allocation skips ahead. */
	advanceImageNum(n) {
		if (n + 1 > this.highestImageNum) this.highestImageNum = n + 1;
	}
	parseRels() {
		const root = parseXml(this.relsText).documentElement;
		if (!root) return;
		let maxNum = 0;
		for (let i = 0; i < root.childNodes.length; i++) {
			const c = root.childNodes[i];
			if (!c || c.nodeType !== 1) continue;
			const el = c;
			if (el.localName !== "Relationship") continue;
			const id = el.getAttribute("Id") ?? "";
			const target = el.getAttribute("Target") ?? "";
			const type = el.getAttribute("Type") ?? "";
			if (id) {
				this.rels.push({
					id,
					target,
					type
				});
				const m = id.match(/^rId(\d+)$/);
				if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
			}
			const mt = target.match(/^(?:\.\.\/)?media\/image(\d+)\./);
			if (mt) {
				const n = parseInt(mt[1], 10) + 1;
				if (n > this.highestImageNum) this.highestImageNum = n;
			}
		}
		this.nextRId = maxNum + 1;
	}
	allocateRId() {
		return `rId${this.nextRId++}`;
	}
};
/** Minimal XML-attribute escape — Relationship targets like hyperlink hrefs
*  may contain `&` (query strings), `"`, `<`, `>`. Encode just enough to
*  keep the rels XML valid. */
function escapeXmlAttr(s) {
	return s.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

//#endregion
//#region lib/edit/content-types.ts
const DEFAULT_CONTENT_TYPES_XML = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"></Types>";
var ContentTypes = class ContentTypes {
	text;
	extensions = /* @__PURE__ */ new Set();
	overridePartNames = /* @__PURE__ */ new Set();
	dirty = false;
	constructor(initialText) {
		this.text = initialText;
		this.parseExtensions();
		this.parseOverrides();
	}
	static async open(reader) {
		return new ContentTypes((await reader.readText("[Content_Types].xml") ?? DEFAULT_CONTENT_TYPES_XML) || DEFAULT_CONTENT_TYPES_XML);
	}
	/** Ensure a `<Default Extension="..." ContentType="..."/>` entry exists.
	*  Idempotent on `ext` — duplicate calls with the same extension skip
	*  even if `mime` differs (each ext should resolve to exactly one MIME
	*  in practice; if a caller passes inconsistent MIMEs the first wins
	*  silently). */
	ensureDefault(ext, mime) {
		const key = ext.toLowerCase();
		if (this.extensions.has(key)) return;
		const insert = `<Default Extension="${ext}" ContentType="${mime}"/>`;
		if (this.text.includes(insert)) {
			this.extensions.add(key);
			return;
		}
		this.text = this.text.replace("</Types>", `${insert}</Types>`);
		this.extensions.add(key);
		this.dirty = true;
	}
	/** Ensure an `<Override PartName="..." ContentType="..."/>` entry exists.
	*  Used to register headerN.xml / footerN.xml / settings.xml / etc. by
	*  full part name. Idempotent. */
	ensureOverride(partName, contentType) {
		if (this.overridePartNames.has(partName)) return;
		const insert = `<Override PartName="${partName}" ContentType="${contentType}"/>`;
		if (this.text.includes(insert)) {
			this.overridePartNames.add(partName);
			return;
		}
		this.text = this.text.replace("</Types>", `${insert}</Types>`);
		this.overridePartNames.add(partName);
		this.dirty = true;
	}
	isDirty() {
		return this.dirty;
	}
	/** Stage the modified content-types XML in the writer's replacement map.
	*  No-op when nothing was added. Uses the `_setFromAccumulator` escape
	*  hatch — `WritableArchive.set` would reject `[Content_Types].xml` as
	*  one of the forbidden direct-write paths, the very invariant this
	*  accumulator exists to enforce. */
	flushTo(replacements) {
		if (!this.dirty) return;
		replacements.setFromAccumulator("[Content_Types].xml", this.text);
	}
	parseExtensions() {
		const re = /<Default[^>]*\bExtension="([^"]+)"/g;
		let m;
		while ((m = re.exec(this.text)) !== null) this.extensions.add(m[1].toLowerCase());
	}
	parseOverrides() {
		const re = /<Override[^>]*\bPartName="([^"]+)"/g;
		let m;
		while ((m = re.exec(this.text)) !== null) this.overridePartNames.add(m[1]);
	}
};

//#endregion
//#region lib/edit/asset-registry.ts
/**
* Document-asset registry — handles image and external-hyperlink registration
* for a single docx part (body or any header / footer). The class composes:
*
*   - `PartRels` — owns this part's `_rels/<part>.rels` (rId allocation +
*     Relationship appends)
*   - `ContentTypes` — shared `[Content_Types].xml` accumulator across all
*     parts in the apply run (singleton, passed in by the orchestrator)
*   - image-specific state: media binaries staged for the writer, docPr id
*     counter, source-path / href dedup caches
*
* For images, registration touches three parts besides the host XML:
*   1. word/media/imageN.<ext>      — the binary file (per-instance state)
*   2. <relsPath>                   — relationship from rId to media path
*   3. [Content_Types].xml          — Default entry for the file extension
*
* For external hyperlinks: just the rels entry (TargetMode="External").
* Internal hyperlinks (`#anchor` form) skip the registry entirely — they
* use `<w:hyperlink w:anchor="name">` with no rId.
*
* One instance serves one part. The body of document.xml gets one
* registry pointing at `word/_rels/document.xml.rels`; each header /
* footer part gets its own registry pointing at `word/_rels/headerN.xml.rels`
* (etc.). All registries share one ContentTypes accumulator.
*
* Inline drawing XML structure follows ECMA-376 §20.4.2.8 (DrawingML
* picture). Sizes are in EMU (1 pt = 12,700 EMU).
*/
const REL_TYPE_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const REL_TYPE_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture";
var DocxAssetRegistry = class DocxAssetRegistry {
	/** archivePath → bytes. New media binaries staged for the writer. */
	binaryAdditions = /* @__PURE__ */ new Map();
	nextDocPrId = 1;
	/** Cache: srcPath → already-registered rId + ext. Lets two image blocks
	*  pointing at the same file share one media entry. */
	byAbsPath = /* @__PURE__ */ new Map();
	/** Cache: external href → already-registered rId. Multiple hyperlinks to
	*  the same target share one Relationship entry — matches what Word writes
	*  when you paste the same URL twice. */
	byExternalHref = /* @__PURE__ */ new Map();
	constructor(partRels, contentTypes) {
		this.partRels = partRels;
		this.contentTypes = contentTypes;
	}
	/** Expose the shared ContentTypes accumulator. The orchestrator
	*  (apply-styles) flushes it once at the end of the apply run, and HF
	*  part registries get constructed with the same instance so all of
	*  the apply's content-type additions land in one file. */
	getContentTypes() {
		return this.contentTypes;
	}
	/** Expose this part's `PartRels` to peer subsystems that need to
	*  append non-image, non-hyperlink Relationships to the same file.
	*  Currently used by `applyHeaderFooter`: header / footer parts are
	*  referenced from the BODY's rels file (document.xml.rels), so the
	*  HF mutation grabs this registry's PartRels and appends a
	*  `Relationship` per HF part. Keeping rId allocation centralized
	*  here avoids racing the body's image / hyperlink rId stream. */
	getPartRels() {
		return this.partRels;
	}
	/** Convenience: open the body-part registry from a docx reader. Loads
	*  `word/_rels/document.xml.rels` and `[Content_Types].xml`. For header /
	*  footer parts, the orchestrator constructs `PartRels` and the shared
	*  `ContentTypes` directly and passes them in. */
	static async open(reader) {
		return new DocxAssetRegistry(await PartRels.open(reader, "word/_rels/document.xml.rels"), await ContentTypes.open(reader));
	}
	/** Returns the rId to use in <a:blip r:embed=...>. Reads the file, picks
	* a fresh archive path (word/media/imageN.ext), stages the binary, adds
	* the rel + content type if needed. Throws on missing source file. */
	registerImage(srcPath) {
		const abs = resolve(srcPath);
		if (!existsSync(abs)) throw new Error(`image not found: ${abs}`);
		const cached = this.byAbsPath.get(abs);
		if (cached) return {
			rId: cached.rId,
			ext: cached.ext
		};
		const rawExt = extname(abs).slice(1).toLowerCase();
		if (!rawExt) throw new Error(`image source has no extension: ${abs}`);
		const archivePath = this.uniqueMediaPath(rawExt);
		const bytes = new Uint8Array(readFileSync(abs));
		this.binaryAdditions.set(archivePath, bytes);
		this.contentTypes.ensureDefault(rawExt, mimeForExt(rawExt));
		const { rId } = this.partRels.appendRel(REL_TYPE_IMAGE, archivePath.replace(/^word\//, ""));
		this.byAbsPath.set(abs, {
			rId,
			ext: rawExt
		});
		return {
			rId,
			ext: rawExt
		};
	}
	/** Register an external hyperlink target. Returns the rId for the
	*  `<w:hyperlink r:id="...">` element. Identical hrefs share one rId
	*  (matches Word's deduplication). Internal `#anchor` links don't go
	*  through this path — they use `<w:hyperlink w:anchor="...">` with
	*  no rId. */
	registerExternalLink(href) {
		const cached = this.byExternalHref.get(href);
		if (cached) return { rId: cached };
		const { rId } = this.partRels.appendRel(REL_TYPE_HYPERLINK, href, "External");
		this.byExternalHref.set(href, rId);
		return { rId };
	}
	/** Build a single inline <w:drawing> element ready to live inside a <w:r>.
	* EMU-converts the dimensions; `alt` populates docPr name + descr. */
	buildDrawing(rId, width, height, alt, ownerDoc) {
		const cx = String(toEmu(width, "image.width"));
		const cy = String(toEmu(height, "image.height"));
		const docPrId = String(this.nextDocPrId++);
		const altText = alt ?? `Image ${docPrId}`;
		const drawing = ownerDoc.createElementNS(NS.w, "w:drawing");
		const inline = ownerDoc.createElementNS(NS.wp, "wp:inline");
		inline.setAttribute("distT", "0");
		inline.setAttribute("distB", "0");
		inline.setAttribute("distL", "0");
		inline.setAttribute("distR", "0");
		drawing.appendChild(inline);
		const extent = ownerDoc.createElementNS(NS.wp, "wp:extent");
		extent.setAttribute("cx", cx);
		extent.setAttribute("cy", cy);
		inline.appendChild(extent);
		const effectExtent = ownerDoc.createElementNS(NS.wp, "wp:effectExtent");
		effectExtent.setAttribute("l", "0");
		effectExtent.setAttribute("t", "0");
		effectExtent.setAttribute("r", "0");
		effectExtent.setAttribute("b", "0");
		inline.appendChild(effectExtent);
		const docPr = ownerDoc.createElementNS(NS.wp, "wp:docPr");
		docPr.setAttribute("id", docPrId);
		docPr.setAttribute("name", altText);
		if (alt) docPr.setAttribute("descr", alt);
		inline.appendChild(docPr);
		const cNvGraphicFramePr = ownerDoc.createElementNS(NS.wp, "wp:cNvGraphicFramePr");
		const graphicFrameLocks = ownerDoc.createElementNS(NS.a, "a:graphicFrameLocks");
		graphicFrameLocks.setAttribute("noChangeAspect", "1");
		cNvGraphicFramePr.appendChild(graphicFrameLocks);
		inline.appendChild(cNvGraphicFramePr);
		const graphic = ownerDoc.createElementNS(NS.a, "a:graphic");
		inline.appendChild(graphic);
		const graphicData = ownerDoc.createElementNS(NS.a, "a:graphicData");
		graphicData.setAttribute("uri", PIC_NS);
		graphic.appendChild(graphicData);
		const pic = ownerDoc.createElementNS(PIC_NS, "pic:pic");
		graphicData.appendChild(pic);
		const nvPicPr = ownerDoc.createElementNS(PIC_NS, "pic:nvPicPr");
		pic.appendChild(nvPicPr);
		const cNvPr = ownerDoc.createElementNS(PIC_NS, "pic:cNvPr");
		cNvPr.setAttribute("id", docPrId);
		cNvPr.setAttribute("name", altText);
		nvPicPr.appendChild(cNvPr);
		nvPicPr.appendChild(ownerDoc.createElementNS(PIC_NS, "pic:cNvPicPr"));
		const blipFill = ownerDoc.createElementNS(PIC_NS, "pic:blipFill");
		pic.appendChild(blipFill);
		const blip = ownerDoc.createElementNS(NS.a, "a:blip");
		blip.setAttributeNS(NS.r, "r:embed", rId);
		blipFill.appendChild(blip);
		const stretch = ownerDoc.createElementNS(NS.a, "a:stretch");
		stretch.appendChild(ownerDoc.createElementNS(NS.a, "a:fillRect"));
		blipFill.appendChild(stretch);
		const spPr = ownerDoc.createElementNS(PIC_NS, "pic:spPr");
		pic.appendChild(spPr);
		const xfrm = ownerDoc.createElementNS(NS.a, "a:xfrm");
		spPr.appendChild(xfrm);
		const off = ownerDoc.createElementNS(NS.a, "a:off");
		off.setAttribute("x", "0");
		off.setAttribute("y", "0");
		xfrm.appendChild(off);
		const ext = ownerDoc.createElementNS(NS.a, "a:ext");
		ext.setAttribute("cx", cx);
		ext.setAttribute("cy", cy);
		xfrm.appendChild(ext);
		const prstGeom = ownerDoc.createElementNS(NS.a, "a:prstGeom");
		prstGeom.setAttribute("prst", "rect");
		spPr.appendChild(prstGeom);
		prstGeom.appendChild(ownerDoc.createElementNS(NS.a, "a:avLst"));
		return drawing;
	}
	/** Stage every mutation (modified [Content_Types].xml, modified rels,
	* binary media files) into the writer's replacement map. No-op when
	* nothing was registered. Hyperlinks-only flushes write the rels alone;
	* image flushes additionally write content types + media binaries.
	*
	* `relsPath` defaults to the body part's rels file. Header / footer
	* registries pass their own part rels path. */
	flushTo(replacements, relsPath = "word/_rels/document.xml.rels") {
		this.partRels.flushTo(replacements, relsPath);
		for (const [path, bytes] of this.binaryAdditions) replacements.set(path, bytes);
	}
	uniqueMediaPath(ext) {
		while (true) {
			const n = this.partRels.nextImageNum();
			this.partRels.advanceImageNum(n);
			const path = `word/media/image${n}.${ext}`;
			if (!this.binaryAdditions.has(path)) return path;
		}
	}
};
function mimeForExt(ext) {
	switch (ext) {
		case "png": return "image/png";
		case "jpeg":
		case "jpg": return "image/jpeg";
		case "gif": return "image/gif";
		case "bmp": return "image/bmp";
		case "tiff":
		case "tif": return "image/tiff";
		case "svg": return "image/svg+xml";
		case "webp": return "image/webp";
		default: return `image/${ext}`;
	}
}

//#endregion
//#region lib/edit/hyperlink.ts
/**
* Hyperlink inline node — emits `<w:hyperlink>` wrapping one or more runs.
*
*   `{ link: "https://x", text: "..." }`  → `<w:hyperlink r:id="rIdN" w:history="1">`
*   `{ link: "#fig-x",   text: "..." }`  → `<w:hyperlink w:anchor="fig-x">`
*
* External links register a Relationship via the `DocxAssetRegistry` that
* owns the host part (body for normal edits, a header / footer part for
* HF content). TargetMode="External" — internal `#anchor` links use the
* w:anchor attribute directly with no rId allocation. The visible text is
* agent-supplied — Word doesn't compute it like REF/SEQ.
*
* Visual style: the run carries the `Hyperlink` character style (Word's
* blue + underlined convention). `ensureHyperlinkCharStyle` injects the
* style into styles.xml on first use when it isn't already present.
*/
const w$10 = NS.w;
const HYPERLINK_STYLE_ID = "Hyperlink";
/** Parse a link string into either an internal anchor name or an external
*  URI. Schema's refine already enforces `#`-prefix format for anchors;
*  this is a runtime split, not a re-validation. */
function parseLinkTarget(link) {
	if (link.startsWith("#")) return {
		kind: "anchor",
		name: link.slice(1)
	};
	return {
		kind: "url",
		href: link
	};
}
/** Build a `<w:hyperlink>` element containing one run carrying the visible
*  text. The run is bound to the `Hyperlink` character style; any agent-
*  supplied `format` layers on top via run rPr. */
function emitHyperlinkNode(ownerDoc, link, text, format, assetRegistry) {
	const target = parseLinkTarget(link);
	const hyper = ownerDoc.createElementNS(w$10, "w:hyperlink");
	if (target.kind === "anchor") hyper.setAttributeNS(w$10, "w:anchor", target.name);
	else {
		const { rId } = assetRegistry.registerExternalLink(target.href);
		hyper.setAttributeNS(NS.r, "r:id", rId);
		hyper.setAttributeNS(w$10, "w:history", "1");
	}
	const run = ownerDoc.createElementNS(w$10, "w:r");
	const rPr = ownerDoc.createElementNS(w$10, "w:rPr");
	const rStyle = ownerDoc.createElementNS(w$10, "w:rStyle");
	rStyle.setAttributeNS(w$10, "w:val", HYPERLINK_STYLE_ID);
	rPr.appendChild(rStyle);
	if (format) for (const c of buildRPrChildren(format, ownerDoc)) insertChildInOrder(rPr, c, RPR_CHILD_ORDER);
	run.appendChild(rPr);
	const t = ownerDoc.createElementNS(w$10, "w:t");
	t.setAttribute("xml:space", "preserve");
	t.textContent = text;
	run.appendChild(t);
	hyper.appendChild(run);
	return hyper;
}
/** Inject the `Hyperlink` character style into stylesDoc when missing. Word
*  treats `Hyperlink` as a built-in id; the canonical name "Hyperlink" and
*  the conventional 0563C1 / single-underline rPr match what Word
*  generates when you paste a URL. Idempotent — call once per apply, no-op
*  when the style already exists. */
function ensureHyperlinkCharStyle(stylesDoc) {
	const root = stylesDoc.documentElement;
	if (!root) return false;
	for (const s of getChildrenNS(root, w$10, "style")) if (wAttr(s, "styleId") === HYPERLINK_STYLE_ID) return false;
	const style = stylesDoc.createElementNS(w$10, "w:style");
	style.setAttributeNS(w$10, "w:type", "character");
	style.setAttributeNS(w$10, "w:styleId", HYPERLINK_STYLE_ID);
	const name = stylesDoc.createElementNS(w$10, "w:name");
	name.setAttributeNS(w$10, "w:val", "Hyperlink");
	style.appendChild(name);
	const basedOn = stylesDoc.createElementNS(w$10, "w:basedOn");
	basedOn.setAttributeNS(w$10, "w:val", "DefaultParagraphFont");
	style.appendChild(basedOn);
	const uiPriority = stylesDoc.createElementNS(w$10, "w:uiPriority");
	uiPriority.setAttributeNS(w$10, "w:val", "99");
	style.appendChild(uiPriority);
	style.appendChild(stylesDoc.createElementNS(w$10, "w:unhideWhenUsed"));
	const rPr = stylesDoc.createElementNS(w$10, "w:rPr");
	const color = stylesDoc.createElementNS(w$10, "w:color");
	color.setAttributeNS(w$10, "w:val", "0563C1");
	color.setAttributeNS(w$10, "w:themeColor", "hyperlink");
	rPr.appendChild(color);
	const u = stylesDoc.createElementNS(w$10, "w:u");
	u.setAttributeNS(w$10, "w:val", "single");
	rPr.appendChild(u);
	style.appendChild(rPr);
	root.appendChild(style);
	return true;
}

//#endregion
//#region lib/edit/wid-allocator.ts
/**
* Shared `w:id` allocator for the document-wide ID space.
*
* OOXML elements that carry a `w:id` attribute share a single document-scoped
* ID space: `<w:bookmarkStart>` / `<w:bookmarkEnd>`, `<w:ins>` / `<w:del>` /
* `<w:moveFrom>` / `<w:moveTo>` and their range markers, `<w:commentRangeStart>` /
* end / reference, `<w:rPrChange>` / `<w:pPrChange>` / `<w:tblPrChange>` /
* `<w:tcPrChange>` / `<w:trPrChange>` / `<w:sectPrChange>` / `<w:numPrChange>`.
*
* Allocating per-subsystem from 0 collides both with source IDs already present
* in the document and across subsystems (bookmarks vs revisions allocating the
* same number). Route every allocation through one instance per apply so the
* counter is monotonic and starts above every pre-existing ID.
*
* Scan strategy: walk every element and read the `w:id` attribute (the
* attribute IS in the `w` namespace — `getAttributeNS(w, "id")`). Scanning
* by attribute rather than enumerating tag names survives future OOXML
* extensions that introduce new `w:id`-bearing elements without code changes.
*/
const w$9 = NS.w;
/**
* Two-tier `w:id` lookup. `getAttributeNS` alone is unreliable: @xmldom/xmldom
* does not consistently namespace-resolve prefixed attributes on parsed XML,
* so `w:id="N"` may arrive with `namespaceURI === null` and qualified name
* `"w:id"`. The codebase-wide `wAttr` helper adds a third tier — bare
* `getAttribute("id")` — which is correct for most `w:*` attributes but
* **wrong here**: DrawingML wraps inside `w:drawing` carry unrelated `id`
* attributes (`<wp:docPr id="..."/>`, `<pic:cNvPr id="..."/>`) that live in
* a separate ID space from OOXML's `w:id`. Picking those up pollutes the
* counter with millions and breaks deterministic emission.
*/
function readWId(el) {
	const ns = el.getAttributeNS(w$9, "id");
	if (ns !== null && ns !== "") return ns;
	const raw = el.getAttribute("w:id");
	return raw && raw !== "" ? raw : null;
}
var WIdAllocator = class {
	nextId;
	constructor(documentDoc) {
		let max = -1;
		const root = documentDoc.documentElement;
		if (root) {
			const all = root.getElementsByTagName("*");
			for (let i = 0; i < all.length; i++) {
				const idAttr = readWId(all[i]);
				if (!idAttr) continue;
				const n = parseInt(idAttr, 10);
				if (Number.isFinite(n) && n > max) max = n;
			}
		}
		this.nextId = max + 1;
	}
	next() {
		return this.nextId++;
	}
};

//#endregion
//#region lib/edit/bookmark.ts
/**
* Bookmark allocator for cross-reference targets.
*
* OOXML bookmarks (`<w:bookmarkStart w:id="N" w:name="..."/>` and matching
* `<w:bookmarkEnd w:id="N"/>`) pair a numeric id with a string name. The id
* is unique within the document; the name is the handle REF fields use to
* resolve targets at render time. Word auto-emits these with names like
* `_Ref` + 8-digit decimal whenever the user inserts a cross-reference.
*
* Two registration paths:
*
*   1. `getOrAllocate(pEl)` — anonymous bookmark, auto-named `_Ref<8hex>`.
*      Used for InlineRefs that target a paragraph by index; the agent
*      doesn't care what the bookmark is called.
*
*   2. `adoptName(name, pEl)` — caller-supplied stable name (validated
*      against Word's bookmark name rules upstream). Used by
*      ParagraphBlock.anchor so later refs can address the paragraph by
*      name across edits[] (or across apply runs).
*
* Element-keyed: re-registering the same element returns the existing
* assignment, so an InlineRef + an anchor on the same paragraph share
* one bookmark.
*
* Two name tables:
*
*   - `nameIndex` — names bound to a paragraph element (origin `source` /
*     `allocated` / `adopted`). Resolves via `resolveByName`. Source
*     bookmarks whose `<w:bookmarkStart>` doesn't live directly inside a
*     `<w:p>` are intentionally excluded — no paragraph-level target to
*     surface for REF.
*   - `reservations` — pre-scan placeholders for names a forthcoming
*     ParagraphBlock.anchor will adopt. Carry a numbering hint so a
*     forward InlineRef can answer "target is auto-numbered?" before the
*     target element exists. `adoptName` consumes the reservation and
*     binds the name into `nameIndex`.
*
* Lazy by design: only paragraphs actually referenced (or explicitly
* named via anchor) get wrapped at `commit` time.
*/
var BookmarkAllocator = class {
	/** Document-wide `w:id` allocator, shared with TrackContext when both run
	* in the same apply so bookmark IDs and revision IDs don't collide with
	* each other or with pre-existing IDs in the source. Constructed
	* internally when not provided — single-purpose call sites (caption-only
	* pipelines, ad-hoc bookmark work) don't need to coordinate. */
	idAllocator;
	/** Tracks every name in use across all three sources: source bookmarks
	* (including non-paragraph-level ones not surfaced in `nameIndex`),
	* bound records, and pending reservations. Used for collision detection
	* in `reserveName` / `adoptName` and to seed `allocName`'s retry loop. */
	usedNames;
	byElement = /* @__PURE__ */ new Map();
	nameIndex = /* @__PURE__ */ new Map();
	reservations = /* @__PURE__ */ new Map();
	/** Names allocated via `allocateRangeBookmark`. Caption emit uses
	* these (the bookmark wraps prefix..suffix runs inline rather than
	* the commit-phase whole-paragraph wrap). InlineRef's numbering
	* check fast-paths via `isRangeBookmark` so caption-class targets
	* don't need a numPr binding. */
	rangeBookmarks = /* @__PURE__ */ new Set();
	constructor(documentDoc, idAllocator) {
		this.idAllocator = idAllocator ?? new WIdAllocator(documentDoc);
		this.usedNames = /* @__PURE__ */ new Set();
		const root = documentDoc.documentElement;
		if (root) {
			const starts = root.getElementsByTagNameNS(NS.w, "bookmarkStart");
			for (let i = 0; i < starts.length; i++) {
				const el = starts[i];
				const name = wAttr(el, "name");
				if (!name) continue;
				this.usedNames.add(name);
				const parent = el.parentNode;
				if (parent && parent.namespaceURI === NS.w && parent.localName === "p") this.nameIndex.set(name, {
					element: parent,
					origin: "source"
				});
			}
		}
	}
	/** Return the assignment for this paragraph element. Allocates lazily on
	* first call per element with an auto-generated `_Ref<8hex>` name.
	* Stable within an allocator's lifetime so multiple refs to the same
	* paragraph share one bookmark. */
	getOrAllocate(pEl) {
		const cached = this.byElement.get(pEl);
		if (cached) return cached;
		const id = this.idAllocator.next();
		const name = this.allocName();
		const assignment = {
			id,
			name
		};
		this.byElement.set(pEl, assignment);
		this.nameIndex.set(name, {
			element: pEl,
			origin: "allocated"
		});
		return assignment;
	}
	/** Pre-scan reservation: register a name that an upcoming Block will
	* adopt, so a forward InlineRef emitted earlier in the pipeline can
	* still recognize the name. Optional `ctx` captures the Block's
	* `styleId` and whether it declares `numbering` directly, so a forward
	* ref can answer "target is auto-numbered?" via `targetIsAutoNumbered`
	* before the target element exists.
	*
	* Throws on collision with a source bookmark, an existing reservation,
	* or an already-adopted name — the engine runs this once per declared
	* anchor before emit starts, so duplicates surface before any mutation. */
	reserveName(name, ctx) {
		if (this.usedNames.has(name)) throw new Error(`anchor "${name}" ${this.describeCollision(name)}. Pick a unique anchor name.`);
		this.usedNames.add(name);
		this.reservations.set(name, {
			styleId: ctx?.styleId,
			directlyNumbered: ctx?.directlyNumbered,
			isCaption: ctx?.isCaption
		});
	}
	/** True iff `name` has been pre-scan reserved but not yet bound to an
	* element via `adoptName`. Lets the emitter branch into forward-ref
	* handling (style-cascade numbering check, late-bound backfill). */
	isReserved(name) {
		return this.reservations.has(name);
	}
	/** Numbering hint captured at reservation time. Returns undefined when
	* the name isn't reserved. */
	predictedNumberingFor(name) {
		return this.reservations.get(name);
	}
	/** Adopt a caller-supplied name for this paragraph (the `anchor` field
	* on a ParagraphBlock). Two failure modes:
	*
	*   - `name` collides with an existing source bookmark or with a name
	*     already adopted in this run → throws (the agent picked a name
	*     that's not unique).
	*   - The same `pEl` is being re-bound to a different name → throws (a
	*     paragraph carries at most one anchor; if you need a second handle,
	*     refTo it by paragraph index instead).
	*
	* If the name was pre-scan reserved, the reservation is consumed and
	* the name is bound to `pEl` — not a duplicate.
	*
	* Idempotent for the same `(name, pEl)` pair. */
	adoptName(name, pEl) {
		const existing = this.byElement.get(pEl);
		if (existing) {
			if (existing.name === name) return existing;
			throw new Error(`paragraph already declared anchor "${existing.name}" — cannot also declare anchor "${name}". Each paragraph carries at most one anchor.`);
		}
		if (this.nameIndex.has(name)) throw new Error(`anchor "${name}" ${this.describeCollision(name)}. Pick a unique anchor name.`);
		this.reservations.delete(name);
		const id = this.idAllocator.next();
		this.usedNames.add(name);
		const assignment = {
			id,
			name
		};
		this.byElement.set(pEl, assignment);
		this.nameIndex.set(name, {
			element: pEl,
			origin: "adopted"
		});
		return assignment;
	}
	/** Resolve a bookmark name to its paragraph element. Returns undefined
	* when the name isn't bound (either unknown or only pre-scan reserved —
	* caller branches via `isReserved` for the latter). */
	resolveByName(name) {
		return this.nameIndex.get(name);
	}
	/** Reserve a name + id for an inline (caption-emitter) bookmark
	* without binding to any element yet. Caption-emit needs the id+name
	* upfront to emit the bookmarkStart/End XML, but the paragraph
	* doesn't exist yet — it's about to be built around the bookmark.
	*
	* Same collision semantics as adoptName: source / prior-adopted name
	* → throw; reserved name → consume the reservation. The returned
	* assignment is NOT yet in nameIndex; call `bindRangeBookmark` after
	* the caption paragraph is constructed to record the binding so
	* `resolveByName` works for REF cross-references. */
	allocateRangeBookmark(name) {
		if (this.nameIndex.has(name)) throw new Error(`anchor "${name}" ${this.describeCollision(name)}. Pick a unique anchor name.`);
		if (this.usedNames.has(name) && !this.reservations.has(name)) throw new Error(`anchor "${name}" ${this.describeCollision(name)}. Pick a unique anchor name.`);
		this.reservations.delete(name);
		const id = this.idAllocator.next();
		this.usedNames.add(name);
		this.rangeBookmarks.add(name);
		return {
			id,
			name
		};
	}
	/** Post-allocation binding for `allocateRangeBookmark`. Records the
	* name → paragraph mapping in `nameIndex` so REF backfill can resolve
	* the target. Does NOT touch `byElement` — caption-emit already
	* emitted bookmarkStart/End inline, so commit() must NOT wrap the
	* paragraph again. */
	bindRangeBookmark(name, pEl) {
		this.nameIndex.set(name, {
			element: pEl,
			origin: "adopted"
		});
	}
	/** True iff `name` was allocated as a caption range bookmark. Used by
	* InlineRef's emit-time numbering check to fast-path caption-class
	* targets (they're SEQ-numbered, not numPr-bound). */
	isRangeBookmark(name) {
		return this.rangeBookmarks.has(name);
	}
	/** True iff at least one bookmark will be wrapped at commit. Source
	* bookmarks don't count — they're already in the XML. */
	hasAllocations() {
		return this.byElement.size > 0;
	}
	/** Wrap each allocated / adopted paragraph with `<w:bookmarkStart>` /
	* `<w:bookmarkEnd>`. Start goes immediately after `<w:pPr>` (or at the
	* front when no pPr); end goes at the paragraph's tail. Idempotent:
	* subsequent calls do nothing because the internal queue is consumed.
	* Source bookmarks are skipped — they already exist in document.xml. */
	commit(documentDoc) {
		const w = NS.w;
		for (const [pEl, { id, name }] of this.byElement) {
			const start = documentDoc.createElementNS(w, "w:bookmarkStart");
			start.setAttributeNS(w, "w:id", String(id));
			start.setAttributeNS(w, "w:name", name);
			const end = documentDoc.createElementNS(w, "w:bookmarkEnd");
			end.setAttributeNS(w, "w:id", String(id));
			const pPr = getChildren(pEl).find((c) => c.namespaceURI === w && c.localName === "pPr");
			if (pPr && pPr.nextSibling) pEl.insertBefore(start, pPr.nextSibling);
			else if (pPr) pEl.appendChild(start);
			else if (pEl.firstChild) pEl.insertBefore(start, pEl.firstChild);
			else pEl.appendChild(start);
			pEl.appendChild(end);
		}
		this.byElement.clear();
	}
	/** Describe why `name` collides for the error message. Caller guarantees
	* the name is in `usedNames`; this just routes to the right phrase. */
	describeCollision(name) {
		const bound = this.nameIndex.get(name);
		if (bound) return bound.origin === "source" ? "exists in the source document" : "was already adopted in this apply";
		if (this.reservations.has(name)) return "is already reserved by an earlier ParagraphBlock.anchor in this apply";
		return "exists in the source document";
	}
	/** Word convention: `_Ref` + 8 decimal digits, allocator-local counter.
	* Re-rolls on collision with a pre-existing name. The numeric search
	* space is 10^8, so collision is statistically impossible in normal
	* use; the retry is defensive only. */
	allocName() {
		let attempt = 0;
		while (attempt < 1e3) {
			const name = `_Ref${Math.floor(Math.random() * 1e8).toString().padStart(8, "0")}`;
			if (!this.usedNames.has(name)) {
				this.usedNames.add(name);
				return name;
			}
			attempt++;
		}
		throw new Error("BookmarkAllocator.allocName: could not find a unique bookmark name after 1000 retries. This is an engine invariant violation; the 10^8 search space should make collision statistically impossible.");
	}
};

//#endregion
//#region lib/edit/fields/ref-field.ts
/**
* REF field emitter — thin layer over the shared complex-field skeleton.
*
* REF resolves a bookmark at field-update time and renders the result
* per its switches (`\n` lvlText / `\r` paragraph number / `\h` hyperlink
* / etc.). The placeholder run carries the user-visible text Word
* displays BEFORE update; after F9 (or auto-update on open via
* `updateFields=true` in settings.xml) Word replaces it with the
* bookmark's resolved content. Computing the placeholder accurately at
* emit time means the document reads correctly even if a user opens it
* without triggering field update.
*
* Format preservation (rPr replication + MERGEFORMAT) is handled by the
* shared skeleton — see `complex-field.ts`.
*/
/** Emit the 5-run REF sequence. Returns the runs plus the result text
* element so callers can backfill the placeholder after counter
* simulation. */
function emitRefField(ownerDoc, spec) {
	const switchPart = spec.switches.length > 0 ? ` ${spec.switches.join(" ")}` : "";
	const { runs, resultTextEl } = emitComplexField(ownerDoc, {
		instrCode: `REF ${spec.bookmarkName}${switchPart}`,
		initialResult: spec.placeholder,
		format: spec.format
	});
	return {
		runs,
		resultTextEl
	};
}
/** Map a `display` value to the corresponding REF switch list.
*   - "label"  → `\n` (full numbered paragraph text from lvlText,
*                e.g. "图 1" or "1.2.3") + `\h` (hyperlink)
*   - "number" → `\r` (paragraph number in relative context — for single-
*                level schemes equals the counter "1"; for multi-level
*                returns relative position, not the leaf digit) + `\h`
*   - "full"   → no number switches (bookmark text content, e.g. caption
*                title) + `\h`
*
* `\h` is always added so the rendered text is clickable in Word — matches
* what Insert → Cross-reference produces. */
function switchesForDisplay(display) {
	switch (display) {
		case "label": return ["\\n", "\\h"];
		case "number": return ["\\r", "\\h"];
		case "full": return ["\\h"];
	}
}

//#endregion
//#region lib/edit/edit-caption-op.ts
/**
* `edit-caption` op handler.
*
* Replaces caption body text — runs after the primary anchor's
* `<w:bookmarkEnd>`, preserving the fields (SEQ / STYLEREF) and bookmark
* pair so cross-references continue to resolve.
*
* Target resolution:
*   - `{ anchor: name }`: lookup via BookmarkAllocator. The bookmark must
*     have been allocated by an earlier caption emit (range bookmark).
*   - `{ captionId, index }`: walk body in order; count paragraphs that
*     (a) have paragraphStyleId matching captions[<id>].styleId AND
*     (b) contain a SEQ field with identifier matching `<id>`. Pick the
*     1-based Nth.
*
* Throws when:
*   - target paragraph has no `bookmarkEnd` (caption was emitted without
*     an anchor — engine can't locate the body boundary)
*   - resolution fails (unknown anchor / out-of-range index)
*/
const w$8 = NS.w;
const m = NS.m;
/** Resolve target to a caption paragraph element. Throws on miss. */
function resolveEditCaptionTarget(input) {
	let para;
	if ("anchor" in input.target) {
		const rec = input.bookmarkAllocator.resolveByName(input.target.anchor);
		if (!rec) throw new Error(`edit-caption: anchor "${input.target.anchor}" was not found. Declare it on the caption block before editing, or pick a different target.`);
		para = rec.element;
	} else {
		const { captionId, index } = input.target;
		const config = input.captionsConfigs.get(captionId);
		if (!config) throw new Error(`edit-caption: captionId "${captionId}" is not declared in captions table.`);
		const matches = [];
		const body = firstChildNS(input.documentDoc.documentElement, w$8, "body");
		if (!body) throw new Error("edit-caption: document has no body");
		for (const p of walkBodyParagraphs(body)) {
			if (paragraphStyleId(p) !== config.paragraphStyleId) continue;
			if (!paragraphContainsSeq(p, captionId)) continue;
			matches.push(p);
		}
		if (index < 1 || index > matches.length) throw new Error(`edit-caption: captionId "${captionId}" index ${index} out of range. Document has ${matches.length} caption paragraph(s) of this identifier.`);
		para = matches[index - 1];
	}
	if (isEquationCaptionParagraph(para)) throw new Error("edit-caption: target paragraph is an EquationBlock's caption (number-only, no body text). To change the equation, delete the EquationBlock and re-emit with the updated LaTeX.");
	return para;
}
/** Detect EquationBlock caption paragraphs. They live in a 3-col borderless
* table's right cell; the middle cell carries an `<m:oMath>` / `<m:oMathPara>`
* subtree. Walk up to the enclosing `<w:tc>`, then to `<w:tr>`, and scan the
* sibling cells for OMML content. */
function isEquationCaptionParagraph(paragraph) {
	let cell = paragraph.parentNode;
	while (cell && !(cell.namespaceURI === w$8 && cell.localName === "tc")) cell = cell.parentNode;
	if (!cell) return false;
	const row = cell.parentNode;
	if (!row || row.namespaceURI !== w$8 || row.localName !== "tr") return false;
	for (const sibling of getChildren(row)) {
		if (sibling === cell) continue;
		if (sibling.namespaceURI !== w$8 || sibling.localName !== "tc") continue;
		if (descendantsNS(sibling, m, "oMath").length > 0 || descendantsNS(sibling, m, "oMathPara").length > 0) return true;
	}
	return false;
}
/** Replace the caption paragraph's body text — runs after the primary
* `<w:bookmarkEnd>`. Existing body runs are removed; a new bodySeparator
* + text pair is appended. Preserves bookmark and field structure. */
function applyEditCaption(paragraph, text, config, ownerDoc) {
	const children = getChildren(paragraph);
	let boundaryIdx = -1;
	for (let i = children.length - 1; i >= 0; i--) {
		const c = children[i];
		if (c.namespaceURI !== w$8) continue;
		if (c.localName === "bookmarkEnd") {
			boundaryIdx = i;
			break;
		}
	}
	if (boundaryIdx < 0) for (let i = children.length - 1; i >= 0; i--) {
		const c = children[i];
		if (c.namespaceURI !== w$8 || c.localName !== "r") continue;
		const fld = firstChildNS(c, w$8, "fldChar");
		if (fld && wAttr(fld, "fldCharType") === "end") {
			boundaryIdx = i;
			break;
		}
	}
	if (boundaryIdx < 0) throw new Error("edit-caption: target paragraph has no field or bookmark boundary — can't locate the body region. Was this paragraph emitted via the caption pipeline?");
	for (let i = children.length - 1; i > boundaryIdx; i--) paragraph.removeChild(children[i]);
	if (text !== "") {
		paragraph.appendChild(buildPlainTextRun(ownerDoc, config.bodySeparator));
		paragraph.appendChild(buildPlainTextRun(ownerDoc, text));
	}
}
function paragraphContainsSeq(paragraph, identifier) {
	const parsed = parseFieldRuns(paragraphRuns(paragraph));
	for (const entry of parsed) if (entry.kind === "field" && entry.fieldType === "SEQ") {
		if (entry.details.identifier === identifier) return true;
	}
	return false;
}

//#endregion
//#region lib/edit/edit-engine.ts
const w$7 = NS.w;
/**
* Resolve all edit locators against the (pre-edits) document and report what
* each op will touch — without mutating. Used by `apply-styles --dry-run` so
* the change report can show predicted edit effect alongside style /
* pattern_rule effects.
*
* Throws on locator-resolution errors and on blocker conflicts (same checks
* the real apply path does pre-mutation).
*/
function previewEditOps(input) {
	const { documentDoc, parsedParagraphs, edits } = input;
	const resolverCtx = buildResolverContext(documentDoc, parsedParagraphs);
	const blockers = detectBlockers(documentDoc, resolverCtx.indexByElement);
	const resolved = [];
	for (const [i, op] of edits.entries()) try {
		resolved.push(resolveOneEdit(op, resolverCtx, i));
	} catch (err) {
		throw new Error(`edits[${i}] (${op.op}): ${err.message}`, { cause: err });
	}
	validateAgainstBlockers(resolved, blockers, resolverCtx.indexByElement);
	const replacedOrDeletedIndices = /* @__PURE__ */ new Set();
	const entries = [];
	for (const [i, edit] of resolved.entries()) {
		const targetParaIndices = edit.target.paragraphs.map((p) => resolverCtx.indexByElement.get(p) ?? -1);
		const op = edit.op.op;
		let willReplaceOrDeleteIndices = [];
		let willInsertCount = 0;
		let survivorIndex;
		if (op === "replace") {
			willReplaceOrDeleteIndices = [...targetParaIndices];
			willInsertCount = edit.op.with.length;
			for (const idx of targetParaIndices) if (idx >= 0) replacedOrDeletedIndices.add(idx);
		} else if (op === "delete") {
			willReplaceOrDeleteIndices = [...targetParaIndices];
			for (const idx of targetParaIndices) if (idx >= 0) replacedOrDeletedIndices.add(idx);
		} else if (op === "merge") {
			willReplaceOrDeleteIndices = [...targetParaIndices];
			for (const idx of targetParaIndices) if (idx >= 0) replacedOrDeletedIndices.add(idx);
			const keepPPr = edit.op.keepPPr ?? "first";
			const sortedTargets = [...targetParaIndices].sort((a, b) => a - b);
			const survivorIdx = keepPPr === "last" ? sortedTargets[sortedTargets.length - 1] : sortedTargets[0];
			if (survivorIdx !== void 0 && survivorIdx >= 0) {
				survivorIndex = survivorIdx;
				replacedOrDeletedIndices.delete(survivorIdx);
			}
		} else if (op === "insert-before" || op === "insert-after") willInsertCount = edit.op.content.length;
		const container = edit.target.container.namespaceURI === w$7 && edit.target.container.localName === "tc" ? "cell" : "body";
		entries.push({
			index: i,
			op,
			targetParaIndices,
			willReplaceOrDeleteIndices,
			willInsertCount,
			container,
			survivorIndex
		});
	}
	return {
		entries,
		replacedOrDeletedIndices
	};
}
/**
* In-memory edit-ops application — for callers that already have the docx
* open and parsed (e.g. `apply` integrating edits into its single
* pipeline). The caller owns file I/O, validation, and writing.
*
* Mutates `documentDoc` in place. Returns the image registry (caller
* flushes its staged binaries / rels / content-type updates into its
* own replacement map) plus the applied-ops report.
*/
async function runEditOps(input) {
	const { documentDoc, parsedParagraphs, reader, edits, trackChanges, author, stylesDoc, numberingDoc } = input;
	const resolverCtx = buildResolverContext(documentDoc, parsedParagraphs);
	const blockers = detectBlockers(documentDoc, resolverCtx.indexByElement);
	const resolved = [];
	for (const [i, op] of edits.entries()) try {
		resolved.push(resolveOneEdit(op, resolverCtx, i));
	} catch (err) {
		throw new Error(`edits[${i}] (${op.op}): ${err.message}`, { cause: err });
	}
	validateAgainstBlockers(resolved, blockers, resolverCtx.indexByElement);
	if (trackChanges) for (const [i, op] of edits.entries()) {
		if (op.op === "merge") throw new Error(`edits[${i}] (merge): merge under trackChanges=true is not supported. Run merge in a separate apply without trackChanges.`);
		const frag = op.op === "replace" ? op.with : op.op === "insert-before" || op.op === "insert-after" ? op.content : null;
		if (!frag) continue;
		if (fragmentContainsTable(frag)) throw new Error(`edits[${i}] (${op.op}): inserting a TableBlock under trackChanges=true is not supported. Run table insertion in a separate apply without trackChanges, then use trackChanges for subsequent cell edits.`);
		if (fragmentContainsEquation(frag)) throw new Error(`edits[${i}] (${op.op}): inserting an equation (block or inline) under trackChanges=true is not supported. Run equation insertion in a separate apply without trackChanges, then use trackChanges for surrounding edits.`);
	}
	for (const [i, op] of edits.entries()) {
		const frag = fragmentOf(op);
		if (!frag) continue;
		const items = [];
		for (const b of frag) items.push(...collectLatexFromBlock(b));
		if (items.length === 0) continue;
		try {
			await prepareLatex(items);
		} catch (err) {
			throw new Error(`edits[${i}] (${op.op}): ${err.message}`, { cause: err });
		}
	}
	const idAllocator = new WIdAllocator(documentDoc);
	const trackContext = makeTrackContext(trackChanges, idAllocator, { author });
	const stale = /* @__PURE__ */ new Set();
	const imageRegistry = input.imageRegistry ?? await DocxAssetRegistry.open(reader);
	const bookmarkAllocator = new BookmarkAllocator(documentDoc, idAllocator);
	for (const op of edits) {
		const fragment = fragmentOf(op);
		if (!fragment) continue;
		walkBlocksForAnchors(fragment, (hint) => {
			bookmarkAllocator.reserveName(hint.anchor, {
				styleId: hint.styleId,
				directlyNumbered: hint.directlyNumbered,
				isCaption: hint.isCaption
			});
		});
	}
	const pendingBackfills = [];
	const pendingCaptionFills = [];
	const pendingCaptionResets = [];
	const captionsMap = input.captions;
	const emitCtx = {
		emitImage: (src, width, height, alt, ownerDoc) => {
			const { rId } = imageRegistry.registerImage(src);
			return imageRegistry.buildDrawing(rId, width, height, alt, ownerDoc);
		},
		resolveStyle: buildStyleResolver(stylesDoc ?? null),
		emitHyperlink: (link, text, format, ownerDoc) => {
			if (stylesDoc) ensureHyperlinkCharStyle(stylesDoc);
			return emitHyperlinkNode(ownerDoc, link, text, format, imageRegistry);
		},
		captions: captionsMap ? {
			resolve: (identifier) => captionsMap.get(identifier),
			allocateBookmark: (name) => bookmarkAllocator.allocateRangeBookmark(name),
			bindBookmark: (name, pEl) => bookmarkAllocator.bindRangeBookmark(name, pEl),
			registerFill: (fill) => pendingCaptionFills.push(fill),
			registerReset: (reset) => pendingCaptionResets.push(reset)
		} : void 0,
		emitRef: (ref, ownerDoc, defaultFormat) => {
			let targetEl;
			let bookmarkName;
			if (ref.refTo.type === "paragraph") {
				const target = resolverCtx.indexed[ref.refTo.index - 1];
				if (!target || target.index !== ref.refTo.index) throw new Error(`InlineRef: refTo.paragraph=${ref.refTo.index} is out of range. Document has ${resolverCtx.indexed.length} indexed paragraphs.`);
				targetEl = target.element;
				bookmarkName = bookmarkAllocator.getOrAllocate(targetEl).name;
			} else {
				const found = bookmarkAllocator.resolveByName(ref.refTo.name);
				const reserved = !found && bookmarkAllocator.isReserved(ref.refTo.name);
				if (!found && !reserved) throw new Error(`InlineRef: refTo.anchor="${ref.refTo.name}" was not found. Declare a ParagraphBlock.anchor (or EquationBlock.anchor) with this name somewhere in edits[], or reference an existing bookmark on a paragraph in the source document.`);
				targetEl = found?.element ?? null;
				bookmarkName = ref.refTo.name;
			}
			const display = ref.display ?? "label";
			const isCaptionAnchor = ref.refTo.type === "anchor" && (bookmarkAllocator.isRangeBookmark(ref.refTo.name) || bookmarkAllocator.predictedNumberingFor(ref.refTo.name)?.isCaption === true);
			if (display === "full" && isCaptionAnchor) throw new Error(`InlineRef: display="full" is not supported on caption-class anchors (anchor "${ref.refTo.name}"). Caption refs use the SEQ-rendered text (prefix + chapter + counter + suffix); switch to display: "label".`);
			if (display !== "full" && !isCaptionAnchor) {
				const predicted = ref.refTo.type === "anchor" ? bookmarkAllocator.predictedNumberingFor(ref.refTo.name) : void 0;
				if (!targetIsAutoNumbered({
					element: targetEl,
					styleId: predicted?.styleId,
					directlyNumbered: predicted?.directlyNumbered
				}, stylesDoc ?? null)) {
					const where = ref.refTo.type === "paragraph" ? `target paragraph #${ref.refTo.index}` : `target of anchor "${ref.refTo.name}"`;
					throw new Error(`InlineRef: ${where} is not bound to a numbering scheme. display="${display}" requires an auto-numbered target (Word's \\n / \\r switches render from the numbering binding). Fix one of: bind the target's pStyle to a numbering[] level; for a caption-class target use a CaptionBlock or EquationBlock with captionId (a *Caption styleId on a plain paragraph is not a caption); or use display: "full" for plain body text.`);
				}
			}
			const switches = isCaptionAnchor ? ["\\h"] : switchesForDisplay(display);
			const { runs, resultTextEl } = emitRefField(ownerDoc, {
				bookmarkName,
				switches,
				placeholder: "",
				format: ref.format ?? defaultFormat
			});
			pendingBackfills.push({
				placeholderTextEl: resultTextEl,
				targetName: bookmarkName,
				display
			});
			return runs;
		},
		adoptAnchor: (name, pEl) => {
			bookmarkAllocator.adoptName(name, pEl);
		},
		forkNumRestart: numberingDoc ? (pEl, numId, level) => applyParagraphLevelRestart(pEl, numberingDoc, numId, level) : void 0
	};
	const perOp = [];
	for (const [i, edit] of resolved.entries()) {
		for (const p of edit.target.paragraphs) if (stale.has(p)) throw new Error(`edits[${i}] (${edit.op.op}): targets a paragraph that was removed by an earlier edit. Reorder edits, or split into separate apply_edits runs so each pass sees a fresh document.`);
		const perOpCtx = {
			...emitCtx,
			usableWidthTwips: usableWidthForTarget(edit.target, input.sections, resolverCtx)
		};
		let touched;
		try {
			touched = applyOne(edit, documentDoc, trackContext, perOpCtx, stale, resolverCtx, {
				bookmarkAllocator,
				captionsMap: input.captions,
				blockersByElement: blockers.byElement
			});
		} catch (err) {
			const msg = err.message;
			if (msg.startsWith(`edits[${i}]`)) throw err;
			throw new Error(`edits[${i}] (${edit.op.op}): ${msg}`, { cause: err });
		}
		perOp.push({
			index: i,
			op: edit.op.op,
			touched
		});
	}
	return {
		imageRegistry,
		report: {
			applied: resolved.length,
			trackChanges: trackContext.enabled,
			blockerCounts: summarizeBlockers(blockers),
			perOp
		},
		crossRefs: {
			bookmarkAllocator,
			pendingCaptionFills,
			pendingCaptionResets,
			pendingBackfills
		}
	};
}
/** Single auto-numbering check for InlineRef targets. The target is
* represented either as a live paragraph element (already emitted — direct
* `<w:p>/<w:pPr>/<w:numPr>` and pStyle cascade both observable) or as a
* pre-scan hint (forward ref, target not yet emitted — `directlyNumbered`
* mirrors the Block's direct `numbering` field, `styleId` runs the cascade
* walk). Both representations now cover the same fact set; forward refs
* are no longer blind to direct-numbering targets. */
function targetIsAutoNumbered(spec, stylesDoc) {
	if (spec.element) {
		const pPr = firstChildNS(spec.element, w$7, "pPr");
		if (pPr) {
			if (firstChildNS(pPr, w$7, "numPr")) return true;
			const pStyle = firstChildNS(pPr, w$7, "pStyle");
			if (pStyle && stylesDoc) {
				const styleId = pStyle.getAttributeNS(w$7, "val") ?? pStyle.getAttribute("w:val") ?? "";
				if (styleId && styleHasNumPrInCascade(stylesDoc, styleId)) return true;
			}
		}
		return false;
	}
	if (spec.directlyNumbered) return true;
	if (spec.styleId && stylesDoc && styleHasNumPrInCascade(stylesDoc, spec.styleId)) return true;
	return false;
}
function styleHasNumPrInCascade(stylesDoc, styleId) {
	const seen = /* @__PURE__ */ new Set();
	let current = styleId;
	while (current && !seen.has(current)) {
		seen.add(current);
		const styleEl = findStyleById(stylesDoc, current);
		if (!styleEl) return false;
		const sPr = firstChildNS(styleEl, w$7, "pPr");
		if (sPr && firstChildNS(sPr, w$7, "numPr")) return true;
		const basedOn = firstChildNS(styleEl, w$7, "basedOn");
		current = basedOn ? basedOn.getAttributeNS(w$7, "val") ?? basedOn.getAttribute("w:val") : null;
	}
	return false;
}
function findStyleById(stylesDoc, id) {
	const root = stylesDoc.documentElement;
	if (!root) return null;
	for (const s of getChildrenNS(root, w$7, "style")) if ((s.getAttributeNS(w$7, "styleId") ?? s.getAttribute("w:styleId")) === id) return s;
	return null;
}
/** Resolve one op's locator. Single source of truth used by both
* `previewEditOps` (dry-run path) and `runEditOps` — keeps set-run's
* RunLocator special case in one place. */
function resolveOneEdit(op, resolverCtx, _opIndex) {
	if (op.op === "set-run") {
		const r = resolveRunLocator(op.at, resolverCtx);
		return {
			op,
			target: {
				paragraphs: [r.paragraph],
				container: resolverCtx.body
			},
			runRef: r.run
		};
	}
	if (op.op === "edit-caption") return {
		op,
		target: {
			paragraphs: [],
			container: resolverCtx.body
		}
	};
	return {
		op,
		target: resolveLocator(op.at, resolverCtx)
	};
}
/** Resolve the usable content width (twips) for the section the op's target
* lives in. Returns `undefined` for cell-internal targets (we don't read
* tcW in v1; emitters fall back to a constant) and when no section info
* is supplied or no section claims the target's paragraph index. */
function usableWidthForTarget(target, sections, resolverCtx) {
	if (!sections || sections.length === 0) return void 0;
	if (target.container.namespaceURI === w$7 && target.container.localName === "tc") return void 0;
	const firstP = target.paragraphs[0];
	if (firstP === void 0) {
		const last = sections[sections.length - 1];
		if (!last) return void 0;
		const width = sectionUsableWidthTwips(last);
		return width > 0 ? width : void 0;
	}
	const idx = resolverCtx.indexByElement.get(firstP);
	if (idx === void 0) return void 0;
	const sec = sectionForParagraph(sections, idx);
	if (!sec) return void 0;
	const width = sectionUsableWidthTwips(sec);
	return width > 0 ? width : void 0;
}
function validateAgainstBlockers(resolved, blockers, indexByElement) {
	const failures = [];
	for (const [i, edit] of resolved.entries()) {
		const overwriteFields = edit.op.op === "replace" && edit.op.overwriteFields === true;
		for (const p of edit.target.paragraphs) {
			const reason = blockers.byElement.get(p);
			if (reason) {
				if (overwriteFields && reason === "field") continue;
				const idx = indexByElement.get(p);
				const where = idx !== void 0 ? `paragraph #${idx}` : `cell paragraph`;
				failures.push(`edits[${i}] (${edit.op.op}): ${where} is blocked — ${explainBlockerReason(reason)}`);
			}
		}
	}
	if (failures.length > 0) throw new Error(`${failures.length} edit(s) targeted blocked paragraphs:\n  ${failures.join("\n  ")}`);
}
/** Detect any TableBlock anywhere in a fragment (handles direct table
* blocks; doesn't recurse into cell Block[] content because the schema
* already forbids nested tables in cells). */
function fragmentContainsTable(fragment) {
	for (const b of fragment) if (b.type === "table") return true;
	return false;
}
function fragmentContainsEquation(fragment) {
	for (const b of fragment) if (collectLatexFromBlock(b).length > 0) return true;
	return false;
}
function collectLatexFromBlock(b) {
	const out = [];
	if (b.type === "equation") {
		if (b.latex !== void 0) out.push({
			latex: b.latex,
			displayMode: true
		});
	} else if (b.type === "paragraph") {
		if (Array.isArray(b.text)) {
			for (const piece of b.text) if ("math" in piece) out.push({
				latex: piece.math,
				displayMode: false
			});
		}
	} else if (b.type === "table") for (const row of b.rows) for (const cell of row) {
		let cellContent;
		if (typeof cell === "string" || Array.isArray(cell)) cellContent = cell;
		else if (cell && typeof cell === "object" && "content" in cell) cellContent = cell.content;
		if (Array.isArray(cellContent)) {
			for (const piece of cellContent) if (piece && typeof piece === "object") {
				if ("math" in piece && typeof piece.math === "string") out.push({
					latex: piece.math,
					displayMode: false
				});
				else if ("type" in piece) out.push(...collectLatexFromBlock(piece));
			}
		}
	}
	return out;
}
/** Return the inserted/replaced Block[] for an op, or null when the op has
* no fragment (delete / format / set-run). Centralizes the
* `with` / `content` discriminant used by the latex collector and the
* anchor pre-scan. */
function fragmentOf(op) {
	if (op.op === "replace") return op.with;
	if (op.op === "insert-before" || op.op === "insert-after") return op.content;
	return null;
}
function walkBlocksForAnchors(fragment, visit) {
	for (const block of fragment) walkBlockForAnchors(block, visit);
}
function walkBlockForAnchors(block, visit) {
	const b = block;
	if (typeof b.anchor === "string" && b.anchor) {
		const isCaption = !!b.captionId;
		visit({
			anchor: b.anchor,
			styleId: b.styleId,
			directlyNumbered: !!b.numbering || isCaption,
			isCaption
		});
	}
	if (block.type === "table") for (const row of block.rows) for (const cell of row) {
		let cellContent;
		if (typeof cell === "string" || Array.isArray(cell)) cellContent = cell;
		else if (cell && typeof cell === "object" && "content" in cell) cellContent = cell.content;
		if (!Array.isArray(cellContent)) continue;
		for (const piece of cellContent) if (piece && typeof piece === "object" && "type" in piece) walkBlockForAnchors(piece, visit);
	}
}
function applyEditCaptionOp(op, documentDoc, deps) {
	if (!deps.captionsMap) throw new Error("edit-caption: captions table not declared in apply config. Add the identifier to `captions`.");
	const para = resolveEditCaptionTarget({
		documentDoc,
		target: op.target,
		text: op.text,
		bookmarkAllocator: deps.bookmarkAllocator,
		captionsConfigs: deps.captionsMap
	});
	let config;
	if ("captionId" in op.target) config = deps.captionsMap.get(op.target.captionId);
	else {
		const pPr = firstChildNS(para, w$7, "pPr");
		const pStyle = pPr ? firstChildNS(pPr, w$7, "pStyle") : null;
		const styleId = pStyle ? pStyle.getAttributeNS(w$7, "val") ?? null : null;
		if (styleId) {
			for (const c of deps.captionsMap.values()) if (c.paragraphStyleId === styleId) {
				config = c;
				break;
			}
		}
	}
	if (!config) throw new Error("edit-caption: could not resolve caption config for the target paragraph. The paragraph's pStyle doesn't match any captions[<id>].styleId.");
	applyEditCaption(para, op.text, config, documentDoc);
	return 1;
}
function applyOne(edit, documentDoc, trackContext, emitCtx, stale, resolverCtx, deps) {
	switch (edit.op.op) {
		case "replace": return applyReplace(edit.target, edit.op.with, documentDoc, trackContext, emitCtx, stale);
		case "insert-before": return applyInsertBefore(edit.target, edit.op.content, documentDoc, trackContext, emitCtx, resolverCtx);
		case "insert-after": return applyInsertAfter(edit.target, edit.op.content, documentDoc, trackContext, emitCtx, resolverCtx);
		case "delete": return applyDelete(edit.target, documentDoc, trackContext, stale);
		case "format": return applyFormat(edit.target, edit.op, documentDoc, trackContext);
		case "set-run":
			if (!edit.runRef) throw new Error("set-run: missing resolved run reference (internal)");
			return applySetRun(edit.runRef, edit.op, documentDoc, trackContext);
		case "edit-caption": return applyEditCaptionOp(edit.op, documentDoc, deps);
		case "merge": return applyMerge(edit.target, edit.op, stale, deps.blockersByElement);
		default: return assertNever(edit.op);
	}
}
function isParagraphElement(el) {
	return el.namespaceURI === w$7 && el.localName === "p";
}
/** pPr fields where an explicit styleId on the new Block is meant to
* govern. When the engine inherits anchor pPr via MDF, these fields
* are stripped so the styleId cascade reaches them. Without this strip,
* a direct <w:outlineLvl w:val="1"/> on the anchor would override the
* Heading 1 outline level cascaded from ProposalH1, mis-tagging the
* new paragraph as Heading 2. */
const STYLE_GOVERNED_PPR_FIELDS = new Set([
	"outlineLvl",
	"numPr",
	"pageBreakBefore",
	"widowControl",
	"spacing"
]);
function inheritPPrFromAnchor(newP, anchor, ownerDoc, blockHasExplicitStyleId) {
	const anchorPPr = firstChildNS(anchor, w$7, "pPr");
	if (!anchorPPr) return;
	let newPPr = firstChildNS(newP, w$7, "pPr");
	const existingByName = /* @__PURE__ */ new Map();
	if (newPPr) {
		for (const c of getChildren(newPPr)) if (c.namespaceURI === w$7) existingByName.set(c.localName, c);
	}
	const newHasPStyle = !!(newPPr && firstChildNS(newPPr, w$7, "pStyle"));
	const toClone = [];
	for (const c of getChildren(anchorPPr)) {
		if (c.namespaceURI !== w$7) continue;
		if (c.localName === "pPrChange") continue;
		if (c.localName === "rPr" && newHasPStyle) continue;
		if (blockHasExplicitStyleId && STYLE_GOVERNED_PPR_FIELDS.has(c.localName)) continue;
		const existing = existingByName.get(c.localName);
		if (existing) mergeMissingAttrs(c, existing);
		else toClone.push(c.cloneNode(true));
	}
	if (toClone.length === 0) return;
	if (!newPPr) {
		newPPr = ownerDoc.createElementNS(w$7, "w:pPr");
		newP.insertBefore(newPPr, newP.firstChild);
	}
	for (const c of toClone) insertChildInOrder(newPPr, c, PPR_CHILD_ORDER);
}
/** Inherit the anchor's first non-empty run rPr into the new paragraph's
* runs when the Block has no explicit runFormat. This implements the
* run-side of MDF: new paragraphs adopt the anchor's typeface so they
* blend visually into the destination context rather than falling back to
* the docDefault font (Calibri / 宋体). Only fields not already declared
* on the run's rPr are inherited — explicit runFormat choices always win. */
function inheritAnchorRunRPr(newRunRPr, anchorEl, newRunHasExplicitFormat) {
	if (newRunHasExplicitFormat) return;
	const firstNonEmpty = getChildrenNS(anchorEl, w$7, "r").find((r) => textContent(r).trim().length > 0);
	if (!firstNonEmpty) return;
	const anchorRPr = firstChildNS(firstNonEmpty, w$7, "rPr");
	if (!anchorRPr) return;
	const existingNames = new Set(getChildren(newRunRPr).filter((c) => c.namespaceURI === w$7).map((c) => c.localName));
	for (const child of getChildren(anchorRPr)) {
		if (child.namespaceURI !== w$7) continue;
		if (child.localName === "rPrChange") continue;
		if (existingNames.has(child.localName)) continue;
		newRunRPr.appendChild(child.cloneNode(true));
	}
}
/** Copy `source`'s attributes onto `target`, leaving any attribute already
* set on `target` untouched. Namespace-aware. */
function mergeMissingAttrs(source, target) {
	const attrs = source.attributes;
	if (!attrs) return;
	for (let i = 0; i < attrs.length; i++) {
		const a = attrs[i];
		const ns = a.namespaceURI;
		const local = a.localName ?? a.name;
		if (ns ? target.hasAttributeNS(ns, local) : target.hasAttribute(local)) continue;
		if (ns) target.setAttributeNS(ns, a.name, a.value);
		else target.setAttribute(a.name, a.value);
	}
}
function inheritFormatForNewParagraphs(newEls, newBlocks, anchor, ownerDoc) {
	if (!anchor) return;
	for (let i = 0; i < newEls.length; i++) {
		const el = newEls[i];
		const block = newBlocks[i];
		if (block.type !== "paragraph") continue;
		if (!isParagraphElement(el)) continue;
		inheritPPrFromAnchor(el, anchor, ownerDoc, block.styleId !== void 0);
		const blockHasExplicitRunFormat = block.runFormat !== void 0;
		for (const r of getChildrenNS(el, w$7, "r")) {
			const existingRPr = firstChildNS(r, w$7, "rPr");
			if (existingRPr !== null) inheritAnchorRunRPr(existingRPr, anchor, blockHasExplicitRunFormat);
			else if (!blockHasExplicitRunFormat) {
				const newRPr = ownerDoc.createElementNS(w$7, "w:rPr");
				inheritAnchorRunRPr(newRPr, anchor, false);
				if (newRPr.childNodes.length > 0) r.insertBefore(newRPr, r.firstChild);
			}
		}
	}
}
function applyDelete(target, documentDoc, trackContext, stale) {
	if (target.paragraphs.length === 0) return 0;
	const touchedParents = /* @__PURE__ */ new Set();
	for (const p of target.paragraphs) if (trackContext.enabled) {
		wrapParagraphContentInDel(p, documentDoc, trackContext);
		markParagraphMarkDeleted(p, documentDoc, trackContext);
	} else {
		const parent = p.parentNode;
		if (parent) {
			touchedParents.add(parent);
			parent.removeChild(p);
		}
		stale.add(p);
	}
	for (const parent of touchedParents) if (parent.namespaceURI === w$7 && parent.localName === "tc") ensureCellHasParagraph(parent, documentDoc);
	return target.paragraphs.length;
}
/** A `<w:tc>` must contain at least one `<w:p>` per ECMA-376 17.4.66. After
* deletes, if the cell has no paragraph child left, append an empty one so
* Word doesn't flag the file as needing repair. The placeholder lands at the
* end of the cell, after any existing `<w:tcPr>`. */
function ensureCellHasParagraph(tc, ownerDoc) {
	if (firstChildNS(tc, w$7, "p")) return;
	tc.appendChild(ownerDoc.createElementNS(w$7, "w:p"));
}
function applyInsertBefore(target, fragment, documentDoc, trackContext, emitCtx, resolverCtx) {
	const newEls = emitFragment(fragment, documentDoc, emitCtx);
	if (newEls.length === 0) return 0;
	const anchor = target.paragraphs[0] ?? null;
	inheritFormatForNewParagraphs(newEls, fragment, anchor, documentDoc);
	if (trackContext.enabled) for (const el of newEls) markParagraphAsInserted(el, documentDoc, trackContext);
	if (anchor) {
		const parent = anchor.parentNode;
		for (const el of newEls) parent.insertBefore(el, anchor);
		normalizeTableSequencing(parent, documentDoc);
	} else {
		insertAtContainerEnd(target.container, newEls, resolverCtx);
		normalizeTableSequencing(target.container, documentDoc);
	}
	return newEls.length;
}
function applyInsertAfter(target, fragment, documentDoc, trackContext, emitCtx, resolverCtx) {
	const newEls = emitFragment(fragment, documentDoc, emitCtx);
	if (newEls.length === 0) return 0;
	const anchor = target.paragraphs[target.paragraphs.length - 1] ?? null;
	inheritFormatForNewParagraphs(newEls, fragment, anchor, documentDoc);
	if (trackContext.enabled) for (const el of newEls) markParagraphAsInserted(el, documentDoc, trackContext);
	if (anchor) {
		const parent = anchor.parentNode;
		const next = anchor.nextSibling;
		if (next) for (const el of newEls) parent.insertBefore(el, next);
		else for (const el of newEls) parent.appendChild(el);
		normalizeTableSequencing(parent, documentDoc);
	} else {
		insertAtContainerEnd(target.container, newEls, resolverCtx);
		normalizeTableSequencing(target.container, documentDoc);
	}
	return newEls.length;
}
function insertAtContainerEnd(container, newEls, resolverCtx) {
	if (container === resolverCtx.body) {
		const sectPr = trailingBodySectPr(container);
		if (sectPr) {
			for (const el of newEls) container.insertBefore(el, sectPr);
			return;
		}
	}
	for (const el of newEls) container.appendChild(el);
}
function applyMerge(target, op, stale, blockersByElement) {
	if (target.paragraphs.length < 2) throw new Error(`merge: needs >= 2 paragraphs, got ${target.paragraphs.length}. Use a range or cell-paragraph-range locator covering at least two paragraphs.`);
	for (const p of target.paragraphs) {
		const reason = blockersByElement.get(p);
		if (reason) throw new Error(`merge: paragraph contains ${explainBlockerReason(reason)} — refusing to merge. Accept/reject existing tracked changes or resolve fields/controls before merging.`);
	}
	const survivor = (op.keepPPr ?? "first") === "last" ? target.paragraphs[target.paragraphs.length - 1] : target.paragraphs[0];
	const allContent = [];
	for (const p of target.paragraphs) for (const child of Array.from(getChildren(p))) {
		if (child.namespaceURI === w$7 && child.localName === "pPr") continue;
		p.removeChild(child);
		allContent.push(child);
	}
	for (const node of allContent) survivor.appendChild(node);
	for (const p of target.paragraphs) if (p !== survivor && p.parentNode) {
		p.parentNode.removeChild(p);
		stale.add(p);
	}
	return target.paragraphs.length;
}
function applyReplace(target, fragment, documentDoc, trackContext, emitCtx, stale) {
	if (target.paragraphs.length === 0) throw new Error("replace op resolved to zero paragraphs — locator must select at least one");
	const newEls = emitFragment(fragment, documentDoc, emitCtx);
	const anchor = target.paragraphs[0];
	const parent = anchor.parentNode;
	inheritFormatForNewParagraphs(newEls, fragment, anchor, documentDoc);
	if (trackContext.enabled) {
		for (const p of target.paragraphs) {
			wrapParagraphContentInDel(p, documentDoc, trackContext);
			markParagraphMarkDeleted(p, documentDoc, trackContext);
		}
		for (const el of newEls) markParagraphAsInserted(el, documentDoc, trackContext);
		for (const el of newEls) parent.insertBefore(el, anchor);
	} else {
		for (const el of newEls) parent.insertBefore(el, anchor);
		for (const p of target.paragraphs) {
			parent.removeChild(p);
			stale.add(p);
		}
	}
	normalizeTableSequencing(parent, documentDoc);
	return newEls.length;
}
/**
* pPr child local-names that clearDirect should strip.
* Whitelist of format-bearing children only — structural children
* (pStyle, numPr, sectPr, pPrChange) are intentionally absent.
*/
const CLEAR_DIRECT_PPR_STRIP_SET = new Set([
	"spacing",
	"ind",
	"jc",
	"outlineLvl",
	"pageBreakBefore",
	"keepNext",
	"keepLines",
	"widowControl",
	"pBdr",
	"shd",
	"tabs",
	"framePr",
	"rPr",
	"adjustRightInd",
	"snapToGrid",
	"autoSpaceDE",
	"autoSpaceDN",
	"textAlignment",
	"textboxTightWrap",
	"bidi",
	"mirrorIndents",
	"wordWrap",
	"kinsoku",
	"overflowPunct",
	"topLinePunct",
	"contextualSpacing",
	"divId"
]);
function applyFormat(target, op, documentDoc, trackContext) {
	if (target.paragraphs.length === 0) throw new Error("format op resolved to zero paragraphs — locator must select at least one");
	const clear = op.clearDirect;
	const shouldClearPPr = clear === "all" || Array.isArray(clear) && clear.includes("pPr");
	const shouldClearRPr = clear === "all" || Array.isArray(clear) && clear.includes("rPr");
	if (shouldClearPPr || shouldClearRPr) for (const p of target.paragraphs) {
		if (shouldClearPPr) {
			const pPr = firstChildNS(p, w$7, "pPr");
			if (pPr) for (const c of [...getChildren(pPr)]) {
				if (c.namespaceURI !== w$7) continue;
				if (CLEAR_DIRECT_PPR_STRIP_SET.has(c.localName)) pPr.removeChild(c);
			}
		}
		if (shouldClearRPr) for (const r of getChildrenNS(p, w$7, "r")) {
			const rPr = firstChildNS(r, w$7, "rPr");
			if (rPr) r.removeChild(rPr);
		}
	}
	let touched = 0;
	for (const p of target.paragraphs) {
		if (op.runFormat) for (const r of getChildrenNS(p, w$7, "r")) {
			const oldRPr = firstChildNS(r, w$7, "rPr");
			const snapshot = oldRPr ? oldRPr.cloneNode(true) : null;
			let rPr;
			if (oldRPr) {
				rPr = oldRPr;
				for (const c of Array.from(getChildren(rPr))) if (c.namespaceURI === w$7 && RPR_MANAGED_LOCAL_NAMES.has(c.localName)) rPr.removeChild(c);
			} else {
				rPr = documentDoc.createElementNS(w$7, "w:rPr");
				r.insertBefore(rPr, r.firstChild);
			}
			for (const c of buildRPrChildren(op.runFormat, documentDoc)) rPr.appendChild(c);
			attachRPrChange(rPr, snapshot, documentDoc, trackContext);
		}
		if (op.paraFormat || op.styleId) {
			const oldPPr = firstChildNS(p, w$7, "pPr");
			const snapshot = oldPPr ? oldPPr.cloneNode(true) : null;
			let pPr;
			if (oldPPr) pPr = oldPPr;
			else {
				pPr = documentDoc.createElementNS(w$7, "w:pPr");
				p.insertBefore(pPr, p.firstChild);
			}
			if (op.styleId) {
				let pStyle = firstChildNS(pPr, w$7, "pStyle");
				if (!pStyle) {
					pStyle = documentDoc.createElementNS(w$7, "w:pStyle");
					pPr.insertBefore(pStyle, pPr.firstChild);
				}
				pStyle.setAttributeNS(w$7, "w:val", op.styleId);
			}
			if (op.paraFormat) {
				for (const c of Array.from(getChildren(pPr))) if (c.namespaceURI === w$7 && PPR_MANAGED_LOCAL_NAMES.has(c.localName)) pPr.removeChild(c);
				for (const c of buildPPrChildren(op.paraFormat, documentDoc)) insertChildInOrder(pPr, c, PPR_CHILD_ORDER);
			}
			attachPPrChange(pPr, snapshot, documentDoc, trackContext);
		}
		touched++;
	}
	return touched;
}
function applySetRun(runEl, op, documentDoc, trackContext) {
	const oldRPr = firstChildNS(runEl, w$7, "rPr");
	const rPrSnapshot = oldRPr ? oldRPr.cloneNode(true) : null;
	if (op.format) {
		let rPr;
		if (oldRPr) {
			rPr = oldRPr;
			for (const c of Array.from(getChildren(rPr))) if (c.namespaceURI === w$7 && RPR_MANAGED_LOCAL_NAMES.has(c.localName)) rPr.removeChild(c);
		} else {
			rPr = documentDoc.createElementNS(w$7, "w:rPr");
			runEl.insertBefore(rPr, runEl.firstChild);
		}
		for (const c of buildRPrChildren(op.format, documentDoc)) rPr.appendChild(c);
		attachRPrChange(rPr, rPrSnapshot, documentDoc, trackContext);
	}
	for (const c of Array.from(getChildren(runEl))) {
		if (c.namespaceURI !== w$7) continue;
		if (c.localName === "rPr") continue;
		runEl.removeChild(c);
	}
	const t = documentDoc.createElementNS(w$7, "w:t");
	t.setAttribute("xml:space", "preserve");
	t.appendChild(documentDoc.createTextNode(op.with));
	runEl.appendChild(t);
	return 1;
}

//#endregion
//#region lib/edit/pangu-lint.ts
const CJK = "[\\u4e00-\\u9fff]";
const ASCII_LATIN_OR_DIGIT = "[A-Za-z0-9]";
const ASCII_SPACE = " +";
const PANGU_AFTER_CJK = new RegExp(`(${CJK})(${ASCII_SPACE})(${ASCII_LATIN_OR_DIGIT})`, "g");
const PANGU_BEFORE_CJK = new RegExp(`(${ASCII_LATIN_OR_DIGIT})(${ASCII_SPACE})(${CJK})`, "g");
/** Run the lint over a list of edit ops. Returns warnings in op order
* with up to 5 entries per op; the report-side caps the displayed
* count separately. Empty list when no hits. */
function lintPanguInEdits(edits) {
	const warnings = [];
	for (const [i, op] of edits.entries()) {
		const texts = collectStringsFromOp(op);
		for (const text of texts) scanString(text, `edits[${i}]`, warnings);
	}
	return warnings;
}
/** Run the lint over a headerFooter config. Walks top-level header /
*  footer and every section override; source labels mirror the config
*  path (`header.default`, `sections["1"].footer.first`, ...). */
function lintPanguInHeaderFooter(hf) {
	const warnings = [];
	const walkBlocks = (blocks, source) => {
		if (!Array.isArray(blocks)) return;
		const texts = [];
		for (const block of blocks) collectStringsFromBlock(block, texts);
		for (const text of texts) scanString(text, source, warnings);
	};
	const walkSurface = (surface, label) => {
		if (surface.default !== void 0) walkBlocks(surface.default, `${label}.default`);
		if (surface.first !== void 0) walkBlocks(surface.first, `${label}.first`);
		if (surface.even !== void 0) walkBlocks(surface.even, `${label}.even`);
	};
	if (hf.header) walkSurface(hf.header, "header");
	if (hf.footer) walkSurface(hf.footer, "footer");
	if (hf.sections) for (const [key, entry] of Object.entries(hf.sections)) {
		if (entry.header) walkSurface(entry.header, `sections["${key}"].header`);
		if (entry.footer) walkSurface(entry.footer, `sections["${key}"].footer`);
	}
	return warnings;
}
function scanString(text, source, out) {
	const PER_SOURCE_CAP = 5;
	let hits = 0;
	const recordHit = (match) => {
		if (hits >= PER_SOURCE_CAP) return;
		hits++;
		const start = Math.max(0, match.index - 12);
		const end = Math.min(text.length, match.index + match[0].length + 12);
		out.push({
			source,
			snippet: text.slice(start, end),
			hit: match[0]
		});
	};
	PANGU_AFTER_CJK.lastIndex = 0;
	let m;
	while ((m = PANGU_AFTER_CJK.exec(text)) !== null) recordHit(m);
	PANGU_BEFORE_CJK.lastIndex = 0;
	while ((m = PANGU_BEFORE_CJK.exec(text)) !== null) recordHit(m);
}
/** Pull every author-supplied string out of one op's fragment. Walks
* ParagraphBlock / CaptionBlock text (string shorthand + InlineRun
* arrays), ImageBlock alt text, TableBlock cell content (string or
* Block[]), and inline run text inside cells. Skips InlineRef +
* InlineEquation — refs carry no prose, equations are LaTeX and have
* their own conventions. */
function collectStringsFromOp(op) {
	const out = [];
	const fragment = op.op === "replace" ? op.with : op.op === "insert-before" || op.op === "insert-after" ? op.content : null;
	if (!fragment) return out;
	for (const block of fragment) collectStringsFromBlock(block, out);
	return out;
}
function collectStringsFromBlock(block, out) {
	if (block === null || typeof block !== "object") return;
	const b = block;
	switch (b.type) {
		case "paragraph":
		case "caption":
			pushRichText(b.text, out);
			return;
		case "image":
			if (typeof b.alt === "string") out.push(b.alt);
			return;
		case "table":
			pushTableRows(b.rows, out);
			return;
		default: return;
	}
}
function pushRichText(text, out) {
	if (typeof text === "string") {
		out.push(text);
		return;
	}
	if (!Array.isArray(text)) return;
	for (const node of text) if (node && typeof node === "object" && "text" in node) {
		const t = node.text;
		if (typeof t === "string") out.push(t);
	}
}
function pushTableRows(rows, out) {
	if (!Array.isArray(rows)) return;
	for (const row of rows) {
		if (!Array.isArray(row)) continue;
		for (const cell of row) pushCellContent(cell, out);
	}
}
function pushCellContent(cell, out) {
	if (typeof cell === "string") {
		out.push(cell);
		return;
	}
	if (!cell || typeof cell !== "object") return;
	const content = cell.content;
	if (typeof content === "string") {
		out.push(content);
		return;
	}
	if (!Array.isArray(content)) return;
	const first = content[0];
	if (first && typeof first === "object" && "type" in first) for (const block of content) collectStringsFromBlock(block, out);
	else pushRichText(content, out);
}

//#endregion
//#region lib/parse/counter-format.ts
/**
* Shared counter-rendering primitives.
*
* Two distinct counter renderers live in this repo:
*   - `lib/apply/numbering-counter.ts` (Word `numFmt` vocabulary —
*     `chineseCounting` uses `〇`, `ideographTraditional` = heavenly
*     stems, etc.)
*   - `lib/edit/caption-counter.ts` (project `SeqFormat` vocabulary —
*     `chinese` uses `零`, `chinese-formal` uses 壹贰叁, etc.)
*
* Their format dispatch tables don't overlap (different vocabularies),
* so they stay separate. But the underlying numeric → glyph primitives
* (roman, spreadsheet-style alpha) are identical and live here.
*
* Chinese-digit tables intentionally do NOT live here — the two
* renderers spell zero differently (`〇` for Word's numFmt vs `零` for
* the project's caption format) and that divergence is by design.
*/
/** Upper-case "I, II, III, IV, V, ...". Callers wanting lower-case
* lowercase the result. Returns "" for n ≤ 0. */
function toRoman(n) {
	if (n <= 0) return "";
	const pairs = [
		[1e3, "M"],
		[900, "CM"],
		[500, "D"],
		[400, "CD"],
		[100, "C"],
		[90, "XC"],
		[50, "L"],
		[40, "XL"],
		[10, "X"],
		[9, "IX"],
		[5, "V"],
		[4, "IV"],
		[1, "I"]
	];
	let out = "";
	let m = n;
	for (const [v, r] of pairs) while (m >= v) {
		out += r;
		m -= v;
	}
	return out;
}
/** Spreadsheet-column style: A..Z, AA..ZZ, AAA... `upper` picks A-Z vs
* a-z. Returns "" for n ≤ 0. */
function toAlphaCounter(n, upper) {
	if (n <= 0) return "";
	const base = upper ? 65 : 97;
	let result = "";
	let m = n;
	while (m > 0) {
		m--;
		result = String.fromCharCode(base + m % 26) + result;
		m = Math.floor(m / 26);
	}
	return result;
}

//#endregion
//#region lib/apply/numbering-counter.ts
/**
* Numbering counter simulator.
*
* Walks the post-edit `document.xml` in body order, maintaining per-`numId`
* per-level counters, and renders each numbered paragraph's `lvlText` with
* the actual counter values substituted. The output map keys by paragraph
* Element so callers (REF placeholder backfill) can look up exactly what
* Word will render once it updates fields on open.
*
* Why a simulator: Word renders auto-numbering at view time from `<w:numPr>`
* + `numbering.xml`. The XML in `document.xml` doesn't contain the rendered
* digits anywhere — running the counter machinery offline is the only way
* to compute the placeholder text a REF field should display BEFORE Word
* updates fields. Without this, the placeholder either stays empty (Word
* shows blank until F9) or carries an incorrect guess.
*
* Limitations:
*   - Counter reset behavior modelled simply: a higher-level increment
*     resets every lower level to its `start` value. Matches Word's default
*     `lvlRestart` (= -1, restart-on-any-higher-level). Custom
*     `lvlRestart` values are not honored — the placeholder may diverge in
*     that case (Word still updates correctly on F9).
*   - `isLgl` is honored: cross-level placeholders rendered as decimal
*     regardless of the referenced level's `numFmt`.
*   - Exotic `numFmt` values (e.g. `cardinalText`, `ordinalText`) fall back
*     to decimal. Common formats covered: decimal, upper/lowerLetter,
*     upper/lowerRoman, chineseCounting, chineseCountingThousand,
*     ideographTraditional, bullet, none.
*/
/** Build a `paragraphElement → RenderedNumbering` map for every paragraph
* carrying `<w:numPr>` — including paragraphs that inherit numbering from
* a style cascade. Paragraphs without any numbering binding are absent. */
function simulateNumberingCounters(documentDoc, numberingDoc, stylesDoc = null) {
	const out = /* @__PURE__ */ new Map();
	const w = NS.w;
	if (!numberingDoc) return out;
	const numIdToAbstract = buildNumIdMap(numberingDoc);
	const numIdStartOverrides = buildNumStartOverrides(numberingDoc);
	const abstractById = buildAbstractMap(numberingDoc);
	const styleNumPr = stylesDoc ? buildStyleNumPrMap(stylesDoc) : /* @__PURE__ */ new Map();
	const body = firstChildNS(documentDoc.documentElement, w, "body");
	if (!body) return out;
	const counters = /* @__PURE__ */ new Map();
	const ensureCounters = (numId) => {
		let m = counters.get(numId);
		if (!m) {
			m = /* @__PURE__ */ new Map();
			counters.set(numId, m);
		}
		return m;
	};
	for (const pEl of walkBodyParagraphs(body)) {
		const binding = resolveParagraphNumbering(pEl, styleNumPr);
		if (!binding) continue;
		const { numId, level } = binding;
		if (numId === "0") continue;
		const abstractId = numIdToAbstract.get(numId);
		if (abstractId === void 0) continue;
		const abstract = abstractById.get(abstractId);
		if (!abstract) continue;
		const lvlDef = abstract.levels.get(level);
		if (!lvlDef) continue;
		const cmap = ensureCounters(numId);
		const overrideForNumId = numIdStartOverrides.get(numId);
		const startFor = (lvl, abstractStart) => overrideForNumId?.get(lvl) ?? abstractStart;
		const cur = cmap.get(level) ?? startFor(level, lvlDef.start) - 1;
		cmap.set(level, cur + 1);
		for (const [lvl, def] of abstract.levels) if (lvl > level) cmap.set(lvl, startFor(lvl, def.start) - 1);
		out.set(pEl, {
			label: renderLvlText(lvlDef, abstract, cmap),
			number: renderCounter(lvlDef.numFmt, cmap.get(level) ?? lvlDef.start)
		});
	}
	return out;
}
/** Resolve a paragraph's effective numbering binding: direct numPr wins;
* otherwise the paragraph's pStyle is traced through the style cascade
* (built once in `styleNumPr`) until a numId surfaces. */
function resolveParagraphNumbering(pEl, styleNumPr) {
	const pPr = firstChildNS(pEl, NS.w, "pPr");
	if (!pPr) return null;
	const direct = firstChildNS(pPr, NS.w, "numPr");
	if (direct) {
		const numIdEl = firstChildNS(direct, NS.w, "numId");
		const ilvlEl = firstChildNS(direct, NS.w, "ilvl");
		const numId = numIdEl ? wVal(numIdEl) : null;
		if (numId) return {
			numId,
			level: ilvlEl ? parseInt(wVal(ilvlEl) || "0", 10) : 0
		};
	}
	const pStyle = firstChildNS(pPr, NS.w, "pStyle");
	if (!pStyle) return null;
	const styleId = wVal(pStyle);
	if (!styleId) return null;
	return styleNumPr.get(styleId) ?? null;
}
/** Walk stylesDoc to map every styleId to its cascade-resolved numbering
* binding (direct numPr or inherited via basedOn). Built once per
* simulator run; lookups are O(1) thereafter. */
function buildStyleNumPrMap(stylesDoc) {
	const w = NS.w;
	const out = /* @__PURE__ */ new Map();
	const root = stylesDoc.documentElement;
	if (!root) return out;
	const styles = /* @__PURE__ */ new Map();
	for (const s of getChildrenNS(root, w, "style")) {
		const id = wAttr(s, "styleId");
		if (id) styles.set(id, s);
	}
	const resolve = (id, seen) => {
		if (seen.has(id)) return null;
		seen.add(id);
		const el = styles.get(id);
		if (!el) return null;
		const pPr = firstChildNS(el, w, "pPr");
		if (pPr) {
			const numPr = firstChildNS(pPr, w, "numPr");
			if (numPr) {
				const numIdEl = firstChildNS(numPr, w, "numId");
				const ilvlEl = firstChildNS(numPr, w, "ilvl");
				const numId = numIdEl ? wVal(numIdEl) : null;
				if (numId) return {
					numId,
					level: ilvlEl ? parseInt(wVal(ilvlEl) || "0", 10) : 0
				};
			}
		}
		const basedOn = firstChildNS(el, w, "basedOn");
		if (basedOn) {
			const parent = wVal(basedOn);
			if (parent) return resolve(parent, seen);
		}
		return null;
	};
	for (const id of styles.keys()) {
		const r = resolve(id, /* @__PURE__ */ new Set());
		if (r) out.set(id, r);
	}
	return out;
}
/** Read `<w:lvlOverride><w:startOverride>` on every `<w:num>` to produce
* `numId → level → startValue`. Used by the simulator so a forked numId
* (perInstance restart) starts counting from its override instead of the
* abstractNum's `<w:start>`. */
function buildNumStartOverrides(numberingDoc) {
	const out = /* @__PURE__ */ new Map();
	const w = NS.w;
	for (const num of getChildrenNS(numberingDoc.documentElement, w, "num")) {
		const numId = wAttr(num, "numId");
		if (!numId) continue;
		const perLevel = /* @__PURE__ */ new Map();
		for (const ovr of getChildrenNS(num, w, "lvlOverride")) {
			const ilvl = wAttr(ovr, "ilvl");
			if (ilvl === null) continue;
			const startOvr = firstChildNS(ovr, w, "startOverride");
			if (!startOvr) continue;
			const val = wAttr(startOvr, "val");
			if (val === null) continue;
			perLevel.set(parseInt(ilvl, 10), parseInt(val, 10));
		}
		if (perLevel.size > 0) out.set(numId, perLevel);
	}
	return out;
}
function buildNumIdMap(numberingDoc) {
	const out = /* @__PURE__ */ new Map();
	const w = NS.w;
	for (const num of getChildrenNS(numberingDoc.documentElement, w, "num")) {
		const numId = wAttr(num, "numId");
		if (!numId) continue;
		const absRef = firstChildNS(num, w, "abstractNumId");
		if (!absRef) continue;
		const absId = wAttr(absRef, "val") ?? wVal(absRef);
		if (absId) out.set(numId, absId);
	}
	return out;
}
function buildAbstractMap(numberingDoc) {
	const out = /* @__PURE__ */ new Map();
	const w = NS.w;
	for (const abs of getChildrenNS(numberingDoc.documentElement, w, "abstractNum")) {
		const id = wAttr(abs, "abstractNumId");
		if (!id) continue;
		const levels = /* @__PURE__ */ new Map();
		for (const lvl of getChildrenNS(abs, w, "lvl")) {
			const ilvl = wAttr(lvl, "ilvl");
			if (ilvl === null) continue;
			const level = parseInt(ilvl, 10);
			const numFmtEl = firstChildNS(lvl, w, "numFmt");
			const lvlTextEl = firstChildNS(lvl, w, "lvlText");
			const startEl = firstChildNS(lvl, w, "start");
			const isLglEl = firstChildNS(lvl, w, "isLgl");
			levels.set(level, {
				level,
				numFmt: numFmtEl && wAttr(numFmtEl, "val") || "decimal",
				lvlText: lvlTextEl && wAttr(lvlTextEl, "val") || "",
				start: parseInt(startEl && wAttr(startEl, "val") || "1", 10),
				isLgl: !!isLglEl
			});
		}
		out.set(id, { levels });
	}
	return out;
}
function renderLvlText(lvl, abstract, counters) {
	if (lvl.numFmt === "bullet" || lvl.numFmt === "none") return lvl.lvlText;
	return lvl.lvlText.replace(/%(\d)/g, (_match, digit) => {
		const refLevel = parseInt(digit, 10) - 1;
		const counter = counters.get(refLevel);
		if (counter === void 0) return "";
		const refLvlDef = abstract.levels.get(refLevel);
		return renderCounter(lvl.isLgl ? "decimal" : refLvlDef?.numFmt ?? "decimal", counter);
	});
}
function renderCounter(numFmt, value) {
	switch (numFmt) {
		case "decimal": return String(value);
		case "decimalZero": return value < 10 ? `0${value}` : String(value);
		case "upperLetter": return toAlphaCounter(value, true);
		case "lowerLetter": return toAlphaCounter(value, false);
		case "upperRoman": return toRoman(value).toUpperCase();
		case "lowerRoman": return toRoman(value).toLowerCase();
		case "chineseCounting":
		case "chineseCountingThousand": return toChineseCounting(value);
		case "ideographTraditional": return toHeavenlyStem(value);
		case "bullet":
		case "none": return "";
		default: return String(value);
	}
}
const CHINESE_DIGITS = [
	"〇",
	"一",
	"二",
	"三",
	"四",
	"五",
	"六",
	"七",
	"八",
	"九"
];
function toChineseCounting(n) {
	if (n === 0) return "〇";
	if (n < 10) return CHINESE_DIGITS[n];
	if (n < 20) return n === 10 ? "十" : `十${CHINESE_DIGITS[n - 10]}`;
	if (n < 100) {
		const tens = Math.floor(n / 10);
		const ones = n % 10;
		return `${CHINESE_DIGITS[tens]}十${ones === 0 ? "" : CHINESE_DIGITS[ones]}`;
	}
	return String(n).split("").map((d) => CHINESE_DIGITS[parseInt(d, 10)]).join("");
}
const HEAVENLY_STEMS = [
	"甲",
	"乙",
	"丙",
	"丁",
	"戊",
	"己",
	"庚",
	"辛",
	"壬",
	"癸"
];
function toHeavenlyStem(n) {
	if (n < 1) return "";
	return HEAVENLY_STEMS[(n - 1) % 10];
}
/** Extract the visible text from a paragraph's `<w:t>` descendants — used
* for the "full" display option (REF without switches renders the
* bookmark's text content). Skips text inside fldChar=begin/end pairs to
* avoid recursive reference text leaking in. */
function extractParagraphText(pEl) {
	const w = NS.w;
	const parts = [];
	let inField = 0;
	const walk = (el) => {
		for (const c of getChildren(el)) {
			if (c.namespaceURI === w && c.localName === "fldChar") {
				const type = wAttr(c, "fldCharType");
				if (type === "begin") inField++;
				else if (type === "end" && inField > 0) inField--;
				continue;
			}
			if (c.namespaceURI === w && c.localName === "t" && inField === 0) {
				parts.push(c.textContent || "");
				continue;
			}
			walk(c);
		}
	};
	walk(pEl);
	return parts.join("");
}

//#endregion
//#region lib/parse/caption-resolver.ts
/**
* Resolve the raw `captions` config (agent-facing) into per-identifier
* `ResolvedCaptionConfig` records consumed by caption-emit + counter sim.
*
* Resolution work:
*   - Apply defaults: prefix/suffix → "", format → "arabic",
*     chapterPrefix → [], chapterSeparator → ".", bodySeparator → " ",
*     subCounter format → "alphabetic", subCounter prefix/suffix → ""
*   - For each chapterPrefix styleId, look up styleName from styles.xml's
*     `<w:name w:val="..."/>` and outlineLvl from `<w:pPr><w:outlineLvl/>`
*   - Derive SEQ \s switch's outline level from the LAST chapterPrefix
*     entry (deepest chapter level controls counter restart)
*
* Throws on:
*   - chapterPrefix references unknown styleId
*   - chapterPrefix references a style without `<w:name>` (engine bug
*     for built-in styles; agent error for hand-written styles)
*
* Does NOT throw on:
*   - chapterPrefix style without outlineLvl binding (Word will render
*     "0"; emit pre-scan can warn separately)
*/
const w$6 = NS.w;
function resolveCaptions(raw, stylesDoc) {
	const byIdentifier = /* @__PURE__ */ new Map();
	const styleIdToName = /* @__PURE__ */ new Map();
	const warnings = [];
	if (!raw) return {
		byIdentifier,
		styleIdToName,
		warnings
	};
	const styleIndex = indexStyles(stylesDoc);
	const styleUses = /* @__PURE__ */ new Map();
	for (const [captionIdentifier, entry] of Object.entries(raw)) for (const rawEntry of entry.chapterPrefix ?? []) {
		const styleId = typeof rawEntry === "string" ? rawEntry : rawEntry.styleId;
		const use = {
			captionIdentifier,
			form: typeof rawEntry === "string" ? "string" : "object",
			format: typeof rawEntry === "string" ? void 0 : rawEntry.format
		};
		const prior = styleUses.get(styleId);
		if (prior) {
			const first = prior[0];
			if (first.form !== use.form || first.format !== use.format) throw new Error(`captions: chapterPrefix for styleId "${styleId}" is referenced inconsistently — captions["${first.captionIdentifier}"] uses ${describeForm(first)} but captions["${use.captionIdentifier}"] uses ${describeForm(use)}. All chapterPrefix entries for the same styleId must use the SAME form — either the bare string (heading's native rendering) or an object with a single matching format. Mixed forms produce divergent chapter sources at render time (STYLEREF + heading text vs hidden parallel SEQ).`);
		}
		styleUses.set(styleId, [...prior ?? [], use]);
	}
	for (const [identifier, entry] of Object.entries(raw)) {
		const chapterPrefix = [];
		for (const rawEntry of entry.chapterPrefix ?? []) {
			const styleId = typeof rawEntry === "string" ? rawEntry : rawEntry.styleId;
			const format = typeof rawEntry === "string" ? void 0 : rawEntry.format;
			const info = styleIndex.get(styleId);
			if (!info) throw new Error(`captions["${identifier}"].chapterPrefix references unknown styleId "${styleId}". Declare the style in styles[] or fix the reference.`);
			if (info.outlineLevel === void 0) warnings.push(`captions["${identifier}"].chapterPrefix references styleId "${styleId}" but that style has no <w:outlineLvl> binding. STYLEREF will render 0 at runtime — bind the style to an outline-numbering scheme or remove it from chapterPrefix.`);
			chapterPrefix.push({
				styleName: info.name,
				styleId,
				outlineLevel: info.outlineLevel ?? 0,
				format
			});
			styleIdToName.set(styleId, info.name);
		}
		const restartAtOutlineLevel = chapterPrefix.length > 0 ? chapterPrefix[chapterPrefix.length - 1].outlineLevel || void 0 : void 0;
		const paraInfo = styleIndex.get(entry.styleId);
		if (paraInfo) styleIdToName.set(entry.styleId, paraInfo.name);
		const subCounter = entry.subCounter ? {
			format: entry.subCounter.format ?? "alphabetic",
			prefix: entry.subCounter.prefix ?? "",
			suffix: entry.subCounter.suffix ?? ""
		} : void 0;
		byIdentifier.set(identifier, {
			identifier,
			prefix: entry.prefix ?? "",
			labelSeparator: entry.labelSeparator ?? "",
			suffix: entry.suffix ?? "",
			format: entry.format ?? "arabic",
			chapterPrefix,
			chapterSeparator: entry.chapterSeparator ?? ".",
			bodySeparator: entry.bodySeparator ?? " ",
			paragraphStyleId: entry.styleId,
			restartAtOutlineLevel,
			subCounter
		});
	}
	return {
		byIdentifier,
		styleIdToName,
		warnings
	};
}
function describeForm(use) {
	if (use.form === "string") return `the bare-string form (heading's native rendering)`;
	return `the object form with format "${use.format ?? "(unset)"}"`;
}
function indexStyles(stylesDoc) {
	const out = /* @__PURE__ */ new Map();
	if (!stylesDoc) return out;
	const root = stylesDoc.documentElement;
	if (!root) return out;
	for (const styleEl of getChildrenNS(root, w$6, "style")) {
		const id = wAttr(styleEl, "styleId");
		if (!id) continue;
		const nameEl = firstChildNS(styleEl, w$6, "name");
		const name = nameEl ? wAttr(nameEl, "val") ?? id : id;
		let outlineLevel;
		const pPr = firstChildNS(styleEl, w$6, "pPr");
		if (pPr) {
			const lvlEl = firstChildNS(pPr, w$6, "outlineLvl");
			if (lvlEl) {
				const v = wAttr(lvlEl, "val");
				const n = v ? parseInt(v, 10) : NaN;
				if (Number.isFinite(n) && n >= 0 && n <= 8) outlineLevel = n + 1;
			}
		}
		out.set(id, {
			name,
			outlineLevel
		});
	}
	return out;
}

//#endregion
//#region lib/edit/caption-counter.ts
/**
* Caption counter simulator.
*
* Walks the document body in order, advancing per-identifier counters
* for caption-class paragraphs, resolving STYLEREF chapter prefixes
* against outline numbering outputs, producing:
*
*   - Per-field result values to write into SEQ / STYLEREF / sub-counter
*     result runs (so caption paragraphs read correctly without Word F9)
*   - Per-caption-paragraph full rendered text (prefix + chapter +
*     parent + sub + suffix), consumed by REF placeholder backfill so
*     body cross-references like "see (2.3)" render pre-F9
*
* Pure function over input data structures — no styles.xml or outline
* simulator access. Caller (apply pipeline) preprocesses:
*
*   - Walks outline simulator's output to build `outlineParagraphs` map
*     (each tagged with its styleName + rendered number)
*   - Resolves agent's captions config into ResolvedCaptionConfig per
*     identifier (styleId → styleName + outline level via styles.xml)
*   - Caption emitters create `fills` + `resets` records pointing at
*     the actual result `<w:t>` elements
*
* State machine per identifier:
*
*   subGroup omitted (standalone): parent++, sub = 0, close open subgroup
*   subGroup "start":              parent++, sub = 1, open subgroup
*   subGroup "continue":           requires open subgroup; sub++, parent unchanged
*   CaptionCounterReset:           parent = newValue, sub = 0, close subgroup
*
* Subgroup continue without preceding start throws — pre-scan should
* catch this, but the simulator double-checks (the resulting render
* would silently look wrong otherwise).
*/
const w$5 = NS.w;
function simulateCaptions(documentDoc, input) {
	const fieldValues = /* @__PURE__ */ new Map();
	const fullCaptionText = /* @__PURE__ */ new Map();
	const states = /* @__PURE__ */ new Map();
	/** styleName → heading's rendered counter text (e.g. "一"). Used when
	* the chapterPrefix entry has no `format` override — STYLEREF returns
	* the heading's native rendering. */
	const latestHeading = /* @__PURE__ */ new Map();
	/** styleName → integer occurrence count. Used when a chapterPrefix
	* entry's `format` override is in effect — counter sim re-formats the
	* integer per the override (matching Word's `\* <FORMAT>` runtime). */
	const latestHeadingCount = /* @__PURE__ */ new Map();
	const fillByPara = /* @__PURE__ */ new Map();
	for (const f of input.fills) fillByPara.set(f.paragraph, f);
	const resetByPara = /* @__PURE__ */ new Map();
	for (const r of input.resets) resetByPara.set(r.paragraph, r);
	const resetTriggers = /* @__PURE__ */ new Map();
	for (const [identifier, config] of input.configs) {
		if (config.chapterPrefix.length === 0) continue;
		const lastStyleName = config.chapterPrefix[config.chapterPrefix.length - 1].styleName;
		const list = resetTriggers.get(lastStyleName) ?? [];
		list.push(identifier);
		resetTriggers.set(lastStyleName, list);
	}
	const root = documentDoc.documentElement;
	if (!root) return {
		fieldValues,
		fullCaptionText
	};
	const body = firstChildNS(root, w$5, "body");
	if (!body) return {
		fieldValues,
		fullCaptionText
	};
	for (const para of walkBodyParagraphs(body)) {
		const outlineInfo = input.outlineParagraphs.get(para);
		if (outlineInfo) {
			latestHeading.set(outlineInfo.styleName, outlineInfo.rendered);
			latestHeadingCount.set(outlineInfo.styleName, (latestHeadingCount.get(outlineInfo.styleName) ?? 0) + 1);
			const triggered = resetTriggers.get(outlineInfo.styleName);
			if (triggered) for (const identifier of triggered) {
				const state = stateFor(states, identifier);
				state.parent = 0;
				state.sub = 0;
				state.openSubGroup = false;
			}
			continue;
		}
		const reset = resetByPara.get(para);
		if (reset) {
			const state = stateFor(states, reset.identifier);
			state.parent = reset.newValue - 1;
			state.sub = 0;
			state.openSubGroup = false;
			continue;
		}
		const fill = fillByPara.get(para);
		if (!fill) continue;
		const config = input.configs.get(fill.identifier);
		if (!config) continue;
		const state = stateFor(states, fill.identifier);
		applyFillToState(state, fill);
		for (let i = 0; i < fill.chapterPrefixResults.length; i++) {
			const entry = config.chapterPrefix[i];
			const resultEl = fill.chapterPrefixResults[i];
			if (!entry) continue;
			let rendered;
			if (entry.format !== void 0) {
				const count = latestHeadingCount.get(entry.styleName) ?? 0;
				rendered = count > 0 ? formatCounter(count, entry.format) : "0";
			} else rendered = latestHeading.get(entry.styleName) ?? "0";
			fieldValues.set(resultEl, rendered);
		}
		const parentRendered = formatCounter(state.parent, config.format);
		fieldValues.set(fill.parentSeqResult, parentRendered);
		let subRendered = "";
		if (fill.subSeqResult && config.subCounter) {
			subRendered = formatCounter(state.sub, config.subCounter.format);
			fieldValues.set(fill.subSeqResult, subRendered);
		}
		const chapterParts = [];
		for (let i = 0; i < fill.chapterPrefixResults.length; i++) {
			const resultEl = fill.chapterPrefixResults[i];
			chapterParts.push(fieldValues.get(resultEl) ?? "0");
		}
		const chapterStr = chapterParts.length > 0 ? chapterParts.join(config.chapterSeparator) + config.chapterSeparator : "";
		const labelSep = chapterStr !== "" ? config.labelSeparator ?? "" : "";
		let subStr = "";
		if (fill.subSeqResult && config.subCounter && state.sub > 0) subStr = config.subCounter.prefix + subRendered + config.subCounter.suffix;
		const full = config.prefix + labelSep + chapterStr + parentRendered + subStr + config.suffix;
		fullCaptionText.set(para, full);
	}
	return {
		fieldValues,
		fullCaptionText
	};
}
function stateFor(states, identifier) {
	let s = states.get(identifier);
	if (!s) {
		s = {
			parent: 0,
			sub: 0,
			openSubGroup: false
		};
		states.set(identifier, s);
	}
	return s;
}
function applyFillToState(state, fill) {
	if (fill.subGroup === "continue") {
		if (!state.openSubGroup) throw new Error(`Caption "${fill.identifier}": subGroup "continue" without preceding "start". Pre-scan should have caught this — investigate the pipeline if it reaches the simulator.`);
		state.sub++;
	} else if (fill.subGroup === "start") {
		state.parent++;
		state.sub = 1;
		state.openSubGroup = true;
	} else {
		state.parent++;
		state.sub = 0;
		state.openSubGroup = false;
	}
}
function formatCounter(n, format) {
	switch (format) {
		case "arabic": return String(n);
		case "alphabetic": return toAlphaCounter(n, false);
		case "ALPHABETIC": return toAlphaCounter(n, true);
		case "roman": return toRoman(n).toLowerCase();
		case "ROMAN": return toRoman(n);
		case "chinese": return toChineseNum(n, false);
		case "chinese-formal": return toChineseNum(n, true);
	}
}
function toChineseNum(n, formal) {
	const digits = formal ? [
		"零",
		"壹",
		"贰",
		"叁",
		"肆",
		"伍",
		"陆",
		"柒",
		"捌",
		"玖"
	] : [
		"零",
		"一",
		"二",
		"三",
		"四",
		"五",
		"六",
		"七",
		"八",
		"九"
	];
	if (n <= 0) return "";
	if (n < 10) return digits[n];
	if (n < 20) return n === 10 ? "十" : "十" + digits[n - 10];
	if (n < 100) {
		const tens = Math.floor(n / 10);
		const ones = n % 10;
		return digits[tens] + "十" + (ones === 0 ? "" : digits[ones]);
	}
	return String(n);
}

//#endregion
//#region lib/apply/standardize-captions.ts
/**
* Standardize re-emit for caption paragraphs already in the document.
*
* Scans body for paragraphs where:
*   (a) paragraph style ∈ captions.styleIds AND
*   (b) paragraph contains a SEQ field with identifier matching one of
*       the declared captions
*
* For each match (excluding paragraphs freshly emitted in this apply
* pass), rebuilds the pre-body run sequence in place using the current
* captions config. Preserves:
*   - paragraph element identity (REF backfills' resolved targets stay
*     valid)
*   - bookmark id + name (so REFs keep resolving)
*   - SEQ identifier (Word's running counter continues)
*   - body text (everything after the primary bookmarkEnd / last fldChar
*     end)
*
* Identifier mismatch (SEQ exists but its identifier isn't in the
* captions config): pass through unchanged + warn.
*
* Produces fresh PendingCaptionFill entries pointing at the rebuilt
* result text elements; caller appends to the apply pipeline's
* pendingCaptionFills list so the counter sim renders these too.
*/
const w$4 = NS.w;
function standardizeCaptions(documentDoc, captions, bookmarkAllocator, freshlyEmitted) {
	const fills = [];
	const warnings = [];
	if (captions.size === 0) return {
		fills,
		warnings
	};
	const body = firstChildNS(documentDoc.documentElement, w$4, "body");
	if (!body) return {
		fills,
		warnings
	};
	const styleToIdentifier = /* @__PURE__ */ new Map();
	for (const [id, config] of captions) styleToIdentifier.set(config.paragraphStyleId, id);
	let paragraphIndex = 0;
	for (const para of walkBodyParagraphs(body)) {
		paragraphIndex++;
		if (freshlyEmitted.has(para)) continue;
		const styleId = paragraphStyleId(para);
		if (!styleId) continue;
		const expectedIdentifier = styleToIdentifier.get(styleId);
		if (!expectedIdentifier) continue;
		const seqIdentifier = findSeqIdentifier(para);
		if (!seqIdentifier) continue;
		if (seqIdentifier !== expectedIdentifier) {
			warnings.push(`standardize-captions: paragraph #${paragraphIndex} styled as "${styleId}" carries SEQ "${seqIdentifier}" but the captions table expects "${expectedIdentifier}". Passed through unchanged — add the identifier to the captions table to bring it under standardize.`);
			continue;
		}
		const config = captions.get(expectedIdentifier);
		const { bookmarkId, bookmarkName, bodyText, subGroup } = extractCaptionParts(para, expectedIdentifier, config.bodySeparator);
		const fill = rebuildCaptionParagraphInPlace(para, config, bookmarkId, bookmarkName, bodyText, subGroup, documentDoc);
		if (fill) fills.push(fill);
		if (bookmarkName) bookmarkAllocator.bindRangeBookmark(bookmarkName, para);
	}
	return {
		fills,
		warnings
	};
}
/** Replace runs after pPr in the paragraph with a freshly-built caption
* sequence (using current config). Returns a PendingCaptionFill for the
* new SEQ/STYLEREF result elements. */
function rebuildCaptionParagraphInPlace(paragraph, config, bookmarkId, bookmarkName, bodyText, subGroup, ownerDoc) {
	const seq = buildCaptionRunSequence(ownerDoc, {
		captionConfig: config,
		subGroup,
		bookmark: bookmarkId !== void 0 && bookmarkName !== void 0 ? {
			id: bookmarkId,
			name: bookmarkName
		} : void 0,
		body: bodyText === "" ? void 0 : bodyText
	});
	const existingChildren = getChildren(paragraph);
	const existingPPr = existingChildren.find((c) => c.namespaceURI === w$4 && c.localName === "pPr");
	for (const c of existingChildren) if (c !== existingPPr) paragraph.removeChild(c);
	for (const r of seq.runs) paragraph.appendChild(r);
	return {
		paragraph,
		identifier: config.identifier,
		subGroup,
		chapterPrefixResults: seq.chapterPrefixResults,
		parentSeqResult: seq.parentSeqResult,
		subSeqResult: seq.subSeqResult,
		bodyText: bodyText === "" ? void 0 : bodyText
	};
}
function extractCaptionParts(paragraph, identifier, bodySeparator) {
	let bookmarkId;
	let bookmarkName;
	const children = getChildren(paragraph);
	for (const c of children) if (c.namespaceURI === w$4 && c.localName === "bookmarkStart") {
		const id = wAttr(c, "id");
		const name = wAttr(c, "name");
		if (id !== null && name !== null) {
			const n = parseInt(id, 10);
			if (Number.isFinite(n)) {
				bookmarkId = n;
				bookmarkName = name;
			}
		}
		break;
	}
	let boundaryIdx = -1;
	for (let i = children.length - 1; i >= 0; i--) {
		const c = children[i];
		if (c.namespaceURI === w$4 && c.localName === "bookmarkEnd") {
			boundaryIdx = i;
			break;
		}
	}
	if (boundaryIdx < 0) for (let i = children.length - 1; i >= 0; i--) {
		const c = children[i];
		if (c.namespaceURI !== w$4 || c.localName !== "r") continue;
		const fld = firstChildNS(c, w$4, "fldChar");
		if (fld && wAttr(fld, "fldCharType") === "end") {
			boundaryIdx = i;
			break;
		}
	}
	let bodyText = "";
	if (boundaryIdx >= 0) {
		let collected = "";
		for (let i = boundaryIdx + 1; i < children.length; i++) {
			const c = children[i];
			if (c.namespaceURI !== w$4 || c.localName !== "r") continue;
			for (const t of getChildren(c)) if (t.namespaceURI === w$4 && t.localName === "t") collected += t.textContent ?? "";
		}
		bodyText = collected.startsWith(bodySeparator) ? collected.slice(bodySeparator.length) : collected;
	}
	const subGroup = detectSubGroup(paragraph, identifier);
	return {
		bookmarkId,
		bookmarkName,
		bodyText,
		subGroup
	};
}
function detectSubGroup(paragraph, identifier) {
	const parsed = parseFieldRuns(paragraphRuns(paragraph));
	let parentRepeat = false;
	let subSeqPresent = false;
	const subId = `${identifier}Sub`;
	for (const entry of parsed) {
		if (entry.kind !== "field" || entry.fieldType !== "SEQ") continue;
		if (entry.details.identifier === identifier && entry.details.repeat) parentRepeat = true;
		if (entry.details.identifier === subId) subSeqPresent = true;
	}
	if (parentRepeat) return "continue";
	if (subSeqPresent) return "start";
}
function findSeqIdentifier(paragraph) {
	return seqFields(paragraph, { skipRepeat: true })[0]?.identifier;
}

//#endregion
//#region lib/apply/inject-chapter-counters.ts
/**
* For each paragraph whose style is referenced under `chapterPrefix`
* with a `format` override, emit a SIBLING engine-managed paragraph
* immediately after it. The sibling carries the hidden auto-chapter
* `SEQ _chap_<styleId> \* <FORMAT>` field; caption-emit's chapter prefix
* `SEQ _chap_<styleId> \c \* <FORMAT>` reads the counter via Word's `\c`
* (repeat) switch.
*
* Why a parallel hidden counter instead of `STYLEREF "Heading 1" \n
* \* ARABIC`: STYLEREF \n returns the heading's full rendered lvlText
* ("第一章" for chineseCounting), and `\* ARABIC` does NOT extract the
* numeric portion — Word renders "第一章.1" instead of "1.1". The
* hidden SEQ maintains a parallel Arabic counter that Word's F9 keeps
* live on H1 add / remove.
*
* Why a sibling paragraph instead of inline at the heading's tail:
* STYLEREF "heading 1" \* MERGEFORMAT reads the H1 paragraph's full
* text content (including hidden runs) and reformats with the field's
* own rPr — stripping vanish. A trailing hidden SEQ inside H1 then
* surfaces as visible numeric pollution in any header that styleref-
* references the heading. Moving the SEQ to a separate paragraph keeps
* H1's text content clean for STYLEREF / TOC / REF / nav-pane extraction
* while preserving document-order SEQ increment for captions.
*
* Sibling layout: a new `<w:p>` with `_HiddenChapterCounter` pStyle,
* paragraph-mark vanish (pPr.rPr.vanish) so the paragraph collapses to
* zero height when hidden text isn't shown, and the same 5-run SEQ
* skeleton with vanish on each run. The `_`-prefix styleId is the
* engine-managed-paragraph convention — `walkIndexedParagraphs` and the
* `DocumentParser` skip these so agent-facing indices match user-visible
* content order. See CLAUDE.md "Mechanical correctness".
*/
const w$3 = NS.w;
/** styleId used for engine-managed sibling counter paragraphs. The
*  leading underscore flags it as engine-internal — walkers and the
*  parser skip `_`-prefix styleIds so the agent-visible paragraph index
*  reflects content order, not raw DOM order. */
const HIDDEN_CHAPTER_COUNTER_STYLE_ID = "_HiddenChapterCounter";
/** Returns the count of sibling counter paragraphs newly emitted. The
* count surfaces in the dry-run change report so the agent can preview
* chapter-counter coverage before committing. */
function injectChapterCounters(documentDoc, captions) {
	const tracked = /* @__PURE__ */ new Map();
	for (const [captionIdentifier, config] of captions) for (const entry of config.chapterPrefix) if (entry.format !== void 0) {
		const existing = tracked.get(entry.styleId);
		if (existing !== void 0 && existing.format !== entry.format) throw new Error(`captions conflict: chapterPrefix styleId "${entry.styleId}" appears with format "${existing.format}" (captions["${existing.captionIdentifier}"]) and "${entry.format}" (captions["${captionIdentifier}"]). The hidden auto-chapter counter for one style must use a single format — split into two styleIds or unify the format.`);
		tracked.set(entry.styleId, {
			format: entry.format,
			captionIdentifier
		});
	}
	if (tracked.size === 0) return 0;
	const body = firstChildNS(documentDoc.documentElement, w$3, "body");
	if (!body) return 0;
	const targets = [];
	for (const para of walkBodyParagraphs(body)) {
		const styleId = paragraphStyleId(para);
		if (!styleId) continue;
		const entry = tracked.get(styleId);
		if (entry === void 0) continue;
		targets.push({
			para,
			styleId,
			format: entry.format
		});
	}
	for (const { para, styleId, format } of targets) emitCounterSibling(para, styleId, format, documentDoc);
	return targets.length;
}
/** Insert an engine-managed paragraph holding the hidden SEQ field
*  immediately after `heading`. */
function emitCounterSibling(heading, styleId, format, ownerDoc) {
	const sibling = ownerDoc.createElementNS(w$3, "w:p");
	const pPr = ownerDoc.createElementNS(w$3, "w:pPr");
	const pStyle = ownerDoc.createElementNS(w$3, "w:pStyle");
	pStyle.setAttributeNS(w$3, "w:val", HIDDEN_CHAPTER_COUNTER_STYLE_ID);
	pPr.appendChild(pStyle);
	const pRPr = ownerDoc.createElementNS(w$3, "w:rPr");
	pRPr.appendChild(ownerDoc.createElementNS(w$3, "w:vanish"));
	pPr.appendChild(pRPr);
	sibling.appendChild(pPr);
	const { runs } = emitSeqField(ownerDoc, {
		identifier: chapterCounterIdentifier(styleId),
		format
	});
	for (const r of runs) {
		addVanishRPr(r, ownerDoc);
		sibling.appendChild(r);
	}
	heading.parentNode.insertBefore(sibling, heading.nextSibling);
}
/** Inject the `_HiddenChapterCounter` style definition into `stylesDoc`
*  when not already present. Called from the apply pipeline whenever
*  `injectChapterCounters` emits at least one sibling — without the
*  matching style entry the `<w:pStyle>` reference is dangling.
*
*  Style shape:
*   - `<w:hidden/>` + `<w:semiHidden/>` + `<w:unhideWhenUsed/>` so it
*     doesn't clutter the agent / user style panels
*   - `<w:pPr>` with zero spacing and paragraph-mark vanish (defense
*     in depth — paragraphs also set the same vanish on their pPr.rPr)
*   - `<w:rPr>` with `<w:vanish/>` and 2-half-point size so the runs
*     are character-hidden and minimal even on Word versions that
*     ignore the pPr vanish
*
*  Returns true when the style was newly injected. */
function ensureHiddenChapterCounterStyle(stylesDoc) {
	const root = stylesDoc.documentElement;
	if (!root) return false;
	for (const s of getChildrenNS(root, w$3, "style")) if (wAttr(s, "styleId") === "_HiddenChapterCounter") return false;
	const style = stylesDoc.createElementNS(w$3, "w:style");
	style.setAttributeNS(w$3, "w:type", "paragraph");
	style.setAttributeNS(w$3, "w:styleId", HIDDEN_CHAPTER_COUNTER_STYLE_ID);
	const name = stylesDoc.createElementNS(w$3, "w:name");
	name.setAttributeNS(w$3, "w:val", "_Hidden Chapter Counter");
	style.appendChild(name);
	const basedOn = stylesDoc.createElementNS(w$3, "w:basedOn");
	basedOn.setAttributeNS(w$3, "w:val", "Normal");
	style.appendChild(basedOn);
	style.appendChild(stylesDoc.createElementNS(w$3, "w:hidden"));
	style.appendChild(stylesDoc.createElementNS(w$3, "w:semiHidden"));
	style.appendChild(stylesDoc.createElementNS(w$3, "w:unhideWhenUsed"));
	const pPr = stylesDoc.createElementNS(w$3, "w:pPr");
	const spacing = stylesDoc.createElementNS(w$3, "w:spacing");
	spacing.setAttributeNS(w$3, "w:before", "0");
	spacing.setAttributeNS(w$3, "w:after", "0");
	spacing.setAttributeNS(w$3, "w:line", "240");
	spacing.setAttributeNS(w$3, "w:lineRule", "auto");
	pPr.appendChild(spacing);
	style.appendChild(pPr);
	const rPr = stylesDoc.createElementNS(w$3, "w:rPr");
	rPr.appendChild(stylesDoc.createElementNS(w$3, "w:vanish"));
	const sz = stylesDoc.createElementNS(w$3, "w:sz");
	sz.setAttributeNS(w$3, "w:val", "2");
	rPr.appendChild(sz);
	style.appendChild(rPr);
	root.appendChild(style);
	return true;
}

//#endregion
//#region lib/apply/settings-mutation.ts
/**
* `word/settings.xml` mutations.
*
* One concern only at this point: the `<w:updateFields>` flag that tells
* Word to update every field (including REF cross-references) when the
* doc next opens. Without this, an inserted cross-reference shows its
* placeholder text until the user manually triggers Ctrl+A → F9.
*
* Three cases:
*   1. settings.xml exists and already has `<w:updateFields w:val="true"/>` → no-op
*   2. settings.xml exists but lacks the element (or has val="false") → mutate / append
*   3. settings.xml doesn't exist → fabricate a minimal one + register it in
*      [Content_Types].xml and word/_rels/document.xml.rels
*
* Case 3 is rare — every Word-generated docx ships with settings.xml — but
* common enough in hand-built fixtures that handling it cleanly is worth
* 30 extra lines.
*/
const w$2 = NS.w;
const SETTINGS_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
const SETTINGS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
/** Empty stub for the fabrication path. Each mutator adds its own children
*  via `insertSettingsChildInOrder`; leaving the stub child-free avoids
*  baking in flags the caller didn't request. */
const MINIMAL_SETTINGS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings xmlns:w="${w$2}"></w:settings>`;
/**
* Ensure word/settings.xml carries `<w:updateFields w:val="true"/>`, and
* stage the appropriate replacements / new-part registration into the
* replacement map the caller is about to flush.
*
* Idempotent: re-running on an already-flagged doc returns without
* touching anything.
*/
async function ensureUpdateFieldsFlag(reader, replacements, bodyAssetRegistry) {
	await mutateSettings(reader, replacements, bodyAssetRegistry, (root) => {
		let updateFields = firstChildNS(root, w$2, "updateFields");
		if (updateFields) {
			if (wAttr(updateFields, "val") === "true") return false;
			updateFields.setAttributeNS(w$2, "w:val", "true");
			return true;
		}
		updateFields = root.ownerDocument.createElementNS(w$2, "w:updateFields");
		updateFields.setAttributeNS(w$2, "w:val", "true");
		insertSettingsChild(root, updateFields, SETTINGS_AFTER_UPDATEFIELDS);
		return true;
	});
}
/**
* Set or clear `<w:evenAndOddHeaders/>` in settings.xml based on
* `enabled`. The flag is presence-only (no @val); Word treats the empty
* element as enabling distinct even/odd headers and footers. Pair with
* sectPrs that carry header/footer references with `w:type="even"`.
*
* Called by the HF pipeline as the source of truth: with HF config
* declared, the flag is enabled iff at least one surface declared an
* `even` variant. The fabrication path (no settings.xml in source)
* is skipped when `enabled=false` to avoid creating settings.xml just
* to remove a flag that wasn't there.
*
* Idempotent: re-run on an already-correct doc no-ops.
*/
async function setEvenAndOddHeadersFlag(reader, replacements, enabled, bodyAssetRegistry) {
	await mutateSettings(reader, replacements, bodyAssetRegistry, (root) => {
		const existing = firstChildNS(root, w$2, "evenAndOddHeaders");
		if (enabled) {
			if (existing) return false;
			insertSettingsChild(root, root.ownerDocument.createElementNS(w$2, "w:evenAndOddHeaders"), SETTINGS_AFTER_EVENANDODDHEADERS);
			return true;
		}
		if (!existing) return false;
		root.removeChild(existing);
		return true;
	}, { fabricateOnMissing: enabled });
}
/** Shared settings.xml mutation flow. Loads (or reuses an in-flight
*  replacement of) settings.xml, hands the root to `mutate`, and stages
*  the result. Fabricates a minimal settings.xml + registers the part
*  in Content_Types and the body rels when none exists.
*
*  `mutate` returns true when it changed the doc; false (already in
*  desired state) skips re-serialization but still falls through to the
*  fabricate path if no settings.xml existed at all. */
async function mutateSettings(reader, replacements, bodyAssetRegistry, mutate, opts = {}) {
	const fabricateOnMissing = opts.fabricateOnMissing ?? true;
	const pending = replacements.get("word/settings.xml");
	let settingsDoc = null;
	if (typeof pending === "string") settingsDoc = parseXml(pending);
	else settingsDoc = await reader.readXml("word/settings.xml");
	if (settingsDoc) {
		const root = settingsDoc.documentElement;
		if (root) {
			if (mutate(root)) replacements.set("word/settings.xml", serializeXml(settingsDoc));
			return;
		}
	}
	if (!fabricateOnMissing) return;
	const fabricated = parseXml(MINIMAL_SETTINGS_XML);
	const root = fabricated.documentElement;
	if (root) mutate(root);
	replacements.set("word/settings.xml", serializeXml(fabricated));
	bodyAssetRegistry.getContentTypes().ensureOverride("/word/settings.xml", SETTINGS_CT);
	const sharedRels = bodyAssetRegistry.getPartRels();
	if (!sharedRels.hasRelTo("settings.xml")) sharedRels.appendRel(SETTINGS_REL_TYPE, "settings.xml");
}
/** Per ECMA-376 §17.15.1, CT_Settings has a long fixed child order.
* Rather than enumerate the full ~80-element sequence, each ensure
* function declares the set of element local-names that come AFTER its
* own insertion — `insertSettingsChild` then places the new element
* before the first existing child found in that set. Pre-element
* children stay untouched; trailing position is preserved when none of
* the AFTER elements are present. */
const SETTINGS_AFTER_UPDATEFIELDS = new Set([
	"hdrShapeDefaults",
	"footnotePr",
	"endnotePr",
	"compat",
	"docVars",
	"rsids",
	"mathPr",
	"uiCompat97To2003",
	"attachedSchema",
	"themeFontLang",
	"clrSchemeMapping",
	"doNotIncludeSubdocsInStats",
	"doNotAutoCompressPictures",
	"forceUpgrade",
	"captions",
	"readModeInkLockDown",
	"smartTagType",
	"schemaLibrary",
	"shapeDefaults",
	"doNotEmbedSmartTags",
	"decimalSymbol",
	"listSeparator"
]);
/** Elements that come AFTER `<w:evenAndOddHeaders>` in CT_Settings.
* Superset of SETTINGS_AFTER_UPDATEFIELDS plus the run of mid-tail
* elements between defaultTableStyle (position ~47) and updateFields
* (position ~74). Without these, an insert against a source doc that
* already carries updateFields (which most do) misplaces the new
* evenAndOddHeaders AFTER updateFields, producing schema-invalid XML. */
const SETTINGS_AFTER_EVENANDODDHEADERS = new Set([
	"bookFoldRevPrinting",
	"bookFoldPrinting",
	"bookFoldPrintingSheets",
	"drawingGridHorizontalSpacing",
	"drawingGridVerticalSpacing",
	"displayHorizontalDrawingGridEvery",
	"displayVerticalDrawingGridEvery",
	"doNotUseMarginsForDrawingGridOrigin",
	"drawingGridHorizontalOrigin",
	"drawingGridVerticalOrigin",
	"doNotShadeFormData",
	"noPunctuationKerning",
	"characterSpacingControl",
	"printTwoOnOne",
	"strictFirstAndLastChars",
	"noLineBreaksAfter",
	"noLineBreaksBefore",
	"savePreviewPicture",
	"doNotValidateAgainstSchema",
	"saveInvalidXml",
	"ignoreMixedContent",
	"alwaysShowPlaceholderText",
	"doNotDemarcateInvalidXml",
	"saveXmlDataOnly",
	"useXSLTWhenSaving",
	"saveThroughXslt",
	"showXMLTags",
	"alwaysMergeEmptyNamespace",
	"updateStyles",
	"updateFields",
	...SETTINGS_AFTER_UPDATEFIELDS
]);
function insertSettingsChild(root, newEl, afterSet) {
	for (const el of getChildren(root)) if (el.namespaceURI === w$2 && afterSet.has(el.localName)) {
		root.insertBefore(newEl, el);
		return;
	}
	root.appendChild(newEl);
}
/**
* Full CT_Settings child sequence, extracted from `wml.xsd` (98 elements,
* every one optional). The per-call AFTER-sets above cover the two elements
* the engine writes implicitly; this table is what the generic
* `settings.elements[]` channel uses, so a raw settings write can be
* positioned correctly without the caller having to know the sequence.
*
* Note `trackRevisions` (not `trackChanges`) — the config-facing name and
* the OOXML element name differ, and the schema is authoritative here.
*/
const SETTINGS_CHILD_ORDER = [
	"writeProtection",
	"view",
	"zoom",
	"removePersonalInformation",
	"removeDateAndTime",
	"doNotDisplayPageBoundaries",
	"displayBackgroundShape",
	"printPostScriptOverText",
	"printFractionalCharacterWidth",
	"printFormsData",
	"embedTrueTypeFonts",
	"embedSystemFonts",
	"saveSubsetFonts",
	"saveFormsData",
	"mirrorMargins",
	"alignBordersAndEdges",
	"bordersDoNotSurroundHeader",
	"bordersDoNotSurroundFooter",
	"gutterAtTop",
	"hideSpellingErrors",
	"hideGrammaticalErrors",
	"activeWritingStyle",
	"proofState",
	"formsDesign",
	"attachedTemplate",
	"linkStyles",
	"stylePaneFormatFilter",
	"stylePaneSortMethod",
	"documentType",
	"mailMerge",
	"revisionView",
	"trackRevisions",
	"doNotTrackMoves",
	"doNotTrackFormatting",
	"documentProtection",
	"autoFormatOverride",
	"styleLockTheme",
	"styleLockQFSet",
	"defaultTabStop",
	"autoHyphenation",
	"consecutiveHyphenLimit",
	"hyphenationZone",
	"doNotHyphenateCaps",
	"showEnvelope",
	"summaryLength",
	"clickAndTypeStyle",
	"defaultTableStyle",
	"evenAndOddHeaders",
	"bookFoldRevPrinting",
	"bookFoldPrinting",
	"bookFoldPrintingSheets",
	"drawingGridHorizontalSpacing",
	"drawingGridVerticalSpacing",
	"displayHorizontalDrawingGridEvery",
	"displayVerticalDrawingGridEvery",
	"doNotUseMarginsForDrawingGridOrigin",
	"drawingGridHorizontalOrigin",
	"drawingGridVerticalOrigin",
	"doNotShadeFormData",
	"noPunctuationKerning",
	"characterSpacingControl",
	"printTwoOnOne",
	"strictFirstAndLastChars",
	"noLineBreaksAfter",
	"noLineBreaksBefore",
	"savePreviewPicture",
	"doNotValidateAgainstSchema",
	"saveInvalidXml",
	"ignoreMixedContent",
	"alwaysShowPlaceholderText",
	"doNotDemarcateInvalidXml",
	"saveXmlDataOnly",
	"useXSLTWhenSaving",
	"saveThroughXslt",
	"showXMLTags",
	"alwaysMergeEmptyNamespace",
	"updateFields",
	"hdrShapeDefaults",
	"footnotePr",
	"endnotePr",
	"compat",
	"docVars",
	"rsids",
	"mathPr",
	"attachedSchema",
	"themeFontLang",
	"clrSchemeMapping",
	"doNotIncludeSubdocsInStats",
	"doNotAutoCompressPictures",
	"forceUpgrade",
	"captions",
	"readModeInkLockDown",
	"smartTagType",
	"schemaLibrary",
	"shapeDefaults",
	"doNotEmbedSmartTags",
	"decimalSymbol",
	"listSeparator"
];
const SETTINGS_ORDER_INDEX = /* @__PURE__ */ new Map(SETTINGS_CHILD_ORDER.map((n, i) => [n, i]));
/** Place `el` at its CT_Settings position. Elements the sequence doesn't
*  know (Microsoft extensions like `w15:chartTrackingRefBased`) append —
*  Word treats them as ignorable and XSD validation strips them anyway. */
function insertSettingsChildByOrder(root, el) {
	const target = SETTINGS_ORDER_INDEX.get(el.localName);
	if (target !== void 0) for (const child of getChildren(root)) {
		if (child.namespaceURI !== w$2) continue;
		const idx = SETTINGS_ORDER_INDEX.get(child.localName);
		if (idx !== void 0 && idx > target) {
			root.insertBefore(el, child);
			return;
		}
	}
	root.appendChild(el);
}
/**
* Upsert one settings element from a `settings.elements[]` entry.
*
* CT_Settings children are single-occurrence, so a second `<w:updateFields>`
* makes the part schema-invalid — this merges into the existing element
* instead of appending a duplicate. Attributes are overwritten (declared
* value wins), text content is replaced, and declared children are upserted
* by local name one level deep (enough for `mailMerge`, `docVars`, `compat`,
* `rsids` — the CT_Settings children whose content models have children).
* Declaring the same child name more than once (e.g. three `docVar`s inside
* `docVars`) appends a repeatable node per declaration instead of collapsing
* them into a single node.
*
* Returns "created" | "updated".
*/
function upsertSettingsElement(root, spec) {
	const doc = root.ownerDocument;
	let el = firstChildNS(root, w$2, spec.name);
	const created = !el;
	if (!el) el = doc.createElementNS(w$2, `w:${spec.name}`);
	for (const [rawName, rawVal] of Object.entries(spec.attrs ?? {})) {
		const [attrName, nsURI] = resolveAttrNs(rawName);
		el.setAttributeNS(nsURI, attrName, String(rawVal));
	}
	if (spec.text !== void 0) {
		while (el.firstChild) el.removeChild(el.firstChild);
		el.appendChild(doc.createTextNode(String(spec.text)));
	}
	const declaredChildCount = /* @__PURE__ */ new Map();
	for (const c of spec.children ?? []) declaredChildCount.set(c.name, (declaredChildCount.get(c.name) ?? 0) + 1);
	const usedChildren = /* @__PURE__ */ new Map();
	for (const childSpec of spec.children ?? []) {
		const name = childSpec.name;
		const wanted = declaredChildCount.get(name) ?? 1;
		let child = null;
		if (wanted > 1) {
			// Repeatable child (`docVar`, `rsid`, `listItem`…): upsert by key,
			// not by position. A declared identity attribute (`name` / `id`)
			// selects the node to update; when nothing matches, a new node is
			// appended — otherwise three `docVar`s would silently overwrite the
			// first three existing variables instead of adding to them.
			const existing = getChildren(el).filter((n) => n.namespaceURI === w$2 && n.localName === name);
			let match = null;
			let keyed = false;
			for (const attr of ["name", "id"]) {
				const declaredVal = (childSpec.attrs ?? {})[attr] ?? (childSpec.attrs ?? {})[`w:${attr}`];
				if (declaredVal === void 0) continue;
				keyed = true;
				match = existing.find((n) => n.getAttributeNS(w$2, attr) === String(declaredVal)) ?? null;
				break;
			}
			if (!keyed) {
				const nth = usedChildren.get(name) ?? 0;
				usedChildren.set(name, nth + 1);
				match = existing[nth] ?? null;
			}
			child = match;
		} else {
			child = firstChildNS(el, w$2, name);
		}
		if (!child) {
			child = doc.createElementNS(w$2, `w:${name}`);
			el.appendChild(child);
		}
		for (const [rawName, rawVal] of Object.entries(childSpec.attrs ?? {})) {
			const [attrName, nsURI] = resolveAttrNs(rawName);
			child.setAttributeNS(nsURI, attrName, String(rawVal));
		}
		if (childSpec.text !== void 0) {
			while (child.firstChild) child.removeChild(child.firstChild);
			child.appendChild(doc.createTextNode(String(childSpec.text)));
		}
	}
	if (created) insertSettingsChildByOrder(root, el);
	return created ? "created" : "updated";
}
/** Normalise a config-declared attribute name to `[qualifiedName, nsURI]`.
*  Bare names default to the `w:` namespace; an unknown prefix is folded
*  back to `w:` rather than emitted undeclared (which would fail XSD). */
function resolveAttrNs(attrName) {
	const colon = attrName.indexOf(":");
	if (colon < 0) return [`w:${attrName}`, NS.w];
	const prefix = attrName.slice(0, colon);
	const ns = NS[prefix];
	if (!ns) return [`w:${attrName.slice(colon + 1)}`, NS.w];
	return [attrName, ns];
}
/**
* Apply the optional `settings.elements[]` config channel: arbitrary
* `word/settings.xml` children (print flags, proofing state, mail-merge
* setup, document variables, compat settings…), each upserted at its
* schema position. The engine's implicit writers (`updateFields`,
* `evenAndOddHeaders`) run through the same upsert path, so declaring the
* same element here updates it rather than duplicating it.
*/
async function applySettingsElements(reader, replacements, bodyAssetRegistry, elements) {
	const report = [];
	await mutateSettings(reader, replacements, bodyAssetRegistry, (root) => {
		for (const spec of elements) report.push({
			name: spec.name,
			action: upsertSettingsElement(root, spec)
		});
		return report.length > 0;
	});
	return report;
}

//#endregion
//#region lib/apply/doc-props.ts
/**
* `docProps/core.xml` + `docProps/app.xml`.
*
* Why the engine writes them: with only `docProps/custom.xml` (what a
* slot-styled package used to ship), `AUTHOR` / `NUMWORDS` /
* `DOCPROPERTY` fields resolve to nothing, the file-properties pane shows
* an empty Title/Author/word-count, and Explorer's "Word count" column is
* blank. Word always writes both parts.
*
* Both are read-modify-write: an existing part keeps every field the
* engine doesn't manage, managed fields are updated in place at their
* schema position (CT_CoreProperties and CT_Properties are both sequences,
* so a misordered element fails XSD).
*/
const CORE_PROPS_CT = "application/vnd.openxmlformats-package.core-properties+xml";
const EXTENDED_PROPS_CT = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
const CORE_PROPS_REL = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const EXTENDED_PROPS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
const EXTENDED_PROPS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
const DCTERMS_NS = "http://purl.org/dc/terms/";
const DC_NS = NS.dc;
const CP_NS = NS.cp;
const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
/** CT_CoreProperties child sequence (15 elements, all optional). */
const CORE_ORDER_INDEX = /* @__PURE__ */ new Map([
	"category",
	"contentStatus",
	"created",
	"creator",
	"description",
	"identifier",
	"keywords",
	"language",
	"lastModifiedBy",
	"lastPrinted",
	"modified",
	"revision",
	"subject",
	"title",
	"version"
].map((n, i) => [n, i]));
/** CT_Properties child sequence (27 elements, all optional). */
const EXTENDED_ORDER_INDEX = /* @__PURE__ */ new Map([
	"Template",
	"Manager",
	"Company",
	"Pages",
	"Words",
	"Characters",
	"PresentationFormat",
	"Lines",
	"Paragraphs",
	"Slides",
	"Notes",
	"TotalTime",
	"HiddenSlides",
	"MMClips",
	"ScaleCrop",
	"HeadingPairs",
	"TitlesOfParts",
	"LinksUpToDate",
	"CharactersWithSpaces",
	"SharedDoc",
	"HyperlinkBase",
	"HLinks",
	"HyperlinksChanged",
	"DigSig",
	"Application",
	"AppVersion",
	"DocSecurity"
].map((n, i) => [n, i]));
/** Place `el` at its position in `orderIndex`. Unknown names append — the
*  only way a caller reaches that branch is a hand-built part. */
function insertChildByOrder(parent, el, orderIndex) {
	const target = orderIndex.get(el.localName);
	if (target !== undefined) for (const child of getChildren(parent)) {
		const idx = orderIndex.get(child.localName);
		if (idx !== undefined && idx > target) {
			parent.insertBefore(el, child);
			return;
		}
	}
	parent.appendChild(el);
}
const CJK_CHAR_RE = /[\u3400-\u9fff\uf900-\ufaff]/g;
/** Body statistics for `<Words>` / `<Characters>` / `<Paragraphs>` /
*  `<Lines>`. Word counts a CJK character as one word and a run of Latin
*  letters/digits as one word — mirrored here so the number the properties
*  pane shows matches what Word computes after opening. */
function countBodyStats(documentDoc) {
	let paragraphs = 0;
	let cjkChars = 0;
	let latinWords = 0;
	let charsNoSpaces = 0;
	let charsWithSpaces = 0;
	for (const p of descendantsNS(documentDoc, NS.w, "p")) {
		paragraphs += 1;
		let text = "";
		for (const t of descendantsNS(p, NS.w, "t")) text += textContent(t) ?? "";
		cjkChars += (text.match(CJK_CHAR_RE) ?? []).length;
		latinWords += (text.replace(CJK_CHAR_RE, " ").match(/[A-Za-z0-9][A-Za-z0-9'’._-]*/g) ?? []).length;
		charsWithSpaces += text.length;
		charsNoSpaces += text.replace(/\s/g, "").length;
	}
	return {
		paragraphs,
		words: cjkChars + latinWords,
		characters: charsNoSpaces,
		charactersWithSpaces: charsWithSpaces,
		lines: paragraphs
	};
}
async function readPartText(reader, replacements, path) {
	const pending = replacements.get(path);
	if (typeof pending === "string") return pending;
	try {
		return await reader.readText(path);
	} catch {
		return null;
	}
}
function setOrCreateTextChild(doc, parent, ns, qname, localName, orderIndex, value) {
	let el = firstChildNS(parent, ns, localName);
	if (!el) {
		el = doc.createElementNS(ns, qname);
		parent.appendChild(el);
		if (orderIndex) {
			parent.removeChild(el);
			insertChildByOrder(parent, el, orderIndex);
		}
	}
	el.textContent = String(value);
	return el;
}
/**
* Write/refresh `docProps/core.xml` and `docProps/app.xml`, register both as
* package parts when they're new (content type + `_rels/.rels`
* relationship), and recompute the body counters.
*
* `creator` is only applied when the part is new or still carries the
* engine's own previous value — an author set by a human in Word outranks a
* tool default.
*/
async function ensureDocProps(reader, replacements, bodyAssetRegistry, documentDoc, opts = {}) {
	const stats = countBodyStats(documentDoc);
	const stamp = opts.timestamp ?? new Date().toISOString().replace(/\.\d{3}Z$/, "Z");
	const revisionStep = opts.bumpRevision !== false;
	const coreRaw = await readPartText(reader, replacements, "docProps/core.xml");
	let coreDoc;
	let coreCreated = false;
	if (coreRaw) coreDoc = parseXml(coreRaw);
	else {
		coreDoc = parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="${NS.cp}" xmlns:dc="${NS.dc}" xmlns:dcterms="${DCTERMS_NS}" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"/>`);
		coreCreated = true;
	}
	const coreRoot = coreDoc?.documentElement;
	if (!coreRoot) return null;
	const existingCreator = firstChildNS(coreRoot, DC_NS, "creator")?.textContent ?? "";
	const toolName = opts.application ?? "docx-master";
	const creator = existingCreator.trim() && existingCreator !== toolName ? existingCreator : opts.creator ?? toolName;
	setOrCreateTextChild(coreDoc, coreRoot, DCTERMS_NS, "dcterms:created", "created", CORE_ORDER_INDEX, firstChildNS(coreRoot, DCTERMS_NS, "created")?.textContent || stamp);
	setOrCreateTextChild(coreDoc, coreRoot, DC_NS, "dc:creator", "creator", CORE_ORDER_INDEX, creator);
	setOrCreateTextChild(coreDoc, coreRoot, CP_NS, "cp:lastModifiedBy", "lastModifiedBy", CORE_ORDER_INDEX, opts.lastModifiedBy ?? creator);
	setOrCreateTextChild(coreDoc, coreRoot, DCTERMS_NS, "dcterms:modified", "modified", CORE_ORDER_INDEX, stamp);
	const revisionPrev = parseInt(firstChildNS(coreRoot, CP_NS, "revision")?.textContent ?? "0", 10);
	const revision = revisionStep ? String((Number.isFinite(revisionPrev) ? revisionPrev : 0) + 1) : String(revisionPrev || 1);
	setOrCreateTextChild(coreDoc, coreRoot, CP_NS, "cp:revision", "revision", CORE_ORDER_INDEX, revision);
	if (opts.title) setOrCreateTextChild(coreDoc, coreRoot, DC_NS, "dc:title", "title", CORE_ORDER_INDEX, opts.title);
	replacements.set("docProps/core.xml", serializeXml(coreDoc));
	const appRaw = await readPartText(reader, replacements, "docProps/app.xml");
	let appDoc;
	let appCreated = false;
	if (appRaw) appDoc = parseXml(appRaw);
	else {
		appDoc = parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="${EXTENDED_PROPS_NS}" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>`);
		appCreated = true;
	}
	const appRoot = appDoc?.documentElement;
	if (appRoot) {
		setOrCreateTextChild(appDoc, appRoot, EXTENDED_PROPS_NS, "Words", "Words", EXTENDED_ORDER_INDEX, stats.words);
		setOrCreateTextChild(appDoc, appRoot, EXTENDED_PROPS_NS, "Characters", "Characters", EXTENDED_ORDER_INDEX, stats.characters);
		setOrCreateTextChild(appDoc, appRoot, EXTENDED_PROPS_NS, "CharactersWithSpaces", "CharactersWithSpaces", EXTENDED_ORDER_INDEX, stats.charactersWithSpaces);
		setOrCreateTextChild(appDoc, appRoot, EXTENDED_PROPS_NS, "Lines", "Lines", EXTENDED_ORDER_INDEX, stats.lines);
		setOrCreateTextChild(appDoc, appRoot, EXTENDED_PROPS_NS, "Paragraphs", "Paragraphs", EXTENDED_ORDER_INDEX, stats.paragraphs);
		setOrCreateTextChild(appDoc, appRoot, EXTENDED_PROPS_NS, "Application", "Application", EXTENDED_ORDER_INDEX, toolName);
		replacements.set("docProps/app.xml", serializeXml(appDoc));
	}
	const createdPaths = [];
	if (coreCreated) createdPaths.push({
		path: "docProps/core.xml",
		ct: CORE_PROPS_CT,
		rel: CORE_PROPS_REL
	});
	if (appCreated) createdPaths.push({
		path: "docProps/app.xml",
		ct: EXTENDED_PROPS_CT,
		rel: EXTENDED_PROPS_REL
	});
	for (const entry of createdPaths) bodyAssetRegistry.getContentTypes().ensureOverride(`/${entry.path}`, entry.ct);
	if (createdPaths.length > 0) await registerPackageRels(reader, replacements, createdPaths);
	return {
		words: stats.words,
		paragraphs: stats.paragraphs,
		revision,
		coreCreated,
		appCreated
	};
}
/** Add the package-level (`_rels/.rels`) relationships for parts that were
*  fabricated in this run. Existing relationships are left alone — a source
*  document that already points at its core/app properties keeps its own
*  rIds. */
async function registerPackageRels(reader, replacements, entries) {
	const raw = await readPartText(reader, replacements, "_rels/.rels");
	if (!raw) return;
	const doc = parseXml(raw);
	const root = doc?.documentElement;
	if (!root) return;
	const existing = new Set(getChildrenNS(root, PKG_REL_NS, "Relationship").map((r) => r.getAttribute("Type")));
	let next = 0;
	for (const rel of getChildrenNS(root, PKG_REL_NS, "Relationship")) {
		const id = rel.getAttribute("Id") ?? "";
		const m = /^rId(\d+)$/.exec(id);
		if (m) next = Math.max(next, parseInt(m[1], 10));
	}
	let changed = false;
	for (const entry of entries) {
		if (existing.has(entry.rel)) continue;
		next += 1;
		const el = doc.createElementNS(PKG_REL_NS, "Relationship");
		el.setAttribute("Id", `rId${next}`);
		el.setAttribute("Type", entry.rel);
		el.setAttribute("Target", entry.path);
		root.appendChild(el);
		changed = true;
	}
	if (changed) replacements.set("_rels/.rels", serializeXml(doc));
}

//#endregion
//#region lib/apply/section-selector.ts
/**
* Section selectors — `"N"` or `"N-M"` (1-based, inclusive).
*
* Shared by `pageSetup.sections` and `headerFooter.sections`. Schema-level
* validation uses `SECTION_SELECTOR_KEY_RE`; runtime expansion to concrete
* section indices uses `expandSectionSelector`, which throws with a
* caller-supplied `fieldPath` so the error names the config root the user
* wrote (e.g. `pageSetup.sections` vs `headerFooter.sections`).
*/
const SECTION_SELECTOR_KEY_RE = /^\d+(?:-\d+)?$/;
function expandSectionSelector(key, sectionCount, fieldPath) {
	const range = key.match(/^(\d+)-(\d+)$/);
	if (range) {
		const lo = parseInt(range[1], 10);
		const hi = parseInt(range[2], 10);
		if (lo < 1 || hi > sectionCount || lo > hi) throw new Error(`${fieldPath}: range "${key}" out of bounds. Document has ${sectionCount} section(s); valid 1..${sectionCount}.`);
		return Array.from({ length: hi - lo + 1 }, (_, i) => lo + i);
	}
	const n = parseInt(key, 10);
	if (n < 1 || n > sectionCount) throw new Error(`${fieldPath}: section ${n} out of bounds. Document has ${sectionCount} section(s); valid 1..${sectionCount}.`);
	return [n];
}

//#endregion
//#region lib/apply/page-setup-mutation.ts
/**
* Page-setup mutation: `<w:sectPr>` children — `pgSz` / `pgMar` / `cols`.
*
* Sparse-by-design: only declared fields mutate. Margins are merged per-edge
* (declaring `top` doesn't disturb `bottom` etc.). Sections inherit the
* top-level defaults and selectively override via `sections.<selector>`,
* where the selector is `"all"`, `"N"`, or `"N-M"` (1-based, inclusive).
*
* Walking order matches `document-parser.ts`: paragraph-embedded sectPrs in
* document order (sections 1..N-1), then the body-trailing sectPr (section
* N). The Nth element of `collectSectPrs(body)` corresponds to section N
* (1-based), so `SectionInfo[]` indices map directly.
*/
const w$1 = NS.w;
/** Paper size constants in twips. Portrait orientation (width < height).
*  Engine swaps when orientation is landscape. */
const PAPER_SIZE_TWIPS = {
	A3: [16838, 23811],
	A4: [11906, 16838],
	A5: [8392, 11906],
	Letter: [12240, 15840],
	Legal: [12240, 20160],
	B5: [9978, 14173],
	"16K": [10489, 14741]
};
const DEFAULT_COLUMN_SPACE_TWIPS = toTwips("0.5cm", "columns.space");
/** ±1 twip absorbs cm/in round-trip noise (Word stores margins in twips
*  but UI sets them in cm/mm; the same logical length can round to either
*  side of a fractional twip). A 1-twip difference is 0.05mm — far below
*  Word's display granularity. Used by mutatePgSz / mutatePgMar to decide
*  whether a declared value differs from the source meaningfully. */
const TWIP_EQ_TOLERANCE = 1;
/** Walk body in document order, collect every sectPr element. Order matches
*  the parser's `SectionInfo[]` index (1-based). */
function collectSectPrs$1(body) {
	const out = [];
	for (const child of getChildren(body)) {
		if (child.namespaceURI !== w$1) continue;
		if (child.localName === "p") {
			const pPr = firstChildNS(child, w$1, "pPr");
			if (pPr) {
				const sectPr = firstChildNS(pPr, w$1, "sectPr");
				if (sectPr) out.push(sectPr);
			}
		} else if (child.localName === "sectPr") out.push(child);
	}
	return out;
}
/** Build the effective per-section config by layering per-section overrides
*  on top of the top-level defaults. Margins merge field-wise; paperSize /
*  orientation / columns replace wholesale. Returns one entry per section
*  (1-based; entry at index 0 is unused). */
function buildEffectiveConfigs(config, sectionCount) {
	const defaults = {
		paperSize: config.paperSize,
		orientation: config.orientation,
		margins: config.margins,
		columns: config.columns,
		pgNumType: config.pgNumType
	};
	const effective = [];
	for (let i = 0; i < sectionCount; i++) effective.push({
		...defaults,
		margins: defaults.margins ? { ...defaults.margins } : void 0,
		pgNumType: defaults.pgNumType ? { ...defaults.pgNumType } : void 0
	});
	if (!config.sections) return effective;
	for (const [key, override] of Object.entries(config.sections)) {
		const indices = expandSectionSelector(key, sectionCount, "pageSetup.sections");
		for (const idx1 of indices) {
			const slot = effective[idx1 - 1];
			if (override.paperSize !== void 0) slot.paperSize = override.paperSize;
			if (override.orientation !== void 0) slot.orientation = override.orientation;
			if (override.columns !== void 0) slot.columns = override.columns;
			if (override.margins !== void 0) slot.margins = {
				...slot.margins ?? {},
				...override.margins
			};
			if (override.pgNumType !== void 0) slot.pgNumType = {
				...slot.pgNumType ?? {},
				...override.pgNumType
			};
		}
	}
	return effective;
}
/** Resolve a PaperSize input to its (width, height) in twips, portrait
*  orientation. Custom `{width, height}` falls through verbatim. */
function paperSizeToTwips(paper) {
	if (typeof paper === "string") {
		const dims = PAPER_SIZE_TWIPS[paper];
		if (!dims) throw new Error(`pageSetup.paperSize: unknown size "${paper}"`);
		return {
			w: dims[0],
			h: dims[1]
		};
	}
	return {
		w: toTwips(paper.width, "paperSize.width"),
		h: toTwips(paper.height, "paperSize.height")
	};
}
/** Build / replace `<w:pgSz>` on a sectPr. When paperSize is declared the
*  result is (constants or custom) optionally swapped for landscape; when
*  only orientation is declared the existing w/h are read and swapped to
*  match the requested direction. Returns true when anything changed. */
function mutatePgSz(sectPr, paperSize, orientation, doc) {
	if (paperSize === void 0 && orientation === void 0) return false;
	const existing = firstChildNS(sectPr, w$1, "pgSz");
	const curW = existing ? parseInt(existing.getAttributeNS(w$1, "w") || "0", 10) : 0;
	const curH = existing ? parseInt(existing.getAttributeNS(w$1, "h") || "0", 10) : 0;
	const curOrient = existing?.getAttributeNS(w$1, "orient") === "landscape" ? "landscape" : curW > 0 && curH > 0 && curW > curH ? "landscape" : "portrait";
	let portraitW;
	let portraitH;
	if (paperSize !== void 0) {
		const dims = paperSizeToTwips(paperSize);
		portraitW = dims.w;
		portraitH = dims.h;
	} else {
		if (curW === 0 || curH === 0) throw new Error(`pageSetup.orientation: section has no existing <w:pgSz> to pivot from; declare paperSize alongside orientation.`);
		portraitW = Math.min(curW, curH);
		portraitH = Math.max(curW, curH);
	}
	const targetOrient = orientation ?? curOrient;
	const [outW, outH] = targetOrient === "landscape" ? [portraitH, portraitW] : [portraitW, portraitH];
	if (existing && Math.abs(outW - curW) <= TWIP_EQ_TOLERANCE && Math.abs(outH - curH) <= TWIP_EQ_TOLERANCE && targetOrient === "landscape" === (existing.getAttributeNS(w$1, "orient") === "landscape")) return false;
	let pgSz = existing;
	if (!pgSz) {
		pgSz = doc.createElementNS(w$1, "w:pgSz");
		insertChildInOrder(sectPr, pgSz, SECT_PR_CHILD_ORDER);
	}
	pgSz.setAttributeNS(w$1, "w:w", String(outW));
	pgSz.setAttributeNS(w$1, "w:h", String(outH));
	if (targetOrient === "landscape") pgSz.setAttributeNS(w$1, "w:orient", "landscape");
	else pgSz.removeAttributeNS(w$1, "orient");
	return true;
}
const MARGIN_ATTR_KEYS = [
	{
		field: "top",
		attr: "top"
	},
	{
		field: "right",
		attr: "right"
	},
	{
		field: "bottom",
		attr: "bottom"
	},
	{
		field: "left",
		attr: "left"
	},
	{
		field: "header",
		attr: "header"
	},
	{
		field: "footer",
		attr: "footer"
	},
	{
		field: "gutter",
		attr: "gutter"
	}
];
function mutatePgMar(sectPr, margins, doc) {
	if (!margins) return false;
	let pgMar = firstChildNS(sectPr, w$1, "pgMar");
	const fresh = !pgMar;
	if (!pgMar) {
		pgMar = doc.createElementNS(w$1, "w:pgMar");
		insertChildInOrder(sectPr, pgMar, SECT_PR_CHILD_ORDER);
	}
	let changed = fresh;
	for (const { field, attr } of MARGIN_ATTR_KEYS) {
		const v = margins[field];
		if (v === void 0) continue;
		const next = toTwips(v, `margins.${field}`);
		const cur = parseInt(pgMar.getAttributeNS(w$1, attr) || "", 10);
		if (Number.isNaN(cur) || Math.abs(cur - next) > TWIP_EQ_TOLERANCE) {
			pgMar.setAttributeNS(w$1, `w:${attr}`, String(next));
			changed = true;
		}
	}
	return changed;
}
function mutateCols(sectPr, columns, doc) {
	if (columns === void 0) return false;
	let count;
	let space;
	let separator;
	let widths;
	let spaces;
	if (typeof columns === "number") count = columns;
	else if (columns.widths !== void 0) {
		widths = columns.widths.map((v, i) => toTwips(v, `columns.widths[${i}]`));
		count = widths.length;
		if (columns.spaces !== void 0) spaces = columns.spaces.map((v, i) => toTwips(v, `columns.spaces[${i}]`));
		separator = columns.separator;
	} else {
		count = columns.count;
		space = columns.space !== void 0 ? toTwips(columns.space, "columns.space") : void 0;
		separator = columns.separator;
	}
	const cols = doc.createElementNS(w$1, "w:cols");
	cols.setAttributeNS(w$1, "w:num", String(count));
	if (widths === void 0) {
		cols.setAttributeNS(w$1, "w:space", String(space ?? DEFAULT_COLUMN_SPACE_TWIPS));
		cols.setAttributeNS(w$1, "w:equalWidth", "true");
	} else {
		cols.setAttributeNS(w$1, "w:equalWidth", "false");
		for (let i = 0; i < widths.length; i++) {
			const col = doc.createElementNS(w$1, "w:col");
			col.setAttributeNS(w$1, "w:w", String(widths[i]));
			if (i < widths.length - 1) {
				const sp = spaces ? spaces[i] : DEFAULT_COLUMN_SPACE_TWIPS;
				col.setAttributeNS(w$1, "w:space", String(sp));
			}
			cols.appendChild(col);
		}
	}
	if (separator !== void 0) cols.setAttributeNS(w$1, "w:sep", separator ? "true" : "false");
	const existing = firstChildNS(sectPr, w$1, "cols");
	if (existing && colsElementsEqual(existing, cols)) return false;
	if (existing) sectPr.removeChild(existing);
	insertChildInOrder(sectPr, cols, SECT_PR_CHILD_ORDER);
	return true;
}
/** Structural comparison for two `<w:cols>` elements (built or DOM). Returns
*  true when the attribute set (num/space/equalWidth/sep) and `<w:col>`
*  children match. Skips ECMA-376 attributes we don't author (none today),
*  so a doc with foreign attrs would falsely diff — acceptable since no-op
*  detection is an optimization, not a correctness invariant. */
function colsElementsEqual(a, b) {
	for (const at of [
		"num",
		"space",
		"equalWidth",
		"sep"
	]) if (a.getAttributeNS(w$1, at) !== b.getAttributeNS(w$1, at)) return false;
	const aCols = Array.from(getChildren(a)).filter((c) => c.namespaceURI === w$1 && c.localName === "col");
	const bCols = Array.from(getChildren(b)).filter((c) => c.namespaceURI === w$1 && c.localName === "col");
	if (aCols.length !== bCols.length) return false;
	for (let i = 0; i < aCols.length; i++) {
		if (aCols[i].getAttributeNS(w$1, "w") !== bCols[i].getAttributeNS(w$1, "w")) return false;
		if (aCols[i].getAttributeNS(w$1, "space") !== bCols[i].getAttributeNS(w$1, "space")) return false;
	}
	return true;
}
/** Build / replace `<w:pgNumType>` on a sectPr. Idempotent on identical
*  attribute sets so re-runs don't bump touchedCount. */
function mutatePgNumType(sectPr, pgNumType, doc) {
	if (pgNumType === void 0) return false;
	const existing = firstChildNS(sectPr, w$1, "pgNumType");
	const curFmt = existing?.getAttributeNS(w$1, "fmt") ?? null;
	const curStart = existing?.getAttributeNS(w$1, "start") ?? null;
	const wantFmt = pgNumType.fmt ?? null;
	const wantStart = pgNumType.start !== void 0 ? String(pgNumType.start) : null;
	if (existing && curFmt === wantFmt && curStart === wantStart) return false;
	if (existing) sectPr.removeChild(existing);
	const el = doc.createElementNS(w$1, "w:pgNumType");
	if (wantFmt !== null) el.setAttributeNS(w$1, "w:fmt", wantFmt);
	if (wantStart !== null) el.setAttributeNS(w$1, "w:start", wantStart);
	insertChildInOrder(sectPr, el, SECT_PR_CHILD_ORDER);
	return true;
}
/** Apply pageSetup to every relevant sectPr in the document. Returns a
*  per-section before/after snapshot for the dry-run report. */
function applyPageSetup(documentDoc, config) {
	const body = firstChildNS(documentDoc.documentElement, w$1, "body");
	if (!body) return {
		sections: [],
		touchedCount: 0
	};
	const sectPrs = collectSectPrs$1(body);
	const sectionCount = sectPrs.length;
	if (sectionCount === 0) return {
		sections: [],
		touchedCount: 0
	};
	const effective = buildEffectiveConfigs(config, sectionCount);
	const report = [];
	let touched = 0;
	for (let i = 0; i < sectionCount; i++) {
		const sectPr = sectPrs[i];
		const fields = effective[i];
		const before = snapshotSection(sectPr);
		const a = mutatePgSz(sectPr, fields.paperSize, fields.orientation, documentDoc);
		const b = mutatePgMar(sectPr, fields.margins, documentDoc);
		const c = mutateCols(sectPr, fields.columns, documentDoc);
		const d = mutatePgNumType(sectPr, fields.pgNumType, documentDoc);
		const changed = a || b || c || d;
		const after = snapshotSection(sectPr);
		if (changed) touched++;
		report.push({
			index: i + 1,
			before,
			after,
			changed
		});
	}
	return {
		sections: report,
		touchedCount: touched
	};
}
/** Read a section's display-relevant attributes from the current DOM state.
*  Called twice per section (before + after mutation); margin attrs default
*  to 0 when no <w:pgMar> is present. */
function snapshotSection(sectPr) {
	const pgSz = firstChildNS(sectPr, w$1, "pgSz");
	const pgMar = firstChildNS(sectPr, w$1, "pgMar");
	const widthTwips = pgSz ? parseInt(pgSz.getAttributeNS(w$1, "w") || "0", 10) : 0;
	const heightTwips = pgSz ? parseInt(pgSz.getAttributeNS(w$1, "h") || "0", 10) : 0;
	const orient = pgSz?.getAttributeNS(w$1, "orient") === "landscape" ? "landscape" : widthTwips > heightTwips && widthTwips > 0 ? "landscape" : "portrait";
	const readMar = (attr) => pgMar ? parseInt(pgMar.getAttributeNS(w$1, attr) || "0", 10) : 0;
	return {
		paperSize: matchPaperLabel(widthTwips, heightTwips),
		orientation: orient,
		margins: {
			top: readMar("top"),
			bottom: readMar("bottom"),
			left: readMar("left"),
			right: readMar("right")
		}
	};
}
/** Map a twips dimension pair back to a paper size label for display. Both
*  orientations match; tolerance of 1 twip absorbs rounding. */
function matchPaperLabel(width, height) {
	if (width <= 0 || height <= 0) return void 0;
	const portraitW = Math.min(width, height);
	const portraitH = Math.max(width, height);
	for (const [label, [pw, ph]] of Object.entries(PAPER_SIZE_TWIPS)) if (Math.abs(portraitW - pw) <= 1 && Math.abs(portraitH - ph) <= 1) return label;
}

//#endregion
//#region lib/apply/header-footer-mutation.ts
const w = NS.w;
const VARIANT_ORDER = [
	"default",
	"first",
	"even"
];
const SURFACE_ORDER = ["header", "footer"];
const HEADER_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
const FOOTER_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
const HEADER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
const FOOTER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
/** Attach a paragraph border to the endpoint paragraph of an HF part:
*  header → last child `<w:p>` (bottom edge); footer → first child `<w:p>`
*  (top edge). Returns `true` when attached, `false` when the endpoint is
*  not a `<w:p>` (table-only or empty variant) — caller aggregates
*  effectiveness across a surface's variants to decide whether to warn.
*  If a `<w:pBdr>` already exists (e.g. variant ends with a
*  `horizontal-rule` block), the user's HF-level separator wins on the
*  relevant edge — duplicate edge children would be schema-invalid. */
function attachSeparator(root, edge, separator, ownerDoc) {
	const target = edge === "bottom" ? findLastChildEl(root, w, "p") : findFirstChildEl(root, w, "p");
	if (!target) return false;
	const borderEdge = separator === true ? "single" : separator;
	let pPr = firstChildNS(target, w, "pPr");
	if (!pPr) {
		pPr = ownerDoc.createElementNS(w, "w:pPr");
		target.insertBefore(pPr, target.firstChild);
	}
	let pBdr = firstChildNS(pPr, w, "pBdr");
	if (!pBdr) {
		pBdr = ownerDoc.createElementNS(w, "w:pBdr");
		insertChildInOrder(pPr, pBdr, PPR_CHILD_ORDER);
	}
	const existingEdge = firstChildNS(pBdr, w, edge);
	if (existingEdge) pBdr.removeChild(existingEdge);
	pBdr.appendChild(buildBorderElement(`w:${edge}`, borderEdge, ownerDoc, 1));
	return true;
}
function findFirstChildEl(parent, ns, local) {
	for (const c of getChildren(parent)) if (c.namespaceURI === ns && c.localName === local) return c;
	return null;
}
function findLastChildEl(parent, ns, local) {
	let found = null;
	for (const c of getChildren(parent)) if (c.namespaceURI === ns && c.localName === local) found = c;
	return found;
}
/** Namespace declarations a body / header / footer part may need once
*  post-processing splices drawing or VML content into it. The engine's own
*  emitters only use w / r / wp / a / pic / m, but a source document can
*  carry — and an extension script can add — shapes, canvases, VML
*  watermarks and the w14/w15 markups. A part declaring only w+r makes
*  every one of those schema-invalid, which is exactly the
*  "Namespace 'wps' in Ignorable but not declared" failure. */
const EXTRA_ROOT_NAMESPACES = {
	"mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
	"v": "urn:schemas-microsoft-com:vml",
	"o": "urn:schemas-microsoft-com:office:office",
	"w10": "urn:schemas-microsoft-com:office:word",
	"wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
	"wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
	"wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
	"w14": "http://schemas.microsoft.com/office/word/2010/wordml",
	"w15": "http://schemas.microsoft.com/office/word/2012/wordml",
	"wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
};
/** Prefixes Word marks ignorable (extensions whose absence is harmless).
*  The shape/canvas family is deliberately absent: `wps` / `wpg` / `wpc`
*  carry real rendering, and marking them ignorable would let a consumer
*  drop visible content. */
const IGNORABLE_EXTRA_PREFIXES = ["w14", "w15", "wp14"];
/** Prefixes whose elements render real content; listing one in
*  `mc:Ignorable` tells a consumer it may drop it, which would silently
*  delete visible shapes. */
const RENDERING_PREFIXES = ["wps", "wpg", "wpc"];
/** Splice missing declarations onto the root start tag of a serialized XML
*  part and merge the ignorable-extension list.
*
*  Why textual rather than DOM: the hand-rolled DOM can't round-trip
*  `xmlns:` attributes (a plain setAttribute drops them, and no
*  namespace-aware writer is exposed), which is why the extension scripts
*  rewrite the root tag after serialization. Same approach here. Existing
*  declarations always win, so a source document's own bindings are left
*  exactly as Word wrote them. */
function ensureRootDeclarations(xml, extras = EXTRA_ROOT_NAMESPACES, ignorable = IGNORABLE_EXTRA_PREFIXES) {
	// Lazy attribute group on purpose: `[^>"']` would also swallow the `/`
	// of a self-closing root tag, leaving `attrs` with a trailing slash and
	// the `(\/?)` group permanently empty. Lazy matching stops at the first
	// `>` the optional-slash group can satisfy.
	const m = /<(?!\?|!)([\w.-]+(?::[\w.-]+)?)((?:"[^"]*"|'[^']*'|[^>"'])*?)(\/?)>/.exec(xml);
	if (!m) return xml;
	const [full, name, attrs, selfClose] = m;
	let nextAttrs = attrs;
	for (const [prefix, uri] of Object.entries(extras)) {
		if (new RegExp(`\\sxmlns:${prefix}\\s*=`).test(nextAttrs)) continue;
		nextAttrs += ` xmlns:${prefix}="${uri}"`;
	}
	if (ignorable.length > 0) {
		const ign = /\s*mc:Ignorable\s*=\s*"([^"]*)"/.exec(nextAttrs);
		const decls = new Set((ign?.[1] ?? "").split(/\s+/).filter(Boolean));
		// A shape/canvas prefix must never be marked ignorable: the flag
		// means "safe to drop if unknown", and these carry visible content.
		// Some hand-built documents (and earlier revisions of the extension
		// scripts) list them anyway, so drop them rather than inherit.
		for (const bad of RENDERING_PREFIXES) decls.delete(bad);
		for (const p of ignorable) decls.add(p);
		const merged = [...decls].join(" ");
		if (ign) nextAttrs = nextAttrs.replace(ign[0], "");
		nextAttrs += ` mc:Ignorable="${merged}"`;
	}
	return xml.slice(0, m.index) + `<${name}${nextAttrs}${selfClose}>` + xml.slice(m.index + full.length);
}
/** Skeleton XML for a fresh `<w:hdr>` or `<w:ftr>` part. Carries every
*  namespace a body paragraph might bring along (drawingML, math, mc, the
*  Word 2010 shape/canvas family) so serialization doesn't have to splice
*  declarations onto descendants. `mc:Ignorable` mirrors what Word writes:
*  the extension markups whose absence is harmless, not the shape elements
*  themselves (`wps` / `wpg` / `wpc` carry real rendering and must never be
*  marked ignorable). */
function buildHfSkeleton(surface) {
	return parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<${surface === "header" ? "w:hdr" : "w:ftr"} xmlns:w="${NS.w}" xmlns:r="${NS.r}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:pic="${NS.pic}" xmlns:m="${NS.m}" xmlns:mc="${NS.mc}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" mc:Ignorable="w14 w15 wp14"/>`);
}
/** Scan the source archive for existing `word/header<N>.xml` /
*  `word/footer<N>.xml` entries. Returns the next free index — new parts
*  start at this number and increment per-emission, preventing collision
*  with parts left over from a prior apply. */
function nextFreePartIndex(reader) {
	let max = 0;
	for (const entry of reader.listEntries()) {
		const m = entry.match(/^word\/(?:header|footer)(\d+)\.xml$/);
		if (m) {
			const n = parseInt(m[1], 10);
			if (n > max) max = n;
		}
	}
	return max + 1;
}
/** Materialise one (surface, variant) pair: emit blocks, register
*  rels/content-types, stage the XML + binary additions. Mutates
*  `replacements` and the shared `contentTypes` / `bodyPartRels` in place;
*  returns the metadata the report and sectPr binding need. */
function emitOnePart(args) {
	const { surface, variant, blocks, separator, partIndex, bodyPartRels, contentTypes, replacements, resolveStyle } = args;
	const partFileName = `${surface}${partIndex}.xml`;
	const partPath = `word/${partFileName}`;
	const partNameSlash = `/${partPath}`;
	const relsPath = `word/_rels/${partFileName}.rels`;
	const partRegistry = new DocxAssetRegistry(new PartRels("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>"), contentTypes);
	const doc = buildHfSkeleton(surface);
	const root = doc.documentElement;
	let hasHyperlinks = false;
	const ctx = {
		emitImage: (src, width, height, alt, ownerDoc) => {
			const { rId } = partRegistry.registerImage(src);
			return partRegistry.buildDrawing(rId, width, height, alt, ownerDoc);
		},
		resolveStyle,
		emitHyperlink: (link, text, format, ownerDoc) => {
			if (!link.startsWith("#")) hasHyperlinks = true;
			return emitHyperlinkNode(ownerDoc, link, text, format, partRegistry);
		}
	};
	for (const block of blocks) root.appendChild(emitBlock(block, doc, ctx));
	let separatorAttached;
	if (separator !== void 0) separatorAttached = attachSeparator(root, surface === "header" ? "bottom" : "top", separator, doc);
	replacements.set(partPath, ensureRootDeclarations(serializeXml(doc)));
	partRegistry.flushTo(replacements, relsPath);
	contentTypes.ensureOverride(partNameSlash, surface === "header" ? HEADER_CONTENT_TYPE : FOOTER_CONTENT_TYPE);
	const { rId } = bodyPartRels.appendRel(surface === "header" ? HEADER_REL_TYPE : FOOTER_REL_TYPE, partFileName);
	return {
		surface,
		variant,
		rId,
		partName: partNameSlash,
		blockCount: blocks.length,
		hasHyperlinks,
		separatorAttached
	};
}
/** Layer per-section overrides on top of the top-level header/footer
*  surfaces. Per-section overrides REPLACE the surface wholesale (parallels
*  `pageSetup.sections` semantics for paperSize / columns; the surface object
*  is the atomic unit because variants + separator have to stay coherent). */
function buildEffectiveSections(config, sectionCount) {
	const result = [];
	for (let i = 0; i < sectionCount; i++) result.push({
		header: config.header,
		footer: config.footer
	});
	if (!config.sections) return result;
	for (const [key, override] of Object.entries(config.sections)) {
		const indices = expandSectionSelector(key, sectionCount, "headerFooter.sections");
		for (const idx1 of indices) {
			const slot = result[idx1 - 1];
			if (override.header !== void 0) slot.header = override.header;
			if (override.footer !== void 0) slot.footer = override.footer;
		}
	}
	return result;
}
/** Bucket sections by JSON-equal effective config. Sections with identical
*  HF share one set of parts (no part-per-section explosion when 50 sections
*  share the same default header). JSON.stringify works because the
*  effective config objects are built deterministically (we know the field
*  order: header before footer; nested config came from a strictObject
*  schema, so its key order is stable across same-shaped inputs). */
function groupSections(effective) {
	const groups = /* @__PURE__ */ new Map();
	const order = [];
	for (let i = 0; i < effective.length; i++) {
		const cfg = effective[i];
		const key = JSON.stringify(cfg);
		let g = groups.get(key);
		if (!g) {
			g = {
				sectionIndices: [],
				config: cfg
			};
			groups.set(key, g);
			order.push(key);
		}
		g.sectionIndices.push(i + 1);
	}
	return order.map((k) => groups.get(k));
}
/** Emit every declared HF part, registering each in Content_Types and the
*  body rels. The caller is responsible for: (a) flushing the body's own
*  `PartRels` (since this function appends to it) and (b) flushing the
*  shared `ContentTypes` (single flush at end of apply, like the body's
*  asset registry — multi-registry race is avoided by the orchestrator
*  owning the single flush).
*
*  When `config.sections` partitions the document, each section group emits
*  its own parts; identical sections share parts (group dedup). When no
*  sections override is declared, every section falls into one group bound
*  to the top-level header/footer — preserves the pre-sections behaviour. */
async function applyHeaderFooter(reader, documentDoc, config, bodyPartRels, contentTypes, replacements, stylesDoc) {
	const body = firstChildNS(documentDoc.documentElement, w, "body");
	const sectionCount = (body ? collectSectPrs(body) : []).length;
	const parts = [];
	const groups = [];
	const separatorWarnings = [];
	let hasEven = false;
	let partIndex = nextFreePartIndex(reader);
	const resolveStyle = buildStyleResolver(stylesDoc);
	if (sectionCount === 0) return {
		parts,
		groups,
		hasEven,
		separatorWarnings
	};
	const pending = groupSections(buildEffectiveSections(config, sectionCount));
	for (const pg of pending) {
		let hasFirst = false;
		const groupParts = [];
		for (const surface of SURFACE_ORDER) {
			const surfaceCfg = pg.config[surface];
			if (!surfaceCfg) continue;
			const separator = surface === "header" ? surfaceCfg.underline : surfaceCfg.overline;
			const surfaceRecords = [];
			for (const variant of VARIANT_ORDER) {
				const blocks = surfaceCfg[variant];
				if (blocks === void 0) continue;
				if (variant === "first") hasFirst = true;
				if (variant === "even") hasEven = true;
				const record = emitOnePart({
					surface,
					variant,
					blocks,
					separator,
					partIndex,
					reader,
					bodyPartRels,
					contentTypes,
					replacements,
					resolveStyle
				});
				parts.push(record);
				groupParts.push(record);
				surfaceRecords.push(record);
				partIndex++;
			}
			if (separator !== void 0 && surfaceRecords.length > 0 && surfaceRecords.every((r) => r.separatorAttached === false)) {
				const groupTag = scopeTagForGroup(pg, sectionCount);
				const field = surface === "header" ? "underline" : "overline";
				separatorWarnings.push(`${groupTag}${surface}.${field}: declared but no variant has a <w:p> endpoint (every variant ends in a table or is empty [])`);
			}
		}
		groups.push({
			sectionIndices: pg.sectionIndices,
			parts: groupParts,
			hasFirst
		});
	}
	return {
		parts,
		groups,
		hasEven,
		separatorWarnings
	};
}
/** Human-readable scope tag for a group's warning. Drops the prefix when
*  the group covers every section (no need to say "Sections 1-N: ..."
*  in single-group documents). */
function scopeTagForGroup(pg, sectionCount) {
	if (pg.sectionIndices.length === sectionCount) return "";
	if (pg.sectionIndices.length === 1) return `Section ${pg.sectionIndices[0]}: `;
	return `Sections ${pg.sectionIndices.join(", ")}: `;
}
/** Collect every sectPr in body order — paragraph-embedded sectPrs first
*  (sections 1..N-1) then the body-trailing one (section N). Mirrors
*  `collectSectPrs` in page-setup-mutation; duplicated here rather than
*  exported to keep page-setup's surface narrow (the HF binding pass
*  doesn't need the indexing semantics page-setup exposes). */
function collectSectPrs(body) {
	const out = [];
	for (const child of getChildren(body)) {
		if (child.namespaceURI !== w) continue;
		if (child.localName === "p") {
			const pPr = firstChildNS(child, w, "pPr");
			if (pPr) {
				const sectPr = firstChildNS(pPr, w, "sectPr");
				if (sectPr) out.push(sectPr);
			}
		} else if (child.localName === "sectPr") out.push(child);
	}
	return out;
}
/** Remove any pre-existing `<w:headerReference>` / `<w:footerReference>`
*  children from a sectPr. We replace wholesale rather than merge — v1
*  rule: HF declarations apply to EVERY section (decision 2). Leftover
*  references from a prior apply would mix old + new and confuse Word
*  about which part wins for each variant. */
function stripExistingReferences(sectPr) {
	const stale = [];
	for (const c of getChildren(sectPr)) {
		if (c.namespaceURI !== w) continue;
		if (c.localName === "headerReference" || c.localName === "footerReference") stale.push(c);
	}
	for (const el of stale) sectPr.removeChild(el);
}
/** Idempotent set/clear of `<w:titlePg/>` on a sectPr based on `enabled`.
*  HF config is the source of truth: with the block declared, titlePg
*  is set when any surface has a `first` variant, cleared otherwise.
*  Returns true when the DOM actually changed; false on no-op. */
function setTitlePg(sectPr, doc, enabled) {
	const existing = firstChildNS(sectPr, w, "titlePg");
	if (enabled) {
		if (existing) return false;
		insertChildInOrder(sectPr, doc.createElementNS(w, "w:titlePg"), SECT_PR_CHILD_ORDER);
		return true;
	}
	if (!existing) return false;
	sectPr.removeChild(existing);
	return true;
}
/** Inject the built-in `Header` and `Footer` paragraph styles into
*  stylesDoc when missing. Word treats these styleIds as well-known —
*  the canonical name "header" / "footer", uiPriority="99",
*  basedOn="Normal" match what Word generates when you insert a header
*  via the UI.
*
*  Deliberately minimal: no tab stops (the classic split layout uses a
*  3-column borderless table — see `references/header-footer.md`) and
*  no bottom border (Word's default Header style ships one, but it's a
*  divisive cosmetic choice — agents wanting it declare paraFormat).
*
*  Idempotent. Returns true when at least one style was injected. */
function ensureHeaderFooterStyles(stylesDoc) {
	const root = stylesDoc.documentElement;
	if (!root) return false;
	const existing = /* @__PURE__ */ new Set();
	for (const s of getChildrenNS(root, w, "style")) {
		const id = wAttr(s, "styleId");
		if (id) existing.add(id);
	}
	let injected = false;
	for (const { id, name } of [{
		id: "Header",
		name: "header"
	}, {
		id: "Footer",
		name: "footer"
	}]) {
		if (existing.has(id)) continue;
		const style = stylesDoc.createElementNS(w, "w:style");
		style.setAttributeNS(w, "w:type", "paragraph");
		style.setAttributeNS(w, "w:styleId", id);
		const nm = stylesDoc.createElementNS(w, "w:name");
		nm.setAttributeNS(w, "w:val", name);
		style.appendChild(nm);
		const basedOn = stylesDoc.createElementNS(w, "w:basedOn");
		basedOn.setAttributeNS(w, "w:val", "Normal");
		style.appendChild(basedOn);
		const uiPriority = stylesDoc.createElementNS(w, "w:uiPriority");
		uiPriority.setAttributeNS(w, "w:val", "99");
		style.appendChild(uiPriority);
		style.appendChild(stylesDoc.createElementNS(w, "w:unhideWhenUsed"));
		root.appendChild(style);
		injected = true;
	}
	return injected;
}
/** Plug each section's HF part records into its sectPr.
*
*   - clears pre-existing headerReference / footerReference children
*   - appends one `<w:headerReference>` per header variant in the section's
*     group, one `<w:footerReference>` per footer variant
*   - sets `<w:titlePg/>` on the sectPr when this section's group declared
*     a `first` variant; clears otherwise
*
*  evenAndOddHeaders activation lives in settings.xml — see
*  `setEvenAndOddHeadersFlag` in settings-mutation. */
function applyHeaderFooterBinding(documentDoc, report) {
	const body = firstChildNS(documentDoc.documentElement, w, "body");
	if (!body) return {
		sectionCount: 0,
		titlePgApplied: false
	};
	const sectPrs = collectSectPrs(body);
	if (sectPrs.length === 0) return {
		sectionCount: 0,
		titlePgApplied: false
	};
	const sectionGroup = /* @__PURE__ */ new Map();
	for (const group of report.groups) for (const idx of group.sectionIndices) sectionGroup.set(idx, group);
	let titlePgApplied = false;
	for (let i = 0; i < sectPrs.length; i++) {
		const sectPr = sectPrs[i];
		const group = sectionGroup.get(i + 1);
		stripExistingReferences(sectPr);
		if (group) for (const part of group.parts) {
			const ref = documentDoc.createElementNS(w, part.surface === "header" ? "w:headerReference" : "w:footerReference");
			ref.setAttributeNS(w, "w:type", part.variant);
			ref.setAttributeNS(NS.r, "r:id", part.rId);
			insertChildInOrder(sectPr, ref, SECT_PR_CHILD_ORDER);
		}
		if (setTitlePg(sectPr, documentDoc, group?.hasFirst ?? false)) titlePgApplied = true;
	}
	return {
		sectionCount: sectPrs.length,
		titlePgApplied
	};
}

//#endregion
//#region lib/apply/apply-styles.ts
/**
* Normalize a ValidationError into a stable key for baseline comparison.
* Strips file paths, line/column coordinates, and paragraph indices — all
* of which differ between source and output but don't indicate a new
* structural problem introduced by apply. Two errors that vary only in
* position (same schema violation, different location in same part) map to
* the same key so a pre-existing positional shift doesn't create a false
* "new" error.
*/
function normalizeValidationError(e) {
	const msg = e.message.replace(/\/[^\s"']+\.docx/g, "<docx>").replace(/\bline \d+\b/g, "line N").replace(/\bcol \d+\b/g, "col M").replace(/\bparagraph #\d+\b/g, "paragraph #N").replace(/\binput_\d+\.xml\b/g, "input_N.xml").trim();
	return `${e.part}||${msg}`;
}
/** Heuristic checks on a resolved style entry. The engine doesn't reject
* these — they're informational signals surfaced in dry-run + the change
* report so the agent can fix before commit. Narrow by design: only fire
* on Mode A extraction artifacts, not on Mode B explicit declarations
* (agent may legitimately want the CJK font's Latin glyphs for English /
* digit text, which is what `fontLatin` controls). */
function detectStyleResolutionWarnings(def, final) {
	const out = [];
	if (def.fromParagraph !== void 0 && typeof final.fontLatin === "string" && /[一-鿿]/.test(final.fontLatin) && final.fontCJK === void 0) out.push(`fromParagraph extracted fontLatin="${final.fontLatin}" with no fontCJK — source paragraph likely had the CJK font on its ascii/hAnsi slot only (common Word authoring artifact). CJK characters fall through to docDefaults eastAsia at render; add fontCJK to this style if you want consistent CJK appearance, or ignore if the Latin-only slot was intentional.`);
	return out;
}
/** Build per-styleId target paragraph sets from the post-rules-pass doc
*  state, then attach a vs-direct classification to each StyleResolutionEntry.
*  Walks the live documentDoc to read each paragraph's CURRENT pStyle (after
*  edits + rules), matches back to parsed.paragraphs by index for the
*  direct-format snapshot captured at parse time. edits[]-inserted paragraphs
*  whose direct format came from agent's own paraFormat/runFormat would
*  trivially appear redundant — they're included for now (acceptable noise
*  given the small footprint of inserts vs chrome). */
/** True iff `el` is reachable from `doc.documentElement` (not removed). Walks
* parent links rather than relying on any cached structure; constant-time per
* call for shallow trees and bounded for paragraphs. */
function isAttachedToDoc(el, doc) {
	let cur = el;
	while (cur) {
		if (cur === doc.documentElement) return true;
		cur = cur.parentNode;
	}
	return false;
}
function populateVsDirectReports(documentDoc, paragraphs, declaredStyles, styleResolutions, resolver) {
	const w = NS.w;
	const paraByIndex = /* @__PURE__ */ new Map();
	for (const p of paragraphs) paraByIndex.set(p.index, p);
	const defaultStyleId = resolver.getDefaultParagraphStyleId() ?? "Normal";
	const targetsByStyleId = /* @__PURE__ */ new Map();
	for (const indexed of walkIndexedParagraphs(documentDoc)) {
		const p = paraByIndex.get(indexed.index);
		if (!p) continue;
		const pPr = firstChildNS(indexed.element, w, "pPr");
		const pStyleEl = pPr ? firstChildNS(pPr, w, "pStyle") : null;
		const currentStyleId = pStyleEl && wAttr(pStyleEl, "val") || defaultStyleId;
		let bucket = targetsByStyleId.get(currentStyleId);
		if (!bucket) {
			bucket = [];
			targetsByStyleId.set(currentStyleId, bucket);
		}
		bucket.push(p);
	}
	const resolutionsById = /* @__PURE__ */ new Map();
	for (const r of styleResolutions) resolutionsById.set(r.styleId, r);
	for (const def of declaredStyles) {
		const entry = resolutionsById.get(def.id);
		if (!entry) continue;
		const targets = targetsByStyleId.get(def.id) ?? [];
		let fromFp;
		if (def.fromParagraph !== void 0) fromFp = paraByIndex.get(def.fromParagraph)?.fingerprint;
		entry.vsDirect = analyzeVsDirect(def, targets, fromFp);
	}
}
async function applyStyles(source, output, config) {
	config.styles ??= [];
	if (config.captions && config.numbering) {
		const captionStyleIds = /* @__PURE__ */ new Set();
		for (const entry of Object.values(config.captions)) captionStyleIds.add(entry.styleId);
		const numSchemes = Array.isArray(config.numbering) ? config.numbering : [config.numbering];
		for (const scheme of numSchemes) for (const lvl of scheme.levels) if (captionStyleIds.has(lvl.styleId)) throw new Error(`Config invariant: styleId "${lvl.styleId}" is bound to both numbering[] (numPr) and captions[] (SEQ). Caption-class styleIds must not appear in numbering[]; they get their counter from the captions table instead. Drop the numbering[] level referencing "${lvl.styleId}".`);
	}
	if (!config.dryRun) {
		mkdirSync(dirname(output), { recursive: true });
		copyFileSync(source, output);
	}
	const reader = await DocxReader.open(config.dryRun ? source : output);
	const stylesDoc = await reader.readXml("word/styles.xml") ?? blankStylesDoc();
	const numberingDoc = await reader.readXml("word/numbering.xml") ?? blankNumberingDoc();
	const documentDoc = await reader.readXml("word/document.xml");
	if (!documentDoc) throw new Error("word/document.xml not found");
	if (config.mergeRuns) {
		// Opt-in pre-pass. Word splits a paragraph's text across many runs
		// (revision ids, proofErr markers), so a phrase visible on the page is
		// often not a contiguous string in the XML — which is exactly what
		// character-level locators (`set-run` at a runIndex, text-matched
		// anchors) look for. Merging has to happen HERE, before paragraphs are
		// parsed and before any locator is resolved, so every index in the
		// config refers to the same run layout the edits operate on. Rendering
		// is unchanged and tracked-change structure is preserved.
		mergeRunsInDocument(documentDoc);
	}
	const themeDoc = await reader.readXml("word/theme/theme1.xml");
	let themeMutated = false;
	if (config.theme?.fonts && themeDoc) {
		applyThemeFontOverrides(themeDoc, config.theme.fonts);
		themeMutated = true;
	}
	const resolver = new StyleResolver(stylesDoc, themeDoc);
	if (config.theme?.fonts) resolver.setThemeFontOverrides(config.theme.fonts);
	resolver.expandThemedFontsInStyles(stylesDoc);
	let parsed = new DocumentParser(documentDoc, resolver, numberingDoc).parse();
	let fpResult = new Fingerprinter().assign(parsed.paragraphs, resolver);
	const hashToLetter = /* @__PURE__ */ new Map();
	for (const s of fpResult.summary) hashToLetter.set(s.hash, s.label);
	const preTemplateSourceNumIds = /* @__PURE__ */ new Set();
	for (const numEl of getChildrenNS(numberingDoc.documentElement, NS.w, "num")) {
		const idStr = wAttr(numEl, "numId");
		if (idStr) {
			const id = parseInt(idStr, 10);
			if (Number.isFinite(id)) preTemplateSourceNumIds.add(id);
		}
	}
	let templateImport = null;
	if (config.template) {
		const tplCfg = config.template;
		const tplPath = resolve(tplCfg.source);
		if (!existsSync(tplPath)) throw new Error(`template not found: ${tplPath}`);
		templateImport = await importTemplateStyles(tplPath, tplCfg.styles, stylesDoc, numberingDoc, { importNumbering: tplCfg.importNumbering });
	}
	if (config.requirements) {
		const declared = new Set(config.styles.map((s) => s.id));
		for (const styleId of Object.keys(config.requirements)) if (!declared.has(styleId)) throw new Error(`requirements: style "${styleId}" is not declared in styles[].\n  Declared: [${[...declared].join(", ")}]\n  Note: requirements is keyed by styleId (not a free-form label).\n  If you renamed a style (e.g. to bind to an existing id like "a3"),\n  update the requirements key to match — the side-by-side report\n  pairs entries by exact id.`);
	}
	const styleResolutions = [];
	const resolvedStyles = config.styles.map((def) => {
		const priorState = extractPriorDisplayFields(resolver, def.id);
		const priorUsage = priorState !== null ? resolver.getStyleDefinition(def.id)?.usageCount ?? 0 : void 0;
		const final = resolveStyleDef(def, parsed.paragraphs);
		const warnings = detectStyleResolutionWarnings(def, final);
		styleResolutions.push({
			styleId: def.id,
			userSpec: config.requirements?.[def.id] ?? null,
			resolved: extractDisplayFields(final),
			priorState,
			priorUsage,
			warnings: warnings.length > 0 ? warnings : void 0
		});
		return final;
	});
	const canonicalNameKey = canonicalStyleName;
	const sourceCanonicalToStyleId = /* @__PURE__ */ new Map();
	const sourceCanonicalToOriginalName = /* @__PURE__ */ new Map();
	for (const styleEl of getChildrenNS(stylesDoc.documentElement, NS.w, "style")) {
		const sid = wAttr(styleEl, "styleId");
		const nameEl = firstChildNS(styleEl, NS.w, "name");
		const nm = nameEl ? wAttr(nameEl, "val") : null;
		if (sid && nm) {
			const key = canonicalNameKey(nm);
			sourceCanonicalToStyleId.set(key, sid);
			sourceCanonicalToOriginalName.set(key, nm);
		}
	}
	const styleNameConflicts = [];
	const sourceStyleIds = new Set(getChildrenNS(stylesDoc.documentElement, NS.w, "style").map((s) => wAttr(s, "styleId")).filter((id) => id !== null));
	for (const def of resolvedStyles) {
		if (def.name === void 0 && sourceStyleIds.has(def.id)) continue;
		const key = canonicalNameKey(def.name ?? def.id);
		const collidingId = sourceCanonicalToStyleId.get(key);
		if (collidingId && collidingId !== def.id) {
			const existingName = sourceCanonicalToOriginalName.get(key);
			styleNameConflicts.push({
				def,
				collidingId,
				existingName
			});
		}
	}
	if (styleNameConflicts.length > 0) {
		const lines = [];
		lines.push(`${styleNameConflicts.length} style name collision(s) — Word treats matching names (and locale aliases) as the same built-in identity and would silently drop the new style's rPr at render:`);
		for (const c of styleNameConflicts) {
			const effectiveName = c.def.name ?? c.def.id;
			const aliasNote = c.existingName !== effectiveName ? ` (locale alias of existing "${c.existingName}")` : "";
			const nameDisplay = c.def.name !== void 0 ? `name="${c.def.name}"` : `name omitted (defaults to id "${c.def.id}")`;
			lines.push(`  styles[].id="${c.def.id}" ${nameDisplay}${aliasNote} → already used by source styleId="${c.collidingId}"`);
		}
		lines.push("");
		lines.push("  Resolve each by either:");
		lines.push(`    (a) override the existing style: set styles[].id to the source styleId`);
		lines.push(`        (above) — minimum-change default. Declare only fields the user`);
		lines.push(`        spec explicitly requires; Mode A fromParagraph or piled-on`);
		lines.push(`        locale defaults (CJK 2-char indent, etc.) rewrite more than`);
		lines.push(`        the user asked to change.`);
		lines.push(`    (b) use the canonical English built-in name when the conflict is a`);
		lines.push(`        locale alias (e.g. name="Body Text" for id="BodyText").`);
		lines.push(`    (c) use a fresh styleId + name pair (e.g. id="BodyMain", name="正文主体")`);
		lines.push(`        when the source has no styleId for this role, or compresses many`);
		lines.push(`        roles onto one styleId so override can't separate them. \`name\` is`);
		lines.push(`        what end users see in Word's style panel — pick a human-readable`);
		lines.push(`        label, not the styleId.`);
		if (config.dryRun) {
			lines.unshift("=== Style Name Conflicts (dry-run; would FAIL on real apply) ===");
			console.error(lines.join("\n") + "\n");
		} else throw new Error(lines.join("\n"));
	}
	const injected = [];
	const updated = [];
	const derivedFrom = /* @__PURE__ */ new Map();
	for (let i = 0; i < resolvedStyles.length; i++) {
		const def = resolvedStyles[i];
		if (upsertStyle(stylesDoc, def) === "created") injected.push(def.id);
		else updated.push(def.id);
		const src = config.styles[i];
		if (src.fromParagraph !== void 0) derivedFrom.set(def.id, src.fromParagraph);
	}
	const installedSchemes = [];
	if (config.numbering) {
		const numberingSchemes = Array.isArray(config.numbering) ? config.numbering : [config.numbering];
		const declaredIds = new Set(config.styles.map((s) => s.id));
		const existingStyleIds = /* @__PURE__ */ new Set();
		for (const s of getChildrenNS(stylesDoc.documentElement, NS.w, "style")) {
			const sid = wAttr(s, "styleId");
			if (sid) existingStyleIds.add(sid);
		}
		const validNumberingTargets = new Set([...declaredIds, ...existingStyleIds]);
		const allNumIds = /* @__PURE__ */ new Set();
		for (const numEl of getChildrenNS(numberingDoc.documentElement, NS.w, "num")) {
			const idStr = wAttr(numEl, "numId");
			if (idStr) {
				const id = parseInt(idStr, 10);
				if (Number.isFinite(id)) allNumIds.add(id);
			}
		}
		const templateImportedNumIds = new Set([...allNumIds].filter((id) => !preTemplateSourceNumIds.has(id)));
		const claimedNumIds = new Set(allNumIds);
		for (const [schemeIdx, scheme] of numberingSchemes.entries()) {
			if (scheme.numId === void 0) continue;
			const path = numberingSchemes.length === 1 ? "numbering" : `numbering[${schemeIdx}]`;
			if (preTemplateSourceNumIds.has(scheme.numId)) throw new Error(`${path}.numId = ${scheme.numId} collides with an existing <w:num w:numId="${scheme.numId}"/> in the source. Either pin to a different numId, or remove the explicit numId to let the engine allocate (the engine will pick an id not in use).`);
			if (templateImportedNumIds.has(scheme.numId)) throw new Error(`${path}.numId = ${scheme.numId} was just claimed by template import in this apply run. Either pin to a different numId, or remove the explicit numId to let the engine allocate (the auto-allocator already skips template-imported ids).`);
			if (claimedNumIds.has(scheme.numId)) {
				const prev = numberingSchemes.findIndex((s, i) => i < schemeIdx && s.numId === scheme.numId);
				const prevPath = numberingSchemes.length === 1 ? "numbering" : `numbering[${prev}]`;
				throw new Error(`${path}.numId ${scheme.numId} conflicts with ${prevPath}.numId ${scheme.numId} — two schemes cannot share the same numId. Set a different numId on one of them, or remove numId to let the engine allocate.`);
			}
			claimedNumIds.add(scheme.numId);
		}
		for (const [schemeIdx, scheme] of numberingSchemes.entries()) {
			if (scheme.levels.length === 0) continue;
			const path = numberingSchemes.length === 1 ? "numbering" : `numbering[${schemeIdx}]`;
			for (const [i, lvl] of scheme.levels.entries()) {
				if (!validNumberingTargets.has(lvl.styleId)) throw new Error(`${path}.levels[${i}]: styleId "${lvl.styleId}" doesn't exist.\n  Declared in styles[]: [${[...declaredIds].join(", ")}]\n  Existing in styles.xml: [${[...existingStyleIds].sort().join(", ")}]`);
				const numPlaceholders = (lvl.lvlText.match(/%\d/g) ?? []).length;
				for (const p of lvl.stripPrefixPatterns ?? []) {
					const pn = (p.match(/%\d/g) ?? []).length;
					if (pn > numPlaceholders) console.error(`Warning: ${path}.levels[${i}].stripPrefixPatterns "${p}" has ${pn} placeholders but lvlText "${lvl.lvlText}" has only ${numPlaceholders}. Pattern may match more than intended.`);
				}
				if (lvl.numFmt === "bullet" || lvl.numFmt === "none") continue;
				const ownPlaceholder = `%${lvl.level + 1}`;
				if (!lvl.lvlText.includes(ownPlaceholder)) {
					const referenced = lvl.lvlText.match(/%\d/g) ?? [];
					const refDesc = referenced.length ? `only references ${referenced.map((p) => `${p} (level ${Number(p[1]) - 1})`).join(", ")}` : "references no counter";
					console.error(`Warning: ${path}.levels[${i}].lvlText "${lvl.lvlText}" does not reference its own counter ${ownPlaceholder} — ${refDesc}.\n  All level-${lvl.level} items will display the same number, restarting only when a higher level changes.\n  %N is positional (1-indexed): %1 → level 0, %2 → level 1, ... — so level ${lvl.level} needs ${ownPlaceholder} to render its own counter.\n  Did you mean lvlText: "${lvl.lvlText.replace(/%\d/, ownPlaceholder)}"?`);
				}
			}
			const { numId, abstractNumId } = injectNumbering(numberingDoc, scheme, {
				claimedNumIds,
				requestedNumId: scheme.numId
			});
			for (const lvl of scheme.levels) attachNumberingToStyle(stylesDoc, lvl.styleId, numId, lvl.level);
			installedSchemes.push({
				levels: scheme.levels.map((l) => ({
					level: l.level,
					styleId: l.styleId,
					restart: l.restart ?? "continuous"
				})),
				numId,
				abstractNumId,
				numIdExplicit: scheme.numId !== void 0
			});
		}
	}
	reorderAgentTouchedStylesFirst(stylesDoc, [...templateImport?.imported ?? [], ...config.styles.map((s) => s.id)]);
	const resolvedCaptions = resolveCaptions(config.captions, stylesDoc);
	if (!config.dryRun) for (const wmsg of resolvedCaptions.warnings) console.warn(`Warning: ${wmsg}`);
	const bodyAssetRegistry = await DocxAssetRegistry.open(reader);
	let editsApplied = 0;
	let editsTrackChanges = false;
	let editsPreview = [];
	let editTouchedIndices;
	let crossRefsTouched = false;
	let listRestartApplied = false;
	let bookmarkAllocator = null;
	let pendingBackfills = [];
	let pendingCaptionFills = [];
	let pendingCaptionResets = [];
	const panguWarnings = [];
	const panguLintEnabled = config.lint?.panguSpacing !== false;
	if (panguLintEnabled && config.headerFooter) panguWarnings.push(...lintPanguInHeaderFooter(config.headerFooter));
	if (config.edits && config.edits.length > 0) {
		if (panguLintEnabled) panguWarnings.push(...lintPanguInEdits(config.edits));
		const preview = previewEditOps({
			documentDoc,
			parsedParagraphs: parsed.paragraphs,
			edits: config.edits
		});
		editsPreview = preview.entries;
		editTouchedIndices = preview.replacedOrDeletedIndices;
		const result = await runEditOps({
			documentDoc,
			parsedParagraphs: parsed.paragraphs,
			reader,
			edits: config.edits,
			trackChanges: config.trackChanges ?? false,
			author: config.author,
			stylesDoc,
			sections: parsed.sections,
			captions: resolvedCaptions.byIdentifier,
			imageRegistry: bodyAssetRegistry,
			numberingDoc
		});
		editsApplied = result.report.applied;
		editsTrackChanges = result.report.trackChanges;
		({bookmarkAllocator, pendingBackfills, pendingCaptionFills, pendingCaptionResets} = result.crossRefs);
		parsed = new DocumentParser(documentDoc, resolver, numberingDoc).parse();
		fpResult = new Fingerprinter().assign(parsed.paragraphs, resolver);
		hashToLetter.clear();
		for (const s of fpResult.summary) hashToLetter.set(s.hash, s.label);
	}
	const declaredStyleIds = new Set(config.styles.map((s) => s.id));
	const checkStyleId = (id, where) => {
		if (!id) return;
		if (!declaredStyleIds.has(id)) {
			const all = [...declaredStyleIds].join(", ");
			throw new Error(`${where}: style "${id}" is not declared in styles[].\n  Declared: [${all || "(none)"}]`);
		}
	};
	const validIndices = new Set(parsed.paragraphs.map((p) => p.index));
	const checkParaIndex = (idx, where) => {
		if (validIndices.has(idx)) return;
		const max = parsed.paragraphs.length;
		const range = max > 0 ? `#${parsed.paragraphs[0].index}–#${parsed.paragraphs[max - 1].index}` : "(none)";
		throw new Error(`${where}: paragraph #${idx} not found. Document has ${max} indexed paragraphs (${range}). Paragraphs inside data tables are not indexed.`);
	};
	const excludeSet = new Set(config.exclude ?? []);
	for (const idx of excludeSet) checkParaIndex(idx, `exclude`);
	const assignmentMap = /* @__PURE__ */ new Map();
	for (const [i, a] of (config.assignments ?? []).entries()) {
		if (a.action === "restyle") checkStyleId(a.style, `assignments[${i}]`);
		checkParaIndex(a.para, `assignments[${i}].para`);
		assignmentMap.set(a.para, a);
	}
	const bulkMap = /* @__PURE__ */ new Map();
	const declaredFingerprints = new Set(parsed.paragraphs.map((p) => p.fingerprint));
	for (const [i, b] of (config.bulk_rules ?? []).entries()) {
		checkStyleId(b.style, `bulk_rules[${i}]`);
		let letter;
		if (declaredFingerprints.has(b.fingerprint)) letter = b.fingerprint;
		else if (hashToLetter.has(b.fingerprint)) letter = hashToLetter.get(b.fingerprint);
		if (!letter) {
			const letters = [...declaredFingerprints].sort().join(", ");
			const hashes = [...hashToLetter.keys()].sort().join(", ");
			throw new Error(`bulk_rules[${i}]: fingerprint "${b.fingerprint}" doesn't exist in this document.\n  Available letters: [${letters}]\n  Available hashes:  [${hashes}]`);
		}
		bulkMap.set(letter, b.style);
	}
	const restyleStats = /* @__PURE__ */ new Map();
	const flags = [];
	const manualNumberingRemoved = /* @__PURE__ */ new Map();
	const manualNumberingByStyle = /* @__PURE__ */ new Map();
	const patternMatchStats = /* @__PURE__ */ new Map();
	const patternStripStats = /* @__PURE__ */ new Map();
	const patternRules = [];
	for (const [i, p] of (config.pattern_rules ?? []).entries()) {
		if (!declaredStyleIds.has(p.style)) throw new Error(`pattern_rules[${i}].style "${p.style}" is not declared in styles[]. Defined: [${[...declaredStyleIds].join(", ")}]`);
		let re;
		try {
			re = new RegExp(p.regex, p.flags ?? "");
		} catch (err) {
			throw new Error(`pattern_rules[${i}].regex "${p.regex}" is invalid: ${err.message}`, { cause: err });
		}
		patternRules.push({
			regex: re,
			style: p.style,
			stripMatch: p.stripMatch ?? false,
			source: p.regex
		});
	}
	const numLvlTextByStyle = /* @__PURE__ */ new Map();
	if (config.numbering) {
		const numberingSchemes = Array.isArray(config.numbering) ? config.numbering : [config.numbering];
		for (const scheme of numberingSchemes) for (const lvl of scheme.levels) {
			const patterns = lvl.stripPrefixPatterns && lvl.stripPrefixPatterns.length > 0 ? lvl.stripPrefixPatterns : [lvl.lvlText];
			numLvlTextByStyle.set(lvl.styleId, patterns);
		}
	}
	const stylePPrCascade = buildStyleChildCascade(stylesDoc, "pPr");
	const styleRPrCascade = buildStyleChildCascade(stylesDoc, "rPr");
	const samples = /* @__PURE__ */ new Map();
	const implicitKeepByFingerprint = /* @__PURE__ */ new Map();
	const unstrippedByStyle = /* @__PURE__ */ new Map();
	const ctx = {
		excludeSet,
		assignmentMap,
		bulkMap,
		patternRules,
		patternMatchStats,
		patternStripStats,
		paragraphs: parsed.paragraphs,
		restyleStats,
		flags,
		manualNumberingRemoved,
		manualNumberingByStyle,
		numLvlTextByStyle,
		samples,
		samplesPerStyleCap: 5,
		implicitKeepByFingerprint,
		unstrippedByStyle,
		editTouchedIndices,
		stylePPrCascade,
		styleRPrCascade
	};
	applyToBody(documentDoc, ctx);
	if (installedSchemes.length > 0 && !listRestartApplied) {
		applyListRestartPass(documentDoc, numberingDoc, installedSchemes, buildHeadingStyleIdSet(stylesDoc));
		listRestartApplied = true;
	}
	let captionsPreview = null;
	if (resolvedCaptions.byIdentifier.size > 0 || pendingBackfills.length > 0 || pendingCaptionFills.length > 0 || pendingCaptionResets.length > 0) {
		if (!bookmarkAllocator) bookmarkAllocator = new BookmarkAllocator(documentDoc);
		const chapterSeqsInjected = injectChapterCounters(documentDoc, resolvedCaptions.byIdentifier);
		if (chapterSeqsInjected > 0) ensureHiddenChapterCounterStyle(stylesDoc);
		const freshlyEmitted = new Set(pendingCaptionFills.map((f) => f.paragraph));
		const standardizeResult = standardizeCaptions(documentDoc, resolvedCaptions.byIdentifier, bookmarkAllocator, freshlyEmitted);
		for (const w of standardizeResult.warnings) console.warn(w);
		const allCaptionFills = [...pendingCaptionFills, ...standardizeResult.fills];
		const captionSimOutput = allCaptionFills.length > 0 || pendingCaptionResets.length > 0 ? simulateCaptions(documentDoc, {
			fills: allCaptionFills,
			resets: pendingCaptionResets,
			configs: resolvedCaptions.byIdentifier,
			outlineParagraphs: buildOutlineParagraphsMap(documentDoc, numberingDoc, stylesDoc, resolvedCaptions.styleIdToName)
		}) : {
			fieldValues: /* @__PURE__ */ new Map(),
			fullCaptionText: /* @__PURE__ */ new Map()
		};
		for (const [el, text] of captionSimOutput.fieldValues) el.textContent = text;
		if (captionSimOutput.fieldValues.size > 0 || captionSimOutput.fullCaptionText.size > 0) crossRefsTouched = true;
		if (config.dryRun) {
			const previewSamples = [];
			for (const fill of allCaptionFills) {
				if (previewSamples.length >= 5) break;
				const counterText = captionSimOutput.fullCaptionText.get(fill.paragraph);
				if (counterText !== void 0) {
					const captionConfig = resolvedCaptions.byIdentifier.get(fill.identifier);
					const body = fill.bodyText;
					const text = body !== void 0 && captionConfig !== void 0 ? counterText + captionConfig.bodySeparator + body : counterText;
					previewSamples.push({
						identifier: fill.identifier,
						text
					});
				}
			}
			captionsPreview = {
				chapterSeqsInjected,
				standardizeReemitted: standardizeResult.fills.length,
				freshlyEmitted: pendingCaptionFills.length,
				samples: previewSamples
			};
		}
		if (bookmarkAllocator.hasAllocations() || captionSimOutput.fullCaptionText.size > 0) {
			const resolvedBackfills = pendingBackfills.map((pending) => {
				const rec = bookmarkAllocator.resolveByName(pending.targetName);
				if (!rec) throw new Error(`InlineRef: target bookmark "${pending.targetName}" could not be resolved post-emit. This is an engine invariant violation; pre-scan should have caught a missing anchor.`);
				return {
					...pending,
					targetParagraph: rec.element
				};
			});
			for (const pending of resolvedBackfills) if (!isAttachedToDoc(pending.targetParagraph, documentDoc)) throw new Error("InlineRef target paragraph was removed by a later edit op. Reorder edits so the ref op runs before any op that replaces / deletes the target, or remove the conflicting edit.");
			const counters = simulateNumberingCounters(documentDoc, numberingDoc, stylesDoc);
			for (const pending of resolvedBackfills) {
				const captionText = captionSimOutput.fullCaptionText.get(pending.targetParagraph);
				if (captionText !== void 0) {
					pending.placeholderTextEl.textContent = captionText;
					continue;
				}
				const rendered = counters.get(pending.targetParagraph);
				if (!rendered && pending.display !== "full") continue;
				const text = pending.display === "number" ? rendered?.number ?? "" : pending.display === "full" ? extractParagraphText(pending.targetParagraph) : rendered?.label ?? "";
				pending.placeholderTextEl.textContent = text;
			}
			bookmarkAllocator.commit(documentDoc);
			crossRefsTouched = true;
		}
	}
	if (config.dryRun) populateVsDirectReports(documentDoc, parsed.paragraphs, config.styles, styleResolutions, resolver);
	let pageSetupReport;
	if (config.pageSetup) pageSetupReport = applyPageSetup(documentDoc, config.pageSetup);
	const replacements = new WritableArchive();
	let headerFooterReport;
	let headerFooterBindingReport;
	if (config.headerFooter) {
		ensureHeaderFooterStyles(stylesDoc);
		headerFooterReport = await applyHeaderFooter(reader, documentDoc, config.headerFooter, bodyAssetRegistry.getPartRels(), bodyAssetRegistry.getContentTypes(), replacements, stylesDoc);
		headerFooterBindingReport = applyHeaderFooterBinding(documentDoc, headerFooterReport);
		if (headerFooterReport.parts.some((p) => p.hasHyperlinks)) ensureHyperlinkCharStyle(stylesDoc);
		await setEvenAndOddHeadersFlag(reader, replacements, headerFooterReport.hasEven, bodyAssetRegistry);
	}
	replacements.set("word/styles.xml", serializeXml(stylesDoc));
	if (numberingDoc) {
		// Any part we write MUST also be registered — content type + relationship.
		// Creating from blank (no `config.numbering` declared) still yields a
		// numberingDoc, so gating registration on the *config* left the freshly
		// written word/numbering.xml unreferenced: PRO's validator calls that
		// `Unreferenced file`, and Word treats the package as corrupt.
		replacements.set("word/numbering.xml", serializeXml(numberingDoc));
		const sharedCT = bodyAssetRegistry.getContentTypes();
		const sharedRels = bodyAssetRegistry.getPartRels();
		sharedCT.ensureOverride("/word/numbering.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml");
		if (!sharedRels.hasRelTo("numbering.xml")) sharedRels.appendRel("http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering", "numbering.xml");
	}
	replacements.set("word/document.xml", ensureRootDeclarations(serializeXml(documentDoc)));
	if (themeMutated && themeDoc) replacements.set("word/theme/theme1.xml", serializeXml(themeDoc));
	if (crossRefsTouched) await ensureUpdateFieldsFlag(reader, replacements, bodyAssetRegistry);
	let settingsReport;
	if (config.settings?.elements?.length) settingsReport = await applySettingsElements(reader, replacements, bodyAssetRegistry, config.settings.elements);
	let docPropsReport;
	if (config.docProps?.enabled !== false) docPropsReport = await ensureDocProps(reader, replacements, bodyAssetRegistry, documentDoc, {
		title: config.docProps?.title,
		creator: config.docProps?.creator
	});
	bodyAssetRegistry.flushTo(replacements);
	bodyAssetRegistry.getContentTypes().flushTo(replacements);
	if (!config.dryRun) {
		const baselineErrors = await validateDocxFile(source);
		const baselineCounts = /* @__PURE__ */ new Map();
		for (const e of baselineErrors) {
			const k = normalizeValidationError(e);
			baselineCounts.set(k, (baselineCounts.get(k) ?? 0) + 1);
		}
		await reader.copyAndModify(output, replacements);
		const outputErrors = await validateDocxFile(output);
		const available = new Map(baselineCounts);
		const newErrors = [];
		const preExistingErrors = [];
		for (const e of outputErrors) {
			const k = normalizeValidationError(e);
			const n = available.get(k) ?? 0;
			if (n > 0) {
				available.set(k, n - 1);
				preExistingErrors.push(e);
			} else newErrors.push(e);
		}
		if (preExistingErrors.length > 0) console.error(`Validation: ${preExistingErrors.length} pre-existing error(s) carried through from source (non-fatal).`);
		if (newErrors.length > 0) {
			const lines = newErrors.slice(0, 20).map((e) => `  ${e.part}: ${e.message}`);
			const more = newErrors.length > 20 ? `\n  …${newErrors.length - 20} more` : "";
			if (config.allowValidationWarnings) console.error(`Validation: ${newErrors.length} new error(s) introduced by this run (--allow-validation-warnings set; output kept):\n${lines.join("\n")}${more}`);
			else {
				if (existsSync(output)) try {
					unlinkSync(output);
				} catch {}
				console.error(`Validation FAILED — ${newErrors.length} new error(s) introduced by this run:\n${lines.join("\n")}${more}\n  (Use --allow-validation-warnings to keep the output despite new errors.)`);
				process.exit(1);
			}
		}
	}
	const numberingBindings = (() => {
		if (!config.numbering) return [];
		return (Array.isArray(config.numbering) ? config.numbering : [config.numbering]).flatMap((scheme) => scheme.levels.map((lvl) => {
			const { suff, effectiveLvlText } = resolveSuff(lvl.lvlText, lvl.suff);
			return {
				styleId: lvl.styleId,
				level: lvl.level,
				lvlText: effectiveLvlText,
				suff
			};
		}));
	})();
	const manualNumberingDetected = config.dryRun ? detectManualNumbering(config.edits, new Set(numLvlTextByStyle.keys())) : void 0;
	const excludeSamples = config.dryRun && excludeSet.size > 0 ? [...excludeSet].sort((a, b) => a - b).map((idx) => {
		const text = parsed.paragraphs.find((q) => q.index === idx)?.text.trim() ?? "";
		return {
			index: idx,
			snippet: text ? text.slice(0, 60) + (text.length > 60 ? "…" : "") : "(empty)"
		};
	}) : void 0;
	printReport({
		source,
		injected,
		updated,
		restyleStats,
		flags,
		manualNumberingRemoved,
		manualNumberingByStyle,
		patternMatchStats,
		patternStripStats,
		styleResolutions,
		derivedFrom,
		output,
		dryRun: !!config.dryRun,
		samples: ctx.samples,
		implicitKeepByFingerprint: ctx.implicitKeepByFingerprint,
		unstrippedByStyle: ctx.unstrippedByStyle,
		manualNumberingDetected,
		excludeSamples,
		numberingBindings,
		numberingAllocation: installedSchemes.map((s, i) => ({
			schemeIndex: i,
			numId: s.numId,
			explicit: s.numIdExplicit
		})),
		templateImport,
		editsPreview,
		captionsPreview,
		panguWarnings: panguWarnings.length > 0 ? panguWarnings : void 0,
		pageSetup: pageSetupReport,
		headerFooter: headerFooterReport,
		headerFooterBinding: headerFooterBindingReport,
		settingsElements: settingsReport && settingsReport.length > 0 ? settingsReport : void 0,
		docProps: docPropsReport ?? void 0,
		totalParagraphs: config.dryRun ? parsed.paragraphs[parsed.paragraphs.length - 1]?.index ?? parsed.paragraphs.length : void 0
	});
	if (editsApplied > 0 && !config.dryRun) console.error(`\nEdit pass: ${editsApplied} op(s) applied${editsTrackChanges ? " (track-changes)" : ""}. New paragraphs participated in pattern_rules / bulk_rules cleanup uniformly with the original chrome.`);
}
/**
* Build a function that maps a style display name to a locale-independent
* canonical key. Word's built-in styles have one English name and one
* localized name per UI language; both forms render to the same internal
* identity, so two styleIds claiming "Normal" and "正文" respectively
* collide just as if they both claimed "Normal".
*
* The pair list below covers the en-US ↔ zh-CN aliases for the built-in
* styles agents typically inject (Normal, headings, Title, Body Text, list
* styles, captions, headers/footers, TOC entries). Unknown names fall
* through to themselves — string equality still catches direct collisions
* outside the listed pairs. Add other locale pairs (zh-TW, ja, ko) when a
* real-doc collision surfaces; this is OOXML-spec data, not natural-
* language enumeration.
*/
/** Build styleId → set of `<childName>` (`pPr` or `rPr`) child localNames the
* style's cascade declares (style + every basedOn ancestor). The rules pass
* uses this to strip a paragraph's direct properties only when the new style
* actually provides values for them; otherwise stripping would fall back to
* docDefaults / Normal and lose chrome-baked values (line spacing, font
* size, bold, etc.). */
function buildStyleChildCascade(stylesDoc, childName) {
	const w = NS.w;
	const styles = getChildrenNS(stylesDoc.documentElement, w, "style");
	const directChildren = /* @__PURE__ */ new Map();
	const basedOn = /* @__PURE__ */ new Map();
	for (const s of styles) {
		const id = wAttr(s, "styleId");
		if (!id) continue;
		const direct = /* @__PURE__ */ new Set();
		const container = firstChildNS(s, w, childName);
		if (container) {
			for (const c of getChildren(container)) if (c.namespaceURI === w) direct.add(c.localName);
		}
		directChildren.set(id, direct);
		const bo = firstChildNS(s, w, "basedOn");
		basedOn.set(id, bo ? wAttr(bo, "val") : null);
	}
	const cascade = /* @__PURE__ */ new Map();
	const walk = (id, seen) => {
		const cached = cascade.get(id);
		if (cached) return cached;
		if (seen.has(id)) return /* @__PURE__ */ new Set();
		seen.add(id);
		const merged = new Set(directChildren.get(id) ?? []);
		const parent = basedOn.get(id);
		if (parent && directChildren.has(parent)) for (const n of walk(parent, seen)) merged.add(n);
		cascade.set(id, merged);
		return merged;
	};
	for (const id of directChildren.keys()) walk(id, /* @__PURE__ */ new Set());
	return cascade;
}
/** Build `outlineParagraphs: Map<Element, { styleName, rendered }>` for the
* caption simulator. Walks the body, finds paragraphs whose pStyle matches
* a styleId referenced by some caption's chapterPrefix, looks up its
* rendered counter number, returns the pair. Caption simulator uses this
* to resolve STYLEREF chapter prefixes. */
function buildOutlineParagraphsMap(documentDoc, numberingDoc, stylesDoc, styleIdToName) {
	const out = /* @__PURE__ */ new Map();
	if (styleIdToName.size === 0) return out;
	const counters = simulateNumberingCounters(documentDoc, numberingDoc, stylesDoc);
	const w = NS.w;
	for (const [paraEl, rendered] of counters) {
		const pPr = firstChildNS(paraEl, w, "pPr");
		if (!pPr) continue;
		const pStyle = firstChildNS(pPr, w, "pStyle");
		if (!pStyle) continue;
		const styleId = wAttr(pStyle, "val");
		if (!styleId) continue;
		const styleName = styleIdToName.get(styleId);
		if (!styleName) continue;
		out.set(paraEl, {
			styleName,
			rendered: rendered.number
		});
	}
	return out;
}

//#endregion
//#region lib/apply/blank-source.ts
/**
* Locator for the bundled `blank.docx` template that apply uses when a
* config omits `source`. The asset ships two ways:
*
*   1. Dev mode (running TS directly from the repo) — sits at
*      `lib/apply/_assets/blank.docx` relative to the repo root.
*   2. Bundled mode (running from `dist/plugin/skills/docx-master/scripts/...`) —
*      `build-skill.ts` copies the asset into `<scripts>/_assets/blank.docx`
*      so a single basename-relative lookup resolves to the same file.
*
* Resolution mirrors `findSchemasDir` in `docx-validate.ts` — two
* candidates, first hit wins. Throws on miss with a build-skill hint.
*/
const MODULE_DIR = dirname(fileURLToPath(import.meta.url));
/** Absolute path to the bundled blank.docx template. The dev candidate
*  resolves against this file's own location (`lib/apply/_assets/`);
*  the bundled candidate accounts for tsdown's chunking — depending on
*  whether this module gets inlined into the apply.js entry or split
*  into `_shared/`, `import.meta.url` resolves to a different output
*  file, so we look both next to the resolved location and one hop up.
*  Throws on miss. */
function getBlankTemplatePath() {
	const candidates = [join(MODULE_DIR, "_assets", "blank.docx"), join(MODULE_DIR, "..", "_assets", "blank.docx")];
	for (const p of candidates) if (existsSync(p)) return p;
	throw new Error(`bundled blank.docx not found. Looked in: ${candidates.join(", ")}.`);
}

//#endregion
//#region node_modules/zod/v4/core/core.js
var _a;
function $constructor(name, initializer, params) {
	function init(inst, def) {
		if (!inst._zod) Object.defineProperty(inst, "_zod", {
			value: {
				def,
				constr: _,
				traits: /* @__PURE__ */ new Set()
			},
			enumerable: false
		});
		if (inst._zod.traits.has(name)) return;
		inst._zod.traits.add(name);
		initializer(inst, def);
		const proto = _.prototype;
		const keys = Object.keys(proto);
		for (let i = 0; i < keys.length; i++) {
			const k = keys[i];
			if (!(k in inst)) inst[k] = proto[k].bind(inst);
		}
	}
	const Parent = params?.Parent ?? Object;
	class Definition extends Parent {}
	Object.defineProperty(Definition, "name", { value: name });
	function _(def) {
		var _a;
		const inst = params?.Parent ? new Definition() : this;
		init(inst, def);
		(_a = inst._zod).deferred ?? (_a.deferred = []);
		for (const fn of inst._zod.deferred) fn();
		return inst;
	}
	Object.defineProperty(_, "init", { value: init });
	Object.defineProperty(_, Symbol.hasInstance, { value: (inst) => {
		if (params?.Parent && inst instanceof params.Parent) return true;
		return inst?._zod?.traits?.has(name);
	} });
	Object.defineProperty(_, "name", { value: name });
	return _;
}
var $ZodAsyncError = class extends Error {
	constructor() {
		super(`Encountered Promise during synchronous parse. Use .parseAsync() instead.`);
	}
};
(_a = globalThis).__zod_globalConfig ?? (_a.__zod_globalConfig = {});
const globalConfig = globalThis.__zod_globalConfig;
function config(newConfig) {
	if (newConfig) Object.assign(globalConfig, newConfig);
	return globalConfig;
}

//#endregion
//#region node_modules/zod/v4/core/util.js
function getEnumValues(entries) {
	const numericValues = Object.values(entries).filter((v) => typeof v === "number");
	return Object.entries(entries).filter(([k, _]) => numericValues.indexOf(+k) === -1).map(([_, v]) => v);
}
function jsonStringifyReplacer(_, value) {
	if (typeof value === "bigint") return value.toString();
	return value;
}
function cached(getter) {
	return { get value() {
		{
			const value = getter();
			Object.defineProperty(this, "value", { value });
			return value;
		}
		throw new Error("cached value already set");
	} };
}
function nullish(input) {
	return input === null || input === void 0;
}
function cleanRegex(source) {
	const start = source.startsWith("^") ? 1 : 0;
	const end = source.endsWith("$") ? source.length - 1 : source.length;
	return source.slice(start, end);
}
const EVALUATING = /* @__PURE__ */ Symbol("evaluating");
function defineLazy(object, key, getter) {
	let value = void 0;
	Object.defineProperty(object, key, {
		get() {
			if (value === EVALUATING) return;
			if (value === void 0) {
				value = EVALUATING;
				value = getter();
			}
			return value;
		},
		set(v) {
			Object.defineProperty(object, key, { value: v });
		},
		configurable: true
	});
}
const captureStackTrace = "captureStackTrace" in Error ? Error.captureStackTrace : (..._args) => {};
function isObject(data) {
	return typeof data === "object" && data !== null && !Array.isArray(data);
}
function isPlainObject(o) {
	if (isObject(o) === false) return false;
	const ctor = o.constructor;
	if (ctor === void 0) return true;
	if (typeof ctor !== "function") return true;
	const prot = ctor.prototype;
	if (isObject(prot) === false) return false;
	if (Object.prototype.hasOwnProperty.call(prot, "isPrototypeOf") === false) return false;
	return true;
}
const propertyKeyTypes = /* @__PURE__ */ new Set([
	"string",
	"number",
	"symbol"
]);
function escapeRegex(str) {
	return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
function clone(inst, def, params) {
	const cl = new inst._zod.constr(def ?? inst._zod.def);
	if (!def || params?.parent) cl._zod.parent = inst;
	return cl;
}
function normalizeParams(_params) {
	const params = _params;
	if (!params) return {};
	if (typeof params === "string") return { error: () => params };
	if (params?.message !== void 0) {
		if (params?.error !== void 0) throw new Error("Cannot specify both `message` and `error` params");
		params.error = params.message;
	}
	delete params.message;
	if (typeof params.error === "string") return {
		...params,
		error: () => params.error
	};
	return params;
}
function optionalKeys(shape) {
	return Object.keys(shape).filter((k) => {
		return shape[k]._zod.optin === "optional" && shape[k]._zod.optout === "optional";
	});
}
const NUMBER_FORMAT_RANGES = {
	safeint: [Number.MIN_SAFE_INTEGER, Number.MAX_SAFE_INTEGER],
	int32: [-2147483648, 2147483647],
	uint32: [0, 4294967295],
	float32: [-34028234663852886e22, 34028234663852886e22],
	float64: [-Number.MAX_VALUE, Number.MAX_VALUE]
};
function aborted(x, startIndex = 0) {
	if (x.aborted === true) return true;
	for (let i = startIndex; i < x.issues.length; i++) if (x.issues[i]?.continue !== true) return true;
	return false;
}
function explicitlyAborted(x, startIndex = 0) {
	if (x.aborted === true) return true;
	for (let i = startIndex; i < x.issues.length; i++) if (x.issues[i]?.continue === false) return true;
	return false;
}
function prefixIssues(path, issues) {
	return issues.map((iss) => {
		var _a;
		(_a = iss).path ?? (_a.path = []);
		iss.path.unshift(path);
		return iss;
	});
}
function unwrapMessage(message) {
	return typeof message === "string" ? message : message?.message;
}
function finalizeIssue(iss, ctx, config) {
	const message = iss.message ? iss.message : unwrapMessage(iss.inst?._zod.def?.error?.(iss)) ?? unwrapMessage(ctx?.error?.(iss)) ?? unwrapMessage(config.customError?.(iss)) ?? unwrapMessage(config.localeError?.(iss)) ?? "Invalid input";
	const { inst: _inst, continue: _continue, input: _input, ...rest } = iss;
	rest.path ?? (rest.path = []);
	rest.message = message;
	if (ctx?.reportInput) rest.input = _input;
	return rest;
}
function getLengthableOrigin(input) {
	if (Array.isArray(input)) return "array";
	if (typeof input === "string") return "string";
	return "unknown";
}
function issue(...args) {
	const [iss, input, inst] = args;
	if (typeof iss === "string") return {
		message: iss,
		code: "custom",
		input,
		inst
	};
	return { ...iss };
}

//#endregion
//#region node_modules/zod/v4/core/errors.js
const initializer = (inst, def) => {
	inst.name = "$ZodError";
	Object.defineProperty(inst, "_zod", {
		value: inst._zod,
		enumerable: false
	});
	Object.defineProperty(inst, "issues", {
		value: def,
		enumerable: false
	});
	inst.message = JSON.stringify(def, jsonStringifyReplacer, 2);
	Object.defineProperty(inst, "toString", {
		value: () => inst.message,
		enumerable: false
	});
};
const $ZodError = $constructor("$ZodError", initializer);
const $ZodRealError = $constructor("$ZodError", initializer, { Parent: Error });

//#endregion
//#region node_modules/zod/v4/core/parse.js
const _parse = (_Err) => (schema, value, _ctx, _params) => {
	const ctx = _ctx ? {
		..._ctx,
		async: false
	} : { async: false };
	const result = schema._zod.run({
		value,
		issues: []
	}, ctx);
	if (result instanceof Promise) throw new $ZodAsyncError();
	if (result.issues.length) {
		const e = new (_params?.Err ?? _Err)(result.issues.map((iss) => finalizeIssue(iss, ctx, config())));
		captureStackTrace(e, _params?.callee);
		throw e;
	}
	return result.value;
};
const parse = /* @__PURE__ */ _parse($ZodRealError);
const _parseAsync = (_Err) => async (schema, value, _ctx, params) => {
	const ctx = _ctx ? {
		..._ctx,
		async: true
	} : { async: true };
	let result = schema._zod.run({
		value,
		issues: []
	}, ctx);
	if (result instanceof Promise) result = await result;
	if (result.issues.length) {
		const e = new (params?.Err ?? _Err)(result.issues.map((iss) => finalizeIssue(iss, ctx, config())));
		captureStackTrace(e, params?.callee);
		throw e;
	}
	return result.value;
};
const parseAsync = /* @__PURE__ */ _parseAsync($ZodRealError);
const _safeParse = (_Err) => (schema, value, _ctx) => {
	const ctx = _ctx ? {
		..._ctx,
		async: false
	} : { async: false };
	const result = schema._zod.run({
		value,
		issues: []
	}, ctx);
	if (result instanceof Promise) throw new $ZodAsyncError();
	return result.issues.length ? {
		success: false,
		error: new (_Err ?? $ZodError)(result.issues.map((iss) => finalizeIssue(iss, ctx, config())))
	} : {
		success: true,
		data: result.value
	};
};
const safeParse = /* @__PURE__ */ _safeParse($ZodRealError);
const _safeParseAsync = (_Err) => async (schema, value, _ctx) => {
	const ctx = _ctx ? {
		..._ctx,
		async: true
	} : { async: true };
	let result = schema._zod.run({
		value,
		issues: []
	}, ctx);
	if (result instanceof Promise) result = await result;
	return result.issues.length ? {
		success: false,
		error: new _Err(result.issues.map((iss) => finalizeIssue(iss, ctx, config())))
	} : {
		success: true,
		data: result.value
	};
};
const safeParseAsync = /* @__PURE__ */ _safeParseAsync($ZodRealError);

//#endregion
//#region node_modules/zod/v4/core/regexes.js
const dateSource = `(?:(?:\\d\\d[2468][048]|\\d\\d[13579][26]|\\d\\d0[48]|[02468][048]00|[13579][26]00)-02-29|\\d{4}-(?:(?:0[13578]|1[02])-(?:0[1-9]|[12]\\d|3[01])|(?:0[469]|11)-(?:0[1-9]|[12]\\d|30)|(?:02)-(?:0[1-9]|1\\d|2[0-8])))`;
const date = /* @__PURE__ */ new RegExp(`^${dateSource}$`);
const string$1 = (params) => {
	const regex = params ? `[\\s\\S]{${params?.minimum ?? 0},${params?.maximum ?? ""}}` : `[\\s\\S]*`;
	return new RegExp(`^${regex}$`);
};
const number$1 = /^-?\d+(?:\.\d+)?$/;
const boolean$1 = /^(?:true|false)$/i;
const _null$2 = /^null$/i;

//#endregion
//#region node_modules/zod/v4/core/checks.js
const $ZodCheck = /* @__PURE__ */ $constructor("$ZodCheck", (inst, def) => {
	var _a;
	inst._zod ?? (inst._zod = {});
	inst._zod.def = def;
	(_a = inst._zod).onattach ?? (_a.onattach = []);
});
const numericOriginMap = {
	number: "number",
	bigint: "bigint",
	object: "date"
};
const $ZodCheckLessThan = /* @__PURE__ */ $constructor("$ZodCheckLessThan", (inst, def) => {
	$ZodCheck.init(inst, def);
	const origin = numericOriginMap[typeof def.value];
	inst._zod.onattach.push((inst) => {
		const bag = inst._zod.bag;
		const curr = (def.inclusive ? bag.maximum : bag.exclusiveMaximum) ?? Number.POSITIVE_INFINITY;
		if (def.value < curr) if (def.inclusive) bag.maximum = def.value;
		else bag.exclusiveMaximum = def.value;
	});
	inst._zod.check = (payload) => {
		if (def.inclusive ? payload.value <= def.value : payload.value < def.value) return;
		payload.issues.push({
			origin,
			code: "too_big",
			maximum: typeof def.value === "object" ? def.value.getTime() : def.value,
			input: payload.value,
			inclusive: def.inclusive,
			inst,
			continue: !def.abort
		});
	};
});
const $ZodCheckGreaterThan = /* @__PURE__ */ $constructor("$ZodCheckGreaterThan", (inst, def) => {
	$ZodCheck.init(inst, def);
	const origin = numericOriginMap[typeof def.value];
	inst._zod.onattach.push((inst) => {
		const bag = inst._zod.bag;
		const curr = (def.inclusive ? bag.minimum : bag.exclusiveMinimum) ?? Number.NEGATIVE_INFINITY;
		if (def.value > curr) if (def.inclusive) bag.minimum = def.value;
		else bag.exclusiveMinimum = def.value;
	});
	inst._zod.check = (payload) => {
		if (def.inclusive ? payload.value >= def.value : payload.value > def.value) return;
		payload.issues.push({
			origin,
			code: "too_small",
			minimum: typeof def.value === "object" ? def.value.getTime() : def.value,
			input: payload.value,
			inclusive: def.inclusive,
			inst,
			continue: !def.abort
		});
	};
});
const $ZodCheckMinLength = /* @__PURE__ */ $constructor("$ZodCheckMinLength", (inst, def) => {
	var _a;
	$ZodCheck.init(inst, def);
	(_a = inst._zod.def).when ?? (_a.when = (payload) => {
		const val = payload.value;
		return !nullish(val) && val.length !== void 0;
	});
	inst._zod.onattach.push((inst) => {
		const curr = inst._zod.bag.minimum ?? Number.NEGATIVE_INFINITY;
		if (def.minimum > curr) inst._zod.bag.minimum = def.minimum;
	});
	inst._zod.check = (payload) => {
		const input = payload.value;
		if (input.length >= def.minimum) return;
		const origin = getLengthableOrigin(input);
		payload.issues.push({
			origin,
			code: "too_small",
			minimum: def.minimum,
			inclusive: true,
			input,
			inst,
			continue: !def.abort
		});
	};
});
const $ZodCheckStringFormat = /* @__PURE__ */ $constructor("$ZodCheckStringFormat", (inst, def) => {
	var _a, _b;
	$ZodCheck.init(inst, def);
	inst._zod.onattach.push((inst) => {
		const bag = inst._zod.bag;
		bag.format = def.format;
		if (def.pattern) {
			bag.patterns ?? (bag.patterns = /* @__PURE__ */ new Set());
			bag.patterns.add(def.pattern);
		}
	});
	if (def.pattern) (_a = inst._zod).check ?? (_a.check = (payload) => {
		def.pattern.lastIndex = 0;
		if (def.pattern.test(payload.value)) return;
		payload.issues.push({
			origin: "string",
			code: "invalid_format",
			format: def.format,
			input: payload.value,
			...def.pattern ? { pattern: def.pattern.toString() } : {},
			inst,
			continue: !def.abort
		});
	});
	else (_b = inst._zod).check ?? (_b.check = () => {});
});
const $ZodCheckRegex = /* @__PURE__ */ $constructor("$ZodCheckRegex", (inst, def) => {
	$ZodCheckStringFormat.init(inst, def);
	inst._zod.check = (payload) => {
		def.pattern.lastIndex = 0;
		if (def.pattern.test(payload.value)) return;
		payload.issues.push({
			origin: "string",
			code: "invalid_format",
			format: "regex",
			input: payload.value,
			pattern: def.pattern.toString(),
			inst,
			continue: !def.abort
		});
	};
});

//#endregion
//#region node_modules/zod/v4/core/versions.js
const version = {
	major: 4,
	minor: 4,
	patch: 3
};

//#endregion
//#region node_modules/zod/v4/core/schemas.js
const $ZodType = /* @__PURE__ */ $constructor("$ZodType", (inst, def) => {
	var _a;
	inst ?? (inst = {});
	inst._zod.def = def;
	inst._zod.bag = inst._zod.bag || {};
	inst._zod.version = version;
	const checks = [...inst._zod.def.checks ?? []];
	if (inst._zod.traits.has("$ZodCheck")) checks.unshift(inst);
	for (const ch of checks) for (const fn of ch._zod.onattach) fn(inst);
	if (checks.length === 0) {
		(_a = inst._zod).deferred ?? (_a.deferred = []);
		inst._zod.deferred?.push(() => {
			inst._zod.run = inst._zod.parse;
		});
	} else {
		const runChecks = (payload, checks, ctx) => {
			let isAborted = aborted(payload);
			let asyncResult;
			for (const ch of checks) {
				if (ch._zod.def.when) {
					if (explicitlyAborted(payload)) continue;
					if (!ch._zod.def.when(payload)) continue;
				} else if (isAborted) continue;
				const currLen = payload.issues.length;
				const _ = ch._zod.check(payload);
				if (_ instanceof Promise && ctx?.async === false) throw new $ZodAsyncError();
				if (asyncResult || _ instanceof Promise) asyncResult = (asyncResult ?? Promise.resolve()).then(async () => {
					await _;
					if (payload.issues.length === currLen) return;
					if (!isAborted) isAborted = aborted(payload, currLen);
				});
				else {
					if (payload.issues.length === currLen) continue;
					if (!isAborted) isAborted = aborted(payload, currLen);
				}
			}
			if (asyncResult) return asyncResult.then(() => {
				return payload;
			});
			return payload;
		};
		const handleCanaryResult = (canary, payload, ctx) => {
			if (aborted(canary)) {
				canary.aborted = true;
				return canary;
			}
			const checkResult = runChecks(payload, checks, ctx);
			if (checkResult instanceof Promise) {
				if (ctx.async === false) throw new $ZodAsyncError();
				return checkResult.then((checkResult) => inst._zod.parse(checkResult, ctx));
			}
			return inst._zod.parse(checkResult, ctx);
		};
		inst._zod.run = (payload, ctx) => {
			if (ctx.skipChecks) return inst._zod.parse(payload, ctx);
			if (ctx.direction === "backward") {
				const canary = inst._zod.parse({
					value: payload.value,
					issues: []
				}, {
					...ctx,
					skipChecks: true
				});
				if (canary instanceof Promise) return canary.then((canary) => {
					return handleCanaryResult(canary, payload, ctx);
				});
				return handleCanaryResult(canary, payload, ctx);
			}
			const result = inst._zod.parse(payload, ctx);
			if (result instanceof Promise) {
				if (ctx.async === false) throw new $ZodAsyncError();
				return result.then((result) => runChecks(result, checks, ctx));
			}
			return runChecks(result, checks, ctx);
		};
	}
	defineLazy(inst, "~standard", () => ({
		validate: (value) => {
			try {
				const r = safeParse(inst, value);
				return r.success ? { value: r.data } : { issues: r.error?.issues };
			} catch (_) {
				return safeParseAsync(inst, value).then((r) => r.success ? { value: r.data } : { issues: r.error?.issues });
			}
		},
		vendor: "zod",
		version: 1
	}));
});
const $ZodString = /* @__PURE__ */ $constructor("$ZodString", (inst, def) => {
	$ZodType.init(inst, def);
	inst._zod.pattern = [...inst?._zod.bag?.patterns ?? []].pop() ?? string$1(inst._zod.bag);
	inst._zod.parse = (payload, _) => {
		if (def.coerce) try {
			payload.value = String(payload.value);
		} catch (_) {}
		if (typeof payload.value === "string") return payload;
		payload.issues.push({
			expected: "string",
			code: "invalid_type",
			input: payload.value,
			inst
		});
		return payload;
	};
});
const $ZodNumber = /* @__PURE__ */ $constructor("$ZodNumber", (inst, def) => {
	$ZodType.init(inst, def);
	inst._zod.pattern = inst._zod.bag.pattern ?? number$1;
	inst._zod.parse = (payload, _ctx) => {
		if (def.coerce) try {
			payload.value = Number(payload.value);
		} catch (_) {}
		const input = payload.value;
		if (typeof input === "number" && !Number.isNaN(input) && Number.isFinite(input)) return payload;
		const received = typeof input === "number" ? Number.isNaN(input) ? "NaN" : !Number.isFinite(input) ? "Infinity" : void 0 : void 0;
		payload.issues.push({
			expected: "number",
			code: "invalid_type",
			input,
			inst,
			...received ? { received } : {}
		});
		return payload;
	};
});
const $ZodBoolean = /* @__PURE__ */ $constructor("$ZodBoolean", (inst, def) => {
	$ZodType.init(inst, def);
	inst._zod.pattern = boolean$1;
	inst._zod.parse = (payload, _ctx) => {
		if (def.coerce) try {
			payload.value = Boolean(payload.value);
		} catch (_) {}
		const input = payload.value;
		if (typeof input === "boolean") return payload;
		payload.issues.push({
			expected: "boolean",
			code: "invalid_type",
			input,
			inst
		});
		return payload;
	};
});
const $ZodNull = /* @__PURE__ */ $constructor("$ZodNull", (inst, def) => {
	$ZodType.init(inst, def);
	inst._zod.pattern = _null$2;
	inst._zod.values = new Set([null]);
	inst._zod.parse = (payload, _ctx) => {
		const input = payload.value;
		if (input === null) return payload;
		payload.issues.push({
			expected: "null",
			code: "invalid_type",
			input,
			inst
		});
		return payload;
	};
});
const $ZodNever = /* @__PURE__ */ $constructor("$ZodNever", (inst, def) => {
	$ZodType.init(inst, def);
	inst._zod.parse = (payload, _ctx) => {
		payload.issues.push({
			expected: "never",
			code: "invalid_type",
			input: payload.value,
			inst
		});
		return payload;
	};
});
function handleArrayResult(result, final, index) {
	if (result.issues.length) final.issues.push(...prefixIssues(index, result.issues));
	final.value[index] = result.value;
}
const $ZodArray = /* @__PURE__ */ $constructor("$ZodArray", (inst, def) => {
	$ZodType.init(inst, def);
	inst._zod.parse = (payload, ctx) => {
		const input = payload.value;
		if (!Array.isArray(input)) {
			payload.issues.push({
				expected: "array",
				code: "invalid_type",
				input,
				inst
			});
			return payload;
		}
		payload.value = Array(input.length);
		const proms = [];
		for (let i = 0; i < input.length; i++) {
			const item = input[i];
			const result = def.element._zod.run({
				value: item,
				issues: []
			}, ctx);
			if (result instanceof Promise) proms.push(result.then((result) => handleArrayResult(result, payload, i)));
			else handleArrayResult(result, payload, i);
		}
		if (proms.length) return Promise.all(proms).then(() => payload);
		return payload;
	};
});
function handlePropertyResult(result, final, key, input, isOptionalIn, isOptionalOut) {
	const isPresent = key in input;
	if (result.issues.length) {
		if (isOptionalIn && isOptionalOut && !isPresent) return;
		final.issues.push(...prefixIssues(key, result.issues));
	}
	if (!isPresent && !isOptionalIn) {
		if (!result.issues.length) final.issues.push({
			code: "invalid_type",
			expected: "nonoptional",
			input: void 0,
			path: [key]
		});
		return;
	}
	if (result.value === void 0) {
		if (isPresent) final.value[key] = void 0;
	} else final.value[key] = result.value;
}
function normalizeDef(def) {
	const keys = Object.keys(def.shape);
	for (const k of keys) if (!def.shape?.[k]?._zod?.traits?.has("$ZodType")) throw new Error(`Invalid element at key "${k}": expected a Zod schema`);
	const okeys = optionalKeys(def.shape);
	return {
		...def,
		keys,
		keySet: new Set(keys),
		numKeys: keys.length,
		optionalKeys: new Set(okeys)
	};
}
function handleCatchall(proms, input, payload, ctx, def, inst) {
	const unrecognized = [];
	const keySet = def.keySet;
	const _catchall = def.catchall._zod;
	const t = _catchall.def.type;
	const isOptionalIn = _catchall.optin === "optional";
	const isOptionalOut = _catchall.optout === "optional";
	for (const key in input) {
		if (key === "__proto__") continue;
		if (keySet.has(key)) continue;
		if (t === "never") {
			unrecognized.push(key);
			continue;
		}
		const r = _catchall.run({
			value: input[key],
			issues: []
		}, ctx);
		if (r instanceof Promise) proms.push(r.then((r) => handlePropertyResult(r, payload, key, input, isOptionalIn, isOptionalOut)));
		else handlePropertyResult(r, payload, key, input, isOptionalIn, isOptionalOut);
	}
	if (unrecognized.length) payload.issues.push({
		code: "unrecognized_keys",
		keys: unrecognized,
		input,
		inst
	});
	if (!proms.length) return payload;
	return Promise.all(proms).then(() => {
		return payload;
	});
}
const $ZodObject = /* @__PURE__ */ $constructor("$ZodObject", (inst, def) => {
	$ZodType.init(inst, def);
	if (!Object.getOwnPropertyDescriptor(def, "shape")?.get) {
		const sh = def.shape;
		Object.defineProperty(def, "shape", { get: () => {
			const newSh = { ...sh };
			Object.defineProperty(def, "shape", { value: newSh });
			return newSh;
		} });
	}
	const _normalized = cached(() => normalizeDef(def));
	defineLazy(inst._zod, "propValues", () => {
		const shape = def.shape;
		const propValues = {};
		for (const key in shape) {
			const field = shape[key]._zod;
			if (field.values) {
				propValues[key] ?? (propValues[key] = /* @__PURE__ */ new Set());
				for (const v of field.values) propValues[key].add(v);
			}
		}
		return propValues;
	});
	const isObject$1 = isObject;
	const catchall = def.catchall;
	let value;
	inst._zod.parse = (payload, ctx) => {
		value ?? (value = _normalized.value);
		const input = payload.value;
		if (!isObject$1(input)) {
			payload.issues.push({
				expected: "object",
				code: "invalid_type",
				input,
				inst
			});
			return payload;
		}
		payload.value = {};
		const proms = [];
		const shape = value.shape;
		for (const key of value.keys) {
			const el = shape[key];
			const isOptionalIn = el._zod.optin === "optional";
			const isOptionalOut = el._zod.optout === "optional";
			const r = el._zod.run({
				value: input[key],
				issues: []
			}, ctx);
			if (r instanceof Promise) proms.push(r.then((r) => handlePropertyResult(r, payload, key, input, isOptionalIn, isOptionalOut)));
			else handlePropertyResult(r, payload, key, input, isOptionalIn, isOptionalOut);
		}
		if (!catchall) return proms.length ? Promise.all(proms).then(() => payload) : payload;
		return handleCatchall(proms, input, payload, ctx, _normalized.value, inst);
	};
});
function handleUnionResults(results, final, inst, ctx) {
	for (const result of results) if (result.issues.length === 0) {
		final.value = result.value;
		return final;
	}
	const nonaborted = results.filter((r) => !aborted(r));
	if (nonaborted.length === 1) {
		final.value = nonaborted[0].value;
		return nonaborted[0];
	}
	final.issues.push({
		code: "invalid_union",
		input: final.value,
		inst,
		errors: results.map((result) => result.issues.map((iss) => finalizeIssue(iss, ctx, config())))
	});
	return final;
}
const $ZodUnion = /* @__PURE__ */ $constructor("$ZodUnion", (inst, def) => {
	$ZodType.init(inst, def);
	defineLazy(inst._zod, "optin", () => def.options.some((o) => o._zod.optin === "optional") ? "optional" : void 0);
	defineLazy(inst._zod, "optout", () => def.options.some((o) => o._zod.optout === "optional") ? "optional" : void 0);
	defineLazy(inst._zod, "values", () => {
		if (def.options.every((o) => o._zod.values)) return new Set(def.options.flatMap((option) => Array.from(option._zod.values)));
	});
	defineLazy(inst._zod, "pattern", () => {
		if (def.options.every((o) => o._zod.pattern)) {
			const patterns = def.options.map((o) => o._zod.pattern);
			return new RegExp(`^(${patterns.map((p) => cleanRegex(p.source)).join("|")})$`);
		}
	});
	const first = def.options.length === 1 ? def.options[0]._zod.run : null;
	inst._zod.parse = (payload, ctx) => {
		if (first) return first(payload, ctx);
		let async = false;
		const results = [];
		for (const option of def.options) {
			const result = option._zod.run({
				value: payload.value,
				issues: []
			}, ctx);
			if (result instanceof Promise) {
				results.push(result);
				async = true;
			} else {
				if (result.issues.length === 0) return result;
				results.push(result);
			}
		}
		if (!async) return handleUnionResults(results, payload, inst, ctx);
		return Promise.all(results).then((results) => {
			return handleUnionResults(results, payload, inst, ctx);
		});
	};
});
const $ZodDiscriminatedUnion = /* @__PURE__ */ $constructor("$ZodDiscriminatedUnion", (inst, def) => {
	def.inclusive = false;
	$ZodUnion.init(inst, def);
	const _super = inst._zod.parse;
	defineLazy(inst._zod, "propValues", () => {
		const propValues = {};
		for (const option of def.options) {
			const pv = option._zod.propValues;
			if (!pv || Object.keys(pv).length === 0) throw new Error(`Invalid discriminated union option at index "${def.options.indexOf(option)}"`);
			for (const [k, v] of Object.entries(pv)) {
				if (!propValues[k]) propValues[k] = /* @__PURE__ */ new Set();
				for (const val of v) propValues[k].add(val);
			}
		}
		return propValues;
	});
	const disc = cached(() => {
		const opts = def.options;
		const map = /* @__PURE__ */ new Map();
		for (const o of opts) {
			const values = o._zod.propValues?.[def.discriminator];
			if (!values || values.size === 0) throw new Error(`Invalid discriminated union option at index "${def.options.indexOf(o)}"`);
			for (const v of values) {
				if (map.has(v)) throw new Error(`Duplicate discriminator value "${String(v)}"`);
				map.set(v, o);
			}
		}
		return map;
	});
	inst._zod.parse = (payload, ctx) => {
		const input = payload.value;
		if (!isObject(input)) {
			payload.issues.push({
				code: "invalid_type",
				expected: "object",
				input,
				inst
			});
			return payload;
		}
		const opt = disc.value.get(input?.[def.discriminator]);
		if (opt) return opt._zod.run(payload, ctx);
		if (def.unionFallback || ctx.direction === "backward") return _super(payload, ctx);
		payload.issues.push({
			code: "invalid_union",
			errors: [],
			note: "No matching discriminator",
			discriminator: def.discriminator,
			options: Array.from(disc.value.keys()),
			input,
			path: [def.discriminator],
			inst
		});
		return payload;
	};
});
const $ZodTuple = /* @__PURE__ */ $constructor("$ZodTuple", (inst, def) => {
	$ZodType.init(inst, def);
	const items = def.items;
	inst._zod.parse = (payload, ctx) => {
		const input = payload.value;
		if (!Array.isArray(input)) {
			payload.issues.push({
				input,
				inst,
				expected: "tuple",
				code: "invalid_type"
			});
			return payload;
		}
		payload.value = [];
		const proms = [];
		const optinStart = getTupleOptStart(items, "optin");
		const optoutStart = getTupleOptStart(items, "optout");
		if (!def.rest) {
			if (input.length < optinStart) {
				payload.issues.push({
					code: "too_small",
					minimum: optinStart,
					inclusive: true,
					input,
					inst,
					origin: "array"
				});
				return payload;
			}
			if (input.length > items.length) payload.issues.push({
				code: "too_big",
				maximum: items.length,
				inclusive: true,
				input,
				inst,
				origin: "array"
			});
		}
		const itemResults = new Array(items.length);
		for (let i = 0; i < items.length; i++) {
			const r = items[i]._zod.run({
				value: input[i],
				issues: []
			}, ctx);
			if (r instanceof Promise) proms.push(r.then((rr) => {
				itemResults[i] = rr;
			}));
			else itemResults[i] = r;
		}
		if (def.rest) {
			let i = items.length - 1;
			const rest = input.slice(items.length);
			for (const el of rest) {
				i++;
				const result = def.rest._zod.run({
					value: el,
					issues: []
				}, ctx);
				if (result instanceof Promise) proms.push(result.then((r) => handleTupleResult(r, payload, i)));
				else handleTupleResult(result, payload, i);
			}
		}
		if (proms.length) return Promise.all(proms).then(() => handleTupleResults(itemResults, payload, items, input, optoutStart));
		return handleTupleResults(itemResults, payload, items, input, optoutStart);
	};
});
function getTupleOptStart(items, key) {
	for (let i = items.length - 1; i >= 0; i--) if (items[i]._zod[key] !== "optional") return i + 1;
	return 0;
}
function handleTupleResult(result, final, index) {
	if (result.issues.length) final.issues.push(...prefixIssues(index, result.issues));
	final.value[index] = result.value;
}
function handleTupleResults(itemResults, final, items, input, optoutStart) {
	for (let i = 0; i < items.length; i++) {
		const r = itemResults[i];
		const isPresent = i < input.length;
		if (r.issues.length) {
			if (!isPresent && i >= optoutStart) {
				final.value.length = i;
				break;
			}
			final.issues.push(...prefixIssues(i, r.issues));
		}
		final.value[i] = r.value;
	}
	for (let i = final.value.length - 1; i >= input.length; i--) if (items[i]._zod.optout === "optional" && final.value[i] === void 0) final.value.length = i;
	else break;
	return final;
}
const $ZodRecord = /* @__PURE__ */ $constructor("$ZodRecord", (inst, def) => {
	$ZodType.init(inst, def);
	inst._zod.parse = (payload, ctx) => {
		const input = payload.value;
		if (!isPlainObject(input)) {
			payload.issues.push({
				expected: "record",
				code: "invalid_type",
				input,
				inst
			});
			return payload;
		}
		const proms = [];
		const values = def.keyType._zod.values;
		if (values) {
			payload.value = {};
			const recordKeys = /* @__PURE__ */ new Set();
			for (const key of values) if (typeof key === "string" || typeof key === "number" || typeof key === "symbol") {
				recordKeys.add(typeof key === "number" ? key.toString() : key);
				const keyResult = def.keyType._zod.run({
					value: key,
					issues: []
				}, ctx);
				if (keyResult instanceof Promise) throw new Error("Async schemas not supported in object keys currently");
				if (keyResult.issues.length) {
					payload.issues.push({
						code: "invalid_key",
						origin: "record",
						issues: keyResult.issues.map((iss) => finalizeIssue(iss, ctx, config())),
						input: key,
						path: [key],
						inst
					});
					continue;
				}
				const outKey = keyResult.value;
				const result = def.valueType._zod.run({
					value: input[key],
					issues: []
				}, ctx);
				if (result instanceof Promise) proms.push(result.then((result) => {
					if (result.issues.length) payload.issues.push(...prefixIssues(key, result.issues));
					payload.value[outKey] = result.value;
				}));
				else {
					if (result.issues.length) payload.issues.push(...prefixIssues(key, result.issues));
					payload.value[outKey] = result.value;
				}
			}
			let unrecognized;
			for (const key in input) if (!recordKeys.has(key)) {
				unrecognized = unrecognized ?? [];
				unrecognized.push(key);
			}
			if (unrecognized && unrecognized.length > 0) payload.issues.push({
				code: "unrecognized_keys",
				input,
				inst,
				keys: unrecognized
			});
		} else {
			payload.value = {};
			for (const key of Reflect.ownKeys(input)) {
				if (key === "__proto__") continue;
				if (!Object.prototype.propertyIsEnumerable.call(input, key)) continue;
				let keyResult = def.keyType._zod.run({
					value: key,
					issues: []
				}, ctx);
				if (keyResult instanceof Promise) throw new Error("Async schemas not supported in object keys currently");
				if (typeof key === "string" && number$1.test(key) && keyResult.issues.length) {
					const retryResult = def.keyType._zod.run({
						value: Number(key),
						issues: []
					}, ctx);
					if (retryResult instanceof Promise) throw new Error("Async schemas not supported in object keys currently");
					if (retryResult.issues.length === 0) keyResult = retryResult;
				}
				if (keyResult.issues.length) {
					if (def.mode === "loose") payload.value[key] = input[key];
					else payload.issues.push({
						code: "invalid_key",
						origin: "record",
						issues: keyResult.issues.map((iss) => finalizeIssue(iss, ctx, config())),
						input: key,
						path: [key],
						inst
					});
					continue;
				}
				const result = def.valueType._zod.run({
					value: input[key],
					issues: []
				}, ctx);
				if (result instanceof Promise) proms.push(result.then((result) => {
					if (result.issues.length) payload.issues.push(...prefixIssues(key, result.issues));
					payload.value[keyResult.value] = result.value;
				}));
				else {
					if (result.issues.length) payload.issues.push(...prefixIssues(key, result.issues));
					payload.value[keyResult.value] = result.value;
				}
			}
		}
		if (proms.length) return Promise.all(proms).then(() => payload);
		return payload;
	};
});
const $ZodEnum = /* @__PURE__ */ $constructor("$ZodEnum", (inst, def) => {
	$ZodType.init(inst, def);
	const values = getEnumValues(def.entries);
	const valuesSet = new Set(values);
	inst._zod.values = valuesSet;
	inst._zod.pattern = new RegExp(`^(${values.filter((k) => propertyKeyTypes.has(typeof k)).map((o) => typeof o === "string" ? escapeRegex(o) : o.toString()).join("|")})$`);
	inst._zod.parse = (payload, _ctx) => {
		const input = payload.value;
		if (valuesSet.has(input)) return payload;
		payload.issues.push({
			code: "invalid_value",
			values,
			input,
			inst
		});
		return payload;
	};
});
const $ZodLiteral = /* @__PURE__ */ $constructor("$ZodLiteral", (inst, def) => {
	$ZodType.init(inst, def);
	if (def.values.length === 0) throw new Error("Cannot create literal schema with no valid values");
	const values = new Set(def.values);
	inst._zod.values = values;
	inst._zod.pattern = new RegExp(`^(${def.values.map((o) => typeof o === "string" ? escapeRegex(o) : o ? escapeRegex(o.toString()) : String(o)).join("|")})$`);
	inst._zod.parse = (payload, _ctx) => {
		const input = payload.value;
		if (values.has(input)) return payload;
		payload.issues.push({
			code: "invalid_value",
			values: def.values,
			input,
			inst
		});
		return payload;
	};
});
function handleOptionalResult(result, input) {
	if (input === void 0 && (result.issues.length || result.fallback)) return {
		issues: [],
		value: void 0
	};
	return result;
}
const $ZodOptional = /* @__PURE__ */ $constructor("$ZodOptional", (inst, def) => {
	$ZodType.init(inst, def);
	inst._zod.optin = "optional";
	inst._zod.optout = "optional";
	defineLazy(inst._zod, "values", () => {
		return def.innerType._zod.values ? new Set([...def.innerType._zod.values, void 0]) : void 0;
	});
	defineLazy(inst._zod, "pattern", () => {
		const pattern = def.innerType._zod.pattern;
		return pattern ? new RegExp(`^(${cleanRegex(pattern.source)})?$`) : void 0;
	});
	inst._zod.parse = (payload, ctx) => {
		if (def.innerType._zod.optin === "optional") {
			const input = payload.value;
			const result = def.innerType._zod.run(payload, ctx);
			if (result instanceof Promise) return result.then((r) => handleOptionalResult(r, input));
			return handleOptionalResult(result, input);
		}
		if (payload.value === void 0) return payload;
		return def.innerType._zod.run(payload, ctx);
	};
});
const $ZodCustom = /* @__PURE__ */ $constructor("$ZodCustom", (inst, def) => {
	$ZodCheck.init(inst, def);
	$ZodType.init(inst, def);
	inst._zod.parse = (payload, _) => {
		return payload;
	};
	inst._zod.check = (payload) => {
		const input = payload.value;
		const r = def.fn(input);
		if (r instanceof Promise) return r.then((r) => handleRefineResult(r, payload, input, inst));
		handleRefineResult(r, payload, input, inst);
	};
});
function handleRefineResult(result, payload, input, inst) {
	if (!result) {
		const _iss = {
			code: "custom",
			input,
			inst,
			path: [...inst._zod.def.path ?? []],
			continue: !inst._zod.def.abort
		};
		if (inst._zod.def.params) _iss.params = inst._zod.def.params;
		payload.issues.push(issue(_iss));
	}
}

//#endregion
//#region node_modules/zod/v4/core/api.js
/* @__NO_SIDE_EFFECTS__ */
function _string(Class, params) {
	return new Class({
		type: "string",
		...normalizeParams(params)
	});
}
/* @__NO_SIDE_EFFECTS__ */
function _number(Class, params) {
	return new Class({
		type: "number",
		checks: [],
		...normalizeParams(params)
	});
}
/* @__NO_SIDE_EFFECTS__ */
function _boolean(Class, params) {
	return new Class({
		type: "boolean",
		...normalizeParams(params)
	});
}
/* @__NO_SIDE_EFFECTS__ */
function _null$1(Class, params) {
	return new Class({
		type: "null",
		...normalizeParams(params)
	});
}
/* @__NO_SIDE_EFFECTS__ */
function _never(Class, params) {
	return new Class({
		type: "never",
		...normalizeParams(params)
	});
}
/* @__NO_SIDE_EFFECTS__ */
function _lte(value, params) {
	return new $ZodCheckLessThan({
		check: "less_than",
		...normalizeParams(params),
		value,
		inclusive: true
	});
}
/* @__NO_SIDE_EFFECTS__ */
function _gte(value, params) {
	return new $ZodCheckGreaterThan({
		check: "greater_than",
		...normalizeParams(params),
		value,
		inclusive: true
	});
}
/* @__NO_SIDE_EFFECTS__ */
function _minLength(minimum, params) {
	return new $ZodCheckMinLength({
		check: "min_length",
		...normalizeParams(params),
		minimum
	});
}
/* @__NO_SIDE_EFFECTS__ */
function _regex(pattern, params) {
	return new $ZodCheckRegex({
		check: "string_format",
		format: "regex",
		...normalizeParams(params),
		pattern
	});
}
/* @__NO_SIDE_EFFECTS__ */
function _refine(Class, fn, _params) {
	return new Class({
		type: "custom",
		check: "custom",
		fn,
		...normalizeParams(_params)
	});
}

//#endregion
//#region node_modules/zod/v4/mini/schemas.js
const ZodMiniType = /* @__PURE__ */ $constructor("ZodMiniType", (inst, def) => {
	if (!inst._zod) throw new Error("Uninitialized schema in ZodMiniType.");
	$ZodType.init(inst, def);
	inst.def = def;
	inst.type = def.type;
	inst.parse = (data, params) => parse(inst, data, params, { callee: inst.parse });
	inst.safeParse = (data, params) => safeParse(inst, data, params);
	inst.parseAsync = async (data, params) => parseAsync(inst, data, params, { callee: inst.parseAsync });
	inst.safeParseAsync = async (data, params) => safeParseAsync(inst, data, params);
	inst.check = (...checks) => {
		return inst.clone({
			...def,
			checks: [...def.checks ?? [], ...checks.map((ch) => typeof ch === "function" ? { _zod: {
				check: ch,
				def: { check: "custom" },
				onattach: []
			} } : ch)]
		}, { parent: true });
	};
	inst.with = inst.check;
	inst.clone = (_def, params) => clone(inst, _def, params);
	inst.brand = () => inst;
	inst.register = ((reg, meta) => {
		reg.add(inst, meta);
		return inst;
	});
	inst.apply = (fn) => fn(inst);
});
const ZodMiniString = /* @__PURE__ */ $constructor("ZodMiniString", (inst, def) => {
	$ZodString.init(inst, def);
	ZodMiniType.init(inst, def);
});
/* @__NO_SIDE_EFFECTS__ */
function string(params) {
	return _string(ZodMiniString, params);
}
const ZodMiniNumber = /* @__PURE__ */ $constructor("ZodMiniNumber", (inst, def) => {
	$ZodNumber.init(inst, def);
	ZodMiniType.init(inst, def);
});
/* @__NO_SIDE_EFFECTS__ */
function number(params) {
	return _number(ZodMiniNumber, params);
}
const ZodMiniBoolean = /* @__PURE__ */ $constructor("ZodMiniBoolean", (inst, def) => {
	$ZodBoolean.init(inst, def);
	ZodMiniType.init(inst, def);
});
/* @__NO_SIDE_EFFECTS__ */
function boolean(params) {
	return _boolean(ZodMiniBoolean, params);
}
const ZodMiniNull = /* @__PURE__ */ $constructor("ZodMiniNull", (inst, def) => {
	$ZodNull.init(inst, def);
	ZodMiniType.init(inst, def);
});
/* @__NO_SIDE_EFFECTS__ */
function _null(params) {
	return _null$1(ZodMiniNull, params);
}
const ZodMiniNever = /* @__PURE__ */ $constructor("ZodMiniNever", (inst, def) => {
	$ZodNever.init(inst, def);
	ZodMiniType.init(inst, def);
});
/* @__NO_SIDE_EFFECTS__ */
function never(params) {
	return _never(ZodMiniNever, params);
}
const ZodMiniArray = /* @__PURE__ */ $constructor("ZodMiniArray", (inst, def) => {
	$ZodArray.init(inst, def);
	ZodMiniType.init(inst, def);
});
/* @__NO_SIDE_EFFECTS__ */
function array(element, params) {
	return new ZodMiniArray({
		type: "array",
		element,
		...normalizeParams(params)
	});
}
const ZodMiniObject = /* @__PURE__ */ $constructor("ZodMiniObject", (inst, def) => {
	$ZodObject.init(inst, def);
	ZodMiniType.init(inst, def);
	defineLazy(inst, "shape", () => def.shape);
});
/* @__NO_SIDE_EFFECTS__ */
function strictObject(shape, params) {
	return new ZodMiniObject({
		type: "object",
		shape,
		catchall: /* @__PURE__ */ never(),
		...normalizeParams(params)
	});
}
const ZodMiniUnion = /* @__PURE__ */ $constructor("ZodMiniUnion", (inst, def) => {
	$ZodUnion.init(inst, def);
	ZodMiniType.init(inst, def);
});
/* @__NO_SIDE_EFFECTS__ */
function union(options, params) {
	return new ZodMiniUnion({
		type: "union",
		options,
		...normalizeParams(params)
	});
}
const ZodMiniDiscriminatedUnion = /* @__PURE__ */ $constructor("ZodMiniDiscriminatedUnion", (inst, def) => {
	$ZodDiscriminatedUnion.init(inst, def);
	ZodMiniType.init(inst, def);
});
/* @__NO_SIDE_EFFECTS__ */
function discriminatedUnion(discriminator, options, params) {
	return new ZodMiniDiscriminatedUnion({
		type: "union",
		options,
		discriminator,
		...normalizeParams(params)
	});
}
const ZodMiniTuple = /* @__PURE__ */ $constructor("ZodMiniTuple", (inst, def) => {
	$ZodTuple.init(inst, def);
	ZodMiniType.init(inst, def);
});
/* @__NO_SIDE_EFFECTS__ */
function tuple(items, _paramsOrRest, _params) {
	const hasRest = _paramsOrRest instanceof $ZodType;
	const params = hasRest ? _params : _paramsOrRest;
	return new ZodMiniTuple({
		type: "tuple",
		items,
		rest: hasRest ? _paramsOrRest : null,
		...normalizeParams(params)
	});
}
const ZodMiniRecord = /* @__PURE__ */ $constructor("ZodMiniRecord", (inst, def) => {
	$ZodRecord.init(inst, def);
	ZodMiniType.init(inst, def);
});
/* @__NO_SIDE_EFFECTS__ */
function record(keyType, valueType, params) {
	if (!valueType || !valueType._zod) return new ZodMiniRecord({
		type: "record",
		keyType: /* @__PURE__ */ string(),
		valueType: keyType,
		...normalizeParams(valueType)
	});
	return new ZodMiniRecord({
		type: "record",
		keyType,
		valueType,
		...normalizeParams(params)
	});
}
const ZodMiniEnum = /* @__PURE__ */ $constructor("ZodMiniEnum", (inst, def) => {
	$ZodEnum.init(inst, def);
	ZodMiniType.init(inst, def);
	inst.options = Object.values(def.entries);
});
/* @__NO_SIDE_EFFECTS__ */
function _enum(values, params) {
	return new ZodMiniEnum({
		type: "enum",
		entries: Array.isArray(values) ? Object.fromEntries(values.map((v) => [v, v])) : values,
		...normalizeParams(params)
	});
}
const ZodMiniLiteral = /* @__PURE__ */ $constructor("ZodMiniLiteral", (inst, def) => {
	$ZodLiteral.init(inst, def);
	ZodMiniType.init(inst, def);
});
/* @__NO_SIDE_EFFECTS__ */
function literal(value, params) {
	return new ZodMiniLiteral({
		type: "literal",
		values: Array.isArray(value) ? value : [value],
		...normalizeParams(params)
	});
}
const ZodMiniOptional = /* @__PURE__ */ $constructor("ZodMiniOptional", (inst, def) => {
	$ZodOptional.init(inst, def);
	ZodMiniType.init(inst, def);
});
/* @__NO_SIDE_EFFECTS__ */
function optional(innerType) {
	return new ZodMiniOptional({
		type: "optional",
		innerType
	});
}
const ZodMiniCustom = /* @__PURE__ */ $constructor("ZodMiniCustom", (inst, def) => {
	$ZodCustom.init(inst, def);
	ZodMiniType.init(inst, def);
});
/* @__NO_SIDE_EFFECTS__ */
function refine(fn, _params = {}) {
	return _refine(ZodMiniCustom, fn, _params);
}

//#endregion
//#region lib/config/zod-primitives.ts
/**
* Shared zod validators used across both `ApplyConfigSchema` and
* `EditConfigSchema`. Both files repeated these identically; consolidating
* here so future tightening (e.g. trim semantics for NonEmptyString)
* lands in one place.
*/
/** Non-empty string. */
const NonEmptyString = string().check(_minLength(1));
/** Length value:
*   - `number` — pt (bare).
*   - `"Npt" / "Ncm" / "Nmm" / "Nin"` — explicit unit.
*  Validated eagerly at parse time so unit typos surface at config-read,
*  not at the emit site.
*/
const LengthValue = union([number(), string()]).check(refine((v) => {
	try {
		parseLengthPt(v);
		return true;
	} catch {
		return false;
	}
}, { error: (issue) => {
	try {
		parseLengthPt(issue.input);
		return "length parse failed";
	} catch (e) {
		return e.message;
	}
} }));
/** Paragraph indent value: a Length plus the special `"Nchar"` form (round-trips
*  Word's `w:firstLineChars` / `w:hangingChars`, auto-scales with font size).
*  Explicit `null` opts out of any indent attribute. */
const IndentValue = union([
	number(),
	string(),
	_null()
]).check(refine((v) => {
	try {
		parseIndent(v);
		return true;
	} catch {
		return false;
	}
}, { error: (issue) => {
	try {
		parseIndent(issue.input);
		return "indent parse failed";
	} catch (e) {
		return e.message;
	}
} }));
/** CSS-style padding: 1-4 Length values.
*   `4`             → all four edges 4pt
*   `[4]`           → same
*   `[4, 8]`        → vertical 4pt, horizontal 8pt
*   `[4, 8, 6]`     → top 4pt, horizontal 8pt, bottom 6pt
*   `[4, 8, 6, 2]`  → top / right / bottom / left
*  Negative values are refused (OOXML cell margins are unsigned).
*/
const PaddingLength = union([number(), string()]);
const PaddingValue = union([
	PaddingLength,
	tuple([PaddingLength]),
	tuple([PaddingLength, PaddingLength]),
	tuple([
		PaddingLength,
		PaddingLength,
		PaddingLength
	]),
	tuple([
		PaddingLength,
		PaddingLength,
		PaddingLength,
		PaddingLength
	])
]).check(refine((v) => {
	try {
		parsePadding(v);
		return true;
	} catch {
		return false;
	}
}, { error: (issue) => {
	try {
		parsePadding(issue.input);
		return "padding parse failed";
	} catch (e) {
		return e.message;
	}
} }));
/** Line spacing value carries the OOXML `lineRule` choice via type:
*   - `number` — multiplier (auto), e.g. `1.5`
*   - `"Npt" / "Ncm" / ...` — exact line height
*   - `{ atLeast: <Length> }` — at-least line height
*  No magnitude heuristic — bare numbers are always multipliers. */
const LineSpacingValue = union([
	number(),
	string(),
	strictObject({ atLeast: union([number(), string()]) })
]).check(refine((v) => {
	try {
		parseLineSpacing(v);
		return true;
	} catch {
		return false;
	}
}, { error: (issue) => {
	try {
		parseLineSpacing(issue.input);
		return "lineSpacing parse failed";
	} catch (e) {
		return e.message;
	}
} }));

//#endregion
//#region lib/config/zod-format.ts
/** Convert a zod issue.path (PropertyKey[]) to dotted/bracketed JS form,
* e.g. `["numbering","levels",0,"styleId"]` → `numbering.levels[0].styleId`. */
function formatPath(path) {
	let out = "";
	for (const seg of path) if (typeof seg === "number") out += `[${seg}]`;
	else out += out ? `.${String(seg)}` : String(seg);
	return out;
}
/** Walk `path` into `raw` and return the value at that location, or undefined
* if any segment is missing. Used so error messages can echo the actual
* value the agent supplied — `issue.input` is unreliable across issue codes
* in zod mini, the raw config is always the ground truth. */
function valueAtPath(raw, path) {
	let cur = raw;
	for (const seg of path) {
		if (cur == null || typeof cur !== "object") return void 0;
		cur = cur[seg];
	}
	return cur;
}
/** Format a single zod issue into an agent-readable line. Falls back to the
* issue's default message only when no specific code matches; for the common
* codes we render concrete details (unknown keys, expected types, allowed
* enum values) instead of the bare "Invalid input" zod emits at low cost. */
function formatIssue(issue, raw) {
	const got = valueAtPath(raw, issue.path);
	const gotStr = got === void 0 ? "(missing)" : JSON.stringify(got);
	switch (issue.code) {
		case "unrecognized_keys": {
			const keys = issue.keys ?? [];
			return `unknown ${keys.length === 1 ? "key" : "keys"} ${keys.map((k) => `"${k}"`).join(", ")}`;
		}
		case "invalid_type": {
			const expected = issue.expected ?? "value";
			return got === void 0 ? `missing required field (expected ${expected})` : `expected ${expected}, got ${gotStr}`;
		}
		case "invalid_value": return `invalid value ${gotStr}. Allowed: [${(issue.values ?? []).map((v) => JSON.stringify(v)).join(", ")}]`;
		case "too_small": {
			const min = issue.minimum;
			const origin = issue.origin;
			if (origin === "string") return `must be a non-empty string (got ${gotStr})`;
			if (origin === "array") return `must contain at least ${String(min)} item(s)`;
			return `value too small (minimum ${String(min)})`;
		}
		case "too_big": {
			const max = issue.maximum;
			return `value too large (maximum ${String(max)})`;
		}
		case "invalid_format": return `format check failed (${issue.format ?? "unknown"})`;
		default: return issue.message;
	}
}
/** When zod reports `invalid_union`, the real issues are nested per-variant.
* Pick the variant that matched the discriminator (its issues are not just
* `invalid_value` on a `type` / `op` field) and surface those. Recurse if
* the matched variant itself fails another union. */
function flattenUnionIssues(issue) {
	if (issue.code !== "invalid_union") return [issue];
	const variants = issue.errors ?? [];
	let best = null;
	for (const variant of variants) {
		const realIssues = variant.filter((sub) => !(sub.code === "invalid_value" && sub.path.length === 1 && (sub.path[0] === "type" || sub.path[0] === "op")));
		if (realIssues.length === 0) continue;
		const score = realIssues.length;
		if (!best || score < best.score) best = {
			issues: realIssues,
			score
		};
	}
	if (!best) {
		const first = variants.flat()[0];
		if (first) return [first];
		return [issue];
	}
	return best.issues.flatMap((sub) => {
		return flattenUnionIssues({
			...sub,
			path: [...issue.path, ...sub.path]
		});
	});
}
/** Format a ZodError as a multi-line, agent-readable string. Each issue
* gets a path prefix and either the supplied domain hint, a code-specific
* render from `formatIssue`, or zod's default message as a final fallback.
* `invalid_union` issues drill into the matched variant so the agent sees
* the inner field path, not just `edits[0]: Invalid input`. */
function formatZodError(error, raw, hint) {
	const lines = [];
	for (const top of error.issues) for (const issue of flattenUnionIssues(top)) {
		const pathStr = formatPath(issue.path);
		const msg = hint?.(issue, pathStr, raw) ?? null ?? formatIssue(issue, raw);
		lines.push(pathStr ? `${pathStr}: ${msg}` : msg);
	}
	return lines.join("\n");
}

//#endregion
//#region lib/config/edit-config-schema.ts
/**
* Zod-mini schema for the `edit` command. Mirrors the config-schema.ts
* conventions: strictObject everywhere, NonEmptyString helper, custom
* formatter that maps known issue shapes to agent-friendly hints.
*
* Single source of truth for the EditConfig shape. The TS types in
* edit-types.ts are inferred from these schemas, so adding/renaming a
* field flows through to every consumer's compile-time check.
*
* What lives here:
*   - shape (required vs optional, types, enums, regex constraints)
*   - rejection of unknown keys (strictObject everywhere)
*   - locally-derivable invariants (`range.from <= range.to`, color hex
*     format, level bounds 0..8, image dimensions positive)
*
* What does NOT live here (left to edit-engine / locator / fragment-emit):
*   - cross-doc validation: paragraph indices in range, table count,
*     numId existence, styleId existence — needs the parsed docx
*   - locator reachability under blockers — needs blocker scan results
*   - track-changes-mode constraint enforcement (none currently; Phase 1
*     supports tracking format ops via rPrChange / pPrChange too)
*/
/** RGB hex without leading "#". Six-hex form, case-insensitive. Same shape
* as styles in lib/config-schema.ts (which is laxer there for legacy
* reasons; we tighten here since this is a new surface). */
const ColorHex = string().check(_regex(/^[0-9a-fA-F]{6}$/));
const RunFormatSchema = strictObject({
	bold: optional(boolean()),
	italic: optional(boolean()),
	underline: optional(boolean()),
	strike: optional(boolean()),
	doubleStrike: optional(boolean()),
	charSpacing: optional(LengthValue),
	emphasisMark: optional(_enum([
		"dot",
		"comma",
		"circle",
		"underDot",
		"none"
	])),
	color: optional(ColorHex),
	fontLatin: optional(string()),
	fontCJK: optional(string()),
	size: optional(LengthValue),
	vertAlign: optional(_enum([
		"superscript",
		"subscript",
		"baseline"
	])),
	/** Right-to-left script run (`<w:rtl/>`). Pair with the paragraph's
	* `paraFormat.bidi` for RTL languages (Arabic / Hebrew); Word uses rtl
	* to pick the cs font and to mirror punctuation. Schema position is
	* after `vertAlign`, before `w:em`. */
	rtl: optional(boolean())
});
const ParagraphFormatSchema = strictObject({
	alignment: optional(_enum([
		"left",
		"center",
		"right",
		"both"
	])),
	spaceBefore: optional(LengthValue),
	spaceAfter: optional(LengthValue),
	lineSpacing: optional(LineSpacingValue),
	firstLineIndent: optional(IndentValue),
	hangingIndent: optional(IndentValue),
	indentLeft: optional(IndentValue),
	indentRight: optional(IndentValue),
	outlineLevel: optional(number().check(_gte(0), _lte(9))),
	/** RTL paragraph direction (`<w:bidi/>`). Flips the base direction of
	* the paragraph; pair with `runFormat.rtl` on the RTL runs. Schema
	* position is before `w:spacing`, so it is emitted first. */
	bidi: optional(boolean())
});
const InlineRunSchema = strictObject({
	text: string(),
	format: optional(RunFormatSchema)
});
/** Word bookmark name. ECMA-376 §17.13.7 mandates start-with-letter-or-
* underscore + alphanumerics; Word's UI lenient set adds hyphens. Capped
* at 40 chars to match Word's UI bookmark name limit (file format permits
* more but going above 40 is asking for cross-tool friction). */
const AnchorNameSchema = string().check(refine((s) => /^[A-Za-z_][A-Za-z0-9_-]{0,39}$/.test(s), { error: "anchor name must start with a letter or underscore and contain only letters, digits, underscores, or hyphens (max 40 chars). Example: \"fig-architecture\", \"ref_smith2024\"." }));
const RefToParagraphSchema = strictObject({
	type: literal("paragraph"),
	index: number().check(_gte(1))
});
const RefToAnchorSchema = strictObject({
	type: literal("anchor"),
	name: AnchorNameSchema
});
/** Inline cross-reference to an auto-numbered paragraph. Emits as an
* OOXML REF field; Word resolves the visible text from the target's
* bookmark at render time. Two locator forms:
*
*   `{ type: "paragraph", index: N }` — resolves against the pre-edit
*     document state. Target paragraph must be in the source and bound
*     to a numbering scheme (unless display === "full").
*
*   `{ type: "anchor", name: "..." }` — resolves against
*     (a) any ParagraphBlock.anchor / EquationBlock.anchor declared
*     anywhere in this same edits[] array (a pre-scan reserves the names
*     before emit, so refs can address anchors that emit later — forward
*     refs are fine), or (b) bookmarks already present in the source
*     document. Lets refs target paragraphs created in the same apply
*     run, which the paragraph-index form cannot do.
*
* The target must be bound to a numbering scheme when display is "label"
* or "number"; "full" resolves to the target paragraph's body text and
* works on any paragraph. See references/cross-references.md. */
const InlineRefSchema = strictObject({
	refTo: union([RefToParagraphSchema, RefToAnchorSchema]),
	display: optional(_enum([
		"full",
		"label",
		"number"
	])),
	format: optional(RunFormatSchema)
});
/** Inline math expression. Embedded as OMML alongside ordinary text runs in
* the paragraph. The `math` field is LaTeX (Temml subset — see
* references/equations.md for the supported command set and the v1 n-ary
* operand bug). No `format` field: math runs carry their own OMML typography
* (italic variables, upright numerals) which doesn't map cleanly to w:rPr. */
const InlineEquationSchema = strictObject({ math: NonEmptyString });
/** Hyperlink. `link` is a URI; `#name` prefix targets an internal bookmark
*  (any source bookmark or `anchor` declared on a Block in this apply run).
*  Anything else is treated as an external URL — `https:` / `http:` /
*  `mailto:` / `tel:` / `ftp:` etc.; Word resolves the protocol at click
*  time. The visible `text` is agent-supplied (unlike `refTo`, which lets
*  Word resolve the visible text from the target's numbering). `format`
*  overrides individual run properties; the Hyperlink character style
*  (color 0563C1, single underline — Word's default) is applied
*  automatically and injected into styles.xml when missing. */
const InlineHyperlinkSchema = strictObject({
	link: NonEmptyString.check(refine((s) => s.startsWith("#") ? /^#[A-Za-z_][A-Za-z0-9_-]{0,39}$/.test(s) : true, { error: "link \"#...\" (internal anchor) must start with a letter or underscore and contain only letters, digits, underscores, or hyphens (max 40 chars after #)." })),
	text: NonEmptyString,
	format: optional(RunFormatSchema)
});
/** Word complex field that resolves at render time:
*   "page"     → PAGE       (current page number)
*   "numPages" → NUMPAGES   (total page count)
*   "date"     → DATE       (current date, Word's default format)
*  Use in headers / footers / body text wherever the dynamic value matters.
*  Word's `updateFields` flag is set during apply so each field resolves on
*  next open without manual F9. */
const InlineFieldSchema = strictObject({
	field: _enum([
		"page",
		"numPages",
		"date"
	]),
	format: optional(RunFormatSchema)
});
/** STYLEREF field — renders the nearest paragraph bound to `styleRef`. Most
*  common use is "chapter title in the page header" by setting
*  `styleRef: "Heading 1"`. With `numberOnly: true` (the OOXML `\n` switch)
*  only the heading's auto-number renders, no body text. `styleRef` must
*  match an existing styleId in the source document (or one installed via
*  `styles[]` in this same apply); apply throws on a missing styleId. */
const InlineStyleRefSchema = strictObject({
	styleRef: NonEmptyString,
	numberOnly: optional(boolean()),
	format: optional(RunFormatSchema)
});
/** Inline break — emits `<w:r><w:br/></w:r>` (line) or `<w:r><w:br
*  w:type="page|column"/></w:r>`. "line" is a soft line break within the
*  paragraph; "page" / "column" force the next run onto the next page or
*  column. Use sparingly — prefer paragraph-level structure over inline
*  breaks where semantics allow. */
const InlineBreakSchema = strictObject({ break: union([
	literal("line"),
	literal("page"),
	literal("column")
]) });
const InlineNodeSchema = union([
	InlineRunSchema,
	InlineRefSchema,
	InlineEquationSchema,
	InlineHyperlinkSchema,
	InlineFieldSchema,
	InlineStyleRefSchema,
	InlineBreakSchema
]);
/** Plain string is shorthand for a single run with no inline formatting.
* The emitter expands strings on the fly — most paragraphs are plain text. */
const RichTextSchema = union([string(), array(InlineNodeSchema)]);
const NumberingRefSchema = strictObject({
	numId: NonEmptyString,
	level: number().check(_gte(0), _lte(8)),
	/** Force a counter restart at this paragraph. The engine forks a fresh
	* `<w:num>` pointing to the same abstractNumId with
	* `<w:startOverride val="1"/>` so this item and subsequent items on the
	* same scheme display 1, 2, 3 … from here. Use for mid-list resets that
	* scheme-level `restart` can't express (e.g. the second list inside a
	* section that already has `restart: "byHeading"` but needs a manual reset
	* at a non-heading boundary). */
	restart: optional(boolean())
});
const ParagraphBlockSchema = strictObject({
	type: literal("paragraph"),
	text: RichTextSchema,
	styleId: optional(NonEmptyString),
	paraFormat: optional(ParagraphFormatSchema),
	runFormat: optional(RunFormatSchema),
	numbering: optional(NumberingRefSchema),
	/** Optional stable name that InlineRefs in the same edits[] (or future
	* apply runs) can target via `refTo: { type: "anchor", name }`. The
	* engine wraps the emitted paragraph with `<w:bookmarkStart>` /
	* `<w:bookmarkEnd>` carrying this name. Names live in a flat namespace
	* with source bookmarks; collisions fail at pre-scan, before any
	* mutation. Refs can address an anchor regardless of declaration order
	* (forward refs are pre-scanned and resolved at backfill time). */
	anchor: optional(AnchorNameSchema)
});
const ImageBlockSchema = strictObject({
	type: literal("image"),
	src: NonEmptyString,
	width: LengthValue,
	height: LengthValue,
	alt: optional(string()),
	/** Optional paragraph-level style binding for the wrapping `<w:p>` —
	* lets a figure binding ("FigureImage" or similar) control centering /
	* keep-with-next / spaceBefore-After. Without it the image paragraph
	* has no pPr (default left alignment, no managed spacing). */
	styleId: optional(NonEmptyString),
	/** Per-call paragraph-format override. Same fields and override
	* precedence as on ParagraphBlock — `paraFormat` wins over `styleId`'s
	* same-named cascade output. Use for one-off tweaks (e.g. spaceAfter:0
	* to butt the figure against its caption) without minting a new style. */
	paraFormat: optional(ParagraphFormatSchema)
});
const PageBreakBlockSchema = strictObject({ type: literal("page-break") });
const HorizontalRuleBlockSchema = strictObject({ type: literal("horizontal-rule") });
/** Display equation. Becomes its own paragraph carrying `<m:oMathPara>`
* (centered by Word's default OMML rendering, or override via styleId /
* paraFormat). Caption + numbering follow the same caption-paragraph
* pattern as figures and tables — see references/equations.md. */
const EquationBlockSchema = strictObject({
	type: literal("equation"),
	latex: optional(NonEmptyString),
	omml: optional(NonEmptyString),
	styleId: optional(NonEmptyString),
	paraFormat: optional(ParagraphFormatSchema),
	captionId: optional(NonEmptyString),
	subGroup: optional(_enum(["start", "continue"])),
	anchor: optional(AnchorNameSchema)
}).check(refine((eq) => {
	if (eq.latex !== void 0 === (eq.omml !== void 0)) return false;
	if (eq.subGroup !== void 0 && eq.captionId === void 0) return false;
	if (eq.anchor !== void 0 && eq.captionId === void 0) return false;
	return true;
}, { error: (issue) => {
	const eq = issue.input;
	const hasLatex = eq.latex !== void 0;
	const hasOmml = eq.omml !== void 0;
	if (hasLatex && hasOmml) return "EquationBlock: latex and omml are mutually exclusive — set exactly one.";
	if (!hasLatex && !hasOmml) return "EquationBlock: must set one of latex or omml. Use latex for standard LaTeX input; omml is the escape hatch when temml fails on the expression.";
	if (eq.subGroup !== void 0 && eq.captionId === void 0) return `EquationBlock: subGroup="${eq.subGroup}" requires captionId. Set captionId to opt into caption numbering, or remove subGroup for a standalone equation.`;
	if (eq.anchor !== void 0 && eq.captionId === void 0) return `EquationBlock: anchor="${eq.anchor}" requires captionId — without numbering the bookmark has no resolved target for REF cross-references. Set captionId, or remove anchor.`;
	return "EquationBlock: invalid field combination";
} }));
/** A caption paragraph (figure title, table title, theorem statement,
* etc.). Replaces the older pattern of `{ type: "paragraph", styleId:
* "FigureCaption", ... }` paired with `numbering[]` binding — captions
* config carries the numbering shape; this block carries the text. */
const CaptionBlockSchema = strictObject({
	type: literal("caption"),
	captionId: NonEmptyString,
	text: string(),
	anchor: optional(AnchorNameSchema)
});
/** Counter reset marker for a caption identifier. Emits a hidden SEQ
* field at this position; counter sim resets accordingly. Use for
* appendix sequences or multi-section docs where the default
* outline-level restart isn't enough. */
const CaptionCounterResetSchema = strictObject({
	type: literal("caption-counter-reset"),
	captionId: NonEmptyString,
	newValue: optional(number().check(_gte(0)))
});
/** Subset of blocks that may appear INSIDE a table cell. Excludes
* TableBlock — nested tables via Block[] cell content are not supported in
* v1 (schema would be cyclic, blocking type inference). Agents needing a
* table inside an existing cell use a separate apply with a `cell` locator
* + insert op. */
const CellBlockSchema = union([
	ParagraphBlockSchema,
	ImageBlockSchema,
	PageBreakBlockSchema,
	HorizontalRuleBlockSchema,
	EquationBlockSchema,
	CaptionBlockSchema,
	CaptionCounterResetSchema
]);
/** Per-edge or per-table-side border specification. */
const BorderEdgeStyle = _enum([
	"single",
	"thick",
	"double",
	"dotted",
	"dashed"
]);
const BorderEdgeObjectSchema = strictObject({
	style: BorderEdgeStyle,
	/** Line size as a Length (`number` = pt, or `"Npt|Ncm|Nmm|Nin"`). Engine
	*  converts to OOXML's 1/8-pt units. Defaults: "single" → 0.5pt,
	*  "thick" → 1.5pt. */
	size: optional(LengthValue),
	/** Hex RGB without leading "#", or "auto" to inherit document defaults. */
	color: optional(union([ColorHex, literal("auto")]))
});
/** A single table border edge. String shorthand selects style with default
* size + color "auto"; object form for full control. `"none"` suppresses
* the edge — useful as a per-cell override. */
const BorderEdgeSchema = union([_enum([
	"none",
	"single",
	"thick",
	"double",
	"dotted",
	"dashed"
]), BorderEdgeObjectSchema]);
/** Border preset for the entire table.
*
*   "all"        every edge thin black (Word's default-looking table)
*   "none"       no borders anywhere
*   "outer"      only the four outer edges
*   "three-line" academic three-line table: thick top + thick bottom +
*                thin line under header row. Requires `headerRows >= 1`
*                to render the header-bottom line; with `headerRows: 0`
*                degrades silently to "top + bottom only".
*/
const BordersPresetSchema = _enum([
	"all",
	"none",
	"outer",
	"three-line"
]);
const BordersCustomSchema = strictObject({
	top: optional(BorderEdgeSchema),
	bottom: optional(BorderEdgeSchema),
	left: optional(BorderEdgeSchema),
	right: optional(BorderEdgeSchema),
	insideH: optional(BorderEdgeSchema),
	insideV: optional(BorderEdgeSchema)
});
const BordersSchema = union([BordersPresetSchema, BordersCustomSchema]);
/** Column width: `"auto"` (Word fits content) or a Length (pt by default,
* or `"Ncm" / "Nmm" / "Nin"`). Percentage widths were considered but require
* coordinated tblW + per-cell tcW emission with OOXML's fiftiethPercent
* units; deferred out of v1 — use fixed widths with `layout: "fixed"` for
* predictable sizing. */
const TableWidthSchema = union([literal("auto"), LengthValue]);
const ColSpecSchema = strictObject({ width: TableWidthSchema });
/** Per-cell properties available when content is wrapped in object form.
* `borders` here override the table-level borders for THIS cell only.
* Note that adjacent cells do not auto-coordinate — see tables.md. */
const TableCellObjectSchema = strictObject({
	content: union([RichTextSchema, array(CellBlockSchema)]),
	/** Number of columns this cell spans (gridSpan). Default 1. */
	colspan: optional(number().check(_gte(1))),
	/** Number of rows this cell spans (vMerge restart). Engine inserts
	* continuation cells in subsequent rows at the same column position;
	* the agent must NOT declare cells at those claimed positions. */
	rowspan: optional(number().check(_gte(1))),
	vAlign: optional(_enum([
		"top",
		"center",
		"bottom"
	])),
	borders: optional(BordersCustomSchema),
	/** Cell background color, hex RGB. */
	shading: optional(ColorHex),
	/** Per-cell padding override — wins over `TableBlock.padding`. CSS shorthand
	* (1-4 Lengths). All four edges are emitted as `<w:tcMar>` children; omit
	* the field entirely to inherit. */
	padding: optional(PaddingValue),
	/** Cell text direction (`<w:tcPr><w:textDirection>`). Word's UI exposes
	* three values; the OOXML enum has two more (tbRlV, tbLtlV — legacy). */
	textDirection: optional(_enum([
		"lrTb",
		"tbRl",
		"btLr",
		"lrTbV",
		"tbLtlV"
	])),
	/** Cell table-formula field (Word's Layout 公式 / Table Formula).
	* Must start with `=` — e.g. `"=SUM(ABOVE)"`, `"=AVERAGE(LEFT)"`,
	* `"=SUM(A1:B3)"`, `"=B2*C2"`. Emitted as a 5-run complex field
	* (`=…` instruction); Word re-evaluates on F9 / open-with-updateFields.
	* The cell's plain-string `content` is used as the cached result text
	* Word displays before the first field update; pass the precomputed
	* value (or leave `content: ""` for an empty placeholder — the cell
	* still renders because the field skeleton provides runs). Non-string
	* content with `formula` is ignored (single-field cell). */
	formula: optional(NonEmptyString)
});
/** Cell content in four progressive forms:
*
*   string             plain text, single paragraph, no formatting
*   InlineNode[]       mixed-format text or inline cross-refs, single para
*   Block[]            multi-paragraph / images / (no nested tables — see below)
*   { content, ... }   any of the above wrapped in an object that also
*                      carries spans, vAlign, per-cell borders, shading
*
* The first three are unambiguously discriminated: string by type, the two
* arrays by element shape (Block has `type` literal, InlineNode has only
* `text` or `refTo`). The object form is recognized by its `content` key.
*
* Block[] excludes nested TableBlock in v1 to keep the schema acyclic
* (recursion breaks TS inference). To put a table inside an existing
* cell, use a separate apply with a `cell` locator + insert op.
*/
const TableCellSchema = union([
	RichTextSchema,
	array(CellBlockSchema),
	TableCellObjectSchema
]);
const TableBlockSchema = strictObject({
	type: literal("table"),
	rows: array(array(TableCellSchema).check(_minLength(1))).check(_minLength(1)),
	/** Number of top rows that repeat as the header on page breaks
	* (`<w:tblHeader/>`). Default 0 — no repeating header. Does NOT
	* auto-bold the header text; bind a styled paragraph via
	* `headerStyle`, or format each header cell explicitly. */
	headerRows: optional(number().check(_gte(0))),
	/** styleId to apply to each cell paragraph in the header rows (top
	* `headerRows` rows). Cells that already carry an explicit styleId
	* via Block[] form win — this is a default, not a override. */
	headerStyle: optional(NonEmptyString),
	/** Per-column widths. Length must match the effective column count
	* (declared cells + cells claimed by ongoing rowspans, expanded by
	* colspans). When omitted, engine emits one `<w:gridCol>` per
	* effective column, seeded with an equal share of the usable width. */
	cols: optional(array(ColSpecSchema)),
	/** Table-level borders. Default `"all"`. */
	borders: optional(BordersSchema),
	/** Horizontal alignment of the table on the page. Default Word
	* behavior (no `<w:jc>` emitted) is left. Academic / formal
	* documents typically center tables. */
	alignment: optional(_enum([
		"left",
		"center",
		"right"
	])),
	/** Default vertical alignment for every cell whose object form
	* doesn't carry its own `vAlign`. Skill default is `"center"`
	* (academic / formal typography norm); set `"top"` for form-style
	* layouts where labels should hug the top of each cell. The
	* per-cell `vAlign` on the object form still wins per-cell. */
	vAlign: optional(_enum([
		"top",
		"center",
		"bottom"
	])),
	/** Column-width interpretation. `"autofit"` (default) lets Word
	* adjust columns to content; `"fixed"` enforces declared widths even
	* if total exceeds page width (content may overflow). */
	layout: optional(_enum(["fixed", "autofit"])),
	/** RTL table direction (`<w:tblPr><w:bidiVisual/>`). Mirrors the column
	* order so column 1 renders on the right — Word's 表格→从右向左.
	* Combine with RTL cell paragraphs for full RTL tables. */
	bidiVisual: optional(boolean()),
	/** Default cell padding for every cell without its own `padding`. CSS
	* shorthand (1-4 Lengths). Emitted as `<w:tblCellMar>` on tblPr — all
	* four edges are written, overriding any inherited TableNormal margins.
	* When omitted, the `"three-line"` preset injects top/bottom 4pt
	* (left/right inherit TableNormal); other presets emit nothing and
	* inherit fully. Set `padding: 0` to flatten. */
	padding: optional(PaddingValue),
	/** Write-time row sort (static equivalent of Word's 布局→排序 with
	* "无标题行" semantics). `column` is a 1-based index into the
	* effective column grid. `order` defaults to `"asc"`. Header rows
	* (`headerRows`) are pinned in place; only body rows are reordered.
	* Sort key: the target cell's plain text (strings, inline-run text,
	* or first text of object-form content). Keys that are numeric on
	* BOTH sides compare numerically (commas stripped — `1,024` works);
	* otherwise comparison is locale-aware `"zh-Hans-CN"` collation, so
	* Chinese sorts by pinyin. Rowspans are rejected under `sort` —
	* reordering rows would orphan vMerge continuations; sort the data
	* before passing it to apply instead. */
	sort: optional(strictObject({
		column: number().check(_gte(1)),
		order: optional(_enum([
			"asc",
			"desc"
		]))
	}))
}).check(refine((block) => block.headerRows === void 0 || block.headerRows <= block.rows.length, { error: "table: headerRows cannot exceed rows.length" }));
const BlockSchema = union([
	ParagraphBlockSchema,
	ImageBlockSchema,
	PageBreakBlockSchema,
	HorizontalRuleBlockSchema,
	TableBlockSchema,
	EquationBlockSchema,
	CaptionBlockSchema,
	CaptionCounterResetSchema
]);
const FragmentSchema = array(BlockSchema);
const ParagraphLocatorSchema = strictObject({
	type: literal("paragraph"),
	index: number().check(_gte(1))
});
const RangeLocatorSchema = strictObject({
	type: literal("range"),
	from: number().check(_gte(1)),
	to: number().check(_gte(1))
}).check(refine((loc) => loc.from <= loc.to, { error: "range locator: from must be <= to" }));
const CellLocatorSchema = strictObject({
	type: literal("cell"),
	table: number().check(_gte(1)),
	row: number().check(_gte(1)),
	col: number().check(_gte(1)),
	/** 1-based paragraph index WITHIN this cell. When set, the locator
	* resolves to just that one paragraph instead of all paragraphs in
	* the cell. Pair with `to` for a contiguous range within the cell. */
	paragraph: optional(number().check(_gte(1))),
	/** 1-based "to" paragraph index within the cell, inclusive. Only
	* meaningful with `paragraph`; defines a range [paragraph, to]. */
	to: optional(number().check(_gte(1)))
}).check(refine((loc) => !(loc.to !== void 0 && loc.paragraph === void 0), { error: "cell locator: `to` requires `paragraph` to be set" })).check(refine((loc) => loc.to === void 0 || loc.paragraph === void 0 || loc.to >= loc.paragraph, { error: "cell locator: `to` must be >= `paragraph`" }));
const HeadingLocatorSchema = strictObject({
	type: literal("heading"),
	text: NonEmptyString,
	level: optional(number().check(_gte(0), _lte(9)))
});
const WholeBodyLocatorSchema = strictObject({ type: literal("whole-body") });
/** Run-level locator. Targets one specific <w:r> within a paragraph — used
* by `set-run` to replace blank/placeholder run text while preserving the
* surrounding label runs. Pick the run by 1-based `runIndex`, or by `blank`
* (Kth run whose text is whitespace-only and rPr carries `<w:u/>` — typical
* form-fill placeholder). When both omitted, defaults to the first blank
* run (`blank: 1`).
*
* Two forms:
*   Global: `{ type: "run", paragraph: N, blank?, runIndex? }` — targets
*     paragraph N in the indexed (body + layout-table-cell) scope.
*   Cell:   `{ type: "run", table: T, row: R, col: C, paragraph: K, blank?,
*              runIndex? }` — targets paragraph K inside data-table cell
*     (T, R, C). Mirrors the `cell` locator coordinate scheme. */
const RunLocatorGlobalSchema = strictObject({
	type: literal("run"),
	paragraph: number().check(_gte(1)),
	blank: optional(number().check(_gte(1))),
	runIndex: optional(number().check(_gte(1)))
});
const RunLocatorCellSchema = strictObject({
	type: literal("run"),
	table: number().check(_gte(1)),
	row: number().check(_gte(1)),
	col: number().check(_gte(1)),
	paragraph: number().check(_gte(1)),
	blank: optional(number().check(_gte(1))),
	runIndex: optional(number().check(_gte(1)))
});
const RunLocatorSchema = union([RunLocatorGlobalSchema, RunLocatorCellSchema]).check(refine((loc) => !(loc.blank !== void 0 && loc.runIndex !== void 0), { error: "run locator: pass either `blank` or `runIndex`, not both" }));
const LocatorSchema = union([
	ParagraphLocatorSchema,
	RangeLocatorSchema,
	CellLocatorSchema,
	HeadingLocatorSchema,
	WholeBodyLocatorSchema
]);
const ReplaceOpSchema = strictObject({
	op: literal("replace"),
	at: LocatorSchema,
	with: FragmentSchema,
	/** When true, allow replacing paragraphs that contain SEQ / REF /
	* other complex fields. Default false (blocker rejects them). Use
	* for caption / cross-ref iteration: a previous apply emitted SEQ
	* fields, you want to rebuild the cell content. Revisions
	* (<w:ins> / <w:del>) and SDT controls are still blocking. */
	overwriteFields: optional(boolean())
});
const InsertBeforeOpSchema = strictObject({
	op: literal("insert-before"),
	at: LocatorSchema,
	content: FragmentSchema
});
const InsertAfterOpSchema = strictObject({
	op: literal("insert-after"),
	at: LocatorSchema,
	content: FragmentSchema
});
const DeleteOpSchema = strictObject({
	op: literal("delete"),
	at: LocatorSchema
});
const FormatOpSchema = strictObject({
	op: literal("format"),
	at: LocatorSchema,
	styleId: optional(NonEmptyString),
	runFormat: optional(RunFormatSchema),
	paraFormat: optional(ParagraphFormatSchema),
	clearDirect: optional(union([array(union([literal("pPr"), literal("rPr")])), literal("all")]))
}).check(refine((op) => !!(op.styleId || op.runFormat || op.paraFormat || op.clearDirect), { error: "format op needs at least one of: styleId, runFormat, paraFormat, clearDirect" }));
const SetRunOpSchema = strictObject({
	op: literal("set-run"),
	at: RunLocatorSchema,
	with: string(),
	format: optional(RunFormatSchema)
});
/** Caption body-text edit. Targets an existing caption paragraph by
* anchor name or by (captionId, 1-based index in body order). Replaces
* the runs after the primary bookmarkEnd; the SEQ / STYLEREF fields and
* bookmark itself stay intact so cross-references continue to resolve.
*
* Throws when target is EquationBlock (no body to edit) — agents
* needing to change an equation use delete + re-emit instead. */
const EditCaptionTargetSchema = union([strictObject({ anchor: AnchorNameSchema }), strictObject({
	captionId: NonEmptyString,
	index: number().check(_gte(1))
})]);
const EditCaptionOpSchema = strictObject({
	op: literal("edit-caption"),
	target: EditCaptionTargetSchema,
	text: string()
});
const MergeOpSchema = strictObject({
	op: literal("merge"),
	at: LocatorSchema,
	/** Keep the pPr of which paragraph: "first" (default) or "last". */
	keepPPr: optional(union([literal("first"), literal("last")]))
});
const EditOpSchema = discriminatedUnion("op", [
	ReplaceOpSchema,
	InsertBeforeOpSchema,
	InsertAfterOpSchema,
	DeleteOpSchema,
	FormatOpSchema,
	SetRunOpSchema,
	EditCaptionOpSchema,
	MergeOpSchema
]);
const EditConfigSchema = strictObject({
	source: NonEmptyString,
	output: NonEmptyString,
	edits: array(EditOpSchema).check(_minLength(1)),
	trackChanges: optional(boolean()),
	/** See ApplyConfigSchema.author — same semantics: written to revision
	* markup's `w:author` attribute when trackChanges is on; omitted means
	* Word shows "Unknown Author". Never defaulted to a tool brand. */
	author: optional(NonEmptyString)
}).check(refine((cfg) => !(cfg.author !== void 0 && cfg.author.trim() === ""), { error: "author: must be non-empty / non-whitespace when set. To leave revisions unattributed, omit the field entirely (Word will display 'Unknown Author')." }), refine((cfg) => !(cfg.author !== void 0 && cfg.trackChanges !== true), { error: "author: only meaningful when `trackChanges: true`. Either enable trackChanges or remove the author field." }));

//#endregion
//#region lib/config/config-schema.ts
/**
* Zod-mini schema for the `apply` CLI's config. Replaces the hand-written
* `typeof` / `Set.has()` validation that previously lived inline in
* apply-styles.ts.
*
* What lives here:
*   - shape (required vs optional, types, enums)
*   - rejection of unknown keys (strictObject everywhere)
*   - non-empty constraints on critical strings (id / name / lvlText / ...)
*   - pure-config invariants that don't need runtime docx context
*     (e.g. stripPrefixPatterns must be ordered by descending placeholder
*     count — fully derivable from the strings themselves)
*
* What does NOT live here:
*   - cross-field checks that need the parsed docx (styleId existence in
*     stylesDoc, paragraph-index validity, fingerprint resolution, name
*     collision detection). Those stay in apply-styles.ts where the runtime
*     context is available.
*   - regex compilation for pattern_rules. The compiled RegExp is consumed
*     downstream, so compile-once-and-store stays in apply-styles.ts.
*
* The error formatter (`formatConfigError`) maps known zod issues to
* domain-specific hints so the agent gets the same "stripPrefixPatterns
* belongs INSIDE each level" / "Allowed: [tab, space, nothing]. Omit to
* auto-infer..." guidance as before. Generic issues fall back to zod's
* default messages with paths.
*/
/** Fields shared between Mode B (`styles[i].*`) direct values and the
* `styles[i].overrides` block. Kept as a plain object so it can be spread
* into both schemas. */
const styleFormatFields = {
	basedOn: optional(string()),
	fontLatin: optional(string()),
	fontCJK: optional(string()),
	size: optional(LengthValue),
	bold: optional(boolean()),
	italic: optional(boolean()),
	color: optional(string()),
	alignment: optional(_enum([
		"left",
		"center",
		"right",
		"both"
	])),
	lineSpacing: optional(LineSpacingValue),
	spaceBefore: optional(LengthValue),
	spaceAfter: optional(LengthValue),
	firstLineIndent: optional(IndentValue),
	hangingIndent: optional(IndentValue),
	indentLeft: optional(IndentValue),
	indentRight: optional(IndentValue),
	outlineLevel: optional(number().check(_gte(0), _lte(9))),
	vertAlign: optional(_enum([
		"superscript",
		"subscript",
		"baseline"
	]))
};
const StyleOverridesSchema = strictObject(styleFormatFields);
const StyleEntrySchema = strictObject({
	id: NonEmptyString,
	name: optional(NonEmptyString),
	fromParagraph: optional(number()),
	...styleFormatFields,
	overrides: optional(StyleOverridesSchema)
});
const NumRPrSchema = strictObject({
	fontLatin: optional(string()),
	fontCJK: optional(string()),
	size: optional(LengthValue),
	bold: optional(boolean()),
	italic: optional(boolean()),
	color: optional(string())
});
const NumLevelSchema = strictObject({
	level: number().check(_gte(0), _lte(8)),
	numFmt: NonEmptyString,
	lvlText: NonEmptyString,
	styleId: NonEmptyString,
	start: optional(number()),
	stripPrefixPatterns: optional(array(string())),
	suff: optional(_enum([
		"tab",
		"space",
		"nothing"
	])),
	numRPr: optional(NumRPrSchema),
	isLgl: optional(boolean()),
	restart: optional(union([_enum([
		"continuous",
		"perInstance",
		"byHeading"
	]), strictObject({ atStyleChange: NonEmptyString })]))
}).check(refine((lvl) => {
	const counts = (lvl.stripPrefixPatterns ?? []).map((p) => (p.match(/%\d/g) ?? []).length);
	for (let j = 1; j < counts.length; j++) if (counts[j] > counts[j - 1]) return false;
	return true;
}, { error: (issue) => {
	const patterns = issue.input.stripPrefixPatterns ?? [];
	const counts = patterns.map((p) => (p.match(/%\d/g) ?? []).length);
	for (let j = 1; j < counts.length; j++) if (counts[j] > counts[j - 1]) return `stripPrefixPatterns must be ordered by descending placeholder count — "${patterns[j - 1]}" (${counts[j - 1]} placeholders) before "${patterns[j]}" (${counts[j]} placeholders) means the shorter pattern matches first and strips a prefix the longer one wanted. Reorder so longer patterns come first, e.g. ["%1.%2", "%1."].`;
	return "stripPrefixPatterns ordering invalid";
} }));
const NumberingSchema = strictObject({
	/** Optional explicit numId. When set, the engine pins this scheme to that
	* id (clones / creates as needed). Use to make block-level
	* `numbering: { numId: K }` references resolve deterministically.
	* Two schemes requesting the same numId cause apply to throw. */
	numId: optional(number().check(_gte(1))),
	levels: array(NumLevelSchema).check(_minLength(1))
});
/** Numeric format for SEQ counters. Maps to Word's `\*` switches —
* see lib/edit/fields/seq-field.ts. */
const CaptionFormatSchema = _enum([
	"arabic",
	"alphabetic",
	"ALPHABETIC",
	"roman",
	"ROMAN",
	"chinese",
	"chinese-formal"
]);
const CaptionSubCounterSchema = strictObject({
	format: optional(CaptionFormatSchema),
	prefix: optional(string()),
	suffix: optional(string())
});
/** A chapterPrefix entry. Bare string = use heading's native rendering
* (e.g. H1 with chineseCounting numFmt renders "一"). Object form with
* `format` forces the chapter number to render in that format — useful
* for the common Chinese academic style where H1 displays "第一章" but
* captions need Arabic "1.1". */
const ChapterPrefixEntrySchema = union([NonEmptyString, strictObject({
	styleId: NonEmptyString,
	format: optional(CaptionFormatSchema)
})]);
/** Per-identifier caption configuration. Block-level `captionId`
* references the key. `styleId` is required — every caption identifier
* must declare what paragraph style to use; no implicit fallback. */
const CaptionEntrySchema = strictObject({
	prefix: optional(string()),
	/** Text inserted between `prefix` and the chapter counter, e.g. `" "`
	* turns prefix `"图"` + chapter `2` + separator `.` + counter `1`
	* into `图 2.1` instead of `图2.1`. Defaults to "" (direct concat, the
	* historical behaviour). Only meaningful when a chapterPrefix is
	* declared — without chapter numbering the caption is `图1` either way
	* and the separator would just add a stray gap. */
	labelSeparator: optional(string()),
	suffix: optional(string()),
	format: optional(CaptionFormatSchema),
	chapterPrefix: optional(array(ChapterPrefixEntrySchema)),
	chapterSeparator: optional(string()),
	bodySeparator: optional(string()),
	styleId: NonEmptyString,
	subCounter: optional(CaptionSubCounterSchema)
});
/** Caption identifier — free string EXCEPT the `_chap_` prefix, which
* the engine reserves for the hidden auto-chapter SEQ counter paired
* with chapterPrefix `format` overrides. */
const CaptionIdentifierSchema = string().check(refine((s) => s.length > 0 && !s.startsWith("_chap_"), { error: "caption identifier must be a non-empty string and must not start with \"_chap_\" (engine-reserved for hidden auto-chapter counters paired with chapterPrefix format overrides)." }));
/** Map of identifier → caption entry. Identifier is a free string;
* conventional values "Equation" / "Figure" / "Table" / "Theorem" live
* in ref docs, not the schema. */
const CaptionsSchema = record(CaptionIdentifierSchema, CaptionEntrySchema);
const TemplateSchema = strictObject({
	source: NonEmptyString,
	styles: array(NonEmptyString).check(_minLength(1)),
	importNumbering: optional(boolean())
});
const ThemeFontsSchema = strictObject({
	majorLatin: optional(string()),
	majorEastAsia: optional(string()),
	minorLatin: optional(string()),
	minorEastAsia: optional(string())
});
const ThemeSchema = strictObject({ fonts: optional(ThemeFontsSchema) });
const PAPER_SIZES = [
	"A4",
	"A3",
	"A5",
	"Letter",
	"Legal",
	"B5",
	"16K"
];
const PaperSizeEnum = _enum(PAPER_SIZES);
const PaperSizeCustom = strictObject({
	width: LengthValue,
	height: LengthValue
});
const PaperSizeSchema = union([PaperSizeEnum, PaperSizeCustom]);
const MarginsSchema = strictObject({
	top: optional(LengthValue),
	bottom: optional(LengthValue),
	left: optional(LengthValue),
	right: optional(LengthValue),
	header: optional(LengthValue),
	footer: optional(LengthValue),
	gutter: optional(LengthValue)
});
/** Columns: `number` shorthand = equal-width count. Object form supports both
*  equal-width (specify `count`) and unequal-width (specify `widths` array).
*  `widths` and `count` are mutually exclusive — `widths.length` IS the count.
*  When `widths` is given, OOXML `w:equalWidth="false"` is set automatically. */
const ColumnsSchema = union([number().check(_gte(1)), strictObject({
	count: optional(number().check(_gte(1))),
	space: optional(LengthValue),
	separator: optional(boolean()),
	widths: optional(array(LengthValue).check(_minLength(1))),
	spaces: optional(array(LengthValue))
}).check(refine((c) => !(c.count !== void 0 && c.widths !== void 0), { error: "columns: specify either `count` (equal-width) or `widths` (unequal-width), not both. `widths.length` is the column count." })).check(refine((c) => c.count !== void 0 || c.widths !== void 0, { error: "columns object: at least one of `count` or `widths` is required." })).check(refine((c) => !c.spaces || c.widths !== void 0 && c.widths.length >= 2 && c.spaces.length === c.widths.length - 1, { error: "columns.spaces: only valid alongside `widths` with at least 2 columns; length must equal `widths.length - 1` (one space between each pair)." }))]);
/** Page numbering format + restart per section. `fmt` controls the numeral
*  shape; `start` restarts the counter at that value (omit to continue). */
const PgNumTypeSchema = strictObject({
	fmt: optional(_enum([
		"decimal",
		"upperRoman",
		"lowerRoman",
		"upperLetter",
		"lowerLetter"
	])),
	start: optional(number().check(_gte(1)))
});
/** Fields shared between top-level (default for all sections) and per-section
*  overrides under `sections`. */
const pageSetupFields = {
	paperSize: optional(PaperSizeSchema),
	orientation: optional(_enum(["portrait", "landscape"])),
	margins: optional(MarginsSchema),
	columns: optional(ColumnsSchema),
	pgNumType: optional(PgNumTypeSchema)
};
const PageSetupSectionFieldsSchema = strictObject(pageSetupFields);
/** Per-section overrides keyed by section selector:
*   "N"     — section N (1-based, matching `inspect_section`)
*   "N-M"   — sections N through M inclusive
*  No "all" selector — top-level fields already serve that role. */
const PageSetupSectionsSchema = record(string(), PageSetupSectionFieldsSchema).check(refine((rec) => Object.keys(rec).every((k) => SECTION_SELECTOR_KEY_RE.test(k)), { error: (issue) => {
	const rec = issue.input;
	return `key "${Object.keys(rec).find((k) => !SECTION_SELECTOR_KEY_RE.test(k))}" is invalid. Use "N" (1-based section index) or "N-M" (inclusive range).`;
} }));
const PageSetupSchema = strictObject({
	...pageSetupFields,
	sections: optional(PageSetupSectionsSchema)
});
/** Block subset allowed inside a header / footer part.
*  Excluded:
*    - page-break  (meaningless inside HF — Word ignores it)
*    - equation / caption / caption-counter-reset  (numbering-counter state
*      lives in body; using these inside HF double-increments and breaks the
*      counter sim)
*  Table is allowed — the common header layout "left text | center text |
*  right text" is typically a single-row 3-column borderless table. */
const HeaderFooterBlockSchema = union([
	ParagraphBlockSchema,
	ImageBlockSchema,
	HorizontalRuleBlockSchema,
	TableBlockSchema
]);
/** Paragraph styleIds that misbehave inside HF: heading IDs carry outline
*  level + numbering bindings the body's chapter SEQ counter watches, so a
*  Heading1 paragraph in a header would re-trigger chapter restarts. Title
*  / Subtitle / BodyText carry doc-flow assumptions that look wrong in HF
*  context. Agents wanting "this header text is bold 14pt" should use the
*  built-in `Header` / `Footer` styleId plus a paraFormat / runFormat
*  override, or mint a custom non-heading styleId. */
const HF_FORBIDDEN_STYLE_IDS = /^(Heading[1-9]|Title|Subtitle|BodyText)$/;
/** Block `type` values disallowed at any nesting depth inside HF content,
*  including inside table cells (where `CellBlockSchema` would otherwise
*  permit them). caption / caption-counter-reset / equation pin counters
*  to the body's numbering / SEQ state; emitting them inside HF would
*  double-increment and break the counter sim. page-break is a no-op
*  inside an HF part — Word ignores it. */
const HF_FORBIDDEN_NESTED_BLOCK_TYPES = new Set([
	"caption",
	"caption-counter-reset",
	"equation",
	"page-break"
]);
/** DFS the block tree pushing into table cells via every cell-content
*  shape that can carry Block[]. Visitor returns true to stop (violation
*  found). Returns true when visitor signals stop, false otherwise. */
function hfWalk(blocks, visit) {
	const stack = [...blocks];
	while (stack.length > 0) {
		const node = stack.pop();
		if (!node || typeof node !== "object") continue;
		const obj = node;
		if (visit(obj)) return true;
		if (obj.type === "table" && Array.isArray(obj.rows)) for (const row of obj.rows) {
			if (!Array.isArray(row)) continue;
			for (const cell of row) if (Array.isArray(cell)) stack.push(...cell);
			else if (cell && typeof cell === "object" && "content" in cell) {
				const content = cell.content;
				if (Array.isArray(content)) stack.push(...content);
			}
		}
	}
	return false;
}
const HeaderFooterContentSchema = array(HeaderFooterBlockSchema).check(refine((blocks) => !hfWalk(blocks, (obj) => {
	const t = obj.type;
	return typeof t === "string" && HF_FORBIDDEN_NESTED_BLOCK_TYPES.has(t);
}), { error: "header/footer cannot contain caption / equation / caption-counter-reset / page-break blocks (allowed at any nesting depth, including inside table cells). caption/equation/caption-counter-reset bind to body counters and would double-increment; page-break is a no-op inside header/footer. Move these blocks to body edits[]." })).check(refine((blocks) => !hfWalk(blocks, (obj) => {
	if (obj.type !== "paragraph") return false;
	const id = obj.styleId;
	return typeof id === "string" && HF_FORBIDDEN_STYLE_IDS.test(id);
}), { error: "header/footer paragraphs cannot use heading/body-text styleIds (Heading1..9 / Title / Subtitle / BodyText) at any nesting depth. These styles carry outline-level and numbering bindings that misbehave outside the body. Use the built-in `Header` / `Footer` styleId with a paraFormat / runFormat override, or mint a custom non-heading styleId." }));
/** Per-surface variants. ECMA-376 sectPr supports three reference types:
*
*   default  — applies to every page that isn't covered by first / even
*   first    — applies to the section's first page; needs <w:titlePg/> on
*              the sectPr (engine auto-sets when this variant is declared)
*   even     — applies to even-numbered pages; needs
*              <w:evenAndOddHeaders/> in settings.xml (engine auto-sets
*              when this variant is declared anywhere)
*
*  At least one variant must be declared per surface. Empty array `[]` is
*  legal: it means "the variant exists for the trigger flags (titlePg /
*  evenAndOdd) but renders no content" — useful for blanking the cover
*  page's header.
*
*  `underline` (header) / `overline` (footer) — separator line between the
*  HF surface and the body. Engine attaches `<w:pBdr>` to the variant's
*  endpoint paragraph (header → last; footer → first). `true` is sugar for
*  `"single"` 0.5pt black. Skipped silently when the variant has no
*  paragraph at the endpoint (empty `[]`, or endpoint block is a table /
*  image). */
const HeaderVariantsSchema = strictObject({
	default: optional(HeaderFooterContentSchema),
	first: optional(HeaderFooterContentSchema),
	even: optional(HeaderFooterContentSchema),
	underline: optional(union([BorderEdgeSchema, literal(true)]))
}).check(refine((v) => v.default !== void 0 || v.first !== void 0 || v.even !== void 0, { error: "header: at least one of `default` / `first` / `even` must be declared" }));
const FooterVariantsSchema = strictObject({
	default: optional(HeaderFooterContentSchema),
	first: optional(HeaderFooterContentSchema),
	even: optional(HeaderFooterContentSchema),
	overline: optional(union([BorderEdgeSchema, literal(true)]))
}).check(refine((v) => v.default !== void 0 || v.first !== void 0 || v.even !== void 0, { error: "footer: at least one of `default` / `first` / `even` must be declared" }));
/** Per-section override: replaces the top-level header/footer surface wholesale
*  for the matched sections. A section that names only `header` keeps the
*  top-level `footer` (and vice versa). At least one surface must be named —
*  empty entries serve no purpose. */
const HeaderFooterSectionFieldsSchema = strictObject({
	header: optional(HeaderVariantsSchema),
	footer: optional(FooterVariantsSchema)
}).check(refine((s) => s.header !== void 0 || s.footer !== void 0, { error: "headerFooter.sections entry: at least one of `header` or `footer` must be declared" }));
const HeaderFooterSectionsSchema = record(string(), HeaderFooterSectionFieldsSchema).check(refine((rec) => Object.keys(rec).every((k) => SECTION_SELECTOR_KEY_RE.test(k)), { error: (issue) => {
	const rec = issue.input;
	return `key "${Object.keys(rec).find((k) => !SECTION_SELECTOR_KEY_RE.test(k))}" is invalid. Use "N" (1-based section index) or "N-M" (inclusive range).`;
} }));
const HeaderFooterSchema = strictObject({
	header: optional(HeaderVariantsSchema),
	footer: optional(FooterVariantsSchema),
	sections: optional(HeaderFooterSectionsSchema)
}).check(refine((hf) => hf.header !== void 0 || hf.footer !== void 0 || hf.sections !== void 0 && Object.keys(hf.sections).length > 0, { error: "headerFooter: at least one of `header`, `footer`, or non-empty `sections` must be declared" }));
const AssignmentSchema = strictObject({
	para: number(),
	action: _enum([
		"keep",
		"restyle",
		"flag"
	]),
	style: optional(string()),
	reason: optional(string())
});
const BulkRuleSchema = strictObject({
	fingerprint: NonEmptyString,
	style: NonEmptyString
});
const PatternRuleSchema = strictObject({
	regex: NonEmptyString,
	flags: optional(string()),
	style: NonEmptyString,
	stripMatch: optional(boolean())
});
/** Advisory lints. None of these block a write — they surface text the
* engine thinks the author should look at twice. Each is individually
* switchable so a project with deliberate conventions can silence just
* the notice it disagrees with instead of losing the whole report. */
const LintSchema = strictObject({
	/** "Possible Pangu spacing" notice: CJK↔Latin runs in author-supplied
	* text that already contain manual spaces. Word's autoSpace would add
	* the gap anyway, so a hand-typed space is usually redundant — but some
	* authors space on purpose. false = don't report. */
	panguSpacing: optional(boolean())
});
/** One `<w:settings>` child written through the raw channel. `attrs` keys
* may be bare (`val`) or prefixed (`w:val`, `r:id`); bare keys resolve to
* the settings namespace. `children` covers the one-level content models
* the settings schema actually has (`mailMerge`, `docVars`, `compat`,
* `rsids`, `footnotePr` / `endnotePr`) — child order inside a parent is
* the author's responsibility, matching how the spec's own content models
* are written. */
const SettingsChildSchema = strictObject({
	name: NonEmptyString,
	attrs: optional(record(string(), union([string(), number(), boolean()]))),
	text: optional(string())
});
const SettingsElementSchema = strictObject({
	name: NonEmptyString,
	attrs: optional(record(string(), union([string(), number(), boolean()]))),
	text: optional(string()),
	children: optional(array(SettingsChildSchema))
});
const SettingsSchema = strictObject({ elements: array(SettingsElementSchema).check(_minLength(1)) });
/** `docProps/core.xml` + `docProps/app.xml` maintenance. On by default —
*  without those two parts, AUTHOR / NUMWORDS fields have nothing to read
*  and the file-properties pane (title / author / word count) is blank.
*  `creator` is only applied when the part is new or still carries the
*  tool's own previous value, so a human-set author is never overwritten. */
const DocPropsSchema = strictObject({
	enabled: optional(boolean()),
	title: optional(string()),
	creator: optional(string())
});
const ApplyConfigSchema = strictObject({
	source: optional(NonEmptyString),
	output: NonEmptyString,
	dryRun: optional(boolean()),
	/** Merge adjacent identically-formatted runs (and strip `rsid*` /
	* `proofErr` noise) in `word/document.xml` before anything else runs.
	* Opt-in, because it changes the run layout: a `set-run` locator written
	* against the fragment-laden original would shift. Reach for it when
	* editing a document of unknown provenance (Word-authored, or round-tripped
	* through another tool) where text-based locators come up empty. Rendering
	* is unchanged and tracked-change structure is preserved.
	* CLI: `--merge-runs`. Standalone equivalent: `scripts/merge_runs.js`. */
	mergeRuns: optional(boolean()),
	lint: optional(LintSchema),
	settings: optional(SettingsSchema),
	docProps: optional(DocPropsSchema),
	allowValidationWarnings: optional(boolean()),
	template: optional(TemplateSchema),
	theme: optional(ThemeSchema),
	pageSetup: optional(PageSetupSchema),
	headerFooter: optional(HeaderFooterSchema),
	styles: optional(array(StyleEntrySchema)),
	numbering: optional(union([NumberingSchema, array(NumberingSchema)])),
	captions: optional(CaptionsSchema),
	assignments: optional(array(AssignmentSchema)),
	bulk_rules: optional(array(BulkRuleSchema)),
	pattern_rules: optional(array(PatternRuleSchema)),
	requirements: optional(record(string(), string())),
	exclude: optional(array(number())),
	edits: optional(array(EditOpSchema)),
	trackChanges: optional(boolean()),
	/** Author name written to `<w:ins>` / `<w:del>` / `<w:rPrChange>` /
	* `<w:pPrChange>`'s `w:author` attribute when `trackChanges: true`.
	* No default — when omitted, an empty string is written (Word displays
	* "Unknown Author"). Set to claim authorship for review handoff;
	* never defaulted to a tool brand. */
	author: optional(NonEmptyString)
}).check(refine((cfg) => !(cfg.template !== void 0 && cfg.source === void 0), { error: "template: incompatible with blank-source mode (omitted `source`). Blank-source starts a clean document; template-import is a delta operation that transplants styles into an existing host with its own style cascade — combining them is conceptually inconsistent. Declare `source` when you want template-import, or drop `template` to use blank-source without the transplant." }), refine((cfg) => !(cfg.author !== void 0 && cfg.author.trim() === ""), { error: "author: must be non-empty / non-whitespace when set. To leave revisions unattributed, omit the field entirely (Word will display 'Unknown Author')." }), refine((cfg) => !(cfg.author !== void 0 && cfg.trackChanges !== true), { error: "author: only meaningful when `trackChanges: true`. Either enable trackChanges or remove the author field." }));
/** Domain-specific hints that override or augment zod's default message for
* particular issue shapes. Returning null means "use the default zod
* message". The hints replicate the agent-friendly guidance the hand-written
* checks used to emit. */
/** Walk up the issue path to the parent op entry and return its `op` literal,
* or undefined when the path doesn't pass through an edit op. Shared by the
* set-run hint with `edit-config-schema.ts` — same logic, same lookup. */
function lookupOpLiteral(raw, issuePath) {
	if (issuePath.length < 2) return void 0;
	const opPath = issuePath.slice(0, -2);
	let cursor = raw;
	for (const seg of opPath) if (cursor && typeof cursor === "object") cursor = cursor[seg];
	else return;
	if (cursor && typeof cursor === "object" && "op" in cursor) return cursor.op;
}
const customHint = (issue, pathStr, raw) => {
	if (issue.code === "invalid_value" && pathStr.endsWith(".at.type")) {
		if (lookupOpLiteral(raw, issue.path) === "set-run") return `set-run requires at.type === "run" (use a RunLocator: { type: "run", paragraph, blank|runIndex }). For paragraph-range edits use op: "replace" / "format" / "insert-before" / "insert-after" instead.`;
	}
	if (issue.code === "unrecognized_keys" && pathStr === "numbering" && "keys" in issue && Array.isArray(issue.keys) && issue.keys.includes("stripPrefixPatterns")) return "unknown field \"stripPrefixPatterns\" on numbering — stripPrefixPatterns belongs INSIDE each level (numbering.levels[i].stripPrefixPatterns), not at the top of numbering.";
	if (issue.code === "invalid_value" && pathStr.endsWith(".suff") && "values" in issue) {
		const allowed = issue.values;
		const got = valueAtPath(raw, issue.path);
		return `invalid suff: ${JSON.stringify(got)}. Allowed: [${allowed.map((v) => String(v)).join(", ")}]. Omit to auto-infer from trailing whitespace in lvlText (0 spaces → nothing, 1 → space, 2+ → tab).`;
	}
	return null;
};
function formatConfigError(error, raw) {
	return formatZodError(error, raw, customHint);
}
/** Parse + validate a raw JSON config. Throws an Error with a multi-line,
* path-annotated message on shape failure. The thrown message preserves the
* agent-friendly hints the previous hand-written validator emitted. */
function parseConfig(raw) {
	const result = safeParse(ApplyConfigSchema, raw);
	if (!result.success) throw new Error(formatConfigError(result.error, raw));
	return result.data;
}

//#endregion
//#region lib/shared/cli-helpers.ts
async function runCli(spec) {
	const args = process.argv.slice(2);
	const dryRun = args.includes("--dry-run");
	const allowValidationWarnings = args.includes("--allow-validation-warnings");
	const mergeRuns = args.includes("--merge-runs");
	const configPath = args.filter((a) => !a.startsWith("--"))[0];
	if (!configPath) {
		console.error(`Usage: node scripts/${spec.script} [--dry-run] [--allow-validation-warnings] [--merge-runs] <config.json>`);
		process.exit(1);
	}
	let raw;
	try {
		raw = JSON.parse(readFileSync(configPath, "utf8"));
	} catch (err) {
		console.error(`Cannot read config: ${err.message}`);
		process.exit(1);
	}
	if (raw && typeof raw === "object") {
		const rawObj = raw;
		if (dryRun) rawObj.dryRun = true;
		if (allowValidationWarnings) rawObj.allowValidationWarnings = true;
		if (mergeRuns) rawObj.mergeRuns = true;
	}
	let config;
	try {
		config = parseConfig(raw);
	} catch (err) {
		console.error(err.message);
		process.exit(1);
	}
	try {
		spec.validate(config);
	} catch (err) {
		console.error(err.message);
		process.exit(1);
	}
	const source = config.source !== void 0 ? resolve(config.source) : getBlankTemplatePath();
	const output = resolve(config.output);
	if (!config.dryRun && source === output) {
		console.error("output must differ from source");
		process.exit(1);
	}
	if (!existsSync(source)) {
		console.error(config.source !== void 0 ? `source not found: ${source}` : `bundled blank template not found at ${source} — run \`bun run build:skill\` to stage it`);
		process.exit(1);
	}
	try {
		await applyStyles(source, output, config);
	} catch (err) {
		if (existsSync(output)) try {
			unlinkSync(output);
		} catch {}
		console.error(`Error: ${err.message}`);
		process.exit(1);
	}
}

//#endregion
//#region skill/tools/apply.ts
/**
* CLI entry for `apply` — the unified docx mutation orchestrator.
*
* Accepts a config that can combine any of:
*   - `styles[]` — install / override named paragraph styles
*   - `numbering` — install / extend multi-level numbering schemes (single
*     scheme or array)
*   - `template` — import named styles from a template document
*   - `theme.fonts` — set the document's design-layer font scheme
*   - `assignments` / `pattern_rules` / `bulk_rules` / `exclude` — paragraph
*     restyling rules (rule precedence: exclude > assignments > pattern_rules
*     > bulk_rules)
*   - `edits[]` — surgical content insertions / replacements / deletions /
*     format ops
*   - `trackChanges` — emit edits as Word revision markup
*
* Pipeline order: see SKILL.md §Commands.
*
* Sparse by design: only declared blocks are applied. Untouched styles,
* numbering schemes, paragraphs, and theme stay as they are.
*
* Pure content-only edits (no style/numbering install) still go through
* `apply` with `edits[]` as the only populated block.
*/
runCli({
	command: "apply",
	script: "apply.js",
	validate(config) {
		const hasStyles = (config.styles?.length ?? 0) > 0;
		const hasNumbering = (() => {
			if (!config.numbering) return false;
			return (Array.isArray(config.numbering) ? config.numbering : [config.numbering]).some((s) => s.levels.length > 0);
		})();
		const hasTemplate = !!config.template?.styles?.length;
		const hasThemeOverride = !!config.theme?.fonts && Object.values(config.theme.fonts).some((v) => typeof v === "string" && v);
		const hasEdits = (config.edits?.length ?? 0) > 0;
		const hasCaptions = !!config.captions && Object.keys(config.captions).length > 0;
		const hasPageSetup = !!config.pageSetup && (config.pageSetup.paperSize !== void 0 || config.pageSetup.orientation !== void 0 || config.pageSetup.margins !== void 0 || config.pageSetup.columns !== void 0 || config.pageSetup.pgNumType !== void 0 || config.pageSetup.sections !== void 0 && Object.keys(config.pageSetup.sections).length > 0);
		const hasHeaderFooter = !!config.headerFooter;
		const hasMergeRuns = config.mergeRuns === true;
		if (!hasStyles && !hasNumbering && !hasTemplate && !hasThemeOverride && !hasEdits && !hasCaptions && !hasPageSetup && !hasHeaderFooter && !hasMergeRuns) throw new Error("config has no operation: provide at least one of styles[] (non-empty), numbering, template, theme.fonts, pageSetup, headerFooter, edits[], captions, or mergeRuns");
	}
});

//#endregion
export {  };