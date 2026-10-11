import { a as firstChildNS, d as textContent, h as NS, l as paragraphStyleId$1, o as getChildren, s as getChildrenNS } from "./xml-utils.js";
import { t as summarizeTable } from "./table-classifier.js";

//#region lib/config/edit-types.ts
function makeTrackContext(enabled, idAllocator, options) {
	return {
		enabled,
		nextId: () => idAllocator.next(),
		author: options?.author ?? "",
		date: options?.isoDate ?? (/* @__PURE__ */ new Date()).toISOString().replace(/\.\d{3}Z$/, "Z")
	};
}
/**
* Compile-time enforcement that every union variant is handled. Place in the
* `default` of a switch over a discriminated union — TypeScript narrows the
* residual type to `never`, so adding a new variant without updating the
* switch becomes a type error here.
*/
function assertNever(x) {
	throw new Error(`unhandled variant: ${JSON.stringify(x)}`);
}

//#endregion
//#region lib/edit/locator.ts
/** Walk the document in DocumentParser order and return every indexed
* paragraph with its element + container. Skips paragraphs inside data
* tables (they're unindexed; reachable only via cell locator). Also skips
* engine-managed scaffolding paragraphs — those carry a styleId starting
* with `_` (currently `_HiddenChapterCounter`, holds the hidden chapter
* SEQ counter as a sibling next to each H1 so STYLEREF on headings stays
* clean). Skipping them here means the agent-facing 1-based paragraph
* index matches user-visible content order, not the raw DOM order. */
function walkIndexedParagraphs(documentDoc) {
	const out = [];
	const root = documentDoc.documentElement;
	if (!root) return out;
	const body = firstChildNS(root, NS.w, "body");
	if (!body) return out;
	let nextIndex = 1;
	const recur = (parent) => {
		for (const child of getChildren(parent)) {
			if (child.namespaceURI !== NS.w) continue;
			if (child.localName === "p") {
				const styleId = paragraphStyleId$1(child);
				if (styleId !== void 0 && styleId.startsWith("_")) continue;
				out.push({
					index: nextIndex++,
					element: child,
					container: parent
				});
			} else if (child.localName === "tbl") {
				if (summarizeTable(child).classification === "layout") for (const tr of getChildrenNS(child, NS.w, "tr")) for (const tc of getChildrenNS(tr, NS.w, "tc")) recur(tc);
			}
		}
	};
	recur(body);
	return out;
}
function walkTopLevelTables(documentDoc) {
	const out = [];
	const root = documentDoc.documentElement;
	if (!root) return out;
	const body = firstChildNS(root, NS.w, "body");
	if (!body) return out;
	let i = 1;
	for (const child of getChildren(body)) if (child.namespaceURI === NS.w && child.localName === "tbl") out.push({
		tableIndex: i++,
		element: child
	});
	return out;
}
function buildResolverContext(documentDoc, parsed) {
	const root = documentDoc.documentElement;
	if (!root) throw new Error("document has no root element");
	const body = firstChildNS(root, NS.w, "body");
	if (!body) throw new Error("document has no body");
	const indexed = walkIndexedParagraphs(documentDoc);
	const indexByElement = /* @__PURE__ */ new Map();
	for (const p of indexed) indexByElement.set(p.element, p.index);
	return {
		documentDoc,
		body,
		indexed,
		indexByElement,
		tables: walkTopLevelTables(documentDoc),
		parsed
	};
}
function resolveLocator(loc, ctx) {
	switch (loc.type) {
		case "paragraph": return resolveParagraph(loc.index, ctx);
		case "range": return resolveRange(loc.from, loc.to, ctx);
		case "cell": return resolveCell(loc.table, loc.row, loc.col, loc.paragraph, loc.to, ctx);
		case "heading": return resolveHeading(loc.text, loc.level, ctx);
		case "whole-body": return resolveWholeBody(ctx);
		default: return assertNever(loc);
	}
}
function resolveParagraph(index, ctx) {
	const hit = ctx.indexed[index - 1];
	if (!hit || hit.index !== index) {
		const max = ctx.indexed.length;
		throw new Error(`paragraph #${index} not found. Document has ${max} indexed paragraph(s) (range: #1${max ? `–#${max}` : ""}). Paragraphs inside data tables are unindexed and only reachable via a cell locator.`);
	}
	return {
		paragraphs: [hit.element],
		container: hit.container
	};
}
function resolveRange(from, to, ctx) {
	const fromHit = ctx.indexed[from - 1];
	const toHit = ctx.indexed[to - 1];
	if (!fromHit || fromHit.index !== from) throw new Error(`range.from: paragraph #${from} not found (max #${ctx.indexed.length}).`);
	if (!toHit || toHit.index !== to) throw new Error(`range.to: paragraph #${to} not found (max #${ctx.indexed.length}).`);
	if (fromHit.container !== toHit.container) throw new Error(`range #${from}–#${to} crosses a structural boundary (one endpoint in body, the other inside a layout-table cell). Splitting cross-boundary ranges has no clean OOXML semantics; split into two separate edits.`);
	const paragraphs = [];
	for (let i = from - 1; i <= to - 1; i++) {
		const p = ctx.indexed[i];
		if (p.container !== fromHit.container) throw new Error(`range #${from}–#${to}: paragraph #${p.index} sits in a different container — range cannot span structural boundaries.`);
		paragraphs.push(p.element);
	}
	return {
		paragraphs,
		container: fromHit.container
	};
}
function resolveCell(table, row, col, paragraph, to, ctx) {
	if (table < 1 || table > ctx.tables.length) throw new Error(`cell.table: index ${table} out of range. Document has ${ctx.tables.length} top-level table(s); valid 1..${ctx.tables.length}.`);
	const tbl = ctx.tables[table - 1].element;
	const rows = getChildrenNS(tbl, NS.w, "tr");
	if (row < 1 || row > rows.length) throw new Error(`cell.row: index ${row} out of range. Table ${table} has ${rows.length} row(s); valid 1..${rows.length}.`);
	const cells = getChildrenNS(rows[row - 1], NS.w, "tc");
	if (col < 1 || col > cells.length) throw new Error(`cell.col: index ${col} out of range. Table ${table} row ${row} has ${cells.length} cell(s); valid 1..${cells.length}.`);
	const tc = cells[col - 1];
	const allParagraphs = getChildrenNS(tc, NS.w, "p");
	if (paragraph === void 0) return {
		paragraphs: allParagraphs,
		container: tc
	};
	const cellCount = allParagraphs.length;
	if (paragraph < 1 || paragraph > cellCount) throw new Error(`cell.paragraph: index ${paragraph} out of range. Table ${table} row ${row} col ${col} has ${cellCount} paragraph(s); valid 1..${cellCount}.`);
	const end = to ?? paragraph;
	if (end > cellCount) throw new Error(`cell.to: index ${end} out of range. Table ${table} row ${row} col ${col} has ${cellCount} paragraph(s); valid ${paragraph}..${cellCount}.`);
	return {
		paragraphs: allParagraphs.slice(paragraph - 1, end),
		container: tc
	};
}
function resolveHeading(text, level, ctx) {
	const target = text.trim();
	for (const p of ctx.parsed) {
		if (p.text.trim() !== target) continue;
		if (level !== void 0 && p.pPr.outlineLevel !== level) continue;
		if (level === void 0 && p.pPr.outlineLevel === void 0) continue;
		const hit = ctx.indexed[p.index - 1];
		if (!hit) continue;
		return {
			paragraphs: [hit.element],
			container: hit.container
		};
	}
	const lvlPart = level !== void 0 ? ` at outline level ${level}` : "";
	throw new Error(`heading "${text}"${lvlPart} not found. Use find_paragraphs to locate by regex, then refer by paragraph index instead.`);
}
function resolveWholeBody(ctx) {
	return {
		paragraphs: ctx.indexed.map((p) => p.element),
		container: ctx.body
	};
}
/** Resolve a `set-run` op's RunLocator. Returns the target paragraph element
* and the specific run element to mutate. The run is selected by:
*   - `runIndex: M`     → 1-based run index in the paragraph
*   - `blank: K`        → Kth (1-based) run that's a "blank" placeholder
*                         (whitespace-only text + rPr containing `<w:u/>`)
*   - neither given     → first blank run (= `blank: 1`)
*
* Accepts both locator forms:
*   - Global: `{ paragraph: N, ... }` — indexed scope (body + layout-table cells)
*   - Cell:   `{ table: T, row: R, col: C, paragraph: K, ... }` — data-table cell
*
* Throws with a clear message when the paragraph isn't found, the index is
* out of range, or the requested blank doesn't exist (lists what blanks ARE
* present so the agent can adjust). */
function resolveRunLocator(loc, ctx) {
	let pEl;
	let locusDesc;
	if ("table" in loc) {
		const cellResolved = resolveCell(loc.table, loc.row, loc.col, loc.paragraph, void 0, ctx);
		if (cellResolved.paragraphs.length !== 1) throw new Error(`run locator cell-form: expected exactly 1 paragraph, got ${cellResolved.paragraphs.length}. This is an internal error.`);
		pEl = cellResolved.paragraphs[0];
		locusDesc = `table ${loc.table} row ${loc.row} col ${loc.col} paragraph ${loc.paragraph}`;
	} else {
		const hit = ctx.indexed[loc.paragraph - 1];
		if (!hit || hit.index !== loc.paragraph) {
			const max = ctx.indexed.length;
			throw new Error(`paragraph #${loc.paragraph} not found. Document has ${max} indexed paragraph(s).`);
		}
		pEl = hit.element;
		locusDesc = `paragraph #${loc.paragraph}`;
	}
	const runs = getChildrenNS(pEl, NS.w, "r");
	if (runs.length === 0) throw new Error(`${locusDesc} has no runs to target.`);
	if (loc.runIndex !== void 0) {
		if (loc.runIndex < 1 || loc.runIndex > runs.length) throw new Error(`${locusDesc}: runIndex ${loc.runIndex} out of range (paragraph has ${runs.length} run(s); valid 1..${runs.length}).`);
		return {
			paragraph: pEl,
			run: runs[loc.runIndex - 1]
		};
	}
	const blankK = loc.blank ?? 1;
	const blanks = runs.filter(isBlankRun);
	if (blanks.length === 0) throw new Error(`${locusDesc}: no blank runs found. A blank run is one whose text is whitespace-only and rPr carries <w:u/> (typical form-fill placeholder). Use \`runIndex\` to target a specific run by 1-based index instead.`);
	if (blankK < 1 || blankK > blanks.length) throw new Error(`${locusDesc}: blank ${blankK} out of range (paragraph has ${blanks.length} blank run(s); valid 1..${blanks.length}).`);
	return {
		paragraph: pEl,
		run: blanks[blankK - 1]
	};
}
/** Heuristic: a "blank" run is one whose text content is whitespace-only and
* whose rPr declares underline (`<w:u/>`). Captures form-fill placeholder
* runs without false-positive on legitimate empty runs that happen to lack
* underline (those are usually inter-text spacers, not fillable slots). */
function isBlankRun(r) {
	const rPr = firstChildNS(r, NS.w, "rPr");
	if (!rPr) return false;
	if (!firstChildNS(rPr, NS.w, "u")) return false;
	let text = "";
	for (const c of getChildren(r)) if (c.namespaceURI === NS.w && c.localName === "t") text += textContent(c);
	return text.trim() === "";
}
/** Read the styleId of a paragraph element for blocker logging /
* heading disambiguation. Returns "Normal" when no explicit `<w:pStyle>`
* is set — Word's effective style at that absence, applied here because
* downstream blockers compare against concrete styleIds and want the
* effective value, not "unset". Tools that need to distinguish "absent"
* from "= Normal" call `paragraphStyleIdOrUndefined` directly. */
function paragraphStyleId(pEl) {
	return paragraphStyleId$1(pEl) ?? "Normal";
}
/** Plain-text content of a paragraph (concatenates <w:t> textContent across
* all runs). Used by blockers / debug logging. */
function paragraphText(pEl) {
	let out = "";
	for (const r of getChildrenNS(pEl, NS.w, "r")) for (const c of getChildren(r)) if (c.namespaceURI === NS.w && c.localName === "t") out += textContent(c);
	return out;
}
/** Read the trailing sectPr at body level (if any). Used for "append to
* body" semantics so we don't insert after the section descriptor. */
function trailingBodySectPr(body) {
	const children = getChildren(body);
	for (let i = children.length - 1; i >= 0; i--) {
		const c = children[i];
		if (c.namespaceURI !== NS.w) continue;
		if (c.localName === "sectPr") return c;
		if (c.localName === "p" || c.localName === "tbl") return null;
	}
	return null;
}

//#endregion
export { resolveRunLocator as a, walkTopLevelTables as c, resolveLocator as i, assertNever as l, paragraphStyleId as n, trailingBodySectPr as o, paragraphText as r, walkIndexedParagraphs as s, buildResolverContext as t, makeTrackContext as u };