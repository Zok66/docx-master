import { a as firstChildNS, f as wAttr, h as NS, o as getChildren } from "./xml-utils.js";

//#region lib/edit/blockers.ts
/**
* Detect paragraphs that aren't safe to edit. The edit engine consults this
* before applying any op; if any resolved target paragraph is blocked, the
* op is refused with a specific reason. Inspect tools surface the same data
* so the agent can plan around blockers before composing edits.
*
* What counts as a blocker:
*   - tracked-change: paragraph sits inside an existing <w:ins> / <w:del>.
*     Mutating it would nest tracked changes; Word's review UI handles the
*     nesting unevenly. Refuse.
*   - field: paragraph is inside a complex field region (between
*     <w:fldChar begin> and <w:fldChar end>). Editing field internals
*     breaks STYLEREF / TOC / REF / cross-reference. Refuse.
*   - sdt: paragraph is inside a <w:sdt> content control (block- or
*     inline-level). Content controls have their own update semantics
*     (see Phase 2 SDT-specialized path).
*
* Field-region detection runs a pseudo-state-machine over the body in
* document order: we increment depth on `begin` and decrement on `end`,
* marking every paragraph encountered while depth > 0. Field regions can
* span paragraph boundaries — `STYLEREF` in particular often does.
*/
const w = NS.w;
function detectBlockers(documentDoc, indexByElement) {
	const byElement = /* @__PURE__ */ new Map();
	const byIndex = /* @__PURE__ */ new Map();
	const root = documentDoc.documentElement;
	if (!root) return {
		byElement,
		byIndex
	};
	const body = firstChildNS(root, NS.w, "body");
	if (!body) return {
		byElement,
		byIndex
	};
	let fieldDepth = 0;
	const recordBlocker = (pEl, reason) => {
		if (byElement.has(pEl)) return;
		byElement.set(pEl, reason);
		const idx = indexByElement.get(pEl);
		if (idx !== void 0) byIndex.set(idx, reason);
	};
	/** Walk the descendants of an element, updating fieldDepth on fldChar
	* tokens and marking any <w:p> ancestor of a tracked-change / sdt /
	* field-region run. The `inheritedReason` is whatever ancestor blocker
	* already applies to the current subtree. */
	const walk = (node, currentParagraph, inheritedReason) => {
		if (node.namespaceURI === w && node.localName === "p") {
			currentParagraph = node;
			if (inheritedReason) recordBlocker(node, inheritedReason);
			if (fieldDepth > 0) recordBlocker(node, "field");
		}
		let scopeReason = inheritedReason;
		if (node.namespaceURI === w) {
			if (node.localName === "ins" || node.localName === "del") {
				scopeReason = "tracked-change";
				if (currentParagraph) recordBlocker(currentParagraph, "tracked-change");
			} else if (node.localName === "sdt") {
				scopeReason = "sdt";
				if (currentParagraph) recordBlocker(currentParagraph, "sdt");
			} else if (node.localName === "fldChar") {
				const ftype = wAttr(node, "fldCharType");
				if (ftype === "begin") {
					fieldDepth++;
					if (currentParagraph) recordBlocker(currentParagraph, "field");
				} else if (ftype === "end") {
					if (fieldDepth > 0) fieldDepth--;
				}
			}
		}
		for (const child of getChildren(node)) walk(child, currentParagraph, scopeReason);
	};
	walk(body, null, null);
	return {
		byElement,
		byIndex
	};
}
/** Reason → human-readable explanation for the agent. */
function explainBlockerReason(reason) {
	switch (reason) {
		case "tracked-change": return "paragraph already contains tracked changes (existing <w:ins>/<w:del>) — edit refused to avoid nested revision markup";
		case "field": return "paragraph is inside a complex field region (TOC / STYLEREF / cross-reference) — editing field internals breaks the field";
		case "sdt": return "paragraph is inside a content control (<w:sdt>) — use the SDT specialized path (Phase 2; not yet available)";
	}
}
/** Tally counts per reason for inspect_blockers' summary line. */
function summarizeBlockers(scan) {
	const out = {
		"tracked-change": 0,
		field: 0,
		sdt: 0
	};
	for (const reason of scan.byElement.values()) out[reason]++;
	return out;
}

//#endregion
export { explainBlockerReason as n, summarizeBlockers as r, detectBlockers as t };