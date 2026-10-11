import { h as NS, v as serializeXml } from "./xml-utils.js";

//#region lib/merge-runs.ts
/**
 * Coalesce adjacent identically-formatted runs in a WordprocessingML part.
 *
 * Word fragments a paragraph's text across many `<w:r>` elements (revision
 * ids, spell-check markers, editing history), so the string you can see on
 * the page is often not a contiguous string in `word/document.xml`. That
 * breaks character-level locating (`find_text`) and `set-run`, which look for
 * a phrase inside a single run. Merging first makes the source text findable.
 *
 * Per part, in order:
 *   1. strip every `<w:proofErr>` marker
 *   2. strip `rsid*` attributes from every run
 *   3. merge adjacent runs whose `<w:rPr>` is identical (both present and
 *      byte-equal, or both absent), appending the source run's non-rPr
 *      children onto the target run
 *   4. coalesce each run's adjacent `<w:t>` / `<w:delText>` children
 *
 * Rendering is unchanged. Text is compared the way Word draws it, not the way
 * the bytes sit in the file: an element without `xml:space="preserve"` has its
 * edge whitespace trimmed before it reaches the page, so
 * `<w:t>Hello </w:t><w:t>world</w:t>` reads "Helloworld" and merges to exactly
 * that — `xml:space="preserve"` is re-applied only when the merged text
 * actually carries edge whitespace.
 *
 * Runs sitting inside *different* `<w:ins>` / `<w:del>` wrappers are never
 * merged — that would collapse separate revisions into one and rewrite the
 * tracked-change structure. `<w:delText>` is a separate text channel and is
 * never folded into `<w:t>`.
 *
 * Semantics mirror `pro/scripts/merge_runs.py`; this is the zero-dependency
 * Node implementation, so `apply` can pre-process a document of unknown
 * provenance without crossing into the Python toolchain.
 */

const ELEMENT_NODE = 1;
const TEXT_NODE = 3;
const CDATA_NODE = 4;

/** Same character class Python's `str.strip(" \t\r\n")` uses. */
const XML_SPACE = new Set([" ", "\t", "\r", "\n"]);

function stripXmlSpace(s) {
	let i = 0;
	let j = s.length;
	while (i < j && XML_SPACE.has(s[i])) i++;
	while (j > i && XML_SPACE.has(s[j - 1])) j--;
	return s.slice(i, j);
}

function isW(node, local) {
	return node !== null && node.nodeType === ELEMENT_NODE && node.namespaceURI === NS.w && node.localName === local;
}

/** Every element with `localName === local` in the w namespace, root included. */
function collectW(root, local, out = []) {
	if (isW(root, local)) out.push(root);
	for (let c = root.firstChild; c; c = c.nextSibling) {
		if (c.nodeType === ELEMENT_NODE) collectW(c, local, out);
	}
	return out;
}

/** Direct children only. */
function childrenW(parent, local) {
	const out = [];
	for (let c = parent.firstChild; c; c = c.nextSibling) {
		if (isW(c, local)) out.push(c);
	}
	return out;
}

function firstChildW(parent, local) {
	for (let c = parent.firstChild; c; c = c.nextSibling) {
		if (isW(c, local)) return c;
	}
	return null;
}

function isRun(node) {
	return isW(node, "r");
}

/** Serialized form without the XML declaration — used to compare two rPr. */
function xmlOf(node) {
	return serializeXml(node).replace(/^<\?xml[^>]*\?>\s*/, "");
}

//#region text semantics

function rawText(elem) {
	let out = "";
	for (let c = elem.firstChild; c; c = c.nextSibling) {
		if (c.nodeType === TEXT_NODE || c.nodeType === CDATA_NODE) out += c.data;
	}
	return out;
}

function hasPreserve(elem) {
	return elem.getAttribute("xml:space") === "preserve";
}

/** The text this element actually contributes to the page. */
function renderedText(elem) {
	const text = rawText(elem);
	return hasPreserve(elem) ? text : stripXmlSpace(text);
}

//#endregion

//#region merging

function nextElementSibling(node) {
	for (let s = node.nextSibling; s; s = s.nextSibling) {
		if (s.nodeType === ELEMENT_NODE) return s;
	}
	return null;
}

/**
 * Next run sibling, skipping over any non-run elements in between. A
 * bookmarkStart or comment anchor between two runs does not hide the second
 * run from the merge walk — though `canMerge` refuses to join runs that are
 * not directly adjacent anyway.
 */
function nextSiblingRun(node) {
	for (let s = node.nextSibling; s; s = s.nextSibling) {
		if (isRun(s)) return s;
	}
	return null;
}

function firstChildRun(container) {
	for (let c = container.firstChild; c; c = c.nextSibling) {
		if (isRun(c)) return c;
	}
	return null;
}

/** Only nothing-but-whitespace text and comments may sit between the two. */
function isAdjacent(a, b) {
	for (let n = a.nextSibling; n; n = n.nextSibling) {
		if (n === b) return true;
		if (n.nodeType === ELEMENT_NODE) return false;
		if (n.nodeType === TEXT_NODE && stripXmlSpace(n.data) !== "") return false;
	}
	return false;
}

function canMerge(target, source) {
	const rpr1 = firstChildW(target, "rPr");
	const rpr2 = firstChildW(source, "rPr");
	if ((rpr1 === null) !== (rpr2 === null)) return false;
	if (rpr1 === null) return true;
	return xmlOf(rpr1) === xmlOf(rpr2);
}

/** Move the source run's content (everything but its rPr) onto the target. */
function mergeRunContent(target, source) {
	for (let c = source.firstChild; c; c = c.nextSibling) {
		if (c.nodeType !== ELEMENT_NODE) continue;
		if (c.namespaceURI === NS.w && c.localName === "rPr") continue;
		target.appendChild(c);
	}
}

function consolidateTextElements(run, tag) {
	const nodes = childrenW(run, tag);
	for (let i = nodes.length - 1; i > 0; i--) {
		const curr = nodes[i];
		const prev = nodes[i - 1];
		if (!isAdjacent(prev, curr)) continue;

		const merged = renderedText(prev) + renderedText(curr);
		const hadPreserve = hasPreserve(prev) || hasPreserve(curr);
		const doc = run.ownerDocument;

		for (const node of Array.from(prev.childNodes)) {
			if (node.nodeType === TEXT_NODE || node.nodeType === CDATA_NODE) prev.removeChild(node);
			else run.insertBefore(node, curr);
		}
		prev.appendChild(doc.createTextNode(merged));
		for (const node of Array.from(curr.childNodes)) {
			if (node.nodeType !== TEXT_NODE && node.nodeType !== CDATA_NODE) run.insertBefore(node, curr);
		}

		if (merged !== stripXmlSpace(merged) || hadPreserve) prev.setAttribute("xml:space", "preserve");
		else if (prev.hasAttribute("xml:space")) prev.removeAttribute("xml:space");

		run.removeChild(curr);
	}
}

function consolidateText(run) {
	consolidateTextElements(run, "t");
	consolidateTextElements(run, "delText");
}

function mergeRunsIn(container) {
	let count = 0;
	let run = firstChildRun(container);
	while (run) {
		for (;;) {
			const next = nextElementSibling(run);
			if (next && isRun(next) && canMerge(run, next)) {
				mergeRunContent(run, next);
				container.removeChild(next);
				count += 1;
			} else break;
		}
		consolidateText(run);
		run = nextSiblingRun(run);
	}
	return count;
}

//#endregion

//#region parts

const HEADER_RE = /^word\/header\d*\.xml$/;
const FOOTER_RE = /^word\/footer\d*\.xml$/;
const NOTE_PARTS = new Set(["word/footnotes.xml", "word/endnotes.xml"]);

/**
 * Which archive entries to merge. `word/document.xml` is always included (it
 * is the point of the tool); a spec widens the set. Unknown names are ignored
 * rather than rejected, so `all` stays forward-compatible.
 */
function selectMergeParts(entries, spec) {
	const all = spec === "all";
	const wanted = new Set(
		(spec ?? "")
			.split(",")
			.map((s) => s.trim().toLowerCase())
			.filter(Boolean)
	);
	const parts = [];
	for (const entry of entries) {
		if (entry === "word/document.xml") {
			parts.push(entry);
			continue;
		}
		if (NOTE_PARTS.has(entry)) {
			if (all || wanted.has("notes") || wanted.has("footnotes") || wanted.has("endnotes")) parts.push(entry);
			continue;
		}
		if (HEADER_RE.test(entry)) {
			if (all || wanted.has("headers") || wanted.has("header")) parts.push(entry);
			continue;
		}
		if (FOOTER_RE.test(entry)) {
			if (all || wanted.has("footers") || wanted.has("footer")) parts.push(entry);
		}
	}
	return parts;
}

//#endregion

/**
 * Merge runs in place inside one parsed part. Returns the number merged.
 * `doc` must be a WordprocessingML document node (w:document).
 */
function mergeRunsInDocument(doc) {
	const root = doc.documentElement;
	if (!root) return 0;

	for (const marker of collectW(root, "proofErr")) marker.parentNode?.removeChild(marker);

	const runs = collectW(root, "r");
	for (const run of runs) {
		for (const attr of Array.from(run.attributes ?? [])) {
			if (attr.name.toLowerCase().includes("rsid")) run.removeAttribute(attr.name);
		}
	}

	const containers = new Set();
	for (const run of runs) if (run.parentNode) containers.add(run.parentNode);

	let count = 0;
	for (const container of containers) count += mergeRunsIn(container);
	return count;
}

//#endregion
export { mergeRunsInDocument, selectMergeParts };
