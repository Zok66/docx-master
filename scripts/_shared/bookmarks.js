/**
 * Bookmark & cross-reference analysis — pure string work, zero dependencies.
 *
 * WHY THIS EXISTS
 * A bookmark is the one document object with no visible rendering. Lose one and
 * the file still opens, still validates against the XSD, still passes every
 * schema check — only Word quietly shows "Error! Bookmark not defined." in the
 * middle of a sentence, or a TOC entry stops jumping. Nothing in the OOXML
 * toolchain looks at whether a name in a REF / PAGEREF / `#anchor` actually
 * exists somewhere in the package.
 *
 * WHAT IT DOES
 * Walks the XML text of every part (body, headers, footers, footnotes, endnotes)
 * as a token stream, and reports:
 *   - every `w:bookmarkStart` (name, id, part, paragraph ordinal, origin)
 *   - every reference to a bookmark: REF / PAGEREF / NOTEREF fields,
 *     `HYPERLINK \l "name"`, `TOC \b name`, and structural `w:hyperlink w:anchor`
 *   - problems: dangling references, unclosed / unopened bookmarks, duplicate
 *     ids, duplicate names, empty names, references with no target
 *
 * The field scanner is a real `fldChar` state machine with a stack, because
 * Word nests fields: an updated TOC contains one PAGEREF field per entry,
 * *inside* the TOC field. A flat regex would attribute every inner instruction
 * to the outer field and misreport the references.
 *
 * Names are resolved across the whole package, not per part — a header REF can
 * legitimately point at a body bookmark.
 */

const REF_FIELDS = new Set(["REF", "PAGEREF", "NOTEREF"]);

/** Word's own bookmark families. Anything else was written by a tool or a human. */
const ORIGIN_RULES = [
	[/^_Toc\d+$/, "word-toc"],
	[/^_Ref\d+$/, "word-xref"],
	[/^_GoBack$/, "word-cursor"],
	[/^_Hlt/, "word-highlight"],
	[/^_OLE_LINK/, "word-ole"],
	[/^_Pic\d+$/, "word-picture"],
	[/^_Sio/, "word-signature"],
	[/^_Mail[A-Z]/, "word-envelope"]
];

export const ORIGIN_LABEL = {
	"word-toc": "Word TOC entry",
	"word-xref": "Word cross-reference",
	"word-cursor": "Word cursor position",
	"word-highlight": "Word highlight",
	"word-ole": "Word OLE link",
	"word-picture": "Word picture",
	"word-signature": "Word signature",
	"word-envelope": "Word envelope",
	custom: "custom / tool-written"
};

export function originOf(name) {
	for (const [re, origin] of ORIGIN_RULES) if (re.test(name)) return origin;
	return "custom";
}

/** True for names Word allocates and rewrites on its own. A missing one of these
 * is usually benign (Word dropped an entry it no longer needs); a missing
 * `custom` name is a defect, because a tool wrote a reference to it and nothing
 * will ever recreate it. */
export function isWordGenerated(name) {
	return originOf(name) !== "custom";
}

// One tokenizer pass over a part. Named groups keep the branch readable.
const TOKEN_RE = new RegExp(
	[
		"<w:bookmarkStart\\b(?<bsAttr>[^>]*?)/?>",
		"<w:bookmarkEnd\\b(?<beAttr>[^>]*?)/?>",
		"<w:fldChar\\b(?<fcAttr>[^>]*?)/?>",
		"<w:instrText\\b[^>]*>(?<instr>[\\s\\S]*?)</w:instrText>",
		"<w:hyperlink\\b(?<hlAttr>[^>]*?)>",
		"<w:fldSimple\\b(?<fsAttr>[^>]*?)>"
	].join("|"),
	"g"
);

const attr = (attrs, name) => {
	const m = String(attrs || "").match(new RegExp(`\\b${name}="([^"]*)"`));
	return m ? m[1] : "";
};

const stripQuotes = (s) => String(s || "").replace(/^"|"$/g, "").trim();

/** First bare (non-switch) token: `_Toc123 \h \* MERGEFORMAT` → `_Toc123`. */
function firstBareToken(s) {
	for (const piece of String(s || "").split(/\s+/)) {
		if (!piece || piece.startsWith("\\")) continue;
		return stripQuotes(piece);
	}
	return "";
}

/**
 * Instructions that name a bookmark. Returns [] when the field references no
 * bookmark (STYLEREF reads a style; an external HYPERLINK has a rel id, not an
 * anchor).
 */
function refsFromInstruction(instr) {
	const flat = String(instr || "").replace(/[\r\n\t]+/g, " ").replace(/\s+/g, " ").trim();
	if (!flat) return [];
	const m = flat.match(/^([A-Za-z]+)\b\s*([\s\S]*)$/);
	if (!m) return [];
	const kind = m[1].toUpperCase();
	const rest = m[2] || "";
	if (REF_FIELDS.has(kind)) return [{ kind, target: firstBareToken(rest), raw: flat }];
	if (kind === "STYLEREF") return [];
	if (kind === "HYPERLINK") {
		const l = rest.match(/\\l\s+("[^"]*"|\S+)/);
		return l ? [{ kind: "HYPERLINK \\l", target: stripQuotes(l[1]), raw: flat }] : [];
	}
	if (kind === "TOC" || kind === "TOA") {
		const b = rest.match(/\\b\s+("[^"]*"|\S+)/);
		return b ? [{ kind: `${kind} \\b`, target: stripQuotes(b[1]), raw: flat }] : [];
	}
	return [];
}

/**
 * Analyze one part's XML text.
 *
 * @param {string} part  part name, e.g. `word/document.xml`
 * @param {string} xml   its text
 * @returns {{bookmarks: object[], refs: object[], problems: object[]}}
 *   `refs[].resolved` is left undefined here — cross-part resolution happens in
 *   `analyzeParts`, which sees the whole package.
 */
export function analyzePart(part, xml) {
	const bookmarks = [];
	const refs = [];
	const problems = [];
	if (!xml) return { bookmarks, refs, problems };

	/** Field stack. Each level collects its own instruction while the level is
	 * still in its instruction phase (`collecting`), so nested fields don't bleed
	 * into their parent. */
	const stack = [];
	/** Currently-open bookmarks, keyed by `w:id` — the pairing key OOXML uses. */
	const openIds = new Map();
	/** Opens whose id was later reused by another `bookmarkStart`. The second
	 * start steals the id, so the first one can now never be closed; keep it
	 * here instead of dropping it, or an unclosed bookmark goes unreported. */
	const orphanedOpens = [];
	const seenNames = new Map();
	let para = 0;
	let cursor = 0;

	const noteRef = (r, paraOrdinal) => refs.push({ ...r, part, para: paraOrdinal });
	const noteField = (instr, paraOrdinal) => {
		for (const r of refsFromInstruction(instr)) noteRef(r, paraOrdinal);
	};

	for (const m of xml.matchAll(TOKEN_RE)) {
		// Paragraph ordinal: count `<w:p` / `<w:p>` openings before this token.
		para += (xml.slice(cursor, m.index).match(/<w:p[ >]/g) || []).length;
		cursor = m.index + m[0].length;
		const g = m.groups || {};

		if (g.bsAttr !== undefined) {
			const name = attr(g.bsAttr, "w:name");
			const id = attr(g.bsAttr, "w:id");
			const entry = { name, id, part, para };
			if (!name) problems.push({ kind: "empty-name", part, para, detail: `bookmarkStart with w:id="${id}" carries no w:name` });
			else if (seenNames.has(name)) problems.push({ kind: "duplicate-name", part, para, name, detail: `bookmark "${name}" declared twice (first at paragraph ${seenNames.get(name)})` });
			else seenNames.set(name, para);
			if (id && openIds.has(id)) {
				problems.push({ kind: "duplicate-id", part, para, name, detail: `bookmark id ${id} already used by "${openIds.get(id).name}"` });
				orphanedOpens.push(openIds.get(id));
				openIds.delete(id);
			}
			if (id) openIds.set(id, entry);
			bookmarks.push(entry);
			continue;
		}

		if (g.beAttr !== undefined) {
			const id = attr(g.beAttr, "w:id");
			if (id && !openIds.has(id)) problems.push({ kind: "unopened", part, para, detail: `bookmarkEnd with w:id="${id}" has no matching bookmarkStart` });
			else if (id) openIds.delete(id);
			continue;
		}

		if (g.fcAttr !== undefined) {
			const type = attr(g.fcAttr, "w:fldCharType");
			if (type === "begin") stack.push({ instr: "", collecting: true, para });
			else if (type === "separate" && stack.length) stack[stack.length - 1].collecting = false;
			else if (type === "end") {
				const top = stack.pop();
				if (top) noteField(top.instr, top.para);
			}
			continue;
		}

		if (g.instr !== undefined) {
			if (stack.length && stack[stack.length - 1].collecting) stack[stack.length - 1].instr += g.instr;
			continue;
		}

		if (g.fsAttr !== undefined) {
			// `w:fldSimple` carries its whole instruction as an attribute.
			const instr = attr(g.fsAttr, "w:instr").replace(/&quot;/g, '"');
			noteField(instr, para);
			continue;
		}

		if (g.hlAttr !== undefined) {
			const anchor = attr(g.hlAttr, "w:anchor");
			if (anchor && anchor !== "_top") noteRef({ kind: "w:hyperlink", target: anchor, raw: `w:anchor="${anchor}"` }, para);
			continue;
		}
	}

	// Anything still open when the part ends was never closed in this part. Word
	// never does this; a hand-written fragment does, and the reference to it then
	// resolves to nothing.
	for (const entry of [...openIds.values(), ...orphanedOpens]) {
		problems.push({ kind: "unclosed", part, para: entry.para, name: entry.name, detail: `bookmark "${entry.name}" (w:id=${entry.id}) is never closed` });
	}
	for (const level of stack) {
		if (level.instr.trim()) problems.push({ kind: "unclosed-field", part, para: level.para, detail: `field "${level.instr.trim().slice(0, 40)}" has no end` });
	}

	return { bookmarks, refs, problems };
}

/**
 * Analyze a whole package. `parts` is `[{ name, xml }]` — pass every part that
 * can carry bookmarks or references (document, headers, footers, footnotes,
 * endnotes; comments and textboxes live in document.xml).
 */
export function analyzeParts(parts) {
	const all = { bookmarks: [], refs: [], problems: [], byPart: {} };
	for (const p of parts || []) {
		if (!p || !p.xml) continue;
		const r = analyzePart(p.name, p.xml);
		all.bookmarks.push(...r.bookmarks);
		all.refs.push(...r.refs);
		all.problems.push(...r.problems);
		all.byPart[p.name] = { bookmarks: r.bookmarks.length, refs: r.refs.length };
	}

	const names = new Set(all.bookmarks.map((b) => b.name));
	for (const r of all.refs) {
		r.resolved = names.has(r.target);
		if (!r.target) all.problems.push({ kind: "empty-target", part: r.part, para: r.para, refKind: r.kind, detail: `field "${r.raw}" names no bookmark` });
		else if (!r.resolved) all.problems.push({ kind: "dangling", part: r.part, para: r.para, name: r.target, refKind: r.kind, detail: `${r.kind} points at "${r.target}", which no bookmarkStart defines` });
	}

	all.dangling = all.problems.filter((p) => p.kind === "dangling");
	all.names = names;
	all.countByOrigin = all.bookmarks.reduce((acc, b) => {
		const o = originOf(b.name);
		acc[o] = (acc[o] || 0) + 1;
		return acc;
	}, {});
	all.refKinds = all.refs.reduce((acc, r) => {
		acc[r.kind] = (acc[r.kind] || 0) + 1;
		return acc;
	}, {});
	return all;
}

/** Short one-line-per-fact summary used by both the CLI and the delivery gate. */
export function summarize(scan) {
	const lines = [];
	const origins = Object.entries(scan.countByOrigin || {}).map(([k, n]) => `${ORIGIN_LABEL[k] || k}: ${n}`);
	lines.push(`bookmarks: ${scan.bookmarks.length}${origins.length ? ` (${origins.join(", ")})` : ""}`);
	const kinds = Object.entries(scan.refKinds || {}).map(([k, n]) => `${k}: ${n}`);
	lines.push(`references: ${scan.refs.length}${kinds.length ? ` (${kinds.join(", ")})` : ""}`);
	lines.push(`dangling references: ${(scan.dangling || []).length}`);
	return lines;
}
