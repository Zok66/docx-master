import { c as paragraphRuns, h as NS, o as getChildren } from "./xml-utils.js";

//#region lib/edit/fields/field-parse.ts
/**
* Parse a run sequence containing complex fields into structured form.
*
* Walks `<w:r>` elements; consecutive runs forming a complex field
* (`fldChar begin` → `fldChar end`) fold into a single `field` entry
* with `instrText`, `result`, `fieldType`, and parsed `details`. Plain
* text runs pass through as `text` entries.
*
* Consumed by:
*   - `inspect-range` / `inspect-runs` — surface field structure in agent
*     output instead of raw fldChar XML
*   - `migrate-captions` — distinguish SEQ-based captions (skip) from
*     manually-numbered candidates (transform)
*   - `inspect-caption` — aggregate occurrences per identifier
*   - standardize re-emit — detect existing caption paragraphs to
*     rebuild surrounding runs
*
* Unrecognized field types appear as `fieldType: "OTHER"` with raw
* instrText preserved. The parser tokenizes respecting double-quoted
* substrings (required for STYLEREF style names containing spaces).
*
* instrText spanning multiple `<w:instrText>` elements (long field codes
* Word may split across runs) is concatenated in order.
*/
const w = NS.w;
/** Parse an array of `<w:r>` elements (or any container's run children)
* into structured form. Returns entries in document order. */
function parseFieldRuns(runs) {
	const out = [];
	let i = 0;
	while (i < runs.length) {
		const run = runs[i];
		if (firstFldCharType(run) === "begin") {
			const group = [run];
			i++;
			while (i < runs.length) {
				const r = runs[i];
				group.push(r);
				i++;
				if (firstFldCharType(r) === "end") break;
			}
			out.push(parseFieldGroup(group));
		} else {
			out.push({
				kind: "text",
				text: textOf(run),
				rPrEl: rPrOf(run)
			});
			i++;
		}
	}
	return out;
}
/** SEQ field details in a paragraph, in document order. When
* `skipRepeat` is true, drops `\c` (repeat) entries — engine-injected
* chapter prefixes carry `\c` to read the current counter value without
* advancing it, so they neither own their identifier nor shadow the
* caption identifier that follows them in the same paragraph. Callers
* needing the "primary" identifier take `[0].identifier`; callers
* needing all advancing counters iterate the array. */
function seqFields(paragraph, opts) {
	const parsed = parseFieldRuns(paragraphRuns(paragraph));
	const out = [];
	for (const entry of parsed) {
		if (entry.kind !== "field" || entry.fieldType !== "SEQ") continue;
		if (opts.skipRepeat && entry.details.repeat) continue;
		out.push(entry.details);
	}
	return out;
}
function firstFldCharType(run) {
	for (const child of getChildren(run)) if (child.namespaceURI === w && child.localName === "fldChar") {
		const t = child.getAttributeNS(w, "fldCharType");
		if (t === "begin" || t === "separate" || t === "end") return t;
	}
}
function rPrOf(run) {
	for (const child of getChildren(run)) if (child.namespaceURI === w && child.localName === "rPr") return child;
}
function textOf(run) {
	let out = "";
	for (const child of getChildren(run)) if (child.namespaceURI === w && (child.localName === "t" || child.localName === "instrText")) out += child.textContent ?? "";
	return out;
}
function parseFieldGroup(runs) {
	let instrText = "";
	let result = "";
	let sawSeparate = false;
	let resultRPr;
	for (const r of runs) {
		const fc = firstFldCharType(r);
		if (fc === "separate") {
			sawSeparate = true;
			continue;
		}
		if (fc) continue;
		for (const child of getChildren(r)) {
			if (child.namespaceURI !== w) continue;
			if (child.localName === "instrText" && !sawSeparate) instrText += child.textContent ?? "";
			else if (child.localName === "t" && sawSeparate) {
				result += child.textContent ?? "";
				if (!resultRPr) resultRPr = rPrOf(r);
			}
		}
	}
	const trimmed = instrText.trim();
	const { fieldType, details } = parseInstrText(trimmed);
	return {
		kind: "field",
		fieldType,
		instrText: trimmed,
		result,
		rPrEl: resultRPr,
		details
	};
}
function parseInstrText(s) {
	const tokens = tokenizeInstrText(s);
	if (tokens.length === 0) return {
		fieldType: "OTHER",
		details: {}
	};
	const type = tokens[0].toUpperCase();
	const rest = tokens.slice(1);
	if (type === "REF") return {
		fieldType: "REF",
		details: parseRefDetails(rest)
	};
	if (type === "SEQ") return {
		fieldType: "SEQ",
		details: parseSeqDetails(rest)
	};
	if (type === "STYLEREF") return {
		fieldType: "STYLEREF",
		details: parseStyleRefDetails(rest)
	};
	return {
		fieldType: "OTHER",
		details: {}
	};
}
/** Tokenize respecting double-quoted substrings. Used for STYLEREF
* `"<name>"` where the name can contain spaces. Backslash-prefixed
* switches are single tokens (no escape processing inside switches). */
function tokenizeInstrText(s) {
	const tokens = [];
	let i = 0;
	while (i < s.length) {
		while (i < s.length && /\s/.test(s[i])) i++;
		if (i >= s.length) break;
		if (s[i] === "\"") {
			i++;
			let token = "";
			while (i < s.length && s[i] !== "\"") {
				token += s[i];
				i++;
			}
			if (i < s.length) i++;
			tokens.push(token);
		} else {
			let token = "";
			while (i < s.length && !/\s/.test(s[i])) {
				token += s[i];
				i++;
			}
			tokens.push(token);
		}
	}
	return tokens;
}
function parseRefDetails(tokens) {
	const switches = [];
	let bookmarkName;
	for (const t of tokens) if (t.startsWith("\\")) switches.push(t);
	else if (bookmarkName === void 0) bookmarkName = t;
	return {
		bookmarkName,
		switches
	};
}
const SWITCH_TO_FORMAT = {
	ARABIC: "arabic",
	alphabetic: "alphabetic",
	ALPHABETIC: "ALPHABETIC",
	roman: "roman",
	ROMAN: "ROMAN",
	CHINESENUM2: "chinese",
	CHINESENUM3: "chinese-formal"
};
function parseSeqDetails(tokens) {
	if (tokens.length === 0) return {};
	const details = { identifier: tokens[0] };
	let i = 1;
	while (i < tokens.length) {
		const tok = tokens[i];
		if (tok === "\\*" && i + 1 < tokens.length) {
			const fmt = SWITCH_TO_FORMAT[tokens[i + 1]];
			if (fmt) details.format = fmt;
			i += 2;
		} else if (tok === "\\s" && i + 1 < tokens.length) {
			const n = parseInt(tokens[i + 1], 10);
			if (Number.isFinite(n)) details.restartAtOutlineLevel = n;
			i += 2;
		} else if (tok === "\\r" && i + 1 < tokens.length) {
			const n = parseInt(tokens[i + 1], 10);
			if (Number.isFinite(n)) details.resetTo = n;
			i += 2;
		} else if (tok === "\\c") {
			details.repeat = true;
			i++;
		} else if (tok === "\\h") {
			details.hidden = true;
			i++;
		} else i++;
	}
	return details;
}
function parseStyleRefDetails(tokens) {
	if (tokens.length === 0) return {};
	const switches = [];
	for (const t of tokens.slice(1)) if (t.startsWith("\\")) switches.push(t);
	return {
		styleName: tokens[0],
		switches
	};
}

//#endregion
export { seqFields as n, parseFieldRuns as t };