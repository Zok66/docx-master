import { a as firstChildNS, f as wAttr, h as NS, o as getChildren, s as getChildrenNS } from "./_shared/xml-utils.js";
import { t as summarizeTable } from "./_shared/table-classifier.js";
import { t as loadDocx } from "./_shared/load.js";
import { s as walkIndexedParagraphs } from "./_shared/locator.js";
import { r as pad } from "./_shared/format.js";

//#region lib/edit/text-search.ts
/**
* Character-level text search across paragraphs. Atomic find primitive: locate
* a literal or regex pattern in the document body, return per-match metadata
* (paragraph index, run index, offset, length, context preview, and any
* structural region the match falls inside). The agent decides what to do
* with the coordinates — replace via `set-run`, deeper inspection via
* `inspect_runs`, coverage validation, or just browsing.
*
* Design choices:
*   - Paragraph is the search unit. Matches don't span paragraphs.
*   - <w:tab/> and <w:br/> become sentinel chars (\t / \n) in the projection;
*     matches that would cross those boundaries are rejected (text flow does
*     not naturally span them).
*   - <w:drawing> / <w:object> / <m:oMath> become ￼ (object replacement
*     character) in the projection — invisible to literal/regex matches that
*     don't deliberately include that codepoint.
*   - Run-index in the output refers to direct `<w:r>` children of `<w:p>` —
*     this matches `set-run`'s runIndex semantics. Matches in nested runs
*     (inside <w:hyperlink> / <w:ins> / <w:del> / <w:sdt>) report no runIndex;
*     the structural region annotation tells the agent where they are.
*   - Field-region detection uses the same fldChar depth state-machine as
*     `lib/blockers.ts`; a match between <w:fldChar begin> and the matching
*     <w:fldChar end> is reported as `field`.
*/
const w = NS.w;
const TAB_SENTINEL = "	";
const BREAK_SENTINEL = "\n";
const OBJECT_PLACEHOLDER = "￼";
function searchDocument(documentDoc, opts) {
	if (!opts.pattern || opts.pattern.length === 0) throw new Error("search pattern is empty");
	const ctxN = opts.contextChars ?? 25;
	const [bL, bR] = opts.bracket ?? ["‹", "›"];
	const limit = opts.limit ?? Infinity;
	const re = opts.regex ? new RegExp(opts.pattern, "g") : null;
	const hits = [];
	const fieldState = { depth: 0 };
	for (const p of walkIndexedParagraphs(documentDoc)) {
		if (opts.paraIndex !== void 0 && p.index !== opts.paraIndex) continue;
		if (opts.paraRange) {
			if (p.index < opts.paraRange.from || p.index > opts.paraRange.to) continue;
		}
		const proj = buildParagraphProjection(p.element, fieldState);
		if (proj.text.length === 0) continue;
		const matches = findInProjection(proj.text, opts.pattern, re);
		for (const m of matches) {
			const hit = buildHit(proj, m, fieldState, ctxN, bL, bR);
			if (!hit) continue;
			hits.push({
				paragraphIndex: p.index,
				...hit
			});
			if (hits.length >= limit) return hits;
		}
	}
	if (opts.paraIndex === void 0 && !opts.paraRange) {
		const root = documentDoc.documentElement;
		const body = root ? firstChildNS(root, NS.w, "body") : null;
		if (body) {
			let tableIdx = 0;
			for (const child of getChildren(body)) {
				if (child.namespaceURI !== NS.w || child.localName !== "tbl") continue;
				tableIdx++;
				if (summarizeTable(child).classification !== "data") continue;
				const rows = getChildrenNS(child, NS.w, "tr");
				for (let ri = 0; ri < rows.length; ri++) {
					const cells = getChildrenNS(rows[ri], NS.w, "tc");
					for (let ci = 0; ci < cells.length; ci++) {
						const cellParas = getChildrenNS(cells[ci], NS.w, "p");
						for (let pi = 0; pi < cellParas.length; pi++) {
							const proj = buildParagraphProjection(cellParas[pi], fieldState);
							if (proj.text.length === 0) continue;
							const matches = findInProjection(proj.text, opts.pattern, re);
							for (const m of matches) {
								const hit = buildHit(proj, m, fieldState, ctxN, bL, bR);
								if (!hit) continue;
								hits.push({
									cell: {
										table: tableIdx,
										row: ri + 1,
										col: ci + 1,
										paragraph: pi + 1
									},
									...hit
								});
								if (hits.length >= limit) return hits;
							}
						}
					}
				}
			}
		}
	}
	return hits;
}
/** Build the common match-hit fields from a projection + raw match position.
* Returns null if the match crosses a sentinel boundary. */
function buildHit(proj, m, _fieldState, ctxN, bL, bR) {
	const slice = proj.text.slice(m.start, m.start + m.length);
	if (slice.includes(TAB_SENTINEL) || slice.includes(BREAK_SENTINEL)) return null;
	const overlapping = proj.segments.filter((seg) => seg.start < m.start + m.length && seg.end > m.start);
	if (overlapping.length === 0) return null;
	let firstRunIdx = null;
	let lastRunIdx = null;
	let region = null;
	for (const seg of overlapping) {
		if (seg.runIndex !== null) {
			if (firstRunIdx === null) firstRunIdx = seg.runIndex;
			lastRunIdx = seg.runIndex;
		}
		if (region === null && seg.region !== null) region = seg.region;
	}
	const crossRun = firstRunIdx !== null && lastRunIdx !== null && firstRunIdx !== lastRunIdx;
	const matched = slice;
	const beforeRaw = proj.text.slice(Math.max(0, m.start - ctxN), m.start);
	const afterRaw = proj.text.slice(m.start + m.length, m.start + m.length + ctxN);
	const before = beforeRaw.replace(/[\t\n￼]/g, " ");
	const after = afterRaw.replace(/[\t\n￼]/g, " ");
	const context = `${m.start > ctxN ? "..." : ""}${before}${bL}${matched}${bR}${after}${m.start + m.length + ctxN < proj.text.length ? "..." : ""}`;
	return {
		ch: m.start,
		len: m.length,
		matched,
		runIndex: firstRunIdx,
		runIndexEnd: lastRunIdx,
		crossRun,
		region,
		context
	};
}
function buildParagraphProjection(pEl, fieldState) {
	const segments = [];
	let text = "";
	let directRunIdx = 1;
	function appendNode(node, scopeRegion) {
		if (node.namespaceURI !== w) return;
		const local = node.localName;
		let effectiveRegion = scopeRegion;
		if (local === "ins" || local === "del") effectiveRegion ??= "tracked-change";
		else if (local === "sdt") effectiveRegion ??= "sdt";
		else if (local === "hyperlink") effectiveRegion ??= "hyperlink";
		if (local === "r") {
			const isDirect = node.parentNode === pEl;
			const idx = isDirect ? directRunIdx : null;
			if (isDirect) directRunIdx++;
			appendRun(node, idx, effectiveRegion);
			return;
		}
		if (local === "hyperlink" || local === "ins" || local === "del" || local === "sdt") {
			const recurseRoot = local === "sdt" ? firstChildNS(node, w, "sdtContent") ?? node : node;
			for (const child of getChildren(recurseRoot)) appendNode(child, effectiveRegion);
			return;
		}
		if (local === "pPr") return;
	}
	function appendRun(rEl, runIndex, region) {
		for (const c of getChildren(rEl)) {
			if (c.namespaceURI !== w) {
				addSegment(OBJECT_PLACEHOLDER, runIndex, region);
				continue;
			}
			const ln = c.localName;
			if (ln === "t") {
				const s = c.textContent ?? "";
				if (s) addSegment(s, runIndex, region);
			} else if (ln === "tab") addSegment(TAB_SENTINEL, runIndex, region);
			else if (ln === "br" || ln === "cr") addSegment(BREAK_SENTINEL, runIndex, region);
			else if (ln === "fldChar") {
				const ftype = wAttr(c, "fldCharType");
				if (ftype === "begin") fieldState.depth++;
				else if (ftype === "end") {
					if (fieldState.depth > 0) fieldState.depth--;
				}
			} else if (ln === "instrText") {} else if (ln === "drawing" || ln === "object" || ln === "pict") addSegment(OBJECT_PLACEHOLDER, runIndex, region);
			else if (ln === "sym") addSegment(OBJECT_PLACEHOLDER, runIndex, region);
		}
	}
	function addSegment(s, runIndex, region) {
		const start = text.length;
		text += s;
		const end = text.length;
		const finalRegion = fieldState.depth > 0 ? "field" : region;
		segments.push({
			start,
			end,
			runIndex,
			region: finalRegion
		});
	}
	for (const node of getChildren(pEl)) {
		if (node.namespaceURI !== w) continue;
		appendNode(node, null);
	}
	return {
		text,
		segments
	};
}
function findInProjection(text, pattern, re) {
	const out = [];
	if (re) for (const m of text.matchAll(re)) {
		if (m[0].length === 0) continue;
		out.push({
			start: m.index ?? 0,
			length: m[0].length
		});
	}
	else {
		let from = 0;
		while (true) {
			const idx = text.indexOf(pattern, from);
			if (idx < 0) break;
			out.push({
				start: idx,
				length: pattern.length
			});
			from = idx + pattern.length;
		}
	}
	return out;
}
function describeRegion(region) {
	switch (region) {
		case "tracked-change": return "in tracked change";
		case "field": return "in field result";
		case "sdt": return "in content control";
		case "hyperlink": return "in hyperlink";
	}
}

//#endregion
//#region skill/tools/find-text.ts
/**
* find_text <docx> <pattern> [--regex] [--paragraph N | --range A-B] [--limit N] [--context N]
*
* Character-level text search returning per-match coordinates: paragraph
* index, run index (when match is in a direct-child run), char offset,
* length, and a context preview with the match bracketed.
*
* Read-only primitive: locate text by literal or regex; what the agent
* does next (replace via set-run, deeper inspection, coverage validation,
* just browsing) is up to them.
*/
/** Unique location key for grouping hits by paragraph/cell in the summary. */
function locationKey(h) {
	if (h.cell) return `T${h.cell.table}R${h.cell.row}C${h.cell.col}K${h.cell.paragraph}`;
	return `#${h.paragraphIndex}`;
}
async function main() {
	const argv = process.argv.slice(2);
	let isRegex = false;
	let paragraph;
	let range;
	let limit = 50;
	let contextChars = 25;
	const positional = [];
	for (let i = 0; i < argv.length; i++) {
		const a = argv[i];
		if (a === "--regex") isRegex = true;
		else if (a === "--paragraph") {
			const v = argv[++i];
			if (!v || isNaN(parseInt(v, 10))) {
				console.error("--paragraph requires a positive integer");
				process.exit(1);
			}
			paragraph = parseInt(v, 10);
		} else if (a === "--range") {
			const m = argv[++i]?.match(/^(\d+)-(\d+)$/);
			if (!m) {
				console.error("--range requires the form \"A-B\" (e.g. \"30-50\")");
				process.exit(1);
			}
			range = {
				from: parseInt(m[1], 10),
				to: parseInt(m[2], 10)
			};
		} else if (a === "--limit") {
			const v = argv[++i];
			if (!v || isNaN(parseInt(v, 10))) {
				console.error("--limit requires a non-negative integer (0 = no cap)");
				process.exit(1);
			}
			limit = parseInt(v, 10);
		} else if (a === "--context") {
			const v = argv[++i];
			if (!v || isNaN(parseInt(v, 10))) {
				console.error("--context requires a non-negative integer");
				process.exit(1);
			}
			contextChars = parseInt(v, 10);
		} else if (a.startsWith("--")) {
			console.error(`unknown flag: ${a}`);
			process.exit(1);
		} else positional.push(a);
	}
	const file = positional[0];
	const pattern = positional[1];
	if (!file || pattern === void 0) {
		console.error("Usage: node scripts/find_text.js <docx-path> <pattern> [--regex] [--paragraph N | --range A-B] [--limit N] [--context N]");
		process.exit(1);
	}
	if (isRegex) try {
		new RegExp(pattern);
	} catch (err) {
		console.error(`Invalid regex /${pattern}/: ${err.message}`);
		process.exit(1);
	}
	try {
		const hits = searchDocument((await loadDocx(file)).documentDoc, {
			pattern,
			regex: isRegex,
			paraIndex: paragraph,
			paraRange: range,
			limit: limit === 0 ? Infinity : limit,
			contextChars
		});
		const lines = [];
		const scopeStr = paragraph !== void 0 ? `paragraph #${paragraph}` : range ? `paragraphs #${range.from}–#${range.to}` : "whole body";
		lines.push(`Pattern: ${isRegex ? `/${pattern}/` : `literal "${pattern}"`}`);
		lines.push(`Scope:   ${scopeStr}`);
		lines.push("");
		if (hits.length === 0) {
			lines.push("No matches.");
			console.log(lines.join("\n"));
			return;
		}
		const locSet = new Set(hits.map(locationKey));
		const annotation = formatAnnotations(summarize(hits));
		lines.push(`${hits.length} matches across ${locSet.size} paragraphs${annotation ? ` (${annotation})` : ""}:`);
		lines.push("");
		for (const h of hits) lines.push(formatMatchLine(h));
		if (limit !== 0 && hits.length === limit) {
			lines.push("");
			lines.push(`(hit limit ${limit}; pass --limit 0 to see all matches)`);
		}
		console.log(lines.join("\n"));
	} catch (err) {
		console.error(`Error: ${err.message}`);
		process.exit(1);
	}
}
function formatMatchLine(h) {
	const locCol = formatLocColumn(h);
	const runCol = formatRunColumn(h);
	const ch = `ch=${String(h.ch).padStart(3, " ")}`;
	const len = `len=${h.len}`;
	const tags = [];
	if (h.crossRun) tags.push("cross-run");
	if (h.region) tags.push(describeRegion(h.region));
	const tagStr = tags.length > 0 ? `   (${tags.join("; ")})` : "";
	return `  ${locCol}  ${runCol}  ${ch}  ${len}   ${h.context}${tagStr}`;
}
function formatLocColumn(h) {
	const PAD = 14;
	if (h.cell) {
		const { table, row, col, paragraph } = h.cell;
		return `T${table}R${row}C${col} K${paragraph}`.padEnd(PAD, " ");
	}
	return `#${pad(h.paragraphIndex)}`.padEnd(PAD, " ");
}
function formatRunColumn(h) {
	const PAD = 12;
	if (h.runIndex === null) return " ".repeat(PAD);
	return (h.crossRun && h.runIndexEnd !== null ? `run[${h.runIndex}..${h.runIndexEnd}]` : `run[${h.runIndex}]`).padEnd(PAD, " ");
}
function summarize(hits) {
	const out = {};
	for (const h of hits) {
		if (h.crossRun) out["cross-run"] = (out["cross-run"] ?? 0) + 1;
		if (h.region) {
			const k = describeRegion(h.region);
			out[k] = (out[k] ?? 0) + 1;
		}
	}
	return out;
}
function formatAnnotations(counts) {
	const keys = Object.keys(counts);
	if (keys.length === 0) return "";
	return keys.map((k) => `${counts[k]} ${k}`).join("; ");
}
main();

//#endregion
export {  };