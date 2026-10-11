import { a as firstChildNS, d as textContent, f as wAttr, h as NS, o as getChildren, s as getChildrenNS } from "./_shared/xml-utils.js";
import { t as summarizeTable } from "./_shared/table-classifier.js";
import { t as loadDocx } from "./_shared/load.js";
import { r as pad } from "./_shared/format.js";

//#region skill/tools/inspect-runs.ts
/**
* inspect_runs <docx> <paraIndex>
* inspect_runs <docx> --table T --row R --col C --paragraph K
*
* Dumps every run inside a paragraph: text, rPr, and a "run-level diversity"
* summary that pinpoints which character properties differ across runs
* (intentional inline emphasis) vs. which are uniform (style-controllable).
*
* This tool exists because `inspect_range` only shows the paragraph's
* computed/dominant rPr — it cannot reveal mixed-format paragraphs like
*   <w:r><w:b/>幻觉与安全性</w:r><w:r>与传统软件不同…</w:r>
* where the leading bold phrase is a separate run from the non-bold body.
* For that, you need this tool.
*
* Cell form: --table / --row / --col / --paragraph address a paragraph inside
* a data-table cell (coordinates match find_text cell-coord hit format).
*/
async function main() {
	const argv = process.argv.slice(2);
	if (!argv[0]) {
		console.error("Usage: node scripts/inspect_runs.js <docx-path> <paragraph-index>\n       node scripts/inspect_runs.js <docx-path> --table T --row R --col C --paragraph K");
		process.exit(1);
	}
	const file = argv[0];
	let cellTarget;
	let paragraphIdx;
	let i = 1;
	while (i < argv.length) {
		const a = argv[i];
		if (a === "--table" || a === "--row" || a === "--col" || a === "--paragraph") {
			const v = argv[++i];
			const n = v !== void 0 ? parseInt(v, 10) : NaN;
			if (isNaN(n) || n < 1) {
				console.error(`${a} requires a positive integer`);
				process.exit(1);
			}
			if (!cellTarget) cellTarget = {
				table: 0,
				row: 0,
				col: 0,
				paragraph: 0
			};
			if (a === "--table") cellTarget.table = n;
			else if (a === "--row") cellTarget.row = n;
			else if (a === "--col") cellTarget.col = n;
			else if (a === "--paragraph") cellTarget.paragraph = n;
		} else if (!a.startsWith("--")) {
			const n = parseInt(a, 10);
			if (isNaN(n) || n < 1) {
				console.error("Invalid paragraph index (must be 1-based positive integer)");
				process.exit(1);
			}
			paragraphIdx = n;
		} else {
			console.error(`unknown flag: ${a}`);
			process.exit(1);
		}
		i++;
	}
	try {
		const doc = await loadDocx(file);
		if (cellTarget !== void 0) {
			const missing = [];
			if (!cellTarget.table) missing.push("--table");
			if (!cellTarget.row) missing.push("--row");
			if (!cellTarget.col) missing.push("--col");
			if (!cellTarget.paragraph) missing.push("--paragraph");
			if (missing.length > 0) {
				console.error(`Cell form requires all four flags. Missing: ${missing.join(", ")}`);
				process.exit(1);
			}
			const pEl = findCellParagraphElement(doc.documentDoc, cellTarget);
			if (!pEl) {
				console.error(`Cell T${cellTarget.table}R${cellTarget.row}C${cellTarget.col} K${cellTarget.paragraph} not found. Verify coordinates via find_text output.`);
				process.exit(1);
			}
			const runs = extractRuns(pEl);
			const locLabel = `T${cellTarget.table}R${cellTarget.row}C${cellTarget.col} K${cellTarget.paragraph}`;
			console.log(renderCellReport(locLabel, runs));
		} else {
			if (paragraphIdx === void 0) {
				console.error("Usage: node scripts/inspect_runs.js <docx-path> <paragraph-index>\n       node scripts/inspect_runs.js <docx-path> --table T --row R --col C --paragraph K");
				process.exit(1);
			}
			const targetIdx = paragraphIdx;
			const para = doc.paragraphs.find((p) => p.index === targetIdx);
			if (!para) {
				const max = doc.paragraphs.length;
				const closest = doc.paragraphs.reduce((best, p) => Math.abs(p.index - targetIdx) < Math.abs(best - targetIdx) ? p.index : best, doc.paragraphs[0]?.index ?? 0);
				console.error(`Paragraph #${targetIdx} not found. Document has ${max} indexed paragraphs (1..${doc.paragraphs[max - 1]?.index ?? 0}). Closest: #${closest}.`);
				console.error("Note: paragraphs inside data tables are not indexed; use --table/--row/--col/--paragraph.");
				process.exit(1);
			}
			const pEl = findParagraphElement(doc.documentDoc, targetIdx);
			if (!pEl) {
				console.error(`Internal error: paragraph #${targetIdx} index lookup failed`);
				process.exit(1);
			}
			const runs = extractRuns(pEl);
			console.log(renderReport(targetIdx, para.fingerprint, para.text, para.styleId, runs));
		}
	} catch (err) {
		console.error(`Error: ${err.message}`);
		process.exit(1);
	}
}
/**
* Resolve a data-table cell paragraph by (1-based) table/row/col/paragraph
* coordinates (matching find_text cell-coord hit format). Counts only
* top-level tables in document body.
*/
function findCellParagraphElement(doc, coords) {
	const body = firstChildNS(doc.documentElement, NS.w, "body");
	if (!body) return null;
	let tableCount = 0;
	for (const child of getChildren(body)) {
		if (child.namespaceURI !== NS.w || child.localName !== "tbl") continue;
		tableCount++;
		if (tableCount !== coords.table) continue;
		const rows = getChildrenNS(child, NS.w, "tr");
		if (coords.row < 1 || coords.row > rows.length) return null;
		const cells = getChildrenNS(rows[coords.row - 1], NS.w, "tc");
		if (coords.col < 1 || coords.col > cells.length) return null;
		const paras = getChildrenNS(cells[coords.col - 1], NS.w, "p");
		if (coords.paragraph < 1 || coords.paragraph > paras.length) return null;
		return paras[coords.paragraph - 1];
	}
	return null;
}
/**
* Walks the body in the same order as DocumentParser to assign indices,
* returning the <w:p> element that matches the target paragraph index.
* Includes paragraphs inside layout tables; skips data tables.
*/
function findParagraphElement(doc, targetIdx) {
	const body = firstChildNS(doc.documentElement, NS.w, "body");
	if (!body) return null;
	let counter = 0;
	let found = null;
	const visit = (parent) => {
		if (found) return;
		for (const child of getChildren(parent)) {
			if (found) return;
			if (child.namespaceURI !== NS.w) continue;
			if (child.localName === "p") {
				counter++;
				if (counter === targetIdx) {
					found = child;
					return;
				}
			} else if (child.localName === "tbl") {
				if (summarizeTable(child).classification === "layout") for (const tr of getChildrenNS(child, NS.w, "tr")) for (const tc of getChildrenNS(tr, NS.w, "tc")) {
					visit(tc);
					if (found) return;
				}
			}
		}
	};
	visit(body);
	return found;
}
function extractRuns(pEl) {
	const w = NS.w;
	const out = [];
	let runIdx = 0;
	for (const r of getChildrenNS(pEl, w, "r")) {
		runIdx++;
		let text = "";
		let hasDrawing = false;
		let hasMath = false;
		let hasBreak = false;
		let hasField = false;
		for (const c of getChildren(r)) if (c.namespaceURI === w) {
			if (c.localName === "t") text += textContent(c);
			else if (c.localName === "tab") text += "	";
			else if (c.localName === "br") hasBreak = true;
			else if (c.localName === "drawing") hasDrawing = true;
			else if (c.localName === "fldChar" || c.localName === "instrText") hasField = true;
		} else if (c.namespaceURI === NS.m) hasMath = true;
		const rPrEl = firstChildNS(r, w, "rPr");
		const rPr = rPrEl ? serializeRPr(rPrEl) : {};
		out.push({
			index: runIdx,
			text,
			rPr,
			hasDrawing,
			hasMath,
			hasBreak,
			hasField
		});
	}
	return out;
}
/**
* Convert <w:rPr>'s children into a flat property → string map. Boolean
* toggles ("b", "i", ...) collapse to "on" / "off" so they compare cleanly
* across runs. Font / size / color preserve their `w:val` (and rFonts'
* ascii/hAnsi/eastAsia attrs).
*/
function serializeRPr(rPr) {
	const w = NS.w;
	const out = {};
	for (const c of getChildren(rPr)) {
		if (c.namespaceURI !== w) continue;
		const name = c.localName;
		out[name] = signatureForRPrChild(c, name);
	}
	return out;
}
function signatureForRPrChild(el, name) {
	if (name === "rFonts") return [
		"ascii",
		"hAnsi",
		"eastAsia",
		"cs"
	].map((a) => `${a}=${wAttr(el, a) ?? ""}`).filter((p) => !p.endsWith("=")).join(" ");
	if (TOGGLE_PROPS.has(name)) {
		const v = wAttr(el, "val");
		return v === "0" || v === "false" ? "off" : "on";
	}
	return wAttr(el, "val") ?? "(present)";
}
const TOGGLE_PROPS = new Set([
	"b",
	"bCs",
	"i",
	"iCs",
	"caps",
	"smallCaps",
	"strike",
	"dstrike",
	"vanish",
	"snapToGrid",
	"noProof",
	"outline",
	"shadow",
	"emboss",
	"imprint"
]);
const ROLE_HINT_PROPS = [
	"rFonts",
	"sz",
	"szCs",
	"b",
	"bCs",
	"i",
	"iCs",
	"color",
	"u",
	"highlight",
	"strike",
	"vertAlign"
];
function renderReport(paraIdx, fingerprint, fullText, styleId, runs) {
	const lines = [];
	lines.push(`#${pad(paraIdx)} [${fingerprint}]  pStyle="${styleId}"  ${runs.length} run${runs.length === 1 ? "" : "s"}`);
	const truncated = fullText.length > 80 ? fullText.slice(0, 77) + "…" : fullText;
	lines.push(`  text: ${JSON.stringify(truncated)}`);
	lines.push("");
	appendRunsAndDiversity(lines, runs);
	return lines.join("\n");
}
/** Render a cell-paragraph hit using the same run + diversity format as
* renderReport, but with a cell-coord header instead of a paragraph index. */
function renderCellReport(locLabel, runs) {
	const lines = [];
	lines.push(`${locLabel}  ${runs.length} run${runs.length === 1 ? "" : "s"}`);
	lines.push("");
	appendRunsAndDiversity(lines, runs);
	return lines.join("\n");
}
function appendRunsAndDiversity(lines, runs) {
	if (runs.length === 0) {
		lines.push("  (paragraph has no runs — likely image-only or empty)");
		return;
	}
	const idxWidth = String(runs.length).length;
	for (const r of runs) {
		const tag = [];
		if (r.hasDrawing) tag.push("[drawing]");
		if (r.hasMath) tag.push("[math]");
		if (r.hasBreak) tag.push("[break]");
		if (r.hasField) tag.push("[field]");
		const tagStr = tag.length > 0 ? " " + tag.join(" ") : "";
		const idx = String(r.index).padStart(idxWidth, " ");
		const txt = r.text.length > 60 ? r.text.slice(0, 57) + "…" : r.text;
		lines.push(`  run ${idx}: text=${JSON.stringify(txt)}${tagStr}`);
		const rPrStr = formatRPrMap(r.rPr);
		lines.push(`         rPr=${rPrStr}`);
	}
	lines.push("");
	const propValues = /* @__PURE__ */ new Map();
	for (const name of ROLE_HINT_PROPS) propValues.set(name, /* @__PURE__ */ new Set());
	for (const r of runs) for (const name of ROLE_HINT_PROPS) propValues.get(name).add(r.rPr[name] ?? "<absent>");
	const mixed = [];
	const uniform = [];
	for (const [name, vals] of propValues) if (vals.size > 1) {
		const display = Array.from(vals).map((v) => v === "<absent>" ? "—" : v).join(" / ");
		mixed.push(`${name}: ${display}`);
	} else {
		const v = Array.from(vals)[0];
		if (v !== "<absent>") uniform.push(`${name}=${v}`);
	}
	lines.push(`  Run-level diversity:`);
	lines.push(`    mixed (preserved on restyle): ${mixed.length === 0 ? "(none)" : mixed.join(", ")}`);
	lines.push(`    uniform (strippable):         ${uniform.length === 0 ? "(none)" : uniform.join(", ")}`);
}
function formatRPrMap(rPr) {
	const keys = Object.keys(rPr);
	if (keys.length === 0) return "{}";
	const parts = [];
	for (const k of keys) {
		const v = rPr[k];
		if (v === "on") parts.push(k);
		else if (v === "off") parts.push(`${k}(off)`);
		else if (v === "(present)") parts.push(k);
		else parts.push(`${k}=${v}`);
	}
	return `{ ${parts.join(", ")} }`;
}
main();

//#endregion
export {  };