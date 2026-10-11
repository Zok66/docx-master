import { a as firstChildNS, h as NS, l as paragraphStyleId, m as walkBodyParagraphs, o as getChildren } from "./_shared/xml-utils.js";
import { n as parseNumbering, t as loadDocx } from "./_shared/load.js";
import { s as walkIndexedParagraphs } from "./_shared/locator.js";
import { a as truncate, i as paperName, o as tw2mm, r as pad } from "./_shared/format.js";
import { n as sectionUsableWidthTwips } from "./_shared/section-metrics.js";
import { n as seqFields } from "./_shared/field-parse.js";

//#region skill/tools/overview.ts
/**
* overview <docx> [--paras=A..B | --paras=none] [--include-unused]
*
* Default output is the full survey (metadata, page setup, theme, style
* definitions, numbering schemes, visual style summary, direct-format per
* fingerprint, document skeleton). Two scoping flags trim long-doc output
* without losing what an agent needs for first-time orientation:
*
*   --paras=A..B      Render the skeleton only for paragraphs in [A, B].
*                     Structural elements (tables / images / breaks) that
*                     fall within the range are kept; everything else is
*                     dropped. Use for second-look drilling after the
*                     initial full survey.
*   --paras=none      Skip the skeleton entirely. Useful when the agent
*                     only needs the style / numbering / theme dictionary
*                     (e.g. composing a small-scope `edits[]` config).
*   --include-unused  Show styles with usage=0 (hidden by default — Word
*                     and template residue often leaves dozens of declared-
*                     but-unused style definitions that drown the table).
*/
const NS_W = NS.w;
function parseArgs() {
	const argv = process.argv.slice(2);
	let paras = "full";
	let includeUnused = false;
	const positional = [];
	for (const a of argv) if (a === "--include-unused") includeUnused = true;
	else if (a.startsWith("--paras=")) {
		const spec = a.slice(8);
		if (spec === "none") paras = "none";
		else {
			const m = spec.match(/^(\d+)\.\.(\d+)$/);
			if (!m) {
				console.error(`--paras: expected "A..B" or "none", got "${spec}"`);
				process.exit(1);
			}
			const from = parseInt(m[1], 10);
			const to = parseInt(m[2], 10);
			if (from < 1 || to < from) {
				console.error(`--paras: invalid range ${from}..${to} (require 1 ≤ from ≤ to)`);
				process.exit(1);
			}
			paras = {
				from,
				to
			};
		}
	} else if (a.startsWith("--")) {
		console.error(`unknown flag: ${a}`);
		process.exit(1);
	} else positional.push(a);
	const file = positional[0];
	if (!file) {
		console.error("Usage: node scripts/overview.js <docx-path> [--paras=A..B | --paras=none] [--include-unused]");
		process.exit(1);
	}
	return {
		file,
		paras,
		includeUnused
	};
}
async function main() {
	const args = parseArgs();
	try {
		const doc = await loadDocx(args.file);
		const out = [];
		out.push(...renderMetadata(doc));
		out.push("");
		out.push(...renderPageSetup(doc));
		out.push("");
		out.push(...renderTheme(doc));
		out.push("");
		out.push(...renderStyleDefinitions(doc, args.includeUnused));
		out.push("");
		out.push(...renderNumbering(doc));
		out.push("");
		const captionLines = renderCaptions(doc);
		if (captionLines.length > 0) {
			out.push(...captionLines);
			out.push("");
		}
		out.push(...renderVisualSummary(doc));
		out.push("");
		out.push(...renderDirectFormatPerFingerprint(doc));
		out.push("");
		out.push(...renderSkeleton(doc, args.paras));
		console.log(out.join("\n"));
	} catch (err) {
		console.error(`Error: ${err.message}`);
		process.exit(1);
	}
}
function renderMetadata(doc) {
	const lines = ["=== Document Metadata ==="];
	lines.push(`File:    ${doc.metadata.fileName}`);
	lines.push(`Size:    ${formatSize(doc.metadata.fileSize)}`);
	if (doc.metadata.title) lines.push(`Title:   ${doc.metadata.title}`);
	if (doc.metadata.author) lines.push(`Author:  ${doc.metadata.author}`);
	lines.push(`Paragraphs: ${doc.paragraphs.length}`);
	lines.push(`Sections:   ${doc.sections.length}`);
	return lines;
}
function renderPageSetup(doc) {
	const lines = ["=== Page Setup ==="];
	const s = doc.sections[doc.sections.length - 1] || doc.sections[0];
	if (!s) {
		lines.push("(no section properties)");
		return lines;
	}
	const paper = paperName(s.pageSize.width, s.pageSize.height);
	lines.push(`Paper:       ${paper} (${tw2mm(s.pageSize.width)} × ${tw2mm(s.pageSize.height)} mm)`);
	lines.push(`Orientation: ${s.orientation}`);
	lines.push(`Margins:     top=${tw2mm(s.margins.top)} bottom=${tw2mm(s.margins.bottom)} left=${tw2mm(s.margins.left)} right=${tw2mm(s.margins.right)} mm`);
	const usable = sectionUsableWidthTwips(s);
	if (usable > 0) lines.push(`Text width:  ${tw2mm(usable)} mm / ${(usable / 20).toFixed(1)} pt`);
	return lines;
}
function renderTheme(doc) {
	const lines = ["=== Theme ==="];
	const fonts = doc.resolver.getThemeFonts();
	const colors = doc.resolver.getThemeColors();
	if (fonts.majorLatin || fonts.majorEastAsia) lines.push(`Major font: ${fonts.majorLatin || "?"}${fonts.majorEastAsia ? ` / ${fonts.majorEastAsia}` : ""}`);
	if (fonts.minorLatin || fonts.minorEastAsia) lines.push(`Minor font: ${fonts.minorLatin || "?"}${fonts.minorEastAsia ? ` / ${fonts.minorEastAsia}` : ""}`);
	const accents = [
		"accent1",
		"accent2",
		"accent3",
		"accent4",
		"accent5",
		"accent6"
	].filter((s) => colors[s]).map((s) => `${s}=#${colors[s]}`);
	if (accents.length > 0) lines.push(`Accent colors: ${accents.join(" ")}`);
	if (lines.length === 1) lines.push("(no theme)");
	return lines;
}
function renderStyleDefinitions(doc, includeUnused) {
	const lines = ["=== Style Definitions (styles.xml) ==="];
	const allNonDefault = doc.resolver.getAllStyles().filter((s) => !s.isDefault);
	const shown = includeUnused ? allNonDefault : allNonDefault.filter((s) => s.usageCount > 0);
	shown.sort((a, b) => b.usageCount - a.usageCount || a.id.localeCompare(b.id));
	if (shown.length === 0) lines.push("(no styles)");
	else for (const s of shown) {
		const props = [];
		if (s.rPr.size !== void 0) props.push(`${s.rPr.size / 2}pt`);
		if (s.rPr.bold) props.push("Bold");
		if (s.rPr.italic) props.push("Italic");
		if (s.rPr.fontEastAsia || s.rPr.fontAscii) props.push(s.rPr.fontEastAsia || s.rPr.fontAscii || "");
		if (s.pPr.alignment) props.push(s.pPr.alignment);
		if (s.pPr.outlineLevel !== void 0) props.push(`outlineLvl=${s.pPr.outlineLevel}`);
		if (s.pPr.numId) props.push(`numId=${s.pPr.numId}`);
		const based = s.basedOn ? ` basedOn=${s.basedOn}` : "";
		lines.push(`  ${s.id} "${s.name}" [${s.type}]${based}  usage=${s.usageCount}  {${props.join(", ")}}`);
	}
	const hidden = allNonDefault.length - shown.length;
	if (hidden > 0) lines.push(`  (${hidden} unused style${hidden === 1 ? "" : "s"} hidden — pass --include-unused to show)`);
	return lines;
}
function renderNumbering(doc) {
	const lines = ["=== Numbering Definitions (numbering.xml) ==="];
	const defs = parseNumbering(doc.numberingDoc);
	if (defs.length === 0) {
		lines.push("(no numbering)");
		return lines;
	}
	const usageByNumId = /* @__PURE__ */ new Map();
	for (const p of doc.paragraphs) if (p.pPr.numId !== void 0) usageByNumId.set(p.pPr.numId, (usageByNumId.get(p.pPr.numId) ?? 0) + 1);
	const clusters = /* @__PURE__ */ new Map();
	for (const def of defs) {
		const sig = def.levels.map((l) => `${l.level}:${l.format}|${l.text}|${l.pStyle ?? ""}`).join(";;");
		if (!clusters.has(sig)) clusters.set(sig, []);
		clusters.get(sig).push(def);
	}
	let clusterIdx = 0;
	for (const [, group] of clusters) {
		clusterIdx++;
		const sample = group[0];
		const ids = group.map((d) => d.numId).join(", ");
		const starts = new Set(group.map((d) => d.levels[0]?.start ?? 1));
		const startsStr = starts.size === 1 ? `start=${[...starts][0]}` : `starts: {${[...starts].sort((a, b) => a - b).join(", ")}}`;
		const usedBy = group.reduce((sum, d) => sum + (usageByNumId.get(d.numId) ?? 0), 0);
		const usedStr = `used by ${usedBy} paragraph${usedBy === 1 ? "" : "s"}`;
		if (group.length === 1) lines.push(`  numId=${sample.numId} (abstract=${sample.abstractNumId})  ${usedStr}`);
		else lines.push(`  Scheme ${clusterIdx} × ${group.length} numIds: [${ids}]  ${startsStr}  ${usedStr}`);
		const meaningful = sample.levels.filter((l) => l.text.length > 0 || l.pStyle !== void 0);
		const shown = meaningful.length > 0 ? meaningful : sample.levels.slice(0, 1);
		for (const lvl of shown) {
			const ps = lvl.pStyle ? ` pStyle=${lvl.pStyle}` : "";
			const startNote = group.length === 1 ? ` start=${lvl.start}` : "";
			lines.push(`    L${lvl.level}: numFmt=${lvl.format} lvlText="${lvl.text}"${startNote}${ps}`);
		}
	}
	return lines;
}
function renderVisualSummary(doc) {
	const lines = ["=== Visual Style Summary (deduplicated) ==="];
	for (const s of doc.summary) {
		const facts = [`×${s.count}`, `avg ${s.avgTextLength}ch`];
		if (s.boundStyleId) facts.push(s.boundStyleName ? `via "${s.boundStyleName}"/${s.boundStyleId}` : `via ${s.boundStyleId}`);
		lines.push(`${s.label} [${s.hash}]: ${s.description.padEnd(36, " ")} ${facts.join("  ")}`);
	}
	return lines;
}
/** For each fingerprint, list which pPr children the first sample paragraph
* carries DIRECTLY (not via style cascade), plus the union of run-level rPr
* children across the paragraph's `<w:r>` children. Lets the agent decide
* whether an attribute is already set as chrome convention — in which case
* the style shouldn't redeclare it. The values themselves aren't shown here
* (use `inspect_range` for those); presence/absence is what governs the
* "don't override what chrome already provides" rule across both axes. */
function renderDirectFormatPerFingerprint(doc) {
	const lines = ["=== Direct Format per Fingerprint ==="];
	const indexToElement = /* @__PURE__ */ new Map();
	for (const p of walkIndexedParagraphs(doc.documentDoc)) indexToElement.set(p.index, p.element);
	const firstByLabel = /* @__PURE__ */ new Map();
	for (const p of doc.paragraphs) if (!firstByLabel.has(p.fingerprint)) firstByLabel.set(p.fingerprint, p);
	for (const s of doc.summary) {
		const sample = firstByLabel.get(s.label);
		if (!sample) continue;
		const el = indexToElement.get(sample.index);
		if (!el) continue;
		const pPrNames = [];
		const pPr = firstChildNS(el, NS.w, "pPr");
		if (pPr) for (const c of getChildren(pPr)) {
			if (c.namespaceURI !== NS.w) continue;
			if (c.localName === "rPr") pPrNames.push("rPr(pMark)");
			else pPrNames.push(c.localName);
		}
		const rPrNames = /* @__PURE__ */ new Set();
		for (const c of getChildren(el)) {
			if (c.namespaceURI !== NS.w || c.localName !== "r") continue;
			const rPr = firstChildNS(c, NS.w, "rPr");
			if (!rPr) continue;
			for (const cc of getChildren(rPr)) {
				if (cc.namespaceURI !== NS.w) continue;
				rPrNames.add(cc.localName);
			}
		}
		const pPrStr = pPrNames.length > 0 ? pPrNames.join(", ") : "(none)";
		const rPrStr = rPrNames.size > 0 ? [...rPrNames].join(", ") : "(none)";
		lines.push(`  ${s.label}  pPr: ${pPrStr}  |  rPr: ${rPrStr}`);
	}
	return lines;
}
function renderSkeleton(doc, paras) {
	if (paras === "none") return ["=== Document Skeleton ===", "(omitted — pass --paras=A..B or remove the flag to render the skeleton)"];
	const range = paras === "full" ? null : paras;
	const lines = [range === null ? "=== Document Skeleton ===" : `=== Document Skeleton (paragraphs #${pad(range.from)}–#${pad(range.to)} of ${doc.paragraphs.length}) ===`];
	const sectionElements = /* @__PURE__ */ new Map();
	for (const el of doc.elements) {
		const idx = elementSectionIndex(el);
		if (!sectionElements.has(idx)) sectionElements.set(idx, []);
		sectionElements.get(idx).push(el);
	}
	let renderedAnything = false;
	for (const sec of doc.sections) {
		if (range && (sec.paraRange[1] < range.from || sec.paraRange[0] > range.to)) continue;
		const headerLine = `--- Section ${sec.index + 1} (para #${sec.paraRange[0]}-#${sec.paraRange[1]}) ---`;
		lines.push(headerLine);
		if (sec.header !== null) lines.push(`Header: ${truncate(sec.header, 60)}`);
		else lines.push(`Header: (none)`);
		if (sec.footer !== null) lines.push(`Footer: ${truncate(sec.footer, 60)}`);
		else lines.push(`Footer: (none)`);
		if (sec.footerPageNumFormat) lines.push(`Footer page num: ${sec.footerPageNumFormat}`);
		lines.push("");
		const elems = sectionElements.get(sec.index) || [];
		let anchor = sec.paraRange[0];
		for (const el of elems) {
			const elemAnchor = updateAnchor(el, anchor);
			anchor = elemAnchor;
			if (range && !elementInRange(el, range, elemAnchor)) continue;
			lines.push(...renderElement(el, "", range));
			renderedAnything = true;
		}
		lines.push("");
	}
	if (range && !renderedAnything) lines.push(`(no paragraphs in range #${range.from}-#${range.to})`);
	return lines;
}
function updateAnchor(el, prev) {
	if (el.kind === "paragraph") return el.paragraph.index;
	if (el.kind === "emptyRun") return el.firstIndex;
	if (el.kind === "table" && el.classification === "layout" && el.paragraphs.length > 0) return el.paragraphs[0].index;
	return prev;
}
function elementInRange(el, range, anchor) {
	if (el.kind === "paragraph") {
		const i = el.paragraph.index;
		return i >= range.from && i <= range.to;
	}
	if (el.kind === "emptyRun") return el.firstIndex + el.count - 1 >= range.from && el.firstIndex <= range.to;
	if (el.kind === "table" && el.classification === "layout") {
		if (el.paragraphs.length === 0) return false;
		return el.paragraphs.some((p) => p.index >= range.from && p.index <= range.to);
	}
	return anchor >= range.from && anchor <= range.to;
}
function elementSectionIndex(el) {
	if (el.kind === "paragraph") return el.paragraph.context.sectionIndex;
	return el.sectionIndex;
}
function renderElement(el, indent, range) {
	const lines = [];
	if (el.kind === "paragraph") {
		const p = el.paragraph;
		lines.push(`${indent}  #${pad(p.index)} [${p.fingerprint}]  "${truncate(p.text, 40)}"`);
	} else if (el.kind === "table") if (el.classification === "layout") {
		lines.push(`${indent}--- LAYOUT TABLE  reason:${el.classificationReason} ---`);
		for (const p of el.paragraphs) {
			if (range && (p.index < range.from || p.index > range.to)) continue;
			lines.push(`${indent}    #${pad(p.index)} [${p.fingerprint}]  "${truncate(p.text, 40)}"`);
		}
		lines.push(`${indent}--- END LAYOUT TABLE ---`);
	} else if (el.row1Texts.some((t) => t.length > 0)) {
		const formatted = el.row1Texts.map((t) => `"${truncate(t, 12)}"`).slice(0, el.cols).join(",");
		lines.push(`${indent}--- TABLE (${el.rows}×${el.cols}) row1:[${formatted}]  reason:${el.classificationReason} ---`);
	} else lines.push(`${indent}--- TABLE (${el.rows}×${el.cols})  reason:${el.classificationReason} ---`);
	else if (el.kind === "image") lines.push(`${indent}--- IMAGE (${el.widthCm.toFixed(1)}cm × ${el.heightCm.toFixed(1)}cm) ---`);
	else if (el.kind === "equation") lines.push(`${indent}--- EQUATION ---`);
	else if (el.kind === "pageBreak") lines.push(`${indent}--- PAGEBREAK ---`);
	else if (el.kind === "sectionBreak") {} else if (el.kind === "emptyRun") if (el.count === 1) lines.push(`${indent}  #${pad(el.firstIndex)} --- empty ---`);
	else lines.push(`${indent}  #${pad(el.firstIndex)} --- empty ×${el.count} ---`);
	return lines;
}
function formatSize(n) {
	if (n < 1024) return `${n} B`;
	if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
	return `${(n / 1024 / 1024).toFixed(2)} MB`;
}
function renderCaptions(doc) {
	const documentDoc = doc.documentDoc;
	if (!documentDoc) return [];
	const body = firstChildNS(documentDoc.documentElement, NS_W, "body");
	if (!body) return [];
	const byId = /* @__PURE__ */ new Map();
	for (const para of walkBodyParagraphs(body)) {
		const advancingSeqs = seqFields(para, { skipRepeat: true });
		if (advancingSeqs.length === 0) continue;
		const styleId = paragraphStyleId(para);
		for (const details of advancingSeqs) {
			const id = details.identifier;
			if (!id) continue;
			let summary = byId.get(id);
			if (!summary) {
				summary = {
					identifier: id,
					format: details.format,
					restartAtOutlineLevel: details.restartAtOutlineLevel,
					paragraphStyleIds: /* @__PURE__ */ new Set(),
					occurrences: 0
				};
				byId.set(id, summary);
			}
			summary.occurrences++;
			if (styleId) summary.paragraphStyleIds.add(styleId);
		}
	}
	if (byId.size === 0) return [];
	const lines = ["=== Captions detected (SEQ identifiers in body) ==="];
	for (const s of byId.values()) {
		const styleNote = s.paragraphStyleIds.size === 0 ? "(unstyled)" : [...s.paragraphStyleIds].join(", ");
		const fmt = s.format ?? "(unknown)";
		const restart = s.restartAtOutlineLevel ? `restart@H${s.restartAtOutlineLevel}` : "global";
		lines.push(`  ${s.identifier.padEnd(20)} format=${fmt.padEnd(10)} ${restart.padEnd(14)} occurrences=${s.occurrences}  style=${styleNote}`);
	}
	lines.push(`  (run \`inspect_caption <doc> <identifier>\` for per-occurrence details)`);
	return lines;
}
main();

//#endregion
export {  };