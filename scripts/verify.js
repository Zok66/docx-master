#!/usr/bin/env node
/**
 * `verify` — the delivery gate. Two jobs, both aimed at the same blind spot:
 * schema validation proves a file is well-formed; it says nothing about
 * whether the file is *what the user asked for*.
 *
 * 1. ASSERTIONS. Declarative expectations about the finished artifact, checked
 *    mechanically, with a non-zero exit on any failure. This is what catches
 *    the class of defect no XSD can see — e.g. a footer that carries an extra
 *    title nobody asked for:
 *
 *      { "type": "notContains", "text": "My Book", "scope": "footer" }
 *
 * 2. CHANGE MANIFEST. A structural diff of output vs. baseline that spells out
 *    every difference the run introduced: sections, paragraph count, parts,
 *    header/footer text, fields, styles. It exists to be READ. Each line is a
 *    change you made; anything on the list the user did not ask for is a
 *    defect, not a feature.
 *
 * Zero dependencies — reads the OPC package directly (Node's zlib only).
 *
 * Usage:
 *   node scripts/verify.js <verify.json> [--diff <baseline.docx>] [--json] [--quiet]
 *   node scripts/verify.js <file.docx> [--diff <baseline.docx>]      # analysis + manifest only
 *
 * verify.json:
 * {
 *   "file": "out.docx",              // required — the artifact to check
 *   "baseline": "src.docx",          // optional — enables the change manifest
 *   "assert": [
 *     { "type": "sectionCount", "eq": 3 },
 *     { "type": "paragraphCount", "min": 400 },
 *     { "type": "headerText", "regex": "佐罗斯" },
 *     { "type": "notContains", "text": "Linux 基础命令笔记", "scope": "footer" },
 *     { "type": "footerText", "regex": "^\\s*第\\s*<PAGE>\\s*页\\s*\\/\\s*共\\s*<NUMPAGES>\\s*页\\s*$" },
 *     { "type": "hasField", "field": "TOC", "scope": "body" },
 *     { "type": "pageSize", "w": 11906, "h": 16838 },
 *     { "type": "styleExists", "styleId": "Heading1" }
 *   ]
 * }
 *
 * Assertion types
 *   Counts      sectionCount · paragraphCount · tableCount · imageCount
 *   Presence    contains · notContains · textMatches       (scope: body | header | footer | all)
 *   Parts       headerText · footerText                    (regex / equals, <PAGE> = field marker)
 *   Fields      hasField · noField
 *   Setup       pageSize {w,h} · partExists {name}
 *   Styles      styleExists · styleMissing
 *   Metadata    author · title
 *
 * Comparison keys on count-ish assertions: eq · ne · min · max
 * Exit code: 0 all assertions passed; 1 any failed (or the file is unreadable).
 */
import { readFileSync, existsSync } from "node:fs";
import { inflateRawSync } from "node:zlib";
import { resolve, dirname, isAbsolute } from "node:path";

// ── minimal OPC (zip) reader — no dependencies ──────────────────────────────

const EOCD = 0x06054b50;
const CEN = 0x02014b50;
const LOC = 0x04034b50;

function openZip(path) {
	const buf = readFileSync(path);
	let eocd = -1;
	const floor = Math.max(0, buf.length - 66000);
	for (let i = buf.length - 22; i >= floor; i--) {
		if (buf.readUInt32LE(i) === EOCD) {
			eocd = i;
			break;
		}
	}
	if (eocd < 0) throw new Error("not a zip/OPC package (no end-of-central-directory record)");
	const count = buf.readUInt16LE(eocd + 10);
	let off = buf.readUInt32LE(eocd + 16);
	const entries = new Map();
	for (let i = 0; i < count; i++) {
		if (off + 46 > buf.length || buf.readUInt32LE(off) !== CEN) break;
		const method = buf.readUInt16LE(off + 10);
		const compSize = buf.readUInt32LE(off + 20);
		const nameLen = buf.readUInt16LE(off + 28);
		const extraLen = buf.readUInt16LE(off + 30);
		const commentLen = buf.readUInt16LE(off + 32);
		const localOff = buf.readUInt32LE(off + 42);
		const name = buf.toString("utf8", off + 46, off + 46 + nameLen);
		entries.set(name, { method, compSize, localOff });
		off += 46 + nameLen + extraLen + commentLen;
	}
	return {
		buf,
		names: [...entries.keys()],
		has: (n) => entries.has(n),
		readText(name) {
			const e = entries.get(name);
			if (!e) return null;
			if (buf.readUInt32LE(e.localOff) !== LOC) throw new Error(`corrupt local header for ${name}`);
			const nLen = buf.readUInt16LE(e.localOff + 26);
			const eLen = buf.readUInt16LE(e.localOff + 28);
			const start = e.localOff + 30 + nLen + eLen;
			const raw = buf.subarray(start, start + e.compSize);
			const out = e.method === 0 ? raw : inflateRawSync(raw);
			return out.toString("utf8");
		}
	};
}

// ── XML text extraction (regex-based; sufficient for auditing, not editing) ─

function decode(s) {
	return s
		.replace(/&lt;/g, "<")
		.replace(/&gt;/g, ">")
		.replace(/&quot;/g, '"')
		.replace(/&apos;/g, "'")
		.replace(/&#(\d+);/g, (_, d) => String.fromCharCode(+d))
		.replace(/&amp;/g, "&");
}

const TOKEN_RE =
	/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:instrText(?:\s[^>]*)?>([\s\S]*?)<\/w:instrText>|<w:fldChar\b[^>]*\bw:fldCharType="(\w+)"[^>]*?\/?>|<w:tab\b[^>]*\/>|<w:(?:br|cr)\b[^>]*\/>/g;

function fieldName(instr) {
	const m = instr.trim().match(/^\\?\s*([A-Za-z]+)/);
	return m ? m[1].toUpperCase() : "";
}

/**
 * Ordered visible text of a part. Fields render as `<FIELDNAME>`.
 *
 * Field *results* are skipped: the text between `separate` and `end` is a
 * cache of whatever Word last computed (the literal "1" behind a PAGE field,
 * the entry list behind a TOC). It is not part of the document's authored
 * content and changes on every repagination, so including it would make both
 * text assertions and the change manifest jitter. The field itself is what
 * the document actually declares — that is what gets compared.
 */
function textOf(xml) {
	if (!xml) return "";
	let out = "";
	let depth = 0;
	let inResult = false;
	let m;
	TOKEN_RE.lastIndex = 0;
	while ((m = TOKEN_RE.exec(xml))) {
		if (m[1] !== undefined) {
			if (!inResult) out += decode(m[1]);
		} else if (m[2] !== undefined) {
			const f = fieldName(m[2]);
			if (f && !inResult) out += `<${f}>`;
		} else if (m[3] !== undefined) {
			const type = m[3].toLowerCase();
			if (type === "begin") depth++;
			else if (type === "separate") inResult = true;
			else if (type === "end") {
				depth = Math.max(0, depth - 1);
				if (depth === 0) inResult = false;
			}
		} else if (m[0].startsWith("<w:tab")) {
			if (!inResult) out += "\t";
		} else if (!inResult) out += "\n";
	}
	return out;
}

/** Field names appearing anywhere in a part (including complex-field runs). */
function fieldsOf(xml) {
	const names = new Set();
	if (!xml) return names;
	const re = /<w:instrText(?:\s[^>]*)?>([\s\S]*?)<\/w:instrText>|<w:fldSimple\b[^>]*\bw:instr="([^"]*)"/g;
	let m;
	while ((m = re.exec(xml))) {
		const f = fieldName(m[1] !== undefined ? m[1] : m[2] || "");
		if (f) names.add(f);
	}
	return names;
}

function paragraphsOf(xml) {
	if (!xml) return [];
	const re = /<w:p\b[^>]*\/>|<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
	const list = [];
	let m;
	while ((m = re.exec(xml))) list.push(m[0]);
	return list;
}

function countMatches(xml, re) {
	if (!xml) return 0;
	return (xml.match(re) || []).length;
}

// ── document analysis ──────────────────────────────────────────────────────

function parseRels(xml) {
	const out = [];
	if (!xml) return out;
	const re = /<Relationship\b([^>]*)\/?>/g;
	let m;
	while ((m = re.exec(xml))) {
		const attrs = m[1];
		const id = (attrs.match(/Id="([^"]*)"/) || [])[1] || "";
		const type = (attrs.match(/Type="([^"]*)"/) || [])[1] || "";
		const target = (attrs.match(/Target="([^"]*)"/) || [])[1] || "";
		if (target && !target.startsWith("http")) out.push({ id, type, target: target.replace(/^\/?word\//, "").replace(/^\//, "") });
	}
	return out;
}

function analyze(file) {
	if (!existsSync(file)) throw new Error(`file not found: ${file}`);
	const zip = openZip(file);
	const bodyXml = zip.readText("word/document.xml");
	if (!bodyXml) throw new Error("word/document.xml missing — not a Word document?");

	const rels = parseRels(zip.readText("word/_rels/document.xml.rels"));
	const headers = [];
	const footers = [];
	for (const r of rels) {
		if (/\/header$/.test(r.type)) {
			const t = zip.readText(`word/${r.target}`);
			if (t) headers.push({ target: r.target, xml: t });
		} else if (/\/footer$/.test(r.type)) {
			const t = zip.readText(`word/${r.target}`);
			if (t) footers.push({ target: r.target, xml: t });
		}
	}
	// Fallback for packages with missing/odd rels: scan by conventional name.
	for (const n of zip.names) {
		if (/^word\/header\d*\.xml$/.test(n) && !headers.some((h) => `word/${h.target}` === n)) {
			headers.push({ target: n.replace("word/", ""), xml: zip.readText(n) });
		}
		if (/^word\/footer\d*\.xml$/.test(n) && !footers.some((f) => `word/${f.target}` === n)) {
			footers.push({ target: n.replace("word/", ""), xml: zip.readText(n) });
		}
	}

	const stylesXml = zip.readText("word/styles.xml") || "";
	const coreXml = zip.readText("docProps/core.xml") || "";
	const paragraphs = paragraphsOf(bodyXml);

	// Page size: prefer the body-level sectPr (last one in the part).
	const sectPrs = bodyXml.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g) || [];
	const sizes = sectPrs
		.map((s) => {
			const m = s.match(/<w:pgSz\b[^>]*>/);
			if (!m) return null;
			const w = (m[0].match(/w:w="(\d+)"/) || [])[1];
			const h = (m[0].match(/w:h="(\d+)"/) || [])[1];
			const orient = (m[0].match(/w:orient="(\w+)"/) || [])[1] || "portrait";
			return w && h ? { w: +w, h: +h, orient } : null;
		})
		.filter(Boolean);

	// Style ids are engine-internal (a writer may emit `2`, `10`, … — including
	// this skill's own `apply`); the *name* is the semantic identity a human
	// checks against. Collect both, and let assertions use either.
	const styleIds = new Set();
	const styleNames = new Set();
	const styleIdToName = new Map();
	{
		const re = /<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g;
		let m;
		while ((m = re.exec(stylesXml))) {
			const id = (m[1].match(/w:styleId="([^"]*)"/) || [])[1];
			const name = (m[2].match(/<w:name\b[^>]*w:val="([^"]*)"/) || [])[1];
			if (id) styleIds.add(id);
			if (name && name !== id) styleNames.add(name);
			if (id && name) styleIdToName.set(id, name);
		}
	}

	const allFields = new Set([...fieldsOf(bodyXml)]);
	for (const h of headers) for (const name of fieldsOf(h.xml)) allFields.add(name);
	for (const ft of footers) for (const name of fieldsOf(ft.xml)) allFields.add(name);

	return {
		file,
		zip,
		parts: { body: bodyXml, headers, footers },
		sectionCount: sectPrs.length,
		paragraphCount: paragraphs.length,
		paragraphTexts: paragraphs.map(textOf),
		tableCount: countMatches(bodyXml, /<w:tbl>/g),
		imageCount: zip.names.filter((n) => /^word\/media\//.test(n)).length,
		styles: styleIds,
		styleNames,
		styleIdToName,
		fields: allFields,
		pageSizes: sizes,
		author: (coreXml.match(/<dc:creator>([\s\S]*?)<\/dc:creator>/) || [])[1] || "",
		title: (coreXml.match(/<dc:title>([\s\S]*?)<\/dc:title>/) || [])[1] || "",
		headerText: headers.map((h) => textOf(h.xml)),
		footerText: footers.map((f) => textOf(f.xml)),
		bodyText: textOf(bodyXml)
	};
}

// ── assertion engine ───────────────────────────────────────────────────────

function scopeTexts(a, scope = "all") {
	if (scope === "body") return [a.bodyText];
	if (scope === "header") return a.headerText;
	if (scope === "footer") return a.footerText;
	return [a.bodyText, ...a.headerText, ...a.footerText];
}

function scopeFields(a, scope = "all") {
	if (scope === "body") return fieldsOf(a.parts.body);
	if (scope === "header") return new Set(a.parts.headers.flatMap((h) => [...fieldsOf(h.xml)]));
	if (scope === "footer") return new Set(a.parts.footers.flatMap((f) => [...fieldsOf(f.xml)]));
	return a.fields;
}

function countCmp(actual, spec) {
	if (spec.eq !== undefined) return { pass: actual === spec.eq, want: `= ${spec.eq}` };
	if (spec.ne !== undefined) return { pass: actual !== spec.ne, want: `≠ ${spec.ne}` };
	const parts = [];
	let pass = true;
	if (spec.min !== undefined) {
		parts.push(`≥ ${spec.min}`);
		pass = pass && actual >= spec.min;
	}
	if (spec.max !== undefined) {
		parts.push(`≤ ${spec.max}`);
		pass = pass && actual <= spec.max;
	}
	return { pass, want: parts.join(" and ") || "(no constraint)" };
}

function evalAssertion(a, spec) {
	const t = spec.type;
	const label = () => {
		const scope = spec.scope && spec.scope !== "all" ? ` [${spec.scope}]` : "";
		switch (t) {
			case "contains":
			case "notContains":
				return `${t} ${spec.regex ? `/${spec.regex}/` : JSON.stringify(spec.text)}${scope}`;
			case "textMatches":
				return `textMatches /${spec.regex}/${scope}`;
			case "headerText":
				return `headerText ${spec.regex ? `/${spec.regex}/` : JSON.stringify(spec.equals)}`;
			case "footerText":
				return `footerText ${spec.regex ? `/${spec.regex}/` : JSON.stringify(spec.equals)}`;
			case "hasField":
			case "noField":
				return `${t} ${spec.field}${scope}`;
			case "styleExists":
			case "styleMissing":
				return `${t} ${spec.styleId || spec.name}`;
			case "partExists":
				return `partExists ${spec.name}`;
			case "pageSize":
				return `pageSize ${spec.w ?? "?"}×${spec.h ?? "?"}${spec.orient ? ` ${spec.orient}` : ""}`;
			case "sectionCount":
			case "paragraphCount":
			case "tableCount":
			case "imageCount": {
				const bits = [];
				if (spec.eq !== undefined) bits.push(`= ${spec.eq}`);
				if (spec.ne !== undefined) bits.push(`≠ ${spec.ne}`);
				if (spec.min !== undefined) bits.push(`≥ ${spec.min}`);
				if (spec.max !== undefined) bits.push(`≤ ${spec.max}`);
				return `${t} ${bits.join(" and ") || "(no constraint)"}`;
			}
			case "author":
			case "title":
				return `${t} ${spec.regex ? `/${spec.regex}/` : JSON.stringify(spec.equals)}`;
			default:
				return `${t} ${JSON.stringify(spec).slice(0, 70)}`;
		}
	};

	const texts = scopeTexts(a, spec.scope);
	const joined = texts.join("\n");

	switch (t) {
		case "sectionCount": {
			const r = countCmp(a.sectionCount, spec);
			return { label: label(), pass: r.pass, detail: `actual ${a.sectionCount}, expected ${r.want}` };
		}
		case "paragraphCount": {
			const r = countCmp(a.paragraphCount, spec);
			return { label: label(), pass: r.pass, detail: `actual ${a.paragraphCount}, expected ${r.want}` };
		}
		case "tableCount": {
			const r = countCmp(a.tableCount, spec);
			return { label: label(), pass: r.pass, detail: `actual ${a.tableCount}, expected ${r.want}` };
		}
		case "imageCount": {
			const r = countCmp(a.imageCount, spec);
			return { label: label(), pass: r.pass, detail: `actual ${a.imageCount}, expected ${r.want}` };
		}
		case "contains":
		case "notContains": {
			const hit = spec.regex ? new RegExp(spec.regex).test(joined) : joined.includes(spec.text ?? "");
			const pass = t === "contains" ? hit : !hit;
			const detail = hit
				? `found in ${spec.scope || "all"}`
				: `not found in ${spec.scope || "all"}`;
			return { label: label(), pass, detail: `${detail}${pass ? "" : "  ← the defect this assertion exists to catch"}` };
		}
		case "textMatches": {
			const pass = new RegExp(spec.regex).test(joined);
			return { label: label(), pass, detail: pass ? "matched" : "no match" };
		}
		case "headerText":
		case "footerText": {
			const arr = t === "headerText" ? a.headerText : a.footerText;
			const test = (txt) => (spec.regex ? new RegExp(spec.regex).test(txt) : txt === (spec.equals ?? ""));
			// No `index`: any part matching is a pass (a document may hold several
			// parts — cover-page blanks, per-section overrides — and "the footer
			// shows a page number" is satisfied by one of them). With `index`,
			// that exact part must match.
			if (spec.index === undefined) {
				if (!arr.length) return { label: label(), pass: false, detail: `no ${t === "headerText" ? "header" : "footer"} part at all` };
				const hit = arr.findIndex(test);
				return {
					label: label(),
					pass: hit !== -1,
					detail: hit !== -1 ? `matched part[${hit}]` : `no part matched — parts: ${arr.map((x) => JSON.stringify(x)).join(", ")}`
				};
			}
			const actual = arr[spec.index];
			if (actual === undefined) {
				return { label: label(), pass: false, detail: `no ${t === "headerText" ? "header" : "footer"} part at index ${spec.index} (found ${arr.length})` };
			}
			const pass = test(actual);
			return { label: label(), pass, detail: `${pass ? "ok" : "mismatch"} — actual ${JSON.stringify(actual)}` };
		}
		case "hasField":
		case "noField": {
			const has = scopeFields(a, spec.scope).has(String(spec.field).toUpperCase());
			const pass = t === "hasField" ? has : !has;
			return { label: label(), pass, detail: pass ? "ok" : `${t === "hasField" ? "missing" : "unexpectedly present"} in ${spec.scope || "all"}` };
		}
		case "styleExists":
		case "styleMissing": {
			const byId = spec.styleId ? a.styles.has(spec.styleId) : null;
			const byName = spec.name ? a.styleNames.has(spec.name) : null;
			if (byId === null && byName === null) {
				return { label: label(), pass: false, detail: "assertion needs styleId or name" };
			}
			const has = byId !== null && byName !== null ? byId && byName : byId !== null ? byId : byName;
			const pass = t === "styleExists" ? has : !has;
			const what = spec.styleId || spec.name;
			return { label: label(), pass, detail: pass ? "ok" : `${what} ${has ? "present" : "absent"}` };
		}
		case "partExists": {
			const pass = a.zip.has(spec.name);
			return { label: label(), pass, detail: pass ? "present" : "missing from package" };
		}
		case "pageSize": {
			const idx = spec.index ?? a.pageSizes.length - 1;
			const ps = a.pageSizes[idx];
			if (!ps) return { label: label(), pass: false, detail: "no pgSz found" };
			const pass = (!spec.w || ps.w === spec.w) && (!spec.h || ps.h === spec.h) && (!spec.orient || ps.orient === spec.orient);
			return { label: label(), pass, detail: `actual ${ps.w}×${ps.h} ${ps.orient}` };
		}
		case "author":
		case "title": {
			const actual = a[t];
			const pass = spec.regex ? new RegExp(spec.regex).test(actual) : actual === (spec.equals ?? "");
			return { label: label(), pass, detail: `${pass ? "ok" : "mismatch"} — actual ${JSON.stringify(actual)}` };
		}
		default:
			return { label: label(), pass: false, detail: `unknown assertion type "${t}"` };
	}
}

// ── change manifest (the part a human must read) ────────────────────────────

function manifest(base, cur) {
	const lines = [];
	const arrow = (from, to) => {
		if (from === to) return `${from} (unchanged)`;
		const d = typeof to === "number" && typeof from === "number" ? (to - from >= 0 ? ` (+${to - from})` : ` (${to - from})`) : "  ← changed";
		return `${from} → ${to}${d}`;
	};
	lines.push(["sections", arrow(base.sectionCount, cur.sectionCount)]);
	lines.push(["paragraphs", arrow(base.paragraphCount, cur.paragraphCount)]);
	lines.push(["tables", arrow(base.tableCount, cur.tableCount)]);
	lines.push(["images", arrow(base.imageCount, cur.imageCount)]);
	lines.push([
		"page size",
		base.pageSizes.length && cur.pageSizes.length
			? arrow(`${base.pageSizes[base.pageSizes.length - 1].w}×${base.pageSizes[base.pageSizes.length - 1].h}`, `${cur.pageSizes[cur.pageSizes.length - 1].w}×${cur.pageSizes[cur.pageSizes.length - 1].h}`)
			: `${cur.pageSizes.length} (new)`
	]);
	lines.push(["header parts", arrow(base.headerText.length, cur.headerText.length)]);
	lines.push(["footer parts", arrow(base.footerText.length, cur.footerText.length)]);
	for (let i = 0; i < cur.headerText.length; i++) {
		const before = base.headerText[i];
		if (before !== cur.headerText[i]) lines.push([`  header[${i}]`, before === undefined ? `(new) ${JSON.stringify(cur.headerText[i])}` : `${JSON.stringify(before)} → ${JSON.stringify(cur.headerText[i])}`]);
	}
	for (let i = 0; i < cur.footerText.length; i++) {
		const before = base.footerText[i];
		if (before !== cur.footerText[i]) lines.push([`  footer[${i}]`, before === undefined ? `(new) ${JSON.stringify(cur.footerText[i])}` : `${JSON.stringify(before)} → ${JSON.stringify(cur.footerText[i])}`]);
	}

	const fieldsAdded = [...cur.fields].filter((f) => !base.fields.has(f));
	const fieldsRemoved = [...base.fields].filter((f) => !cur.fields.has(f));
	if (fieldsAdded.length || fieldsRemoved.length) {
		lines.push(["fields", `${[...cur.fields].join(", ") || "(none)"}${fieldsAdded.length ? `   added: ${fieldsAdded.join(", ")}` : ""}${fieldsRemoved.length ? `   removed: ${fieldsRemoved.join(", ")}` : ""}`]);
	}
	// Compare styles by NAME, not styleId: a writer may renumber ids on every
	// run (this skill's own apply does), which would otherwise report every
	// style as both added and removed — pure noise in a report a human reads.
	const styleLabels = (a) => new Set([...a.styles].map((id) => a.styleIdToName.get(id) || id));
	const baseStyles = styleLabels(base);
	const curStyles = styleLabels(cur);
	const stylesAdded = [...curStyles].filter((s) => !baseStyles.has(s));
	if (stylesAdded.length) lines.push(["styles added", stylesAdded.join(", ")]);
	const stylesRemoved = [...baseStyles].filter((s) => !curStyles.has(s));
	if (stylesRemoved.length) lines.push(["styles removed", stylesRemoved.join(", ")]);

	const baseSet = new Set(base.paragraphTexts.map((t) => t.trim()).filter(Boolean));
	const newTexts = cur.paragraphTexts.map((t) => t.trim()).filter((t) => t && !baseSet.has(t));
	if (newTexts.length) {
		const shown = newTexts.slice(0, 12).map((t) => (t.length > 68 ? `${t.slice(0, 68)}…` : t));
		lines.push([`body text added (~${newTexts.length} para)`, shown.join("\n" + " ".repeat(24))]);
	}
	return lines;
}

// ── CLI ────────────────────────────────────────────────────────────────────

const argv = process.argv.slice(2);
const asJson = argv.includes("--json");
const quiet = argv.includes("--quiet");
let diffFlag = null;
const di = argv.indexOf("--diff");
if (di !== -1) diffFlag = argv[di + 1];
const target = argv.find((x) => !x.startsWith("--") && x !== diffFlag);

if (!target) {
	console.error("Usage: node scripts/verify.js <verify.json | file.docx> [--diff <baseline.docx>] [--json] [--quiet]");
	process.exit(1);
}

const targetAbs = resolve(target);
let config = null;
if (/\.json$/i.test(targetAbs)) {
	if (!existsSync(targetAbs)) {
		console.error(`verify config not found: ${targetAbs}`);
		process.exit(1);
	}
	try {
		config = JSON.parse(readFileSync(targetAbs, "utf8"));
	} catch (e) {
		console.error(`verify config is not valid JSON: ${e.message}`);
		process.exit(1);
	}
}
const baseDir = config ? dirname(targetAbs) : process.cwd();
const fileAbs = config && config.file ? (isAbsolute(config.file) ? config.file : resolve(baseDir, config.file)) : targetAbs;
const baselineRel = diffFlag || (config && config.baseline) || null;
const baselineAbs = baselineRel ? (isAbsolute(baselineRel) ? baselineRel : resolve(baseDir, baselineRel)) : null;

let cur;
try {
	cur = analyze(fileAbs);
} catch (e) {
	console.error(`verify: ${e.message}`);
	if (process.env.DOCX_MASTER_DEBUG) console.error(e.stack);
	process.exit(1);
}

const results = [];
if (config && Array.isArray(config.assert)) {
	for (const spec of config.assert) results.push(evalAssertion(cur, spec));
}

let base = null;
if (baselineAbs) {
	try {
		base = analyze(baselineAbs);
	} catch (e) {
		console.error(`verify: baseline unreadable — ${e.message}`);
		process.exit(1);
	}
}

const failed = results.filter((r) => !r.pass);

if (asJson) {
	console.log(JSON.stringify({
		file: fileAbs,
		baseline: baselineAbs,
		summary: { total: results.length, passed: results.length - failed.length, failed: failed.length },
		assertions: results,
		manifest: base ? manifest(base, cur) : null
	}, null, 2));
	process.exit(failed.length ? 1 : 0);
}

// ── report ─────────────────────────────────────────────────────────────────

if (!quiet) {
	console.log(`verify — ${fileAbs}`);
	console.log(`  sections ${cur.sectionCount} · paragraphs ${cur.paragraphCount} · tables ${cur.tableCount} · images ${cur.imageCount}`);
	console.log(`  header parts ${cur.headerText.length} · footer parts ${cur.footerText.length} · fields ${[...cur.fields].join(", ") || "(none)"}`);
}

if (base) {
	console.log(`\nCHANGE MANIFEST   ${baselineAbs}\n                  → ${fileAbs}`);
	console.log("─".repeat(78));
	for (const [k, v] of manifest(base, cur)) console.log(`  ${String(k).padEnd(22)} ${v}`);
	console.log("─".repeat(78));
	console.log("  Every line above is a change this run introduced.");
	console.log("  Read it against the user's own words: anything here they did not ask");
	console.log("  for is a defect, not a feature — remove it before delivering.\n");
}

if (results.length) {
	console.log(`ASSERTIONS (${results.length - failed.length}/${results.length} passed)`);
	console.log("─".repeat(78));
	for (const r of results) console.log(`  [${r.pass ? "PASS" : "FAIL"}] ${r.label}\n         ${r.detail}`);
	console.log("─".repeat(78));
}

if (failed.length) {
	console.log(`\n✗ ${failed.length} assertion(s) FAILED — the artifact does not meet the declared spec.`);
	console.log("  Delivery blocked. Fix the config and re-run the pipeline.");
	process.exit(1);
}
if (results.length) console.log(`\n✓ all ${results.length} assertion(s) passed.`);
