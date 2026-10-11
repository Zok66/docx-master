import { g as DocxReader, v as serializeXml, _ as parseXml, h as NS } from "./_shared/xml-utils.js";
import { mergeRunsInDocument, selectMergeParts } from "./_shared/merge-runs.js";
import { existsSync, mkdirSync, readFileSync } from "node:fs";
import { basename, extname, join, resolve } from "node:path";

//#region skill/tools/batch-fill-cli.ts
/**
 * `batch_fill` — one template + one data table → N finished documents.
 *
 * A mail-merge that does not need Word: every `{{field}}` in the template is
 * resolved from the matching table column, and one .docx comes out per row.
 *
 * Why it is a separate tool rather than a loop around `apply`: the hard part
 * is not the substitution, it is that a placeholder visible on the page is
 * usually **not** one contiguous string in the XML. Word splits it across
 * runs (`{{客` / `户` / `名}}`) because of revision ids, spell-check markers
 * and editing history, and a per-run `replace` silently misses all of those.
 * So the pipeline mirrors what `merge_runs` + `find_text` do by hand:
 *
 *   1. merge runs in every participating part (headers / footers / notes
 *      included — that is where contract numbers and page stamps usually
 *      live), ONCE, and cache the merged XML;
 *   2. then, per row, stitch whatever runs a placeholder still spans: the
 *      value lands in the run that holds the placeholder's first character
 *      and the runs that held the rest are emptied. No run is ever deleted,
 *      so rPr runs, bookmarks, fields and comment anchors around them keep
 *      their positions;
 *   3. write through `copyAndModify`, so every untouched part is copied
 *      byte-for-byte from the template.
 *
 * Rendering is unchanged and the template itself is never written to.
 *
 * Value handling: `\n` becomes `<w:br/>` and `\t` becomes `<w:tab/>` (a
 * multi-line address stays multi-line instead of collapsing into one run of
 * spaces). `xml:space="preserve"` is re-applied to every `<w:t>` whose text
 * carries edge whitespace, so leading / trailing spaces survive Word.
 *
 * Field syntax: `{{name}}`, whitespace inside the braces is ignored, and
 * `{{name|fallback}}` supplies the text used when the column is absent.
 * `{{_index}}` / `{{_row}}` are always available (1-based row number) — handy
 * for file naming when the table has no unique column.
 *
 * Usage:
 *   node scripts/batch_fill.js tpl.docx data.csv --out-dir out/
 *   node scripts/batch_fill.js tpl.docx data.csv --out-dir out/ --name "{{编号}}-{{客户名称}}.docx"
 *   node scripts/batch_fill.js tpl.docx data.json --out-dir out/ --name-field 客户名称
 *   node scripts/batch_fill.js tpl.docx data.csv --out-dir out/ --dry-run
 *   node scripts/batch_fill.js tpl.docx data.csv --out-dir out/ --validate
 */

const XML_NS = "http://www.w3.org/XML/1998/namespace";
const ELEMENT_NODE = 1;

/** XML whitespace only — a CJK ideographic space is *not* XML whitespace. */
const XML_SPACE_RE = /^[ \t\r\n]|[ \t\r\n]$/;

const DEFAULT_OPEN = "{{";
const DEFAULT_CLOSE = "}}";
const DEFAULT_NAME = "output-{{_index}}.docx";

function usage() {
	return [
		"Usage: node scripts/batch_fill.js <template.docx> <data.csv|data.json> --out-dir <dir> [options]",
		"",
		"One template + one data table -> one document per row (mail merge without Word).",
		"Placeholders: {{field}} · {{field|fallback}} · {{_index}} / {{_row}}",
		"",
		"Options:",
		"  -o, --out-dir <dir>    Output directory (created if missing). Required.",
		"      --name <pattern>   File name pattern, may use {{field}}.",
		`                         Default: ${DEFAULT_NAME}`,
		"      --name-field <col> Shorthand for --name \"{{<col>}}.docx\".",
		"      --parts <list>     Parts to fill: all (default) | headers,footers,notes.",
		"                         word/document.xml is always included.",
		"      --delims <a,b>     Placeholder delimiters. Default: {{,}}",
		"      --limit <n>        Process only the first n rows (debugging).",
		"      --strict           Fail when the template needs a column the table lacks.",
		"      --dry-run          Report the match-up; write nothing.",
		"      --validate         Schema-check every document produced.",
		"  -h, --help             Show this message."
	].join("\n");
}

function parseArgs(argv) {
	const opts = {
		template: null,
		data: null,
		outDir: null,
		name: null,
		nameField: null,
		parts: "all",
		open: DEFAULT_OPEN,
		close: DEFAULT_CLOSE,
		limit: null,
		strict: false,
		dryRun: false,
		validate: false,
		help: false
	};
	for (let i = 0; i < argv.length; i++) {
		const a = argv[i];
		if (a === "-h" || a === "--help") opts.help = true;
		else if (a === "-o" || a === "--out-dir") opts.outDir = argv[++i] ?? null;
		else if (a === "--name") opts.name = argv[++i] ?? null;
		else if (a === "--name-field") opts.nameField = argv[++i] ?? null;
		else if (a === "--parts") opts.parts = argv[++i] ?? "all";
		else if (a === "--delims") {
			const raw = argv[++i] ?? "";
			const idx = raw.indexOf(",");
			if (idx === -1) throw new Error("--delims expects two delimiters, e.g. --delims {{,}}");
			opts.open = raw.slice(0, idx);
			opts.close = raw.slice(idx + 1);
		} else if (a === "--limit") opts.limit = Number(argv[++i] ?? "");
		else if (a === "--strict") opts.strict = true;
		else if (a === "--dry-run") opts.dryRun = true;
		else if (a === "--validate") opts.validate = true;
		else if (a.startsWith("--")) throw new Error(`unknown flag: ${a}`);
		else if (opts.template === null) opts.template = a;
		else if (opts.data === null) opts.data = a;
		else throw new Error(`unexpected argument: ${a}`);
	}
	return opts;
}

//#region table loading

/** RFC 4180 enough: quoted fields, `""` escapes, CRLF, BOM, trailing newline. */
function parseCsv(text) {
	let src = text;
	if (src.charCodeAt(0) === 0xfeff) src = src.slice(1);

	const rows = [];
	let row = [];
	let field = "";
	let quoted = false;

	for (let i = 0; i < src.length; i++) {
		const ch = src[i];
		if (quoted) {
			if (ch === '"') {
				if (src[i + 1] === '"') {
					field += '"';
					i++;
				} else quoted = false;
			} else field += ch;
			continue;
		}
		if (ch === '"') {
			quoted = true;
			continue;
		}
		if (ch === ",") {
			row.push(field);
			field = "";
			continue;
		}
		if (ch === "\r") continue;
		if (ch === "\n") {
			row.push(field);
			rows.push(row);
			row = [];
			field = "";
			continue;
		}
		field += ch;
	}
	if (field !== "" || row.length > 0) {
		row.push(field);
		rows.push(row);
	}
	return rows.filter((r) => !(r.length === 1 && r[0].trim() === ""));
}

function rowsFromObjects(objects) {
	const columns = [];
	const seen = new Set();
	for (const obj of objects) {
		for (const key of Object.keys(obj)) {
			if (!seen.has(key)) {
				seen.add(key);
				columns.push(key);
			}
		}
	}
	const rows = objects.map((obj) => {
		const row = {};
		for (const col of columns) {
			const v = obj[col];
			row[col] = v === undefined || v === null ? "" : String(v);
		}
		return row;
	});
	return { columns, rows };
}

/** CSV -> { columns, rows:[{col:value}] }. Only the first row is a header. */
function loadTable(file) {
	const text = readFileSync(file, "utf8");
	const ext = extname(file).toLowerCase();

	if (ext === ".json") {
		let parsed;
		try {
			parsed = JSON.parse(text);
		} catch (err) {
			throw new Error(`${basename(file)} is not valid JSON: ${err.message}`);
		}
		const list = Array.isArray(parsed) ? parsed : Array.isArray(parsed?.rows) ? parsed.rows : null;
		if (!list) throw new Error(`${basename(file)} must be an array, or an object with a "rows" array`);
		if (list.length === 0) throw new Error(`${basename(file)} contains no rows`);
		if (typeof list[0] !== "object" || list[0] === null) {
			throw new Error(`${basename(file)} rows must be objects ({"column": "value"}), not bare values`);
		}
		return rowsFromObjects(list);
	}

	const table = parseCsv(text);
	if (table.length === 0) throw new Error(`${basename(file)} contains no rows`);
	const header = table[0].map((h, i) => (h.trim() === "" ? `column_${i + 1}` : h.trim()));
	const dupes = header.filter((h, i) => header.indexOf(h) !== i);
	if (dupes.length > 0) throw new Error(`duplicate column name(s) in ${basename(file)}: ${[...new Set(dupes)].join(", ")}`);
	const rows = table.slice(1).map((cells) => {
		const row = {};
		header.forEach((col, i) => {
			row[col] = cells[i] ?? "";
		});
		return row;
	});
	if (rows.length === 0) throw new Error(`${basename(file)} has a header but no data rows`);
	return { columns: header, rows };
}

//#endregion

//#region placeholder resolution

/** `{{name|fallback}}` -> { key: "name", fallback: "fallback" }. */
function parseKey(raw) {
	const s = raw.trim();
	const bar = s.indexOf("|");
	if (bar === -1) return { key: s, fallback: null };
	return { key: s.slice(0, bar).trim(), fallback: s.slice(bar + 1) };
}

/**
 * Column lookup, most specific first: exact, trimmed, case-insensitive.
 * Returns `undefined` when no such column exists at all (as opposed to `""`,
 * which is a column that exists but is empty for this row).
 */
function lookup(row, columns, key) {
	if (Object.prototype.hasOwnProperty.call(row, key)) return row[key];
	const trimmed = key.trim();
	if (Object.prototype.hasOwnProperty.call(row, trimmed)) return row[trimmed];
	const lower = trimmed.toLowerCase();
	const hit = columns.find((c) => c.toLowerCase() === lower);
	return hit === undefined ? undefined : row[hit];
}

function placeholderRegex(opts) {
	return new RegExp(`${escapeRegExp(opts.open)}([\\s\\S]*?)${escapeRegExp(opts.close)}`, "g");
}

function escapeRegExp(s) {
	return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/** Every distinct placeholder key in a run of plain text. */
function collectKeys(text, opts, out) {
	for (const m of text.matchAll(placeholderRegex(opts))) {
		const { key } = parseKey(m[1]);
		if (key !== "" && key !== "_index" && key !== "_row") out.add(key);
	}
	return out;
}

/**
 * Placeholder inventory of a part. Scanned over the *rendered* text of each
 * paragraph, not over the raw XML: a placeholder split across runs would
 * otherwise be read as `{{公</w:t></w:r>...`, inventing a field name out of
 * markup and hiding the real one.
 */
function collectPartKeys(doc, opts, out) {
	const root = doc.documentElement;
	if (!root) return out;
	for (const p of collectParagraphs(root)) {
		const nodes = collectTextNodes(p);
		if (nodes.length === 0) continue;
		collectKeys(
			nodes.map((n) => n.textContent ?? "").join(""),
			opts,
			out
		);
	}
	return out;
}

//#endregion

//#region xml mutation

function isW(node, local) {
	return node.nodeType === ELEMENT_NODE && node.namespaceURI === NS.w && node.localName === local;
}

/** All `w:p` in the part, outermost included. A nested `w:p` (textbox inside
 * a shape) is returned too, but its text is *not* counted as its parent's — so
 * each paragraph is filled exactly once. */
function collectParagraphs(root, out = []) {
	if (isW(root, "p")) out.push(root);
	for (let c = root.firstChild; c; c = c.nextSibling) {
		if (c.nodeType === ELEMENT_NODE) collectParagraphs(c, out);
	}
	return out;
}

/**
 * Text-bearing `w:t` children of a paragraph, in document order. Descends
 * through `w:r`, `w:hyperlink`, `w:ins`, `w:sdt` wrappers. Skips
 * `w:delText` (deleted text is not visible) and `w:instrText` (field codes
 * must never be rewritten), and does not cross into a nested paragraph.
 */
function collectTextNodes(p, out = []) {
	const walk = (node, isRoot) => {
		for (let c = node.firstChild; c; c = c.nextSibling) {
			if (c.nodeType !== ELEMENT_NODE) continue;
			if (c.namespaceURI === NS.w) {
				if (!isRoot && c.localName === "p") continue;
				if (c.localName === "t") {
					out.push(c);
					continue;
				}
				if (c.localName === "delText" || c.localName === "instrText") continue;
			}
			walk(c, false);
		}
	};
	walk(p, true);
	return out;
}

/** Set a `w:t`'s text, keeping `xml:space` in step with its edge whitespace. */
function setText(node, text) {
	while (node.firstChild) node.removeChild(node.firstChild);
	if (text !== "") node.appendChild(node.ownerDocument.createTextNode(text));
	if (text !== "" && XML_SPACE_RE.test(text)) node.setAttributeNS(XML_NS, "xml:space", "preserve");
	else node.removeAttributeNS(XML_NS, "space");
}

/** `a\nb\tc` -> [{text:"a"},{br:true},{text:"b"},{tab:true},{text:"c"}] */
function splitValue(value) {
	const pieces = [];
	const lines = String(value).split(/\r\n|\r|\n/);
	lines.forEach((line, li) => {
		if (li > 0) pieces.push({ br: true });
		line.split("\t").forEach((chunk, ci) => {
			if (ci > 0) pieces.push({ tab: true });
			if (chunk !== "") pieces.push({ text: chunk });
		});
	});
	return pieces;
}

/**
 * Write `head + value + tail` starting at `node`, turning embedded newlines
 * and tabs into real `w:br` / `w:tab` siblings. `node` is always the first
 * `w:t` of the placeholder; extra elements are inserted right after it inside
 * the same run.
 */
function writeValue(node, head, value, tail) {
	const doc = node.ownerDocument;
	const run = node.parentNode;
	const pieces = splitValue(value);

	let anchor = node;
	let textElem = node;
	let pending = head;

	const flush = () => {
		setText(textElem, pending);
		pending = "";
	};

	for (const piece of pieces) {
		if (piece.br || piece.tab) {
			flush();
			const el = doc.createElementNS(NS.w, piece.br ? "w:br" : "w:tab");
			run.insertBefore(el, anchor.nextSibling);
			anchor = el;
			const fresh = doc.createElementNS(NS.w, "w:t");
			run.insertBefore(fresh, anchor.nextSibling);
			anchor = fresh;
			textElem = fresh;
		} else {
			pending += piece.text;
		}
	}
	setText(textElem, pending + tail);
}

/** Locate the text node holding char `start` and the one holding `end - 1`. */
function locate(texts, start, end) {
	let acc = 0;
	let i0 = -1;
	let l0 = 0;
	for (let i = 0; i < texts.length; i++) {
		const len = texts[i].length;
		if (len === 0) continue;
		if (i0 === -1 && start < acc + len) {
			i0 = i;
			l0 = start - acc;
		}
		if (end - 1 < acc + len) {
			if (i0 === -1) return null;
			return { i0, l0, i1: i, l1: end - 1 - acc };
		}
		acc += len;
	}
	return null;
}

/**
 * Fill every placeholder in one paragraph. Matches are applied back-to-front
 * so the offsets of the ones still pending never shift.
 */
function fillParagraph(p, resolve, opts, stats) {
	const nodes = collectTextNodes(p);
	if (nodes.length === 0) return;
	const texts = nodes.map((n) => n.textContent ?? "");
	if (!texts.some((t) => t.includes(opts.open))) return;

	const full = texts.join("");
	const matches = [];
	for (const m of full.matchAll(placeholderRegex(opts))) {
		const { key, fallback } = parseKey(m[1]);
		if (key === "") continue;
		matches.push({ start: m.index, end: m.index + m[0].length, key, fallback });
	}
	if (matches.length === 0) return;

	for (let i = matches.length - 1; i >= 0; i--) {
		const { start, end, key, fallback } = matches[i];
		const value = resolve(key, fallback);
		if (value === undefined) {
			stats.missing.set(key, (stats.missing.get(key) ?? 0) + 1);
			continue;
		}
		const span = locate(texts, start, end);
		if (!span) continue;
		const { i0, l0, i1, l1 } = span;
		const head = texts[i0].slice(0, l0);
		const tail = texts[i1].slice(l1 + 1);

		if (i0 === i1) {
			writeValue(nodes[i0], head, value, tail);
			texts[i0] = `${head}${value}${tail}`;
		} else {
			writeValue(nodes[i0], head, value, "");
			for (let k = i0 + 1; k < i1; k++) setText(nodes[k], "");
			setText(nodes[i1], tail);
			texts[i0] = `${head}${value}`;
			for (let k = i0 + 1; k < i1; k++) texts[k] = "";
			texts[i1] = tail;
		}
		stats.replaced += 1;
	}
}

/** Fill a whole part. Returns { replaced, missing }. */
function fillPart(doc, resolve, opts) {
	const stats = { replaced: 0, missing: new Map() };
	const root = doc.documentElement;
	if (!root) return stats;
	for (const p of collectParagraphs(root)) fillParagraph(p, resolve, opts, stats);
	return stats;
}

//#endregion

//#region naming

function sanitizeFileName(name) {
	// eslint-disable-next-line no-control-regex
	const cleaned = name
		.replace(/[\\/:*?"<>|\u0000-\u001f]/g, "_")
		.replace(/\s+/g, " ")
		.trim();
	return cleaned === "" || cleaned === ".docx" ? "" : cleaned.slice(0, 120);
}

function renderName(pattern, row, columns, index) {
	const rendered = pattern.replace(placeholderRegex({ open: DEFAULT_OPEN, close: DEFAULT_CLOSE }), (all, raw) => {
		const { key, fallback } = parseKey(raw);
		if (key === "_index" || key === "_row") return String(index);
		const value = lookup(row, columns, key);
		if (value === undefined) return fallback ?? "";
		return value;
	});
	let name = sanitizeFileName(rendered);
	if (name === "") name = `output-${index}.docx`;
	if (!/\.(docx|dotx)$/i.test(name)) name = `${name}.${extname(pattern).replace(".", "") || "docx"}`;
	return name;
}

//#endregion

async function main() {
	let opts;
	try {
		opts = parseArgs(process.argv.slice(2));
	} catch (err) {
		console.error(err.message);
		console.error(usage());
		process.exit(1);
	}

	if (opts.help) {
		console.log(usage());
		return;
	}
	if (!opts.template || !opts.data) {
		console.error("Error: a template and a data table are both required");
		console.error(usage());
		process.exit(1);
	}
	if (!opts.dryRun && !opts.outDir) {
		console.error("Error: --out-dir is required (use --dry-run to preview without writing)");
		process.exit(1);
	}

	const template = resolve(opts.template);
	const dataFile = resolve(opts.data);
	if (!existsSync(template)) {
		console.error(`Error: template not found: ${template}`);
		process.exit(1);
	}
	const ext = extname(template).toLowerCase();
	if (ext !== ".docx" && ext !== ".dotx") {
		console.error(`Error: ${basename(template)} is not a .docx / .dotx file (convert a legacy .doc first)`);
		process.exit(1);
	}
	if (!existsSync(dataFile)) {
		console.error(`Error: data table not found: ${dataFile}`);
		process.exit(1);
	}

	let table;
	try {
		table = loadTable(dataFile);
	} catch (err) {
		console.error(`Error: ${err.message}`);
		process.exit(1);
	}
	const rows = opts.limit != null && Number.isFinite(opts.limit) ? table.rows.slice(0, opts.limit) : table.rows;
	if (rows.length === 0) {
		console.error("Error: no data rows to process");
		process.exit(1);
	}

	const namePattern =
		opts.name ?? (opts.nameField ? `${DEFAULT_OPEN}${opts.nameField}${DEFAULT_CLOSE}.docx` : DEFAULT_NAME);

	try {
		const reader = await DocxReader.open(template);
		const parts = selectMergeParts(reader.listEntries(), opts.parts);
		if (parts.length === 0) {
			console.error("Error: no fillable part found (word/document.xml missing?)");
			process.exit(1);
		}

		// Merge once, then reuse the merged XML for every row. Merging is the
		// expensive half and it does not depend on the row data.
		const merged = new Map();
		let mergedRuns = 0;
		for (const part of parts) {
			const doc = await reader.readXml(part);
			if (!doc) continue;
			mergedRuns += mergeRunsInDocument(doc);
			merged.set(part, serializeXml(doc));
		}

		const templateKeys = new Set();
		for (const xml of merged.values()) collectPartKeys(parseXml(xml), opts, templateKeys);

		const missingColumns = [...templateKeys].filter((k) => lookup(rows[0], table.columns, k) === undefined);
		const unusedColumns = table.columns.filter((c) => !templateKeys.has(c));

		console.log(`Template  : ${basename(template)}`);
		console.log(`Parts     : ${[...merged.keys()].map((p) => p.replace(/^word\//, "")).join(", ")}`);
		console.log(`Runs      : merged ${mergedRuns}`);
		console.log(`Fields    : ${templateKeys.size} placeholder(s) in template, ${table.columns.length} column(s) in data`);
		if (templateKeys.size > 0) console.log(`            ${[...templateKeys].sort().join(", ")}`);
		console.log(`Rows      : ${rows.length}`);
		if (templateKeys.size === 0) console.log("  ! no placeholder found — check --delims and the template text");
		if (missingColumns.length > 0) {
			console.log(`Missing   : ${missingColumns.join(", ")} (no such column${opts.strict ? " — fatal" : ", left as-is"})`);
		}
		if (unusedColumns.length > 0) console.log(`Unused    : ${unusedColumns.join(", ")}`);
		if (missingColumns.length > 0 && opts.strict) {
			console.error("Error: --strict and the template needs columns the data does not have");
			process.exit(1);
		}

		if (opts.dryRun) {
			console.log("");
			console.log("Dry run — nothing written. File names would be:");
			const preview = rows.slice(0, 10).map((row, i) => `  ${renderName(namePattern, row, table.columns, i + 1)}`);
			console.log(preview.join("\n"));
			if (rows.length > 10) console.log(`  ... and ${rows.length - 10} more`);
			return;
		}

		const outDir = resolve(opts.outDir);
		if (!existsSync(outDir)) mkdirSync(outDir, { recursive: true });

		const usedNames = new Set();
		const warnings = new Map();
		let written = 0;
		let totalReplaced = 0;
		let validated = 0;
		const validationErrors = [];
		const validateDoc = opts.validate ? (await import("./_shared/docx-validate.js")).t : null;

		for (let i = 0; i < rows.length; i++) {
			const row = rows[i];
			const index = i + 1;
			const resolve = (key, fallback) => {
				if (key === "_index" || key === "_row") return String(index);
				const value = lookup(row, table.columns, key);
				if (value !== undefined) return value;
				return fallback === null ? undefined : fallback;
			};

			const replacements = new Map();
			for (const [part, xml] of merged) {
				const doc = parseXml(xml);
				const stats = fillPart(doc, resolve, opts);
				totalReplaced += stats.replaced;
				for (const [key, count] of stats.missing) warnings.set(key, (warnings.get(key) ?? 0) + count);
				replacements.set(part, serializeXml(doc));
			}

			let fileName = renderName(namePattern, row, table.columns, index);
			if (usedNames.has(fileName.toLowerCase())) {
				const dot = fileName.lastIndexOf(".");
				fileName = `${fileName.slice(0, dot)}-${index}${fileName.slice(dot)}`;
			}
			usedNames.add(fileName.toLowerCase());

			const outPath = join(outDir, fileName);
			await reader.copyAndModify(outPath, replacements);
			written += 1;

			if (validateDoc) {
				const errors = await validateDoc(outPath);
				validated += 1;
				if (errors.length > 0) {
					validationErrors.push({ file: fileName, errors });
				}
			}
			process.stdout.write(`\r  writing ${written}/${rows.length} ...`);
		}
		process.stdout.write("\r");

		console.log("");
		console.log(`Wrote ${written} document(s) to ${outDir}`);
		console.log(`Replacements: ${totalReplaced}`);
		if (warnings.size > 0) {
			const list = [...warnings.entries()].map(([k, c]) => `${k} (${c}x)`).join(", ");
			console.log(`Left unfilled (no column): ${list}`);
		}
		if (validateDoc) {
			if (validationErrors.length === 0) {
				console.log(`Validation: ${validated} document(s) checked, all PASSED`);
			} else {
				for (const { file, errors } of validationErrors) {
					console.error(`Validation FAILED: ${file}`);
					for (const e of errors) console.error(`  ${e.part}: ${e.message}`);
				}
				process.exit(1);
			}
		}
	} catch (err) {
		console.error(`Error: ${err.message}`);
		process.exit(1);
	}
}

main();

//#endregion
export {};
