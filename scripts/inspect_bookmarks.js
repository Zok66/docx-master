#!/usr/bin/env node
/**
 * `inspect_bookmarks` — bookmarks, cross-references, and the references that
 * point at nothing.
 *
 * A bookmark is invisible. Every other defect a document can carry shows up
 * somewhere — a stray footer title is visible, a broken style shifts the text,
 * a lost part fails validation. A dangling reference renders as ordinary text
 * in Word ("Error! Bookmark not defined." inside a sentence, a TOC entry that
 * stops jumping) and passes every schema check on the way. This is the tool
 * that looks at it.
 *
 * Reports, per package:
 *   - bookmarks: name, w:id, part, paragraph ordinal, origin (Word-generated
 *     vs authored) — so a tool-written anchor is distinguishable at a glance
 *     from the 70 `_Toc*` names Word regenerates on every field update
 *   - references: REF / PAGEREF / NOTEREF, `HYPERLINK \l`, `TOC \b`, and
 *     structural `w:hyperlink w:anchor`, each with its resolution status
 *   - problems: dangling references, unclosed / unopened bookmarks, duplicate
 *     ids or names, bookmarks whose range never closes
 *
 * Usage:
 *   node scripts/inspect_bookmarks.js <file.docx> [--json] [--problems-only]
 *                                     [--exit-zero] [--limit N]
 *
 * Exit code: 0 clean · 1 a reference or bookmark is broken (unless --exit-zero)
 */
import { existsSync } from "node:fs";
import { basename, resolve, isAbsolute } from "node:path";
import { openZip, referenceBearingParts } from "./_shared/opc.js";
import { analyzeParts, originOf, ORIGIN_LABEL, isWordGenerated } from "./_shared/bookmarks.js";

/** Problems that make a document quietly wrong. `unclosed-field` is advisory:
 * a field opened in one part and closed in another is unusual but legal. */
const ERROR_KINDS = new Set([
	"dangling",
	"unclosed",
	"unopened",
	"duplicate-id",
	"duplicate-name",
	"empty-name",
	"empty-target"
]);

const LINE = "─".repeat(76);

function parseArgs(argv) {
	const out = { file: "", json: false, problemsOnly: false, exitZero: false, limit: 40 };
	for (let i = 0; i < argv.length; i++) {
		const a = argv[i];
		if (a === "--json") out.json = true;
		else if (a === "--problems-only") out.problemsOnly = true;
		else if (a === "--exit-zero") out.exitZero = true;
		else if (a === "--limit") out.limit = Number(argv[++i]) || out.limit;
		else if (a === "-h" || a === "--help") out.help = true;
		else if (!a.startsWith("-")) out.file = a;
	}
	return out;
}

/** Collapse the 70 identical PAGEREF/HYPERLINK pairs a TOC produces into one
 * line per (kind, target) with a hit count and the first location. */
function groupRefs(refs) {
	const map = new Map();
	for (const r of refs) {
		const key = `${r.kind}\u0000${r.target}`;
		const hit = map.get(key);
		if (hit) {
			hit.count++;
			hit.parts.add(r.part);
		} else {
			map.set(key, { kind: r.kind, target: r.target, count: 1, part: r.part, para: r.para, resolved: r.resolved, parts: new Set([r.part]) });
		}
	}
	return [...map.values()].sort((a, b) => (a.resolved === b.resolved ? b.count - a.count : a.resolved ? 1 : -1));
}

function reportText(scan, file, opt) {
	const L = [];
	const name = basename(file);
	const both = (label, n, detail) => L.push(`${label.padEnd(14)} ${String(n).padStart(4)}   ${detail || ""}`);

	L.push(`bookmarks — ${name}`);
	L.push(LINE);

	const originDetail = Object.entries(scan.countByOrigin)
		.sort((a, b) => b[1] - a[1])
		.map(([k, n]) => `${ORIGIN_LABEL[k] || k} ${n}`)
		.join(", ");
	both("bookmarks", scan.bookmarks.length, originDetail);

	const refDetail = Object.entries(scan.refKinds).map(([k, n]) => `${k} ${n}`).join(", ");
	both("references", scan.refs.length, refDetail);

	const custom = scan.bookmarks.filter((b) => !isWordGenerated(b.name));
	const referenced = new Set(scan.refs.map((r) => r.target));
	const unreferenced = custom.filter((b) => !referenced.has(b.name));
	both("dangling", scan.dangling.length, scan.dangling.length ? "references pointing at nothing" : "");
	both("problems", scan.problems.length, scan.problems.length ? "see below" : "");
	if (custom.length) {
		both("authored", custom.length, `bookmarks you (or a tool) wrote; ${unreferenced.length} unreferenced`);
	}

	if (!opt.problemsOnly) {
		for (const [part, counts] of Object.entries(scan.byPart)) {
			if (!counts.bookmarks && !counts.refs) continue;
			L.push("");
			L.push(`${part}  (${counts.bookmarks} bookmarks, ${counts.refs} references)`);
			const rows = scan.bookmarks.filter((b) => b.part === part);
			if (rows.length) {
				L.push(`  ${"name".padEnd(22)} ${"id".padStart(4)} ${"para".padStart(5)}  origin`);
				for (const b of rows.slice(0, opt.limit)) {
					L.push(`  ${b.name.padEnd(22)} ${String(b.id).padStart(4)} ${String(b.para).padStart(5)}  ${ORIGIN_LABEL[originOf(b.name)]}`);
				}
				if (rows.length > opt.limit) L.push(`  … ${rows.length - opt.limit} more`);
			}
		}

		const groups = groupRefs(scan.refs);
		if (groups.length) {
			L.push("");
			L.push("references");
			L.push(`  ${"kind".padEnd(14)} ${"target".padEnd(24)} ${"part".padEnd(18)} ${"para".padStart(4)}  status`);
			for (const g of groups.slice(0, opt.limit)) {
				const part = g.parts.size > 1 ? `${[...g.parts].length} parts` : g.part;
				const suffix = g.count > 1 ? `×${g.count}` : "";
				L.push(`  ${g.kind.padEnd(14)} ${g.target.padEnd(24)} ${part.padEnd(18)} ${String(g.para).padStart(4)}  ${g.resolved ? "ok" : "DANGLING"} ${suffix}`.trimEnd());
			}
			if (groups.length > opt.limit) L.push(`  … ${groups.length - opt.limit} more distinct references`);
		}
	}

	L.push("");
	if (scan.problems.length) {
		L.push("problems");
		for (const p of scan.problems) L.push(`  [${p.kind}] ${p.part} p#${p.para} — ${p.detail}`);
	} else {
		L.push("problems: none — every reference resolves.");
	}

	L.push(LINE);
	const fatal = scan.problems.filter((p) => ERROR_KINDS.has(p.kind));
	L.push(fatal.length
		? `${fatal.length} broken reference(s) / bookmark(s). Word will render these as "Error! Bookmark not defined." or a dead link after the next field update.`
		: "clean: no dangling references, no malformed bookmarks.");
	return L.join("\n");
}

function main(argv) {
	const opt = parseArgs(argv);
	if (opt.help || !opt.file) {
		console.log("usage: node scripts/inspect_bookmarks.js <file.docx> [--json] [--problems-only] [--exit-zero] [--limit N]");
		return opt.file ? 0 : 2;
	}
	const file = isAbsolute(opt.file) ? opt.file : resolve(process.cwd(), opt.file);
	if (!existsSync(file)) {
		console.error(`inspect_bookmarks: file not found: ${file}`);
		return 1;
	}

	const zip = openZip(file);
	const parts = referenceBearingParts(zip);
	const scan = analyzeParts(parts);

	if (opt.json) {
		console.log(JSON.stringify({
			file,
			bookmarks: scan.bookmarks.map((b) => ({ ...b, origin: originOf(b.name) })),
			references: scan.refs,
			problems: scan.problems,
			counts: {
				bookmarks: scan.bookmarks.length,
				references: scan.refs.length,
				dangling: scan.dangling.length,
				byOrigin: scan.countByOrigin,
				byRefKind: scan.refKinds
			}
		}, null, 2));
	} else {
		console.log(reportText(scan, file, opt));
	}

	const fatal = scan.problems.filter((p) => ERROR_KINDS.has(p.kind));
	return fatal.length && !opt.exitZero ? 1 : 0;
}

process.exit(main(process.argv.slice(2)));
