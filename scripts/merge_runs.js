import { g as DocxReader, v as serializeXml } from "./_shared/xml-utils.js";
import { mergeRunsInDocument, selectMergeParts } from "./_shared/merge-runs.js";
import { existsSync, renameSync, unlinkSync } from "node:fs";
import { basename, dirname, resolve } from "node:path";

//#region skill/tools/merge-runs-cli.ts
/**
 * `merge_runs` — coalesce adjacent identically-formatted runs in a .docx.
 *
 * Run this before a surgical edit on a document of unknown provenance: it
 * makes the text findable again for `find_text` / `set-run`. Word fragments a
 * paragraph's text across many `<w:r>` elements (revision ids, spell-check
 * markers), so a phrase visible on the page is often not a contiguous string
 * in `word/document.xml`.
 *
 * Rendering is unchanged; tracked-change structure is preserved (runs inside
 * different `<w:ins>` / `<w:del>` wrappers are never joined). Full semantics:
 * `_shared/merge-runs.js`.
 *
 * Unlike the Python companion (`pro/scripts/merge_runs.py`, document.xml only),
 * `--parts` can also cover headers, footers, footnotes and endnotes.
 *
 * Usage:
 *   node scripts/merge_runs.js doc.docx                    # rewrite in place
 *   node scripts/merge_runs.js doc.docx -o merged.docx     # write elsewhere
 *   node scripts/merge_runs.js doc.docx --dry-run          # report only
 *   node scripts/merge_runs.js doc.docx --parts all        # + headers/footers/notes
 *   node scripts/merge_runs.js doc.docx --parts headers,notes
 */

function usage() {
	return [
		"Usage: node scripts/merge_runs.js <file.docx> [options]",
		"",
		"Coalesce adjacent identically-formatted runs so text becomes findable.",
		"Run this before a surgical edit on a document of unknown provenance.",
		"",
		"Options:",
		"  -o, --output <path>   Write to <path> instead of rewriting the input.",
		"      --parts <list>     Extra parts to process, comma-separated:",
		"                         headers, footers, notes, all. word/document.xml",
		"                         is always processed.",
		"      --dry-run          Report what would merge; write nothing.",
		"  -h, --help             Show this message."
	].join("\n");
}

function parseArgs(argv) {
	const opts = { input: null, output: null, parts: null, dryRun: false, help: false };
	for (let i = 0; i < argv.length; i++) {
		const a = argv[i];
		if (a === "-h" || a === "--help") opts.help = true;
		else if (a === "-o" || a === "--output") opts.output = argv[++i] ?? null;
		else if (a === "--parts") opts.parts = argv[++i] ?? null;
		else if (a === "--dry-run") opts.dryRun = true;
		else if (a.startsWith("--")) throw new Error(`unknown flag: ${a}`);
		else if (opts.input === null) opts.input = a;
		else throw new Error(`unexpected argument: ${a}`);
	}
	return opts;
}

async function main() {
	let opts;
	try {
		opts = parseArgs(process.argv.slice(2));
	} catch (err) {
		console.error(err.message);
		console.error(usage());
		process.exit(1);
	}

	if (opts.help || !opts.input) {
		console.log(usage());
		process.exit(opts.help ? 0 : 1);
	}

	const input = resolve(opts.input);
	if (!existsSync(input)) {
		console.error(`Error: ${input} not found`);
		process.exit(1);
	}
	const ext = input.slice(input.lastIndexOf(".")).toLowerCase();
	if (ext !== ".docx" && ext !== ".dotx") {
		console.error(`Error: ${basename(input)} is not a .docx / .dotx file`);
		process.exit(1);
	}

	const output = opts.output ? resolve(opts.output) : input;
	if (output !== input && !existsSync(dirname(output))) {
		console.error(`Error: output directory does not exist: ${dirname(output)}`);
		process.exit(1);
	}

	try {
		const reader = await DocxReader.open(input);
		const parts = selectMergeParts(reader.listEntries(), opts.parts);
		if (parts.length === 0) {
			console.error("Error: no mergeable part found (word/document.xml missing?)");
			process.exit(1);
		}

		const replacements = new Map();
		const lines = [];
		let total = 0;
		for (const part of parts) {
			const doc = await reader.readXml(part);
			if (!doc) continue;
			const count = mergeRunsInDocument(doc);
			replacements.set(part, serializeXml(doc));
			total += count;
			lines.push(`  ${part}: merged ${count} run(s)`);
		}

		console.log(`Parts processed: ${replacements.size}`);
		console.log(lines.join("\n"));
		console.log(`Total: ${total} run(s) merged.`);

		if (opts.dryRun) {
			console.log("(dry run — nothing written)");
			return;
		}

		if (output === input) {
			// Write beside the target, then swap — a partially written zip
			// would leave the user with an unopenable document.
			const staging = `${input}.merge-tmp-${process.pid}`;
			try {
				await reader.copyAndModify(staging, replacements);
				renameSync(staging, input);
			} finally {
				if (existsSync(staging)) unlinkSync(staging);
			}
			console.log(`Rewrote ${basename(input)} in place.`);
		} else {
			await reader.copyAndModify(output, replacements);
			console.log(`Wrote ${output}`);
		}
	} catch (err) {
		console.error(`Error: ${err.message}`);
		process.exit(1);
	}
}

main();

//#endregion
export { };
