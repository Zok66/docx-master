#!/usr/bin/env node
/**
 * `pipeline` — single-entry rebuild for multi-step docx-master chains.
 *
 * Why this exists: a real task is rarely one command. Producing a document
 * usually means *generate config → apply → extend → validate*, and every step
 * depends on the previous one. Run by hand, the chain breaks in a way no
 * validator can catch: regenerate the config *source* but forget to re-emit
 * the config JSON, and the writer happily writes the OLD config while every
 * downstream check reports PASS. The failure is silent; the artifact looks
 * fine until a human reads the page.
 *
 * So the chain becomes one command with no skippable links:
 *   - steps run in declared order with a single exit code;
 *   - the first failure ABORTS the run (there is no "continue anyway");
 *   - schema validation is appended automatically unless already declared,
 *     so a build cannot quietly finish unvalidated;
 *   - `verify` assertions run last when declared — a failed assertion fails
 *     the build.
 *
 * Three modes (all describe the same chain):
 *   (default)        execute the steps here, sequentially.
 *   --plan           print the resolved command sequence; execute nothing.
 *   --emit-script    write a `set -e` bash script that runs the chain, so a
 *                    single `bash <script>` gives the same abort-on-first-
 *                    failure semantics. Use this when the caller cannot spawn
 *                    child processes (sandboxed agents), or when the chain
 *                    should be re-runnable by a human.
 *
 * Usage:
 *   node scripts/pipeline.js <pipeline.json> [--plan] [--emit-script [path]]
 *                                            [--dry-run] [--no-auto-validate]
 *                                            [--python <exe>] [--quiet]
 *
 * pipeline.json:
 * {
 *   "cwd": ".",                    // base for relative paths (default: manifest dir)
 *   "steps": [
 *     { "run": "node build_config.js", "desc": "emit config.json from content source" },
 *     { "apply": "config.json", "out": "build/body.docx" },
 *     { "extend": "extend_config.json" },
 *     { "validate": "out.docx" },
 *     { "proValidate": "out.docx", "original": "src.docx" },
 *     { "run": "node post.js", "optional": true }
 *   ],
 *   "deliver": "out.docx",         // handed to the user; auto-validated if no validate step
 *   "verify": "verify.json"        // optional assertions gate (scripts/verify.js)
 * }
 *
 * Exit codes: 0 chain completed clean · 1 a step/assertion failed ·
 *             2 manifest problem · 3 deferred — script emitted, nothing executed.
 */
import { readFileSync, existsSync, writeFileSync, readdirSync } from "node:fs";
import { spawnSync } from "node:child_process";
import { dirname, resolve, isAbsolute, join } from "node:path";
import { fileURLToPath } from "node:url";

const HERE = dirname(fileURLToPath(import.meta.url)); // scripts/
const SKILL_ROOT = resolve(HERE, "..");
const NODE = process.execPath;

// ── environment capability (declared early: picked up below) ───────────────

/** Can this process start child processes at all? Sandboxes sometimes say no. */
function canSpawn() {
	const r = spawnSync(NODE, ["-e", "process.exit(0)"], { encoding: "utf8", shell: false });
	return !r.error && r.status === 0;
}

function probeExe(exe, args) {
	try {
		const r = spawnSync(exe, args, { encoding: "utf8", shell: false, timeout: 15000 });
		return !r.error && r.status === 0;
	} catch {
		return false;
	}
}

/** Interpreters that exist on disk in managed locations (no spawning needed). */
function managedPythons() {
	const out = [];
	const home = process.env.USERPROFILE || process.env.HOME;
	if (!home) return out;
	const envs = join(home, ".workbuddy", "binaries", "python", "envs");
	if (!existsSync(envs)) return out;
	try {
		for (const e of readdirSync(envs)) {
			for (const tail of [["Scripts", "python.exe"], ["bin", "python"]]) {
				const p = join(envs, e, ...tail);
				if (existsSync(p)) out.push(p);
			}
		}
	} catch { /* ignore */ }
	return out;
}

/**
 * Which python should run `extend` / the PRO validator?
 *
 * A bare `python` is not a safe default: it may resolve to an interpreter
 * without lxml/defusedxml while a complete environment sits right next to it
 * (observed on this machine — `python` → 3.13.x base, deps live in
 * envs/default). So: prefer whatever actually imports the deps, then any
 * interpreter, then a managed env on disk, then the bare name.
 */
function pickPython() {
	if (process.env.DOCX_MASTER_PYTHON) return process.env.DOCX_MASTER_PYTHON;
	const managed = managedPythons();
	const all = ["python", "python3", ...managed];
	if (canSpawn()) {
		for (const c of all) if (probeExe(c, ["-c", "import lxml, defusedxml"])) return c;
		for (const c of all) if (probeExe(c, ["--version"])) return c;
	} else {
		// Cannot probe — decide by what is on disk.
		const preferred = managed.find((p) => /envs[\\/]default[\\/]/.test(p));
		if (preferred) return preferred;
		if (managed.length) return managed[0];
	}
	return "python";
}

const argv = process.argv.slice(2);
const flags = {
	plan: argv.includes("--plan"),
	dryRun: argv.includes("--dry-run"),
	noAutoValidate: argv.includes("--no-auto-validate"),
	quiet: argv.includes("--quiet"),
	emit: argv.includes("--emit-script")
};
let pythonExe = null;
const pyFlag = argv.indexOf("--python");
if (pyFlag !== -1 && argv[pyFlag + 1]) pythonExe = argv[pyFlag + 1];
else pythonExe = pickPython();
let emitPath = null;
if (flags.emit) {
	const ei = argv.indexOf("--emit-script");
	const next = argv[ei + 1];
	if (next && !next.startsWith("--") && /\.(sh|bash)$/i.test(next)) emitPath = next;
}
const manifestPath = argv.find((a) => !a.startsWith("--") && a !== pythonExe && a !== emitPath);

if (!manifestPath) {
	console.error("Usage: node scripts/pipeline.js <pipeline.json> [--plan] [--emit-script [out.sh]] [--dry-run] [--no-auto-validate] [--python <exe>] [--quiet]");
	process.exit(2);
}

const C = process.stdout.isTTY
	? { dim: "\x1b[2m", red: "\x1b[31m", green: "\x1b[32m", yellow: "\x1b[33m", bold: "\x1b[1m", off: "\x1b[0m" }
	: { dim: "", red: "", green: "", yellow: "", bold: "", off: "" };
const say = (...a) => {
	if (!flags.quiet) console.log(...a);
};

const abs = (p, base) => (isAbsolute(p) ? p : resolve(base, p));
const q = (s) => (/[\s"'$`\\()]/.test(s) ? `"${String(s).replace(/"/g, '\\"')}"` : String(s));
const shq = (s) => `'${String(s).replace(/'/g, `'\\''`)}'`;

// ── manifest ───────────────────────────────────────────────────────────────

const manifestAbs = resolve(manifestPath);
if (!existsSync(manifestAbs)) {
	console.error(`pipeline manifest not found: ${manifestAbs}`);
	process.exit(2);
}
let manifest;
try {
	manifest = JSON.parse(readFileSync(manifestAbs, "utf8"));
} catch (e) {
	console.error(`pipeline manifest is not valid JSON: ${e.message}`);
	process.exit(2);
}
const base = abs(manifest.cwd || ".", dirname(manifestAbs));
const steps = Array.isArray(manifest.steps) ? manifest.steps : [];
if (!steps.length) {
	console.error("pipeline manifest declares no steps");
	process.exit(2);
}

// ── resolve steps into concrete commands ───────────────────────────────────

/** Returns { label, kind, argv, shell, optional, artifact } — no side effects. */
function resolveSteps() {
	const out = [];
	steps.forEach((step, i) => {
		const n = `[${i + 1}/${steps.length}]`;
		const optional = !!step.optional;
		if (step.run) {
			out.push({ label: `${n} run — ${step.desc || step.run}`, kind: "run", shell: step.run, optional });
		} else if (step.apply) {
			const cfg = abs(step.apply, base);
			const a = [NODE, resolve(HERE, "apply.js")];
			if (flags.dryRun) a.push("--dry-run");
			a.push(cfg);
			out.push({ label: `${n} apply — ${step.desc || cfg}`, kind: "apply", argv: a, optional, artifact: step.out ? abs(step.out, base) : null });
		} else if (step.extend) {
			const cfg = abs(step.extend, base);
			const a = [pythonExe, resolve(HERE, "extend", "extend.py")];
			if (flags.dryRun) a.push("--dry-run");
			a.push(cfg);
			out.push({ label: `${n} extend — ${step.desc || cfg}`, kind: "extend", argv: a, optional });
		} else if (step.validate) {
			const file = abs(step.validate, base);
			out.push({ label: `${n} validate — ${file}`, kind: "validate", argv: [NODE, resolve(HERE, "validate.js"), file], optional, artifact: file, skipInDryRun: true });
		} else if (step.proValidate) {
			const file = abs(step.proValidate, base);
			const a = [pythonExe, resolve(SKILL_ROOT, "pro", "scripts", "validate.py"), file];
			if (step.original) a.push("--original", abs(step.original, base));
			if (step.author) a.push("--author", step.author);
			out.push({ label: `${n} proValidate — ${file}`, kind: "proValidate", argv: a, optional, artifact: file, skipInDryRun: true });
		} else {
			console.error(`${C.red}step ${i + 1} has no recognised op (run / apply / extend / validate / proValidate)${C.off}`);
			process.exit(2);
		}
	});
	return out;
}

const resolved = resolveSteps();
const deliver = manifest.deliver ? abs(manifest.deliver, base) : [...resolved].reverse().find((s) => s.artifact)?.artifact || null;
const validatedSomething = resolved.some((s) => s.kind === "validate" || s.kind === "proValidate");

// Auto-validation: a build may not finish unvalidated by accident.
const autoSteps = [];
if (!flags.dryRun && !flags.noAutoValidate && deliver && !validatedSomething) {
	autoSteps.push({ label: `[auto] validate — ${deliver}`, kind: "validate", argv: [NODE, resolve(HERE, "validate.js"), deliver], optional: false, artifact: deliver });
	autoSteps.push({ label: `[auto] proValidate — ${deliver}`, kind: "proValidate", argv: [pythonExe, resolve(SKILL_ROOT, "pro", "scripts", "validate.py"), deliver], optional: false, artifact: deliver });
}
const allSteps = [...resolved, ...autoSteps];

const verifyStep = manifest.verify
	? { label: `[verify] assertions — ${abs(manifest.verify, base)}`, kind: "verify", argv: [NODE, resolve(HERE, "verify.js"), abs(manifest.verify, base)], optional: false }
	: null;

const cmdLine = (s) => (s.shell ? s.shell : s.argv.map(q).join(" "));

/** Windows path → forward slashes, so the emitted script survives Git Bash. */
const normPath = (s) => (/^[A-Za-z]:\\/.test(s) ? s.replace(/\\/g, "/") : s);
const shLine = (s) => (s.shell ? s.shell : s.argv.map((a) => shq(normPath(a))).join(" "));

// ── mode: plan ─────────────────────────────────────────────────────────────

if (flags.plan) {
	console.log(`${C.bold}docx-master pipeline — PLAN${C.off} ${C.dim}${manifestAbs}${C.off}`);
	console.log(`${C.dim}base: ${base}${flags.dryRun ? "  (dry-run)" : ""}${C.off}\n`);
	allSteps.forEach((s, i) => {
		console.log(`${String(i + 1).padStart(3)}. ${s.label}${s.optional ? `${C.dim} (optional)${C.off}` : ""}`);
		console.log(`     ${C.dim}$${C.off} ${cmdLine(s)}${s.skipInDryRun && flags.dryRun ? `  ${C.dim}(skipped in dry-run)${C.off}` : ""}`);
	});
	if (verifyStep) {
		console.log(`${String(allSteps.length + 1).padStart(3)}. ${verifyStep.label}`);
		console.log(`     ${C.dim}$${C.off} ${cmdLine(verifyStep)}`);
	}
	console.log(`\n${C.bold}Run in order. Abort on the first non-zero exit.${C.off}`);
	console.log(`${C.dim}Or emit a script that enforces this: --emit-script${C.off}`);
	process.exit(0);
}

// ── mode: emit script ──────────────────────────────────────────────────────

const scriptMode = flags.emit;
const spawnOk = scriptMode ? true : canSpawn();

if (scriptMode || !spawnOk) {
	const outPath = emitPath ? abs(emitPath, process.cwd()) : join(dirname(manifestAbs), "pipeline-run.sh");
	const L = [];
	L.push("#!/usr/bin/env bash");
	L.push("# Generated by docx-master pipeline — regenerate instead of editing.");
	L.push(`# manifest: ${manifestAbs}`);
	L.push(`# ${new Date().toISOString()}`);
	L.push("# `set -e` gives the same contract as running the pipeline directly:");
	L.push("# the first failing step aborts the whole chain.");
	L.push("set -euo pipefail");
	L.push(`cd ${shq(normPath(base))}`);
	L.push("");
	const body = [];
	allSteps.forEach((s, i) => {
		body.push(`echo ${shq(`── ${s.label}`)}`);
		if (s.skipInDryRun && flags.dryRun) {
			body.push(`echo "   (skipped in dry-run)"`);
		} else if (s.optional) {
			body.push(`if ! ${shLine(s)}; then echo "! optional step failed — continuing: ${s.label.replace(/"/g, "'")}"; fi`);
		} else {
			body.push(shLine(s));
		}
		body.push("");
	});
	if (verifyStep) {
		body.push(`echo ${shq(`── ${verifyStep.label}`)}`);
		body.push(shLine(verifyStep));
		body.push("");
	}
	body.push(`echo`);
	body.push(`echo "✓ pipeline complete"`);
	if (deliver) body.push(`echo "  deliverable: ${deliver.replace(/"/g, '\\"')}"`);
	L.push(...body);
	writeFileSync(outPath, L.join("\n"), "utf8");

	const why = scriptMode ? "(--emit-script)" : "(child processes blocked in this environment)";
	console.log(`${C.bold}docx-master pipeline${C.off} ${C.dim}${manifestAbs}${C.off}`);
	console.log(`${C.yellow}Deferred ${why} — nothing was executed.${C.off}`);
	console.log(`\nExecutable chain written to:\n  ${C.bold}${outPath}${C.off}`);
	console.log(`\n${allSteps.length + (verifyStep ? 1 : 0)} step(s), abort-on-first-failure:`);
	allSteps.forEach((s, i) => console.log(`  ${String(i + 1).padStart(2)}. ${s.label}`));
	if (verifyStep) console.log(`  ${String(allSteps.length + 1).padStart(2)}. ${verifyStep.label}`);
	console.log(`\nRun it with a single command:`);
	console.log(`  ${C.bold}bash ${q(outPath)}${C.off}`);
	console.log(`\n${C.dim}A non-zero exit from that script IS the build result — do not ship a
document from a chain that did not reach the end.${C.off}`);
	process.exit(scriptMode ? 0 : 3);
}

// ── mode: execute ──────────────────────────────────────────────────────────

function runStep(s) {
	if (s.skipInDryRun && flags.dryRun) {
		say(`  ${C.dim}(skipped in dry-run)${C.off}`);
		return;
	}
	say(`  ${C.dim}$${C.off} ${cmdLine(s)}`);
	const r = s.shell
		? spawnSync(s.shell, { shell: true, cwd: base, encoding: "utf8", stdio: ["ignore", "pipe", "pipe"] })
		: spawnSync(s.argv[0], s.argv.slice(1), { cwd: base, encoding: "utf8", stdio: ["ignore", "pipe", "pipe"] });
	const out = `${r.stdout || ""}${r.stderr || ""}`.trim();
	if (r.status !== 0) {
		if (s.optional) {
			say(`  ${C.yellow}! optional step failed (exit ${r.status}) — continuing${C.off}`);
			return;
		}
		console.error(`\n${C.red}✗ ${s.label} FAILED (exit ${r.status ?? r.error?.code ?? "?"})${C.off}`);
		if (r.error) console.error(`  ${r.error.message}`);
		if (out) console.error(out.split(/\r?\n/).slice(-25).map((l) => `  │ ${l}`).join("\n"));
		console.error(`\n${C.red}Chain aborted here — every later step depends on this one.${C.off}`);
		process.exit(1);
	}
	if (out) {
		const lines = out.split(/\r?\n/).filter(Boolean);
		say(lines.slice(-6).map((l) => `    ${C.dim}│${C.off} ${l}`).join("\n"));
	}
}

say(`${C.bold}docx-master pipeline${C.off} ${C.dim}${manifestAbs}${C.off}`);
say(`${C.dim}base: ${base}${flags.dryRun ? "  (dry-run)" : ""}${C.off}\n`);

for (const s of allSteps) {
	say(`${C.bold}${s.label}${C.off}`);
	runStep(s);
}

if (verifyStep && !flags.dryRun) {
	say(`${C.bold}${verifyStep.label}${C.off}`);
	runStep(verifyStep);
}

console.log(`\n${C.green}✓ pipeline complete${C.off}`);
if (deliver) console.log(`  deliverable: ${deliver}`);
console.log(`  steps: ${allSteps.length + (verifyStep ? 1 : 0)}${flags.dryRun ? `  ${C.yellow}(DRY-RUN — nothing written)${C.off}` : ""}`);
console.log(`\n${C.dim}Validation proves the file is well-formed, not that it says what the user
asked for. Read the change manifest and confirm every change against the user's
own words before delivering — see "Delivery gate" in SKILL.md.${C.off}`);
