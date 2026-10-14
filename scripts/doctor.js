#!/usr/bin/env node
/**
 * `doctor` — environment capability probe.
 *
 * Run this FIRST on a new machine or at the start of a long task. Every
 * optional tool in this skill (LibreOffice, poppler, pandoc, the Python
 * toolchain) is exactly the kind of dependency that fails late: you design a
 * render-verify loop, get nine steps in, and only then discover `soffice` was
 * never installed. Probing up front turns that into a one-line answer, and
 * tells you which documented paths are live and which degrade.
 *
 * Usage:
 *   node scripts/doctor.js [--json] [--python <exe>]
 *
 * Exit code: always 0 — this is a report, not a gate. (The Node engine alone
 * is enough for every core `apply` / `extend` / `validate` path; missing
 * optional tools only disable the paths listed against them.)
 */
import { spawnSync } from "node:child_process";
import { existsSync, readdirSync } from "node:fs";
import { join } from "node:path";

const argv = process.argv.slice(2);
const asJson = argv.includes("--json");
let pythonExe = process.env.DOCX_MASTER_PYTHON || "python";
const pyFlag = argv.indexOf("--python");
if (pyFlag !== -1 && argv[pyFlag + 1]) pythonExe = argv[pyFlag + 1];

function probe(cmd, args, opts = {}) {
	try {
		const r = spawnSync(cmd, args, { encoding: "utf8", shell: false, timeout: 20000 });
		if (r.error) {
			// EBUSY / EPERM here usually means the sandbox blocked process creation,
			// not that the tool is absent — say so, or the report misleads.
			const code = r.error.code || r.error.message;
			const blocked = code === "EBUSY" || code === "EPERM";
			return { ok: false, blocked, detail: blocked ? `blocked by sandbox (${code})` : code };
		}
		const out = `${r.stdout || ""}${r.stderr || ""}`.trim().split(/\r?\n/)[0] || "";
		return { ok: r.status === 0, detail: out.slice(0, 90), status: r.status };
	} catch (e) {
		return { ok: false, detail: e.message };
	}
}

/** LibreOffice ships under several names and locations; try the likely ones. */
function findSoffice() {
	const candidates = [
		"soffice",
		"C:/Program Files/LibreOffice/program/soffice.exe",
		"C:/Program Files (x86)/LibreOffice/program/soffice.exe",
		"/Applications/LibreOffice.app/Contents/MacOS/soffice",
		"/usr/bin/soffice",
		"/usr/local/bin/soffice"
	];
	for (const c of candidates) {
		const isPath = c.includes("/") || c.includes("\\");
		if (isPath && !existsSync(c)) continue;
		const r = probe(c, ["--version"]);
		if (r.ok) return { ok: true, detail: r.detail, cmd: c };
	}
	return { ok: false, detail: "not found" };
}

/** Managed WorkBuddy pythons on disk (existence check only — no spawning). */
function managedEnvPythons() {
	const out = [];
	const home = process.env.USERPROFILE || process.env.HOME;
	if (!home) return out;
	const root = join(home, ".workbuddy", "binaries", "python", "envs");
	if (!existsSync(root)) return out;
	try {
		for (const env of readdirSync(root)) {
			const exe = join(root, env, "Scripts", "python.exe");
			const exeU = join(root, env, "bin", "python");
			if (existsSync(exe)) out.push(exe);
			else if (existsSync(exeU)) out.push(exeU);
		}
	} catch { /* ignore */ }
	return out;
}

function findPython() {
	const managed = managedEnvPythons();
	// Order matters. A provisioned venv beats a bare interpreter even when it
	// is not first on PATH: "python" from PATH is usually the stock runtime
	// without lxml, and probing stops at the first interpreter that merely
	// *starts*. Reporting "MISS Python + lxml" while a fully provisioned
	// venv sits in the managed envs directory sends the agent off to install
	// packages that are already there.
	const preferred = managed.filter((p) => /envs[\\/]default[\\/]/.test(p));
	const rest = managed.filter((p) => !preferred.includes(p));
	const tried = [...preferred, pythonExe, "python3", "python", ...rest];
	const seen = new Set();
	let fallback = null;
	const probedList = [];
	for (const exe of tried) {
		if (seen.has(exe)) continue;
		seen.add(exe);
		const r = probe(exe, ["--version"]);
		if (!r.ok) continue;
		const deps = probe(exe, ["-c", "import lxml, defusedxml"]);
		probedList.push(`${exe}${deps.ok ? "" : " (no lxml)"}`);
		const candidate = {
			ok: true,
			cmd: exe,
			detail: r.detail,
			deps: deps.ok,
			depsDetail: deps.ok ? "lxml + defusedxml" : "missing lxml/defusedxml",
			probed: probedList
		};
		if (deps.ok) return candidate;
		// Usable but incomplete — remember it and keep looking for a better one.
		if (!fallback) fallback = candidate;
	}
	if (fallback) {
		// Say where the packages could come from instead of leaving the agent
		// to guess that the machine has only one interpreter.
		if (managed.length) {
			fallback.detail += `  → venv(s) found: ${managed.join(", ")}`;
		}
		return fallback;
	}
	// Nothing probed successfully. Distinguish "absent" from "present but
	// unprobeable": a sandbox that blocks process creation would otherwise make
	// an installed interpreter look missing, sending you off to install it.
	if (managed.length) {
		const pick = preferred[0] || managed[0];
		return {
			ok: false,
			unprobeable: true,
			detail: `${managed.length} interpreter(s) present but unprobeable (process creation blocked) — e.g. ${pick}`
		};
	}
	return { ok: false, detail: "not found" };
}

const checks = [];

checks.push({
	name: "Node engine",
	why: "apply / extend-driver / validate / inspect / batch_fill — the whole core",
	res: { ok: true, detail: `${process.version} (${process.execPath})` }
});

// Can this process start child processes at all? `pipeline` orchestrates the
// chain as child processes; a sandbox that blocks creation degrades it to
// --emit-script, which is a different (still one-command) workflow.
const spawnProbe = spawnSync(process.execPath, ["-e", "process.exit(0)"], { encoding: "utf8" });
const canSpawn = !spawnProbe.error && spawnProbe.status === 0;
checks.push({
	name: "Subprocess spawning",
	why: "`pipeline` runs apply / extend / validate as child processes; without it use --emit-script and run the generated .sh",
	res: canSpawn
		? { ok: true, detail: "available" }
		: { ok: false, detail: `blocked (${spawnProbe.error?.code || `status ${spawnProbe.status}`}) — child processes are forbidden here; pipeline --emit-script still works` }
});

const py = findPython();
checks.push({
	name: "Python + lxml",
	why: "`extend` (TOC / fields / watermarks / drawings / charts) and the PRO validator",
	res: py.ok ? { ok: py.deps, detail: `${py.cmd} — ${py.detail}${py.deps ? "" : " [tools import failed]"}` } : py,
	cmd: py.cmd
});
if (py.ok && !py.deps) checks[checks.length - 1].res.detail += "  → pip install lxml defusedxml";

checks.push({
	name: "LibreOffice (soffice)",
	why: "legacy .doc → .docx conversion; PDF render for visual self-check",
	res: findSoffice()
});

checks.push({
	name: "pdftoppm (poppler)",
	why: "PDF → image, the last leg of the render-and-eyeball path",
	res: probe("pdftoppm", ["-v"])
});

checks.push({
	name: "pandoc",
	why: "foreign-format import (md / html / odt → docx)",
	res: probe("pandoc", ["--version"])
});

checks.push({
	name: "npm docx",
	why: "docx-js scaffolding path only",
	res: (() => {
		const r = probe("node", ["-e", "require.resolve('docx')"]);
		return r;
	})()
});

if (asJson) {
	console.log(JSON.stringify(checks.map((c) => ({ name: c.name, ok: !!c.res.ok, detail: c.res.detail, why: c.why, cmd: c.cmd })), null, 2));
	process.exit(0);
}

const line = "─".repeat(74);
console.log(`docx-master doctor\n${line}`);
for (const c of checks) {
	const skip = c.res.unprobeable || c.res.blocked;
	const tag = c.res.ok ? "OK  " : skip ? "SKIP" : "MISS";
	console.log(`${tag} ${c.name}`);
	console.log(`     ${c.res.detail || ""}`);
	console.log(`     need for: ${c.why}`);
}
console.log(line);
const missing = checks.filter((c) => !c.res.ok && !c.res.unprobeable && !c.res.blocked);
const unprobeable = checks.filter((c) => !c.res.ok && (c.res.unprobeable || c.res.blocked));
if (unprobeable.length) {
	console.log(`SKIP = cannot be probed here (a sandbox is blocking process creation);`);
	console.log(`       this says nothing about whether the tool is installed.`);
}
if (!missing.length) {
	console.log("Everything probed is available — including the render-verify loop.");
} else {
	console.log(`${missing.length} capability gap(s):`);
	for (const m of missing) console.log(`  · ${m.name} — ${m.why}`);
	console.log("\nCore writes and schema validation work without these. Stay on the Node engine,");
	console.log("or install the missing tool before relying on the path that needs it.");
}
