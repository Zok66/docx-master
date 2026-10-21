#!/usr/bin/env node
/**
 * docx-master 回归测试 —— 零依赖。
 *
 *   node tests/run.mjs            # 全部检查
 *   node tests/run.mjs --quiet    # 只打印失败项
 *   node tests/run.mjs --smoke    # 额外跑 tests/cases/*.json（需要能 spawn 子进程）
 *
 * 检查四类事实：
 *   1. 结构契约 —— frontmatter 合法、体积上限、description 长度
 *   2. 规则契约 —— tests/contract.json 里每条"不可丢失的强制规则"仍出现在 SKILL.md
 *   3. 资源契约 —— 契约清单里声明的 references / scripts 文件真实存在
 *   4. 链接完整性 —— 所有 md 的相对链接（含 #锚点）可解析
 *
 * 退出码：0 = 全绿；1 = 有失败项。
 */
import fs from 'node:fs';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(HERE, '..');
const QUIET = process.argv.includes('--quiet');
const SMOKE = process.argv.includes('--smoke');

const C = {
  ok: '\u001b[32m', bad: '\u001b[31m', dim: '\u001b[2m',
  bold: '\u001b[1m', off: '\u001b[0m',
};

let pass = 0;
const failures = [];

function ok(label) {
  pass++;
  if (!QUIET) console.log(`  ${C.ok}✓${C.off} ${label}`);
}
function bad(label, detail) {
  failures.push(`${label}${detail ? ` — ${detail}` : ''}`);
  console.log(`  ${C.bad}✗${C.off} ${label}${detail ? `\n      ${C.dim}${detail}${C.off}` : ''}`);
}
function group(name) {
  if (!QUIET) console.log(`\n${C.bold}${name}${C.off}`);
}

// ---------- 工具函数 ----------

/** GitHub 风格锚点（支持 CJK）：lowercase → 去标点 → 每个空白字符转一个连字符，重复标题加 -1/-2 */
function slugify(text) {
  return text
    .toLowerCase()
    .replace(/<[^>]*>/g, '')
    .replace(/[^\p{L}\p{N}\p{M}\s-]/gu, '')
    .trim()
    .replace(/\s/g, '-');
}

function stripInlineMd(t) {
  return t
    .replace(/`([^`]*)`/g, '$1')
    .replace(/\*\*([^*]*)\*\*/g, '$1')
    .replace(/\*([^*]*)\*/g, '$1')
    .replace(/\[([^\]]*)\]\([^)]*\)/g, '$1');
}

/** 抽出 md 里所有 ATX 标题的 slug */
function headingSlugs(md) {
  const seen = new Map();
  const out = new Set();
  for (const line of md.split(/\r?\n/)) {
    const m = line.match(/^(#{1,6})\s+(.*)$/);
    if (!m) continue;
    const base = slugify(stripInlineMd(m[2]));
    const n = seen.get(base) ?? 0;
    seen.set(base, n + 1);
    out.add(n === 0 ? base : `${base}-${n}`);
  }
  return out;
}

function walkMd(dir, acc = []) {
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    if (e.name === '.git' || e.name === 'node_modules' || e.name === '__pycache__') continue;
    const p = path.join(dir, e.name);
    if (e.isDirectory()) walkMd(p, acc);
    else if (e.name.endsWith('.md')) acc.push(p);
  }
  return acc;
}

function rel(p) {
  return path.relative(ROOT, p).split(path.sep).join('/');
}

// ---------- 1. 结构契约 ----------

group('结构契约');

const contract = JSON.parse(fs.readFileSync(path.join(HERE, 'contract.json'), 'utf8'));
const skillMdPath = path.join(ROOT, 'SKILL.md');
const skillMd = fs.readFileSync(skillMdPath, 'utf8');
const skillBytes = Buffer.byteLength(skillMd, 'utf8');

const fm = skillMd.match(/^---\r?\n([\s\S]*?)\r?\n---/);
if (!fm) {
  bad('SKILL.md 有合法 frontmatter');
} else {
  ok('SKILL.md 有合法 frontmatter');
  const fmBody = fm[1];
  const name = fmBody.match(/^name:\s*(.+)$/m);
  const desc = fmBody.match(/^description:\s*"([\s\S]*?)"\s*$/m);
  name && name[1].trim() === 'docx-master'
    ? ok('frontmatter name = docx-master')
    : bad('frontmatter name = docx-master', name ? `实际 "${name[1].trim()}"` : '缺失');
  desc
    ? ok(`frontmatter description 存在（${desc[1].length} 字符）`)
    : bad('frontmatter description 存在');

  if (desc && desc[1].length > contract.skillMdDescriptionMaxChars) {
    bad(`description ≤ ${contract.skillMdDescriptionMaxChars} 字符`,
      `实际 ${desc[1].length} —— 触发条件不该塞能力清单全量`);
  } else if (desc) {
    ok(`description ≤ ${contract.skillMdDescriptionMaxChars} 字符`);
  }
}

skillBytes <= contract.skillMdMaxBytes
  ? ok(`SKILL.md 体积 ≤ ${contract.skillMdMaxBytes} B（实际 ${skillBytes} B）`)
  : bad(`SKILL.md 体积 ≤ ${contract.skillMdMaxBytes} B`, `实际 ${skillBytes} B —— 该下移到 references/`);

// ---------- 2. 规则契约 ----------

group('规则契约（不可丢失的强制规则）');

const textCache = new Map([['SKILL.md', skillMd]]);
function textOf(relFile) {
  if (!textCache.has(relFile)) {
    const p = path.join(ROOT, relFile);
    textCache.set(relFile, fs.existsSync(p) ? fs.readFileSync(p, 'utf8') : null);
  }
  return textCache.get(relFile);
}

for (const { phrase, why, in: inFile } of contract.skillMdRequiredPhrases) {
  const target = inFile ?? 'SKILL.md';
  const text = textOf(target);
  if (text === null) {
    bad(`保留：${phrase}`, `规则声明的定义处 ${target} 不存在`);
  } else if (text.includes(phrase)) {
    ok(`保留：${phrase}  ${C.dim}(${target})${C.off}`);
  } else {
    bad(`保留：${phrase}`, `${target} 中已找不到。原因：${why}`);
  }
}

// ---------- 3. 资源契约 ----------

group('资源契约');

for (const f of contract.requiredReferenceFiles) {
  fs.existsSync(path.join(ROOT, f)) ? ok(`存在 ${f}`) : bad(`存在 ${f}`, '文件缺失');
}
for (const f of contract.requiredScripts) {
  fs.existsSync(path.join(ROOT, f)) ? ok(`存在 ${f}`) : bad(`存在 ${f}`, '文件缺失');
}
for (const { file, heading, why } of contract.readmeRequiredHeadings || []) {
  const p = path.join(ROOT, file);
  const hit = fs.existsSync(p) &&
    fs.readFileSync(p, 'utf8').split(/\r?\n/).some((line) => line.trim() === heading);
  hit ? ok(`${file} 含「${heading}」`) : bad(`${file} 含「${heading}」`, why);
}

// ---------- 4. 链接完整性 ----------

group('链接完整性');

const mdFiles = walkMd(ROOT).filter(
  (p) => !p.includes(path.sep + '.git' + path.sep),
);
const slugCache = new Map();
function slugsOf(file) {
  if (!slugCache.has(file)) slugCache.set(file, headingSlugs(fs.readFileSync(file, 'utf8')));
  return slugCache.get(file);
}

let linkCount = 0;
for (const file of mdFiles) {
  const md = fs.readFileSync(file, 'utf8');
  const re = /\[[^\]]*\]\(([^)\s]+)\)/g;
  let m;
  while ((m = re.exec(md))) {
    const target = m[1];
    if (/^(https?:|mailto:|tel:|data:)/.test(target)) continue;
    if (target.startsWith('#')) continue; // 同文件锚点自查在下面统一做
    linkCount++;

    const [rawPath, anchor] = target.split('#');
    const resolved = path.resolve(path.dirname(file), decodeURIComponent(rawPath));
    if (!fs.existsSync(resolved)) {
      bad(`链接可解析：${rel(file)} → ${target}`, '目标文件不存在');
      continue;
    }
    if (anchor && resolved.endsWith('.md')) {
      const slugs = slugsOf(resolved);
      slugs.has(anchor)
        ? ok(`锚点有效：${rel(file)} → ${target}`)
        : bad(`锚点有效：${rel(file)} → ${target}`, `目标标题 slug 不存在（可能标题被改动）`);
    }
  }
  // 同文件纯锚点
  const re2 = /\[[^\]]*\]\((#[^)\s]+)\)/g;
  while ((m = re2.exec(md))) {
    const anchor = m[1].slice(1);
    slugsOf(file).has(anchor)
      ? ok(`同文件锚点：${rel(file)} → #${anchor}`)
      : bad(`同文件锚点：${rel(file)} → #${anchor}`, '本文件无此标题');
  }
}
if (linkCount === 0) bad('存在可校验的相对链接', '一个都没扫到 —— 正则或文件集有问题');

// ---------- 5. smoke（可选） ----------

if (SMOKE) {
  group('smoke cases');
  const casesDir = path.join(HERE, 'cases');
  const files = fs.existsSync(casesDir)
    ? fs.readdirSync(casesDir).filter((f) => f.endsWith('.json') && !f.endsWith('.config.json'))
    : [];
  if (!files.length) {
    ok('无 smoke case（tests/cases/*.json 为空，跳过）');
  }
  for (const f of files) {
    const spec = JSON.parse(fs.readFileSync(path.join(casesDir, f), 'utf8'));
    const cfg = path.resolve(ROOT, spec.config);
    if (!fs.existsSync(cfg)) { bad(`case ${spec.name}`, `config 不存在 ${spec.config}`); continue; }
    const r = spawnSync(process.execPath, [path.join(ROOT, 'scripts', 'apply.js'), '--dry-run', cfg], {
      cwd: ROOT, encoding: 'utf8', timeout: 120000,
    });
    const want = spec.expectExit ?? 0;
    const got = r.status ?? -1;
    if (got === want) ok(`case ${spec.name}（exit ${got}）`);
    else bad(`case ${spec.name}`, `期望 exit ${want}，实际 ${got}\n      ${(r.stderr || r.stdout || '').split('\n').slice(0, 6).join('\n      ')}`);
  }
}

// ---------- 汇总 ----------

const total = pass + failures.length;
console.log(`\n${C.bold}${failures.length ? C.bad : C.ok}${failures.length ? 'FAIL' : 'PASS'}${C.off}  ${total - failures.length}/${total} 通过`);
if (failures.length) {
  console.log(`\n${C.bad}失败项：${C.off}`);
  failures.forEach((f, i) => console.log(`  ${i + 1}. ${f}`));
}
process.exit(failures.length ? 1 : 0);
