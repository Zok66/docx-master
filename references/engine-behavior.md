# 引擎自动行为（Engine behavior）

`apply` / `extend` 在每次写入时**替你完成**的事项。你不需要主动调用或声明它们，
但设计配置、解读校验输出、预判产出时要知道它们的存在——否则容易把"引擎已经处理好的事"
重复声明一遍，或者误读校验失败的原因。

主控 [`SKILL.md`](../SKILL.md) 的 Hard rules 只列「Agent 必须主动遵守」的规则；本文件是其镜像。
两边不重复：某条规则**只有一个**定义处，改动时同步核对另一侧的指向。

## 校验与预检

- **Baseline-diff validation.** Every applying CLI writes a fresh copy and validates before
  keeping it; only errors *introduced* by the run are fatal. Override:
  `--allow-validation-warnings` ([config-schema.md](config-schema.md)).
- **Style-name preflight.** Before any mutation the engine catches `<w:name>` collisions
  (including en / zh-CN locale aliases) and aborts with a fix hint; `--dry-run` lists
  collisions as warnings instead.

## 写入时的自动处理

- **Section properties are sparse-by-design**: untouched unless declared via `pageSetup`
  ([config-schema.md §page setup](config-schema.md#page-setup)).
- **Restyle semantics**: run-level direct formatting uniform across all runs gets stripped;
  per-run differences are preserved as intentional inline emphasis.
- **`settings.elements` is order-safe and duplicate-safe**: each entry merges into an
  existing element of the same name, otherwise inserts at the XSD-correct position from the
  authoritative 98-child order. `updateFields`, `evenAndOddHeaders`, etc. never appear twice.
  Repeatable children work too — declaring the same child name more than once (e.g. three
  `docVar`s inside `docVars`, or several `rsid`s) appends one node per declaration instead of
  collapsing them into a single node.
- **`docProps/core.xml` + `app.xml`** are created or refreshed automatically
  (`docProps.enabled`, default on): title / creator / revision plus `Words` / `Characters` /
  `Paragraphs` counted the way Word counts Chinese (CJK per character, Latin per word).
  These are exactly what `NUMPAGES` / `NUMWORDS` / `AUTHOR` fields read.
- **Root namespace declarations are repaired** on every part written — `xmlns:w14` / `w15` /
  `wp14` / `wps` / `wpg` / `wpc` / `v` / `o` / `mc` as needed, with rendering prefixes
  (`wps` / `wpg` / `wpc`) explicitly kept **out** of `mc:Ignorable`. Template files declaring
  only `xmlns:w` + `xmlns:r` are the normal case, not the exception.

## 为什么这些不放在主控

主控每次触发都会全量注入。上面每条都是「引擎已经做了、你只需知道」的事实，
而不是「你必须记得做」的动作——把它们留在主控只会稀释真正需要当场遵守的规则。
需要时读本文件即可；`tests/contract.json` 保证它们不会被误删。
