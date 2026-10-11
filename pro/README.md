# `pro/` —— 可选 Python 工具链

本目录是 **`docx-master` 的可选扩展层，不是独立技能**。默认链路仍是 Node 引擎
（`scripts/apply.js` + `inspect_*` / `find_*`）；只有 Node 引擎唯一不覆盖的场景才走这里。

| 需要什么 | 用哪个 |
|---|---|
| 加批注 / 回复批注 | `scripts/comment.py` |
| 接受全部修订，产出干净副本 | `scripts/accept_changes.py` |
| 合并碎片化 run（让文本在 XML 中重新连续、可查找） | **Node 优先**：`../scripts/merge_runs.js`（还能覆盖页眉 / 页脚 / 脚尾注，或给 `apply` 配 `mergeRuns: true`）；文档已解包成目录时用 `scripts/merge_runs.py` |
| XSD 校验 + 修订完整性核查（`--author`） | `scripts/validate.py` |
| 查任意元素的 XSD 强制子元素顺序 | `scripts/element_order.py` |
| 渲染成图、人工核对版式 | `scripts/soffice.py` + `pdftoppm` |
| 用 docx-js（npm `docx`）从零建文档 | 方法论与陷阱清单见 [`zh-CN.md`](zh-CN.md) |

## 目录内容

| 路径 | 内容 |
|---|---|
| `scripts/` | CLI 入口（批注 / 修订 / run 合并 / 校验 / 元素顺序 / 渲染 / docx-js 辅助） |
| `templates/` | 批注六件套模板（`comments*.xml` / `people.xml`） |
| `schemas/` | ECMA-376 + ISO-IEC29500 XSD 全量集（供 `validate.py` 做 XSD 校验） |
| `zh-CN.md` | docx-js 创建与 pandoc 读取的方法论原文 |

## 进一步阅读

- **能力路由、依赖状态、调用约定、常用工作流、边界与冲突、安全审计结论、回退路径**
  → [`../references/pro-toolchain.md`](../references/pro-toolchain.md)
- 主技能入口 → [`../SKILL.md`](../SKILL.md)

## 来源

整合自市场技能 **`docx-pro`**（display_name「Word 文档处理」，v1.0.0，category `document`，
author `zxh`，skillId `skill_2099517733028458496`），整合日期 2026-10-10，
逐文件安全审计结论见 [`../references/pro-toolchain.md`](../references/pro-toolchain.md) §6。

原技能目录已从本机移除，**`pro/` 是唯一留存副本**；`docx-master` 不依赖任何外部技能即可独立运行。
本层未随附原技能的许可证文件 —— 对外分发本技能前需自行确认授权条款。
