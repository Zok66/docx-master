# 贡献指南

感谢关注 docx-master！这是一个直接操作 OOXML 的 Word 文档工具箱，欢迎提 issue、报缺陷、补能力。

## 开始之前

- 本仓库有**两个分支**：`main`（标准版）与 `pro`（全功能版）。两版共用同一套代码，差异只在 `scripts/extend/extend.py` 的 `OPERATIONS` 注册表。动手前先确认目标分支。
- 设计原则见 [README「设计原则」](README.md#设计原则) —— 尤其是「原文件永不被修改」与「稀疏声明」。

## 开发流程

1. Fork / 拉分支（从 `main` 起步）
2. 改代码 —— **一条规则只在一处定义**：SKILL.md 只放主流程与硬规则，细节进 `references/`
3. 跑回归：

   ```bash
   node tests/run.mjs --smoke     # 结构 + 规则 + 资源 + 链接 四类契约 + smoke case
   ```

4. 新增能力时在 `tests/contract.json` 同步补一条契约，在 `CHANGELOG.md` 记一轮
5. 提交 PR，描述里写清：改了什么、为什么、怎么验证的

## 硬约束

- **SKILL.md ≤ 20000 B**（回归测试会卡），塞不下就下移到 `references/` 并在 SKILL.md 里留一行指向
- 写盘命令必须产出**新文件**，绝不原地改
- 所有文档内链（含锚点）必须可解析，回归测试会查

## 缺陷报告

Issue 里请带上：最小复现 config、输入文档的关键结构（`overview` 输出即可）、期望 vs 实际、
`node tests/run.mjs` 结果。有 XSD 报错的一并贴上 —— 它通常直指根因。
