# Third-Party Notices — 第三方组件声明

本仓库包含下列第三方组件，按其各自许可证再分发。本仓库整体以 MIT 授权（见 [LICENSE](LICENSE)），
以下组件的授权条款独立生效。

---

## 1. Temml（TeX → MathML 转换器）

- **位置**：`scripts/temml/`（vendored，未做修改）
- **版本**：0.13.02
- **来源**：https://github.com/ronkok/Temml
- **许可证**：MIT License，Copyright (c) 2020 Ron Kok
- **许可证全文**：[`scripts/temml/LICENSE`](scripts/temml/LICENSE)
- **用途**：公式（`w:oMath`）生成链路中把 TeX 语法转换为 MathML

## 2. ECMA-376 / OOXML Schema（XSD 校验架构集）

- **位置**：`pro/schemas/`
- **来源**：ECMA-376（Office Open XML 文件格式，ISO/IEC 29500）官方规范附带的 XSD 文件，
  含 Microsoft 2010 扩展（`wml-2010.xsd` 等）
- **授权依据**：
  - ECMA-376 规范文本与附件允许为实现目的自由使用与再分发（ECMA Patent / Copyright 政策：
    实现规范所需的使用即视为已获许可）
  - Microsoft 对 OOXML 实现做出 [Open Specification Promise](https://learn.microsoft.com/openspecifications/dev-center/osp)，
    承诺不就实现该规范主张必要专利
- **用途**：`pro/scripts/validate.py` 对产出的 docx 做 XSD 校验（写盘前自动执行）
- **说明**：这些文件按原样再分发、未做修改；如 ECMA/微软要求移除，可随时用官方规范包重新获取

---

## 3. `pro/` 可选 Python 工具链（衍生自 docx-pro）

- **位置**：`pro/scripts/`（`comment.py`、`accept_changes.py`、`merge_runs.py`、`soffice.py`、
  `validate.py`、`validators_*.py`、`element_order.py`、`helpers.py`）、`pro/templates/`、`pro/zh-CN.md`
- **来源**：技能市场中的 `docx-pro` 技能（展示名「Word 文档处理」，作者 `zxh`，版本 1.0.3）
- **授权**：该技能未在其元数据中声明独立许可证。本仓库按技能市场的分发条款继续使用这部分代码，
  不对其主张额外权利。
- **本地改动**：在本仓库中做过适配与扩展（如新增 `render.py`、调整校验器调用约定）。
- **不包含**：`pro/schemas/` 的 XSD 文件属第 2 节所述的 ECMA-376 规范文件，与本条无关。
- **异议处理**：若权利人认为此处的再分发方式不当，请提交 issue，将立即移除该目录或替换为等价的自研实现。

---

## 本仓库原创部分的授权

除第 1–3 节所列目录外，本仓库其余全部内容（`scripts/*.js`、`scripts/extend/*.py`、`references/`、
`tests/`、文档与资产）均为原创，以 MIT License 授权：

> MIT License · Copyright (c) 2026 Zok66 · 详见 [LICENSE](LICENSE)
