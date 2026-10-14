# Out of scope — and what to use instead

每一节回答同一个问题：**为什么不做**，以及**替代路径是什么**。列在这里的能力都有一个共同点——离线静态构造无法可靠完成，或者根本不属于「文档层」而属于 Word 应用层 / 外部服务。写进来只会产出 Word 会弹修复框、或被 XSD 拒绝的文件。

## 需要二进制部件

### 宏 — `.vbaProject.bin`

`vbaProject.bin` 是 OLE 复合文档（CFB）容器，内含压缩的 VBA 源码与编译中间态。格式无公开完整规范，且不能只塞进一段 XML 就生效——Word 还要校验签名与工程校验和。

**替代：** 无。需要宏就让用户自己在 Word 里录制；或用 PRO `accept_changes.py` 把宏执行后的结果固化下来（只处理结果，不生成宏）。

### ActiveX 控件

控件的类型信息、持久化属性都是二进制 OLE 流（`word/activeX/*.bin` + `activeX1.xml` 的属性流）。静态构造出的控件 Word 会拒载。

**替代：** 内容控件 SDT（`extend` 的 `sdt` op），支持纯文本、下拉列表、日期等常见形态，且完全在文本 XML 内。需要 ActiveX 特有行为（如绑外部 DLL）时，明确告诉用户这一项做不到。

### 真 SmartArt — `dgm:*` 部件

SmartArt 的版式部件（`word/diagrams/data1.xml` + `layout1.xml` + `quickStyle1.xml` + `colors1.xml`）里是 Word 内置的**版式计算规则**（`forEach` / `alg` / `constrLst`），不是画好的坐标。离线硬拼出的 `dgm` 部件与版式算法不匹配时，Word 打开即弹修复框。

**替代：** `extend` 的 `smartArt` op——用 `wpc` 绘图画布 + 预设形状 + 连线，自动算出 process / cycle / list / hierarchy 四种版式的坐标。节点文字可编辑、可整体缩放。差别：它不是 `dgm` 对象，没有「SmartArt 工具」选项卡与一键换版式。

### InkML 手写笔迹

Word 原生墨迹是 `word/ink/ink1.xml`（InkML）+ `<w14:ink>` 扩展 + 笔迹画布。缺可验证的 XSD 覆盖，笔顺 / 压感 / 回放都依赖 Word 的墨迹引擎。

**替代：** `extend` 的 `drawing` op 的 `ink` 字段——VML 折线（`v:shape` + `v:path`），静态渲染一致，代价是没有压感与笔顺回放。

## 需要外部应用或服务

### `.dotx` / `.dotm` 模板导出

技能的输出通道固定 `.docx`，没有内容类型切换与模板元数据（`docProps` 的模板位）写入路径。技能里的 `template` 块做的是反方向的「**从**模板导入样式与编号」，不是导出。

**替代：** 产出 `.docx` 后由用户另存为模板（Word：另存为 → Word 模板）。若确实需要机器侧生成，可复制 `.docx` 改扩展名并改 `[Content_Types].xml` 的主文档类型——但这是脚本层的活，不在本技能范围内。

### PDF / XPS 导出

依赖 LibreOffice（`soffice --convert-to pdf`）或 Word 本身。

**替代：** `doctor.js` 确认机器上有没有 `soffice`；有则用 PRO `soffice.py`，配 `pdftoppm` 还能把 PDF 拆成图做视觉核对。没有就交付 `.docx`，让用户自己导出。

### 文档翻译 / 语音听写

外部服务能力（Word 的翻译走云端、听写走系统语音栈）。

**替代：** 内容层面的翻译由本技能之外的流程完成——产出译文后再用 `apply` 重排。

### 信封 / 标签批量打印

需要信封节（`w:sectPr` 的信封尺寸组合）与标签网格版式生成器，且强依赖具体打印机驱动。

**替代：** 邮件合并。`apply` 能写 `settings.xml` 的 `mailMerge`（主文档类型 / 数据源类型 / 输出方式）+ 行内 `MERGEFIELD` 域；真正批量产出走 `scripts/batch_fill.js`（模板 + CSV/JSON → 每行一份文档，页眉页脚脚注一并替换）。

## 属于 Word 应用层，不属于文档层

写了也不会生效——它们保存在用户配置或外部部件里，文档打开时的行为由 Word 自己决定。

| 能力 | 原因 | 替代 |
|---|---|---|
| 自动更正词条 | 应用层设置（`AutoCorrect` 表在 Word 配置里，不在文档里） | 无 |
| 拼写 / 语法词典、同义词库 | 依赖 Word 校对引擎与已装语言包 | 只能写标记：`proofState`（`spelling` / `grammar`）、`hideSpellingErrors`；真正的检查由 Word 执行 |
| 自定义词典 | 外部文件（`.dic`） | 无 |
| 双面打印 / 打印机设置 | `printerSettings` 部件是驱动相关二进制 | `settings.xml` 的 `printTwoOnOne` 缩打可以写 |
| 样式集（快速样式集） | `stylesWithEffects.xml` + `stylePaneFormatFilter` 的主题化机制 | 写入样式定义（`styles[]`）与主题（`theme`）；按钮式「样式集」切换不做 |

## 相关

- 能力逐项对照与实测记录：`docx-master实现情况清单.md`（演示工程产出）
- 各 op 的字段与边界：[extend.md](extend.md)、[config-schema.md](config-schema.md)
- 环境探测：[delivery-gate.md](delivery-gate.md)、[pro-toolchain.md](pro-toolchain.md)
