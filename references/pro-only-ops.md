# pro 独有操作（标准版不含）

标准版面向**日常办公 / 论文 / 实训报告**：页数、目录、分页、页码、页眉页脚、脚注尾注、
图表题注、交叉引用、批注、修订、批量套打、交付门——够用且轻。

以下 8 个 `extend` op 只在 **pro 版**（`git checkout pro`）提供。标准版调用会被明确拒绝，
报错信息带 `UNSUPPORTED_OPS` 提示，不会静默降级。

| op | 能力 | 为什么归 pro |
|---|---|---|
| `watermark` | 页眉里的 VML 水印（`w:pict` + `v:textpath`） | 内部资料 / 保密标识，非日常排版 |
| `drawing` | 形状、文本框、艺术字变形、符号段落、手写墨迹、连接线 | 图文设计类，需精细调参 |
| `canvas` | 绘图画布（多形状按厘米坐标摆放） | 同上 |
| `smartArt` | 智能图形（水平流程 / 环形 / 垂直列表 / 树形） | 需生成 `dgm` 部件与布局，属演示型排版 |
| `chart` | 图表（bar / line / pie …，内嵌工作簿数据） | 数据分析场景，非文档排版核心 |
| `ole` | OLE 嵌入对象（载荷工作簿 + 预览图） | 冷门，且需构造复合文档 |
| `glossary` | 构建基块 / 自动图文集（`word/glossary/document.xml`） | 模板制作场景 |
| `customXml` | 自定义 XML 部件（`customXml/itemN.xml`） | 系统对接场景 |

## 切到 pro

```bash
cd ~/.workbuddy/skills/docx-master
git checkout pro && git pull
```

pro 分支同时保留完整 `pro/` 工具链（XSD 全套架构 + 批注 / 接受修订 / 渲染 / `.doc` 转换）。
切回标准版：`git checkout main`。

## 标准版仍有的画图能力

**嵌入式图片**（`image` op）在标准版完整保留：本地图片自动入库、裁剪、亮度对比度、`effectLst`
影效、浮动环绕。论文 / 实训报告需要的插图、截图、实验结果图都走这条。

## 历史实现

被剥离的函数体（`op_watermark` / `op_drawing` / `op_canvas` / `op_smart_art` / `op_chart` /
`op_ole` / `op_glossary` / `op_custom_xml`）仍在 `scripts/extend/extend.py` 中，只是未注册进
`OPERATIONS`。这样 pro 分支合并时不会冲突，也方便将来把某个能力挪回标准版——把名字加回
`OPERATIONS` 即可。
