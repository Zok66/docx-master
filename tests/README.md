# 回归测试

零依赖，一条命令：

```bash
node tests/run.mjs            # 全部检查
node tests/run.mjs --quiet    # 只打印失败项
node tests/run.mjs --smoke    # 额外跑 tests/cases/*.json（需要能 spawn 子进程）
```

退出码：`0` = 全绿；`1` = 有失败项。可直接接进 CI。

## 检查四类事实

| # | 类别 | 内容 |
|---|---|---|
| 1 | **结构契约** | SKILL.md frontmatter 合法、体积 ≤ `skillMdMaxBytes`、description 长度上限 |
| 2 | **规则契约** | `contract.json` 里每条「不可丢失的强制规则」仍出现在其声明的文件里 |
| 3 | **资源契约** | 契约声明的 `references/*.md` 与 `scripts/*.js` 真实存在；README / README_EN 的必备章节（FAQ / 开发规范 / 版本与更新日志）不得丢失 |
| 4 | **链接完整性** | 所有 md 的相对链接（含 `#锚点`，按 GitHub slug 规则）可解析 |

第 2 类是关键：它把「不能为了省 token 删掉必要约束」变成了一条**可执行的断言**。
主控精简时误删任何一条硬规则，测试立刻变红。

## 怎么加一条「不许删的规则」

在 `tests/contract.json` 的 `skillMdRequiredPhrases` 里加一条：

```json
{ "phrase": "原文里必须逐字出现的字符串", "why": "这条规则为什么重要", "in": "references/xxx.md" }
```

- `in` 省略 → 默认查 `SKILL.md`；规则定义在子文件时写子文件的相对路径。
- `phrase` 必须与文件里的文本**逐字一致**（含反引号、大小写、标点）。
  写错了跑一次 `node tests/run.mjs --quiet` 就会报出来。
- 规则只在一处定义。若某条规则从 SKILL.md 下移到子文件，记得同步改 `in`。

## 怎么加一个 smoke case

1. 在 `tests/cases/` 放一个 apply config，文件名以 `.config.json` 结尾
   （后缀用于把它和 case 声明区分开）。
2. 写 case 声明，文件名任意 `.json`：

```json
{ "name": "用例名", "config": "tests/cases/xxx.config.json", "expectExit": 0, "why": "这条用例守的是什么" }
```

`expectExit` 默认 `0`；故意要失败的用例写 `1`。

## 改完 skill 之后的动线

```bash
node tests/run.mjs --smoke
```

改主控、下移规则、调整链接之后都跑一遍 —— 它保证「没丢规则、没断链、没超预算」。
