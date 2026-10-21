# Push-Sync —— GitHub → Gitee 镜像同步指南

> 给 AI 助手的操作手册。目标：把 GitHub 仓库的**全部分支、标签、发行版、许可证与仓库元信息**
> 镜像到 Gitee，并保证每次发版后两边一致。
>
> 本文以 `Zok66/docx-master` 为实例，但流程适用于任何仓库——把变量换掉即可。

**配套脚本**（已实测可用）：

| 脚本 | 作用 |
|---|---|
| [`scripts/sync-releases.sh`](scripts/sync-releases.sh) | 把 GitHub Release 幂等镜像到 Gitee |
| [`scripts/verify-sync.sh`](scripts/verify-sync.sh) | 双向核对分支 / Tag / Release / 元信息 / SHA |

**依赖**：`gh`（已登录）、`curl`、`python`。**不需要 `jq`**——脚本用 Python 解析 JSON（本机实测无 jq）。

---

## 0. 同步范围（哪些能同步、哪些不能）

| 对象 | 能同步 | 机制 | 备注 |
|---|---|---|---|
| 提交历史 | ✅ | `git push` | 天然一致 |
| 分支 Branches | ✅ | `git push <remote> <branch>` | 需显式指定每个分支，或用 `--all` |
| 标签 Tag | ✅ | `git push --tags` / `--follow-tags` | annotated tag 要带 `^{}` 才算完整 |
| 发行版 Releases | ✅ | **API**（Gitee 无 git 协议） | 需脚本，见 §4 |
| 许可证 License | ⚠️ 半自动 | 仓库元信息 PATCH | 见 §5 的坑：Gitee 转公开必须先设 license |
| 默认分支 | ⚠️ 手动 | 仓库设置 / API | 见 §5 |
| 仓库可见性（公开/私有） | ⚠️ 手动 | PATCH `private` | 建仓时传 `private=false` **不生效** |
| 仓库描述 | ⚠️ 手动 | PATCH `description` | |
| Issues / PR / Wiki / Actions | ❌ | — | Gitee 无对应同步机制，只能人工迁移 |
| Star / Watch / Fork | ❌ | — | 平台独立数据 |
| GitHub Actions 工作流 | ❌ | — | Gitee 有「GoGo 流水线」，语法不同，需重写 |

**结论**：`git push` 解决代码与 tag；**Releases、license、可见性、默认分支只能靠 API**。

---

## 1. 前置：认证

### 1.1 GitHub

```bash
gh auth status                     # 应显示 Logged in ... account <user>
```

本机已用 SSH（`git@github.com:`），`~/.ssh/config` 配了 `Host github.com → ssh.github.com:443`。

### 1.2 Gitee

需要一个 **私人令牌**（Gitee → 设置 → 私人令牌，勾选 `projects` 权限）。

```bash
# 建议放环境变量，不要硬编码进仓库
export GITEE_TOKEN='<你的令牌>'
export GITEE_OWNER='Zok66'
export REPO='docx-master'
```

> ⚠️ **令牌安全**：不要把令牌写进任何被 git 跟踪的文件。本仓库当前把令牌存在
> `.git/config` 的 remote URL 里（不被跟踪，但明文落盘）。更稳妥的做法见 §7。

验证令牌：

```bash
curl -s "https://gitee.com/api/v5/user?access_token=$GITEE_TOKEN" | head -c 200
```

---

## 2. 配置双远端推送（一次性）

让 `origin` 同时推两个平台，之后一条命令搞定代码同步。

```bash
# 先确认现状：origin 的 fetch 指向 GitHub
git remote -v

# 给 origin 添加第二个 push 地址（注意：这会替换默认 push 行为）
git remote set-url --add --push origin git@github.com:Zok66/docx-master.git
git remote set-url --add --push origin https://Zok66:<TOKEN>@gitee.com/Zok66/docx-master.git

# 验证：应看到两个 (push)
git remote -v
```

预期输出：

```
origin  git@github.com:Zok66/docx-master.git (fetch)
origin  git@github.com:Zok66/docx-master.git (push)
origin  https://Zok66:<TOKEN>@gitee.com/Zok66/docx-master.git (push)
```

> **为什么不用两个 remote？**
> 也可以 `git remote add gitee ...` 然后分别推，但双 push 地址能让
> `git push origin main` 一次推两处，减少漏推。取舍：单点失败会整体报错（第一条成功、
> 第二条失败时，GitHub 已更新而 Gitee 没有，需重推）。

---

## 3. 同步分支与标签

### 3.1 常规推送（含 tag）

```bash
# 当前分支 + 所有 tag
git -c http.sslBackend=openssl -c http.sslVerify=false push origin main pro --follow-tags
```

**为什么要加 `-c http.sslBackend=openssl -c http.sslVerify=false`？**
Windows 的 Git 默认用 schannel，访问 `https://gitee.com` 时报
`CRYPT_E_NO_REVOCATION_CHECK`（证书吊销检查不可达）。这条命令**只对本次调用**切换
OpenSSL 后端并跳过校验，不改全局配置。

GitHub 走 SSH，不受此参数影响。

### 3.2 全部分支

```bash
git push origin --all              # 推所有本地分支
git push origin --tags             # 单独补推 tag（--all 不含 tag）
```

### 3.3 删除同步

远端有、本地已删的分支不会自动消失：

```bash
git push origin --delete <branch>          # 逐个删
git remote prune origin                    # 清理过期的远程跟踪引用
```

---

## 4. 同步 Releases（必须走 API）

git 协议不传 Release。Gitee 的 Releases 只能通过 REST API 创建。

### 4.1 拉取 GitHub 的 Release 清单

```bash
gh api repos/Zok66/docx-master/releases --paginate \
  --jq '.[] | {tag: .tag_name, name: .name, body: .body, prerelease: .prerelease, draft: .draft}'
```

### 4.2 在 Gitee 上创建

```bash
curl -X POST "https://gitee.com/api/v5/repos/$GITEE_OWNER/$REPO/releases" \
  -d "access_token=$GITEE_TOKEN" \
  -d "tag_name=v0.16.0" \
  -d "target_commitish=main" \
  --data-urlencode "name=v0.16.0 标准版 —— 日常办公 / 论文 / 实训报告" \
  --data-urlencode "body=<发布说明 Markdown>"
```

关键参数：

| 字段 | 说明 |
|---|---|
| `tag_name` | **必须已存在于 Gitee**（先推 tag，再建 Release；否则会失败） |
| `target_commitish` | 分支名或 commit sha；写分支名便于阅读 |
| `name` | 标题 |
| `body` | Markdown 正文，用 `--data-urlencode` 避免特殊字符被截断 |
| `prerelease` | `true` / `false`，与 GitHub 对齐 |

### 4.3 字段差异（重要）

| GitHub 有 | Gitee 有 | 影响 |
|---|---|---|
| `tag_name` `name` `body` `prerelease` `target_commitish` `assets` `created_at` | ✅ | 可直接映射 |
| `draft` | ❌ | Gitee 无草稿概念，草稿版应跳过不推 |
| `latest`（Latest 标记） | ❌ | Gitee 按创建时间倒序显示，**最新创建的在最前**。若 GitHub 的 Latest 不是最新创建的，Gitee 无法精确复刻 |
| `published_at` | ❌ | 用 `created_at` 代替 |

> **Latest 的处理**：Gitee 没有 Latest 字段。若想让它显示“最新”，只能**最后创建**
> 你想置顶的那个版本。因此推 Release 时要按“从旧到新”的顺序，把目标 Latest 放最后。

### 4.4 幂等：已存在则更新

```bash
# 查是否已存在
curl -s "https://gitee.com/api/v5/repos/$GITEE_OWNER/$REPO/releases/tags/v0.16.0?access_token=$GITEE_TOKEN"

# 存在则 PATCH（release_id 从上一步取 .id）
curl -X PATCH "https://gitee.com/api/v5/repos/$GITEE_OWNER/$REPO/releases/$RELEASE_ID" \
  -d "access_token=$GITEE_TOKEN" \
  --data-urlencode "name=<新标题>" \
  --data-urlencode "body=<新正文>"
```

### 4.5 完整同步脚本

仓库里已有可直接运行的版本：[`scripts/sync-releases.sh`](scripts/sync-releases.sh)。核心逻辑：

```bash
#!/usr/bin/env bash
set -euo pipefail
: "${GITEE_TOKEN:?需要 export GITEE_TOKEN}"

GH_REPO="Zok66/docx-master"
GE_REPO="Zok66/docx-master"

# 从旧到新处理，让最新创建的排在 Gitee 最前；用 python 解析，不依赖 jq
gh api "repos/$GH_REPO/releases" --paginate \
| python -c '
import sys, json
for r in reversed(json.load(sys.stdin)):
    print(json.dumps({"tag": r["tag_name"], "name": r.get("name") or r["tag_name"],
                      "body": r.get("body") or "", "pre": bool(r.get("prerelease")),
                      "draft": bool(r.get("draft")),
                      "target": r.get("target_commitish") or "main"}, ensure_ascii=False))
' | while IFS= read -r line; do
  f() { printf '%s' "$line" | python -c 'import sys,json;v=json.load(sys.stdin).get(sys.argv[1],"");print("True" if v is True else ("False" if v is False else v))' "$1"; }

  tag=$(f tag); [ -z "$tag" ] && continue
  [ "$(f draft)" = "True" ] && { echo "跳过草稿 $tag"; continue; }

  # 是否已存在
  rid=$(curl -s "https://gitee.com/api/v5/repos/$GE_REPO/releases/tags/$tag?access_token=$GITEE_TOKEN" \
        | python -c 'import sys,json
try: print(json.load(sys.stdin).get("id",""))
except Exception: print("")')

  if [ -n "$rid" ]; then
    curl -s -X PATCH "https://gitee.com/api/v5/repos/$GE_REPO/releases/$rid" \
      -d "access_token=$GITEE_TOKEN" -d "prerelease=$(f pre)" \
      --data-urlencode "name=$(f name)" --data-urlencode "body=$(f body)" >/dev/null
    echo "更新 $tag"
  else
    curl -s -X POST "https://gitee.com/api/v5/repos/$GE_REPO/releases" \
      -d "access_token=$GITEE_TOKEN" -d "tag_name=$tag" \
      -d "target_commitish=$(f target)" -d "prerelease=$(f pre)" \
      --data-urlencode "name=$(f name)" --data-urlencode "body=$(f body)" >/dev/null
    echo "创建 $tag"
  fi
done
```

> **`--data-urlencode` 不能省**：Release 正文是 Markdown，含 `**`、`|`、`#`、换行，
> 直接 `-d "body=$body"` 会被截断或报错。

---

## 5. 同步仓库元信息（license / 可见性 / 默认分支 / 描述）

这些不是 git 对象，全走 `PATCH /repos/{owner}/{repo}`。

```bash
curl -X PATCH "https://gitee.com/api/v5/repos/$GITEE_OWNER/$REPO" \
  -d "access_token=$GITEE_TOKEN" \
  -d "name=$REPO" \
  -d "private=false" \
  -d "license=MIT" \
  -d "default_branch=main" \
  --data-urlencode "description=直接操作 OOXML 的 Word 文档技能（标准版 / pro 版）"
```

### ⚠️ 三个必须知道的坑

1. **建仓时 `private=false` 不生效** —— `POST /user/repos` 会返回成功，但仓库仍是私有的。
   必须再发一次 `PATCH` 才能真正公开。
2. **PATCH 必须带 `name`** —— 缺了会报 `{"messages":["name is missing"]}`。
3. **转公开必须先设 license** —— Gitee 要求公开仓库选定开源许可证，`PATCH` 里
   `private=false` 与 `license=MIT` 要同时发。license 取值可查：

   ```bash
   curl -s "https://gitee.com/api/v5/licenses" | jq -r '.[].key' | grep MIT
   ```

### GitHub 侧对应查询

```bash
gh api repos/Zok66/docx-master \
  --jq '{private, license: .license.spdx_id, default: .default_branch, desc: .description}'
```

---

## 6. 校验：同步是否完整

同步后**必须双向核对**，不能只看 push 命令有没有报错。仓库已有可直接运行的版本：
[`scripts/verify-sync.sh`](scripts/verify-sync.sh)。

```bash
export GITEE_TOKEN='xxx'
bash scripts/verify-sync.sh
```

输出示例（本仓库实测）：

```
校验 Zok66/docx-master ←→ Zok66/docx-master
  ✅ 分支   一致 (2 项)
  ✅ Tag      一致 (4 项)
  ✅ Release  一致 (2 项)
  ✅ 元信息 一致（False MIT main）
  ✅ SHA      main 同一提交 (e7d395f)

全部一致。
```

退出码 **0 = 一致，1 = 有差异**，可直接接进 CI 或发版脚本。

它比对 5 项：

| 检查项 | 方法 |
|---|---|
| 分支 | GitHub `/branches` vs Gitee `/branches` 的 name 集合 |
| Tag | 同上，`/tags` |
| Release | 同上，`/releases` 的 `tag_name` 集合 |
| 元信息 | `private` / `license` / `default_branch` 三元组 |
| SHA | 默认分支是否指向同一 commit（比 tag 名更严格） |

**SHA 级手工核对**（怀疑 tag 被移动时用）：

```bash
git rev-parse main
git ls-remote origin main | cut -c1-40
curl -s "https://gitee.com/api/v5/repos/Zok66/docx-master/branches/main?access_token=$GITEE_TOKEN" \
  | python -c 'import sys,json;print(json.load(sys.stdin)["commit"]["sha"])'
```

> **Windows 上的坑**：`curl` / `gh` 输出带 `\r`，直接 `diff` 会把相同的两行判成不同
> （现象：`1c1 < v0.16.0 --- > v0.16.0`，看着一样却报错）。脚本里已统一 `tr -d '\r'`。

---

## 7. 令牌安全建议

当前状态：令牌明文存在 `.git/config` 的 remote URL 里（不被跟踪、不进版本库，但落盘明文）。

更稳妥的方案：

```bash
# 方案 A：系统凭据存储（本机沙盒会拦截写 ~/.git-credentials，需在沙盒外执行）
git config --global credential.helper store
printf 'https://<user>:<token>@gitee.com\n' >> ~/.git-credentials
chmod 600 ~/.git-credentials
# 之后 remote 用不含凭据的 URL
git remote set-url --add --push origin https://gitee.com/Zok66/docx-master.git

# 方案 B：每次推送临时注入（最干净，但命令变长）
GITEE_TOKEN=xxx git -c credential.helper='!f(){ echo "username=Zok66"; echo "password=$GITEE_TOKEN"; };f' push ...
```

**若令牌曾出现在聊天记录 / 命令行历史里，应当作已泄露处理，立即吊销重发。**

---

## 8. 每次发版的完整流程（日常操作清单）

```bash
# 1) 代码就绪后提交
git add -A && git commit -m "feat: ..."

# 2) 推代码 + tag（一次推两平台）
git tag -a v0.17.0 -m "v0.17.0 ..."
git -c http.sslBackend=openssl -c http.sslVerify=false push origin main pro --follow-tags

# 3) GitHub 建 Release（gh 最容易）
gh release create v0.17.0 --target main \
  --title "v0.17.0 ..." --notes-file notes.md

# 4) 镜像 Release 到 Gitee（脚本见 §4.5）
./scripts/sync-releases.sh

# 5) 核对（脚本见 §6）
./scripts/verify-sync.sh
```

**顺序不能乱**：tag 必须先到 Gitee，Release 才能挂上去。

---

## 9. 故障排查

| 现象 | 原因 | 处理 |
|---|---|---|
| `CRYPT_E_NO_REVOCATION_CHECK` | Windows schannel 无法校验 Gitee 证书 | 加 `-c http.sslBackend=openssl -c http.sslVerify=false` |
| `Everything up-to-date` 但 Gitee 没更新 | 双 push 地址中 Gitee 那条失败被吞 | 单独 `git push <gitee-url> main` 看完整报错 |
| push 成功但 `git status` 仍显示 ahead | 沙盒环境远程跟踪引用写入失败 | 以 `git ls-remote origin main` 为准；必要时 `printf '<sha>\n' > .git/refs/remotes/origin/main` |
| Gitee Release 建不上，报 tag 不存在 | 先建了 Release、tag 还没推 | 先 `git push origin --tags` |
| `private` 改不动 | PATCH 缺 `name` 或 `license` | 按 §5 三件套一起发 |
| Gitee 显示的顺序与 GitHub 的 Latest 不符 | Gitee 无 `latest` 字段 | 把要置顶的版本**最后创建** |
| 令牌失效 / 401 | 已吊销或过期 | 重新生成并更新 `.git/config` |
| `jq: command not found` | 本机未装 jq | 用配套脚本（已改 Python 解析）；或 `choco install jq` / `scoop install jq` |
| diff 显示两边内容一样却报不一致 | Windows 输出带 `\r` | `diff <(a \| tr -d '\r') <(b \| tr -d '\r')` |
| 仓库一直显示私有 | `POST` 建仓时 `private=false` 无效 | 再发一次 `PATCH`（§5 三件套） |

---

## 附：本仓库当前同步状态（2026-10-20 实测 `verify-sync.sh`）

```
校验 Zok66/docx-master ←→ Zok66/docx-master
  ✅ 分支   一致 (2 项)        main / pro
  ✅ Tag      一致 (4 项)      v0.15.0 · v0.15.0-pro · v0.16.0 · v0.16.0-pro
  ✅ Release  一致 (2 项)      v0.16.0 · v0.16.0-pro
  ✅ 元信息 一致（False MIT main）
  ✅ SHA      main 同一提交 (e7d395f)
```

| 项目 | GitHub | Gitee | 状态 |
|---|---|---|---|
| 分支 | `main` `pro` | `main` `pro` | ✅ |
| Tag | 4 个 | 4 个 | ✅ |
| Release | v0.16.0（Latest）、v0.16.0-pro | v0.16.0、v0.16.0-pro | ✅ |
| 可见性 | public | public | ✅ |
| License | MIT | MIT | ✅ |
| 默认分支 | main | main | ✅ |
| 描述 | 空 | 已填 | ⚠️ GitHub 侧可补 |
| Issues / PR / Wiki / Actions | 有 | 无对应同步 | ❌ 平台差异 |

**已知平台差异（无需强求同步）**：

- **Release 的 Latest 标记**：GitHub 有，Gitee 没有（只能靠创建顺序控制谁在最前）
- **社区健康标签**：GitHub 会在仓库页点亮 "Code of conduct / Contributing / Security"，
  Gitee 不做这套检测——文件都在，只是没有入口标签
- **Actions**：GitHub 的 workflow 在 Gitee 不会运行，需要另配 GoGo 流水线
