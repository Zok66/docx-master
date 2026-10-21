#!/usr/bin/env bash
# sync-releases.sh —— 把 GitHub Release 镜像到 Gitee（幂等）
#
# 用法：
#   export GITEE_TOKEN='xxx'
#   ./scripts/sync-releases.sh [gh_owner/repo] [gitee_owner/repo]
#
# 默认镜像 Zok66/docx-master。Gitee 无 latest 字段，按"从旧到新"创建，
# 所以最后一个创建的会显示在最前。
# 依赖：gh（已登录）、curl、python —— 不依赖 jq。
set -euo pipefail

GH_REPO="${1:-Zok66/docx-master}"
GE_REPO="${2:-Zok66/docx-master}"
: "${GITEE_TOKEN:?需要 export GITEE_TOKEN（Gitee 私人令牌，含 projects 权限）}"

command -v gh   >/dev/null || { echo "缺少 gh CLI"; exit 1; }
command -v curl >/dev/null || { echo "缺少 curl";   exit 1; }
PY="${PYTHON:-python}"
"$PY" -c "" 2>/dev/null || { echo "缺少 python"; exit 1; }

echo "镜像 $GH_REPO → $GE_REPO"

# 取 GitHub Release 清单，每条输出一行 JSON（从旧到新）
gh api "repos/$GH_REPO/releases" --paginate 2>/dev/null \
  | "$PY" -c '
import sys, json
data = json.load(sys.stdin)
for r in reversed(data):
    print(json.dumps({
        "tag": r.get("tag_name", ""),
        "name": r.get("name") or r.get("tag_name", ""),
        "body": r.get("body") or "",
        "prerelease": bool(r.get("prerelease")),
        "draft": bool(r.get("draft")),
        "target": r.get("target_commitish") or "main",
    }, ensure_ascii=False))
' | while IFS= read -r line; do
  # 用 python 逐字段取值，避免 shell 分词破坏中文与换行
  field() {
    printf '%s' "$line" | "$PY" -c '
import sys, json
d = json.load(sys.stdin)
v = d.get(sys.argv[1], "")
print("True" if v is True else ("False" if v is False else v))
' "$1"
  }

  tag=$(field tag)
  [ -z "$tag" ] && continue

  if [ "$(field draft)" = "True" ]; then
    echo "  跳过草稿  $tag"
    continue
  fi

  rid=$(curl -s -m 30 \
    "https://gitee.com/api/v5/repos/$GE_REPO/releases/tags/$tag?access_token=$GITEE_TOKEN" \
    | "$PY" -c '
import sys, json
try:
    print(json.load(sys.stdin).get("id", ""))
except Exception:
    print("")')

  name=$(field name)
  body=$(field body)
  pre=$(field prerelease)
  pre=$([ "$pre" = "True" ] && echo true || echo false)
  target=$(field target)

  if [ -n "$rid" ]; then
    curl -s -m 30 -X PATCH \
      "https://gitee.com/api/v5/repos/$GE_REPO/releases/$rid" \
      -d "access_token=$GITEE_TOKEN" \
      -d "prerelease=$pre" \
      --data-urlencode "name=$name" \
      --data-urlencode "body=$body" > /dev/null
    echo "  更新      $tag"
  else
    curl -s -m 30 -X POST "https://gitee.com/api/v5/repos/$GE_REPO/releases" \
      -d "access_token=$GITEE_TOKEN" \
      -d "tag_name=$tag" \
      -d "target_commitish=$target" \
      -d "prerelease=$pre" \
      --data-urlencode "name=$name" \
      --data-urlencode "body=$body" > /dev/null
    echo "  创建      $tag"
  fi
done

echo "完成。用 scripts/verify-sync.sh 核对。"
