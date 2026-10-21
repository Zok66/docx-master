#!/usr/bin/env bash
# verify-sync.sh —— 比对 GitHub 与 Gitee 的分支 / tag / Release / 元信息
#
# 用法：
#   export GITEE_TOKEN='xxx'
#   ./scripts/verify-sync.sh [gh_owner/repo] [gitee_owner/repo]
#
# 依赖：gh（已登录）、curl、python（无 jq 依赖）
# 退出码：0 = 全部一致；1 = 有差异
set -uo pipefail

GH="${1:-Zok66/docx-master}"
GE="${2:-Zok66/docx-master}"
: "${GITEE_TOKEN:?需要 export GITEE_TOKEN}"

API="https://gitee.com/api/v5/repos/$GE"
PY="${PYTHON:-python}"
fail=0

# jq 的轻量替代：从 stdin 读 JSON，按点号路径逐键取值。
#   pyj name          → 取顶层 name（若顶层是数组，则对每项取 name）
#   pyj ""            → 数组展开为每行（标量原样，对象转 JSON）
#   pyj commit.sha    → 嵌套取值
pyj() { "$PY" -c '
import sys, json
path = sys.argv[1] if len(sys.argv) > 1 else ""
try:
    d = json.load(sys.stdin)
except Exception:
    sys.exit(0)

def walk(obj):
    """按路径逐层下钻；遇到列表则对每项分别下钻后拼接。"""
    if not path:
        return [obj]
    out = []
    for k in [p for p in path.split(".") if p]:
        nxt = []
        for o in (obj if isinstance(obj, list) else [obj]):
            if isinstance(o, dict) and k in o:
                nxt.append(o[k])
            elif isinstance(o, list):
                nxt.extend(x.get(k) if isinstance(x, dict) else None for x in o)
        obj = nxt
    return obj if isinstance(obj, list) else [obj]

for v in walk(d):
    if v is None:
        continue
    print(v if not isinstance(v, (dict, list)) else json.dumps(v, ensure_ascii=False))
' "$@"; }

ge_get() { curl -s -m 30 "$API/$1" -G -d "access_token=$GITEE_TOKEN"; }

cmp_list() {
  local label=$1 gh_list=$2 ge_list=$3
  # 统一去 CR：Windows 下 curl/gh 输出可能带 \r，否则内容相同也会 diff 报差异
  if diff <(echo "$gh_list" | tr -d '\r') <(echo "$ge_list" | tr -d '\r') > /tmp/sync-diff.txt 2>&1; then
    printf '  ✅ %-8s 一致 (%d 项)\n' "$label" "$(echo "$gh_list" | grep -c . || true)"
  else
    printf '  ❌ %-8s 不一致：\n' "$label"
    sed 's/^/       /' /tmp/sync-diff.txt
    fail=1
  fi
}

echo "校验 $GH ←→ $GE"

cmp_list "分支" \
  "$(gh api "repos/$GH/branches" --paginate --jq '.[].name' | sort)" \
  "$(ge_get branches | pyj name | sort)"

cmp_list "Tag" \
  "$(gh api "repos/$GH/tags" --paginate --jq '.[].name' | sort)" \
  "$(ge_get tags | pyj name | sort)"

cmp_list "Release" \
  "$(gh api "repos/$GH/releases" --paginate --jq '.[].tag_name' | sort)" \
  "$(ge_get releases | pyj tag_name | sort)"

# 元信息：private / license / default_branch
gh_json=$(gh api "repos/$GH" 2>/dev/null)
gh_meta=$(echo "$gh_json" | "$PY" -c 'import sys,json;d=json.load(sys.stdin);print(d.get("private"), (d.get("license") or {}).get("spdx_id","-"), d.get("default_branch"))' 2>/dev/null)
ge_json=$(ge_get "")
ge_meta=$(echo "$ge_json" | "$PY" -c 'import sys,json;d=json.load(sys.stdin);print(d.get("private"), d.get("license") or "-", d.get("default_branch"))' 2>/dev/null)

if [ -n "$gh_meta" ] && [ "$gh_meta" = "$ge_meta" ]; then
  printf '  ✅ %-8s 一致（%s）\n' "元信息" "$gh_meta"
else
  printf '  ❌ %-8s 不一致 → GH:[%s] GE:[%s]\n' "元信息" "$gh_meta" "$ge_meta"
  fail=1
fi

# SHA 级核对：默认分支是否指向同一 commit
def_branch=$(echo "$gh_meta" | awk '{print $3}')
def_branch=${def_branch:-main}
gh_sha=$(gh api "repos/$GH/commits/$def_branch" --jq '.sha' 2>/dev/null | cut -c1-7)
ge_sha=$(ge_get "branches/$def_branch" | "$PY" -c 'import sys,json;d=json.load(sys.stdin);print((d.get("commit") or {}).get("sha","")[:7])' 2>/dev/null)
if [ -n "$gh_sha" ] && [ "$gh_sha" = "$ge_sha" ]; then
  printf '  ✅ %-8s %s 同一提交 (%s)\n' "SHA" "$def_branch" "$gh_sha"
else
  printf '  ⚠️  %-8s GH:[%s] GE:[%s] —— 若分支名不同属正常\n' "SHA" "$gh_sha" "$ge_sha"
fi

echo
[ "$fail" -eq 0 ] && echo "全部一致。" || echo "存在差异，见上。"
exit "$fail"
