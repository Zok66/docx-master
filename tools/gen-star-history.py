#!/usr/bin/env python3
"""生成 docx-master 的 Star History SVG（浅色 / 深色两份）。

每 12 小时由 .github/workflows/star-history.yml 拉取 GitHub stargazers API
并重绘；曲线按当前保留的星标汇总（stargazers API 含取关后重星的历史）。

用法：GITHUB_TOKEN=xxx python tools/gen-star-history.py [owner/repo]
输出：assets/star-history-light.svg / assets/star-history-dark.svg
"""
import json
import os
import sys
import urllib.request
from datetime import datetime, timedelta, timezone

REPO = sys.argv[1] if len(sys.argv) > 1 else "Zok66/docx-master"
OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "assets")

W, H = 1000, 420
ML, MR, MT, MB = 24, 190, 84, 64          # 边距（右上留给标签）
PW, PH = W - ML - MR, H - MT - MB         # 绘图区


def fetch_stars(repo: str):
    """带 starred_at 时间戳的分页拉取（按星的时间升序）。"""
    token = os.environ.get("GITHUB_TOKEN", "")
    url = f"https://api.github.com/repos/{repo}/stargazers?per_page=100"
    events = []
    page = 1
    while True:
        req = urllib.request.Request(f"{url}&page={page}")
        req.add_header("Accept", "application/vnd.github.star+json")
        if token:
            req.add_header("Authorization", f"Bearer {token}")
        with urllib.request.urlopen(req, timeout=30) as resp:
            batch = json.load(resp)
        if not batch:
            break
        events += [e["starred_at"] for e in batch if e.get("starred_at")]
        page += 1
    events.sort()
    return events


def cumulative(events):
    """[(datetime, count)] —— 每个事件后的累计值。"""
    out, n = [], 0
    for iso in events:
        n += 1
        out.append((datetime.fromisoformat(iso.replace("Z", "+00:00")), n))
    return out


def to_points(series, axis_max):
    """把累计序列换算成 SVG 坐标。

    - 纵向以 axis_max（与右侧刻度一致）归一化，避免曲线被自身峰值顶满格；
    - 起点补一个 0 值台阶，让"0 → 第一颗星"的爬升能看出来。
    """
    if not series:
        t0 = datetime.now(timezone.utc)
        return [(ML, MT + PH), (ML + PW, MT + PH)], (t0, t0)
    t0, t1 = series[0][0], series[-1][0]
    # 时间轴起点提前，给"0 → 第一颗星"的台阶留出可见宽度
    t0 = t0 - timedelta(seconds=3600)
    if (t1 - t0).total_seconds() < 86400:
        t1 = t0.fromtimestamp(t0.timestamp() + 86400, tz=timezone.utc)
    span = (t1 - t0).total_seconds() or 1
    scale = axis_max or 1

    def px(t):
        return ML + (t - t0).total_seconds() / span * PW

    def py(c):
        return MT + PH - min(c / scale, 1) * PH

    # 起手 0 值台阶：从左侧 0 线水平走到第一颗星之前，再竖直抬起
    first_t, first_c = series[0]
    pts = [(ML, py(0)), (px(first_t), py(0))]
    pts.append((px(first_t), py(first_c)))
    for t, c in series[1:]:
        pts.append((px(t), py(c)))
    # 收尾补一个"现在"的点，让曲线顶到右边界
    _, last_c = series[-1]
    pts.append((ML + PW, py(last_c)))
    return pts, (t0, t1)


def nice_axis(peak: int):
    """纵轴向上取整到一个好看的刻度。"""
    for v in (10, 20, 50, 100, 200, 500, 1000, 2000, 5000, 10000):
        if peak <= v:
            return v
    return peak


def render(pts, rng, peak_now, axis_max, dark: bool) -> str:
    if dark:
        bg, grid, line, fill, strong, muted = "none", "#2d3748", "#4a9eff", "#1a365d88", "#e2e8f0", "#a0aec0"
    else:
        bg, grid, line, fill, strong, muted = "none", "#e2e8f0", "#3182ce", "#bee3f888", "#1a202c", "#4a5568"

    t0, t1 = rng
    grid_lines = []
    for i in range(5):
        y = MT + PH * i / 4
        val = axis_max * (1 - i / 4)
        grid_lines.append(
            f'<line x1="{ML}" y1="{y:.0f}" x2="{ML+PW}" y2="{y:.0f}" stroke="{grid}" stroke-width="1"/>'
            f'<text x="{ML+PW+14}" y="{y+5:.0f}" font-size="15" fill="{muted}">{val}</text>'
        )
    for fx, label in ((0, t0.strftime("%m-%d")), (0.5, ""), (1, t1.strftime("%m-%d"))):
        x = ML + PW * fx
        grid_lines.append(
            f'<text x="{x}" y="{MT+PH+30}" font-size="15" fill="{muted}" text-anchor="{"start" if fx == 0 else "end"}">{label}</text>'
        )

    poly = " ".join(f"{x:.1f},{y:.1f}" for x, y in pts)
    area = f"{ML},{MT+PH} " + poly + f" {ML+PW},{MT+PH}"
    asof = datetime.now(timezone.utc).strftime("%Y-%m-%d")

    return f'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {W} {H}" font-family="-apple-system, 'Segoe UI', 'Microsoft YaHei', sans-serif">
<rect width="{W}" height="{H}" fill="{bg}"/>
<text x="{ML}" y="44" font-size="26" font-weight="600" fill="{strong}">Star History</text>
<text x="{ML}" y="68" font-size="14" fill="{muted}">{REPO}</text>
<text x="{W-MR+14}" y="44" font-size="15" fill="{muted}">{peak_now} stars</text>
{''.join(grid_lines)}
<polygon points="{area}" fill="{fill}"/>
<polyline points="{poly}" fill="none" stroke="{line}" stroke-width="2.5" stroke-linejoin="round"/>
<text x="{ML}" y="{H-14}" font-size="13" fill="{muted}">数据截至 {asof} · GitHub stargazers API</text>
</svg>'''


def main():
    events = fetch_stars(REPO)
    series = cumulative(events)
    peak_now = series[-1][1] if series else 0
    axis_max = nice_axis(peak_now)
    pts, rng = to_points(series, axis_max)
    os.makedirs(OUT_DIR, exist_ok=True)
    for name, dark in (("star-history-light.svg", False), ("star-history-dark.svg", True)):
        path = os.path.join(OUT_DIR, name)
        with open(path, "w", encoding="utf-8") as f:
            f.write(render(pts, rng, peak_now, axis_max, dark))
        print(f"{path}  ({peak_now} stars, 轴顶 {axis_max})")


if __name__ == "__main__":
    main()
