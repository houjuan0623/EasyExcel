# Copyright (c) 2026 AbeLin · MIT License
"""固定 16:9 看板的布局校验和离线 HTML；不改变分析数据。"""
import re
from html import escape
from pathlib import Path


def dashboard_cards(request):
    """校验新方案并为旧网格方案补齐位置，在产生导出文件前拒绝无效布局。"""
    cards = request.get("cards")
    version = request.get("version", 1)
    if type(version) is not int or version not in (1, 2):
        raise ValueError("分析方案版本无效")
    if not isinstance(cards, list) or not 1 <= len(cards) <= 12:
        raise ValueError("一个 dashboard 可包含 1～12 个图表")
    if version == 2 and request.get("canvas") != {"width": 1920, "height": 1080}:
        raise ValueError("看板画布必须为 1920 × 1080（16:9）")
    cols = 1 if len(cards) == 1 else 2 if len(cards) <= 4 else 3 if len(cards) <= 9 else 4
    rows = (len(cards) + cols - 1) // cols
    width, height = (1872 - 24 * (cols - 1)) // cols, (936 - 24 * (rows - 1)) // rows
    result, seen = [], set()
    fields = ("inputPath", "blocks", "token", "kind", "title", "unit", "note", "combine", "joinKeys", "joinType",
              "x", "y", "aggregation", "sort", "timeBucket", "bins", "limit", "fieldLabels")
    for index, card in enumerate(cards):
        if not isinstance(card, dict) or not isinstance(card.get("config"), dict):
            raise ValueError("图表配置无效")
        identifier = card.get("id")
        if (not isinstance(identifier, str) or not re.fullmatch(r"[a-zA-Z0-9_-]{1,80}", identifier)
                or identifier in {"__proto__", "constructor", "prototype"} or identifier in seen):
            raise ValueError("图表标识无效或重复")
        seen.add(identifier)
        layout = ({"x": 24 + index % cols * (width + 24), "y": 120 + index // cols * (height + 24),
                   "width": width, "height": height} if version == 1 else card.get("layout"))
        if not isinstance(layout, dict) or any(type(layout.get(k)) is not int for k in ("x", "y", "width", "height")):
            raise ValueError("图表位置和尺寸必须是整数")
        x, y, w, h = (layout[k] for k in ("x", "y", "width", "height"))
        if x < 24 or y < 120 or w < 240 or h < 168 or x + w > 1896 or y + h > 1056:
            raise ValueError("图表位置或尺寸超出画布范围")
        result.append({"id": identifier, "config": {k: card["config"][k] for k in fields if k in card["config"]},
                       "layout": {"x": x, "y": y, "width": w, "height": h}})
    return result


def dashboard_html(title, cards, results, html_table):
    """SVG viewBox 统一缩放字体和图表；无 JavaScript、网络请求或外部资源。"""
    tiles, details = [], []
    for index, (card, result) in enumerate(zip(cards, results, strict=True), 1):
        layout = card["layout"]
        style = (f"left:{layout['x']}px;top:{layout['y']}px;width:{layout['width']}px;"
                 f"height:{layout['height']}px;z-index:{index}")
        visual = (f'<img src="{escape(result["image"], quote=True)}" alt="{escape(result["title"], quote=True)}">'
                  if result["image"] else '<div class="dashboard-metric"><strong>'
                  + ("—" if result["metric"] is None else f'{result["metric"]:,.4f}'.rstrip("0").rstrip("."))
                  + '</strong><span>' + escape(result["unit"]) + '</span></div>')
        note = '<p class="dashboard-tile-note">' + escape(result["note"]) + '</p>' if result["note"] else ''
        count = len(result["warnings"])
        warning = f'<p class="dashboard-tile-status">{count} 条数据处理说明 · 见下方明细</p>' if count else ''
        tiles.append(f'<article class="dashboard-tile" data-card-id="{card["id"]}" style="{style}">'
                     f'<header class="dashboard-tile-header"><h3><a href="#chart-{index}">{escape(result["title"])}</a></h3></header>'
                     f'<div class="dashboard-tile-body">{visual}</div>{note}{warning}</article>')
        source = "；".join(f"{s['file']} / {s['sheet']}!{s['range']}（{s['rows']} 条记录）" for s in result["sources"])
        warnings = "".join("<li>" + escape(w) + "</li>" for w in result["warnings"])
        details.append(f'<section id="chart-{index}" class="report-detail"><h2>{index}. {escape(result["title"])}</h2>'
                       f'<p>{escape(result["method"])}</p><p class="note">{escape(result["note"])}</p>'
                       f'<p class="source">{escape(source)}</p><ul>{warnings}</ul>'
                       '<details><summary>图表数据与描述统计</summary>' + html_table(result["table"])
                       + html_table(result["statistics"]) + '</details></section>')
    css = (Path(__file__).parent / "shared" / "dashboard-canvas.css").read_text(encoding="utf-8")
    return ("""<!doctype html><html lang="zh-CN"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; style-src 'unsafe-inline'"><title>"""
            + escape(title) + "</title><style>" + css + """
body{margin:0;padding:24px;background:#e4ebe7;color:#284c3a;font:14px 'Segoe UI','Microsoft YaHei',sans-serif}
.report-intro{margin:0 0 18px;line-height:1.8;color:#627f6c;font-size:12px}
.report-stage{display:block;width:100%;height:auto;aspect-ratio:16/9;border-radius:8px;overflow:hidden}
.dashboard-tile a{color:inherit;text-decoration:none}.dashboard-tile a:hover{text-decoration:underline}
.report-details{max-width:1200px;margin:24px auto;display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:20px}
.report-detail{padding:22px;background:white;border:1px solid #d7e3db;border-radius:14px;min-width:0}
.report-detail h2{font-size:18px}.source,li{font-size:12px;color:#6b7a70}.report-detail p{line-height:1.8}.note{white-space:pre-wrap}
.scroll{overflow:auto}table{border-collapse:collapse;white-space:nowrap;font-size:12px;margin:14px 0}th,td{padding:7px;border-bottom:1px solid #e0e8e2}summary{cursor:pointer}
@media(max-width:700px){body{padding:12px}.report-details{grid-template-columns:1fr}}
@media print{body{padding:0;background:white}.report-stage{break-after:page}.report-details{display:block}.report-detail{break-inside:avoid}details{display:block}}
</style><p class="report-intro">EASYEXCEL · 16:9 本地统计看板 · 点击图表标题查看明细。此报告包含真实统计结果。公式使用工作簿已保存的缓存，未进行重算。</p>
<svg class="report-stage" viewBox="0 0 1920 1080" preserveAspectRatio="xMidYMid meet" role="group" aria-label="16:9 统计看板"><foreignObject width="1920" height="1080"><div class="dashboard-artboard"><header class="dashboard-canvas-heading"><span class="dashboard-brand">EASYEXCEL</span><h1 class="dashboard-title">"""
            + escape(title) + '</h1></header>' + "".join(tiles)
            + '</div></foreignObject></svg><main class="report-details">' + "".join(details) + '</main></html>')
