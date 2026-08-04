# -*- coding: utf-8 -*-
"""本地 Markdown 知识库 — 类似 Obsidian，支持 [[双向链接]]、搜索、编辑"""

import os
import re
import json
from datetime import datetime
from pathlib import Path
from flask import Flask, render_template_string, request, jsonify, redirect, url_for
import markdown as md_lib

BASE_DIR = Path(__file__).parent
NOTES_DIR = BASE_DIR / "notes"
NOTES_DIR.mkdir(exist_ok=True)

app = Flask(__name__)

# ── Wikilink 正则 ──
WIKILINK_RE = re.compile(r"\[\[([^\]]+?)\]\]")

# ── Markdown 渲染 + wikilink 处理 ──
def render_markdown(text: str, current_note: str = "") -> str:
    """渲染 Markdown，将 [[xxx]] 转为可点击链接"""
    def replace_wikilink(m):
        target = m.group(1).strip()
        # 处理 [[笔记名|显示文字]]
        if "|" in target:
            name, alias = target.split("|", 1)
        else:
            name, alias = target, target
        name = name.strip()
        alias = alias.strip()
        slug = note_slug(name)
        css_class = "wikilink" if note_exists_by_slug(slug) else "wikilink missing"
        return f'<a href="/note/{slug}" class="{css_class}">{alias}</a>'
    
    html = re.sub(WIKILINK_RE, replace_wikilink, text)
    # 使用 markdown 库渲染
    return md_lib.markdown(html, extensions=["fenced_code", "tables", "nl2br"])

# ── 笔记文件名处理 ──
def note_slug(title: str) -> str:
    """标题转文件名"""
    # 保留中文、字母、数字，其他转 -
    slug = re.sub(r"[^\u4e00-\u9fa5a-zA-Z0-9]+", "-", title).strip("-")
    return slug or "untitled"

def note_exists_by_slug(slug: str) -> bool:
    return (NOTES_DIR / f"{slug}.md").exists()

def get_all_notes() -> list[dict]:
    """获取所有笔记"""
    notes = []
    for f in sorted(NOTES_DIR.glob("*.md"), key=lambda x: x.stat().st_mtime, reverse=True):
        content = f.read_text(encoding="utf-8")
        title = f.stem
        # 从内容第一行 # 提取实际标题
        first_line = content.strip().split("\n")[0]
        if first_line.startswith("# "):
            title = first_line[2:].strip()
        notes.append({
            "slug": f.stem,
            "title": title,
            "mtime": datetime.fromtimestamp(f.stat().st_mtime).strftime("%Y-%m-%d %H:%M"),
            "size": len(content),
        })
    return notes

def get_backlinks(slug: str) -> list[dict]:
    """查找引用了当前笔记的其他笔记"""
    backlinks = []
    for f in NOTES_DIR.glob("*.md"):
        if f.stem == slug:
            continue
        content = f.read_text(encoding="utf-8")
        pattern = re.compile(rf"\[\[{re.escape(slug)}(\||\]\])")
        # 也尝试匹配显示标题
        # 简单匹配
        if f"[[{slug}" in content:
            title = f.stem
            first_line = content.strip().split("\n")[0]
            if first_line.startswith("# "):
                title = first_line[2:].strip()
            backlinks.append({"slug": f.stem, "title": title})
    return backlinks

# ── 路由 ──
@app.route("/")
def index():
    notes = get_all_notes()
    return render_template_string(INDEX_TEMPLATE, notes=notes, note=None, 
                                   content_html="", backlinks=[], 
                                   search_results=None, query="")

@app.route("/note/<slug>")
def view_note(slug):
    note_path = NOTES_DIR / f"{slug}.md"
    if not note_path.exists():
        return "笔记不存在", 404
    
    content = note_path.read_text(encoding="utf-8")
    title = slug
    first_line = content.strip().split("\n")[0]
    if first_line.startswith("# "):
        title = first_line[2:].strip()
    
    content_html = render_markdown(content, current_note=slug)
    backlinks = get_backlinks(slug)
    notes = get_all_notes()
    
    note = {"slug": slug, "title": title, "content": content}
    return render_template_string(INDEX_TEMPLATE, notes=notes, note=note,
                                   content_html=content_html, backlinks=backlinks,
                                   search_results=None, query="")

@app.route("/search")
def search():
    query = request.args.get("q", "").strip()
    if not query:
        return redirect("/")
    
    results = []
    for f in NOTES_DIR.glob("*.md"):
        content = f.read_text(encoding="utf-8")
        if query.lower() in content.lower():
            title = f.stem
            first_line = content.strip().split("\n")[0]
            if first_line.startswith("# "):
                title = first_line[2:].strip()
            # 找关键词上下文
            idx = content.lower().find(query.lower())
            start = max(0, idx - 40)
            end = min(len(content), idx + len(query) + 80)
            snippet = content[start:end].replace("\n", " ")
            results.append({
                "slug": f.stem,
                "title": title,
                "snippet": f"...{snippet}...",
            })
    
    notes = get_all_notes()
    return render_template_string(INDEX_TEMPLATE, notes=notes, note=None,
                                   content_html="", backlinks=[],
                                   search_results=results, query=query)

@app.route("/edit/<slug>", methods=["GET", "POST"])
def edit_note(slug):
    if request.method == "POST":
        data = request.get_json()
        content = data.get("content", "")
        note_path = NOTES_DIR / f"{slug}.md"
        note_path.write_text(content, encoding="utf-8")
        return jsonify({"ok": True, "slug": slug})
    
    note_path = NOTES_DIR / f"{slug}.md"
    content = note_path.read_text(encoding="utf-8") if note_path.exists() else f"# {slug}\n\n"
    notes = get_all_notes()
    note = {"slug": slug, "title": slug, "content": content}
    return render_template_string(EDIT_TEMPLATE, notes=notes, note=note)

@app.route("/create", methods=["POST"])
def create_note():
    data = request.get_json()
    title = data.get("title", "").strip()
    if not title:
        return jsonify({"ok": False, "error": "标题不能为空"}), 400
    
    slug = note_slug(title)
    note_path = NOTES_DIR / f"{slug}.md"
    if note_path.exists():
        return jsonify({"ok": False, "error": "笔记已存在"}), 409
    
    note_path.write_text(f"# {title}\n\n", encoding="utf-8")
    return jsonify({"ok": True, "slug": slug})

@app.route("/delete/<slug>", methods=["POST"])
def delete_note(slug):
    note_path = NOTES_DIR / f"{slug}.md"
    if note_path.exists():
        note_path.unlink()
    return jsonify({"ok": True})

@app.route("/api/graph")
def graph_data():
    """返回知识图谱数据"""
    nodes = []
    links = []
    node_set = set()
    
    for f in NOTES_DIR.glob("*.md"):
        slug = f.stem
        if slug not in node_set:
            title = slug
            content = f.read_text(encoding="utf-8")
            first_line = content.strip().split("\n")[0]
            if first_line.startswith("# "):
                title = first_line[2:].strip()
            nodes.append({"id": slug, "label": title, "size": len(content)})
            node_set.add(slug)
        
        # 找链接
        for match in WIKILINK_RE.finditer(content):
            target = match.group(1).split("|")[0].strip()
            target_slug = note_slug(target)
            links.append({"source": slug, "target": target_slug})
    
    return jsonify({"nodes": nodes, "links": links})

# ── 模板 ──
INDEX_TEMPLATE = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>知识库</title>
<style>
    :root {
        --bg: #1e1e2e;
        --sidebar-bg: #181825;
        --text: #cdd6f4;
        --text-muted: #a6adc8;
        --accent: #89b4fa;
        --accent2: #a6e3a1;
        --border: #313244;
        --hover: #45475a;
        --card: #1e1e2e;
        --input-bg: #313244;
        --danger: #f38ba8;
        --wikilink: #89b4fa;
        --wikilink-missing: #f38ba8;
    }
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body { font-family: -apple-system, "Microsoft YaHei", sans-serif; background: var(--bg); color: var(--text); display: flex; height: 100vh; }
    
    /* 侧边栏 */
    .sidebar {
        width: 260px; min-width: 260px; background: var(--sidebar-bg);
        border-right: 1px solid var(--border); display: flex; flex-direction: column; overflow: hidden;
    }
    .sidebar-header {
        padding: 16px; border-bottom: 1px solid var(--border);
        display: flex; gap: 8px; align-items: center;
    }
    .sidebar-header input {
        flex: 1; padding: 8px 12px; border-radius: 6px; border: 1px solid var(--border);
        background: var(--input-bg); color: var(--text); font-size: 14px;
    }
    .sidebar-header button {
        padding: 8px 14px; border-radius: 6px; border: none;
        background: var(--accent); color: #1e1e2e; cursor: pointer; font-size: 14px; font-weight: 600;
    }
    .note-list { flex: 1; overflow-y: auto; padding: 8px; }
    .note-item {
        display: block; padding: 10px 12px; border-radius: 6px; cursor: pointer;
        color: var(--text); text-decoration: none; font-size: 14px;
        margin-bottom: 2px; transition: background .15s;
    }
    .note-item:hover { background: var(--hover); }
    .note-item.active { background: var(--accent); color: #1e1e2e; font-weight: 600; }
    .note-item .time { font-size: 11px; color: var(--text-muted); display: block; }
    .note-item.active .time { color: #1e1e2e; opacity: .7; }
    
    /* 主区域 */
    .main {
        flex: 1; display: flex; flex-direction: column; overflow: hidden;
    }
    .toolbar {
        padding: 12px 20px; border-bottom: 1px solid var(--border);
        display: flex; gap: 10px; align-items: center;
    }
    .toolbar button, .toolbar a {
        padding: 6px 14px; border-radius: 6px; border: 1px solid var(--border);
        background: transparent; color: var(--text); cursor: pointer;
        text-decoration: none; font-size: 13px;
    }
    .toolbar button:hover, .toolbar a:hover { background: var(--hover); }
    .toolbar button.primary { background: var(--accent); color: #1e1e2e; border-color: var(--accent); }
    
    .content-area {
        flex: 1; display: flex; overflow: hidden;
    }
    .note-content {
        flex: 1; padding: 24px 32px; overflow-y: auto; line-height: 1.8;
    }
    .note-content h1 { font-size: 1.8em; margin-bottom: 16px; border-bottom: 1px solid var(--border); padding-bottom: 8px; }
    .note-content h2 { font-size: 1.4em; margin: 20px 0 10px; }
    .note-content h3 { font-size: 1.15em; margin: 16px 0 8px; }
    .note-content p { margin: 8px 0; }
    .note-content ul, .note-content ol { padding-left: 24px; margin: 8px 0; }
    .note-content code {
        background: var(--input-bg); padding: 2px 6px; border-radius: 4px;
        font-family: "JetBrains Mono", "Cascadia Code", monospace; font-size: .9em;
    }
    .note-content pre {
        background: var(--input-bg); padding: 14px 18px; border-radius: 8px;
        overflow-x: auto; margin: 12px 0; border: 1px solid var(--border);
    }
    .note-content pre code { background: none; padding: 0; }
    .note-content blockquote {
        border-left: 3px solid var(--accent); padding: 4px 16px; margin: 12px 0;
        color: var(--text-muted);
    }
    .note-content table { border-collapse: collapse; width: 100%; margin: 12px 0; }
    .note-content th, .note-content td { border: 1px solid var(--border); padding: 8px 12px; text-align: left; }
    .note-content th { background: var(--sidebar-bg); }
    .note-content a { color: var(--accent); text-decoration: none; }
    .note-content a:hover { text-decoration: underline; }
    .note-content a.wikilink { color: var(--wikilink); border-bottom: 1px dashed var(--wikilink); }
    .note-content a.wikilink:hover { border-bottom-style: solid; text-decoration: none; }
    .note-content a.wikilink.missing { color: var(--wikilink-missing); border-color: var(--wikilink-missing); }
    
    /* 反向链接 */
    .backlinks-panel {
        width: 220px; min-width: 220px; border-left: 1px solid var(--border);
        padding: 16px; overflow-y: auto; font-size: 13px;
    }
    .backlinks-panel h3 { font-size: 13px; color: var(--text-muted); margin-bottom: 10px; text-transform: uppercase; letter-spacing: .5px; }
    .backlinks-panel a {
        display: block; padding: 6px 10px; border-radius: 4px; color: var(--accent);
        text-decoration: none; margin-bottom: 2px;
    }
    .backlinks-panel a:hover { background: var(--hover); }
    
    /* 欢迎页 */
    .welcome {
        flex: 1; display: flex; flex-direction: column; align-items: center;
        justify-content: center; color: var(--text-muted); gap: 12px;
    }
    .welcome h2 { color: var(--text); font-size: 1.4em; }
    .welcome kbd {
        background: var(--input-bg); padding: 4px 10px; border-radius: 4px;
        font-family: monospace; font-size: .9em;
    }
    
    /* 搜索结果 */
    .search-results { padding: 24px 32px; overflow-y: auto; }
    .search-result {
        padding: 12px 16px; border-radius: 8px; margin-bottom: 8px;
        border: 1px solid var(--border); cursor: pointer; transition: background .15s;
    }
    .search-result:hover { background: var(--hover); }
    .search-result h3 { font-size: 15px; margin-bottom: 4px; }
    .search-result h3 a { color: var(--accent); text-decoration: none; }
    .search-result .snippet { font-size: 13px; color: var(--text-muted); }
    
    /* 空状态 */
    .empty { text-align: center; padding: 40px; color: var(--text-muted); }
</style>
</head>
<body>
<div class="sidebar">
    <div class="sidebar-header">
        <input type="text" id="search-input" placeholder="搜索笔记..." value="{{ query }}"
               onkeydown="if(event.key==='Enter')window.location='/search?q='+encodeURIComponent(this.value)">
        <button onclick="createNote()">+</button>
    </div>
    <div class="note-list">
        {% for n in notes %}
        <a href="/note/{{ n.slug }}" class="note-item {% if note and note.slug == n.slug %}active{% endif %}">
            {{ n.title }}
            <span class="time">{{ n.mtime }}</span>
        </a>
        {% endfor %}
        {% if not notes %}
        <div class="empty">还没有笔记<br>点击 + 创建第一篇</div>
        {% endif %}
    </div>
</div>

<div class="main">
    {% if search_results is not none %}
    <!-- 搜索结果 -->
    <div class="toolbar">
        <span style="color:var(--text-muted)">搜索 "{{ query }}" — {{ search_results|length }} 个结果</span>
        <a href="/">✕ 清除</a>
    </div>
    <div class="search-results">
        {% for r in search_results %}
        <div class="search-result" onclick="location.href='/note/{{ r.slug }}'">
            <h3><a href="/note/{{ r.slug }}">{{ r.title }}</a></h3>
            <div class="snippet">{{ r.snippet }}</div>
        </div>
        {% endfor %}
        {% if not search_results %}
        <div class="empty">没有找到匹配的笔记</div>
        {% endif %}
    </div>
    {% elif note %}
    <!-- 笔记内容 -->
    <div class="toolbar">
        <a href="/edit/{{ note.slug }}">✏️ 编辑</a>
        <button onclick="if(confirm('确定删除「{{ note.title }}」？')){fetch('/delete/{{ note.slug }}',{method:'POST'}).then(()=>location.href='/')}">🗑 删除</button>
        <a href="/api/graph" target="_blank" style="margin-left:auto">🔗 图谱</a>
    </div>
    <div class="content-area">
        <div class="note-content">{{ content_html | safe }}</div>
        {% if backlinks %}
        <div class="backlinks-panel">
            <h3>🔗 反向链接 ({{ backlinks|length }})</h3>
            {% for bl in backlinks %}
            <a href="/note/{{ bl.slug }}">{{ bl.title }}</a>
            {% endfor %}
        </div>
        {% endif %}
    </div>
    {% else %}
    <!-- 首页 -->
    <div class="welcome">
        <h2>📝 知识库</h2>
        <p>类似 Obsidian 的本地 Markdown 笔记</p>
        <p>使用 <kbd>[[笔记名]]</kbd> 创建双向链接</p>
        <p style="margin-top:8px">{{ notes|length }} 篇笔记</p>
    </div>
    {% endif %}
</div>

<script>
async function createNote() {
    const title = prompt('📝 笔记标题：');
    if (!title) return;
    const resp = await fetch('/create', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({title: title})
    });
    const data = await resp.json();
    if (data.ok) {
        location.href = '/edit/' + data.slug;
    } else {
        alert('创建失败: ' + data.error);
    }
}
</script>
</body>
</html>"""

EDIT_TEMPLATE = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>编辑 — {{ note.title }}</title>
<style>
    :root {
        --bg: #1e1e2e; --sidebar-bg: #181825; --text: #cdd6f4;
        --text-muted: #a6adc8; --accent: #89b4fa; --border: #313244;
        --hover: #45475a; --input-bg: #313244;
    }
    * { margin:0; padding:0; box-sizing:border-box; }
    body { font-family: -apple-system, "Microsoft YaHei", sans-serif; background: var(--bg); color: var(--text); display: flex; height: 100vh; }
    
    .sidebar {
        width: 240px; min-width: 240px; background: var(--sidebar-bg);
        border-right: 1px solid var(--border); padding: 16px; overflow-y: auto;
    }
    .sidebar a { display: block; padding: 6px 10px; color: var(--text); text-decoration: none; border-radius: 4px; font-size: 13px; }
    .sidebar a:hover { background: var(--hover); }
    .sidebar a.active { background: var(--accent); color: #1e1e2e; font-weight: 600; }
    
    .main { flex:1; display:flex; flex-direction:column; }
    .toolbar {
        padding: 12px 20px; border-bottom: 1px solid var(--border);
        display: flex; gap: 10px; align-items: center;
    }
    .toolbar button {
        padding: 8px 18px; border-radius: 6px; border: none;
        cursor: pointer; font-size: 14px; font-weight: 600;
    }
    .toolbar .save { background: var(--accent); color: #1e1e2e; }
    .toolbar .cancel { background: transparent; color: var(--text); border: 1px solid var(--border); }
    
    .editor-wrapper { flex:1; display:flex; overflow:hidden; }
    .editor-pane { flex:1; display:flex; flex-direction:column; }
    .editor-pane textarea {
        flex:1; background: var(--bg); color: var(--text); border: none;
        padding: 24px; font-family: "JetBrains Mono", "Cascadia Code", monospace;
        font-size: 14px; line-height: 1.7; resize: none; outline: none;
    }
    .preview-pane {
        flex:1; padding: 24px; overflow-y: auto; border-left: 1px solid var(--border);
        line-height: 1.8;
    }
    .preview-pane h1 { font-size:1.6em; border-bottom:1px solid var(--border); padding-bottom:8px; }
    .preview-pane h2 { font-size:1.3em; margin:16px 0 8px; }
    .preview-pane code { background:var(--input-bg); padding:2px 6px; border-radius:4px; }
    .preview-pane pre { background:var(--input-bg); padding:12px; border-radius:6px; overflow-x:auto; }
    .preview-pane blockquote { border-left:3px solid var(--accent); padding:4px 16px; color:var(--text-muted); }
    .preview-pane a { color: var(--accent); }
    
    .help-bar {
        padding: 8px 20px; border-top: 1px solid var(--border);
        color: var(--text-muted); font-size: 12px; display: flex; gap: 16px;
    }
    .help-bar span { color: var(--accent); }
</style>
</head>
<body>
<div class="sidebar">
    <div style="margin-bottom:12px;color:var(--text-muted);font-size:12px;">📝 笔记列表</div>
    {% for n in notes %}
    <a href="/note/{{ n.slug }}" class="{% if n.slug == note.slug %}active{% endif %}">{{ n.title }}</a>
    {% endfor %}
</div>

<div class="main">
    <div class="toolbar">
        <button class="save" onclick="save()">💾 保存</button>
        <button class="cancel" onclick="location.href='/note/{{ note.slug }}'">取消</button>
        <span style="margin-left:12px;color:var(--text-muted);font-size:13px;">{{ note.slug }}.md</span>
    </div>
    <div class="editor-wrapper">
        <div class="editor-pane">
            <textarea id="editor" oninput="updatePreview()">{{ note.content }}</textarea>
        </div>
        <div class="preview-pane" id="preview"></div>
    </div>
    <div class="help-bar">
        <span>[[笔记名]]</span> 创建链接 &nbsp;
        <span>[[笔记名|别名]]</span> 别名链接 &nbsp;
        <span>Ctrl+S</span> 保存
    </div>
</div>

<script>
const slug = "{{ note.slug }}";
const editor = document.getElementById("editor");

// 实时预览
let previewTimer;
function updatePreview() {
    clearTimeout(previewTimer);
    previewTimer = setTimeout(async () => {
        const resp = await fetch('/api/preview', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({content: editor.value})
        });
        const data = await resp.json();
        document.getElementById("preview").innerHTML = data.html;
    }, 300);
}

async function save() {
    const resp = await fetch('/edit/' + slug, {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({content: editor.value})
    });
    const data = await resp.json();
    if (data.ok) {
        location.href = '/note/' + data.slug;
    } else {
        alert('保存失败');
    }
}

document.addEventListener('keydown', e => {
    if ((e.ctrlKey || e.metaKey) && e.key === 's') {
        e.preventDefault(); save();
    }
});

// 初始预览
updatePreview();
</script>
</body>
</html>"""

# ── 预览 API ──
@app.route("/api/preview", methods=["POST"])
def api_preview():
    data = request.get_json()
    content = data.get("content", "")
    html = render_markdown(content)
    return jsonify({"html": html})



# ── 知识图谱页面 ──
GRAPH_TEMPLATE = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>知识图谱</title>
<script src="https://d3js.org/d3.v7.min.js"></script>
<style>
    :root { --bg: #1e1e2e; --text: #cdd6f4; --accent: #89b4fa; }
    body { margin:0; background:var(--bg); color:var(--text); font-family:-apple-system,"Microsoft YaHei",sans-serif; overflow:hidden; }
    .header { position:absolute; top:0; left:0; right:0; padding:12px 20px; z-index:10; display:flex; gap:16px; align-items:center; }
    .header a { color:var(--accent); text-decoration:none; font-size:14px; }
    svg { width:100vw; height:100vh; }
    .node circle { fill: var(--accent); stroke: #1e1e2e; stroke-width: 2px; cursor: pointer; }
    .node text { fill: var(--text); font-size: 13px; cursor: pointer; }
    .link { stroke: #45475a; stroke-width: 1.5px; stroke-opacity: 0.6; }
</style>
</head>
<body>
<div class="header">
    <a href="/">← 返回知识库</a>
    <span style="color:#a6adc8;font-size:13px;">拖拽节点 | 滚轮缩放 | 点击跳转</span>
</div>
<svg id="graph"></svg>
<script>
const svg = d3.select("#graph"),
      width = window.innerWidth,
      height = window.innerHeight;
const g = svg.append("g");
svg.call(d3.zoom().scaleExtent([0.2, 4]).on("zoom", (e) => g.attr("transform", e.transform)));
fetch("/api/graph")
  .then(r => r.json())
  .then(data => {
    const nodes = data.nodes.map(n => ({...n}));
    const links = data.links.filter(l => nodes.some(n => n.id === l.target));
    const sim = d3.forceSimulation(nodes)
      .force("link", d3.forceLink(links).id(d => d.id).distance(120))
      .force("charge", d3.forceManyBody().strength(-300))
      .force("center", d3.forceCenter(width/2, height/2))
      .force("collide", d3.forceCollide(30));
    const link = g.append("g").selectAll("line").data(links).join("line").attr("class", "link");
    const node = g.append("g").selectAll(".node").data(nodes).join("g").attr("class", "node")
      .call(d3.drag()
        .on("start", (e,d) => { if(!e.active) sim.alphaTarget(0.3).restart(); d.fx = d.x; d.fy = d.y; })
        .on("drag", (e,d) => { d.fx = e.x; d.fy = e.y; })
        .on("end", (e,d) => { if(!e.active) sim.alphaTarget(0); d.fx = null; d.fy = null; }));
    node.append("circle").attr("r", d => Math.max(8, Math.min(20, Math.sqrt(d.size) / 10)))
      .on("click", (e,d) => window.open("/note/"+d.id, "_self"));
    node.append("text").text(d => d.label.length > 8 ? d.label.slice(0,8)+"…" : d.label)
      .attr("dy", -14).attr("text-anchor", "middle")
      .on("click", (e,d) => window.open("/note/"+d.id, "_self"));
    sim.on("tick", () => {
      link.attr("x1", d => d.source.x).attr("y1", d => d.source.y)
          .attr("x2", d => d.target.x).attr("y2", d => d.target.y);
      node.attr("transform", d => `translate(${d.x},${d.y})`);
    });
  });
</script>
</body>
</html>"""

@app.route("/graph")
def graph_page():
    return render_template_string(GRAPH_TEMPLATE)
if __name__ == "__main__":
    print("📝 知识库启动: http://127.0.0.1:8940")
    app.run(host="127.0.0.1", port=8940, debug=False)

