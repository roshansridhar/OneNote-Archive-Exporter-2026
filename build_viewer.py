#!/usr/bin/env python3
import argparse
import json
from pathlib import Path


def rel_posix(path: Path, root: Path) -> str:
    return str(path.relative_to(root)).replace("\\", "/")


def build_section_page_tree(section_data: dict) -> list:
    pages = section_data.get("pages", [])
    # OneNote page tree should follow Graph's explicit "order", not filename/title sort.
    pages = sorted(
        pages,
        key=lambda p: (
            p.get("order") is None,
            p.get("order") if p.get("order") is not None else 10**9,
            p.get("title", "").lower(),
        ),
    )
    roots = []
    stack = []

    for p in pages:
        html_path = p.get("html")
        if not html_path:
            continue

        try:
            level = int(p.get("level", 0) or 0)
        except (TypeError, ValueError):
            level = 0
        node = {
            "name": p.get("title", "untitled"),
            "path": html_path,
            "type": "page",
            "children": [],
        }

        if level <= 0:
            roots.append(node)
            stack = [node]
            continue

        parent = None
        if len(stack) >= level:
            parent = stack[level - 1]
        elif stack:
            parent = stack[-1]

        if parent is None:
            roots.append(node)
            stack = [node]
            continue

        parent["children"].append(node)
        stack = stack[:level]
        stack.append(node)

    def normalize(n: dict) -> dict:
        if n["children"]:
            n["type"] = "page-group"
            n["children"] = [normalize(c) for c in n["children"]]
        else:
            n["type"] = "file"
            n.pop("children", None)
        return n

    return [normalize(n) for n in roots]


def build_tree(root: Path, current: Path):
    rel = current.relative_to(root)
    node = {
        "name": current.name if current != root else root.name,
        "path": str(rel).replace("\\", "/"),
        "type": "dir",
        "children": [],
    }

    section_json = current / "_section.json"
    if section_json.exists():
        try:
            section_data = json.loads(section_json.read_text(encoding="utf-8"))
            node["children"] = build_section_page_tree(section_data)
            return node
        except Exception:
            pass

    children = sorted(current.iterdir(), key=lambda p: (p.is_file(), p.name.lower()))
    for child in children:
        if child.name.startswith("."):
            continue
        if child.is_dir():
            if child.name.endswith("_assets"):
                continue
            child_node = build_tree(root, child)
            if child_node["children"]:
                node["children"].append(child_node)
            continue

        if child.suffix.lower() == ".html" and child.name.lower() != "viewer.html":
            node["children"].append(
                {
                    "name": child.stem,
                    "path": rel_posix(child, root),
                    "type": "file",
                }
            )
    return node


def render_html(tree_json: str, title: str) -> str:
    return f"""<!doctype html>
<html lang=\"en\">
<head>
  <meta charset=\"utf-8\" />
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />
  <title>{title}</title>
  <style>
    :root {{
      --bg: #f6f7f9;
      --panel: #ffffff;
      --border: #e3e6eb;
      --text: #1f2937;
      --muted: #667085;
      --accent: #0f766e;
      --accent-soft: #e6fffa;
    }}
    * {{ box-sizing: border-box; }}
    html, body {{ height: 100%; margin: 0; font-family: ui-sans-serif, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif; color: var(--text); background: var(--bg); }}
    .app {{ display: grid; grid-template-columns: 340px 1fr; height: 100%; gap: 12px; padding: 12px; }}
    .panel {{ background: var(--panel); border: 1px solid var(--border); border-radius: 12px; overflow: hidden; }}
    .sidebar {{ display: flex; flex-direction: column; }}
    .header {{ padding: 12px 14px; border-bottom: 1px solid var(--border); }}
    .title {{ font-size: 15px; font-weight: 700; }}
    .sub {{ font-size: 12px; color: var(--muted); margin-top: 4px; }}
    .tree-wrap {{ padding: 8px; overflow: auto; }}
    .node {{ margin-left: 14px; }}
    details {{ margin: 2px 0; }}
    summary {{ cursor: pointer; list-style: none; font-size: 13px; padding: 4px 6px; border-radius: 8px; display:flex; align-items:center; gap:6px; }}
    summary:hover {{ background: #f3f4f6; }}
    summary::-webkit-details-marker {{ display: none; }}
    .folder::before {{ content: '[+]'; color: #64748b; font-size: 11px; }}
    .page-parent::before {{ content: '[-]'; color: #64748b; font-size: 11px; }}
    .page-open {{ border: 0; background: transparent; color: #0f766e; cursor: pointer; padding: 0; font-size: 12px; }}
    .file-btn {{ display: block; width: 100%; text-align: left; border: 0; background: transparent; padding: 5px 6px; border-radius: 8px; cursor: pointer; font-size: 13px; color: #111827; }}
    .file-btn:hover {{ background: #f3f4f6; }}
    .file-btn.active {{ background: var(--accent-soft); color: #0f172a; font-weight: 600; }}
    .viewer {{ display: grid; grid-template-rows: auto 1fr; }}
    .viewer-top {{ display: flex; align-items: center; justify-content: space-between; gap: 8px; padding: 10px 12px; border-bottom: 1px solid var(--border); }}
    .path {{ font-size: 12px; color: var(--muted); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
    .open-link {{ font-size: 12px; color: var(--accent); text-decoration: none; }}
    iframe {{ width: 100%; height: 100%; border: 0; background: #fff; }}
    .empty {{ display: grid; place-items: center; color: var(--muted); font-size: 14px; height: 100%; }}
    @media (max-width: 900px) {{ .app {{ grid-template-columns: 1fr; grid-template-rows: 40% 1fr; }} }}
  </style>
</head>
<body>
  <div class=\"app\">
    <aside class=\"panel sidebar\">
      <div class=\"header\">
        <div class=\"title\">OneNote Export Viewer</div>
        <div class=\"sub\">Uses page hierarchy when available</div>
      </div>
      <div class=\"tree-wrap\" id=\"tree\"></div>
    </aside>

    <main class=\"panel viewer\">
      <div class=\"viewer-top\">
        <div class=\"path\" id=\"path\">Select a note from the tree.</div>
        <a id=\"openLink\" class=\"open-link\" target=\"_blank\" rel=\"noopener noreferrer\" style=\"display:none\">Open in new tab</a>
      </div>
      <div id=\"viewArea\" class=\"empty\">No file selected</div>
    </main>
  </div>

<script>
const treeData = {tree_json};
const treeEl = document.getElementById('tree');
const pathEl = document.getElementById('path');
const viewArea = document.getElementById('viewArea');
const openLink = document.getElementById('openLink');
let activeBtn = null;

function openFile(path, btn) {{
  if (activeBtn) activeBtn.classList.remove('active');
  if (btn) btn.classList.add('active');
  activeBtn = btn || null;

  pathEl.textContent = path;
  openLink.href = path;
  openLink.style.display = 'inline';

  const iframe = document.createElement('iframe');
  iframe.src = path;
  viewArea.innerHTML = '';
  viewArea.className = '';
  viewArea.appendChild(iframe);
}}

function createNode(node) {{
  if (node.type === 'file') {{
    const btn = document.createElement('button');
    btn.className = 'file-btn';
    btn.textContent = node.name;
    btn.addEventListener('click', () => openFile(node.path, btn));
    return btn;
  }}

  const details = document.createElement('details');
  if (!node.path || node.path.split('/').length <= 1) details.open = true;

  const summary = document.createElement('summary');
  summary.className = node.type === 'page-group' ? 'page-parent' : 'folder';

  const label = document.createElement('span');
  label.textContent = node.name;
  summary.appendChild(label);

  if (node.type === 'page-group' && node.path) {{
    const openBtn = document.createElement('button');
    openBtn.className = 'page-open';
    openBtn.textContent = 'open';
    openBtn.addEventListener('click', (e) => {{
      e.preventDefault();
      e.stopPropagation();
      openFile(node.path, null);
    }});
    summary.appendChild(openBtn);
  }}

  details.appendChild(summary);

  const container = document.createElement('div');
  container.className = 'node';
  (node.children || []).forEach((child) => container.appendChild(createNode(child)));
  details.appendChild(container);
  return details;
}}

function init() {{
  treeEl.appendChild(createNode(treeData));
}}

init();
</script>
</body>
</html>
"""


def main():
    p = argparse.ArgumentParser(description="Build local HTML tree viewer for OneNote export")
    p.add_argument("--export-dir", default="export", help="Path to export directory")
    p.add_argument("--out", default="viewer.html", help="Output HTML filename (inside export dir if relative)")
    args = p.parse_args()

    export_dir = Path(args.export_dir).resolve()
    if not export_dir.exists():
        raise SystemExit(f"Export directory does not exist: {export_dir}")

    tree = build_tree(export_dir, export_dir)
    out = Path(args.out)
    if not out.is_absolute():
        out = export_dir / out

    html = render_html(json.dumps(tree, ensure_ascii=False), "OneNote Export Viewer")
    out.write_text(html, encoding="utf-8")
    print(f"Viewer written: {out}")


if __name__ == "__main__":
    main()
