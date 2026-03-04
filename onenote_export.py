#!/usr/bin/env python3
import argparse
import hashlib
import json
import os
import pathlib
import random
import re
import sys
import time
from html import escape
from typing import Dict, List, Optional
from urllib.parse import urljoin, urlparse
from xml.etree import ElementTree as ET

import msal
import requests
from PIL import Image, ImageDraw

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SCOPES = ["Notes.Read", "User.Read"]
REQUEST_DELAY_S = 0.0
MAX_RETRIES = 6
MAX_BACKOFF_S = 30
INK_PREVIEW_BEGIN = "<!-- ONENOTE_INK_PREVIEW_BEGIN -->"
INK_PREVIEW_END = "<!-- ONENOTE_INK_PREVIEW_END -->"


def log(msg: str) -> None:
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)


def slugify(name: str, fallback: str) -> str:
    if not name:
        return fallback
    value = re.sub(r"[^a-zA-Z0-9._-]+", "_", name).strip("_")
    return value[:120] or fallback


def graph_request(
    token: str,
    url: str,
    *,
    params: Dict = None,
    timeout: int = 45,
    allow_redirects: bool = False,
) -> requests.Response:
    headers = {"Authorization": f"Bearer {token}"}
    for attempt in range(MAX_RETRIES):
        if REQUEST_DELAY_S > 0:
            time.sleep(REQUEST_DELAY_S)
        r = requests.get(
            url,
            headers=headers,
            params=params,
            timeout=timeout,
            allow_redirects=allow_redirects,
        )
        if r.status_code in (429, 503):
            wait_s = compute_retry_wait_seconds(r.headers, attempt)
            print(f"Retrying {url} after {wait_s}s due to {r.status_code}...")
            time.sleep(wait_s)
            continue
        if r.status_code >= 400:
            raise RuntimeError(f"Graph GET failed {r.status_code} for {url}: {r.text[:500]}")
        return r
    raise RuntimeError(f"Graph GET repeatedly throttled/unavailable for {url}")


def graph_get(token: str, url: str, params: Dict = None) -> Dict:
    return graph_request(token, url, params=params, timeout=45, allow_redirects=False).json()


def graph_get_bytes(token: str, url: str) -> bytes:
    return graph_request(token, url, timeout=45, allow_redirects=False).content


def guess_ext_from_content_type(content_type: str) -> str:
    ct = (content_type or "").split(";")[0].strip().lower()
    mapping = {
        "image/jpeg": ".jpg",
        "image/png": ".png",
        "image/gif": ".gif",
        "image/webp": ".webp",
        "image/svg+xml": ".svg",
        "application/pdf": ".pdf",
        "text/plain": ".txt",
        "application/octet-stream": ".bin",
    }
    return mapping.get(ct, ".bin")


def graph_download(token: str, url: str) -> tuple:
    r = graph_request(token, url, timeout=60, allow_redirects=True)
    return r.content, r.headers.get("Content-Type", ""), r.url


def compute_retry_wait_seconds(headers: Dict, attempt: int) -> int:
    retry_after = headers.get("Retry-After")
    if retry_after and retry_after.isdigit():
        return max(1, min(int(retry_after), MAX_BACKOFF_S))

    retry_after_ms = headers.get("x-ms-retry-after-ms")
    if retry_after_ms and retry_after_ms.isdigit():
        return max(1, min(int(int(retry_after_ms) / 1000), MAX_BACKOFF_S))

    base = min(2 ** attempt, MAX_BACKOFF_S)
    jittered = int(max(1, min(base * random.uniform(0.85, 1.25), MAX_BACKOFF_S)))
    return jittered


def write_bytes_atomic(path: pathlib.Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_bytes(data)
    tmp.replace(path)


def write_text_atomic(path: pathlib.Path, text: str) -> None:
    write_bytes_atomic(path, text.encode("utf-8"))


def write_json_atomic(path: pathlib.Path, obj: Dict) -> None:
    write_text_atomic(path, json.dumps(obj, indent=2, ensure_ascii=False))


def work_id_asset(url: str) -> str:
    return "asset:" + hashlib.sha256(url.encode("utf-8")).hexdigest()


class StateStore:
    def __init__(self, path: pathlib.Path):
        self.path = path
        self.records: Dict[str, Dict] = {}
        self.path.parent.mkdir(parents=True, exist_ok=True)
        if self.path.exists():
            for line in self.path.read_text(encoding="utf-8", errors="ignore").splitlines():
                line = line.strip()
                if not line:
                    continue
                try:
                    rec = json.loads(line)
                except Exception:
                    continue
                wid = rec.get("work_id")
                if wid:
                    self.records[wid] = rec

    def get(self, work_id: str) -> Optional[Dict]:
        return self.records.get(work_id)

    def is_ok(self, work_id: str) -> bool:
        rec = self.records.get(work_id)
        return bool(rec and rec.get("status") == "ok")

    def mark(self, work_id: str, status: str, detail: str = "", payload: Dict = None) -> None:
        rec = {
            "ts_epoch": int(time.time()),
            "work_id": work_id,
            "status": status,
            "detail": detail[:1000] if detail else "",
        }
        if payload:
            rec["payload"] = payload
        with self.path.open("a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")
        self.records[work_id] = rec


def download_page_assets(
    token: str,
    html: str,
    assets_dir: pathlib.Path,
    state: StateStore,
    global_seen: Dict[str, str] = None,
    global_failed: set = None,
) -> str:
    attr_pat = re.compile(
        r"""(?:src|href|data-fullres-src)\s*=\s*(["'])(.*?)\1""",
        flags=re.IGNORECASE | re.DOTALL,
    )
    refs = [m.group(2) for m in attr_pat.finditer(html)]
    if not refs:
        return html

    assets_dir.mkdir(parents=True, exist_ok=True)
    seen = {}
    idx = 1
    rewritten = html
    global_seen = global_seen if global_seen is not None else {}
    global_failed = global_failed if global_failed is not None else set()
    page_dir = assets_dir.parent

    for ref in refs:
        if not ref or ref.startswith("data:") or ref.startswith("#"):
            continue
        full = urljoin("https://graph.microsoft.com/", ref)
        parsed = urlparse(full)
        if parsed.scheme not in ("http", "https"):
            continue
        if "graph.microsoft.com" not in (parsed.netloc or ""):
            continue
        if "/onenote/resources/" not in (parsed.path or ""):
            continue
        asset_id = work_id_asset(full)
        rec = state.get(asset_id)
        if rec and rec.get("status") == "ok":
            payload = rec.get("payload") or {}
            abs_path = payload.get("abs_path")
            if abs_path and pathlib.Path(abs_path).exists():
                local_rel = os.path.relpath(abs_path, start=str(page_dir)).replace("\\", "/")
                rewritten = rewritten.replace(ref, local_rel)
                continue
        if full in global_failed:
            continue
        if full in global_seen:
            cached_abs = pathlib.Path(global_seen[full])
            if cached_abs.exists():
                local_rel = os.path.relpath(str(cached_abs), start=str(page_dir)).replace("\\", "/")
                rewritten = rewritten.replace(ref, local_rel)
            continue
        if full in seen:
            local_rel = seen[full]
            rewritten = rewritten.replace(ref, local_rel)
            continue

        try:
            content, content_type, final_url = graph_download(token, full)
        except Exception as ex:
            print(f"Skipping asset fetch for {full}: {ex}")
            # Keep transient server-side errors retriable within the same run.
            if is_permanent_asset_failure(str(ex)):
                global_failed.add(full)
                state.mark(asset_id, "permanent_fail", str(ex))
            else:
                state.mark(asset_id, "retryable_fail", str(ex))
            continue
        ext = guess_ext_from_content_type(content_type)
        file_name = f"asset_{idx:04d}{ext}"
        idx += 1
        out_path = assets_dir / file_name
        write_bytes_atomic(out_path, content)
        local_rel = f"{assets_dir.name}/{file_name}"
        seen[full] = local_rel
        seen[final_url] = local_rel
        global_seen[full] = str(out_path.resolve())
        global_seen[final_url] = str(out_path.resolve())
        rewritten = rewritten.replace(ref, local_rel)
        state.mark(asset_id, "ok", payload={"abs_path": str(out_path.resolve()), "url": full})

    return rewritten


def add_metadata_header(html_text: str, title: str, created: str, edited: str) -> str:
    header = (
        '<div style="border:1px solid #ddd;padding:10px 12px;margin:8px 0 16px 0;'
        'font-family:Arial,sans-serif;background:#f8f8f8">'
        f"<div><strong>Title:</strong> {escape(title or '')}</div>"
        f"<div><strong>Created:</strong> {escape(created or '')}</div>"
        f"<div><strong>Last Edited:</strong> {escape(edited or '')}</div>"
        "</div>"
    )
    body_open = re.search(r"<body[^>]*>", html_text, flags=re.IGNORECASE)
    if body_open:
        insert_at = body_open.end()
        return html_text[:insert_at] + header + html_text[insert_at:]
    return header + html_text


def extract_inkml_from_multipart(raw: bytes) -> bytes:
    low = raw.lower()
    marker = b"content-type: application/inkml+xml"
    idx = low.find(marker)
    if idx == -1:
        return b""

    # Find payload start after part headers.
    header_end = raw.find(b"\r\n\r\n", idx)
    if header_end == -1:
        header_end = raw.find(b"\n\n", idx)
    if header_end == -1:
        return b""
    payload_start = header_end + (4 if raw[header_end:header_end + 4] == b"\r\n\r\n" else 2)

    # Next multipart boundary starts at line beginning with "--".
    boundary_idx = raw.find(b"\r\n--", payload_start)
    if boundary_idx == -1:
        boundary_idx = raw.find(b"\n--", payload_start)
    payload_end = boundary_idx if boundary_idx != -1 else len(raw)
    return raw[payload_start:payload_end].strip()


def render_inkml_to_png(inkml_bytes: bytes, png_path: pathlib.Path) -> bool:
    if not inkml_bytes:
        return False
    try:
        root = ET.fromstring(inkml_bytes)
    except Exception:
        return False

    ns = {"inkml": "http://www.w3.org/2003/InkML"}
    traces = root.findall(".//inkml:trace", ns)
    strokes = []
    min_x = float("inf")
    min_y = float("inf")
    max_x = float("-inf")
    max_y = float("-inf")

    for tr in traces:
        raw = (tr.text or "").strip()
        if not raw:
            continue
        pts = []
        for chunk in raw.split(","):
            vals = [v for v in re.split(r"\s+", chunk.strip()) if v]
            if len(vals) < 2:
                continue
            try:
                x = float(vals[0])
                y = float(vals[1])
            except Exception:
                continue
            pts.append((x, y))
            min_x = min(min_x, x)
            min_y = min(min_y, y)
            max_x = max(max_x, x)
            max_y = max(max_y, y)
        if len(pts) > 1:
            strokes.append(pts)

    if not strokes:
        return False

    pad = 48
    target_w = 3200
    target_h = 2400
    width = max(1.0, max_x - min_x)
    height = max(1.0, max_y - min_y)
    scale = min((target_w - 2 * pad) / width, (target_h - 2 * pad) / height)
    draw_w = int(width * scale + 2 * pad)
    draw_h = int(height * scale + 2 * pad)
    draw_w = max(1200, min(draw_w, 6000))
    draw_h = max(900, min(draw_h, 6000))

    # Supersample then downsample for smoother strokes.
    ss = 2
    img = Image.new("RGB", (draw_w * ss, draw_h * ss), "white")
    draw = ImageDraw.Draw(img)
    stroke_w = max(2, int(2.0 * ss))
    for stroke in strokes:
        points = []
        for x, y in stroke:
            px = (pad + (x - min_x) * scale) * ss
            py = (pad + (y - min_y) * scale) * ss
            points.append((px, py))
        draw.line(points, fill="black", width=stroke_w, joint="curve")

    img = img.resize((draw_w, draw_h), Image.Resampling.LANCZOS)

    png_path.parent.mkdir(parents=True, exist_ok=True)
    tmp = png_path.with_suffix(png_path.suffix + ".tmp")
    img.save(tmp, format="PNG")
    tmp.replace(png_path)
    return True


def inject_ink_preview_image(html_text: str, img_rel_path: str) -> str:
    block = (
        f'{INK_PREVIEW_BEGIN}<div style="margin:10px 0 16px 0">'
        f'<div style="font-family:Arial,sans-serif;font-size:12px;color:#444;margin-bottom:6px">'
        "Handwriting preview</div>"
        f'<img src="{escape(img_rel_path)}" style="max-width:100%;height:auto;border:1px solid #ddd" />'
        f"</div>{INK_PREVIEW_END}"
    )
    # Replace new-style block safely by explicit begin/end markers.
    if INK_PREVIEW_BEGIN in html_text and INK_PREVIEW_END in html_text:
        start = html_text.find(INK_PREVIEW_BEGIN)
        end = html_text.find(INK_PREVIEW_END, start)
        if start != -1 and end != -1:
            end += len(INK_PREVIEW_END)
            return html_text[:start] + block + html_text[end:]

    # Replace old one-off marker block format from previous versions.
    old_block_pattern = (
        r"<!-- ONENOTE_INK_PREVIEW -->\s*"
        r"<div[^>]*>\s*"
        r"<div[^>]*>Handwriting preview</div>\s*"
        r"<img[^>]*>\s*"
        r"</div>"
    )
    if re.search(old_block_pattern, html_text, flags=re.DOTALL):
        html_text = re.sub(old_block_pattern, block, html_text, flags=re.DOTALL)
        return html_text
    body_open = re.search(r"<body[^>]*>", html_text, flags=re.IGNORECASE)
    if body_open:
        insert_at = body_open.end()
        return html_text[:insert_at] + block + html_text[insert_at:]
    return block + html_text


def is_permanent_asset_failure(error_text: str) -> bool:
    match = re.search(r"Graph GET failed\s+(\d{3})", error_text)
    if not match:
        return False
    code = int(match.group(1))
    if code in (429, 503):
        return False
    if 500 <= code <= 599:
        return False
    return 400 <= code <= 499


def paged_values(token: str, start_url: str, params: Dict = None) -> List[Dict]:
    out = []
    next_url = start_url
    first = True
    while next_url:
        payload = graph_get(token, next_url, params if first else None)
        first = False
        out.extend(payload.get("value", []))
        next_url = payload.get("@odata.nextLink")
    return out


def acquire_token(client_id: str, tenant: str, cache_path: pathlib.Path) -> str:
    authority = f"https://login.microsoftonline.com/{tenant}"
    cache = msal.SerializableTokenCache()
    if cache_path.exists():
        cache.deserialize(cache_path.read_text(encoding="utf-8"))

    app = msal.PublicClientApplication(client_id=client_id, authority=authority, token_cache=cache)

    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(f"Failed to create device flow: {json.dumps(flow, indent=2)}")

    print("Open this URL and enter the code:")
    print(flow["verification_uri"])
    print(f"Code: {flow['user_code']}")

    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(f"Auth failed: {json.dumps(result, indent=2)}")

    if cache.has_state_changed:
        cache_path.write_text(cache.serialize(), encoding="utf-8")

    return result["access_token"]


def export_onenote(token: str, out_dir: pathlib.Path, state: StateStore) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    log(f"Starting export into: {out_dir}")
    notebooks = paged_values(token, f"{GRAPH_BASE}/me/onenote/notebooks")
    log(f"Found notebooks: {len(notebooks)}")
    manifest = {
        "exported_at_epoch": int(time.time()),
        "notebooks": [],
    }

    skipped_notebooks = 0
    skipped_sections = 0
    skipped_pages = 0

    for n_idx, nb in enumerate(notebooks, start=1):
        nb_name = nb.get("displayName", f"notebook_{n_idx}")
        nb_id = nb["id"]
        nb_dir = out_dir / slugify(nb_name, f"notebook_{n_idx}")
        nb_dir.mkdir(exist_ok=True)
        log(f"[Notebook {n_idx}/{len(notebooks)}] {nb_name}")

        try:
            sections = paged_values(
                token,
                f"{GRAPH_BASE}/me/onenote/notebooks/{nb_id}/sections",
                params={"$select": "id,displayName,createdDateTime,lastModifiedDateTime"},
            )
        except Exception as ex:
            skipped_notebooks += 1
            log(f"Skipping notebook {nb_id} due to throttling/error: {ex}")
            continue
        log(f"  Sections: {len(sections)}")

        nb_entry = {"id": nb_id, "displayName": nb_name, "sections": []}

        for s_idx, section in enumerate(sections, start=1):
            s_name = section.get("displayName", f"section_{s_idx}")
            s_id = section["id"]
            s_dir = nb_dir / slugify(s_name, f"section_{s_idx}")
            s_dir.mkdir(exist_ok=True)
            log(f"  [Section {s_idx}/{len(sections)}] {s_name}")

            try:
                pages = paged_values(
                    token,
                    f"{GRAPH_BASE}/me/onenote/sections/{s_id}/pages",
                    params={
                        "$select": "id,title,level,order,createdDateTime,lastModifiedDateTime,links",
                        "pagelevel": "true",
                    },
                )
            except Exception as ex:
                skipped_sections += 1
                log(f"Skipping section {s_id} due to throttling/error: {ex}")
                continue
            log(f"    Pages: {len(pages)}")

            s_entry = {"id": s_id, "displayName": s_name, "pages": []}

            for p_idx, page in enumerate(pages, start=1):
                title = page.get("title", f"page_{p_idx}")
                p_id = page["id"]
                p_slug = slugify(title, f"page_{p_idx}")

                page_meta_path = s_dir / f"{p_slug}.meta.json"
                write_json_atomic(page_meta_path, page)

                html_path = s_dir / f"{p_slug}.html"
                page_work_id = f"page_html:{p_id}"
                page_done = state.is_ok(page_work_id) and html_path.exists()
                rewritten_html = ""
                if page_done:
                    log(f"    [Page {p_idx}/{len(pages)}] resume html: {title}")
                    rewritten_html = html_path.read_text(encoding="utf-8", errors="replace")
                else:
                    log(f"    [Page {p_idx}/{len(pages)}] fetch html: {title}")
                    try:
                        html_bytes = graph_get_bytes(token, f"{GRAPH_BASE}/me/onenote/pages/{p_id}/content")
                    except Exception as ex:
                        skipped_pages += 1
                        state.mark(page_work_id, "retryable_fail", str(ex))
                        log(f"Skipping page content for {p_id}: {ex}")
                        continue
                    html_text = html_bytes.decode("utf-8", errors="replace")
                    html_text = add_metadata_header(
                        html_text=html_text,
                        title=title,
                        created=page.get("createdDateTime", ""),
                        edited=page.get("lastModifiedDateTime", ""),
                    )
                    assets_dir = s_dir / f"{p_slug}_assets"
                    rewritten_html = download_page_assets(token, html_text, assets_dir, state=state)
                    write_text_atomic(html_path, rewritten_html)
                    state.mark(page_work_id, "ok")

                # Pull ink markup and persist as PNG preview in page HTML when available.
                ink_work_id = f"ink:{p_id}"
                ink_png_path = s_dir / f"{p_slug}_ink.png"
                if not (state.is_ok(ink_work_id) and ink_png_path.exists()):
                    log(f"    [Page {p_idx}/{len(pages)}] fetch ink: {title}")
                    try:
                        multipart = graph_get_bytes(
                            token,
                            f"{GRAPH_BASE}/me/onenote/pages/{p_id}/content?includeInkML=true",
                        )
                        inkml = extract_inkml_from_multipart(multipart)
                        if inkml:
                            inkml_path = s_dir / f"{p_slug}.inkml.xml"
                            write_bytes_atomic(inkml_path, inkml)
                            if render_inkml_to_png(inkml, ink_png_path):
                                rewritten_html = inject_ink_preview_image(
                                    rewritten_html,
                                    f"{ink_png_path.name}",
                                )
                                write_text_atomic(html_path, rewritten_html)
                        state.mark(ink_work_id, "ok")
                    except Exception as ex:
                        state.mark(ink_work_id, "retryable_fail", str(ex))
                        log(f"Skipping ink preview for page {p_id}: {ex}")
                else:
                    log(f"    [Page {p_idx}/{len(pages)}] resume ink: {title}")

                s_entry["pages"].append(
                    {
                        "id": p_id,
                        "title": title,
                        "level": page.get("level", 0),
                        "order": page.get("order"),
                        "meta": str(page_meta_path.relative_to(out_dir)),
                        "html": str(html_path.relative_to(out_dir)),
                    }
                )

            write_json_atomic(s_dir / "_section.json", s_entry)
            nb_entry["sections"].append(s_entry)

        manifest["notebooks"].append(nb_entry)

    write_json_atomic(out_dir / "manifest.json", manifest)
    log(
        "Export summary: "
        f"skipped_notebooks={skipped_notebooks}, skipped_sections={skipped_sections}, skipped_pages={skipped_pages}"
    )


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Export OneNote via Microsoft Graph")
    p.add_argument("--client-id", required=True, help="Azure app registration client/application ID")
    p.add_argument("--tenant", default="common", help="Tenant ID or 'common' (default)")
    p.add_argument("--out", default="export", help="Output directory")
    p.add_argument("--cache", default=".token_cache.json", help="MSAL token cache file")
    p.add_argument(
        "--request-delay",
        type=float,
        default=1.2,
        help="Seconds to wait before each Graph HTTP request (useful to reduce throttling)",
    )
    p.add_argument(
        "--max-retries",
        type=int,
        default=20,
        help="Max retry attempts per throttled/unavailable Graph request",
    )
    p.add_argument(
        "--max-backoff",
        type=int,
        default=120,
        help="Max seconds for exponential backoff between retries",
    )
    return p.parse_args()


def main() -> int:
    global REQUEST_DELAY_S, MAX_RETRIES, MAX_BACKOFF_S
    args = parse_args()
    out_dir = pathlib.Path(args.out).resolve()
    cache_path = pathlib.Path(args.cache).resolve()
    REQUEST_DELAY_S = max(0.0, float(args.request_delay))
    MAX_RETRIES = max(1, int(args.max_retries))
    MAX_BACKOFF_S = max(1, int(args.max_backoff))
    state = StateStore(out_dir / ".export_state.jsonl")

    token = acquire_token(args.client_id, args.tenant, cache_path)
    export_onenote(token, out_dir, state=state)

    log(f"Export complete: {out_dir}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
