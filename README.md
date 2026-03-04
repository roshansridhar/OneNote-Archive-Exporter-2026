# OneNote cold-storage export

Export OneNote notebooks through Microsoft Graph into a local folder with:
- page metadata (`.meta.json`)
- page HTML (`.html`)
- handwriting preview image when available (`*_ink.png`)
- optional local asset folders (`*_assets`) for images/files
- top-level `manifest.json`

## Prerequisites

- Python 3.9+
- A Microsoft Entra app registration
- A Microsoft account that can access your OneNote notebooks

## 1) Create and configure the Entra app
1. In Azure portal: `Microsoft Entra ID` -> `App registrations` -> `New registration`. https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-register-app
2. Name it (example: `onenote-migration-app`).
3. Under **Supported account types**, choose based on your scenario:
   - Work/school only: `Accounts in this organizational directory only`
   - Personal Microsoft account (Outlook/Hotmail/Live) (I chose this): `Accounts in any organizational directory and personal Microsoft accounts`
4. Save the **Application (client) ID**.
5. In `Manage` -> `API permissions`, add delegated Microsoft Graph permissions:
   - `Notes.Read`
   - `User.Read`
   https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-configure-app-access-web-apis?utm_source=chatgpt.com
6. In `Authentication`, enable public client/device flow:
   - `Allow public client flows` = `Yes`

### Personal-account compatibility notes

If personal account sign-in fails:
- confirm supported account type includes personal Microsoft accounts
- in app `Manifest`, ensure:
  - `"requestedAccessTokenVersion": 2`
- then retry auth with `--tenant consumers`


## 2) Install

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## 3) Run export

### Work/school tenant

```bash
python3 onenote_export.py --client-id "<APP_CLIENT_ID>" --tenant "<TENANT_ID>" --out export
```

### Personal Microsoft account

```bash
python3 onenote_export.py --client-id "<APP_CLIENT_ID>" --tenant consumers --cache .token_cache_personal.json --out export
```

On first run you will get a device login URL + code. Complete sign-in in browser, then export continues.

## Resume and partial-download safety

The exporter is resumable by default. It writes:
- `.export_state.jsonl` in the output directory (append-only work log)
- files using atomic temp-write + rename (prevents half-written final files)

If a run is interrupted (Ctrl+C, laptop sleep, network error, 429 burst), rerun the same command with the same `--out` directory. Already completed units (`page_html:*`, `asset:*`, `ink:*`) are skipped, and remaining work continues.

Quick checks:

```bash
tail -n 20 export/.export_state.jsonl
```

```bash
rg -n '"status": "retryable_fail"|"status": "permanent_fail"' export/.export_state.jsonl
```


## 4) Images and handwriting behavior

Notes:
- full export downloads Graph-hosted resources into `*_assets` folders and rewrites page HTML links.
- full export also attempts InkML extraction and writes `*_ink.png` into page HTML when pen data is available.
- throttling-friendly defaults are enabled; you can slow further with custom flags.

Very slow profile (use when a fresh full pull keeps hitting `429`):

```bash
python3 onenote_export.py \
  --client-id "<APP_CLIENT_ID>" \
  --tenant consumers \
  --cache .token_cache_personal.json \
  --out export \
  --request-delay 2.0 \
  --max-retries 30 \
  --max-backoff 180
```

## Output structure

```text
export/
  manifest.json
  <Notebook>/
    <Section>/
      <Page>.meta.json
      <Page>.html
      <Page>_assets/
        asset_0001.jpg
        asset_0002.png
```

## Troubleshooting

- `AADSTS50059`:
  - tenant/account type mismatch; use explicit tenant ID or `consumers` for personal account.
- `30121 The tenant does not have a valid SharePoint license`:
  - work tenant/license issue; use an account with OneNote/SharePoint service enabled.
- `429` / `20166 too many requests`:
  - wait for cooldown and rerun with higher `--request-delay` and `--max-retries`.
  - rerun in the same `--out` folder to continue from `.export_state.jsonl` without redoing completed items.

## Long-term cold storage recommendations

For archival durability:
- keep raw export (`JSON + HTML + assets`)
- also create normalized plain text/Markdown copies
- keep periodic `tar` snapshots with `sha256` checksums
- store in at least 3 locations (2 cloud + 1 offline)

## Security

- Do not commit secrets/token caches:
  - `.token_cache*.json`
  - app client secrets
- Rotate any secret that has ever been shared in chat/logs.

## HTML tree viewer

Generate a local viewer for browsing notebooks/sections/pages:

```bash
python3 build_viewer.py --export-dir export --out viewer.html
```

Then open:
- `export/viewer.html`

Features:
- collapsible tree of folders and note pages
- click a page to preview in a right-side pane
- open selected page in a new tab
