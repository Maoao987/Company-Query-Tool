# Company Query Tool Handoff

## Current status

- Current version: `1.1.3`
- GitHub repo: `https://github.com/Maoao987/Company-Query-Tool`
- Current release strategy:
  - same-version hotfix: overwrite existing release assets
  - new version release: bump patch version and publish a new release

## What is already done

- Windows installer packaging is working
- GitHub repo is connected and source code is already pushed
- Built-in updater is connected to GitHub Releases
- Light mode is enforced for the UI
- Batch query supports:
  - unified business number
  - stock code
  - company name
- Stock price date supports calendar input
- If selected date is not a trading day, the system automatically aligns to the nearest prior trading day
- Director/supervisor data is included in company info and reports
- PDF export, Excel export, CSV export, and source snapshot flow are already wired in
- Installer version and in-app version are now aligned through shared version files

## Important implementation notes

### Version control

- Main source of truth: `version.txt`
- Installer version sync helper:
  - `sync_version.ps1`
- Inno Setup include file:
  - `version.iss.inc`

If you want to change the version, update `version.txt` first, then run the release script.

### Release automation

- Release script:
  - `publish_release.ps1`
- Release usage guide:
  - `RELEASE_WORKFLOW.md`

Typical commands:

```powershell
powershell -ExecutionPolicy Bypass -File .\publish_release.ps1 -Mode fix
```

```powershell
powershell -ExecutionPolicy Bypass -File .\publish_release.ps1 -Mode release
```

```powershell
powershell -ExecutionPolicy Bypass -File .\publish_release.ps1 -Mode release -Version 1.1.4
```

### Important release behavior

- If you keep the same version number, GitHub Release assets can be overwritten.
- But users already on the exact same version will not be prompted by the updater again.
- If you want users to see an update automatically, publish a new version number.

## Known caveats

- Emerging stock (`興櫃`) official history uses average traded price, not the same closing-price definition used by listed/OTC markets.
- Some stock-to-company resolution still depends on matching company name with official and registry data. Most common cases are already handled, but edge cases may still require manual review.
- If findbiz rate-limits or token retrieval fails, registry queries can temporarily fail and should be retried later.

## Suggested setup on another computer

1. Install:
   - Git
   - Python 3.12
   - GitHub CLI (`gh`)
   - Inno Setup 6

2. Clone repo:

```powershell
git clone https://github.com/Maoao987/Company-Query-Tool.git
cd Company-Query-Tool
```

3. Login GitHub CLI:

```powershell
gh auth login
```

4. When you want to build or publish:

```powershell
powershell -ExecutionPolicy Bypass -File .\publish_release.ps1 -Mode fix
```

or

```powershell
powershell -ExecutionPolicy Bypass -File .\publish_release.ps1 -Mode release
```

## Files you will touch often

- `app.py`
- `company_query.py`
- `pdf_report.py`
- `findbiz_scraper.py`
- `update_manager.py`
- `version.txt`
- `publish_release.ps1`

## If you continue work with Codex later

Tell Codex:

- this repo is already connected to GitHub
- version source of truth is `version.txt`
- release script is `publish_release.ps1`
- keep hotfixes on same version only when you intentionally want asset replacement without updater bump
- for normal user-visible updates, bump to a new patch version
