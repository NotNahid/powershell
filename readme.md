<div align="center">

# Nahid Power Utilities

A Windows utility toolkit built with PowerShell, Windows Forms, and a live remote catalog.

[![Stars](https://img.shields.io/github/stars/NotNahid/powershell?style=flat-square)](https://github.com/NotNahid/powershell/stargazers)
[![Forks](https://img.shields.io/github/forks/NotNahid/powershell?style=flat-square)](https://github.com/NotNahid/powershell/network/members)
[![Issues](https://img.shields.io/github/issues/NotNahid/powershell?style=flat-square)](https://github.com/NotNahid/powershell/issues)
[![Last commit](https://img.shields.io/github/last-commit/NotNahid/powershell?style=flat-square)](https://github.com/NotNahid/powershell/commits/main)
[![Commit activity](https://img.shields.io/github/commit-activity/m/NotNahid/powershell?style=flat-square)](https://github.com/NotNahid/powershell/graphs/commit-activity)
[![Repository size](https://img.shields.io/github/repo-size/NotNahid/powershell?style=flat-square)](https://github.com/NotNahid/powershell)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=flat-square&logo=powershell&logoColor=white)](https://learn.microsoft.com/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?style=flat-square&logo=windows&logoColor=white)](https://www.microsoft.com/windows)

</div>

---

## Overview

Nahid Power Utilities is a collection of small Windows tools with one graphical launcher. The launcher retrieves a JSON catalog over HTTP, validates each entry, renders the available utilities, and downloads a selected script only when it is launched.

The repository also includes standalone PowerShell tools, Python companions, Windows context-menu utilities, and a web landing page.

<div align="center">

<a href="https://github.com/NotNahid/powershell"><img src="https://github-stats-extended.vercel.app/api/pin/?username=NotNahid&repo=powershell&theme=transparent&hide_border=true" alt="Repository statistics" width="420"></a>

<a href="https://github.com/NotNahid"><img src="https://github-stats-extended.vercel.app/api?username=NotNahid&show_icons=true&theme=transparent&hide_border=true&rank=false" alt="GitHub statistics" width="420"></a>

<a href="https://github.com/NotNahid"><img src="https://github-stats-extended.vercel.app/api/top-langs/?username=NotNahid&layout=compact&theme=transparent&hide_border=true&langs_count=8" alt="Top languages" width="420"></a>

[![GitHub streak](https://streak-stats.demolab.com/?user=NotNahid&theme=transparent&hide_border=true)](https://github.com/DenverCoder1/github-readme-streak-stats)

</div>

## Install

### Run the hosted launcher

The short command below downloads the current launcher through the project's redirect URL.

```
irm is.gd/notnahid | iex
```

Only use this option after reviewing and trusting the source. It downloads and executes remote PowerShell code.

### Clone and run locally

```
git clone https://github.com/NotNahid/powershell.git
Set-Location powershell
Set-ExecutionPolicy -Scope Process Bypass
.\ultimate.ps1
```

Running locally is the recommended option for inspection, development, and offline changes to the catalog.

## What the launcher does

```
utilities.json -> HTTP request -> validation -> Windows Forms UI
                                           |
                                  select a utility
                                           |
                                  download script
                                           |
                                  start background job
```

`ultimate.ps1` currently provides:

- Remote catalog loading with a local fallback

- Required-field validation for every catalog entry

- Searchable utility cards

- Category and tag metadata

- Dark and light themes

- Detail dialogs and tooltips

- Background downloads and launches

- Status updates while a utility is starting

## Utility catalog

The dashboard is driven by [`Main Links File Json/utilities.json`](./Main%20Links%20File%20Json/utilities.json). The catalog is intentionally separate from the UI so new tools can be added without changing `ultimate.ps1`.

| Utility | Category | Description |
| --- | --- | --- |
| Ghost Typer | Automation | Auto-type text anywhere |
| Portfolio | Personal | Open the portfolio website |
| Rename Tool | File Management | Bulk rename files |
| FFmpeg With GUI | Editing | Use FFmpeg through a graphical interface |
| Network Speed Monitor | Network | Monitor network speed |
| File Organizer | File Management | Organize files with a dedicated tool |
| Auto Clicker | Productivity | Automate repeated clicks |
| Live Wallpaper | Fun | Launch the live wallpaper workflow |
| Quick Tools in System Tray | Productivity | Keep utilities available from the system tray |

## Live APIs and endpoints

The launcher uses ordinary HTTP endpoints instead of a hard-coded tool list.

| Purpose | Endpoint |
| --- | --- |
| Hosted launcher | `https://is.gd/notnahid` |
| Live catalog | `https://raw.githubusercontent.com/NotNahid/powershell/refs/heads/main/Main%20Links%20File%20Json/utilities.json` |
| Repository API | `https://api.github.com/repos/NotNahid/powershell` |
| Latest commits API | `https://api.github.com/repos/NotNahid/powershell/commits` |
| Repository contents API | `https://api.github.com/repos/NotNahid/powershell/contents` |
| Stats cards | `https://github-stats-extended.vercel.app/api` |
| Streak stats | `https://streak-stats.demolab.com/` |
| Badge API | `https://img.shields.io/` |

Example: retrieve the current catalog from PowerShell.

```
$catalogUrl = "https://raw.githubusercontent.com/NotNahid/powershell/refs/heads/main/Main%20Links%20File%20Json/utilities.json"
$catalog = Invoke-RestMethod -Uri $catalogUrl
$catalog | Format-Table Name, Category, Desc
```

Example: query repository metadata from the GitHub REST API.

```
$repo = Invoke-RestMethod `
  -Uri "https://api.github.com/repos/NotNahid/powershell" `
  -Headers @{ Accept = "application/vnd.github+json" }

$repo | Select-Object full_name, stargazers_count, forks_count, open_issues_count
```

The statistics shown near the top of this page are also generated dynamically by GitHub-compatible badge and activity APIs rather than being hard-coded in the README.

### Third-party widget sources

The visual widgets above are live images generated when GitHub renders this page:

- [Shields.io](https://shields.io/) provides repository metadata badges such as stars, last commit, commit activity, and repository size.

- [GitHub Stats Extended](https://github.com/stats-organization/github-stats-extended) provides the repository card, account statistics, and language card. It is the maintained successor to the older `github-readme-stats` service.

- [GitHub Readme Streak Stats](https://github.com/DenverCoder1/github-readme-streak-stats) provides total contributions, current streak, and longest streak data through `streak-stats.demolab.com`.

These services are external dependencies. If a provider is rate-limited or unavailable, the project itself still works normally; only the visual cards may be temporarily unavailable.

## Add a utility

Add the script to the repository, then append an object to `utilities.json`:

```json
{
  "Name": "My Utility",
  "Link": "https://raw.githubusercontent.com/your-user/your-repo/main/my-utility.ps1",
  "Desc": "Short description of the utility",
  "Category": "Automation",
  "Tags": ["windows", "productivity", "utility"]
}
```

The launcher validates these fields:

| Field | Type | Required |
| --- | --- | --- |
| `Name` | string | Yes |
| `Link` | URL string | Yes |
| `Desc` | string | Yes |
| `Category` | string | Yes |
| `Tags` | array of strings | Yes |

Use raw HTTPS URLs, keep descriptions short, and test the script independently before adding it to the live catalog.

## Repository layout

```
.
├── ultimate.ps1
├── Main Links File Json/
│   └── utilities.json
├── Auto Clicker.ps1
├── File Organizer.ps1
├── FFmpeg With GUI.ps1
├── Ghost Typer.ps1
├── Network Speed Monitor.ps1
├── Quick Tools in System Trey.ps1
├── Rename Tool.ps1
├── live_wallpaper.ps1
├── Portfolio.ps1
├── Right-click/
├── Python/
├── index.html
└── SECURITY.md
```

## Requirements

- Windows 10 or newer

- Windows PowerShell 5.1 or PowerShell 7+

- Windows Forms support for the graphical launcher

- Internet access when using the live catalog or remote script links

- Python for the Python-based companion utilities

## Development

```
# Clone the project
git clone https://github.com/NotNahid/powershell.git
Set-Location powershell

# Run a local utility
.\Rename Tool.ps1

# Run the dashboard
.\ultimate.ps1

# Inspect the catalog
Get-Content ".\Main Links File Json\utilities.json" -Raw | ConvertFrom-Json
```

Before opening a pull request:

1. Test the script on Windows.

1. Check that it does not contain credentials or machine-specific paths.

1. Document downloads, permissions, and side effects.

1. Add the catalog entry with an accurate category and description.

1. Keep the change focused.

## Security

The launcher can download and execute remote PowerShell scripts. Treat every link in the catalog as executable code.

Review scripts before running them, use the local clone when auditing changes, and avoid running unfamiliar tools on systems containing sensitive data. For vulnerability reports, see [`SECURITY.md`](./SECURITY.md). For normal bugs, open an [issue](https://github.com/NotNahid/powershell/issues) with reproduction steps and environment details.

## Related project page

The repository includes [`index.html`](./index.html), an interactive web page that presents the toolkit and its utilities in a browser.

## Author

**Nahid**

- GitHub: [@NotNahid](https://github.com/NotNahid)

- Website: [notnahid.rf.gd](https://notnahid.rf.gd/)

- Email: [nahidul.live@gmail.com](mailto:nahidul.live@gmail.com)

---

<div align="center">

[View source](https://github.com/NotNahid/powershell) · [Open an issue](https://github.com/NotNahid/powershell/issues) · [Browse the catalog](./Main%20Links%20File%20Json/utilities.json)

</div>
