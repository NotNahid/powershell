# ⚡ Nahid's PowerShell Toolkit

A collection of small, useful Windows utilities — all launched from one slick dashboard, no cloning or setup required.

```powershell
irm is.gd/notnahid | iex
```

Run that in PowerShell and the dashboard opens, ready to launch any tool in the collection.

![PowerShell](https://img.shields.io/badge/PowerShell-69.5%25-5391FE?logo=powershell&logoColor=white)
![Python](https://img.shields.io/badge/Python-30.5%25-3776AB?logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/platform-Windows-0078D6)
![Stars](https://img.shields.io/github/stars/NotNahid/powershell?style=social)

---

## ✨ What is this?

This repo is a grab-bag of standalone PowerShell (and a few Python) scripts, each solving one small annoying problem — renaming files, organizing downloads, checking network speed, and so on. Instead of digging through the repo for the script you want, `ultimate.ps1` acts as a **dashboard**: a dark/light-themed GUI that lists every tool, lets you search and read what each one does, and downloads + runs it on demand in the background.

It's basically a personal app store for your own scripts.

## 🚀 Quick Start

Open PowerShell and run:

```powershell
irm is.gd/notnahid | iex
```

This downloads and launches `ultimate.ps1`, which pulls the live tool catalog and shows it in a card-based UI. Click **Run** on any tool, confirm, and it launches in the background — no manual downloading, no `git clone`.

> **Heads up:** this one-liner downloads and executes remote code on your machine. That's the whole point of the convenience, but you (and anyone else running it) should know that's what `irm | iex` does. Feel free to open `ultimate.ps1` and read it first if you want to see exactly what it does before running it.

## 🧰 What's inside

| Tool | What it does |
|---|---|
| **ultimate.ps1** | The main dashboard — searchable, themeable GUI launcher for every tool in this repo |
| **Portfolio.ps1** | A little desktop app showing an About / Projects / Contact view — basically a PowerShell-native portfolio |
| **Ghost Typer.ps1** | Auto-types text anywhere on screen |
| **type-pro.ps1** | An enhanced/extended auto-typing utility |
| **File Organizer.ps1** | Sorts and organizes files automatically (e.g. by type or date) |
| **Rename Tool.ps1** | Batch renames files |
| **Network Speed Monitor.ps1** | Tracks and displays live network/internet speed |
| **FFmpeg With GUI.ps1** | A graphical front-end for running FFmpeg commands without the command line |
| **Quick Tools in System Trey.ps1** | A system tray menu for quick access to small utilities |
| **live_wallpaper.ps1** | Sets an animated/live wallpaper on the desktop |
| **Python/** | Companion Python scripts used by some of the tools |
| **Main Links File Json/** | `utilities.json` — the live catalog the dashboard reads from |

## 🛠 How the dashboard works

`ultimate.ps1` doesn't hardcode its tool list. On launch it:

1. Fetches `utilities.json` from this repo (falls back to a built-in default list if that fails)
2. Validates each entry (name, link, description, category, tags)
3. Renders one card per tool in a Catppuccin-inspired dark/light UI
4. On **Run**, downloads that tool's script and executes it as a background job, with live status updates (Downloading → Starting → Launching → Running → Done)

That means adding a new tool to the dashboard is just a matter of adding an entry to `utilities.json` — no need to touch `ultimate.ps1` itself.

### Adding your own tool

Add an object like this to `Main Links File Json/utilities.json`:

```json
{
  "Name": "My Cool Tool",
  "Link": "https://raw.githubusercontent.com/you/repo/main/my-tool.ps1",
  "Desc": "One-line description of what it does",
  "Category": "Automation",
  "Tags": ["productivity", "windows"]
}
```

It'll show up in the dashboard automatically next time it loads the catalog.

## 📋 Requirements

- Windows (Windows Forms is used for the GUI)
- PowerShell 5.1+ (ships with Windows by default)
- Internet connection (to fetch the catalog and individual tool scripts)

## 🤝 Contributing

Got a small utility you think belongs here? Open a PR adding the script plus an entry in `utilities.json`. Keep scripts self-contained where possible so they can be downloaded and run independently.

## 👤 Author

**Nahid** — Web developer (front-end, diving into full-stack)

- GitHub: [@NotNahid](https://github.com/NotNahid)
- Website: [nahid.rf.gd](https://nahid.rf.gd/)
- Email: nahidul.live@gmail.com

## 🐞 Known Issues

Some tools have a bug where they don't close properly when you hit the close/exit button — being worked on. Until it's fixed, if a tool won't close, open **Task Manager** (`Ctrl+Shift+Esc`), find the process, and click **End Task**.

## ⚠️ Disclaimer

These are personal utility scripts shared as-is. Review any script before running it on a machine you care about, especially since the dashboard's whole model is "download and execute remote code." No warranty, no guarantees — just tools that made my own workflow easier.