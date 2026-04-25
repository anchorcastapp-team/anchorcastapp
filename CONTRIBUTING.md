# Contributing to AnchorCast

Thank you for contributing to **AnchorCast** — a free, open-source church presentation platform built to serve churches of all sizes.

We welcome contributions from developers, designers, church media operators, testers, documentation writers, and anyone with a heart to help the local church through technology.

---

## Ways to Contribute

- 🐛 **Bug reports** — spotted something broken? Tell us.
- 💡 **Feature requests** — have a real church use case? We want to hear it.
- 🔧 **Code contributions** — bug fixes, new features, performance improvements.
- 🎨 **UI/UX improvements** — making the operator experience simpler and more reliable.
- 📖 **Documentation** — setup guides, workflow explanations, troubleshooting tips.
- 🧪 **Testing in real services** — the best feedback comes from Sunday mornings.
- 🎨 **Themes and presets** — share what works well in your church.

---

## Before You Start

1. Search existing issues before opening a new one — your question may already be answered.
2. For significant changes, open an issue first to discuss the approach before writing code.
3. Keep pull requests focused — one problem, one fix. Smaller PRs get reviewed faster.
4. Be kind, respectful, and constructive. AnchorCast exists to serve the church — let that spirit carry into our collaboration.

---

## Development Setup

### Requirements

- **Node.js** 18+
- **npm**
- **Electron** environment for desktop testing
- **Python 3.10–3.12** — optional, for local Whisper transcription

### Install

```bash
npm install
```

### Run the desktop app

```bash
npm start
```

### Run in web mode

```bash
npm run web
```

### Build installers

```bash
npm run build:win    # Windows installer (includes AnchorCast Timer shortcut)
npm run build:timer  # Standalone AnchorCast Timer installer
npm run build:mac    # macOS DMG
```

> **Before building on Windows:** run `setup_whisper.bat` to create the `python\` folder. Optionally add `vc_redist.x64.exe` and a `models\` folder for a fully offline installer.

---

## Project Structure

| Path | Description |
|---|---|
| `src/main.js` | Electron main process — IPC, projection, timer engine, NDI |
| `src/timer-main.js` | Standalone AnchorCast Timer entry point |
| `src/preload.js` | Electron IPC bridge |
| `src/renderer/js/app.js` | Main operator workflow and UI logic |
| `src/renderer/projection.html` | Projection rendering layer |
| `src/renderer/countdown-window.html` | Timer control UI |
| `src/renderer/js/ai-detection.js` | Bible verse detection engine |
| `server.js` | Web-mode Express backend |
| `whisper_server.py` | Local Whisper AI transcription server |
| `installer.nsh` | NSIS installer hooks (VC++, models, timer shortcut) |

---

## Coding Guidelines

### General principles

- Keep patches small and targeted — one concern per change.
- Prefer clear, readable code over clever shortcuts.
- Preserve offline-first behavior — the app must work without internet.
- Do not break desktop mode while fixing web mode, or vice versa.
- Keep the operator workflow simple and predictable — this runs during live services.

### IPC discipline

- Every channel invoked in `preload.js` must have a matching handler in `main.js`.
- Every event sent from `main.js` must be in the `preload.js` allowlist.
- Avoid duplicate `ipcMain.handle` registrations — they silently override each other.
- New channels added to `preload.js` must also be stubbed in `timer-main.js` if applicable.

### Timer and projection changes

Take extra care with:

- **State synchronization** — timer state lives in the main process; projection and timer windows pull from it.
- **Window reopen behavior** — projection and timer windows can be closed and reopened at any time; always sync state on `did-finish-load`.
- **Duplicate events** — coalesce rapid updates; avoid redundant redraws.
- **Standalone vs embedded timer** — `timer-main.js` is a separate entry point; IPC changes must be reflected in both `main.js` and `timer-main.js`.

### UI changes

Test the full operator loop:
- Preview → Project Live → Clear → navigate slides
- What the operator sees vs. what the congregation sees
- Rapid repeated actions (fast clicking, quick navigation)
- Both the main-app timer and the standalone AnchorCast Timer

---

## Pull Request Guidelines

A good pull request should include:

- **What problem is being solved** — link to the issue if one exists
- **Files changed** — a brief list
- **How it was tested** — desktop mode, web mode, or both? Tested in a real service?
- **Screenshots or GIFs** — for any UI changes

### Examples of good PRs

> *Fixes timer desync after projection window reopen. Updates `src/main.js` and `src/renderer/projection.html`. Tested with projection open, closed, and reopened. Standalone timer verified still working.*

> *Adds dark mode support for the Operator Command Center. Updates `src/renderer/js/app.js` and `src/renderer/css/anchorcast-theme.css`. Tested on Windows 11 with system dark mode.*

---

## Bug Reports

A helpful bug report includes:

- What you **expected** to happen
- What **actually** happened
- **Steps to reproduce** — the simpler, the better
- Whether it occurred in **desktop mode** or **web mode**
- Your **OS and AnchorCast version**
- **Screenshots, logs, or a screen recording** if available

### Examples of clear bug reports

> *"Projection window stops updating after being closed and reopened during a live service."*

> *"Remote control shows Connected but does not advance slides in Songs mode."*

> *"Timer starts from zero instead of the correct remaining time after reopening the projection window."*

---

## Feature Requests

Feature requests are welcome — especially those grounded in real Sunday morning situations.

Please include:

- The **real church use case** — what problem does this solve for your team?
- **Who benefits** — operator, pastor, worship leader, or congregation?
- Which part of the system it touches — projection, timer, remote, transcription, songs, or schedule
- Whether it needs to **work offline**

---

## Testing Priorities

When contributing, pay special attention to these areas:

| Area | Why It Matters |
|---|---|
| Scripture detection | Core mission — must be accurate and fast |
| Live projection | Runs during services — no tolerance for crashes or freezes |
| Timer (embedded + standalone) | Visible to the entire congregation |
| Projection reopen sync | Operators close and move windows constantly |
| Song manager | High-frequency Sunday morning workflow |
| Remote control | Pastors and operators depend on this during services |
| Offline behavior | Many churches have limited or no internet access |
| Windows installer | First impression for new churches |

---

## Documentation

Documentation contributions are just as valuable as code. Useful areas include:

- Getting started guide for first-time church media operators
- Projection setup for common screen and projector configurations
- Timer setup and usage
- Song import from EasyWorship and ProPresenter
- NDI setup for OBS and vMix
- Remote control setup and PIN configuration
- Troubleshooting common issues
- Church media team onboarding checklist

---

## Community Standards

AnchorCast exists to serve the local church. Please bring that same spirit into how we work together:

- Be kind and patient — contributors come from many backgrounds and skill levels
- Offer constructive feedback — critique ideas, not people
- Remember that the end users are church volunteers running live Sunday morning services

---

## Changelog

All notable changes between versions are documented in [CHANGELOG.md](CHANGELOG.md).

---

## License

By contributing, you agree that your contributions will be released under the project's **MIT License** — keeping AnchorCast free for every church, forever.
