## AnchorCast v1.7.0 — Theme Editor Bug Fix
### 🐞 Bug Fix

This release fixes a longstanding issue where background images and videos chosen in the Theme Editor would silently fail to display.

---

### 🐞 Bug Fix

- **Theme Editor background image/video not displaying** — background files are served through AnchorCast's `media://` protocol, which only allows files that live inside AnchorCast's own data folder. A picture or video picked from anywhere else (Downloads, Desktop, external drives, etc.) was silently rejected, so the background never rendered — not even in the Theme Editor's own preview, which made this look like a deeper bug than it was. Background files picked via the file dialog or dragged and dropped are now automatically copied into AnchorCast's data folder first, so any file from any location works correctly.

---

### 📦 Downloads
| Platform | Variant | Description |
|----------|---------|-------------|
| Windows | **Full** | Includes Python + Whisper model (~600 MB) — ready to use offline immediately |
| Windows | **Light** | Includes Python only (~200 MB) — downloads Whisper model on first use |
| macOS Apple Silicon | **Full** | For M1/M2/M3 Macs with model bundled |
| macOS Apple Silicon | **Light** | For M1/M2/M3 Macs, downloads model on first use |
| macOS Intel | **Full** | For Intel Macs with model bundled |
| macOS Intel | **Light** | For Intel Macs, downloads model on first use |

> **Upgrading from v1.6.0?** Your data, settings, Whisper models, songs, transcripts, presets, and installed Bible translations are fully preserved — no extra steps needed. If you had trouble setting a Theme background image or video, this release fixes it.

> **Upgrading from v1.5.0 or earlier?** See the [v1.6.0 release notes](https://github.com/anchorcastapp-team/anchorcastapp/releases/tag/v1.6.0) for the Bible import fix, transcript autosave, and projection display fix first.

[![Download anchorcastapp](https://a.fsdn.com/con/app/sf-download-button)](https://sourceforge.net/projects/anchorcastapp/files/v1.7.0/)
