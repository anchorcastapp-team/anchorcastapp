# ⚓ AnchorCast v1.1.0
### *Free & Open-Source AI Church Presentation Software*

<p align="center">
  <a href="https://github.com/gbkeku/anchorcastapp/stargazers">
    <img src="https://img.shields.io/github/stars/gbkeku/anchorcastapp?style=for-the-badge" />
  </a>
  <a href="https://github.com/gbkeku/anchorcastapp/releases">
    <img src="https://img.shields.io/github/v/release/gbkeku/anchorcastapp?style=for-the-badge&label=version" />
  </a>
  <a href="https://github.com/gbkeku/anchorcastapp/blob/main/LICENSE">
    <img src="https://img.shields.io/github/license/gbkeku/anchorcastapp?style=for-the-badge" />
  </a>
  <a href="https://github.com/gbkeku/anchorcastapp/issues">
    <img src="https://img.shields.io/github/issues/gbkeku/anchorcastapp?style=for-the-badge" />
  </a>
  <a href="https://www.anchorcastapp.com">
    <img src="https://img.shields.io/badge/website-anchorcastapp.com-blue?style=for-the-badge" />
  </a>
</p>

<p align="center">
  <strong>AnchorCast listens to your preacher, detects Bible verses in real time, and displays them on your projection screen — automatically.</strong><br/>
  Songs · Scripture · Media · Timer · Transcription · NDI — all in one free app.
</p>

> 🆓 **Free forever. No subscription. No trial. No credit card.**

---

## 🎬 Live Demos

### 🔥 AI Scripture Detection — verses appear the instant the preacher speaks them
![Scripture Detection](./assets/gifs/scripture-detection.gif)

### 🎵 Song Projection Workflow
![Song Projection](./assets/gifs/songs-projection.gif)

### ⏱ Countdown Timer — synced to projection, no lag
![Timer Sync](./assets/gifs/timer-sync.gif)

### 📱 Wi-Fi Remote Control — control from any phone or tablet
![Remote Control](./assets/gifs/remote-control.gif)

---

## ⚡ Quick Start

```bash
git clone https://github.com/gbkeku/anchorcastapp.git
cd anchorcastapp
npm install
npm start
```

On Windows, double-click `start.bat`.

---

## 📥 Download

👉 **[anchorcastapp.com](https://www.anchorcastapp.com)** — Windows installer (Python + Whisper bundled, fully offline)  
👉 **[GitHub Releases](https://github.com/gbkeku/anchorcastapp/releases)** — all versions

---

## ✨ Features

### 🎙 AI Transcription & Scripture
- **Real-time sermon transcription** — local Whisper AI (offline) or Deepgram cloud (online)
- **Automatic Bible verse detection** — detects direct quotes and paraphrased references as the preacher speaks
- **AI context search** — find verses by topic, not just reference
- **Adaptive Memory** — learns speaker vocabulary, accents, and correction patterns over time
- **Detection Review** — approve, reject, or teach custom trigger phrases for future services

### 🎵 Worship Tools
- **Song Manager** — full library with lyrics editor, auto slide formatting, and Genius lyrics search
- **Import from EasyWorship or ProPresenter XML** — or paste raw lyrics for instant AI formatting
- **Theme Designer** — custom fonts, colors, backgrounds, and logos; per-category themes for Scripture, Songs, and Presentations
- **Presentation Editor** — 1920×1080 canvas designer with text, images, shapes, and backgrounds
- **Service Schedule** — drag-and-drop planner to stage verses, songs, and media in order

### 📺 Projection System
- **Live preview + projection** — operator sees the next slide while the congregation sees the current one
- **Multi-monitor support** — drag the projection window to any connected screen or projector
- **NDI output** — send to OBS, vMix, or Wirecast for live streaming
- **MJPEG fallback** — for any browser-based streaming client
- **Logo, alerts, and live captions** — overlay layers managed independently of the main display

### 📱 Remote Control
- Access at `http://[your-ip]:5000/remote` from any phone or tablet on Wi-Fi
- Live projection preview, Scripture / Songs / Media mode tabs
- Prev / Next / Clear / Go Live — full control from anywhere in the building
- PIN authentication with role-based access (admin, scripture, songs, media, monitor)

### ⏱ Timer System
- **Countdown and count-up** modes with custom title, font, color, and position
- **Displayed on the projection screen** — fully synced, no lag
- **Flashes red** when time expires and shows a closing message to the speaker
- **Clock & date overlay** for pre-service display
- **Standalone AnchorCast Timer** — installs alongside the main app as a separate shortcut

### 📝 Sermon Intelligence
- **Live transcript** — auto-saved after every session
- **AI Sermon Notes** — generate structured notes from the transcript with one click
- **Sermon Intelligence** — AI title suggestions, keyword analysis, and structure insights
- **Service Archive** — searchable history of every service by title, speaker, or verse
- **Analytics Dashboard** — cross-service stats: most-used books, quoted verses, speaker patterns

### 🔒 Reliability
- **Offline-first** — works entirely without internet using local AI and local Bible data
- **KJV bundled** by default; import any translation (NKJV, NIV, ESV, NLT, NASB, ASV) via JSON
- **Role-based remote access** — PIN-protected operator controls
- **NDI + MJPEG** — professional and fallback streaming options built in

---

## 💻 Desktop vs Web Mode

| Feature | Desktop | Web |
|---|---|---|
| Bible display + projection | ✅ | ✅ |
| Song management + Genius search | ✅ | ✅ |
| Theme designer | ✅ | ✅ |
| Presentation editor | ✅ | ✅ |
| Service schedule | ✅ | ✅ |
| Remote control | ✅ PIN + roles | ✅ open |
| NDI output | ✅ | ❌ |
| Local Whisper transcription | ✅ | ❌ |
| Countdown timer | ✅ | ❌ |
| PPTX import | ✅ | ❌ |
| Multi-display management | ✅ | ❌ |
| Full app backup / restore | ✅ | ❌ |

---

## ⌨ Keyboard Shortcuts

| Action | Shortcut |
|---|---|
| Toggle Go Live | `Ctrl+L` |
| Send Preview → Live | `Enter` |
| Next / Prev verse | `↓ / ↑` |
| Clear display | `Ctrl+Backspace` |
| Open Projection | `Ctrl+P` |
| Open Settings | `Ctrl+,` |
| Help | `F1` |

---

## 🔨 Build

```bash
npm run build:win    # Windows installer — includes AnchorCast Timer shortcut
npm run build:mac    # macOS DMG
npm run build:linux  # AppImage
npm run build:timer  # Standalone AnchorCast Timer installer
```

> **Before building on Windows:** run `setup_whisper.bat` to create the `python\` folder. Optionally place `vc_redist.x64.exe` in the project root for a fully offline installer. Models in a `models\` folder are bundled automatically; otherwise the installer offers to download them.

---

## 🗂 Project Structure

```
src/
├── main.js                    # Electron main process — IPC, projection, timer, NDI
├── timer-main.js              # Standalone AnchorCast Timer entry point
├── preload.js                 # Electron IPC bridge
└── renderer/
    ├── index.html             # Main operator dashboard
    ├── projection.html        # Projection output window
    ├── countdown-window.html  # Timer control window
    └── js/
        ├── app.js             # Main app logic and operator workflow
        ├── ai-detection.js    # Bible verse detection engine
        ├── bible.js           # Bible database + n-gram indexes
        └── electron-shim.js   # Web-mode compatibility layer

server.js            # Express web server (web mode)
whisper_server.py    # Local Whisper AI transcription server
installer.nsh        # NSIS installer hooks (VC++, models, timer shortcut)
data/kjv.json        # Bundled KJV Bible
ndi-addon/           # Native NDI SDK addon
assets/              # App icons, splash screen, demo GIFs
```

---

## 📋 Requirements

- **Node.js** 18+
- **Electron** 31+ (desktop mode)
- **Python 3.10–3.12** — optional, for local Whisper transcription (bundled in the Windows installer)
- **NDI 6 SDK** — optional, for NDI output

---

## 🌍 Open Source & Contributing

AnchorCast is MIT licensed and open to contributions from developers, designers, church media operators, and testers.

See [CONTRIBUTING.md](CONTRIBUTING.md) for full guidelines and [CHANGELOG.md](CHANGELOG.md) for the full version history.

```bash
git clone https://github.com/gbkeku/anchorcastapp.git
```

---

## ❤️ Mission

To give every church, regardless of size or budget, a powerful, reliable, and free tools for worship and the Word.

---

*Built with love for churches everywhere. ✝*
