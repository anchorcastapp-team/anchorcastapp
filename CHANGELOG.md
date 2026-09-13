# Changelog

All notable changes to **AnchorCast** are documented here.

---

## [1.6.0] — 2026-08

### Fixed
- **"Save handler not registered" error on Bible upload/import** — the `save-bible-version` IPC handler was accidentally dropped during earlier refactoring, breaking both the General tab's manual Upload and the Custom Translation (Advanced) paste-JSON import. Restored.
- **"Clear all" on the AI Detections panel silently wiping the live transcript** — clicking Clear all on the Detections panel was also erasing the entire in-progress sermon transcript with no warning, so anything transcribed before that click never made it into the saved history. The two are now fully independent.
- **Projection window opening on the main screen instead of the external display** — a timing race between positioning the window and transitioning to fullscreen could cause the projection window to land on the wrong display. The window is now positioned and verified on the target display before going fullscreen, with an automatic self-correct if it still lands wrong.

### Added
- **Live transcript autosave** — the in-progress transcript is now periodically saved to a recovery file while recording. If AnchorCast is interrupted mid-service (crash, force-quit, power loss), the transcript is automatically recovered into History on next launch instead of being lost.
- **Per-translation default on import** — importing 21st Century King James (KJ21) now automatically activates it as your default translation and remembers that choice for future launches, the same way manually switching translations already does. (KJ21 is copyrighted and must still be imported manually — see Bible Setup docs.)

---

## [1.5.0] — 2026-07

### Added
- **In-app Bible download** — new "Download Bible" tab in Settings → Bible Versions. Search and download translations directly, no manual scripts or file handling required.
- **License-aware downloads** — translations are labeled Public Domain, Free License, or License Required. Public domain and free license translations download and import with one click; copyrighted translations show a License Required notice and route to manual import instead of downloading.
- **Auto-flatten on upload** — Bible JSON uploads (both the Download tab's License Required flow and the General tab's manual upload) are automatically converted to AnchorCast's format, supporting flat-list and nested books→chapters→verses structures.

### Changed
- Bible Setup help documentation updated to reflect the new in-app download flow and copyright/licensing guidance; outdated command-line script instructions removed in favor of the license-aware in-app system.

---

## [1.4.0] — 2026-05

### Security

This release is a focused security hardening update covering all source layers. All changes address vulnerabilities identified during a full audit of the v1.3.0 codebase across five passes: Electron configuration, IPC handlers, remote control server, Whisper server, renderer JS, renderer HTML, and build/installer configuration.

**Critical fixes**

- **Settings API leaks credentials** — `GET /api/settings` previously returned every stored API key (`claudeApiKey`, `deepgramApiKey`, `geniusApiKey`) and all remote control PINs to any HTTP client on the local network without authentication. Sensitive fields are now stripped for non-loopback callers
- **Settings API unauthenticated write** — `POST /api/settings` accepted full configuration rewrites from any LAN device with no auth. Protected by `requireLocalOrToken` middleware
- **`open-external` arbitrary process launch** — The `open-external` IPC handler passed caller-supplied URLs directly to `shell.openExternal()` with no scheme check, allowing `file://`, `ms-msdt://`, and other dangerous protocol handlers to be triggered from the renderer. Now restricted to `https:`, `http:`, and `mailto:` only. Same fix applied in `electron-shim.js`

**High severity fixes**

- **`webSecurity: false` removed** — Both `mainWindow` and `projectionWindow` had the same-origin policy disabled. This allowed cross-origin fetch from any URL and compounded every other renderer-side finding. Removed from both windows; cross-origin media continues to work via the existing `media://` custom protocol handler
- **`media://` protocol arbitrary file read** — The custom media protocol handler served any absolute filesystem path with no directory restriction. A renderer-side call to `media:///C:/Users/user/.ssh/id_rsa` or `media:///etc/passwd` would succeed and return the file contents. Now restricted to `APPDATA_ROOT` and the app's own asset directories
- **Remote PIN stored in plaintext** — Remote control PINs were stored as plain strings in `~/.anchorcast/settings.json` and compared with direct string equality. Replaced with PBKDF2 (100,000 iterations, SHA-256, per-PIN salt). Legacy plaintext PINs are detected and migrated to hashed form on next save. Comparison uses `crypto.timingSafeEqual` to prevent timing attacks
- **`import-presentation` PowerShell injection** — The handler passed a caller-supplied `filePath` into a PowerShell command string with only single-quote escaping, which is bypassable. Now resolves and validates the path (extension allowlist, existence check, metacharacter block) before any shell interaction; arguments are passed positionally, not via string interpolation

**Medium severity fixes**

- **`/api/control` unauthenticated** — `POST /api/control` and `GET /api/control/pending`, which drive the live projection display, had no authentication. Any device on the same Wi-Fi could project arbitrary content. Both now protected by `requireLocalOrToken`
- **Whisper `/reinforce` unauthenticated** — Any local process could inject arbitrary text into the Whisper context window via `POST /reinforce`, biasing transcription output. The server now requires an `X-Whisper-Secret` header matching a per-run secret generated by the main process at startup and passed to the Python process via `--secret`. Both `fetch` calls in `main.js` include the secret header
- **Whisper memory exhaustion and lock starvation** — `/transcribe` accepted unbounded audio payloads and held a global threading lock indefinitely. Now enforces a 10 MB payload limit and a 30-second lock acquisition timeout, returning `503` if the server is busy
- **Replay timeline XSS** — `ev.payload.ref/title/text/summary` from loaded service archives was injected raw into `innerHTML`. A crafted archive file could execute arbitrary JS on replay open. All payload fields now wrapped in `escapeHtml()`
- **Live transcript XSS** — Whisper interim transcript text was injected directly into `innerHTML` in `updateInterim()`. A transcript chunk containing `<img onerror=...>` would execute. Now escaped
- **`splashWindow` missing preload** — The splash screen window was created without a preload script, leaving no contextIsolation bridge and creating pressure to add `nodeIntegration: true` as a workaround in future. Preload added for consistency
- **No Content-Security-Policy** — No CSP was set on any Express response or Electron session. A CSP middleware is now applied to all Express responses covering `script-src`, `style-src`, `connect-src`, `media-src`, and `img-src`

**Informational fixes**

- **Remote session token in `localStorage`** — The remote control session token was persisted in `localStorage`, meaning it survived browser restarts and persisted on shared tablets. Moved to `sessionStorage` so the token is cleared when the tab is closed
- **Electron version upgraded** — Pinned from `^31.0.0` (EOL) to `^33.0.0`. Electron 31 has reached end-of-life with known Chromium CVEs; 33 is the current stable release
- **`asar` packaging enforced** — All electron-builder configs now explicitly set `"asar": true`, ensuring the app bundle is always packed as a protected asar archive
- **`package-lock.json` version synced** — The lockfile was stale; synced to `1.4.0`. `npm run audit` and `npm run audit:fix` scripts added

**HTML renderer fixes (XSS)**

- `history.html` — Session titles, detected verse references, and action button IDs injected raw into `innerHTML`. Session IDs moved to `data-sid` attributes; all user-derived strings escaped with `escH()`
- `about.html` — Registration name, email, and church name from `registration.json` rendered unescaped. Escaped with `escH()`
- `settings.html` / `bible-manager.html` — `e.message` from caught exceptions and `abbrev` (user-supplied translation abbreviation) injected into status banners. Both escaped
- `presentation-editor.html` — Bible search result `r.ref` and `r.text` rendered directly into `innerHTML`. Escaped
- `projection.html` — `sanitizeProjectionSongHtml()` missed unquoted `onX=value`, backtick-quoted handlers, and `style=url(javascript:...)`. Sanitizer hardened to cover all three quoting styles plus `style=`, `href=`, `src=`, `xlink:href=`, `action=`, and `formaction=`

**Installer (NSIS) fixes**

- **TOCTOU race on temp scripts** — `installer.nsh` wrote PowerShell and Python scripts to predictable paths (`$TEMP\dl_vc.ps1`, `$TEMP\ac_get_model.py`, `$TEMP\ac_run_model.ps1`). Any process running as the same user could replace these between write and execute. All temp paths now use `GetTempFileName` which returns a system-ACL-protected unique path
- **No hash verification on downloaded vc_redist** — The fallback VC++ download was executed immediately without any integrity check. Now verifies SHA-256 with `Get-FileHash` before execution; download is deleted if verification fails

### Fixed

- **macOS black screen on relaunch** — Closing AnchorCast via Cmd+Q and reopening from the Dock showed a blank dark window. The root cause was the `before-quit` handler shutting down the renderer HTTP server but the macOS process remaining alive — so the `activate` handler re-created the main window with no server running, causing `loadURL()` to silently fail. The `activate` handler now restarts the renderer server if `rendererPort` is `0` before creating any windows, and resets `rendererPort` properly on shutdown
- **macOS splash screen hang on Dock relaunch** — Clicking the Dock icon after closing the app caused the splash to appear and get stuck at the progress bar indefinitely. Stale window references and state were not fully reset between close and reactivate. `createSplashAndMain()` now resets all window refs, restarts the renderer server cleanly, and includes a 12-second fallback timeout so the splash never hangs forever
- **macOS "Could not check for updates" HTTP 404** — The Mac update checker was constructing the wrong yml filename by appending an extra `-mac` suffix (e.g. `latest-mac-arm64-mac.yml` instead of `latest-mac-arm64.yml`). Corrected to match the filenames electron-builder actually generates
- **Media import silently failing (images and video)** — Removing `webSecurity: false` (security fix F4) had a side effect: Electron no longer populates `file.path` on `File` objects from drag-and-drop or file dialogs when same-origin policy is enforced. Files appeared to be accepted but nothing was imported. Fixed by using `webUtils.getPathForFile()` — the correct Electron API — exposed via `contextBridge` in `preload.js`
- **PowerPoint import failing on Windows** — The F5 security fix switched PowerShell from `-Command` with string interpolation to `param()` + `-Args`. However `-Args` only passes variables into scripts run with `-File`, not `-Command` — so `$src` and `$outDir` were always empty, causing a COM error. PowerShell script is now written to a `GetTempFileName`-derived temp file and executed with `-File` + positional arguments
- **Remote control Copy URL / Share buttons not working** — `navigator.clipboard.writeText()` requires a secure context (HTTPS or localhost). The remote is accessed from phones on the local network via `http://192.168.x.x` which is neither, so clipboard writes silently failed. Fixed with a 3-method fallback: Electron IPC → `navigator.clipboard` → `textarea` + `execCommand('copy')`, the last of which works on any HTTP origin
- **Remote control showing "Error 429" with no action** — When a stale session token was present, repeated failed auth attempts triggered a rate-limit (429) response. The error was displayed as a raw `"Error 429"` toast with no guidance. All error responses in the remote now show friendly user-facing messages; 401 and 429 both clear the stale token and show the PIN entry screen automatically
- **Deepgram "Invalid API key" false error** — When `whisperSource` was set to `deepgram`, the main process `processChunk()` function had no handler for it — both `useLocal` and `useOnline` were false — so it fell through to the `else` branch which fired `transcript-no-key`. This toast appeared even when Deepgram was transcribing correctly in the renderer. Added `useDeepgram` guard to skip main-process chunk handling entirely when Deepgram is selected
- **Sermon Notes limited to ~5 points** — `max_tokens: 2000` was too small for sermons with many points — the JSON response was silently truncated mid-way and the parser only received however many points fit before the cutoff. Increased to `4096` and added explicit prompt instruction to capture all points without limiting

---

## [1.3.0] — 2026-05

### Added

**Auto-Update**
- App checks for updates automatically every 6 hours and on startup (5-second delay)
- Gold banner appears when a new version is available — user clicks Download to begin
- Green banner when update is ready — "Restart & Install" button applies the update
- Update packages are small (~30–80 MB) — Whisper model never re-downloaded on update
- Separate update channels per platform and architecture (Windows, Mac ARM64, Mac Intel)
- Windows: downloads and installs via `electron-updater` with user consent (no auto-install)
- macOS: checks GitHub releases and shows banner with direct `.dmg` download link

**macOS Support**
- AnchorCast now runs natively on macOS — Apple Silicon (M1/M2/M3) and Intel
- Bundled portable Python 3.12 for offline Whisper AI transcription on Mac
- Whisper model bundled in Full builds; downloaded on first use in Light builds
- Mac-specific PATH restoration — Electron no longer strips Homebrew and user paths
- `after-pack.js` hook restores executable permissions on bundled Python binaries after build
- Traffic light buttons (close/min/max) no longer overlap the AnchorCast title bar
- Cmd+C / Cmd+V / Cmd+X / Cmd+A / Cmd+Z clipboard shortcuts now work in all text fields
- Cmd+P freed for Paste — projection shortcut moved to Cmd+Shift+P

**Remote Control — Sign Out**
- Sign Out button added to the Remote Control UI header
- Clicking Sign Out clears the session token server-side and locally, returns to PIN screen
- `POST /api/signout` endpoint invalidates the session token and clears IP lockout

**Build Variants**
- **Full installer** (~600 MB Windows / ~500 MB Mac) — Python + Whisper model bundled, ready immediately
- **Light installer** (~200 MB Windows / ~150 MB Mac) — Python bundled, model downloaded on first use
- **Update package** (~30–80 MB) — app code + Python, no model — silent auto-update for existing users
- Separate build configs: `electron-builder-win-full.json`, `electron-builder-win-light.json`, `electron-builder-win-update.json`, and Mac equivalents

**Dynamic Versioning**
- Version number is now read from `package.json` in a single place
- About, Settings, Bible Manager, Help, and Welcome pages all update automatically on release
- No more manually updating version strings across multiple files

**Developer Experience**
- DevTools accessible in production via F12 / Ctrl+Shift+I / Cmd+Option+I
- View menu always shows Developer Tools and Reload options
- CSP headers added to Electron renderer server (allows `wss://api.deepgram.com`)

### Fixed

- **HDMI disconnect** — Projection window now parks off-screen silently when projector is unplugged; auto-restores fullscreen and content when reconnected (no user action needed)
- **Win+D minimizing projection** — `minimizable: false` + minimize event intercept prevents Win+D and taskbar from minimizing the projection screen
- **Canvas word-wrap** — Verse text no longer clips at the right edge of Program Preview and Live Display canvases
- **Remote Control auth** — X-Remote-Role header no longer trusted as authentication; sessions properly invalidated on PIN change
- **History window Copy button** — Fixed with 3-method fallback (Electron IPC → navigator.clipboard → execCommand)
- **Remote Control Copy URL / Share buttons** — Same 3-method clipboard fix applied
- **Projection timer z-index** — Timer now renders above song slides (z-index raised from 7 to 50)
- **Media context menu** — Smart viewport positioning prevents menu from rendering off-screen; z-index raised to 9999
- **Whisper not found on Mac** — Fixed Electron stripping PATH on launch, preventing detection of Homebrew Python
- **Wrong userData path on Mac** — Fixed `anchorcast` vs `AnchorCast` case mismatch causing models not to be found
- **Whisper server not killed on app close** — `stopWhisperServer()` now called on `before-quit`; orphaned processes cleaned up via `pkill` on Mac
- **PayPal donation link** — Updated to correct URL across About, Registration Status, and donate modal
- **Deepgram `no_delay` parameter** — Removed invalid parameter that caused HTTP 400 connection failures
- **Deepgram key name mismatch** — Normalized `deepgramApiKey` (server.js) to `deepgramKey` (app.js) in `applySettings()`

### Improvements

**Deepgram Transcription (Live)**
- `utterance_end_ms` reduced from 1200ms to 600ms — final results appear faster
- Audio send buffer batches chunks to 100ms before sending — reduces WebSocket overhead
- Audio worklet flush size increased from 2048 to 4096 samples — fewer, larger sends

**Local Whisper Transcription**
- Audio chunk size reduced from 4.5s to 3.5s — slightly faster transcription turnaround
- Overlap reduced from 0.75s to 0.5s
- Clicking "Start Transcript" before Whisper model has loaded now shows a loading state and auto-starts when ready, instead of showing a false "not installed" error
- `whisperLocalReady` check removed from source toggle — Local can always be selected; Whisper starts on demand
- Whisper setup banner delayed 6 seconds with cancellation — suppresses false "not installed" flash during post-update restart

**AI Detection**
- Navigation buffer suppresses verse :1 false positives when preacher announces a chapter
- Dedup window upgraded to 5 minutes; 60-minute presented suppression (re-allows at ≥97% confidence)
- Sentence loop collapse added to `normalizeTranscriptText` (Deepgram streaming glitch fix)
- Chapter-range fix: "Genesis 5 and 6" no longer fires Genesis 5:6
- Offline Whisper hallucination guard added
- 30+ biblical word corrections from live sermons added to all 3 correction layers

**Projection**
- Guard interval pauses during GPU teardown on HDMI events — prevents white screen freeze
- `mainWindow` guard in `window-all-closed` prevents premature app quit during display changes

---

## [1.2.0] — 2026-04

### Added

**Timer — Single Source of Truth**
- Centralized timer state lives in the main process — the timer no longer depends on the projection window being open to run
- Projection window rebuilds the running timer correctly from `startedAt` + server time when closed and reopened — no more restarting from zero
- Timer is blocked from starting if the projection screen is not open, with a clear actionable error message
- **Flash Faster button** — hold to speed up the red warning blink on both the preview and projection screen
- Timer enters a red warning phase during the final 60 seconds, not only at 00:00
- End message displayed on projection when countdown expires: *"Thank you. Kindly conclude your session."*
- Count-up mode hides the Duration section (not relevant to count-up)

**Standalone AnchorCast Timer**
- `AnchorCast Timer` installs alongside the main app as a separate shortcut (`AnchorCast.exe --timer`)
- Standalone timer opens its own projection window automatically on a secondary display when the timer starts
- Standalone projection supports custom background control (solid color, image, video), clock & date overlay, and Flash Faster
- `npm run build:timer` produces a fully independent standalone installer

**Clock & Date Display**
- Optional clock overlay on the projection screen — 12h or 24h format, show/hide date, 5 position options, custom color
- Clock state persists correctly when the projection window is closed and reopened
- Clock defaults to Top Left position to avoid overlapping the edge timer

**Projection Background Control** *(Standalone Timer only)*
- Set projection background to Default, Solid color, Image, or Video from the timer control window
- Greyed out with a clear message when running inside the main AnchorCast app

**Projection Pipeline**
- Timer rendering interval tightened to 100ms for smoother countdown display
- Coalesced duplicate render-state updates into a single frame — fewer redraws, less lag
- Skipped exact duplicate scripture / song / media redraws
- Direct projection events now routed through shared pipeline — scripture and song changes feel faster

**Whisper AI — Installer & Runtime**
- `vc_redist.x64.exe` bundled in the installer — Visual C++ installs silently with no user action required
- `python\` folder and `faster-whisper` bundled in the installer — AI transcription works immediately after install with no internet required
- `models\` folder bundled if present at build time; installer offers to download the model (~244 MB) if not bundled
- App detects a missing model on startup and shows the correct banner immediately, rather than silently downloading for 1–2 minutes
- "Set Up Now" button uses the bundled Python directly — no console window, no `setup_whisper.bat` confusion
- All HuggingFace Hub warnings (`HF_TOKEN`, symlinks) suppressed from Whisper server output
- `AppData\Roaming\AnchorCast\` path is now consistent between the installer and the app (fixed a `anchorcast` vs `AnchorCast` case mismatch)
- Whisper models stored in `AppData\Roaming\AnchorCast\AnchorCastData\WhisperModels\` — survives reinstalls and app updates

**Transcript & Sermon Notes**
- Transcript auto-saves to history every time recording stops — `💾 Transcript auto-saved` toast confirmation
- Close safeguard dialog when closing the app with an unsaved transcript — options to Generate Notes, Close anyway, or Stay
- Sermon Notes modal includes a Transcript Picker — browse saved transcripts, select one, or use the current live transcript
- Generate Sermon Notes uses whichever transcript is selected (saved or live)

**Settings — Whisper Model Manager**
- Model manager in **Settings → Audio & AI** shows all 6 Whisper models (tiny.en, base.en, small.en, tiny, base, small) with Fast / Balanced / Accurate labels
- ✓ green badge for installed models; ↓ Download button for uninstalled models
- Downloads run in the background with live progress; model selection dropdown updates automatically on completion

### Fixed

- Startup Whisper banner now correctly identifies what is actually missing — Python not installed, faster-whisper not installed, or only the model file missing
- Projection timer restarting from zero after closing and reopening the projection window
- Main-app timer incorrectly showing a failure toast when the timer started successfully but projection was not open
- Flash Faster button not activating during the warning phase — was only working at 00:00
- NSIS `IfFileExists` using `*.*` which does not match folders — model copy silently skipped during install
- Model download dialog not appearing during installation when models were not bundled
- `anchorcast` vs `AnchorCast` AppData path case mismatch causing the app not to find models after install
- VC++ detection now checks 4 registry paths — correctly detected when installed via Visual Studio 2022, Office, or games
- Blank CMD window appearing during model download — now fully hidden via PowerShell `-WindowStyle Hidden`
- HuggingFace Hub warnings (`HF_TOKEN`, symlinks) appearing in setup terminal output
- NDI `.dylib` (macOS) no longer incorrectly bundled in Windows builds

---

## [1.1.0] — Previous

### Added

Full initial platform release including:

- Live transcription — local Whisper AI + cloud Deepgram
- AI verse detection with approval workflow
- Bundled KJV Bible; multi-translation import support
- Live projection with multi-monitor management
- NDI output and MJPEG stream fallback
- Wi-Fi phone remote with PIN-based role access
- Countdown and count-up timer with projection display
- Song Manager with Genius lyrics search and EasyWorship import
- Theme Designer with per-category themes
- Presentation Editor (1920×1080 canvas)
- Service schedule with drag-and-drop ordering
- Context search (AI-assisted topic-to-verse)
- Sermon History with auto-save
- Adaptive Memory — learns vocabulary and accent patterns per speaker
- Verse Detection Review — approve, reject, teach custom trigger phrases
- Sermon Intelligence — AI title suggestions, keyword analysis, structure insights
- Live Smart Suggestions — real-time scripture and song recommendations
- Analytics Dashboard — cross-service historical stats
- Operator Command Center — compact floating live-service panel
- Post-Service Report — auto-generated service summary
- Service Replay Timeline — replayable event timeline with scrubber
- Service Archive — searchable history by title, speaker, verse, or transcript
- Auto Service Builder — history-based service planning suggestions
- Clip Generator — sermon highlight packages for social media
- Offline-first architecture — local AI, local Bible, no cloud required

---

## [Project Notes]

### Architecture
- Electron desktop mode with IPC-based state management
- Web mode via Express + REST/SSE (most features supported)
- Local JSON-based persistence — no database required
- Projection, timer, remote control, and detection systems share centralized state through the main process

### Focus areas since v1.3.0
- Full security audit and hardening across all layers (Electron config, IPC, HTTP server, Whisper server, renderer JS, renderer HTML, build pipeline)
- Electron upgraded to v33 (EOL v31 retired)
- Installer hardened against TOCTOU attacks and unverified downloads
- macOS stability: black screen on relaunch fixed, update check channel names corrected
