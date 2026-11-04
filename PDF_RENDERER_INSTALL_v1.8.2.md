> **PDF lifecycle correction r2:** If the local smoke test reports `pdf.destroy is not a function`, apply the replacement files in the PDF Lifecycle Fix and follow `PDF_LIFECYCLE_FIX_v1.8.2.md`. Do not reinstall Python for that error. Existing staged PDF.js 6.5.438 resources can be retained. On macOS, `node scripts/verify-pdf-renderer.cjs --platform=mac` checks only the Mac build configurations; use `--platform=all` before distributing both families.

# AnchorCast v1.8.2 — bundled PDF renderer (manual source patch)

## Scope and status

Apply on top of the latest Presentations persistence + Python Path Fix source and the v1.8.2 release-documentation update. Version remains **1.8.2**. This ZIP is not an EXE/DMG and is not a complete app. It contains complete replacement/new source files, not an updater.

The new PDF route uses **Mozilla PDF.js 6.5.438, legacy browser build**, staged locally in `assets/pdfjs`. It runs in an isolated Electron window and writes ordered PNG slides into the existing managed presentation store. PDFs no longer need Python, PyMuPDF, Poppler, LibreOffice or PowerPoint when this bundled route is present and functioning. Existing external routes are retained as fallbacks. PPT/PPTX/ODP conversion requirements are different: PDF.js does not convert Office documents; export to PDF or keep the existing Office/Python conversion routes.

The dependency download could not run in the development environment for this patch. **Actual PDF.js decoding and native Electron/Windows/macOS behavior have not been tested here.** Run the included native conversion test on your machine before building and test an installed copy before distributing it. The build guard checks missing resources; it is not a replacement for a real import test.

## 1. Copy the source files

Close AnchorCast after saving any work. Back up your project and presentation data.

Merge this patch into the project root, replacing only matching files. Do not replace entire `src` or `scripts` folders. Preserve independent edits inside `app.js` or `presentation-import.js` before replacing their complete files.

Application files:

```
src/presentation-import.js             replace (retains the Python path correction)
src/presentation-pdf.js                new
src/pdf-renderer/index.html            new
src/pdf-renderer/render.mjs            new
src/pdf-renderer/style.css             new
src/renderer/js/app.js                 replace (cancel result + PDF error title + 1.8.2 About label)
```

The supplied app.js is based on the latest manual Presentations patch and retains the unified newest-first detection timeline. No other detection logic is changed.

Keep the existing `require('./presentation-library').install({...})` hook in `src/main.js` **once**. No extra hook or change to `main.js` is needed. No preload replacement, Python installation, registration change or SMTP change is included.

## 2. Install and stage PDF support — developer/build-machine step

From the actual project folder on Windows or macOS:

```sh
npm install --save-dev --save-exact pdfjs-dist@6.5.438
node scripts/prepare-pdf-renderer.cjs
```

This npm operation needs internet access **on the build machine**. It updates your existing package and lockfile without replacing their dependency graph. Use a Node version supported by the package; do not ignore an unsupported-engine error.

The staging command copies the matching library, worker, resource data and license notices from the installed dependency into the app's existing assets tree. It downloads nothing itself. It replaces only the managed `assets/pdfjs` directory and preserves the previous staged bundle if staging fails. Do not put your own files in that managed directory. Do not remove its resource folders after staging.

The package's optional Node canvas module is not used by this implementation: PDF rendering uses Electron's browser canvas. The PDF engine is a development dependency because the needed runtime resources are explicitly staged into `assets` for packaging. Native converter executables are not needed for the built-in PDF route.

**End users will not run npm, pip, Homebrew or this preparation script.** They receive the staged renderer inside your rebuilt installer and can import PDFs offline.

## 3. Add a packaging guard

For **each desktop installer configuration you distribute**, add this top-level property (alongside `files`, `appId`, etc.):

```json
"beforePack": "scripts/check-pdf-renderer-build.cjs"
```

The previously supplied Windows/macOS configs already include `"assets/**/*"` and `"src/**/*"` in their `files` arrays. Retain those rules. If you use a custom narrower files list, add `"assets/pdfjs/**/*"` and `"src/pdf-renderer/**/*"`, plus the new `src/presentation-pdf.js`, without excluding them later.

Relevant configurations in the supplied project:

```
electron-builder-win-full.json
electron-builder-win-light.json
electron-builder-win-update.json
electron-builder-mac-full.json
electron-builder-mac-light.json
electron-builder-mac-update.json
electron-builder-mac-update-arm.json
electron-builder-mac-update-x64.json
```

Do not change the standalone Timer configuration; this feature belongs to the main app. Do not replace an existing custom `beforePack` hook. Instead call the new check from your existing hook:

```js
await require('./check-pdf-renderer-build.cjs')(context);
```

That relative path assumes your existing hook is also inside `scripts`; adjust the relative path if it is elsewhere. Keep its other work unchanged. The source verifier expects the standard direct hook string; for a composed custom hook, review the integration manually rather than bypassing the build guard.

No `asarUnpack` entry is required for this renderer: it serves packaged assets through Electron's ASAR-aware filesystem APIs. Do not remove your existing unpack rules for Python, FFmpeg, NDI or media.

The beforePack check validates the pinned version, resource hashes and asset inclusion. A missing/incomplete bundle fails the build rather than silently shipping a PDF importer with no renderer.

## 4. Verify and run a real conversion

```sh
node scripts/verify-presentations.cjs
node scripts/verify-pdf-renderer.cjs
node --test scripts/test-presentation-persistence.cjs scripts/test-bundled-pdf.cjs
npx --no-install electron scripts/smoke-pdf-renderer-electron.cjs
```

The last command is important. It uses your **actual installed Electron and the real staged PDF.js**, converts the included two-page PDF, validates PNG decoding and page count, reloads/deletes the saved entry, checks the original file hash and forbids external converter fallbacks. It uses a temporary data directory; it does not touch your church library.

If any step fails, stop before building and keep its complete output. Do not mark an unavailable/failed native conversion check as passed. The ordinary Node tests use mocked PDF.js/Electron where stated and cannot establish that the actual dependency works on your target machine.

After the native check succeeds:

```sh
npm start
```

Test your own PDF. Close/reopen AnchorCast and confirm its pages still load. Delete it and confirm it stays deleted after reopening. Then rebuild your normal Windows/macOS installers. Test a freshly installed copy on a machine without Poppler/PyMuPDF to verify the shipped bundle, not a developer fallback. Retain third-party notices distributed with the dependency.

## Runtime behavior

- Built-in PDF renderer first; external Poppler/PyMuPDF remain fallback routes.
- Local progress window and Cancel. Closing this import window also cancels that import.
- One PDF job at a time; 128 MB source size, 500 pages, 1920-pixel longest side, 1 GB total PNG output, 90-second idle timeout, 15-minute total timeout.
- PDF scripts are not enabled. The isolated window has Node integration off, context isolation and sandboxing on, restricted local resource loading and no external navigation/download/permission access.
- Password-protected files request an authorized unlocked copy; this patch does not implement a password-entry form. Pure XFA documents request a flattened PDF.
- Invalid, cancelled or incomplete imports do not commit a library record. Partial generated pages are removed; the original source is unchanged.
- Converted PDF pages are static images, not editable PowerPoint objects or animations.
- Missing bundled resources produce a specific error in the app. Source-development instructions are distinguished from the instruction for installed-app users to install a corrected release.

## Unchanged / excluded

The main.js hook, preloads, SMTP settings, transcription providers, registration restrictions, timers, sermon audio recorder/voiceovers/trimming, Bible data and existing library files are not replaced. This patch does not claim to fix the Intel Mac Settings loop or the earlier macOS audio-duration tests.

## References

- PDF.js getting started: https://mozilla.github.io/pdf.js/getting_started/
- PDF.js supported environments / legacy builds: https://github.com/mozilla/pdf.js/wiki/Frequently-Asked-Questions
- PDF.js API options: https://mozilla.github.io/pdf.js/api/draft/module-pdfjsLib.html
- Electron BrowserWindow: https://www.electronjs.org/docs/latest/api/browser-window
- Electron security: https://www.electronjs.org/docs/latest/tutorial/security

## Rollback

Close AnchorCast and restore the backed-up `presentation-import.js`, `app.js`, documents and manifests. Restore the previous build hooks and package/lockfile from your backup if rolling back the dependency addition. The managed `assets/pdfjs` folder can then be removed. Do not delete the application's saved presentations or user-data folder.
