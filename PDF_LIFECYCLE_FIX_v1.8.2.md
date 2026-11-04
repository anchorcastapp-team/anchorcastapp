# AnchorCast v1.8.2 — PDF lifecycle correction (macOS and Windows)

## What your macOS output shows

The bundled PDF.js resources, importer connection and existing main-process hook passed. The real Electron conversion failed with `pdf.destroy is not a function`. That is an application cleanup bug in the previous source patch, not a request to install Python. The native smoke test deliberately disallows all external converters, so its fallback Python warning does not diagnose your system Python.

Separately, a process-reload test obtained the correct number of saved slides (`2`) decorated with ANSI colour escapes and compared that output to a plain string. Release/version metadata, an older manifest and missing `beforePack` entries also need attention. A successful unit-test run does not replace those release setup checks.

## Scope

This is an incremental, manual replacement-file patch for the preceding v1.8.2 Bundled PDF Renderer patch. It is not a complete app or an EXE/DMG installer. The app version remains **1.8.2**. The only runtime application file changed is:

```
src/pdf-renderer/render.mjs
```

The archive also includes the refreshed source manifest, corrected/read-only verifiers, focused tests, the two optional presentation UI test files previously absent on your Mac, and current documentation. It contains no replacement `main.js`, preload, package.json, lockfile, build JSON, SMTP configuration, app.js, presentation-import.js, recordings, Bible files or vendor font files. Preserve independent edits before replacing matching files.

## 1. Copy manually

Stop/save current work and fully quit AnchorCast. Back up the matching files. Extract this ZIP separately, then copy its **folder contents** into:

```
/Users/gbkeku/Desktop/Coding/Java/anchorcastapp
```

Merge folders, not entire-folder replacement. In particular, do not replace your whole `src`, `scripts` or `assets` directory. Copy `PRESENTATIONS_PATCH_MANIFEST.json` and the companion scripts, not just render.mjs. Keep the existing presentation-library install block in src/main.js exactly once; do not add it again.

Your already-staged `assets/pdfjs` passed verification. Keep it. No npm install, Python/Poppler setup or voiceover regeneration is needed for this correction. If the actual staged dependency is missing on a different build machine, follow the original PDF renderer staging instructions there.

## 2. Correct release metadata without changing dependencies

In your existing JSON files, set only these application fields:

- `package.json`: top-level `version` = `1.8.2`.
- `package-lock.json`: top-level `version` = `1.8.2`.
- `package-lock.json`: `packages[""].version` = `1.8.2`, when present.

Keep the package name/identity and dependency versions unchanged. Do not globally replace every version number in the lockfile. Keep any newer custom version labels in your renderer files consistent with the intended release.

## 3. Add the PDF build guard to Mac installer configurations

For each Mac build configuration you distribute, add this property at the **top level**, alongside `appId`, `files`, etc. It does not belong inside `mac` or `extraMetadata`:

```json
"beforePack": "scripts/check-pdf-renderer-build.cjs"
```

The configurations reported in your log were:

```
electron-builder-mac-full.json
electron-builder-mac-light.json
electron-builder-mac-update-arm.json
electron-builder-mac-update-x64.json
electron-builder-mac-update.json
```

Keep the existing `files` inclusion `assets/**/*`; it includes `assets/pdfjs`. Do not remove source/Python/audio/NDI/signing rules or an existing `afterPack` hook. The direct `beforePack` value above is what this verifier checks. If you already have a custom `beforePack`, do not discard it: call the PDF guard from that hook and review the composite hook separately; the simple verifier does not execute or certify arbitrary custom build hooks.

For Windows releases, add the same guard to the three reported Windows configurations too. Running the verifier with `--platform=mac` intentionally leaves Windows unchecked; it does not make a future Windows installer verified.

## 4. Verify and test on the Mac

Run in the actual project root, one command at a time. Stop at a failure:

```bash
cd /Users/gbkeku/Desktop/Coding/Java/anchorcastapp
node scripts/verify-presentations.cjs
node scripts/verify-pdf-renderer.cjs --platform=mac
node --test scripts/test-presentation-persistence.cjs scripts/test-bundled-pdf.cjs scripts/test-pdf-renderer-lifecycle.cjs scripts/test-pdf-fix-verifiers.cjs
npx --no-install electron scripts/smoke-pdf-renderer-electron.cjs
```

The final command uses your real installed Electron and PDF.js. It must report successful conversion to two decodable PNGs, reload, deletion and unchanged original PDF. External converters are intentionally disabled; a working Poppler fallback must not hide a broken bundled renderer.

The test scripts sometimes print `Staged PDF.js ...: 9 files` while creating **temporary test fixtures**. Those messages are not a count of the real package in your project; they do not modify your `assets/pdfjs`. Native conversion is assessed by the final Electron test, not by those stub-package fixtures.

If a checksum still differs after copying the refreshed manifest, compare your independent edits with the supplied source instead of blindly replacing unrelated code or manually making the manifest accept unknown content. The manifest now covers current runtime/tooling files and this correction's documents; old release notes/test reports are not runtime prerequisites. An invalid manifest is reported before any stale hash comparisons.

After all checks pass:

```bash
npm start
```

Import your own PDF, close/reopen AnchorCast, reopen its preview, delete the deck, restart, and confirm the deletion persists. Rebuild the normal Apple Silicon/Intel DMGs, then test an installed copy with no external fallback renderer available. A source patch does not update previously distributed installers.

## Why the code changes

PDFDocumentLoadingTask controls the loading/worker lifetime and exposes async destroy(). PDFDocumentProxy is the resolved document and the current API does not expose that method. The new renderer destroys the loading task once and never calls pdf.destroy(). Completion is sent only after all pages have been acknowledged and task cleanup has succeeded. Error cleanup is caught so it cannot overwrite the original failure or throw a second unhandled error.

The existing two-slide reload assertion is preserved. Its child process uses process.stdout.write(String(count)) instead of formatting a number with console.log. The test explicitly enables forced colour to guard against recurrence.

## Validation boundaries

See PDF_LIFECYCLE_FIX_TEST_RESULTS.json for the exact test run. Tests in this environment use Linux, the actual application source, real Poppler/LibreOffice conversion of synthetic fixtures, and a PDF.js API contract fixture for the browser/worker lifecycle. The contract fixture intentionally has no document-level destroy method.

The pinned PDF.js package could not be downloaded here. No real PDF.js decoding, native Electron execution, native Windows/macOS installer or physical-device test is claimed by this patch. The user's local native smoke test is still required. This correction does not address the separately reported Intel Mac Settings loop or earlier audio-duration/path tests.

References: Mozilla's PDFDocumentLoadingTask API documents destroy() at https://mozilla.github.io/pdf.js/api/draft/module-pdfjsLib-PDFDocumentLoadingTask.html and the PDFDocumentProxy API is at https://mozilla.github.io/pdf.js/api/draft/module-pdfjsLib-PDFDocumentProxy.html . These are technical references, not runtime network dependencies.
