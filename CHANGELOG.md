# Changelog

All notable changes to **AnchorCast** are documented here.

## [1.8.2] — Unreleased

The presentation-library and import fixes previously prepared under the v1.8.1 draft are assigned to **v1.8.2**. **v1.8.0 and v1.8.1 have already been released**; their history is retained below. Release dates have not been added where none were supplied.

### Fixed — bundled PDF renderer compatibility (r2)

- Corrected PDF.js cleanup to use `PDFDocumentLoadingTask.destroy()` instead of calling the removed `destroy()` method on a `PDFDocumentProxy`. This resolves the `pdf.destroy is not a function` failure reported by the macOS native smoke test. Cleanup runs once across success, failure, cancellation and password-request paths; a cleanup error does not overwrite the primary rendering error.
- Fixed the separate-process presentation reload test to write an unformatted numeric result to stdout. The assertion still requires exactly two persisted slides; it now passes even with terminal colours forced on.
- Refreshed the presentation/PDF verification manifest and included the previously missing optional presentation UI test tools. An outdated manifest is rejected before comparing its stale file hashes.
- Added `--platform=mac`, `--platform=win` and `--platform=all` to the read-only PDF build verifier. All matching build variants on the selected platform still require assets and the beforePack guard. Selecting one platform does not verify the other platform's installer.
- Added renderer lifecycle contract tests whose document proxy deliberately lacks `destroy()`. Actual PDF.js decoding and rebuilt native installers still require the real local smoke test; no native compatibility certification is claimed.

### Added — bundled PDF import

- Added a local PDF.js rendering route for Windows, macOS and Linux desktop builds, ahead of the existing Poppler/PyMuPDF fallbacks. The renderer resources are staged under `assets/pdfjs` for inclusion in Full, Light and Update installers; PDF imports no longer depend on separately installed Python packages when the bundle is present.
- Added a local import-progress window with Cancel, sequential page rendering, bounded size/page/output limits and cleanup of incomplete output. Closing that window cancels the import without changing the original PDF or committing a partial library record.
- Added specific password-protected, malformed, unsupported-XFA, missing-bundle and timeout messages. PDF failures no longer instruct operators to install PowerPoint or LibreOffice for an existing PDF.
- Added developer staging, a `beforePack` resource check, read-only source/dependency verification and a native Electron conversion smoke test. PDF.js is pinned to 6.5.438 and staged with its matching worker/resources/notices. End users receive the renderer in rebuilt installers rather than installing npm/Python tools themselves.
- Rendering uses an isolated, sandboxed browser window with no Node integration or arbitrary filesystem/IPC bridge; resources and job endpoints are limited to local allowlisted paths. PDF scripting is not enabled.
- Validation scope: source/bridge/packaging tests run locally, but the real PDF.js dependency could not be downloaded in the patch environment. Actual PDF.js decoding and native Windows/macOS installer behavior must be checked on the build machine before release. This entry describes the implemented source path, not a completed native certification.

### Fixed — presentation library, saved previews and import diagnostics

- Cross-platform interpreter discovery now builds paths with the explicitly requested platform's path rules. Fixed the Mac candidate assertion failing when the suite runs on Windows (`/Resources/...` becoming `\Resources\...`); added Windows/POSIX host simulations without relaxing assertions. Mac per-user LibreOffice paths follow the same rule; Windows `py.exe` launcher paths also receive `-3`.
- Imported PowerPoint/PDF records now retain an explicit imported type. Legacy imported entries are normalized on load, so restarting does not send them to the editable-slide path or the wrong deletion store.
- Persist deletions before removing cards from the interface. Both imported and custom decks use consistent string IDs; failed writes leave the card visible and report an error. Remove corresponding entries from the current in-memory schedule.
- Atomic library replacement with an adjacent previous-content `.bak` file; malformed saved JSON is reported, not silently replaced with an empty library.
- Editor Save awaits acknowledged disk persistence before notifying the main window and closing. Targeted upserts avoid rewriting an unrelated stale list.
- Editor images/backgrounds are copied into persistent managed storage through the existing media bridge. Removed temporary blob URLs and direct `file://` previews for new image imports. Already-expired blob assets must be reselected; this update cannot reconstruct them.
- Normalize older title/body slide layouts and generated text style aliases. Render only after the preview container is visible, remove stale imported-image elements when switching to custom slides, and include image objects in custom thumbnails.
- Imported slides use their actual saved filenames, including padded and uppercase names. Repair managed paths when the existing slide folder can be found; retain missing-asset entries with an explicit re-import message.
- Late async slide fetches cannot overwrite a newer selection. Keep a persistent explanatory import-error dialog instead of an easily missed toast.
- Move **New** out of each presentation row into the library toolbar beside **Import** and **Schedule**.
- Presentation import checks bundled Python plus system interpreters, including common macOS GUI-launch paths and Windows `py -3`. Missing interpreters, missing Python packages, missing conversion tools, invalid documents and helper failures are no longer all reported as “Python not found.”
- Discover the macOS LibreOffice application and common PDF-tool locations. Stage the Python helper outside Electron's virtual ASAR archive. The optional basic PPTX fallback reports its layout limitations; no dependencies are downloaded automatically.

### Release scope

- Retains the already-released v1.8.1 global newest-first timeline for confirmed detections and Suggestions / Review. Detection thresholds, grouping, History, Auto/Manual behavior, recording, voiceovers, trimming, timers and registration remain unchanged by this maintenance patch.
- Includes the cross-platform Python discovery correction together with presentation persistence, restored previews, durable deletion and the relocated **New** toolbar action.
- The release is numbered **1.8.2**. Source builds must use matching application-version metadata; the documentation update alone does not change `package.json`, lockfiles or displayed application labels.
- PDF.js is a new build dependency; its staged browser assets are included in the installer. No new preload permission is introduced. The existing presentation-library hook in `src/main.js` must remain installed once; do not add a second copy.
- The separate Intel Mac Settings-window loop and earlier macOS audio-duration/path test failures are not claimed as fixed by this release.

---

## [1.8.1] — Released

### Fixed — global detection panel ordering (corrective revision r2)

- Confirmed references, quotations and **Suggestions / Review** now share **one newest-first timeline**. Fixed the first v1.8.1 update's remaining layout error: correctly sorted suggestions were still pinned below all confirmed cards by separate rendering containers.
- Order uses the latest accepted detection or meaningful evidence update across all categories. A suggestion can appear above older confirmed cards; a later confirmed update can move ahead of that suggestion.
- Each suggestion retains an amber indicator and an explicit **Suggestion / Review — not a confirmed citation** label. Removed the separate bottom-pinned Suggestions wrapper, not the underlying suggestions.
- Meaningful supporting evidence updates its existing card. Identical repeated results, routine redraws and presentation-state changes do not fabricate activity or duplicate events. Same-millisecond updates use the existing stable sequence.
- Promoted suggestions keep one card in the shared timeline with updated confirmation and correct chronology.
- Passage expansion, keyboard focus and individual Present/Schedule/Review actions survive reordering. Scrolled operators are not forcibly returned to the top by new arrivals or supporting updates.
- Display activity remains separate from spoken evidence, Auto freshness and projection authorization. Showing a suggestion higher does not grant Auto permission or allow stale speech to replace a newer passage.

### Changed

- Retained v1.8.1 `updatedAt`, `activitySequence` and grouped activity timestamps; legacy records keep their creation/source-time fallback.
- Debug export reports `interfaceRevision: detection-global-order-v1.8.1-r2`, `panelLayout: unified-timeline`, and **actual rendered card order** in addition to model order.
- Application metadata remains **1.8.1**. This correction includes a complete manually replaceable `app.js`, read-only verification and regression checks. No new runtime dependency or installation script.

This maintenance update retains v1.8.0 confidence benchmarks, grouping, History, Auto/Manual controls, MP3 recording, voiceovers, trimming, timers and registration behavior. It does not address the separate Intel Mac Settings-window loop or macOS audio-duration/path test failures.

---

## [1.8.0] — Released

This release combines the changes prepared for the **unpublished v1.7.2 draft** with the complete **v1.8.0 Sermon Audio Studio** updates. There is no separate v1.7.2 release entry; those changes are included below as part of v1.8.0. The v1.7.1 and earlier history is preserved unchanged.

### Added

#### Independent sermon audio recording

- **Local MP3 recorder** — an opt-in **Record Audio** control below **Start Transcript** records the microphone selected in **Settings → Audio & AI**. Includes an input meter, elapsed capture time, **Start Recording**, **Pause/Resume**, **Stop & Save**, **Save MP3 As…** and **Open Recordings**. Recording remains independent of transcription and does not start when the controls are merely opened.
- **Streaming and recovery** — microphone audio is encoded locally as a genuine **128 kbps mono MP3**, with bounded PCM buffers, input/session validation, interrupted-file retention where possible and an app-close recording guard. Paused time is excluded. No additional transcription-provider call is introduced, and the recorder does not require a cloud API or Whisper model. Check interrupted audio before sharing it.

#### Recording settings and reusable voiceovers

- **Recording tab under Audio & AI** — editable introduction and closing scripts, optional contact fields/placeholders, a default formal male Onyx narration preset, voice/speed selection, editable voice direction, local preview and optional OpenAI speech generation. Generation uses the saved OpenAI API key, requires explicit consent and may incur API charges. Only the selected script, resolved contacts and voice direction are sent; this feature does not upload the sermon recording.
- **Prepared voiceover library** — independent introduction and closing assets saved locally for offline reuse. Contact/text/voice changes invalidate affected prepared speech; music changes require local **Rebuild mix** only. Authorized narration can also be imported without cloud generation. Unresolved placeholders block preparation rather than being spoken.
- **Introduction and closing background music** — includes the original synthesized **Warm Welcome** instrumental, a music-off option and custom-track import, with adjustable volume, lead-in, ending and transition gap. Fades are applied around the narration. Music accompanies the introduction and closing only, not the preacher's sermon. The asset notice and reproducible synthesis source are included.
- **Generic editable samples** — introduction and closing starter text is not tied to a particular church and does not require contact placeholders. **Use sample text** is available for each slot, with confirmation before replacing nonempty text.

#### Sermon trimming and finishing

- **Save the original first** — explicit **Stop & Save** and **Stop & Save, then Close** save the original MP3 before offering optional finishing. Existing recordings can be opened through **Trim / Add Intro / Closing…**. Introduction and closing inclusion is opt-in; both checkboxes start unchecked.
- **Start/end trim editor** — choose the sermon section to keep with **Start at / End at**, two sliders or **Set start/end to playhead**. Shows original and kept duration, with **Play kept section** and **Reset trim**. This trims the beginning and/or end, not middle sections or multiple tracks.
- **Non-destructive, sample-boundary export** — trim only the sermon input, then concatenate the unchanged prepared introduction and/or closing. Export a trimmed sermon alone, either voiceover, or both. Finished **128 kbps mono MP3** copies go to **Recordings/Exports**, with preview, **Save finished MP3 As…** and **Open finished messages**. The original and previously finished copies are not overwritten.
- **Efficient original-audio preview** — decode duration locally with the existing encoder, cache it only while the source file is unchanged and stream preview through revocable loopback-only capability URLs with byte-range support. The whole sermon is not copied into renderer memory.
- **Explicit Cancel and Escape** — Cancel remains reachable in the finishing dialog footer. Before export it closes the editor; while that dialog is loading or exporting it cancels its own operation, waits for cleanup and then closes. Cancelling after a completed save keeps both the finished MP3 and the original. Unrelated audio jobs are not cancelled.
- **Validation and failure handling** — block invalid, empty or reversed selections; reject source changes since preview; clean unfinished derivatives after cancellation/errors; and retain trim choices for retry after a failed export. Timestamps accept seconds, MM:SS or HH:MM:SS with millisecond precision. Native shutdown/interruption paths retain their original-audio behavior rather than opening an optional finishing prompt.
- **Native audio-generation/export protection** — allowlisted local file operations, bounded imports, a fixed HTTPS speech endpoint, redacted provider errors, scoped job cancellation and exit confirmation while preparing/exporting. Prior prepared clips and original sermons survive failed or cancelled operations.

#### Diagnostics and source tooling

- **Readiness labels** — distinguish waiting for reference completion, suggestions for review, ready in Manual, ready/queued for Auto, sent-to-output and held/cancelled states. Confirmation colour is separate from age and presentation readiness.
- **Expanded and grouped diagnostics** — export effective threshold, Manual/Auto state, passage context, confirmation decisions, evidence upgrades, parser-error stage, Auto holds/cancellations and projection IPC results. Include discussion IDs, announced-passage membership, contributing methods, grouped-card counts and separate occurrence counts. IPC acceptance is not proof of physical-screen display. Logs can include private transcript excerpts; review them before sharing.
- **Setup and verification tooling** — checked, repeatable merge scripts for native bridges, encoder packaging and version metadata; source manifests; Node regression tests; optional Chromium checks; and rebuild checklists. Native-file backups are kept outside the project. Incremental trim setup merges its required preview and cancellation methods into both preloads without replacing unrelated code or SMTP settings.

### Changed

#### Confidence benchmarks and detection canvas

- **One selected benchmark** — **Low 60%, Medium 75%, High 85%** remain selectable. The saved threshold is inclusive and applies before grouping and again before Auto commit. Removed the hidden 80% online floor, 88% Auto gate and conflicting legacy per-type overrides from these decisions. Missing/invalid settings default to 75%; an existing valid selection is retained.
- **Confirmation is separate from confidence** — a match needs a confirmed reference, supported quotation or operator-approved learned phrase for Auto; a high score alone does not confirm a Scripture. Scores are heuristic match scores, not measured accuracy guarantees.
- **One card per ongoing Scripture discussion** — an announced reference, its following quotation and continued supporting phrases share a discussion identity. This addresses the repeated active cards reported for **1 John 1:9** and **Philippians 1:6**. Each accepted occurrence remains independently recorded and reviewable; scores are neither averaged nor artificially boosted.
- **Genuine repetitions retained** — separately re-announced references, returns after another confirmed passage or a longer gap can begin another discussion. There is no whole-service ban on detecting the same verse again. Existing saved sessions and historical detection lists are not rewritten.
- **Suggestions separated from confirmed detections** — unrelated thematic matches appear in an initially collapsed **Suggestions / Review** section. Supporting suggestions for the active verse remain attached without downgrading its confirmation or transferring their score/Auto eligibility to it.
- **Cleaner detection canvas** — removed recorded-match counts, evidence rows, timestamps and per-match controls from active cards. Detailed records remain available through Review, History and diagnostic export. Expandable announced passages remain available.
- **Stable confirmation colours** — confirmed references/quotations meeting the selected benchmark remain green even when they become earlier occurrences. Suggestions remain amber; rejected, superseded or below-threshold results are not shown as confirmed green. Stale Auto presentation remains blocked independently, with age/readiness explained in text.

#### Recording status, narration and storage

- **Visible recording indicator** — the main Audio button and recorder monitor show a pulsing red dot and **REC HH:MM:SS** while recording. Minimized recorder controls do not hide the main-button indicator or elapsed time; the timer text stays steady.
- **Distinct non-recording states** — paused audio shows a steady amber pause marker. Startup, saving, stopped and error states do not show a false recording pulse. Screen-reader announcements occur on state changes; reduced-motion/high-contrast preferences use a steady marker.
- **Script-only voiceovers** — introduction and closing speech requests contain only the resolved editable script. Removed the automatically added AI-narration sentence and updated the default voice direction.
- **Written listener notice** — Settings and the finishing dialog provide a non-spoken publishing reminder for selected AI-generated clips. Imported-only narration and original sermon audio are not falsely labelled as generated. Include an appropriate disclosure in the episode description or a visible listener-facing player notice when sharing.
- **Shared bundled audio folder** — Warm Welcome and its notice use the existing root **assets/audio** directory rather than a separate **src/audio-assets** folder. Runtime lookup, synthesis scripts, packaging and verification use this location; existing **assets/Audio** casing is supported. Prepared voiceovers, recordings, exports and custom tracks keep their writable user-data paths. The folder move does not require regenerating narration.
- **Speech-provider requirements** — this implementation supports OpenAI speech generation or imported narration, not a Claude-key voice provider. Recording settings explain the requirement; recognizable Anthropic credentials entered in the OpenAI field are rejected locally before a speech request. No Claude script-writing feature is introduced.

### Fixed

#### Scripture parsing and passage context

- **Premature references in both modes** — chapter-only mentions establish context without becoming verse 1. Unfinished verse numbers and ambiguous number sequences are held before creating a verse card. Complete references have a short settling period for nearby corrections.
- **Spoken-digit references** — marked chapter/verse numbers support phrases such as “Psalm one one nine, verse one zero five.” Ordinary “Romans one one” remains Romans 1:1 and numeric hyphens remain ranges. Added singular “Hebrew,” chapter-before-book phrasing and book/number forms without spaces.
- **Field-aware chapter/verse corrections** — “Psalm 89, I mean verse 89” can restore the earlier Psalm 119 context rather than changing the active chapter to 89. A later correction marks the old occurrence superseded instead of deleting History.
- **Side references hijacking the main reading** — track recent explicit passages and supported quotation context separately. Ambiguous bare verse numbers wait for explicit book/chapter context or corroborating installed Bible text rather than inheriting the last side reference automatically. “Read on” alone does not invent the next verse; ambiguous shared quotations remain held outside the active reading.
- **Stronger evidence discarded** — a stronger overlapping match updates its existing occurrence and can become eligible for Auto without another card. Rejected/dismissed cards and cleared panels are not silently reactivated, and operator-corrected references are not overwritten by late evidence.

#### Announced passages and automatic presentation

- **Verse ranges stay together** — **Daniel 3:16–27** is one expandable passage with all twelve verses in ascending order, **Present verse 16**, individual verse presentation and **Schedule passage**. Non-contiguous lists remain lists; unannounced intermediate verses are not invented.
- **Auto range queue overflow** — only the opening verse of an announcement enters the Auto queue. Unread members no longer fill the eight pending slots or discard the opening verses. Auto waits for subsequent explicit verse references or supported quotations instead of racing through every verse. The existing opt-in complete-passage display remains available for contiguous ranges.
- **Unread range transcription hints** — send only the opening verse to the local transcription hint bridge initially, rather than letting every unread tail verse replace it.
- **Reading position after a range** — “Next verse” after announcing 16–27 advances from 16 to 17, not to 28. Supported subsequent quotations within the range can update the reading position.
- **Repeated Auto presentation during one discussion** — Auto uses discussion and verse identity, not just a method's event key. Supporting phrases do not repeatedly send the same ongoing verse discussion to live output; genuinely new discussions remain eligible.
- **Manual/Auto synchronization** — applying Manual settings disables Auto; toolbar mode changes persist. An unchanged mode control in an already-open Options window does not overwrite a newer toolbar choice. Manual selection, panel/transcript clearing, rejected matches and manual output actions cancel pending Auto work.
- **Rapid or stale automatic overwrites** — scheduling briefly settles candidates, prioritizes current evidence and spaces live changes. Before committing, it rechecks mode, live state, threshold, status and freshness. Late responses cannot override newer manual selections. Enabling Auto can reconsider the newest fresh eligible result without replaying an old backlog.
- **Detection card flicker and focus loss** — readiness/supporting-evidence updates do not replay every card's entry animation or unnecessarily reset scroll position. Expanded passages and focused controls survive redraws.

#### Online response handling

- **JSON parsing failures** — read all provider text blocks, accept one complete array (including a valid empty array with harmless surrounding commentary), validate item schema/evidence and reject truncated or ambiguous responses without inventing results. Diagnostics distinguish response decoding, schema/evidence issues, output truncation and HTTP/network failures. Parsing failures use a short retry cadence rather than the old minute-long network backoff; local detection continues independently.

#### Recorder controls and theme styling

- **Explicit Minimize (−)** — hides recorder controls while capture continues; a paused session stays paused. The main Audio button restores the same session and current time.
- **Guarded Close (×)** — asks before stopping active or paused recording. **Stop & Save, then Close** waits for successful finalization and offers optional finishing. **Cancel** leaves the recording state unchanged. Idle/saved controls close normally; failed or interrupted saves remain visible.
- **Safe Escape and cancellation** — Escape uses the close-confirmation path while recording and dismisses that question when already open. Backdrop clicks do nothing. Startup cancellation waits for cleanup and duplicate save clicks are blocked.
- **Recorder theme override** — state-specific colours take precedence over the application theme and hover styles. Frequent audio updates do not recreate the animated dot.

### Upgrade and compatibility

- **Single consolidated release** — application and lockfile metadata remain **1.8.0**, with the unpublished v1.7.2 work incorporated into this release. Existing application identity and data paths are unchanged; there is no requirement to publish or install a separate v1.7.2 release first.
- **Existing features retained** — v1.7.1 negative overtime/red blinking, **Hide on Projection / Show on Projection** while timing continues, selected-translation loading, independent detection History/recovery, Theme Designer behavior and existing registration restrictions. AI Transcription, additional Bible translations, NDI and Remote Control retain their registration requirements; basic audio recording remains separate.
- **Saved scripts protected** — a one-time Recording-settings update replaces only untouched former sample defaults, allowing formatting-only differences. Custom scripts, contacts, imported narration, music and voice selections are retained. A byte-for-byte backup is created beside the settings file before the atomic revision update.
- **Old generated prefaces require regeneration** — previously prepared AI clips containing the added spoken sentence are kept but marked **Generate voice again — remove old spoken preface**. They are not silently reused as current clips. **Generate voice**, not **Rebuild mix**, removes the preface from future narration. Imported audio and earlier finished sermon files are unchanged; no paid generation happens during migration.
- **Encoder dependency** — the combined release adds **ffmpeg-static** where absent. Install dependencies on the appropriate build machine, synchronize the lockfile before using **npm ci**, and include the matching encoder in each target installer. These source packages do not supply platform-specific encoder binaries. Review the included third-party notices and platform/signing notes.
- **Data and configuration** — no automatic rewrite of existing transcripts, historical detections or Bible data. Native merge scripts preserve unrelated main-process/preload configuration, including local SMTP settings. Existing API-key storage is not changed or newly claimed to be encrypted.
- **Validation scope** — implementation results and limitations are documented in **START_HERE_v1.8.0.md**, **RECORDING_TRIM_TEST_RESULTS.json** and the earlier feature test notes. This release-document consolidation changes no application code and is not a new native-installation or microphone test.

---

## [1.7.1] — 2026-10

### Added
- **Hide on Projection / Show on Projection** — hide the running timer, its label and conclusion message without stopping, pausing or restarting it. Countdown, Count Up and negative overtime continue in the background. The Timer Monitor and quick-controls readout remain available to the operator. The visibility setting stays with the current timer when projection/control windows reopen; a new Start defaults to visible. Available in the timer window (including standalone Timer), inline controls and timer modal. Slides, media, clock and other overlays are unaffected.
- **Negative overtime countdown** — at 00:00 the timer remains running and continues as -00:01, -00:02, and so on, showing how far the speaker has exceeded their allotted time. Red blinking, the conclusion message and Flash Faster remain available. Applies to the timer control window, operator live preview, main-app projection and standalone Timer.
- **Detection Debug log** — exports processed-line decisions, actual thresholds, missing Bible text, selected/fallback translation and panel-versus-history counts for troubleshooting. API settings and tokens are not exported; transcript excerpts are included and should be reviewed before sharing.
- **Detection verification and offline replay tools** — includes source verification, parser/history regression tests, optional Chromium component checks, and a finalized-text replay command using a user-supplied flat Bible file. No new runtime dependencies.

### Fixed
- **Overtime lost on window reopen** — expired countdowns are no longer discarded by the main/standalone timer state handlers. Opening the timer controls or reopening the projection during the same app session reconstructs the original deadline and current negative overtime rather than hiding or restarting the timer. Count-up timers also preserve their elapsed time on reopen.
- **Operator timer disappearing at zero** — the live timer overlay stays visible and blinks red throughout overtime unless the operator hides it, stops it or starts a new timer. A hidden timer continues in the operator monitor. Rejected start requests no longer show a false running timer.
- **Inconsistent countdown boundary and formatting** — a shared timing helper keeps the displays on the same rounding rules, uses the original duration for warning thresholds, and formats the negative sign without a -00:00 display. Styling changes do not replace the current preview with the configured starting duration. The conclusion message no longer overwrites the editable session title.
- **Detected verses missing from saved History after clearing the panel** — History and crash recovery now use an independent current-transcript detection ledger, not the dismissible AI Detections list. Clearing the panel preserves previous references and subsequent appearances of the same Scripture; clearing the transcript explicitly starts a new ledger. Failed permanent saves retain the recovery file.
- **Later repeated Scripture references suppressed** — detection identity now follows the contributing finalized transcript occurrence rather than suppressing identical reference/text pairs for the whole service. Repeated citations and keyword/learned quotations remain eligible; overlapping evidence from the same occurrence is coalesced.
- **Spoken and split Scripture references** — adds a shared grammar for complete spoken numbers, ordinal book names, adjacent verse continuations, lists/ranges, chapter-only context and bounded self-corrections. Examples include “Ephesians 3:20” followed by “And 21,” “Ephesians chapter one. Verse eighteen and nineteen,” and “Romans eight seven, I mean eleven.” Does not invent verse 1 for chapter-only mentions or silently change valid numeric references to match a quotation.
- **Selected Bible translation ignored by detection** — repairs the missing reload function and incorrect translation-availability checks; detection now uses the selected installed translation. KJV fallback is explicitly identified, unavailable verse text is diagnosed, database readiness is reported per translation, and replacing same-sized Bible data invalidates its search index.
- **Quote matching delayed during continuous speech / cloud failure** — local detection remains active with or without optional online detection, has a bounded pending-search wait, and uses source-attributed conservative quotation matching. A keyword index reduces full-Bible rescans. Pending asynchronous results are bounded at save and invalidated across transcript resets.
- **Compound Psalm/verse numbers changed by transcript normalization** — complete number phrases such as “Psalm twenty three” and “chapter eight verse twenty six” remain intact instead of partially converting the number.

### Changed
- Application version and root lockfile metadata updated to **1.7.1**; installer naming continues to use the package version. Help, changelog and release notes reflect this release.
- Shared timer helper is included in the standalone Timer packaging allowlist. No runtime dependencies added.
- Added timer regression checks and refreshed the existing source verification hashes for this combined release.

---

## [1.7.0] — 2026-09

### Fixed
- **Theme Editor background image/video not displaying** — the `media://` protocol used to serve local background files only allows files inside AnchorCast's own data folder; any picture or video picked from outside it (Downloads, Desktop, etc.) was silently rejected with a 403, so the background never rendered — even in the Theme Editor's own preview. Background files are now automatically copied into AnchorCast's data folder on pick/drop before being referenced, for both the file picker and drag-and-drop.
- **Remote Control and NDI registration locks** — both controls are greyed out for unregistered devices and explain the restriction with **Register now** / **Not now** rather than silently toggling or doing nothing. Covers the toolbar, Remote popover, Settings and External Output panel. Registration status is rechecked before enabling; activation refreshes controls without automatically starting a service.
- **Remote Control false ON state** — the UI now waits for successful server binding and checks blocked/error responses before changing the switch. Failed or repeated requests no longer misleadingly show Remote Control as active.
- **Network feature startup bypass** — registration is checked inside the desktop Remote/NDI service start functions, including startup, saved-settings and network-adapter restart paths. Unregistered settings cannot persist enable flags through the settings-save flow; unavailable registration status denies access. Remote URLs/role links and NDI stream URLs are not returned while locked.
- **Register from Settings** — the registration form can now be opened from the trusted Settings window; sender validation is retained. The registration prompt supports mouse/keyboard operation, focus return and Escape dismissal.
- **AI Transcription still available after skipping registration** — Start Transcript now checks verified registration before microphone access and offers **Register now** / **Not now**. The same check covers Local Whisper, Deepgram Live, cloud transcription, audio-file import, and legacy/manual transcript start paths. PCM submission and processing now also enforce registration in the main process; a running model or configured API key does not grant access.
- **Repeated/cancelled transcript starts** — concurrent Start clicks share an in-flight registration check; stopping during microphone permission or a pending Deepgram connection prevents a late start. A registration-status failure blocks transcription with an actionable message instead of silently allowing it. Existing transcript contents are not cleared.
- **Registration failure blocking first launch** — dismissing the registration form no longer quits AnchorCast. A persistent **Continue without registering** button works before, during and after an email request. It reveals the basic app without activating the device or removing registration-required feature restrictions.
- **Invisible main window after activation** — successful email activation, registration dismissal and normal startup share a per-window reveal path that restores opacity and sends `app-ready` once. Late startup callbacks cannot affect replacement windows. Update initialization is not duplicated on subsequent reveals.
- **Registration email error handling** — the existing 20-second socket idle timeout remains; the TLS socket now has its own idle timeout, unexpected disconnects settle immediately, and a separate 45-second total request deadline prevents unbounded waits. The form supports retry and shows a clear failure message. Concurrent email-send requests are rejected while one is in progress.
- **False email failure after SMTP acceptance** — success is recognized when the SMTP server accepts the message after DATA; it no longer depends on the later QUIT response. Server acceptance is not proof of delivery to an inbox.
- **Help → Registration → Open registration form** — the button now opens the form instead of only closing the status window.
- **Quit during registration** — deliberate Quit/Cmd+Q is not treated as Skip. Cancelling the existing unsaved-transcript quit prompt no longer stops the application's servers prematurely.

### Added
- Reproducible registration regression tests and patch verification script.

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
