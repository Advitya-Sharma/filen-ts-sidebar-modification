# @filen/office

Document engines and editors shared by the Filen apps. `pnpm run lint`, `pnpm run typecheck` and `pnpm test` are what CI runs.

## DOCX viewer

- `@filen/office/dom/docx` runs the engine (package → model → styles → fonts → layout) and streams pages as plain data: `serveDocxWorker` + `workerDocxHost` over a worker (web, one worker per document), or `inThreadDocxHost` with `windowServices` on the page's own thread (Expo DOM components). The document's bytes stay with the caller; the engine reads ranges.
- It also holds the viewer's controls as serialisable commands and state (`DocxViewerCommand`, `DocxViewerState`, `DOCX_VIEWER_COMMANDS`): apps draw their own toolbars, a native toolbar can send them over the DOM-component bridge (`parseDocxViewerCommand`).
- `@filen/office/react/docx` is unstyled: `useDocxDocument`, `<DocxPages>` (virtualised, real DOM text), `useDocxViewer`, find and link helpers. `@filen/office/css/docx.css` is structural; every visual value is a `--fo-*` variable.
- Apps serve `assets/fonts/**` (with each family's licence file) from their own origin and pass its URL to `workerServices`; fonts are registered from bytes, so `font-src 'self'` holds.
- Apps that bundle the React layer from source dedupe `react` and `react-dom` (Vite `resolve.dedupe`).

## DOCX editing core

- `@filen/office/core/docx/edit` turns a story into a ProseMirror document whose nodes and marks keep their source items, so a writer can copy whatever was not edited as its original bytes. Blocks that cannot be edited safely become locked blocks. Invariant plugins keep markers, section breaks, fields and source identity intact, and `DocxEditorSession` is the headless command/state API (`DOCX_EDITOR_COMMANDS`, `parseDocxEditorCommand`, `snapshot()`).
- `@filen/office/core/docx/spelling` is the spelling of an editing session, headless: `DocxSpelling` gives the misspelled and repeated words as ranges (`marks()`, `subscribe()`; pass them to `DocxEditor`'s `marks`) and the actions on them (`spellingAt`, `suggestions`, `replaceWord`, `ignoreAll` for the session, `addToDictionary`). Language comes from each run's `w:lang` slot for the word's script; Word's proofing rules apply (noProof, hidden and deleted text, `w:hideSpellingErrors`, uppercase words, words with digits, addresses). The app supplies the `SpellingChecker` (web: `@filen/spellcheck`); words are found with the UAX #29 word segmenter in `core/text`.
- ProseMirror (`prosemirror-model`, `-state`, `-transform`, `-history`) is an optional peer dependency. An app that ships the editor adds these packages as its own dependencies, dedupes them with React (Vite `resolve.dedupe`, Metro `resolveRequest`), and adds ProseMirror's MIT notice to its third-party notices.

## Real-world corpus

The corpus holds 1,454 licensed real-world files in every supported format (DOCX, XLSX, CSV/TSV, PPTX, ODT, ODS, ODP, RTF, DOC, XLS, PPT and their template/macro variants); the corpus tests (`src/tests/*Corpus.test.ts`) run each engine over its own family. The files are never committed; `corpus/manifest.json` lists each one with its download URL, licence, size and SHA-256, and `corpus/README.md` lists the sources, licences and families. `src/tests/corpusFiles.test.ts` checks every fetched family, with an engine or not, against the manifest.

```sh
pnpm run corpus:fetch   # downloads ~504 MB into .corpus/ (gitignored); re-runs only fetch what is missing
pnpm test               # the corpus tests now run instead of skipping
```

- Every file is verified against its SHA-256 before it is written; files no longer in the manifest are removed.
- `--strict` fails on any unavailable file (by default up to 10 may be), `--check-remote` also checks that the sources of cached files still exist, `--dir <directory>` fetches elsewhere, `--family pptx,ppt` fetches only some families.
- The tests look for the corpus in `FILEN_OFFICE_CORPUS` (must exist when set), then `.corpus/`, then the gitignored `packages/filen-web/docs/office/corpus` where it was first assembled.
- `FILEN_OFFICE_UNICODE_FULL=1` adds the full Unicode conformance files, downloaded by `pnpm exec tsx src/scripts/unicodeTables.ts`.
- `pnpm run corpus:manifest --from <MANIFEST.json>` regenerates `corpus/` when the corpus is rebuilt.

CI runs all of it in `.github/workflows/office-corpus.yml`, on changes to this package and nightly.

### DOCX editing fuzzers and digests (`src/tests/fuzz`)

Skipped unless turned on; each job runs in its own Node process.

```sh
FILEN_OFFICE_FUZZ=3 FILEN_OFFICE_FUZZ_SEED=21 pnpm exec vitest run src/tests/fuzz/paged.fuzz.test.ts   # 3 seeds × sync, race per file
FILEN_OFFICE_FUZZ=8 FILEN_OFFICE_FUZZ_SEED=8 pnpm exec vitest run src/tests/fuzz/damage.fuzz.test.ts   # 8 sequences per file
FILEN_OFFICE_DIGEST=/tmp/viewer.json pnpm exec vitest run src/tests/fuzz/viewer.fuzz.test.ts         # writes, then compares
```

- Paged fuzz: random user edits through the paged editor; after every settled step the pages shown must equal a fresh layout of the saved file. Damage fuzz: raw transactions saved through Gates 1 and 2 and read back by an independent zip/XML reader; gate refusals are counted, anything else fails. Viewer digest: the unedited layout of every document against a baseline written by the first run.
- `FILEN_OFFICE_FUZZ_FILTER` (file name part), `_STEPS`, `_OPS` (op names), `_MODES` (paged: `sync,race,proof`; damage: `stories,undo,blockmove,undoall,xsd,selftest`), `_VIEW` (`final`, `markup`, `original`), `_JOBS` and `_OUT` (JSON lines) tune a run. One job: `FILEN_OFFICE_FUZZ=1 node --import tsx src/tests/fuzz/run.ts paged <file> <seed> <mode>` or `run.ts damage <file>`.
- Edits live in `src/tests/fuzz/ops/`, one file per feature; ops with a `script` also form the scripted edit suite of `docxWriteCorpus.test.ts`.

Known failures the fuzz finds (they predate the fuzz; run from this package):

- Tracked changes: pasting a range that holds tracked insertions duplicates their revision ids, so the save is refused (`revision:132 appears 2 times`): `FILEN_OFFICE_FUZZ=1 FILEN_OFFICE_FUZZ_OPS=paste node --import tsx src/tests/fuzz/run.ts paged .corpus/docx/lo__tdf123054.docx 8 sync`. The same document shows a paste whose pages differ from the saved file's (seed 23).
- Tracked changes: edits racing layout leave a page laid out differently from the saved file: `FILEN_OFFICE_FUZZ=1 FILEN_OFFICE_FUZZ_OPS=type,typeEnd,delete,enter,join,format,style,list,link,cell,paste node --import tsx src/tests/fuzz/run.ts paged .corpus/docx/oxpt__RA001-Tracked-Revisions-01.docx 47 race` (also seed 60).
