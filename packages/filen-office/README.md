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
- ProseMirror (`prosemirror-model`, `-state`, `-transform`, `-history`) is an optional peer dependency. An app that ships the editor adds these packages as its own dependencies, dedupes them with React (Vite `resolve.dedupe`, Metro `resolveRequest`), and adds ProseMirror's MIT notice to its third-party notices.

## Real-world corpus

The corpus tests (`src/tests/*Corpus.test.ts`) run the engines over 510 licensed real-world documents. The files are never committed; `corpus/manifest.json` lists each one with its download URL, licence, size and SHA-256, and `corpus/README.md` lists the sources and licences.

```sh
pnpm run corpus:fetch   # downloads ~84 MB into .corpus/ (gitignored); re-runs only fetch what is missing
pnpm test               # the corpus tests now run instead of skipping
```

- Every file is verified against its SHA-256 before it is written; files no longer in the manifest are removed.
- `--strict` fails on any unavailable file (by default up to 10 may be), `--check-remote` also checks that the sources of cached files still exist, `--dir <directory>` fetches elsewhere.
- The tests look for the corpus in `FILEN_OFFICE_CORPUS` (must exist when set), then `.corpus/`, then the gitignored `packages/filen-web/docs/office/corpus` where it was first assembled.
- `FILEN_OFFICE_UNICODE_FULL=1` adds the full Unicode conformance files, downloaded by `pnpm exec tsx src/scripts/unicodeTables.ts`.
- `pnpm run corpus:manifest --from <MANIFEST.json>` regenerates `corpus/` when the corpus is rebuilt.

CI runs all of it in `.github/workflows/office-corpus.yml`, on changes to this package and nightly.
