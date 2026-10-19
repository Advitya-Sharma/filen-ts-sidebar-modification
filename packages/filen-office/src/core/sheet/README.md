# core/sheet: the spreadsheet engine's contracts

Three files fix how the parts of the spreadsheet engine meet. Each part can then be built on its own:

- `protocol.ts`: what a renderer and an app exchange with an open workbook (tiles, commands, queries, events, saving and opening), the same on a worker host and an in-thread host;
- `registry.ts`: how a feature adds commands, queries and per-tile, per-layout and per-cell data;
- `../xlsx/model/types.ts`: the lossless workbook model that the reader builds, edits change and the splice writer saves.

These files change only by agreement. Everything else in the engine is written against them.

## Layers

| Layer | Where | Talks through |
|---|---|---|
| Workbook model, reader, writer, gate, refs catalogue, xlsx edits | `core/xlsx/**` | `XlsxWorkbookModel`, `XlsxChangeLog` |
| Session (history, recalculation, tile building), commands, numfmt, styles, csv | `core/sheet/**` | `SheetFeature`, `SheetSessionApi` |
| Hosts: worker server and client, in-thread host, encryption | `dom/sheet/**` | `SheetHost` |
| Grid: tile cache, windowing, selection | `react/sheet/**` | `SheetSessionApi`, `SheetTile` |

Core stays DOM-free and Hermes-safe (the eslint core rules). The in-thread host runs the same session code on the calling thread and passes `Slicing` so that long work yields.

## Protocol rules

- **Order.** A host runs one session's calls in the order they were made and answers them in that order. Events come between answers. A client that applies everything in arrival order is consistent; no version checks are needed beyond that.
- **Tiles.** A tile is `TILE_ROWS` × `TILE_COLS` cells, sparse, with its own string table. Styles are indices into the session's render-style table, which grows by `SheetStyleAppend`.
  - The tiles a client asked for are *held* until it calls `release`.
  - A change patches the held tiles (`SheetTilePatch`) and invalidates regions. Anything else the client cached for those regions is dropped.
- **Transfer.** Every typed array in a message owns its whole buffer, so a worker host transfers it. Never send a view into the store.
- **Refuse, never drop.** A command that the writer could not save is refused with a typed `SheetRefusal`. A save never throws: it answers `SheetSaveResult`.
- **Before a handler runs, the session checks** `SheetCommandInfo`:
  - write commands while view-only, and the wait for the proof;
  - sheet protection, by the action named;
  - workbook structure protection.

## Adding a feature

1. In the feature's own module, declare its names on the open maps. Inside the block, the protocol's own types (`SheetId`, `SheetRange`, …) are in scope without an import:

   ```ts
   declare module "../protocol" {
   	interface SheetCommandMap {
   		"merge.cells": { readonly sheet: SheetId; readonly ranges: readonly SheetRange[]; readonly center: boolean }
   	}
   	interface SheetRefusalReasons {
   		mergeLosesValues: true
   	}
   }
   ```

   The open maps are `SheetCommandMap`, `SheetQueryMap`, `SheetTileLayerMap`, `SheetLayoutExtensionMap`, `SheetCellDetailMap` and `SheetRefusalReasons`.
   - A command is named `<feature>.<verb>`. Commands and queries share one namespace.
   - A new tile flag claims a free bit in `CellFlag` (17–31).

2. Export a `SheetFeature<D>` (`defineSheetFeature`). It holds the handlers for those names, an optional tile producer, layout producers and cell-detail producers. `D` is the document the session's features work on: the xlsx session's own document, or a text file's.
3. Handlers report what they changed to `SheetChangeSink` and return a `SheetEditStep` (undo and redo report the same way), a `{ refused }`, or null.
   - Recalculation, tile patches and history are the session's.
   - An xlsx edit records what the writer must serialise in `XlsxChangeLog`. For a section or part that the writer does not own, it passes the writer function.
4. Add the feature to the session's feature list. `sheetRegistry` throws when two features claim one name, so a clash shows at the first test.

No dispatch by `switch` on a command type. Two features never edit one file.

## Model rules

See the header of `../xlsx/model/types.ts`. In short:

- Spans index the source part's inflated bytes, and parts are inflated again at save.
- Change marks are conservative, and the writer emits source bytes for anything equal to its source.
- Untouched cells in a rewritten row are copied from their source bytes.
- Styles, shared strings and formulas are append-only.

## Encryption

The host classifies and unlocks the file with `@filen/office/dom/encryption` (`openEncryptedPackage` with `tryDefault`, for Excel's default password).

- An unlocked file opens from its decrypted package with a `SheetEncryption`.
- Every save is sealed by `SheetSealer` (`PackageSealer`) after the gate, then checked with `isSealedOfficeFile` before it is handed out.
- A session in another realm than the unlock takes the package and the key through `SheetUnlockedPackage`, as DOCX editing does.
- Closing releases the key.
