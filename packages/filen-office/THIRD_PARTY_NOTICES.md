# Third-party notices

`@filen/office` is licensed under the GNU Affero General Public License v3.0. It contains code derived from the projects below; each derived file carries a header naming its source.

## SuperDoc 1.x

- Source: SuperDoc v1.46.3 (`e096323e7d5e54f95b4817d1d3aefde8c37d52cf`), https://github.com/superdoc/docx-editor
- Copyright: Harbour and the SuperDoc contributors
- Licence: GNU Affero General Public License v3.0
- Derived files, modified for `@filen/office` (AGPL-3.0 §5a), first ported and last changed on 2026-10-11:
  - `src/core/docx/numbering/formats.ts` (from `shared/common/list-numbering`)
  - `src/core/docx/numbering/counter.ts` (from `packages/word-layout/src/numbering-manager.ts`)
  - `src/core/docx/styles/merge.ts` (from `packages/layout-engine/style-engine/src/cascade.ts`)
  - `src/core/docx/styles/resolve.ts` and `src/core/docx/styles/tableStyle.ts` (from `packages/layout-engine/style-engine/src/ooxml/index.ts`)

Only the AGPL-3.0 1.x line is used. No code from `superdoc` 2.3 or later, or from `@superdoc/docx-engine`, is included.
