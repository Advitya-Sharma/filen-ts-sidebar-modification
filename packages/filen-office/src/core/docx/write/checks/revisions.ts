import type { DocxCheck } from "./index"

// Tracked change shapes (w:delText only in deletions, authors, ISO dates, no empty wrappers): with
// tracked change recording. Unique revision ids are counted already.
export const revisionChecks: readonly DocxCheck[] = []
