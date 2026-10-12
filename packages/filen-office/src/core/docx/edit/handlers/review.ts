import type { DocxTransactionTransformer, FeatureHandlers } from "../registry"

// Tracked changes: recording (the transformer), accept and reject, the markup view. None yet: the commands are disabled.
export const REVIEW_HANDLERS: FeatureHandlers = {}

// Records edits as tracked changes while the document's review.tracking is on; null until it does.
export const REVIEW_TRANSFORMER: DocxTransactionTransformer | null = null
