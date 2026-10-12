import type { DocxTransactionTransformer, FeatureHandlers } from "../registry"
import { COMMENT_HANDLERS } from "./comments"
import { DRAWING_HANDLERS } from "./drawings"
import { FIELD_HANDLERS } from "./fields"
import { FORMS_AFTER, FORMS_BEFORE, FORM_HANDLERS } from "./forms"
import { FORMATTING_HANDLERS } from "./formatting"
import { INSERT_HANDLERS } from "./insert"
import { MATH_HANDLERS } from "./math"
import { NOTE_HANDLERS } from "./notes"
import { REVIEW_HANDLERS, REVIEW_TRANSFORMER } from "./review"
import { SECTION_HANDLERS } from "./sections"
import { TABLE_HANDLERS } from "./tables"

// Every feature's handlers (base commands are the session's own).
export const FEATURE_HANDLERS: readonly FeatureHandlers[] = [
	FORMATTING_HANDLERS,
	INSERT_HANDLERS,
	COMMENT_HANDLERS,
	REVIEW_HANDLERS,
	TABLE_HANDLERS,
	SECTION_HANDLERS,
	NOTE_HANDLERS,
	FIELD_HANDLERS,
	FORM_HANDLERS,
	DRAWING_HANDLERS,
	MATH_HANDLERS
]

// Forms around tracked changes: before (a protection that forces tracking turns it on), after (protection,
// content control locks and placeholders, on what tracking made of the transaction).
export const FEATURE_TRANSFORMERS: readonly DocxTransactionTransformer[] = [FORMS_BEFORE, REVIEW_TRANSFORMER, FORMS_AFTER].filter(
	(transform): transform is DocxTransactionTransformer => transform !== null
)
