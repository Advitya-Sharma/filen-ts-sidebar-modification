import type { EditorState } from "prosemirror-state"
import type { DocxEditContext } from "../context"
import type { DocxEditorSession } from "../session"
import { commentOf, type DocxCommentContext } from "./comments"
import { drawingOf, type DocxDrawingContext } from "./drawing"
import { fieldOf, type DocxFieldContext } from "./field"
import { linkOf, valuesOf, type DocxLinkContext, type DocxValuesSnapshot } from "./formats"
import { equationOf, type DocxEquationContext } from "./math"
import { reviewOf, type DocxReviewSnapshot } from "./review"
import { controlOf, type DocxControlContext } from "./sdt"
import { storyOf, type DocxStoryContext } from "./story"
import { tableOf, type DocxTableContext } from "./table"

// The parts of a session's snapshot each feature provides. A provider reads only what the selection's
// ancestor path or an incremental index gives (never the whole document), and returns the same object
// while its answer is the same, so a keystroke changes nothing a toolbar compares.

export interface SnapshotInput {
	readonly session: DocxEditorSession
	readonly state: EditorState
	readonly context: DocxEditContext
	readonly readOnly: boolean
}

// Where the selection is.
export interface DocxContextSnapshot {
	readonly story: DocxStoryContext
	readonly table: DocxTableContext | null
	readonly drawing: DocxDrawingContext | null
	readonly control: DocxControlContext | null
	readonly field: DocxFieldContext | null
	readonly link: DocxLinkContext | null
	readonly comment: DocxCommentContext | null
	readonly equation: DocxEquationContext | null
}

export interface DocxProvidedSnapshot {
	readonly values: DocxValuesSnapshot
	readonly context: DocxContextSnapshot
	readonly review: DocxReviewSnapshot
}

const CONTEXT_KEYS = ["story", "table", "drawing", "control", "field", "link", "comment", "equation"] as const

// The providers' parts for `input`; `previous` (the last snapshot's) is returned as far as nothing changed.
export function providedSnapshot(input: SnapshotInput, previous: DocxProvidedSnapshot | undefined): DocxProvidedSnapshot {
	const context: DocxContextSnapshot = {
		story: storyOf(input),
		table: tableOf(input),
		drawing: drawingOf(input),
		control: controlOf(input),
		field: fieldOf(input),
		link: linkOf(input),
		comment: commentOf(input),
		equation: equationOf(input)
	}
	const values = valuesOf(input)
	const review = reviewOf(input)
	const sameContext = previous !== undefined && CONTEXT_KEYS.every(key => previous.context[key] === context[key])

	if (previous !== undefined && sameContext && previous.values === values && previous.review === review) return previous

	return { values, context: sameContext && previous !== undefined ? previous.context : context, review }
}

export type {
	DocxCommentContext,
	DocxControlContext,
	DocxDrawingContext,
	DocxEquationContext,
	DocxFieldContext,
	DocxLinkContext,
	DocxReviewSnapshot,
	DocxStoryContext,
	DocxTableContext,
	DocxValuesSnapshot
}
