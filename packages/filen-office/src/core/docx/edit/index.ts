export { DocxDocBuilder, type StoryBuildStats, type StoryDoc } from "./build"
export {
	DOCX_EDITOR_COMMANDS,
	DOCX_PARAGRAPH_STYLES,
	DOCX_PARAGRAPH_STYLE_MENU,
	PARAGRAPH_STYLE_NAMES,
	isWriteCommand,
	parseDocxEditorCommand,
	type DocxAlignment,
	type DocxEditorCommand,
	type DocxEditorCommandGroup,
	type DocxEditorCommandInfo,
	type DocxEditorCommandType,
	type DocxParagraphStyle,
	type DocxParagraphStyleInfo
} from "./commands"
export { openDocxEditContext, type DocxEditContext, type EditContextParts } from "./context"
export { FormatResolver, alignmentOf, type ListKind, type RunToggle } from "./formats"
export { RunBases, formattingMarks, type RunBase } from "./formatting"
export {
	changedRanges,
	docxInvariants,
	docxInvariantsKey,
	freshSlice,
	invariantViolation,
	normalizeListAttrs,
	repairTransaction,
	type InvariantOptions,
	type ListAttrs,
	type ListNormalizer
} from "./invariants"
export { conservativeProof, lockBlocks, proveBlocks, type BlockProof, type LockRequest } from "./proof"
export {
	FORMATTING_MARKS,
	NO_MARKERS,
	NO_SOURCE,
	REVISION_MARKS,
	docxSchema,
	hasSource,
	endsWithTable,
	inlineLevel,
	markerKey,
	sourceOf,
	survivorsOf,
	type ColorAttrs,
	type DocxMarkName,
	type DocxNodeName,
	type DocxSchema,
	type IndentAttrs,
	type InlineLevel,
	type LockReason,
	type LockedKind,
	type MarkerAttrs,
	type MarkerKind,
	type MarkerList
} from "./schema"
export { DocxEditorSession, type DocxEditorOptions, type DocxEditorSnapshot, type DocxReadOnlyReason, type FormatState } from "./session"
