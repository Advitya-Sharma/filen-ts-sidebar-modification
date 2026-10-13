export { DocxDocBuilder, type StoryBuildStats, type StoryDoc } from "./build"
export {
	DOCX_COMMAND_FEATURES,
	DOCX_EDITOR_COMMANDS,
	DOCX_PARAGRAPH_STYLES,
	DOCX_PARAGRAPH_STYLE_MENU,
	PARAGRAPH_STYLE_NAMES,
	commandInfo,
	isWriteCommand,
	parseDocxEditorCommand,
	type BaseCommand,
	type DocxAlignment,
	type DocxCommandContext,
	type DocxCommandFeature,
	type DocxCommandInfo,
	type DocxCommandInput,
	type DocxCommandOf,
	type DocxCommandScope,
	type DocxEditorCommand,
	type DocxEditorCommandGroup,
	type DocxEditorCommandInfo,
	type DocxEditorCommandType,
	type DocxParagraphStyle,
	type DocxParagraphStyleInfo
} from "./commands"
export type { CommentCommand } from "./commands/comments"
export { CROP_WHOLE } from "./commands/drawings"
export type { DocxCrop, DocxRelativeH, DocxRelativeV, DocxWrapMode, DrawingCommand } from "./commands/drawings"
export { FIELD_CATALOGUE, fieldCodeType, type FieldCommand } from "./commands/fields"
export { MAX_BOOKMARK_NAME, sessionBookmarks, validBookmarkName, type BookmarkIndex, type DocxBookmark } from "../fields/bookmarks"
export {
	DEFAULT_DATE_PICTURE,
	DEFAULT_TIME_PICTURE,
	UPDATABLE_FIELDS,
	datePicture,
	numericPicture,
	parseFieldCode,
	type FieldCode,
	type FieldSwitch
} from "../fields/evaluate"
export { NO_ENTRIES, tocInstruction, tocOptions, tocRegions, type TocOptions, type TocRegion } from "../fields/toc"
export { DOCX_CONTROL_KINDS, type DocxControlItemInput, type DocxControlKind, type FormCommand } from "./commands/forms"
export {
	DOCX_HIGHLIGHT_COLORS,
	DOCX_THEME_COLORS,
	DOCX_FONT_SIZES,
	DOCX_THEME_BODY_FONT,
	DOCX_THEME_HEADINGS_FONT,
	type DocxHighlightColor,
	type DocxLineSpacingRule,
	type DocxTextColor,
	type DocxThemeColor,
	type FormattingCommand
} from "./commands/formatting"
export type { InsertCommand } from "./commands/insert"
export type { MathCommand } from "./commands/math"
export type { NoteCommand } from "./commands/notes"
export { DOCX_MARKUP_VIEWS, type DocxMarkupView, type ReviewCommand } from "./commands/review"
export {
	DOCX_PAGE_NUMBER_FORMATS,
	type DocxOrientation,
	type DocxPageNumberFormat,
	type DocxPageSetup,
	type DocxSectionBreak,
	type SectionCommand
} from "./commands/sections"
export type { DocxCellVAlign, DocxTableAlign, TableCommand } from "./commands/tables"
export type { ViewCommand } from "./commands/view"
export {
	DocxIds,
	DocxMediaStore,
	authorInitials,
	openDocxEditContext,
	type DocxEditContext,
	type DocxIdSpace,
	type DocxIdentity,
	type DocxMediaEntry,
	type DocxMediaInfo,
	type DocxReviewState,
	type EditContextParts,
	type NewStory,
	type NewStoryKind
} from "./context"
export {
	documentHandler,
	registeredTransformer,
	storyHandler,
	type DocumentCommandContext,
	type DocumentCommandHandler,
	type DocumentHandlers,
	type DocxTransactionTransformer,
	type FeatureHandlers,
	type StoryCommandContext,
	type StoryCommandHandler,
	type StoryCommandResult,
	type StoryHandlers
} from "./registry"
export { readOnlyReasons, textboxHolderLocked } from "./readOnly"
export {
	MAX_IMAGE_BYTES,
	MAX_IMAGE_PIXELS,
	sniffImage,
	type DocxImageFormat,
	type DocxImageInfo,
	type DocxImageRefusal
} from "../objects/images"
export { NO_PROTECTION, commandAllowed, protectionMode, protectionOf, type DocxProtection, type DocxProtectionMode } from "./protection"
export { NO_VALUES } from "./snapshot/formats"
export { DOCX_COMMON_FONTS, docxFontList, type DocxFontChoice, type DocxFontKind } from "./fontList"
export type { DocxFormatValue, DocxIndentValue, DocxLineSpacingValue, DocxParagraphSpacingValue } from "./snapshot/formats"
export { NO_REVIEW } from "./snapshot/review"
export type {
	DocxCommentContext,
	DocxContextSnapshot,
	DocxControlContext,
	DocxDrawingContext,
	DocxEquationContext,
	DocxFieldContext,
	DocxFieldRefusal,
	DocxLinkContext,
	DocxProvidedSnapshot,
	DocxReviewSnapshot,
	DocxStoryContext,
	DocxTableContext,
	DocxValuesSnapshot,
	SnapshotInput
} from "./snapshot/index"
export {
	STORY_FIND_LIMIT,
	findInStory,
	findPattern,
	matchHolds,
	replaceAll,
	replaceMatches,
	selectMatch,
	type StoryFindOptions,
	type StoryMatch
} from "./find"
export { FormatResolver, alignmentOf, type ListKind, type RunToggle } from "./formats"
export { RunBases, formattingMarks, type RunBase } from "./formatting"
export {
	INVARIANT_STAGES,
	changedRanges,
	docxInvariantPlugins,
	docxInvariants,
	docxInvariantsKey,
	freshSlice,
	invariantViolation,
	isCommand,
	normalizeListAttrs,
	repairTransaction,
	writableNew,
	type InvariantOptions,
	type InvariantStage,
	type ListAttrs,
	type ListNormalizer
} from "./invariants/index"
export { conservativeProof, lockBlocks, proveBlocks, type BlockProof, type LockRequest } from "./proof"
export {
	BLOCK_NODE_MARKS,
	FORMATTING_MARKS,
	INLINE_MARK_NAMES,
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
	type DrawingOverrides,
	type DocxMarkName,
	type DocxNodeName,
	type DocxSchema,
	type IndentAttrs,
	type SpacingAttrs,
	type InlineLevel,
	type LockReason,
	type LockedKind,
	type MarkerAttrs,
	type MarkerKind,
	type MarkerList,
	type MarkerNodeAttrs
} from "./schema"
export {
	REVERTIBLE_UNITS,
	findUnsavableBlocks,
	type RevertibleUnit,
	type RevertibleUnits,
	type UnsavableBlocks,
	type UnsavableSearchOptions,
	withRevertedUnits
} from "./unsavable"
export { ReviewIndex, ReviewKind, revisionMarkId, type ReviewEntry } from "../review/index"
export {
	resolveChanges,
	reviewIndexWanted,
	storyReviewIndex,
	wantReviewIndex,
	type ReviewOutcome,
	type ReviewScope
} from "../review/accept"
export {
	flushComposition,
	minuteDate,
	paragraphChange,
	revisionStamp,
	setTrackingPersisted,
	trackingPersisted,
	type ParagraphAttrs,
	type ParagraphChange,
	type RevisionStamp
} from "../review/track"
export { DocxEditorSession, type DocxEditorOptions, type DocxEditorSnapshot, type DocxReadOnlyReason, type FormatState } from "./session"
