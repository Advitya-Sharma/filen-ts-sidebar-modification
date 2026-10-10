export { buildDocxModel, DocxModel, type CommentAnchorTable, type HeaderFooterSet, type MediaRef, type Section } from "./document"
export type { BuildCounts, MaxIds } from "./build"
export type { CommentExtension, CommentInfo, Person } from "./comments"
export type { AnchorInfo, AnchorPosition, DiagramIds, DrawingInfo, DrawingPlacement, VmlInfo, VmlKind, WrapType } from "./drawings"
export {
	BreakClear,
	BreakType,
	FieldCharType,
	ITEM_NAMES,
	Item,
	REVISION_NAMES,
	Revision,
	RevisionFlag,
	RevisionView,
	SdtLevel,
	TRANSPARENT,
	revisionHidden
} from "./kinds"
export type { GraphicKind } from "./namespaces"
export {
	StoryPart,
	type AlternateTable,
	type BookmarkTable,
	type CellTable,
	type FieldTable,
	type HyperlinkTarget,
	type ParagraphTable,
	type PartItems,
	type PartTables,
	type RevisionTable,
	type RowTable,
	type RunTable,
	type SdtTable,
	type TableTable
} from "./part"
export { PartReader, partReader } from "./reader"
export { NO_SETTINGS, readSettings, type CompatSetting, type DocumentProtection, type DocxSettings, type WriteProtection } from "./settings"
export { modelStats, type DocxModelStats } from "./stats"
export type { NoteType, Story, StoryKind } from "./story"
export { ItemWalker } from "./walker"
