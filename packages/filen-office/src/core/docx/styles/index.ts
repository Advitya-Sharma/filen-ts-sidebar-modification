export { TOGGLE_KEYS, mergeCell, mergeIndent, mergeParagraph, mergeRow, mergeRun, mergeTable } from "./merge"
export {
	DEFAULT_FONT,
	DEFAULT_SIZE,
	PropertyResolver,
	resolveSection,
	type CellFormat,
	type ParagraphFormat,
	type ResolverParts,
	type RowPosition,
	type TableFormat
} from "./resolve"
export { DEFAULT_STYLE_SETTINGS, readStyleSettings, styleSettings, type StyleSettings } from "./settings"
export {
	DocxStyles,
	MAX_STYLE_DEPTH,
	TABLE_STYLE_TYPES,
	canonicalStyleName,
	headingLevel,
	parseStyles,
	type LatentStyle,
	type LatentStyles,
	type StyleDef,
	type StyleType,
	type TableStyleCondition,
	type TableStyleType
} from "./styles"
export { CONDITION_ORDER, DEFAULT_TABLE_LOOK, ROW_CONDITIONS, cellConditions, conditionBit, type CellPosition } from "./tableStyle"
