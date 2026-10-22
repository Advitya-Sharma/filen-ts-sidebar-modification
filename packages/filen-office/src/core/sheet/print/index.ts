export {
	EMPTY_SECTIONS,
	parseHeaderFooter,
	runText,
	sectionsEmpty,
	type HeaderFooterField,
	type HeaderFooterRun,
	type HeaderFooterSections,
	type HeaderFooterStyle,
	type HeaderFooterValues
} from "./headerFooter"
export {
	MAX_PRINT_PAGES,
	fitScale,
	headingWidth,
	orientedPaper,
	paginate,
	type PrintAxis,
	type PrintPage,
	type PrintSheet,
	type SheetPagination
} from "./paginate"
export { A4, COMMON_PAPER_SIZES, LETTER, paperPoints, paperSizeKnown, universalMeasure } from "./paper"
export {
	DEFAULT_PAGE_SETUP,
	NORMAL_MARGINS,
	readAreas,
	readPageSetup,
	readTitles,
	type SheetHeaderFooter,
	type SheetPageSetup,
	type SheetPrintMargins,
	type SheetPrintSpan
} from "./setup"
export type { SheetPrintPatch, SheetPrintRows, SheetPrintSetupResult } from "./xlsx"
