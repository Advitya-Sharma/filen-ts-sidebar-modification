import type { DocxCommandLabels, DocxCommandVariant } from "@/features/preview/components/docxEditor.logic"

// Headers and footers (entered from the Insert menu or a double click; the header chip holds their toggles),
// section breaks (Insert menu) and page setup (the ⋯ menu's dialog). Deleting a section break is not offered.
export const SECTIONS_LABELS: DocxCommandLabels<"sections"> = {
	"insert.sectionBreak": "previewDocxSectionBreak",
	"header.edit": "previewDocxHeaderEdit",
	"footer.edit": "previewDocxFooterEdit",
	"header.differentFirst": "previewDocxHeaderDifferentFirst",
	"header.differentOddEven": "previewDocxHeaderDifferentOddEven",
	"header.linkToPrevious": "previewDocxHeaderLinkToPrevious",
	"header.remove": "previewDocxHeaderRemove",
	"footer.remove": "previewDocxFooterRemove",
	"page.setup": "previewDocxPageSetup"
}

// Insert ▸ Section break ▸ each kind.
export const SECTION_BREAK_VARIANTS: readonly DocxCommandVariant[] = [
	{ label: "previewDocxSectionBreakNextPage", command: { type: "insert.sectionBreak", kind: "nextPage" } },
	{ label: "previewDocxSectionBreakContinuous", command: { type: "insert.sectionBreak", kind: "continuous" } },
	{ label: "previewDocxSectionBreakEvenPage", command: { type: "insert.sectionBreak", kind: "evenPage" } },
	{ label: "previewDocxSectionBreakOddPage", command: { type: "insert.sectionBreak", kind: "oddPage" } }
]
