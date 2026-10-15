import type { DocxCommandLabels } from "@/features/preview/components/docxEditor.logic"

export const REVIEW_LABELS: DocxCommandLabels<"review"> = {
	"review.track": "previewDocxTrackChanges",
	"review.accept": "previewDocxAcceptChange",
	"review.reject": "previewDocxRejectChange",
	"review.acceptAll": "previewDocxAcceptAll",
	"review.rejectAll": "previewDocxRejectAll",
	"review.next": "previewDocxNextChange",
	"review.prev": "previewDocxPreviousChange",
	"view.markup": "previewDocxMarkup"
}
