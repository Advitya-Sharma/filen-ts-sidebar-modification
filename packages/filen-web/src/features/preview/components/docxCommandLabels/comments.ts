import type { DocxCommandLabels } from "@/features/preview/components/docxEditor.logic"

export const COMMENTS_LABELS: DocxCommandLabels<"comments"> = {
	"comment.add": "previewDocxNewComment",
	"comment.reply": "previewDocxCommentReply",
	"comment.edit": "previewDocxCommentEdit",
	"comment.resolve": "previewDocxCommentResolve",
	"comment.reopen": "previewDocxCommentReopen",
	"comment.delete": "previewDocxCommentDelete",
	"comment.next": "previewDocxNextComment",
	"comment.prev": "previewDocxPreviousComment"
}
