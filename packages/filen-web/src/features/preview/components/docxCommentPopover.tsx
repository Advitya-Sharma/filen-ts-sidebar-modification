import { useEffect, useState } from "react"
import { useTranslation } from "react-i18next"
import type { DocxObjectPress } from "@filen/office/react/docx/editor"
import type { DocxAuthorGate } from "@/features/preview/components/docxAuthorGate.logic"
import type { DocxObjectControlProps } from "@/features/preview/components/docxFormControls"
import { reviewView, threadAt } from "@/features/preview/components/docxReview.logic"
import { useInputForm } from "@/features/preview/components/docxEditor.logic"
import { CommentField, CommentThreadCard } from "@/features/preview/components/docxReviewLayer"
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover"

export interface DocxCommentPopoverProps extends DocxObjectControlProps {
	withAuthor: DocxAuthorGate
}

// What the popover shows, at a point of the container (px from its top left).
type Opened = { readonly x: number; readonly y: number } & ({ readonly kind: "thread"; readonly id: string } | { readonly kind: "compose" })

// Where the selection ends on screen: its page's place and the caret box on the pages; the browser's selection
// in the reflowed view.
function selectionPoint(props: DocxCommentPopoverProps): { x: number; y: number } | undefined {
	const { editor, container, surface, zoom } = props

	if (container === null) {
		return undefined
	}

	if (surface === "reflow") {
		const selection = window.getSelection()
		const rect = selection !== null && selection.rangeCount > 0 ? selection.getRangeAt(0).getBoundingClientRect() : undefined

		return rect === undefined || (rect.width === 0 && rect.height === 0) ? undefined : { x: rect.right, y: rect.bottom }
	}

	const end = editor.ends().end
	const page = end === undefined ? null : container.querySelector(`[data-fo-page="${String(end.page)}"]`)

	if (end === undefined || page === null) {
		return undefined
	}

	const rect = page.getBoundingClientRect()

	return { x: rect.left + end.x * zoom, y: rect.top + (end.top + end.height) * zoom }
}

// A comment thread in a popover at its anchor, where the margin has no room for balloons (narrow windows, Mobile
// view): a press on a comment marker (a tap on commented text in Mobile view) opens its thread, and a new comment
// is composed here at the selection.
export function DocxCommentPopover(props: DocxCommentPopoverProps) {
	const { editor, container, surface, readOnly, execute, withAuthor, press, onPressDone, inputRequest, onInputDone } = props
	const { t } = useTranslation("preview")
	const view = reviewView(editor)
	const [pressed, setPressed] = useState<Opened | null>(null)
	const [taken, setTaken] = useState<DocxObjectPress | null>(null)
	// Balloons beside the pages take presses and new comments themselves.
	const taking = surface === "reflow" || !view.aside()
	const at = (clientX: number, clientY: number): { x: number; y: number } => {
		const rect = container?.getBoundingClientRect()

		return { x: clientX - (rect?.left ?? 0), y: clientY - (rect?.top ?? 0) }
	}

	// Taken while rendering, told done after.
	if (press !== null && taking && taken !== press) {
		const id = press.hit.part ?? threadAt(view.comments, press.hit.story, press.hit.pos)
		const thread = id === undefined ? undefined : view.comments.threadOf(id)

		setTaken(press)

		if (thread !== undefined) {
			setPressed({ kind: "thread", id: thread, ...at(press.clientX, press.clientY) })
		}
	}

	useEffect(() => {
		if (taken !== null) {
			onPressDone()
		}
	}, [taken, onPressDone])

	// A new comment is composed here at the selection, the request held while it is.
	const composing = useInputForm<Opened>(
		inputRequest,
		request => request.type === "comment.add" && taking,
		() => {
			if (readOnly || !editor.canExecute("comment.add")) {
				return null
			}

			const rect = container?.getBoundingClientRect()
			const point = selectionPoint(props) ?? { x: (rect?.left ?? 0) + (rect?.width ?? 0) / 2, y: (rect?.top ?? 0) + 48 }

			return { kind: "compose", ...at(point.x, point.y) }
		}
	)
	const composingNonce = composing.request?.nonce ?? null
	const opened = composing.form ?? pressed

	useEffect(() => {
		// Asked where no comment can be added: done at once.
		if (composingNonce !== null && composing.form === null) {
			onInputDone(composingNonce)
		}
	}, [composingNonce, composing.form, onInputDone])

	const thread = opened?.kind === "thread" ? view.comments.threads().find(candidate => candidate.id === opened.id) : undefined
	const authors = view.authors()
	const close = (): void => {
		if (composing.form !== null && composingNonce !== null) {
			onInputDone(composingNonce)
		} else {
			setPressed(null)
		}
	}

	if (opened === null || (opened.kind === "thread" && thread === undefined)) {
		return null
	}

	return (
		<Popover
			open
			onOpenChange={open => {
				if (!open) {
					close()
				}
			}}
		>
			<PopoverTrigger
				aria-hidden="true"
				tabIndex={-1}
				className="pointer-events-none absolute size-px opacity-0"
				style={{ left: opened.x, top: opened.y }}
			/>
			<PopoverContent
				align="start"
				aria-label={
					thread === undefined
						? t("previewDocxCommentPlaceholder")
						: t("previewDocxCommentThread", { author: thread.comment.author ?? t("previewDocxUnknownAuthor") })
				}
				className="max-h-96 w-72 gap-3 overflow-y-auto rounded-2xl p-3 text-xs"
				data-docx-review=""
			>
				{thread === undefined ? (
					<CommentField
						initial=""
						placeholder={t("previewDocxCommentPlaceholder")}
						submitLabel={t("previewDocxCommentPost")}
						autoFocus
						onCancel={close}
						// The field stays (with what was typed) until there is a name to comment with.
						onSubmit={text => {
							withAuthor(() => {
								close()
								execute({ type: "comment.add", text })
							})
						}}
					/>
				) : (
					<CommentThreadCard
						thread={thread}
						authors={authors}
						readOnly={readOnly}
						expanded
						execute={execute}
						withAuthor={withAuthor}
					/>
				)}
			</PopoverContent>
		</Popover>
	)
}
