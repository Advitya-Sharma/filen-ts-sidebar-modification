import { useTranslation } from "react-i18next"
import { parseDocxEditorCommand, type DocxEditorCommand, type DocxEditorSnapshot } from "@filen/office/core/docx/edit"
import {
	commandVariants,
	contextCommands,
	controlLockedKey,
	type DocxRequestInput,
	type DocxSurface
} from "@/features/preview/components/docxEditor.logic"
import {
	ContextMenuItem,
	ContextMenuSeparator,
	ContextMenuSub,
	ContextMenuSubContent,
	ContextMenuSubTrigger
} from "@/components/ui/context-menu"

// The context menu's items for where the selection is (after the clipboard's): accept or reject the tracked
// change, a new comment, the table's rows, columns and cells, the control's content (and why it cannot
// change), the link, the field and the note there, the picture's items (no crop in Mobile view, which has no
// handles). Read when the menu renders, never per keystroke.

export interface DocxContextCommandsProps {
	snapshot: DocxEditorSnapshot | undefined
	readOnly: boolean
	execute: (command: DocxEditorCommand) => void
	requestInput: DocxRequestInput
	surface: DocxSurface
}

export function DocxContextCommands({ snapshot, readOnly, execute, requestInput, surface }: DocxContextCommandsProps) {
	const { t } = useTranslation("preview")

	if (snapshot === undefined) {
		return null
	}

	const commands = contextCommands(snapshot, readOnly).filter(command => !(surface === "reflow" && command.info.type === "drawing.crop"))
	const locked = snapshot.context.control?.locked ?? null

	if (commands.length === 0 && locked === null) {
		return null
	}

	return (
		<>
			<ContextMenuSeparator />
			{locked === null ? null : (
				<ContextMenuItem
					disabled
					data-docx-control-locked=""
				>
					{t(controlLockedKey(locked))}
				</ContextMenuItem>
			)}
			{commands.map(({ info, label, action }) =>
				action === "variants" ? (
					<ContextMenuSub key={info.type}>
						<ContextMenuSubTrigger>{t(label)}</ContextMenuSubTrigger>
						<ContextMenuSubContent>
							{(commandVariants(info.type) ?? []).map(variant => (
								<ContextMenuItem
									key={variant.label}
									onClick={() => {
										if ("command" in variant) {
											execute(variant.command)
										} else {
											requestInput(info.type, variant.input)
										}
									}}
								>
									{t(variant.label)}
								</ContextMenuItem>
							))}
						</ContextMenuSubContent>
					</ContextMenuSub>
				) : (
					<ContextMenuItem
						key={info.type}
						onClick={() => {
							if (action === "input") {
								requestInput(info.type)
							} else {
								const command = parseDocxEditorCommand({ type: info.type })

								if (command !== undefined) {
									execute(command)
								}
							}
						}}
					>
						{t(label)}
					</ContextMenuItem>
				)
			)}
		</>
	)
}
