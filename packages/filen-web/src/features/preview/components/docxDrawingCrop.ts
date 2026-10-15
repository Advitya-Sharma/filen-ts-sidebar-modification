import { useSyncExternalStore } from "react"
import type { DocxEditorCommand } from "@filen/office/core/docx/edit"

// Crop mode: the picture whose crop handles show on the pages (docxObjectHandles.tsx), switched by the Picture
// group (docxDrawingMenu.tsx). Kept per editor view, keyed by the view's `execute` (one stable function per
// open editor, which both are given): two editors never share it.

export type DocxViewExecute = (command: DocxEditorCommand) => void

interface CropState {
	// The picture being cropped (its position in the story edited), or null.
	pos: number | null
	readonly listeners: Set<() => void>
}

const STATES = new WeakMap<DocxViewExecute, CropState>()

function stateOf(view: DocxViewExecute): CropState {
	let state = STATES.get(view)

	if (state === undefined) {
		state = { pos: null, listeners: new Set() }
		STATES.set(view, state)
	}

	return state
}

export function setDrawingCrop(view: DocxViewExecute, pos: number | null): void {
	const state = stateOf(view)

	if (state.pos === pos) {
		return
	}

	state.pos = pos

	for (const listener of state.listeners) {
		listener()
	}
}

export function drawingCrop(view: DocxViewExecute): number | null {
	return STATES.get(view)?.pos ?? null
}

export function useDrawingCrop(view: DocxViewExecute): number | null {
	return useSyncExternalStore(
		listener => {
			const { listeners } = stateOf(view)

			listeners.add(listener)

			return () => {
				listeners.delete(listener)
			}
		},
		() => drawingCrop(view),
		() => drawingCrop(view)
	)
}
