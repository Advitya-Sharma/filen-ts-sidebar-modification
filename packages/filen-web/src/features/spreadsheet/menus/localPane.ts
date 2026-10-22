import { useEffect, useRef } from "react"
import { type } from "arktype"
import type { LayerKey, SheetPane } from "@filen/office/core/sheet"
import type { SheetGridModel, SheetGridState } from "@filen/office/react/sheet"
import { kvDelete, kvGetJson, kvSetJson } from "@/lib/storage/adapter"
import { log } from "@/lib/log"
import { samePane } from "@/features/spreadsheet/menus/structure.logic"

// Frozen panes of a text file (a CSV cannot hold them), kept beside it as its sizes are: by the file's stable id in
// the app's key-value store, or for the session alone when it has none. A bounded index forgets the files frozen
// longest ago.

export const LOCAL_PANE_LIMIT = 200

const INDEX_KEY = "spreadsheet.pane.index.v1"
const paneKey = (id: string): string => `spreadsheet.pane.v1.${id}`
const paneSchema = type({
	rows: "number.integer >= 0",
	cols: "number.integer >= 0",
	topRow: "number.integer >= 0",
	leftCol: "number.integer >= 0"
})
const indexSchema = type("string[]")
const sessionPanes = new Map<string, SheetPane>()

export async function loadLocalPane(key: LayerKey): Promise<SheetPane | null> {
	if (key.kind === "session") return sessionPanes.get(key.id) ?? null

	const stored = await kvGetJson(paneKey(key.id), paneSchema)

	// A text file's frozen rows and columns start at its first.
	return stored === null ? null : { ...stored, firstRow: 0, firstCol: 0 }
}

export async function saveLocalPane(key: LayerKey, pane: SheetPane | null): Promise<void> {
	if (key.kind === "session") {
		if (pane === null) sessionPanes.delete(key.id)
		else sessionPanes.set(key.id, pane)

		return
	}

	const index = ((await kvGetJson(INDEX_KEY, indexSchema)) ?? []).filter(id => id !== key.id)

	if (pane === null) {
		await kvDelete(paneKey(key.id))
		await kvSetJson(INDEX_KEY, index)

		return
	}

	index.push(key.id)

	const evicted = index.splice(0, Math.max(0, index.length - LOCAL_PANE_LIMIT))

	await kvSetJson(paneKey(key.id), { rows: pane.rows, cols: pane.cols, topRow: pane.topRow, leftCol: pane.leftCol })
	await Promise.all(evicted.map(id => kvDelete(paneKey(id))))
	await kvSetJson(INDEX_KEY, index)
}

function keyOf(kind: LayerKey["kind"], id: string): LayerKey {
	return kind === "stable" ? { kind: "stable", id } : { kind: "session", id }
}

// A text file's frozen panes: what was kept beside it is frozen again once its session opens, and every change of the
// panes from then on is kept. Off (`enabled` false) for a workbook, which holds its own.
export function useLocalPane(grid: SheetGridState, key: LayerKey, enabled: boolean): void {
	const { model } = grid
	const pane = grid.layout?.pane ?? null
	const { kind, id } = key
	// What the store holds for this model and file, once read; until then nothing is kept (it would replace it).
	const kept = useRef<{ model: SheetGridModel; keyId: string; pane: SheetPane | null } | null>(null)
	const keyId = `${kind}:${id}`

	useEffect(() => {
		if (!enabled) {
			return undefined
		}

		// An object: the load's callbacks read it after awaits.
		const run = { live: true }

		kept.current = null
		// What was kept is frozen again, then kept as the file's from then on (a session that went meanwhile does neither).
		const replay = async (stored: SheetPane | null): Promise<void> => {
			if (stored === null || !run.live) {
				return
			}

			const sheet = model.getSnapshot().sheet

			await model
				.command(
					{ type: "view.freeze", sheet, rows: stored.rows, cols: stored.cols, topRow: stored.topRow, leftCol: stored.leftCol },
					false
				)
				.catch(() => undefined)
		}

		loadLocalPane(keyOf(kind, id))
			.catch(() => null)
			.then(async stored => {
				await replay(stored)
				keep(stored)
			})
			.catch(() => undefined)

		function keep(stored: SheetPane | null): void {
			if (run.live) {
				kept.current = { model, keyId: `${kind}:${id}`, pane: stored }
			}
		}

		return () => {
			run.live = false
		}
	}, [model, kind, id, enabled])

	useEffect(() => {
		const current = kept.current

		if (!enabled || current?.model !== model || current.keyId !== keyId || samePane(current.pane, pane)) {
			return
		}

		kept.current = { ...current, pane }
		saveLocalPane(keyOf(kind, id), pane).catch((e: unknown) => {
			log.error("spreadsheet", "keeping frozen panes failed", e)
		})
	}, [enabled, model, keyId, kind, id, pane])
}
