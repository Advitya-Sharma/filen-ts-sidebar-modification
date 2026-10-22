import type { SheetEvent, SheetId, SheetLayout } from "@filen/office/core/sheet"

// A session's events from the moment it opens: held until the grid's model takes them (it is made after the open
// answers, from the workbook as opened), then passed straight on, in order.
export class SheetEventRelay {
	private held: SheetEvent[] = []
	private target: ((event: SheetEvent) => void) | null = null

	readonly emit = (event: SheetEvent): void => {
		if (this.target === null) {
			this.held.push(event)
		} else {
			this.target(event)
		}
	}

	// The layouts held events carry, made again (the workbook's digit width changed after they were made): each is
	// replaced by `layout`'s for its sheet. Events arriving meanwhile are held after them, already made anew.
	async relayout(layout: (sheet: SheetId) => Promise<SheetLayout>): Promise<void> {
		const count = this.held.length

		for (let k = 0; k < count; k++) {
			const event = this.held[k]

			if (event === undefined) continue

			if (event.type === "sheetReady") {
				this.held[k] = { ...event, layout: await layout(event.sheet.id) }
			} else if (event.type === "change" && event.change.layouts.length > 0) {
				const layouts = await Promise.all(
					// The whole layout (its sheet the patch's own) stands for every field the patch had.
					event.change.layouts.map(async patch => ({ sheet: patch.sheet, layout: await layout(patch.sheet) }))
				)

				this.held[k] = { ...event, change: { ...event.change, layouts } }
			}
		}
	}

	attach(target: (event: SheetEvent) => void): () => void {
		const held = this.held

		this.held = []
		this.target = target

		for (const event of held) {
			target(event)
		}

		return () => {
			if (this.target === target) {
				this.target = null
			}
		}
	}
}
