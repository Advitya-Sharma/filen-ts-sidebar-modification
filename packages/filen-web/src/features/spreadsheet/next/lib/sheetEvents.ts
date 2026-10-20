import type { SheetEvent } from "@filen/office/core/sheet"

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
