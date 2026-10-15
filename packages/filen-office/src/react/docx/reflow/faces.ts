import type { PagedEditor } from "../../../core/docx/paged/index"
import type { CssFaces } from "../../../core/docx/reflow/css"

// The faces the pages draw Word's fonts with (the aliases the document's font faces were registered under),
// read from the layout's text styles: each style names the family it was made for (its run's font) and the
// CSS font it draws with. A family takes the face most of its styles draw with (a style drawn from another of
// the run's fonts, a complex-script run's, is outvoted). Read again only when the layout made new styles.

const CSS_FAMILY = /\d(?:\.\d+)?px (.+)$/

export class LayoutFaces {
	private readonly editor: PagedEditor
	private read = 0
	private readonly votes = new Map<string, Map<string, number>>()
	private readonly chosen = new Map<string, string | undefined>()
	// Changes whenever what `faces` returns may have changed.
	version = 0
	readonly faces: CssFaces

	constructor(editor: PagedEditor) {
		this.editor = editor
		this.faces = family => {
			if (this.chosen.has(family)) return this.chosen.get(family)

			const votes = this.votes.get(family)
			let best: string | undefined
			let most = 0

			if (votes !== undefined) {
				for (const [css, count] of votes) {
					if (count > most) {
						best = css
						most = count
					}
				}
			}

			this.chosen.set(family, best)

			return best
		}
	}

	// Takes in the styles made since the last call; true when the faces may have changed.
	refresh(): boolean {
		const info = this.editor.layout.styleInfo

		if (info.length <= this.read) return false

		for (let id = this.read; id < info.length; id++) {
			const style = info[id]
			const family = style?.run?.fonts.ascii
			// The quoted aliases only: the generic family goes last, after every run font's face.
			const css =
				style === undefined
					? undefined
					: CSS_FAMILY.exec(style.css)?.[1]
							?.match(/"[^"]*"/g)
							?.join(", ")

			// Generic families only (no face of the document's): nothing to learn.
			if (family === undefined || family.length === 0 || css === undefined) continue

			let votes = this.votes.get(family)

			if (votes === undefined) {
				votes = new Map()
				this.votes.set(family, votes)
			}

			votes.set(css, (votes.get(css) ?? 0) + 1)
		}

		this.read = info.length
		this.chosen.clear()
		this.version++

		return true
	}
}
