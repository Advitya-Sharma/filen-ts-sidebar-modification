import { defineSheetFeature } from "../registry"
import { textFeature, type TextSheetDocument } from "../text"
import { entryCells, fillPlan } from "./fill"

// Fill over a text file: each target takes its source's text as it is.

export const textFillFeature = defineSheetFeature<TextSheetDocument>({
	id: "fill",
	commands: {
		"fill.copy": {
			info: { write: true },
			run: (context, command) => {
				const plan = fillPlan(command.ranges, command.direction)

				if (plan === null) return null
				if (plan === "tooLarge") return { refused: { reason: "tooLarge" } }

				const inputs = Array.from(plan.rows, (_, i) => context.doc.text(plan.sourceRows[i] ?? 0, plan.sourceCols[i] ?? 0))
				const set = textFeature.commands?.["cells.set"]

				return set === undefined
					? { refused: { reason: "unsupported" } }
					: set.run(context, {
							type: "cells.set",
							sheet: command.sheet,
							cells: { rows: plan.rows, cols: Uint16Array.from(plan.cols), inputs }
						})
			}
		},
		"fill.entry": {
			info: { write: true },
			run: (context, command) => {
				// A text file merges nothing.
				const cells = entryCells(command.ranges, new Uint32Array(0), command.entry, command.from)
				const set = textFeature.commands?.["cells.set"]

				if (cells === null) return null
				if (cells === "tooLarge") return { refused: { reason: "tooLarge" } }

				return set === undefined
					? { refused: { reason: "unsupported" } }
					: set.run(context, { type: "cells.set", sheet: command.sheet, cells })
			}
		}
	}
})
