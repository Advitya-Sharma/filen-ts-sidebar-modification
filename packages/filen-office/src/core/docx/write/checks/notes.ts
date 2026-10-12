import { DocxVerifyError } from "../failure"
import type { DocxCheck } from "./index"

// Every new note reference resolves to a normal note (ids below 1 are Word's separators). Gate 2 counts
// note and reference ids itself. To come with note writing: separators present.
export const noteChecks: readonly DocxCheck[] = [
	{
		done(input) {
			for (const kind of ["footnote", "endnote"]) {
				for (const id of input.added(`${kind}Reference`)) {
					if (!/^\d+$/.test(id) || Number(id) < 1 || input.ids(`${kind}:${id}`) < 1)
						throw new DocxVerifyError("danglingReference", `A new reference to ${kind} ${id} has no note`)
				}
			}
		}
	}
]
