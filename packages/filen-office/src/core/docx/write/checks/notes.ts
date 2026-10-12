import { DocxVerifyError } from "../failure"
import type { DocxCheck } from "./index"

// Notes hold together: every new note reference resolves to a normal note (ids below 1 are Word's separators)
// and is the only reference to it, and the separator notes a notes part has stay (never touched). Gate 2 counts note and reference ids itself (verify.ts).
export const noteChecks: readonly DocxCheck[] = [
	{
		done(input) {
			for (const kind of ["footnote", "endnote"]) {
				const added = input.added(`${kind}Reference`)

				for (const id of added) {
					if (!/^\d+$/.test(id) || Number(id) < 1 || input.ids(`${kind}:${id}`) < 1)
						throw new DocxVerifyError("danglingReference", `A new reference to ${kind} ${id} has no note`)
					if (input.ids(`${kind}Reference:${id}`) > 1)
						throw new DocxVerifyError("duplicateId", `${kind} ${id} has two references`)
				}

				for (const separator of ["-1", "0"]) {
					const key = `${kind}:${separator}`

					if (input.ids(key) < Math.min(1, input.source(key)))
						throw new DocxVerifyError("content", `The save removed ${kind} separator ${separator}`)
				}
			}
		}
	}
]
