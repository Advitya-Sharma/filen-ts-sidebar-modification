import { DocxVerifyError } from "../failure"
import type { DocxCheck } from "./index"

// Every new comment reference resolves to a comment, and a new comment takes no id the source's ranges or
// references use without a comment (it would take them over). Gate 2 counts comment, range and reference
// ids itself. To come with comment writing: the comments part family consistent.
export const commentChecks: readonly DocxCheck[] = [
	{
		done(input) {
			for (const id of input.added("commentReference")) {
				if (input.ids(`comment:${id}`) < 1)
					throw new DocxVerifyError("danglingReference", `A new reference to comment ${id} has no comment`)
			}

			for (const id of input.added("comment")) {
				if (
					input.source(`comment:${id}`) === 0 &&
					(input.source(`commentRangeStart:${id}`) > 0 || input.source(`commentReference:${id}`) > 0)
				)
					throw new DocxVerifyError("duplicateId", `A new comment ${id} takes the id of a range or reference without a comment`)
			}
		}
	}
]
