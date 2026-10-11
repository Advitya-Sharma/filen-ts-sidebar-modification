import { useEffect, useState } from "react"
import type { DocxAnnotations } from "../../../dom/docx/annotations"
import type { DocxSessionApi } from "../../../dom/docx/protocol"
import type { DocxDocument } from "../store"

// The open document's comments and tracked changes, once every page is laid out (their boxes need the
// final layout); undefined until then, and for a session that cannot tell. `enabled`: false skips asking
// (a document whose open info counted none).
export function useDocxAnnotations(document: DocxDocument, enabled = true): DocxAnnotations | undefined {
	const session = document.session
	const ready = enabled && document.status === "ready"
	const [loaded, setLoaded] = useState<{ readonly session: DocxSessionApi; readonly annotations: DocxAnnotations } | null>(null)

	useEffect(() => {
		if (!ready || session === undefined || session.annotations === undefined) return

		let current = true

		session
			.annotations()
			.then(annotations => {
				if (current && annotations !== null) setLoaded({ session, annotations })
			})
			.catch(() => undefined)

		return () => {
			current = false
		}
	}, [ready, session])

	return loaded !== null && loaded.session === session ? loaded.annotations : undefined
}
