import { useEffect, useRef } from "react"
import { useNavigate } from "@tanstack/react-router"
import { whenTitleKnown } from "@/features/videos/lib/followTitle"

// Follows a corrected item to the title it now belongs to, once the index has it. A save that resolves after the
// page went (its dialog closes while saving) follows nowhere, and a pending follow is dropped when the page goes.
export function useFollowMovedTitle(titleId: string): (id: string) => void {
	const navigate = useNavigate()
	const following = useRef<(() => void) | null>(null)
	const live = useRef(false)

	useEffect(() => {
		live.current = true

		return () => {
			live.current = false
			following.current?.()
			following.current = null
		}
	}, [titleId])

	return (id: string): void => {
		if (id === titleId || !live.current) {
			return
		}

		following.current?.()
		following.current = whenTitleKnown(id, () => {
			following.current = null
			void navigate({ to: "/videos/title/$uuid", params: { uuid: id }, replace: true })
		})
	}
}
