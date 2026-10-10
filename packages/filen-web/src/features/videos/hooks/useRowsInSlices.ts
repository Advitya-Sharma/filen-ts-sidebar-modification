import { useEffect, useState } from "react"

// Rows mounted on the first render, then per idle slice: a long list shows its first screen at once and the rest
// follows in order, so tab order and find-in-page see every row once the slices are done.
export const FIRST_SLICE_ROWS = 40
export const SLICE_ROWS = 80

function whenIdle(run: () => void): () => void {
	if (typeof requestIdleCallback === "function") {
		const handle = requestIdleCallback(run, { timeout: 200 })

		return () => {
			cancelIdleCallback(handle)
		}
	}

	const timer = setTimeout(run, 16)

	return () => {
		clearTimeout(timer)
	}
}

// How many of `total` rows to mount now.
export function useRowsInSlices(total: number): number {
	const [count, setCount] = useState(() => Math.min(total, FIRST_SLICE_ROWS))

	useEffect(() => {
		if (count >= total) {
			return
		}

		return whenIdle(() => {
			setCount(Math.min(total, count + SLICE_ROWS))
		})
	}, [count, total])

	return Math.min(count, total)
}
