// One IntersectionObserver for many card elements, each with its own listener: what a row (or the grid) uses to
// tell its cards when they come near, and the page to tell them when they are actually on screen.

type Listener = (inView: boolean) => void

export class CardObserver {
	private readonly observer: IntersectionObserver
	private readonly listeners = new Map<Element, Listener>()

	public constructor({ root, rootMargin, threshold }: { root: Element | null; rootMargin: string; threshold: number }) {
		this.observer = new IntersectionObserver(
			entries => {
				for (const entry of entries) {
					this.listeners.get(entry.target)?.(entry.isIntersecting && entry.intersectionRatio >= threshold)
				}
			},
			{ root, rootMargin, threshold }
		)
	}

	public observe(element: Element, listener: Listener): () => void {
		this.listeners.set(element, listener)
		this.observer.observe(element)

		return () => {
			this.listeners.delete(element)
			this.observer.unobserve(element)
		}
	}

	public disconnect(): void {
		this.observer.disconnect()
		this.listeners.clear()
	}
}

let visibleCards: CardObserver | null = null

// Half a card on screen: what starts the dwell before a card may cost a download.
export function visibleCardObserver(): CardObserver {
	visibleCards ??= new CardObserver({ root: null, rootMargin: "0px", threshold: 0.5 })

	return visibleCards
}
