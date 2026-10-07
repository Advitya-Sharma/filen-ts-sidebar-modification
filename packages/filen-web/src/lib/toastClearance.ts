import { useSyncExternalStore } from "react"
import { createListenerSet } from "@/lib/listenerSet"
import { computeToastClearance, computeToastSideClearance } from "@/lib/toastClearance.logic"

// Single source of truth for how far toasts must lift: surfaces that can sit in the bottom-right toast
// corner (floating selection bars, the docked audio player, the chat composer) register their root
// element while mounted, and the Toaster reads the measured clearance. Nothing registered — public-link
// routes, an idle shell — reads 0, i.e. sonner's normal flush offset. Panels docked against the right edge
// (the assistant) register as SIDE obstructions instead: they push the column left rather than up.
const obstructions = new Map<Element, ResizeObserver>()
const sideObstructions = new Map<Element, ResizeObserver>()
const listeners = createListenerSet("toastClearance")
let clearance = 0
let sideClearance = 0
let frame: number | null = null

function measure(): void {
	const viewport = {
		width: window.innerWidth,
		height: window.innerHeight
	}

	const nextSide = computeToastSideClearance({
		obstructions: Array.from(sideObstructions.keys(), element => element.getBoundingClientRect()),
		viewport
	})

	const next = computeToastClearance({
		obstructions: Array.from(obstructions.keys(), element => element.getBoundingClientRect()),
		viewport,
		sideClearance: nextSide
	})

	if (next === clearance && nextSide === sideClearance) {
		return
	}

	clearance = next
	sideClearance = nextSide
	listeners.emit()
}

// Deferred to the next frame for changes that move an obstruction without resizing it — another bar
// mounting/unmounting reflows the shell, a viewport resize moves the docked player's top edge.
function scheduleMeasure(): void {
	if (frame !== null) {
		return
	}

	frame = requestAnimationFrame(() => {
		frame = null
		measure()
	})
}

function registeredCount(): number {
	return obstructions.size + sideObstructions.size
}

function register(registry: Map<Element, ResizeObserver>, element: HTMLElement): () => void {
	const observer = new ResizeObserver(measure)

	observer.observe(element)

	if (element.parentElement !== null) {
		observer.observe(element.parentElement)
	}

	if (registeredCount() === 0) {
		window.addEventListener("resize", scheduleMeasure)
	}

	registry.set(element, observer)
	scheduleMeasure()

	return () => {
		observer.disconnect()
		registry.delete(element)

		if (registeredCount() === 0) {
			window.removeEventListener("resize", scheduleMeasure)
		}

		scheduleMeasure()
	}
}

// React 19 ref callback (returns its own cleanup), so call sites need no hook: `ref={toastObstructionRef}`.
// The parent is observed too because a centered bar keeps its size while its wrapper resizes around it
// (sidebar drag, window resize), which moves it into or out of the toast column.
export function toastObstructionRef(element: HTMLElement | null): (() => void) | undefined {
	if (element === null) {
		return undefined
	}

	return register(obstructions, element)
}

// Same contract for a panel docked against the right edge; its width (a sidebar drag resizes it) is what
// moves the column.
export function toastSideObstructionRef(element: HTMLElement | null): (() => void) | undefined {
	if (element === null) {
		return undefined
	}

	return register(sideObstructions, element)
}

function getClearance(): number {
	return clearance
}

function getSideClearance(): number {
	return sideClearance
}

// Pixels between the viewport's bottom edge and the top of the highest obstruction in the toast column.
export function useToastClearance(): number {
	return useSyncExternalStore(listeners.subscribe, getClearance, () => 0)
}

// Pixels between the viewport's right edge and the left edge of the widest right-docked panel.
export function useToastSideClearance(): number {
	return useSyncExternalStore(listeners.subscribe, getSideClearance, () => 0)
}
