// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { createElement, useEffect, useRef } from "react"
import { createRoot, type Root } from "react-dom/client"
import type { RegisteredRouter } from "@tanstack/react-router"
import { navigateChecked } from "@/features/assistant/tools/navigation"
import { ASSISTANT_PANEL_ID, ASSISTANT_TOGGLE_ID } from "@/features/assistant/components/assistantFocus"

let pathname = "/drive"
let resolvedListeners: (() => void)[] = []

const router = {
	state: {
		location: {
			get pathname() {
				return pathname
			}
		}
	},
	subscribe: (_event: string, listener: () => void) => {
		resolvedListeners.push(listener)

		return () => {
			resolvedListeners = resolvedListeners.filter(l => l !== listener)
		}
	}
} as unknown as RegisteredRouter

let composer: HTMLTextAreaElement
let destination: HTMLTextAreaElement
let toggle: HTMLButtonElement

function element<K extends keyof HTMLElementTagNameMap>(tag: K, parent: HTMLElement): HTMLElementTagNameMap[K] {
	const el = document.createElement(tag)

	parent.appendChild(el)

	return el
}

// Lands on `to` and has the destination take focus the way a mount effect does: in a later task.
function goTo(
	to: string,
	steal = () => {
		destination.focus()
	}
): () => Promise<void> {
	return () => {
		pathname = to

		setTimeout(steal, 0)

		return Promise.resolve()
	}
}

function nextTask(): Promise<void> {
	return new Promise(resolve => {
		setTimeout(resolve, 0)
	})
}

beforeEach(() => {
	pathname = "/drive"
	resolvedListeners = []

	const panel = element("aside", document.body)

	panel.id = ASSISTANT_PANEL_ID
	panel.tabIndex = -1
	composer = element("textarea", panel)
	composer.setAttribute("data-assistant-autofocus", "")
	destination = element("textarea", document.body)
	toggle = element("button", document.body)
	toggle.id = ASSISTANT_TOGGLE_ID
	composer.focus()
})

afterEach(async () => {
	vi.useRealTimers()
	// Ends any hold left by a test before the next one records its own.
	toggle.focus()
	await nextTask()
	document.body.replaceChildren()
})

describe("navigateChecked keeps focus the user left in the assistant", () => {
	it("puts focus back in the composer when the destination takes it after landing", async () => {
		await navigateChecked({ router, navigate: goTo("/chats/c1"), landed: p => p === "/chats/c1" })
		await nextTask()

		expect(document.activeElement).toBe(composer)
	})

	it("puts focus back when the destination takes it before the navigation settles", async () => {
		await navigateChecked({
			router,
			navigate: () => {
				pathname = "/chats/c1"
				destination.focus()

				return Promise.resolve()
			},
			landed: p => p === "/chats/c1"
		})
		await Promise.resolve()

		expect(document.activeElement).toBe(composer)
	})

	it("keeps focus in the composer when a mounted chat thread focuses its own input", async () => {
		const host = element("div", document.body)
		const root: Root = createRoot(host)
		let threadFocused = false

		function Thread() {
			const ref = useRef<HTMLTextAreaElement>(null)

			useEffect(() => {
				ref.current?.focus()
				threadFocused = true
			}, [])

			return createElement("textarea", { ref, "aria-label": "message" })
		}

		await navigateChecked({
			router,
			navigate: () => {
				pathname = "/chats/c1"
				root.render(createElement(Thread))

				return Promise.resolve()
			},
			landed: p => p === "/chats/c1"
		})
		await vi.waitFor(() => {
			expect(threadFocused).toBe(true)
		})
		await Promise.resolve()

		expect(document.activeElement).toBe(composer)
		root.unmount()
	})

	it("moves to the panel's autofocus target when the composer it held was replaced", async () => {
		const replacement = document.createElement("textarea")

		replacement.setAttribute("data-assistant-autofocus", "")

		await navigateChecked({
			router,
			navigate: goTo("/chats/c1", () => {
				composer.replaceWith(replacement)
				destination.focus()
			}),
			landed: p => p === "/chats/c1"
		})
		await nextTask()

		expect(document.activeElement).toBe(replacement)
	})

	it("leaves focus where the user clicked", async () => {
		await navigateChecked({
			router,
			navigate: goTo("/chats/c1", () => {
				destination.dispatchEvent(new Event("pointerdown", { bubbles: true }))
				destination.focus()
			}),
			landed: p => p === "/chats/c1"
		})
		await nextTask()

		expect(document.activeElement).toBe(destination)
	})

	it("leaves focus the user tabbed out with", async () => {
		await navigateChecked({
			router,
			navigate: goTo("/chats/c1", () => {
				composer.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", bubbles: true }))
				destination.focus()
			}),
			landed: p => p === "/chats/c1"
		})
		await nextTask()

		expect(document.activeElement).toBe(destination)
	})

	it("leaves focus a key press inside the panel moved on purpose, but not one typed text", async () => {
		await navigateChecked({
			router,
			navigate: goTo("/chats/c1", () => {
				composer.dispatchEvent(new KeyboardEvent("keydown", { key: "a", bubbles: true }))
				destination.focus()
			}),
			landed: p => p === "/chats/c1"
		})
		await nextTask()

		expect(document.activeElement).toBe(composer)

		await navigateChecked({
			router,
			navigate: goTo("/chats/c2", () => {
				composer.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }))
				destination.focus()
			}),
			landed: p => p === "/chats/c2"
		})
		await nextTask()

		expect(document.activeElement).toBe(destination)
	})

	it("lets closing the panel move focus to the rail toggle", async () => {
		await navigateChecked({
			router,
			navigate: goTo("/chats/c1", () => {
				toggle.focus()
			}),
			landed: p => p === "/chats/c1"
		})
		await nextTask()

		expect(document.activeElement).toBe(toggle)
	})

	it("lets a dialog the navigation opened take focus", async () => {
		const dialog = element("div", document.body)

		dialog.setAttribute("role", "alertdialog")

		const keep = element("button", dialog)

		await navigateChecked({
			router,
			navigate: goTo("/chats/c1", () => {
				keep.focus()
			}),
			landed: p => p === "/chats/c1"
		})
		await nextTask()

		expect(document.activeElement).toBe(keep)
	})

	it("does not touch focus that was outside the assistant", async () => {
		destination.focus()

		const other = element("input", document.body)

		await navigateChecked({
			router,
			navigate: goTo("/chats/c1", () => {
				other.focus()
			}),
			landed: p => p === "/chats/c1"
		})
		await nextTask()

		expect(document.activeElement).toBe(other)
	})

	it("gives up on a destination that keeps taking focus back", async () => {
		const grabby = element("input", document.body)
		let grabs = 0

		grabby.addEventListener("blur", () => {
			grabs += 1
			queueMicrotask(() => {
				grabby.focus()
			})
		})

		await navigateChecked({
			router,
			navigate: goTo("/chats/c1", () => {
				grabby.focus()
			}),
			landed: p => p === "/chats/c1"
		})
		await vi.waitFor(() => {
			expect(document.activeElement).toBe(grabby)
		})
		await nextTask()

		expect(grabs).toBeLessThanOrEqual(3)
	})

	it("stops holding focus a while after landing", async () => {
		vi.useFakeTimers()

		const navigation = navigateChecked({ router, navigate: goTo("/chats/c1", () => undefined), landed: p => p === "/chats/c1" })

		await vi.advanceTimersByTimeAsync(0)
		await navigation
		await vi.advanceTimersByTimeAsync(5_000)
		destination.focus()
		await Promise.resolve()

		expect(document.activeElement).toBe(destination)
	})

	it("stops holding focus as soon as a navigation is refused", async () => {
		await expect(
			navigateChecked({
				router,
				navigate: () => Promise.resolve(),
				landed: p => p === "/chats/c1"
			})
		).rejects.toThrow("did not navigate")

		destination.focus()
		await Promise.resolve()

		expect(document.activeElement).toBe(destination)
	})
})
