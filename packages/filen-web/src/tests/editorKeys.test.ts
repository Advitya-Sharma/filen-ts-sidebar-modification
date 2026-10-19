import { describe, expect, it } from "vitest"
import { codeMirrorKeys } from "@/features/preview/lib/editorKeys.logic"

describe("codeMirrorKeys", () => {
	it("maps the app's combos to CodeMirror's notation, alternatives and all", () => {
		expect(codeMirrorKeys("mod+shift+x")).toEqual(["Mod-Shift-x"])
		expect(codeMirrorKeys("mod+alt+f")).toEqual(["Mod-Alt-f"])
		expect(codeMirrorKeys("delete,backspace")).toEqual(["Delete", "Backspace"])
		expect(codeMirrorKeys("ctrl+slash")).toEqual(["Ctrl-/"])
		expect(codeMirrorKeys("")).toEqual([])
	})

	it("maps the recorder's event.code tokens to KeyboardEvent.key names", () => {
		const tokens = [
			["arrowup", "ArrowUp"],
			["arrowdown", "ArrowDown"],
			["arrowleft", "ArrowLeft"],
			["arrowright", "ArrowRight"],
			["pageup", "PageUp"],
			["pagedown", "PageDown"],
			["home", "Home"],
			["end", "End"],
			["bracketleft", "["],
			["bracketright", "]"],
			["minus", "-"],
			["equal", "="],
			["semicolon", ";"],
			["quote", "'"],
			["backquote", "`"],
			["backslash", "\\"],
			["comma", ","],
			["period", "."],
			["slash", "/"],
			["space", "Space"],
			["enter", "Enter"],
			["escape", "Escape"],
			["tab", "Tab"],
			["backspace", "Backspace"],
			["delete", "Delete"],
			["f1", "F1"],
			["f12", "F12"],
			["1", "1"],
			["a", "a"]
		] as const

		for (const [token, key] of tokens) {
			expect(codeMirrorKeys(`mod+${token}`), token).toEqual([`Mod-${key}`])
		}
	})

	it("maps react-hotkeys-hook's aliases and the numpad operators", () => {
		expect(codeMirrorKeys("esc,return,up,shift+down,mod+add,mod+subtract")).toEqual([
			"Escape",
			"Enter",
			"ArrowUp",
			"Shift-ArrowDown",
			"Mod-+",
			"Mod--"
		])
	})
})
