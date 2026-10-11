import type { SpellcheckRequest } from "@filen/spellcheck"
import { serveSpellcheckWorker } from "@filen/spellcheck/worker"

// One spellchecker engine (wasm) and the dictionaries the page asked for (lib/spellcheck.ts runs two of
// these: checks, and suggestions it may terminate). Dictionaries come from the app's own origin and are
// kept in Cache Storage.
serveSpellcheckWorker({
	post: (message, transfer) => {
		self.postMessage(message, { transfer: [...transfer] })
	},
	listen: onMessage => {
		self.addEventListener("message", (event: MessageEvent<SpellcheckRequest>) => {
			onMessage(event.data)
		})
	}
})
