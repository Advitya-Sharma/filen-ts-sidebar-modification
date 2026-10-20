// The versions this tab's own saves made, by uuid, kept across a reload of the tab (sessionStorage): a
// preview that later finds one of them newer than what it opened on (a listing restored from disk at boot)
// shows it as the user's own save, not as a change made elsewhere. Only the latest few: what a listing read
// races is recent. Best-effort: storage failing means an own save may be announced as made elsewhere.
const OWN_SAVES_KEY = "filen.preview.ownSaves"
const OWN_SAVES_KEPT = 32

function read(): string[] {
	try {
		const raw = sessionStorage.getItem(OWN_SAVES_KEY)
		const parsed: unknown = raw === null ? [] : JSON.parse(raw)

		return Array.isArray(parsed) ? parsed.filter((uuid): uuid is string => typeof uuid === "string") : []
	} catch {
		return []
	}
}

export function recordOwnSave(uuid: string): void {
	const kept = read().filter(saved => saved !== uuid)

	kept.push(uuid)

	try {
		sessionStorage.setItem(OWN_SAVES_KEY, JSON.stringify(kept.slice(-OWN_SAVES_KEPT)))
	} catch {
		// Not kept: the save may later be announced as made elsewhere.
	}
}

export function isOwnSave(uuid: string): boolean {
	return read().includes(uuid)
}
