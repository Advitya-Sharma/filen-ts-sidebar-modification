import { refusedForAuthor, trackingRepair } from "../../review/track"
import type { InvariantStage } from "./index"

// Tracked changes: an edit recorded while tracking waits for the author's name (refused until there is one),
// and the marks of what it recorded (insertions, deletions, format and paragraph changes) follow it in the
// same undo step (see review/track.ts). Deleted text staying read-only is a core rule.
export const TRACKING_STAGE: InvariantStage | null = {
	name: "tracking",
	violation: tr => (refusedForAuthor(tr) ? "authorNeeded" : undefined),
	repair: (transactions, state) => trackingRepair(transactions, state)
}
