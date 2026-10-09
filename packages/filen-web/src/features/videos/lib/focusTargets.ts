// Elements focus moves to when what had it goes away (the last library directory removed).
export const VIDEOS_FIRST_RUN_ACTION_ID = "videos-first-run-action"
export const VIDEOS_HEADING_ID = "videos-heading"

// The first-run button when the home shows it, else the screen's heading; true leaves Base UI's default.
export function videosFocusFallback(): HTMLElement | true {
	return document.getElementById(VIDEOS_FIRST_RUN_ACTION_ID) ?? document.getElementById(VIDEOS_HEADING_ID) ?? true
}
