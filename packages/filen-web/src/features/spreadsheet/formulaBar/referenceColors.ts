// The colours of the references of a formula being typed, one per distinct reference in the order they appear, as
// Excel colours them: in the formula's text and as outlines on the cells they name. Two sets with enough contrast
// against the light and the dark grid.
const LIGHT = ["#1f5fd6", "#16843f", "#8a3fd1", "#c62f2f", "#0b7f8a", "#b45309", "#b8327f", "#5b6b00"] as const
const DARK = ["#7aa7ff", "#5fd08a", "#c79bff", "#ff8a80", "#4fd3df", "#f6b25a", "#ff8ccf", "#c4d65a"] as const

export const REFERENCE_COLORS = LIGHT.length

export function referenceColors(dark: boolean): readonly string[] {
	return dark ? DARK : LIGHT
}
