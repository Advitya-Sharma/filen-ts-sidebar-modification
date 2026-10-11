// Layout works in CSS px at 96 dpi; OOXML stores twips (1/20 pt), EMU (1/914400 in), half-points and
// eighths of a point.

export function twips(value: number): number {
	return value / 15
}

export function emu(value: number): number {
	return value / 9525
}

export function halfPoints(value: number): number {
	return (value * 2) / 3
}

export function points(value: number): number {
	return (value * 4) / 3
}

export function eighthPoints(value: number): number {
	return value / 6
}

export const EPSILON = 1e-6
