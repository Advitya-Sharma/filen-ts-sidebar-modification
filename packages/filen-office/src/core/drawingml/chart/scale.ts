// Value axis scaling: explicit c:scaling bounds where given, otherwise an automatic "nice" range close to
// what Office picks (zero-based when the data allows, steps of 1, 2 or 5 times a power of ten, headroom
// when the data touches the top).

export interface Scale {
	readonly min: number
	readonly max: number
	readonly major: number
	// Logarithm base, or 0 for a linear axis.
	readonly logBase: number
}

export interface ScaleRequest {
	readonly dataMin: number
	readonly dataMax: number
	readonly min: number | undefined
	readonly max: number | undefined
	readonly major: number | undefined
	readonly logBase: number
	// The most major intervals that fit the axis length.
	readonly maxIntervals: number
	// Percent-stacked data: the range is 0..1 (or -1..1) unless given.
	readonly percent: boolean
}

function niceStep(span: number, intervals: number): number {
	if (!(span > 0) || !Number.isFinite(span)) return 1

	const raw = span / Math.max(1, intervals)
	const power = 10 ** Math.floor(Math.log10(raw))

	for (const factor of [1, 2, 5, 10]) {
		if (factor * power >= raw - 1e-12 * raw) return factor * power
	}

	return 10 * power
}

function roundTo(value: number, step: number): number {
	// Snaps away float noise (0.30000000000000004) relative to the step.
	const digits = Math.max(0, -Math.floor(Math.log10(step)) + 2)

	return Number(value.toFixed(Math.min(20, digits)))
}

export function computeScale(request: ScaleRequest): Scale {
	if (request.logBase > 1) return logScale(request)

	let low = request.dataMin
	let high = request.dataMax

	if (!Number.isFinite(low) || !Number.isFinite(high)) {
		low = 0
		high = 1
	}

	if (request.percent) {
		low = low < 0 ? -1 : 0
		high = high > 0 ? 1 : 0
	} else {
		// Zero-based unless the data sits in a narrow band far from zero.
		if (low >= 0 && (high === 0 || low < (high * 5) / 6)) low = 0
		if (high <= 0 && (low === 0 || high > (low * 5) / 6)) high = 0
	}

	if (request.min !== undefined) low = request.min
	if (request.max !== undefined) high = request.max
	if (high === low) {
		if (high === 0) high = 1
		else if (high > 0) low = request.min ?? 0
		else high = request.max ?? 0
	}

	if (high < low) [low, high] = [high, low]

	const major = request.major !== undefined && request.major > 0 ? request.major : niceStep(high - low, request.maxIntervals)
	let min = request.min ?? Math.floor(low / major + 1e-9) * major
	let max = request.max ?? Math.ceil(high / major - 1e-9) * major

	// Office leaves a step of headroom when the data comes within 5% of a step of the top (or bottom).
	if (request.max === undefined && !request.percent && request.dataMax > 0 && max - request.dataMax < major * 0.05) max += major
	if (request.min === undefined && !request.percent && request.dataMin < 0 && request.dataMin - min < major * 0.05) min -= major
	if (max <= min) max = min + major

	return { min: roundTo(min, major), max: roundTo(max, major), major, logBase: 0 }
}

function logScale(request: ScaleRequest): Scale {
	const base = request.logBase
	const positiveMin = request.dataMin > 0 ? request.dataMin : 1
	const positiveMax = request.dataMax > 0 ? request.dataMax : base
	const min =
		request.min !== undefined && request.min > 0 ? request.min : base ** Math.floor(Math.log(positiveMin) / Math.log(base) + 1e-9)
	let max = request.max !== undefined && request.max > 0 ? request.max : base ** Math.ceil(Math.log(positiveMax) / Math.log(base) - 1e-9)

	if (max <= min) max = min * base

	return { min, max, major: request.major !== undefined && request.major > 1 ? request.major : base, logBase: base }
}

// Position of `value` along the scale as a fraction 0..1 (unclamped).
export function scaleFraction(scale: Scale, value: number): number {
	if (scale.logBase > 1) {
		if (!(value > 0)) return 0

		const low = Math.log(scale.min)

		return (Math.log(value) - low) / (Math.log(scale.max) - low)
	}

	return (value - scale.min) / (scale.max - scale.min)
}

// Major tick values from min to max.
export function scaleTicks(scale: Scale): number[] {
	const out: number[] = []

	if (scale.logBase > 1) {
		for (let value = scale.min; value <= scale.max * (1 + 1e-9) && out.length < 200; value *= scale.major) out.push(value)

		return out
	}

	const count = Math.min(1000, Math.round((scale.max - scale.min) / scale.major))

	for (let k = 0; k <= count; k++) out.push(roundTo(scale.min + k * scale.major, scale.major))

	return out
}
