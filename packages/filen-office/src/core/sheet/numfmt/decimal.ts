// A non-negative finite number as the decimal digits a format shows: at most 15 significant digits (what a cell
// holds), then rounded half away from zero on those digits, so 1.005 shows 1.01 at two decimals as Excel shows it
// (a binary rounding would give 1.00).
export interface Digits {
	// Significant digits without leading or trailing zeros; "" for zero.
	digits: string
	// The power of ten of the first digit (1234.5: 3; 0.05: -2).
	point: number
}

const scratch: Digits = { digits: "", point: 0 }

// Decomposes |value| into `out` (a shared scratch object unless one is passed).
export function decompose(value: number, out: Digits = scratch): Digits {
	const absolute = value < 0 ? -value : value

	if (absolute === 0 || !Number.isFinite(absolute)) {
		out.digits = ""
		out.point = 0

		return out
	}

	// The shortest text that reads back as the number is those digits when it has at most 15 of them (the 15-digit
	// rounding of the number then lands on it); otherwise round to 15. Either is "ddd.ddd" or "d.ddde±x".
	let text = String(absolute)

	if (significantDigits(text) > 15) text = absolute.toPrecision(15)

	const e = text.indexOf("e")
	const mantissa = e < 0 ? text : text.slice(0, e)
	const dot = mantissa.indexOf(".")
	let digits = dot < 0 ? mantissa : mantissa.slice(0, dot) + mantissa.slice(dot + 1)
	let point = (dot < 0 ? mantissa.length : dot) - 1 + (e < 0 ? 0 : Number(text.slice(e + 1)))
	let start = 0

	while (start < digits.length - 1 && digits.charCodeAt(start) === 48) start++

	point -= start

	let end = digits.length

	while (end > start && digits.charCodeAt(end - 1) === 48) end--

	digits = digits.slice(start, end)
	out.digits = digits
	out.point = digits === "" ? 0 : point

	return out
}

// Digits from the first non-zero one to the last non-zero one of a number's text (before any exponent).
function significantDigits(text: string): number {
	let first = -1
	let last = -1
	let count = 0

	for (let i = 0; i < text.length; i++) {
		const c = text.charCodeAt(i)

		if (c === 101) break

		if (c >= 48 && c <= 57) {
			if (c !== 48) {
				if (first < 0) first = count
				last = count
			}

			count++
		}
	}

	return first < 0 ? 0 : last - first + 1
}

// Rounds `d` in place to `keep` significant digits (half away from zero); keep <= 0 rounds below the first digit.
export function roundDigits(d: Digits, keep: number): Digits {
	const { digits } = d

	if (keep >= digits.length) return d

	if (keep < 0 || (keep === 0 && digits.charCodeAt(0) < 53)) {
		d.digits = ""
		d.point = 0

		return d
	}

	let kept = digits.slice(0, keep)

	if (digits.charCodeAt(keep) >= 53) {
		let i = kept.length - 1

		while (i >= 0 && kept.charCodeAt(i) === 57) i--

		if (i < 0) {
			kept = "1"
			d.point += 1
		} else kept = kept.slice(0, i) + String.fromCharCode(kept.charCodeAt(i) + 1)
	}

	let end = kept.length

	while (end > 0 && kept.charCodeAt(end - 1) === 48) end--

	d.digits = kept.slice(0, end)

	if (d.digits === "") d.point = 0

	return d
}

// The integer part's digits ("" for zero).
export function integerDigits(d: Digits): string {
	if (d.digits === "" || d.point < 0) return ""

	const count = d.point + 1

	return d.digits.length >= count ? d.digits.slice(0, count) : d.digits + "0".repeat(count - d.digits.length)
}

// The fraction's first `count` digits, zero-padded.
export function fractionDigits(d: Digits, count: number): string {
	let out = ""

	for (let i = 0; i < count; i++) {
		const at = d.point + 1 + i

		out += at >= 0 && at < d.digits.length ? d.digits[at] : "0"
	}

	return out
}
