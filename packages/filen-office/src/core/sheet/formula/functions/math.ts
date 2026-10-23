import { FunctionArgumentType, type ImplementedFunctions } from "hyperformula"
import { divError, numError, SheetFunctionPlugin, type PluginCall, type PluginState } from "./plugin"

// The rounding functions as Excel calculates them, where the engine's own differ: INT, MOD, ROUND, ROUNDUP,
// ROUNDDOWN, TRUNC, ISEVEN, ISODD and the CEILING and FLOOR family. Each case is the function's page on
// support.microsoft.com and its ECMA-376 Part 1 §18.17.7 entry:
// - INT rounds down, toward minus infinity: INT(-8.9) is -9 (the engine truncated it to -8).
// - MOD's result has the divisor's sign: MOD(-3,2) is 1 (the engine gave JavaScript's remainder, -1); a multiple of
//   the divisor leaves 0 (below).
// - The ROUND family rounds decimal digits ("digits 0–4 round down, while digits 5–9 round up", §18.17.7.278) of the
//   number at the 15 significant digits Excel holds ("Floating-point arithmetic may give inaccurate results in
//   Excel"): ROUND(1.005,2) is 1.01, ROUNDDOWN(0.29,2) 0.29 and ROUNDDOWN(8.94-8,2) 0.94 as LibreOffice 7.5 gives them,
//   where the binary values are just below (the engine gave 1, 0.28 and 0.93). num_digits is truncated:
//   ROUND(1.234,1.9) rounds to one place.
// - ISEVEN and ISODD truncate their number (§18.17.7.177, §18.17.7.183): ISODD(-15) is TRUE.
// - The CEILING and FLOOR family keeps a number that is a multiple of the significance as it is ("CEILING
//   function": CEILING(1.5,0.1) is 1.5): the quotient is read at the 15 significant digits Excel holds before it is
//   rounded, so CEILING(0.07,0.01) is 0.07 and FLOOR(0.3,0.1) 0.3 where the binary quotients are a step off the whole
//   number (the engine gave 0.08 and 0.2), as LibreOffice 7.5 gives them.
// - CEILING with a significance of 0 is 0, as LibreOffice gives it (the engine's #DIV/0!; the page is silent);
//   FLOOR's is #DIV/0!.
// - CEILING.MATH and FLOOR.MATH take any mode but 0 to reverse the direction for negative numbers:
//   CEILING.MATH(-5.5,2,-1) is -6 (the engine reversed only for 1).
// MROUND, EVEN and ODD stay the engine's: MROUND rounds the binary quotient to the nearest whole number, so a
// multiple is kept, and its page has MROUND(6.05,0.1) as 6 (6.05 is just below it), as LibreOffice gives it.

const NUMBER = { argumentType: FunctionArgumentType.NUMBER }
const DIGITS = { argumentType: FunctionArgumentType.NUMBER, defaultValue: 0 }
const SIGNIFICANCE = { argumentType: FunctionArgumentType.NUMBER, defaultValue: 1 }

const DOWN = 0
const NEAREST = 1
const UP = 2

type Direction = typeof DOWN | typeof NEAREST | typeof UP

// A number at the 15 significant digits Excel holds of it. Whole numbers are exact already.
function significant(value: number): number {
	return Number.isInteger(value) || !Number.isFinite(value) ? value : Number(value.toPrecision(15))
}

// The largest power of ten a double holds (10^308); past it there is nothing to round.
const MAX_PLACES = 308

// Four units in the last place of a double, relative (2^-50).
const ULPS = 2 ** -50

// x rounded to `digits` decimal places (left of the point when negative), down toward zero, to the nearest with
// halves away from zero, or up away from zero.
function roundDigits(x: number, digits: number, direction: Direction): number {
	const places = Math.trunc(digits)

	if (x === 0 || !Number.isFinite(x) || places > MAX_PLACES) return x
	if (places < -MAX_PLACES) return 0

	const scale = 10 ** Math.abs(places)
	const held = Math.abs(significant(x))
	const scaled = significant(places >= 0 ? held * scale : held / scale)

	if (!Number.isFinite(scaled)) return x

	const whole = direction === DOWN ? Math.floor(scaled) : direction === UP ? Math.ceil(scaled) : Math.floor(scaled + 0.5)

	if (whole === 0) return 0

	const signed = x < 0 ? -whole : whole

	return places >= 0 ? signed / scale : signed * scale
}

// x rounded to a multiple of `significance` (whose sign the caller has settled), down or up.
function toMultiple(x: number, significance: number, up: boolean): number {
	if (x === 0 || significance === 0) return 0

	const quotient = significant(x / significance)
	const multiple = (up ? Math.ceil(quotient) : Math.floor(quotient)) * significance

	return multiple === 0 ? 0 : multiple
}

export class MathFunctions extends SheetFunctionPlugin {
	static override implementedFunctions: ImplementedFunctions = {
		INT: { method: "int", parameters: [NUMBER] },
		MOD: { method: "mod", parameters: [NUMBER, NUMBER] },
		ROUND: { method: "round", parameters: [NUMBER, DIGITS] },
		ROUNDUP: { method: "roundUp", parameters: [NUMBER, DIGITS] },
		ROUNDDOWN: { method: "roundDown", parameters: [NUMBER, DIGITS] },
		// TRUNC is ROUNDDOWN (the same arguments).
		TRUNC: { method: "roundDown", parameters: [NUMBER, DIGITS] },
		ISEVEN: { method: "isEven", parameters: [NUMBER] },
		ISODD: { method: "isOdd", parameters: [NUMBER] },
		CEILING: { method: "ceiling", parameters: [NUMBER, NUMBER] },
		"CEILING.MATH": { method: "ceilingMath", parameters: [NUMBER, SIGNIFICANCE, DIGITS] },
		"FLOOR.MATH": { method: "floorMath", parameters: [NUMBER, SIGNIFICANCE, DIGITS] },
		FLOOR: { method: "floor", parameters: [NUMBER, NUMBER] },
		"CEILING.PRECISE": { method: "ceilingPrecise", parameters: [NUMBER, SIGNIFICANCE] },
		// ISO.CEILING is CEILING.PRECISE (the same arguments).
		"ISO.CEILING": { method: "ceilingPrecise", parameters: [NUMBER, SIGNIFICANCE] },
		"FLOOR.PRECISE": { method: "floorPrecise", parameters: [NUMBER, SIGNIFICANCE] }
	}

	int(ast: PluginCall, state: PluginState): unknown {
		return this.runFunction(ast.args, state, this.metadata("INT"), (x: number) => Math.floor(x))
	}

	mod(ast: PluginCall, state: PluginState): unknown {
		return this.runFunction(ast.args, state, this.metadata("MOD"), (n: number, d: number) => {
			if (d === 0) return divError()

			const remainder = n % d
			const signed = remainder < 0 !== d < 0 && remainder !== 0 ? remainder + d : remainder
			// A remainder within a few units of the number's last binary place of 0 or of the divisor is the rounding
			// of a multiple: MOD(6.3,2.1) is 0, as LibreOffice gives it, where the binary remainder is 2.1 less 4e-16.
			const noise = Math.abs(n) * ULPS

			return Math.abs(signed) <= noise || Math.abs(d - signed) <= noise ? 0 : signed
		})
	}

	round(ast: PluginCall, state: PluginState): unknown {
		return this.runFunction(ast.args, state, this.metadata("ROUND"), (x: number, digits: number) => roundDigits(x, digits, NEAREST))
	}

	roundUp(ast: PluginCall, state: PluginState): unknown {
		return this.runFunction(ast.args, state, this.metadata("ROUNDUP"), (x: number, digits: number) => roundDigits(x, digits, UP))
	}

	roundDown(ast: PluginCall, state: PluginState): unknown {
		return this.runFunction(ast.args, state, this.metadata("ROUNDDOWN"), (x: number, digits: number) => roundDigits(x, digits, DOWN))
	}

	isEven(ast: PluginCall, state: PluginState): unknown {
		return this.runFunction(ast.args, state, this.metadata("ISEVEN"), (x: number) => Math.trunc(x) % 2 === 0)
	}

	isOdd(ast: PluginCall, state: PluginState): unknown {
		return this.runFunction(ast.args, state, this.metadata("ISODD"), (x: number) => Math.abs(Math.trunc(x) % 2) === 1)
	}

	// CEILING(number, significance): away from zero; a negative number with a positive significance toward it; a
	// positive number with a negative significance is #NUM!.
	ceiling(ast: PluginCall, state: PluginState): unknown {
		return this.runFunction(ast.args, state, this.metadata("CEILING"), (x: number, significance: number) =>
			x > 0 && significance < 0 ? numError() : toMultiple(x, significance, true)
		)
	}

	ceilingMath(ast: PluginCall, state: PluginState): unknown {
		return this.runFunction(ast.args, state, this.metadata("CEILING.MATH"), (x: number, significance: number, mode: number) => {
			const size = Math.abs(significance)

			return toMultiple(x, mode !== 0 && x < 0 ? -size : size, true)
		})
	}

	floorMath(ast: PluginCall, state: PluginState): unknown {
		return this.runFunction(ast.args, state, this.metadata("FLOOR.MATH"), (x: number, significance: number, mode: number) => {
			const size = Math.abs(significance)

			return toMultiple(x, mode !== 0 && x < 0 ? -size : size, false)
		})
	}

	// FLOOR(number, significance): toward zero; a negative number with a positive significance away from it; a
	// positive number with a negative significance is #NUM!, a significance of 0 #DIV/0!.
	floor(ast: PluginCall, state: PluginState): unknown {
		return this.runFunction(ast.args, state, this.metadata("FLOOR"), (x: number, significance: number) =>
			significance === 0 ? divError() : x > 0 && significance < 0 ? numError() : toMultiple(x, significance, false)
		)
	}

	// CEILING.PRECISE and FLOOR.PRECISE round up (down) whatever the signs: the significance's absolute value
	// ("ISO.CEILING function": ISO.CEILING(4.3,-2) is 6).
	ceilingPrecise(ast: PluginCall, state: PluginState): unknown {
		return this.runFunction(ast.args, state, this.metadata("CEILING.PRECISE"), (x: number, significance: number) =>
			toMultiple(x, Math.abs(significance), true)
		)
	}

	floorPrecise(ast: PluginCall, state: PluginState): unknown {
		return this.runFunction(ast.args, state, this.metadata("FLOOR.PRECISE"), (x: number, significance: number) =>
			toMultiple(x, Math.abs(significance), false)
		)
	}
}
