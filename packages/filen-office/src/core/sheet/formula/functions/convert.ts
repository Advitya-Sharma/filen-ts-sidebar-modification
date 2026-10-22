import { CellError, FunctionArgumentType, type ImplementedFunctions } from "hyperformula"
import { naError, SheetFunctionPlugin, type PluginCall, type PluginState } from "./plugin"

// CONVERT(number, from_unit, to_unit) over the units, metric prefixes and binary prefixes its page lists
// (support.microsoft.com "CONVERT function"), each unit by its exact definition: the inch 0.0254 m, the pound
// 0.45359237 kg, the US gallon 231 in³, the gravity of a pound-force 9.80665 m/s², the IT calorie 4.1868 J, the
// Julian year 365.25 days. Units and prefixes are case-sensitive; an unknown unit, or units of two groups, is #N/A.

const Group = {
	MASS: 0,
	DISTANCE: 1,
	TIME: 2,
	PRESSURE: 3,
	FORCE: 4,
	ENERGY: 5,
	POWER: 6,
	MAGNETISM: 7,
	TEMPERATURE: 8,
	VOLUME: 9,
	AREA: 10,
	INFORMATION: 11,
	SPEED: 12
} as const

type Group = (typeof Group)[keyof typeof Group]

interface Unit {
	readonly group: Group
	// The unit in the group's base (g, m, s, Pa, N, J, W, T, K, m³, m², bit, m/s); a temperature's scale is its
	// degree in kelvin.
	readonly factor: number
	// Takes a metric prefix (and, for information, a binary one).
	readonly prefixed: boolean
	// The power a prefix applies at: 2 for a squared length (m2), 3 for a cubed one (m3).
	readonly power: number
}

const INCH = 0.0254
const FOOT = 12 * INCH
const YARD = 3 * FOOT
const MILE = 5280 * FOOT
const NAUTICAL_MILE = 1852
const SURVEY_MILE = (5280 * 1200) / 3937
const LIGHT_YEAR = 299_792_458 * 365.25 * 86_400
// 648000/π astronomical units (IAU 2012, 2015).
const PARSEC = (648_000 / Math.PI) * 149_597_870_700
const PICA_POINT = INCH / 72
const PICA = INCH / 6
const POUND = 453.59237
const GRAVITY = 9.80665
const POUND_FORCE = (POUND / 1000) * GRAVITY
const ATMOSPHERE = 101_325
const GALLON = 231 * INCH ** 3
const FLUID_OUNCE = GALLON / 128
const UK_GALLON = 4.54609e-3
const HORSEPOWER = 550 * FOOT * POUND_FORCE

// [names, group, factor, prefixed]: names an alias list.
const UNITS: readonly (readonly [string, Group, number, boolean])[] = [
	["g", Group.MASS, 1, true],
	["sg", Group.MASS, (POUND * GRAVITY) / FOOT, false],
	["lbm", Group.MASS, POUND, false],
	["u", Group.MASS, 1.6605390666e-24, true],
	["ozm", Group.MASS, POUND / 16, false],
	["grain", Group.MASS, POUND / 7000, false],
	["cwt shweight", Group.MASS, 100 * POUND, false],
	["uk_cwt lcwt hweight", Group.MASS, 112 * POUND, false],
	["stone", Group.MASS, 14 * POUND, false],
	["ton", Group.MASS, 2000 * POUND, false],
	["uk_ton LTON brton", Group.MASS, 2240 * POUND, false],

	["m", Group.DISTANCE, 1, true],
	["mi", Group.DISTANCE, MILE, false],
	["Nmi", Group.DISTANCE, NAUTICAL_MILE, false],
	["in", Group.DISTANCE, INCH, false],
	["ft", Group.DISTANCE, FOOT, false],
	["yd", Group.DISTANCE, YARD, false],
	["ang", Group.DISTANCE, 1e-10, true],
	["ell", Group.DISTANCE, 45 * INCH, false],
	["ly", Group.DISTANCE, LIGHT_YEAR, false],
	["parsec pc", Group.DISTANCE, PARSEC, false],
	["Picapt Pica", Group.DISTANCE, PICA_POINT, false],
	["pica", Group.DISTANCE, PICA, false],
	["survey_mi", Group.DISTANCE, SURVEY_MILE, false],

	["yr", Group.TIME, 365.25 * 86_400, false],
	["day d", Group.TIME, 86_400, false],
	["hr", Group.TIME, 3600, false],
	["mn min", Group.TIME, 60, false],
	["sec s", Group.TIME, 1, true],

	["Pa p", Group.PRESSURE, 1, true],
	["atm at", Group.PRESSURE, ATMOSPHERE, true],
	["mmHg", Group.PRESSURE, ATMOSPHERE / 760, true],
	["psi", Group.PRESSURE, POUND_FORCE / INCH ** 2, false],
	["Torr", Group.PRESSURE, ATMOSPHERE / 760, false],

	["N", Group.FORCE, 1, true],
	["dyn dy", Group.FORCE, 1e-5, true],
	["lbf", Group.FORCE, POUND_FORCE, false],

	["J", Group.ENERGY, 1, true],
	["e", Group.ENERGY, 1e-7, true],
	["c", Group.ENERGY, 4.184, true],
	["cal", Group.ENERGY, 4.1868, true],
	["eV ev", Group.ENERGY, 1.602176634e-19, true],
	["HPh hh", Group.ENERGY, HORSEPOWER * 3600, false],
	["Wh wh", Group.ENERGY, 3600, true],
	["flb", Group.ENERGY, FOOT * POUND_FORCE, false],
	// The IT British thermal unit: 1055.05585262 J.
	["BTU btu", Group.ENERGY, 1055.05585262, false],

	["HP h", Group.POWER, HORSEPOWER, false],
	["PS", Group.POWER, 75 * GRAVITY, false],
	["W w", Group.POWER, 1, true],

	["T", Group.MAGNETISM, 1, true],
	["ga", Group.MAGNETISM, 1e-4, true],

	["C cel", Group.TEMPERATURE, 1, true],
	["F fah", Group.TEMPERATURE, 5 / 9, false],
	["K kel", Group.TEMPERATURE, 1, true],
	["Rank", Group.TEMPERATURE, 5 / 9, false],
	["Reau", Group.TEMPERATURE, 1.25, false],

	["tsp", Group.VOLUME, FLUID_OUNCE / 6, false],
	["tspm", Group.VOLUME, 5e-6, false],
	["tbs", Group.VOLUME, FLUID_OUNCE / 2, false],
	["oz", Group.VOLUME, FLUID_OUNCE, false],
	["cup", Group.VOLUME, 8 * FLUID_OUNCE, false],
	["pt us_pt", Group.VOLUME, 16 * FLUID_OUNCE, false],
	["uk_pt", Group.VOLUME, UK_GALLON / 8, false],
	["qt", Group.VOLUME, 32 * FLUID_OUNCE, false],
	["uk_qt", Group.VOLUME, UK_GALLON / 4, false],
	["gal", Group.VOLUME, GALLON, false],
	["uk_gal", Group.VOLUME, UK_GALLON, false],
	["l L lt", Group.VOLUME, 1e-3, true],
	["ang3 ang^3", Group.VOLUME, 1e-30, true],
	["barrel", Group.VOLUME, 42 * GALLON, false],
	// 2150.42 cubic inches.
	["bushel", Group.VOLUME, 2150.42 * INCH ** 3, false],
	["ft3 ft^3", Group.VOLUME, FOOT ** 3, false],
	["in3 in^3", Group.VOLUME, INCH ** 3, false],
	["ly3 ly^3", Group.VOLUME, LIGHT_YEAR ** 3, false],
	["m3 m^3", Group.VOLUME, 1, true],
	["mi3 mi^3", Group.VOLUME, MILE ** 3, false],
	["yd3 yd^3", Group.VOLUME, YARD ** 3, false],
	["Nmi3 Nmi^3", Group.VOLUME, NAUTICAL_MILE ** 3, false],
	["Picapt3 Picapt^3 Pica3 Pica^3", Group.VOLUME, PICA_POINT ** 3, false],
	["GRT regton", Group.VOLUME, 100 * FOOT ** 3, false],
	["MTON", Group.VOLUME, 40 * FOOT ** 3, false],

	["uk_acre", Group.AREA, 43_560 * FOOT ** 2, false],
	["us_acre", Group.AREA, 43_560 * (SURVEY_MILE / 5280) ** 2, false],
	["ang2 ang^2", Group.AREA, 1e-20, true],
	["ar", Group.AREA, 100, true],
	["ft2 ft^2", Group.AREA, FOOT ** 2, false],
	["ha", Group.AREA, 10_000, false],
	["in2 in^2", Group.AREA, INCH ** 2, false],
	["ly2 ly^2", Group.AREA, LIGHT_YEAR ** 2, false],
	["m2 m^2", Group.AREA, 1, true],
	["Morgen", Group.AREA, 2500, false],
	["mi2 mi^2", Group.AREA, MILE ** 2, false],
	["Nmi2 Nmi^2", Group.AREA, NAUTICAL_MILE ** 2, false],
	["Picapt2 Pica2 Pica^2 Picapt^2", Group.AREA, PICA_POINT ** 2, false],
	["yd2 yd^2", Group.AREA, YARD ** 2, false],

	["bit", Group.INFORMATION, 1, true],
	["byte", Group.INFORMATION, 8, true],

	["admkn", Group.SPEED, (6080 * FOOT) / 3600, false],
	["kn", Group.SPEED, NAUTICAL_MILE / 3600, false],
	["m/h m/hr", Group.SPEED, 1 / 3600, true],
	["m/s m/sec", Group.SPEED, 1, true],
	["mph", Group.SPEED, MILE / 3600, false]
]

const BY_NAME = new Map<string, Unit>()

for (const [names, group, factor, prefixed] of UNITS) {
	// A prefix on a squared or cubed unit (km2, cm3) scales each of its lengths.
	const power = !prefixed
		? 1
		: names.startsWith("m2") || names.startsWith("ang2")
			? 2
			: names.startsWith("m3") || names.startsWith("ang3")
				? 3
				: 1
	const unit: Unit = { group, factor, prefixed, power }

	for (const name of names.split(" ")) BY_NAME.set(name, unit)
}

const METRIC: ReadonlyMap<string, number> = new Map([
	["Y", 1e24],
	["Z", 1e21],
	["E", 1e18],
	["P", 1e15],
	["T", 1e12],
	["G", 1e9],
	["M", 1e6],
	["k", 1e3],
	["h", 1e2],
	["da", 1e1],
	["e", 1e1],
	["d", 1e-1],
	["c", 1e-2],
	["m", 1e-3],
	["u", 1e-6],
	["n", 1e-9],
	["p", 1e-12],
	["f", 1e-15],
	["a", 1e-18],
	["z", 1e-21],
	["y", 1e-24]
])

const BINARY: ReadonlyMap<string, number> = new Map([
	["Yi", 2 ** 80],
	["Zi", 2 ** 70],
	["Ei", 2 ** 60],
	["Pi", 2 ** 50],
	["Ti", 2 ** 40],
	["Gi", 2 ** 30],
	["Mi", 2 ** 20],
	["ki", 2 ** 10]
])

interface Resolved {
	readonly unit: Unit
	// The prefix's multiplier, at the unit's power.
	readonly scale: number
}

// A unit as written: a unit's own name first, then a binary prefix (information only), then a metric one.
function resolve(name: string): Resolved | null {
	const unit = BY_NAME.get(name)

	if (unit !== undefined) return { unit, scale: 1 }

	const binary = BINARY.get(name.slice(0, 2))
	const binaryUnit = binary === undefined ? undefined : BY_NAME.get(name.slice(2))

	if (binary !== undefined && binaryUnit?.group === Group.INFORMATION) return { unit: binaryUnit, scale: binary }

	for (const length of [2, 1]) {
		const multiplier = METRIC.get(name.slice(0, length))
		const base = multiplier === undefined ? undefined : BY_NAME.get(name.slice(length))

		if (multiplier !== undefined && base !== undefined && base.prefixed) return { unit: base, scale: multiplier ** base.power }
	}

	return null
}

// A temperature in kelvin, from the unit's own scale (`degrees` already times its prefix).
function toKelvin(name: string, degrees: number): number {
	switch (name) {
		case "C":
		case "cel":
			return degrees + 273.15
		case "F":
		case "fah":
			return ((degrees - 32) * 5) / 9 + 273.15
		case "Rank":
			return (degrees * 5) / 9
		case "Reau":
			return degrees * 1.25 + 273.15
		default:
			return degrees
	}
}

function fromKelvin(name: string, kelvin: number): number {
	switch (name) {
		case "C":
		case "cel":
			return kelvin - 273.15
		case "F":
		case "fah":
			return ((kelvin - 273.15) * 9) / 5 + 32
		case "Rank":
			return (kelvin * 9) / 5
		case "Reau":
			return (kelvin - 273.15) / 1.25
		default:
			return kelvin
	}
}

// The base unit name of a temperature written with a prefix ("mK" → "K").
function temperatureName(name: string): string {
	for (const base of ["cel", "kel", "C", "K"]) if (name.endsWith(base) && name !== base) return base

	return name
}

export function convert(value: number, from: string, to: string): number | CellError {
	const source = resolve(from)
	const target = resolve(to)

	if (source === null || target === null || source.unit.group !== target.unit.group) return naError()

	if (source.unit.group === Group.TEMPERATURE) {
		const kelvin = toKelvin(temperatureName(from), value * source.scale)

		return fromKelvin(temperatureName(to), kelvin) / target.scale
	}

	// One factor first, so a conversion between a unit and itself is exact.
	return value * ((source.unit.factor * source.scale) / (target.unit.factor * target.scale))
}

export class ConvertFunctions extends SheetFunctionPlugin {
	static override implementedFunctions: ImplementedFunctions = {
		CONVERT: {
			method: "convert",
			parameters: [
				{ argumentType: FunctionArgumentType.NUMBER },
				{ argumentType: FunctionArgumentType.STRING },
				{ argumentType: FunctionArgumentType.STRING }
			]
		}
	}

	convert(ast: PluginCall, state: PluginState): unknown {
		return this.runFunction(ast.args, state, this.metadata("CONVERT"), convert)
	}
}
