import { FunctionArgumentType, type ImplementedFunctions } from "hyperformula"
import { MAX_SERIAL_1900, MAX_SERIAL_1904, serialDate, type DateParts } from "../../numfmt"
import { numError, SheetFunctionPlugin, type PluginCall, type PluginState } from "./plugin"

// The parts of a date serial as Excel reads them, through the same calendar the cells show: in the 1900 system 0 is
// 1900-01-00 and 60 is 1900-02-29 (ECMA-376 Part 1 §18.17.4.1), where the engine read 0 as 1899-12-31.

const DATE_PART = { parameters: [{ argumentType: FunctionArgumentType.NUMBER, minValue: 0 }] }

export class DateFunctions extends SheetFunctionPlugin {
	static override implementedFunctions: ImplementedFunctions = {
		DAY: { method: "day", ...DATE_PART },
		MONTH: { method: "month", ...DATE_PART },
		YEAR: { method: "year", ...DATE_PART }
	}

	private readonly parts: DateParts = { year: 0, month: 0, day: 0, weekday: 0 }

	private part(ast: PluginCall, state: PluginState, name: string, pick: (parts: DateParts) => number): unknown {
		const date1904 = this.config.nullDate.year === 1904

		return this.runFunction(ast.args, state, this.metadata(name), (serial: number) => {
			const day = Math.floor(serial)

			if (day > (date1904 ? MAX_SERIAL_1904 : MAX_SERIAL_1900)) return numError()

			return pick(serialDate(day, date1904, this.parts))
		})
	}

	day(ast: PluginCall, state: PluginState): unknown {
		return this.part(ast, state, "DAY", parts => parts.day)
	}

	month(ast: PluginCall, state: PluginState): unknown {
		return this.part(ast, state, "MONTH", parts => parts.month)
	}

	year(ast: PluginCall, state: PluginState): unknown {
		return this.part(ast, state, "YEAR", parts => parts.year)
	}
}
