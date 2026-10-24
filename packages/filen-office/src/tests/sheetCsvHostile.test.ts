import { describe, expect, it } from "vitest"
import { parseCsv } from "../core/sheet/csv/csvParse"
import { openTextSession } from "../core/sheet/text"

// A text file past the cell or row limit is refused while it is read, before its rows are built: memory stays near
// what the limit allows, whatever the file holds.

describe("a hostile CSV", () => {
	it("is refused by the cell limit as the parser passes it", () => {
		expect(() => parseCsv("a,b,c\n".repeat(4), ",", { maxCells: 11 })).toThrow("too many cells")
		expect(parseCsv("a,b,c\n".repeat(4), ",", { maxCells: 12 })).toHaveLength(4)
	})

	it("is refused by the row limit as the parser passes it", () => {
		expect(() => parseCsv("\n".repeat(5), ",", { maxRows: 4 })).toThrow("too many rows")
		expect(parseCsv("\n".repeat(4), ",", { maxRows: 4 })).toHaveLength(4)
	})

	it("is refused by the column limit as the parser passes it", () => {
		expect(() => parseCsv("a,b,c", ",", { maxColumns: 2 })).toThrow("too many columns")
	})

	it("of 30 MB of line ends is refused fast, in bounded memory", { timeout: 60_000 }, async () => {
		const bytes = new Uint8Array(30_000_000).fill(0x0a)
		const before = process.memoryUsage().rss
		const cpu = process.cpuUsage()
		const opened = await openTextSession(bytes, { kind: "csv" }, () => undefined)
		const used = process.cpuUsage(cpu)

		expect(opened.kind).toBe("failed")
		expect(opened.kind === "failed" && opened.reason).toBe("tooLarge")
		// The text decoded (60 MB as UTF-16) and at most the limit's rows of one empty field each.
		expect(process.memoryUsage().rss - before).toBeLessThan(400_000_000)
		expect((used.user + used.system) / 1000).toBeLessThan(5_000)
	})
})
