import { describe, expect, it } from "vitest"
import { applyPara, applyTable, defaultPara, defaultRow } from "../core/doc/props"
import { forEachPrl, operandLength, readPrls, SPRM_NAMES, sprmName, sprmSgc } from "../core/doc/sprms"

function walk(bytes: readonly number[]): [string, number][] {
	const data = Uint8Array.from(bytes)
	const out: [string, number][] = []

	forEachPrl(data, 0, data.length, (sprm, _at, length) => out.push([sprmName(sprm), length]))

	return out
}

describe("sprm operands", () => {
	it("sizes operands by spra", () => {
		// spra 0 toggle, 1 byte, 2 word, 3 dword, 7 three bytes, 6 variable (size byte first).
		expect(
			walk([0x35, 0x08, 1, 0x48, 0x2a, 2, 0x43, 0x4a, 24, 0, 0x03, 0x6a, 1, 2, 3, 4, 0x6f, 0xe8, 9, 9, 9, 0x71, 0xca, 2, 7, 7])
		).toEqual([
			["sprmCFBold", 1],
			["sprmCIss", 1],
			["sprmCHps", 2],
			["sprmCPicLocation", 4],
			["sprm0xe86f", 3],
			["sprmCShd", 3]
		])
	})

	it("sizes sprmTDefTable by its 2-byte cb and sprmPChgTabs by its delete and add counts", () => {
		const defTable = [0x08, 0xd6, 6, 0, 1, 0, 0, 0x10, 0x0e]
		// cb 255: one deleted tab (4 bytes), one added tab (3 bytes).
		const chgTabs = [0x15, 0xc6, 255, 1, 0x10, 0, 0x20, 0, 1, 0x30, 0, 0]

		expect(walk([...defTable, ...chgTabs])).toEqual([
			["sprmTDefTable", 7],
			["sprmPChgTabs", 10]
		])
	})

	it("stops at a Prl that runs past the grpprl and keeps what precedes it", () => {
		expect(walk([0x35, 0x08, 1, 0x43, 0x4a, 24])).toEqual([["sprmCFBold", 1]])
		expect(operandLength(Uint8Array.of(9), 0, 1, 0x6a03)).toBe(-1)
	})

	it("lists every sprm of the specification once with a valid sgc", () => {
		expect(SPRM_NAMES.size).toBe(322)

		for (const sprm of SPRM_NAMES.keys()) expect([1, 2, 3, 4, 5]).toContain(sprmSgc(sprm))
	})

	it("reads tab stops added and removed by sprmPChgTabsPapx", () => {
		const para = defaultPara()
		// Add tabs at 720 (left) and 1440 (right, dot leader); then delete 720.
		const add = Uint8Array.of(0x0d, 0xc6, 8, 0, 2, 0xd0, 0x02, 0xa0, 0x05, 0, 0x0a)
		const remove = Uint8Array.of(0x0d, 0xc6, 4, 1, 0xd0, 0x02, 0)

		applyPara(para, add, 0, add.length)
		expect(para.tabs).toEqual([
			{ position: 720, align: 0, leader: 0 },
			{ position: 1440, align: 2, leader: 1 }
		])

		applyPara(para, remove, 0, remove.length)
		expect(para.tabs).toEqual([{ position: 1440, align: 2, leader: 1 }])
	})

	it("defines row cells from sprmTDefTable and inserts and merges cells", () => {
		const row = defaultRow()
		// Two columns: edges 0, 1000, 2500; no TC80s.
		const defTable = [0x08, 0xd6, 8, 0, 2, 0, 0, 0xe8, 0x03, 0xc4, 0x09]
		const insert = [0x21, 0x76, 1, 1, 0xf4, 0x01]
		const merge = [0x24, 0x56, 0, 2]
		const data = Uint8Array.from([...defTable, ...insert, ...merge])

		applyTable(row, data, 0, data.length)

		expect(row.cells.map(cell => [cell.left, cell.right, cell.horzMerge])).toEqual([
			[0, 1000, 1],
			[1000, 1500, 2],
			[1500, 3000, 0]
		])
		expect(readPrls(data, 0, data.length).map(item => sprmName(item.sprm))).toEqual(["sprmTDefTable", "sprmTInsert", "sprmTMerge"])
	})
})
