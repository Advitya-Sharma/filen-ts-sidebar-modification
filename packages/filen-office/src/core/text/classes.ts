// Numeric property values as the generated tables encode them (order of LINE_BREAK_CLASS_NAMES and
// BIDI_CLASS_NAMES in packed.ts).

export const LineBreakClass = {
	BK: 0,
	CR: 1,
	LF: 2,
	NL: 3,
	CM: 4,
	ZWJ: 5,
	WJ: 6,
	ZW: 7,
	GL: 8,
	SP: 9,
	B2: 10,
	BA: 11,
	BB: 12,
	HY: 13,
	HH: 14,
	CB: 15,
	CL: 16,
	CP: 17,
	EX: 18,
	IN: 19,
	NS: 20,
	OP: 21,
	QU: 22,
	// QU with General_Category Pi / Pf.
	QU_PI: 23,
	QU_PF: 24,
	IS: 25,
	NU: 26,
	PO: 27,
	PR: 28,
	SY: 29,
	AK: 30,
	AL: 31,
	AP: 32,
	AS: 33,
	CJ: 34,
	EB: 35,
	EM: 36,
	H2: 37,
	H3: 38,
	HL: 39,
	ID: 40,
	JL: 41,
	JV: 42,
	JT: 43,
	RI: 44,
	VF: 45,
	VI: 46
} as const

export type LineBreakClass = (typeof LineBreakClass)[keyof typeof LineBreakClass]

export const LINE_BREAK_CLASS_MASK = 0x3f
export const EAST_ASIAN_FLAG = 0x40
export const PICTOGRAPHIC_UNASSIGNED_FLAG = 0x80

export const BidiClass = {
	L: 0,
	R: 1,
	AL: 2,
	EN: 3,
	ES: 4,
	ET: 5,
	AN: 6,
	CS: 7,
	NSM: 8,
	BN: 9,
	B: 10,
	S: 11,
	WS: 12,
	ON: 13,
	LRE: 14,
	LRO: 15,
	RLE: 16,
	RLO: 17,
	PDF: 18,
	LRI: 19,
	RLI: 20,
	FSI: 21,
	PDI: 22
} as const

export type BidiClass = (typeof BidiClass)[keyof typeof BidiClass]
