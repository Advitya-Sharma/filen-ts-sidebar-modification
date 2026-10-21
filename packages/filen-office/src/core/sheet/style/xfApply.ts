// An xf record's apply* attributes as written, with quotePrefix and pivotButton (§18.8.45): bits of XlsxXf flags.
export const XfApply = {
	NUMBER_FORMAT: 1 << 0,
	FONT: 1 << 1,
	FILL: 1 << 2,
	BORDER: 1 << 3,
	ALIGNMENT: 1 << 4,
	PROTECTION: 1 << 5,
	QUOTE_PREFIX: 1 << 6,
	PIVOT_BUTTON: 1 << 7
} as const
