// SPDX-License-Identifier: GPL-3.0-only
//
// The MT Extra code table: an adaptation of the Mozilla MathML project's "MT Extra" encoding page
// (https://www-archive.mozilla.org/projects/mathml/fonts/encoding/mtextra), portions © 1998–2009 by
// individual mozilla.org contributors, licensed to us under CC BY-SA 4.0 (Mozilla grants CC BY-SA 3.0
// Unported "or any later version"). Modified: the combining embellishments are mapped to their spacing
// forms, the arrow embellishments to spacing arrows, and the stretchy brace pieces are left out. This
// adaptation is licensed under the GNU General Public License version 3, as Creative Commons permits
// for BY-SA 4.0 adaptations, and is combined with the rest of the package under AGPL-3.0 section 13.
// See THIRD_PARTY_NOTICES.md. No MT Extra font file was read.

import type { SymbolTable } from "../fonts/index"

// MathType's companion font, codes 0x20..0x7F: the Unicode character each draws, 0 where none.
export const MT_EXTRA: SymbolTable = (() => {
	const table = new Array<number>(0xe0).fill(0)
	const set = (code: number, unicode: number): void => {
		table[code - 0x20] = unicode
	}

	set(0x20, 0x20)
	set(0x23, 0x2cb)
	set(0x24, 0x2c6)
	set(0x25, 0x2dc)
	set(0x26, 0x2d9)
	set(0x3c, 0x25c3)
	set(0x3e, 0x25b9)
	set(0x43, 0x2210)
	set(0x44, 0x19b)
	set(0x49, 0x2229)
	set(0x4b, 0x2026)
	set(0x4c, 0x22ef)
	set(0x4d, 0x22ee)
	set(0x4e, 0x22f0)
	set(0x4f, 0x22f1)
	set(0x51, 0x2235)
	set(0x55, 0x222a)
	set(0x60, 0x2035)
	set(0x61, 0x21a6)
	set(0x62, 0x2195)
	set(0x63, 0x21d5)
	set(0x66, 0x227b)
	set(0x68, 0x210f)
	set(0x6c, 0x2113)
	set(0x6d, 0x2213)
	set(0x6f, 0x2218)
	set(0x70, 0x227a)
	set(0x72, 0x2192)
	set(0x73, 0x2190)
	set(0x74, 0x2194)
	set(0x75, 0x23af)
	set(0x76, 0x21c0)
	set(0x77, 0x21bc)
	set(0x7b, 0x23df)
	set(0x7d, 0x23de)

	return table
})()
