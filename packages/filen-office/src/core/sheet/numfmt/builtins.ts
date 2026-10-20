import type { NumFmtLocale } from "./locale"

// The built-in number formats a cell format names by id without a numFmt record (ECMA-376 Part 1 §18.8.30), as
// Excel defines them (MS-OI29500 §2.1.712 a: 14, 22, 37-40 and 47 differ from the standard's), in every language.
export const EXCEL_BUILTIN_FORMATS: ReadonlyMap<number, string> = new Map([
	[0, "General"],
	[1, "0"],
	[2, "0.00"],
	[3, "#,##0"],
	[4, "#,##0.00"],
	[9, "0%"],
	[10, "0.00%"],
	[11, "0.00E+00"],
	[12, "# ?/?"],
	[13, "# ??/??"],
	[14, "m/d/yyyy"],
	[15, "d-mmm-yy"],
	[16, "d-mmm"],
	[17, "mmm-yy"],
	[18, "h:mm AM/PM"],
	[19, "h:mm:ss AM/PM"],
	[20, "h:mm"],
	[21, "h:mm:ss"],
	[22, "m/d/yyyy h:mm"],
	[37, "#,##0_);(#,##0)"],
	[38, "#,##0_);[Red](#,##0)"],
	[39, "#,##0.00_);(#,##0.00)"],
	[40, "#,##0.00_);[Red](#,##0.00)"],
	[45, "mm:ss"],
	[46, "[h]:mm:ss"],
	[47, "mm:ss.0"],
	[48, "##0.0E+0"],
	[49, "@"]
])

// The language-dependent ids of §18.8.30's tables, for an app to put in NumFmtLocale.builtins. Korean 55 is
// Excel's (MS-OI29500 §2.1.712 a).
const ZH_TW_DATE = "[$-404]e/m/d"
const ZH_TW_LONG = '[$-404]e"年"m"月"d"日"'
const ZH_TW_TIME = '上午/下午hh"時"mm"分"'
const ZH_TW_SECONDS = '上午/下午hh"時"mm"分"ss"秒"'
const ZH_CN_MONTH = 'yyyy"年"m"月"'
const ZH_CN_DAY = 'm"月"d"日"'
const ZH_CN_TIME = '上午/下午h"时"mm"分"'
const ZH_CN_SECONDS = '上午/下午h"时"mm"分"ss"秒"'
const JA_DATE = "[$-411]ge.m.d"
const JA_LONG = '[$-411]ggge"年"m"月"d"日"'
const JA_MONTH = 'yyyy"年"m"月"'
const JA_DAY = 'm"月"d"日"'
const KO_DATE = 'yyyy"年" mm"月" dd"日"'

export const IMPLIED_BUILTINS: Readonly<Record<"zh-TW" | "zh-CN" | "ja-JP" | "ko-KR" | "th-TH", Readonly<Record<number, string>>>> = {
	"zh-TW": {
		27: ZH_TW_DATE,
		28: ZH_TW_LONG,
		29: ZH_TW_LONG,
		30: "m/d/yy",
		31: 'yyyy"年"m"月"d"日"',
		32: 'hh"時"mm"分"',
		33: 'hh"時"mm"分"ss"秒"',
		34: ZH_TW_TIME,
		35: ZH_TW_SECONDS,
		36: ZH_TW_DATE,
		50: ZH_TW_DATE,
		51: ZH_TW_LONG,
		52: ZH_TW_TIME,
		53: ZH_TW_SECONDS,
		54: ZH_TW_LONG,
		55: ZH_TW_TIME,
		56: ZH_TW_SECONDS,
		57: ZH_TW_DATE,
		58: ZH_TW_LONG
	},
	"zh-CN": {
		27: ZH_CN_MONTH,
		28: ZH_CN_DAY,
		29: ZH_CN_DAY,
		30: "m-d-yy",
		31: 'yyyy"年"m"月"d"日"',
		32: 'h"时"mm"分"',
		33: 'h"时"mm"分"ss"秒"',
		34: ZH_CN_TIME,
		35: ZH_CN_SECONDS,
		36: ZH_CN_MONTH,
		50: ZH_CN_MONTH,
		51: ZH_CN_DAY,
		52: ZH_CN_MONTH,
		53: ZH_CN_DAY,
		54: ZH_CN_DAY,
		55: ZH_CN_TIME,
		56: ZH_CN_SECONDS,
		57: ZH_CN_MONTH,
		58: ZH_CN_DAY
	},
	"ja-JP": {
		27: JA_DATE,
		28: JA_LONG,
		29: JA_LONG,
		30: "m/d/yy",
		31: 'yyyy"年"m"月"d"日"',
		32: 'h"時"mm"分"',
		33: 'h"時"mm"分"ss"秒"',
		34: JA_MONTH,
		35: JA_DAY,
		36: JA_DATE,
		50: JA_DATE,
		51: JA_LONG,
		52: JA_MONTH,
		53: JA_DAY,
		54: JA_LONG,
		55: JA_MONTH,
		56: JA_DAY,
		57: JA_DATE,
		58: JA_LONG
	},
	"ko-KR": {
		27: KO_DATE,
		28: "mm-dd",
		29: "mm-dd",
		30: "mm-dd-yy",
		31: 'yyyy"년" mm"월" dd"일"',
		32: 'h"시" mm"분"',
		33: 'h"시" mm"분" ss"초"',
		34: "yyyy-mm-dd",
		35: "yyyy-mm-dd",
		36: KO_DATE,
		50: KO_DATE,
		51: "mm-dd",
		52: "yyyy-mm-dd",
		53: "yyyy-mm-dd",
		54: "mm-dd",
		55: "yyyy/mm/dd",
		56: "yyyy-mm-dd",
		57: KO_DATE,
		58: "mm-dd"
	},
	"th-TH": {
		59: "t0",
		60: "t0.00",
		61: "t#,##0",
		62: "t#,##0.00",
		67: "t0%",
		68: "t0.00%",
		69: "t# ?/?",
		70: "t# ??/??",
		71: "ว/ด/ปปปป",
		72: "ว-ดดด-ปป",
		73: "ว-ดดด",
		74: "ดดด-ปป",
		75: "ช:นน",
		76: "ช:นน:ทท",
		77: "ว/ด/ปปปป ช:นน",
		78: "นน:ทท",
		79: "[ช]:นน:ทท",
		80: "นน:ทท.0",
		81: "d/m/bb"
	}
}

// The code of a built-in id: the locale's when it has one, else the table's; undefined for an id with neither.
export function builtinFormatCode(id: number, locale: NumFmtLocale): string | undefined {
	return locale.builtins[id] ?? EXCEL_BUILTIN_FORMATS.get(id)
}
