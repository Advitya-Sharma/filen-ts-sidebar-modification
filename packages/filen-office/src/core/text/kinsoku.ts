import { Kinsoku } from "./lineBreak"

// Default kinsoku lists, as ECMA-376 Part 4 §2.3.1.16 (kinsoku) gives them per language. A document's
// settings (w:noLineBreaksBefore / w:noLineBreaksAfter) replace them per language.

export interface KinsokuLists {
	// Characters that cannot start a line.
	readonly noStart: string
	// Characters that cannot end a line.
	readonly noEnd: string
}

export const KINSOKU_SIMPLIFIED_CHINESE: KinsokuLists = {
	noStart: "!%),.:;>?]}¢¨°·ˇˉ―‖’”…‰′″›℃∶、。〃〉》」』】〕〗〞︶︺︾﹀﹄﹚﹜﹞！＂％＇），．：；？］｀｜｝～￠",
	noEnd: "$([{£¥·‘“〈《「『【〔〖〝﹙﹛﹝＄（．［｛￡￥"
}

export const KINSOKU_TRADITIONAL_CHINESE: KinsokuLists = {
	noStart: "!),.:;?]}¢·–—’”•‥…‧′╴、。〉》」』】〕〞︰︱︳︴︶︸︺︼︾﹀﹂﹄﹏﹐﹑﹒﹔﹕﹖﹗﹚﹜﹞！），．：；？］｜｝､",
	noEnd: "([{£¥‘“‵〈《「『【〔〝︵︷︹︻︽︿﹁﹃﹙﹛﹝（｛"
}

export const KINSOKU_JAPANESE: KinsokuLists = {
	noStart: "!%),.:;?]}¢°’”‰′″℃、。々〉》」』】〕゛゜ゝゞ・ヽヾ！％），．：；？］｝｡｣､･ﾞﾟ￠",
	noEnd: "$([\\{£¥‘“〈《「『【〔＄（［｛｢￡￥"
}

// With w:strictFirstAndLastChars in the document settings.
export const KINSOKU_JAPANESE_STRICT: KinsokuLists = {
	noStart:
		"!%),.:;?]}¢°’”‰′″℃、。々〉》」』】〕ぁぃぅぇぉっゃゅょゎ゛゜ゝゞァィゥェォッャュョヮヵヶ・ーヽヾ！％），．：；？］｝｡｣､･ｧｨｩｪｫｬｭｮｯｰﾞﾟ￠",
	noEnd: KINSOKU_JAPANESE.noEnd
}

export const KINSOKU_KOREAN: KinsokuLists = {
	noStart: "!%),.:;?]}¢°’”′″℃〉》」』】〕！％），．：；？］｝￠",
	noEnd: "$([\\{£¥‘“〈《「『【〔＄（［｛￡￥￦"
}

// The default lists for a BCP 47 East Asian language tag, or null for languages kinsoku does not apply to
// (Word applies it to Chinese, Japanese and Korean text only).
export function kinsokuLists(language: string | undefined, strict: boolean): KinsokuLists | null {
	if (language === undefined) return null

	const tag = language.toLowerCase()

	if (tag === "ja" || tag.startsWith("ja-")) return strict ? KINSOKU_JAPANESE_STRICT : KINSOKU_JAPANESE
	if (tag === "ko" || tag.startsWith("ko-")) return KINSOKU_KOREAN
	if (tag === "zh" || tag.startsWith("zh-")) {
		return /^zh-(?:tw|hk|mo|hant)\b/.test(tag) ? KINSOKU_TRADITIONAL_CHINESE : KINSOKU_SIMPLIFIED_CHINESE
	}

	return null
}

const compiled = new Map<KinsokuLists, Kinsoku>()

// Compiled lists, shared per list object.
export function kinsokuOf(lists: KinsokuLists): Kinsoku {
	let kinsoku = compiled.get(lists)

	if (kinsoku === undefined) {
		kinsoku = new Kinsoku(lists.noStart, lists.noEnd)
		compiled.set(lists, kinsoku)
	}

	return kinsoku
}
