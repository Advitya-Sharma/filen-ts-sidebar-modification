// Sort key for library titles: punctuation that should not affect order removed and a leading English article
// dropped ("The Matrix" sorts under M).

const LEADING_ARTICLE = /^(?:the|a|an)\s+/iu
const AS_SPACE = /[.+%]/gu
const DROPPED = /[,&\-{}'’]/gu
const SPACE_RUN = /\s{2,}/gu

// `nfoSortTitle` (an NFO <sorttitle>) is taken verbatim when present.
export function videoSortName(title: string, nfoSortTitle?: string): string {
	if (nfoSortTitle !== undefined && nfoSortTitle.trim().length > 0) {
		return nfoSortTitle.trim()
	}

	// The article must be followed by whitespace in the title itself: "A.I." keeps its "A".
	const withoutArticle = title.trim().replace(LEADING_ARTICLE, "")
	const base = withoutArticle.length > 0 ? withoutArticle : title.trim()

	return base.replace(AS_SPACE, " ").replace(DROPPED, "").replace(SPACE_RUN, " ").trim()
}
