import { partReader } from "../../model/reader"
import { trackingPersisted } from "../../review/track"
import { Names, Out, spliceChildren } from "../markup"
import { W_NAMESPACE } from "../source"
import { DocxWriteError } from "../values"
import type { PartContributor, SaveParts } from "./index"

// settings.xml: the track switch (w:trackRevisions) where the editor turned it from what the file has; nothing
// else of the part changes. A package without settings gets a part holding only the switch.

const SETTINGS_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
const SETTINGS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"

// CT_Settings (wml.xsd), in schema order.
const SETTINGS: readonly string[] = [
	"writeProtection",
	"view",
	"zoom",
	"removePersonalInformation",
	"removeDateAndTime",
	"doNotDisplayPageBoundaries",
	"displayBackgroundShape",
	"printPostScriptOverText",
	"printFractionalCharacterWidth",
	"printFormsData",
	"embedTrueTypeFonts",
	"embedSystemFonts",
	"saveSubsetFonts",
	"saveFormsData",
	"mirrorMargins",
	"alignBordersAndEdges",
	"bordersDoNotSurroundHeader",
	"bordersDoNotSurroundFooter",
	"gutterAtTop",
	"hideSpellingErrors",
	"hideGrammaticalErrors",
	"activeWritingStyle",
	"proofState",
	"formsDesign",
	"attachedTemplate",
	"linkStyles",
	"stylePaneFormatFilter",
	"stylePaneSortMethod",
	"documentType",
	"mailMerge",
	"revisionView",
	"trackRevisions",
	"doNotTrackMoves",
	"doNotTrackFormatting",
	"documentProtection",
	"autoFormatOverride",
	"styleLockTheme",
	"styleLockQFSet",
	"defaultTabStop",
	"autoHyphenation",
	"consecutiveHyphenLimit",
	"hyphenationZone",
	"doNotHyphenateCaps",
	"showEnvelope",
	"summaryLength",
	"clickAndTypeStyle",
	"defaultTableStyle",
	"evenAndOddHeaders",
	"bookFoldRevPrinting",
	"bookFoldPrinting",
	"bookFoldPrintingSheets",
	"drawingGridHorizontalSpacing",
	"drawingGridVerticalSpacing",
	"displayHorizontalDrawingGridEvery",
	"displayVerticalDrawingGridEvery",
	"doNotUseMarginsForDrawingGridOrigin",
	"drawingGridHorizontalOrigin",
	"drawingGridVerticalOrigin",
	"doNotShadeFormData",
	"noPunctuationKerning",
	"characterSpacingControl",
	"printTwoOnOne",
	"strictFirstAndLastChars",
	"noLineBreaksAfter",
	"noLineBreaksBefore",
	"savePreviewPicture",
	"doNotValidateAgainstSchema",
	"saveInvalidXml",
	"ignoreMixedContent",
	"alwaysShowPlaceholderText",
	"doNotDemarcateInvalidXml",
	"saveXmlDataOnly",
	"useXSLTWhenSaving",
	"saveThroughXslt",
	"showXMLTags",
	"alwaysMergeEmptyNamespace",
	"updateFields",
	"hdrShapeDefaults",
	"footnotePr",
	"endnotePr",
	"compat",
	"docVars",
	"rsids",
	"attachedSchema",
	"themeFontLang",
	"clrSchemeMapping",
	"doNotIncludeSubdocsInStats",
	"doNotAutoCompressPictures",
	"forceUpgrade",
	"captions",
	"readModeInkLockDown",
	"smartTagType",
	"shapeDefaults",
	"doNotEmbedSmartTags",
	"decimalSymbol",
	"listSeparator"
]

async function writeTracking(parts: SaveParts, on: boolean): Promise<void> {
	const pkg = parts.pkg
	const name = pkg.part("settings")
	const xml = name === undefined ? undefined : await parts.xml(name)

	if (name === undefined || xml === undefined) {
		// Nothing to switch off in a package without settings.
		if (!on) return
		if (parts.context.model.strict) throw new DocxWriteError("unsupported", "New settings in a Strict document")

		const main = pkg.mainPart
		const target = parts.freeName(main.replace(/[^/]*$/, "settings.xml"))

		parts.add(
			target,
			new TextEncoder().encode(
				`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:settings xmlns:w="${W_NAMESPACE}"><w:trackRevisions/></w:settings>`
			),
			SETTINGS_CONTENT_TYPE
		)
		parts.relationship(main, { type: SETTINGS_TYPE, target: target.slice(target.lastIndexOf("/") + 1) })

		return
	}

	const reader = partReader(xml)
	const root = xml.documentElement()
	const names = new Names(reader.namespaces.prefixFor(root, W_NAMESPACE) ?? "w", "")
	const out = new Out()

	out.bytes(xml.bytes, 0, xml.start(root))
	spliceChildren(out, reader, root, new Map([["trackRevisions", on ? names.empty("trackRevisions") : ""]]), SETTINGS)
	out.bytes(xml.bytes, xml.end(root), xml.bytes.length)
	parts.replace(name, out.finish())
}

export const settingsParts: PartContributor = {
	async contribute(parts) {
		const context = parts.context
		const on = context.review.tracking

		if (on !== context.model.settings.trackRevisions && trackingPersisted(context)) await writeTracking(parts, on)
	}
}
