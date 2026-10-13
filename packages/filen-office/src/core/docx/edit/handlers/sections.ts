import {
	bindEditor,
	canEditHeader,
	canLink,
	canRemove,
	editHeader,
	removeHeader,
	toggleEvenAndOdd,
	toggleLink,
	toggleTitlePage
} from "../../stories/headers"
import { canBreak, deleteSectionBreak, insertSectionBreak, pageSetup } from "../../stories/sections"
import type { DocumentCommandContext, DocumentHandlers, FeatureHandlers } from "../registry"

// Section breaks, page setup, headers and footers (document commands: they change the body's section
// properties and the stories sections refer to). Enablement is decided per snapshot: it looks at the caret and
// the sections as laid out, never through the document.

function bound(context: DocumentCommandContext): DocumentCommandContext {
	bindEditor(context.context, context.editor)

	return context
}

const writable = (context: DocumentCommandContext): boolean => !bound(context).editor.body.session.readOnly

const document: DocumentHandlers = {
	"insert.sectionBreak": {
		enabled: context => canBreak(bound(context).editor),
		run: (context, command) => insertSectionBreak(bound(context).editor, command.kind)
	},
	"section.deleteBreak": {
		// Exact at run time (the break of the caret's section): here, whether there is a break at all.
		enabled: context => writable(context) && context.editor.edits.model.sections.length > 1,
		run: context => deleteSectionBreak(bound(context).editor)
	},
	"page.setup": {
		enabled: writable,
		run: (context, command) => pageSetup(bound(context).editor, command)
	},
	"header.edit": {
		enabled: context => canEditHeader(bound(context).editor),
		run: context => editHeader(bound(context).editor, "header")
	},
	"footer.edit": {
		enabled: context => canEditHeader(bound(context).editor),
		run: context => editHeader(bound(context).editor, "footer")
	},
	"header.differentFirst": {
		enabled: writable,
		run: context => toggleTitlePage(bound(context).editor)
	},
	"header.differentOddEven": {
		enabled: writable,
		run: context => toggleEvenAndOdd(bound(context).editor)
	},
	"header.linkToPrevious": {
		enabled: context => canLink(bound(context).editor),
		run: context => toggleLink(bound(context).editor)
	},
	"header.remove": {
		enabled: context => canRemove(bound(context).editor, "header"),
		run: context => removeHeader(bound(context).editor, "header")
	},
	"footer.remove": {
		enabled: context => canRemove(bound(context).editor, "footer"),
		run: context => removeHeader(bound(context).editor, "footer")
	}
}

export const SECTION_HANDLERS: FeatureHandlers = { document }
