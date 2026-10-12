import { COMMENT_UNITS, docxComments } from "../../review/comments"
import type { DocumentCommandContext, DocumentHandlers, FeatureHandlers } from "../registry"

// Comments: add, reply, edit, resolve, reopen, delete and step through them (document commands: they touch the
// story being edited, the story a comment is anchored in and the comment's own story). What records an author
// refuses without one (the app asks for the name first).

const store = (context: DocumentCommandContext) => docxComments(context.editor)
const writable = (context: DocumentCommandContext): boolean => !context.editor.body.session.readOnly
const anyComment = (context: DocumentCommandContext): boolean => store(context).any()
const onComment = (context: DocumentCommandContext): boolean => writable(context) && anyComment(context)

const document: DocumentHandlers = {
	"comment.add": {
		enabled: context => store(context).canAdd(),
		run: (context, command) => store(context).add(command.text)
	},
	"comment.reply": {
		enabled: onComment,
		run: (context, command) => store(context).reply(command.id, command.text)
	},
	"comment.edit": {
		enabled: onComment,
		run: (context, command) => store(context).edit(command.id, command.text)
	},
	"comment.resolve": {
		enabled: onComment,
		run: (context, command) => store(context).resolve(command.id, true)
	},
	"comment.reopen": {
		enabled: onComment,
		run: (context, command) => store(context).resolve(command.id, false)
	},
	"comment.delete": {
		enabled: onComment,
		run: (context, command) => store(context).remove(command.id)
	},
	"comment.next": {
		enabled: anyComment,
		run: context => store(context).step(true)
	},
	"comment.prev": {
		enabled: anyComment,
		run: context => store(context).step(false)
	}
}

export const COMMENT_HANDLERS: FeatureHandlers = { document, units: COMMENT_UNITS }
