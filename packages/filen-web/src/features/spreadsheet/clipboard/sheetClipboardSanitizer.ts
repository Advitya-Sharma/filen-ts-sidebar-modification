import DOMPurify from "dompurify"
import { createSheetClipboardSanitizer } from "@filen/shared/dom"

// Clipboard HTML in and out of a sheet goes through the shared sheet clipboard sanitizer, on a DOMPurify instance of
// its own (the rich-text sanitizer's hooks live on the default one).
export const sanitizeSheetClipboard = createSheetClipboardSanitizer(DOMPurify(window))

// Pasted HTML past this many characters is not read (sanitising more would be a long task: about 55 ms a MiB in
// Chromium): the clipboard's text pastes its values.
export const MAX_PASTE_HTML = 768 * 1024
