// "preview" namespace, Word document viewer — Word 97–2003 documents (.doc, .dot) with a password to open. Spread into
// preview.ts; same catalog rules. The password prompt itself uses the shared previewDocxPassword* strings.
export const previewDocxLegacyPassword = {
	/** Docx viewer — the Save as .docx confirm dialog, added when the Word 97–2003 document had a password to open: the copy is saved without it */
	previewDocxSaveAsDocxUnencrypted: "The copy isn't password-protected: anyone who can open it can read it."
} as const
