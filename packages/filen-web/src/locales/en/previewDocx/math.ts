// "preview" namespace, Word document editor — editing equations as LaTeX in a popover with a live preview. Spread into preview.ts; same catalog rules.
export const previewDocxMath = {
	/** Docx editor — Insert menu item opening the equation popover to type a new equation as LaTeX (Word's Insert ▸ Equation) */
	previewDocxEquationInsert: "Equation",
	/** Docx editor — menu item opening the selected equation in the equation popover to change it */
	previewDocxEquationEdit: "Edit equation",
	/** Docx editor — accessible name of the equation popover while a new equation is typed */
	previewDocxEquationInsertTitle: "Insert equation",
	/** Docx editor — accessible name of the equation popover showing an equation that can't be changed here (a view-only document) */
	previewDocxEquationViewTitle: "Equation",
	/** Docx editor — label of the field in the equation popover where the equation is typed in LaTeX notation */
	previewDocxEquationLatex: "LaTeX",
	/** Docx editor — example shown in the empty LaTeX field of the equation popover; keep the LaTeX as it is */
	previewDocxEquationPlaceholder: "For example \\frac{a}{b} + x^2",
	/** Docx editor — accessible name of the live preview of the typed equation in the equation popover */
	previewDocxEquationPreview: "Preview",
	/** Docx editor — switch in the equation popover placing a new equation centered on a line of its own instead of within the text (Word's Display / Inline) */
	previewDocxEquationDisplay: "On its own line",
	/** Docx editor — equation popover button adding the typed equation to the document */
	previewDocxEquationInsertAction: "Insert",
	/** Docx editor — equation popover button replacing the equation with the typed one */
	previewDocxEquationUpdateAction: "Update",
	/** Docx editor — equation popover button closing it without changing the document */
	previewDocxEquationCancel: "Cancel",
	/** Docx editor — equation popover button closing a view-only equation */
	previewDocxEquationClose: "Close",
	/** Docx editor — note in the equation popover when the equation has formatting or parts LaTeX can't express, which an update drops */
	previewDocxEquationLossy: "Updating it as LaTeX simplifies this equation: formatting and parts LaTeX can't express are dropped.",
	/** Docx editor — LaTeX error in the equation popover: the equation is longer than the editor accepts */
	previewDocxEquationErrorTooLong: "The equation is too long.",
	/** Docx editor — LaTeX error in the equation popover: groups, fractions or roots inside each other too many levels deep ({{position}}: character number where reading stopped) */
	previewDocxEquationErrorTooDeep: "It's nested too deeply (at character {{position, number}}).",
	/** Docx editor — LaTeX error in the equation popover: a command the editor doesn't know ({{token}}: the command as typed, e.g. \foo) */
	previewDocxEquationErrorUnknownCommand: "Unknown command {{token}} (at character {{position, number}}).",
	/** Docx editor — LaTeX error in the equation popover: a command or ^/_ without what it applies to ({{token}}: the command as typed) */
	previewDocxEquationErrorMissingArgument: "{{token}} is missing its argument (at character {{position, number}}).",
	/** Docx editor — LaTeX error in the equation popover: a brace, \left, \begin or similar without its partner ({{token}}: as typed) */
	previewDocxEquationErrorUnbalanced: "{{token}} has no matching partner (at character {{position, number}}).",
	/** Docx editor — LaTeX error in the equation popover: a character or command that can't stand where it is, e.g. a second ^ ({{token}}: as typed) */
	previewDocxEquationErrorUnexpected: "{{token}} can't be used here (at character {{position, number}})."
} as const
