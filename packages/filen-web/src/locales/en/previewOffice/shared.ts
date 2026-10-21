// "preview" namespace — strings both office editors use (DOCX editing and the spreadsheet viewer on the office engine).
// Spread into preview.ts; same catalog rules. Builds with neither flag resolve it to the empty stub (vite.config.ts).
export const previewOfficeShared = {
	/** Office editors — dialog title asking for the name comments, notes and tracked changes are signed with */
	previewOfficeAuthorNameTitle: "Your name in documents",
	/** Office editors — dialog body asking for the author name; it is synced and can be changed in settings */
	previewOfficeAuthorNameBody: "Comments and tracked changes show this name. You can change it later in settings.",
	/** Docx editor — body of the name prompt when an edit was not made because track changes needs the author's name */
	previewOfficeAuthorNameTrackBody:
		"Track changes is on, so every edit carries your name. Give it to keep editing, or turn track changes off.",
	/** Office editors — author name field label */
	previewOfficeAuthorNameLabel: "Name",
	/** Office editors — author name dialog submit button */
	previewOfficeAuthorNameSave: "Save",
	/** Docx editor — button in the name prompt turning track changes off instead of giving a name */
	previewOfficeTurnOffTracking: "Turn off track changes",
	/** Office editors — colour picker: automatic text colour (black on light, white on dark) */
	previewOfficeColorAutomatic: "Automatic",
	/** Office editors — colour picker: no highlight or shading */
	previewOfficeColorNone: "No color",
	/** Office editors — colour picker section heading */
	previewOfficeColorThemeColors: "Theme colors",
	/** Office editors — colour picker section heading */
	previewOfficeColorStandardColors: "Standard colors",
	/** Office editors — a theme colour variation's name; {{color}} is the colour, {{percent}} how much lighter */
	previewOfficeColorLighter: "{{color}}, lighter {{percent}}%",
	/** Office editors — a theme colour variation's name; {{color}} is the colour, {{percent}} how much darker */
	previewOfficeColorDarker: "{{color}}, darker {{percent}}%",
	/** Office editors — theme colour name */
	previewOfficeColorBackground1: "Background 1",
	/** Office editors — theme colour name */
	previewOfficeColorText1: "Text 1",
	/** Office editors — theme colour name */
	previewOfficeColorBackground2: "Background 2",
	/** Office editors — theme colour name */
	previewOfficeColorText2: "Text 2",
	/** Office editors — theme colour name */
	previewOfficeColorAccent1: "Accent 1",
	/** Office editors — theme colour name */
	previewOfficeColorAccent2: "Accent 2",
	/** Office editors — theme colour name */
	previewOfficeColorAccent3: "Accent 3",
	/** Office editors — theme colour name */
	previewOfficeColorAccent4: "Accent 4",
	/** Office editors — theme colour name */
	previewOfficeColorAccent5: "Accent 5",
	/** Office editors — theme colour name */
	previewOfficeColorAccent6: "Accent 6",
	/** Office editors — standard colour name */
	previewOfficeColorDarkRed: "Dark red",
	/** Office editors — standard colour name */
	previewOfficeColorRed: "Red",
	/** Office editors — standard colour name */
	previewOfficeColorOrange: "Orange",
	/** Office editors — standard colour name */
	previewOfficeColorYellow: "Yellow",
	/** Office editors — standard colour name */
	previewOfficeColorLightGreen: "Light green",
	/** Office editors — standard colour name */
	previewOfficeColorGreen: "Green",
	/** Office editors — standard colour name */
	previewOfficeColorLightBlue: "Light blue",
	/** Office editors — standard colour name */
	previewOfficeColorBlue: "Blue",
	/** Office editors — standard colour name */
	previewOfficeColorDarkBlue: "Dark blue",
	/** Office editors — standard colour name */
	previewOfficeColorPurple: "Purple"
}
