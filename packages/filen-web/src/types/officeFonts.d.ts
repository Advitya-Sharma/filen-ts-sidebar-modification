// vite/office-fonts-plugin.ts: each bundled font's URL by its path under @filen/office's assets/fonts.
declare module "virtual:office-fonts" {
	export const OFFICE_FONT_URLS: Readonly<Record<string, string>>
}

// vite/office-fonts-plugin.ts: each gzip'd TrueType face PDF export embeds, by its path under assets/fonts.
declare module "virtual:office-pdf-fonts" {
	export const OFFICE_PDF_FONT_URLS: Readonly<Record<string, string>>
}
