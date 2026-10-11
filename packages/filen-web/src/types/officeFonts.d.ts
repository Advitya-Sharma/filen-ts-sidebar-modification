// vite/office-fonts-plugin.ts: each bundled font's URL by its path under @filen/office's assets/fonts.
declare module "virtual:office-fonts" {
	export const OFFICE_FONT_URLS: Readonly<Record<string, string>>
}
