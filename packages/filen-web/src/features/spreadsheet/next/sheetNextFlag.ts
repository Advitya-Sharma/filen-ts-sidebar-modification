// Spreadsheets open in the office engine's viewer (@filen/office core/xlsx, dom/sheet, react/sheet) only in builds
// with VITE_SHEET_NEXT=1 until it reaches parity; every other build keeps the current viewer. Read by dot access
// so the build replaces it with a constant and drops the other viewer.
export const SHEET_NEXT: boolean = import.meta.env.VITE_SHEET_NEXT === "1"
