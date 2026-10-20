// What builds without VITE_SHEET_NEXT resolve spreadsheetView.tsx to (vite.config.ts): the current viewer itself,
// so neither the office engine's viewer nor the switch between them is in the build.
export { SpreadsheetViewer as SpreadsheetView } from "@/features/spreadsheet/components/spreadsheetViewer"
