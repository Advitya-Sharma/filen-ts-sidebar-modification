import { describe, expect, it } from "vitest"
import { detectLanguage } from "../languageDetect"

describe("detectLanguage", () => {
	it.each([
		["en", "I moved the three invoices into your Archive directory."],
		["de", "Ich habe die drei Rechnungen in deinen Archiv-Ordner verschoben."],
		["es", "He movido las tres facturas a tu carpeta de archivo."],
		["fr", "J'ai déplacé les trois factures dans ton dossier d'archives."],
		["pt", "Eu movi as três faturas para a sua pasta de arquivo."],
		["ru", "Я переместил три счёта в папку архива."],
		["tr", "Üç faturayı arşiv klasörüne taşıdım."],
		["ar", "لقد نقلت الفواتير الثلاث إلى مجلد الأرشيف."],
		["hi", "मैंने तीनों बिल आर्काइव फ़ोल्डर में डाल दिए हैं।"],
		["id", "Saya sudah memindahkan tiga tagihan ke folder arsip kamu."],
		["zh", "我已经把三张发票移到了归档文件夹。"],
		["ja", "3つの請求書をアーカイブフォルダに移動しました。"]
	])("recognizes %s", (lang, text) => {
		expect(detectLanguage(text)).toBe(lang)
	})

	it("ignores file names, handles and paths copied from tool results", () => {
		expect(detectLanguage("Готово: Invoice_2024-01_Telekom.pdf, Invoice_2024-03_Stadtwerke.pdf и f3 теперь в /Archive.")).toBe("ru")
	})

	it("is not swayed by names and quotes from another language", () => {
		expect(detectLanguage("Tienes chats abiertos con los siguientes participantes: Anna Schmidt, Ahmet Yılmaz, Мария Иванова.")).toBe(
			"es"
		)
		expect(
			detectLanguage("Anna wrote: \u201cDenk bitte auch an den Termin am Freitag um 10 Uhr.\u201d That is all she said today.")
		).toBe("en")
		expect(detectLanguage("Para criar a nota, preciso que você me diga o título dela. Qual título você prefere?")).toBe("pt")
	})

	it("abstains on text too short to tell", () => {
		expect(detectLanguage("OK")).toBe("und")
	})
})
