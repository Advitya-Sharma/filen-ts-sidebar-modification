import { describe, expect, it } from "vitest"
import type { SheetSessionApi } from "../core/sheet/protocol"
import type { DrawingAnchor } from "../core/xlsx/drawings/anchor"
import type { SheetDrawingObject, SheetDrawings } from "../core/xlsx/drawings/feature"
import { ItemKind, ItemInt, ITEM_INTS } from "../core/drawingml/displayList"
import { buildFixture, savedBytes } from "./sheetSessionFixtures"
import { changeOf, entries, refusalOf, settle } from "./sheetNotesFixtures"
import {
	CONNECTOR,
	DRAWINGS,
	DRAWINGS_BYTES,
	PICTURE,
	PNG,
	SVG,
	TEXT_BOX,
	openDrawings,
	strictOf,
	withMedia,
	wsDr
} from "./sheetDrawingsFixtures"
import { TYPES, relationship } from "./xlsxRefsFixtures"

// Drawing objects: listed with their anchors, drawn as display lists, moved, resized, deleted, given alternative
// text and inserted as pictures, saved through the gate and read back.

// A sheet's objects once its drawing is read (asking for its layout starts that).
async function drawingsOf(session: SheetSessionApi, sheet: number): Promise<SheetDrawings | undefined> {
	await session.layout(sheet)
	await settle()

	return (await session.layout(sheet)).extensions?.drawings
}

function objectOf(drawings: SheetDrawings | undefined, name: string): SheetDrawingObject {
	const found = drawings?.objects.find(object => object.name === name)

	if (found === undefined) throw new Error(`no object ${name}`)

	return found
}

const MOVED: DrawingAnchor = {
	mode: "twoCell",
	editAs: "oneCell",
	from: { row: 10, col: 2, rowOff: 9525, colOff: 0 },
	to: { row: 14, col: 4, rowOff: 9525, colOff: 0 }
}

describe("sheet drawings", () => {
	it("lists every anchored object in z-order with its kind, name, alternative text and anchor", async () => {
		const { session, sheets } = await openDrawings()
		const drawings = await drawingsOf(session, sheets[0] ?? 0)

		expect(drawings?.part).toBe("xl/drawings/drawing1.xml")
		expect(drawings?.objects.map(object => [object.key, object.kind, object.name])).toEqual([
			[0, "picture", "Picture 1"],
			[1, "shape", "TextBox 2"],
			[2, "chart", "Chart 3"],
			[3, "connector", "Straight Connector 4"]
		])
		expect(objectOf(drawings, "Picture 1").alt).toBe("A red dot")
		expect(objectOf(drawings, "Picture 1").anchor).toEqual({
			mode: "twoCell",
			editAs: "oneCell",
			from: { row: 1, col: 1, rowOff: 0, colOff: 0 },
			to: { row: 5, col: 3, rowOff: 0, colOff: 0 }
		})
		expect(objectOf(drawings, "Chart 3").chartPart).toBe("xl/charts/chart1.xml")
		expect(objectOf(drawings, "Straight Connector 4").anchor).toEqual({ mode: "absolute", x: 0, y: 0, cx: 952500, cy: 0 })
		// A sheet without a drawing lists nothing.
		expect(await drawingsOf(session, sheets[1] ?? 0)).toBeUndefined()
		session.close()
	})

	it("draws a picture as an image of its media part and a text box as a filled shape with text; charts are left to their renderer", async () => {
		const { session, sheets } = await openDrawings()
		const sheet = sheets[0] ?? 0
		const drawings = await drawingsOf(session, sheet)
		const picture = objectOf(drawings, "Picture 1")
		const rendered = await session.query("drawing.render", { sheet, key: picture.key, stamp: picture.stamp })

		expect(rendered?.kind).toBe("drawing")

		if (rendered?.kind !== "drawing") return

		const kinds = Array.from(
			{ length: rendered.drawing.itemCount },
			(_, item) => rendered.drawing.itemInts[item * ITEM_INTS + ItemInt.KIND]
		)

		expect(kinds).toContain(ItemKind.IMAGE)
		expect(rendered.drawing.images.map(image => image.partName)).toEqual(["xl/media/image1.png"])
		// Drawn at the frame its xfrm records: 1,219,200 × 762,000 EMU.
		expect([rendered.drawing.width, rendered.drawing.height]).toEqual([128, 80])

		const box = objectOf(drawings, "TextBox 2")
		const text = await session.query("drawing.render", { sheet, key: box.key, stamp: box.stamp })

		expect(text?.kind === "drawing" && text.drawing.texts.length).toBe(1)
		expect(text?.kind === "drawing" && text.drawing.paints.length).toBeGreaterThan(0)

		const chart = objectOf(drawings, "Chart 3")

		expect(await session.query("drawing.render", { sheet, key: chart.key, stamp: chart.stamp })).toEqual({
			kind: "placeholder",
			reason: "chart"
		})
		// A stamp the object no longer has draws nothing.
		expect(await session.query("drawing.render", { sheet, key: picture.key, stamp: "0:0:0" })).toBeNull()
		session.close()
	})

	it("hands out a picture's bytes, and nothing for parts that are not pictures", async () => {
		const { session } = await openDrawings()

		expect(await session.query("drawing.media", { partName: "xl/media/image1.png" })).toEqual({ bytes: PNG, type: "image/png" })
		expect(await session.query("drawing.media", { partName: "xl/workbook.xml" })).toBeNull()
		expect(await session.query("drawing.media", { partName: "xl/media/none.png" })).toBeNull()
		session.close()
	})

	it("moves a picture: its anchor and its own frame written, other parts kept, and an undo saves the source again", async () => {
		const { session, sheets } = await openDrawings()
		const sheet = sheets[0] ?? 0
		const picture = objectOf(await drawingsOf(session, sheet), "Picture 1")

		changeOf(await session.command({ type: "drawing.move", sheet, key: picture.key, stamp: picture.stamp, anchor: MOVED }))
		expect(objectOf(await drawingsOf(session, sheet), "Picture 1").anchor).toEqual(MOVED)

		const files = entries(await savedBytes(session))
		const drawing = files["xl/drawings/drawing1.xml"] ?? ""

		expect(drawing).toContain(
			'<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>9525</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col>'
		)
		// The frame follows: columns of 64 px, rows of 20 px, at 9,525 EMU per px.
		expect(drawing).toContain('<a:off x="1219200" y="1914525"/><a:ext cx="1219200" cy="762000"/>')
		// The other objects keep their markup.
		expect(drawing).toContain(TEXT_BOX)
		expect(drawing).toContain(CONNECTOR)
		expect(files["xl/drawings/_rels/drawing1.xml.rels"]).toBe(entries(DRAWINGS_BYTES)["xl/drawings/_rels/drawing1.xml.rels"])

		const reopened = await openDrawings(await savedBytes(session))

		expect(objectOf(await drawingsOf(reopened.session, reopened.sheets[0] ?? 0), "Picture 1").anchor).toEqual(MOVED)
		reopened.session.close()

		changeOf(await session.undo())
		expect(entries(await savedBytes(session))["xl/drawings/drawing1.xml"]).toBe(entries(DRAWINGS_BYTES)["xl/drawings/drawing1.xml"])
		session.close()
	})

	it("keeps only the drawing's current version parsed: an undo reads the earlier one again", async () => {
		const { session, sheets } = await openDrawings()
		const sheet = sheets[0] ?? 0
		const before = await drawingsOf(session, sheet)
		const picture = objectOf(before, "Picture 1")

		changeOf(await session.command({ type: "drawing.move", sheet, key: picture.key, stamp: picture.stamp, anchor: MOVED }))
		await drawingsOf(session, sheet)
		changeOf(await session.undo())

		const after = await drawingsOf(session, sheet)

		expect(after).toEqual(before)
		// Listed from a parse of its own, not one held since the source was shown.
		expect(after).not.toBe(before)
		session.close()
	})

	it("refuses a move of an object that changed since it was listed, and one of another anchor kind", async () => {
		const { session, sheets } = await openDrawings()
		const sheet = sheets[0] ?? 0
		const picture = objectOf(await drawingsOf(session, sheet), "Picture 1")

		expect(refusalOf(await session.command({ type: "drawing.move", sheet, key: picture.key, stamp: "stale", anchor: MOVED }))).toBe(
			"unsupported"
		)
		expect(
			refusalOf(
				await session.command({
					type: "drawing.move",
					sheet,
					key: picture.key,
					stamp: picture.stamp,
					anchor: {
						mode: "oneCell",
						from: MOVED.mode === "twoCell" ? MOVED.from : { row: 0, col: 0, rowOff: 0, colOff: 0 },
						cx: 1,
						cy: 1
					}
				})
			)
		).toBe("unsupported")
		expect(
			refusalOf(
				await session.command({
					type: "drawing.move",
					sheet,
					key: picture.key,
					stamp: picture.stamp,
					anchor: { ...MOVED, from: { row: 20, col: 0, rowOff: 0, colOff: 0 } }
				})
			)
		).toBe("tooLarge")
		session.close()
	})

	it("refuses to move an object the file locks, and keeps letting it be resized", async () => {
		const locked = wsDr(PICTURE.replace('<a:picLocks noChangeAspect="1"/>', '<a:picLocks noChangeAspect="1" noMove="1"/>'))
		const bytes = withMedia(
			buildFixture({ ...DRAWINGS, parts: { ...DRAWINGS.parts, "xl/drawings/drawing1.xml": { text: locked, type: TYPES.drawing } } }),
			{ "xl/media/image1.png": PNG }
		)
		const { session, sheets } = await openDrawings(bytes)
		const sheet = sheets[0] ?? 0
		const picture = objectOf(await drawingsOf(session, sheet), "Picture 1")

		expect(picture.noMove).toBe(true)
		expect(
			refusalOf(await session.command({ type: "drawing.move", sheet, key: picture.key, stamp: picture.stamp, anchor: MOVED }))
		).toBe("objectLocked")
		changeOf(
			await session.command({
				type: "drawing.move",
				sheet,
				key: picture.key,
				stamp: picture.stamp,
				anchor: {
					mode: "twoCell",
					editAs: "oneCell",
					from: { row: 1, col: 1, rowOff: 0, colOff: 0 },
					to: { row: 7, col: 4, rowOff: 0, colOff: 0 }
				}
			})
		)
		session.close()
	})

	it("sets and removes alternative text", async () => {
		const { session, sheets } = await openDrawings()
		const sheet = sheets[0] ?? 0
		const box = objectOf(await drawingsOf(session, sheet), "TextBox 2")

		changeOf(await session.command({ type: "drawing.alt", sheet, key: box.key, stamp: box.stamp, alt: 'Says "hello" & more' }))

		const named = objectOf(await drawingsOf(session, sheet), "TextBox 2")

		expect(named.alt).toBe('Says "hello" & more')
		expect(entries(await savedBytes(session))["xl/drawings/drawing1.xml"]).toContain(
			'<xdr:cNvPr id="3" name="TextBox 2" descr="Says &quot;hello&quot; &amp; more"/>'
		)
		changeOf(await session.command({ type: "drawing.alt", sheet, key: named.key, stamp: named.stamp, alt: "" }))
		expect(entries(await savedBytes(session))["xl/drawings/drawing1.xml"]).toContain('<xdr:cNvPr id="3" name="TextBox 2"/>')
		session.close()
	})

	it("deletes objects with the relationships only they used; a source picture's media stays", async () => {
		const { session, sheets } = await openDrawings()
		const sheet = sheets[0] ?? 0
		const drawings = await drawingsOf(session, sheet)
		const picture = objectOf(drawings, "Picture 1")
		const box = objectOf(drawings, "TextBox 2")

		changeOf(
			await session.command({
				type: "drawing.delete",
				sheet,
				objects: [
					{ key: picture.key, stamp: picture.stamp },
					{ key: box.key, stamp: box.stamp }
				]
			})
		)
		expect((await drawingsOf(session, sheet))?.objects.map(object => object.name)).toEqual(["Chart 3", "Straight Connector 4"])

		const files = entries(await savedBytes(session))

		expect(files["xl/drawings/drawing1.xml"]).not.toContain("Picture 1")
		expect(files["xl/drawings/_rels/drawing1.xml.rels"]).not.toContain("image1.png")
		expect(files["xl/drawings/_rels/drawing1.xml.rels"]).toContain("chart1.xml")
		expect(files["xl/media/image1.png"]).toBeDefined()

		changeOf(await session.undo())
		expect((await drawingsOf(session, sheet))?.objects).toHaveLength(4)
		expect(entries(await savedBytes(session))["xl/drawings/_rels/drawing1.xml.rels"]).toBe(
			entries(DRAWINGS_BYTES)["xl/drawings/_rels/drawing1.xml.rels"]
		)
		session.close()
	})

	it("inserts a picture on a sheet without a drawing: media, drawing part, relationships and <drawing>, read back", async () => {
		const { session, sheets } = await openDrawings()
		const sheet = sheets[1] ?? 0

		changeOf(
			await session.command({
				type: "drawing.insertImage",
				sheet,
				placement: { from: { row: 2, col: 1, rowOff: 0, colOff: 0 }, cx: 1219200, cy: 381000 },
				bytes: PNG,
				name: "",
				alt: "Logo"
			})
		)

		const listed = await drawingsOf(session, sheet)

		expect(listed?.objects.map(object => [object.kind, object.name, object.alt])).toEqual([["picture", "Picture 1", "Logo"]])
		// Two columns of 64 px and one row of 20 px: the picture ends at C4's top left.
		expect(listed?.objects[0]?.anchor).toEqual({
			mode: "twoCell",
			editAs: "oneCell",
			from: { row: 2, col: 1, rowOff: 0, colOff: 0 },
			to: { row: 4, col: 3, rowOff: 0, colOff: 0 }
		})

		const saved = await savedBytes(session)
		const files = entries(saved)

		expect(files["xl/worksheets/sheet2.xml"]).toMatch(/<drawing xmlns:r="[^"]*" r:id="rId\d+"\/><\/worksheet>$/)
		expect(files["xl/worksheets/_rels/sheet2.xml.rels"]).toMatch(/Type="[^"]*\/drawing" Target="\.\.\/drawings\/drawing2\.xml"/)
		expect(files["xl/drawings/_rels/drawing2.xml.rels"]).toMatch(/Type="[^"]*\/image" Target="\.\.\/media\/image2\.png"/)
		expect(files["[Content_Types].xml"]).toContain(
			'PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"'
		)
		// The package's Default for png types the new media.
		expect(files["[Content_Types].xml"]).toContain('<Default Extension="png" ContentType="image/png"/>')
		expect(files["[Content_Types].xml"]).not.toContain("/xl/media/image2.png")

		const reopened = await openDrawings(saved)
		const read = await drawingsOf(reopened.session, reopened.sheets[1] ?? 0)

		expect(read?.objects.map(object => object.alt)).toEqual(["Logo"])
		expect(await reopened.session.query("drawing.media", { partName: "xl/media/image2.png" })).toEqual({
			bytes: PNG,
			type: "image/png"
		})
		reopened.session.close()

		// Undone: the sheet saves as its source.
		changeOf(await session.undo())
		const undone = entries(await savedBytes(session))

		expect(undone["xl/worksheets/sheet2.xml"]).toBe(entries(DRAWINGS_BYTES)["xl/worksheets/sheet2.xml"])
		expect(undone["xl/media/image2.png"]).toBeUndefined()
		expect(undone["xl/drawings/drawing2.xml"]).toBeUndefined()
		session.close()
	})

	it("inserts an SVG with the PNG Office shows where it cannot draw SVG, and appends to an existing drawing", async () => {
		const { session, sheets } = await openDrawings()
		const sheet = sheets[0] ?? 0

		changeOf(
			await session.command({
				type: "drawing.insertImage",
				sheet,
				placement: { from: { row: 0, col: 8, rowOff: 0, colOff: 0 } },
				bytes: SVG,
				fallback: PNG,
				name: "icon.svg",
				alt: ""
			})
		)

		const listed = await drawingsOf(session, sheet)
		const added = listed?.objects.at(-1)

		expect(added?.name).toBe("icon.svg")
		// cNvPr ids are unique in the part: past the source's highest (5).
		expect(entries(await savedBytes(session))["xl/drawings/drawing1.xml"]).toContain('<xdr:cNvPr id="6" name="icon.svg"/>')

		const rendered = added === undefined ? null : await session.query("drawing.render", { sheet, key: added.key, stamp: added.stamp })

		expect(rendered?.kind === "drawing" && rendered.drawing.images.map(image => [image.partName, image.svgPartName])).toEqual([
			["xl/media/image2.png", "xl/media/image3.svg"]
		])
		session.close()
	})

	it("refuses bytes that are not a picture, and an SVG without its PNG", async () => {
		const { session, sheets } = await openDrawings()
		const sheet = sheets[1] ?? 0
		const placement = { from: { row: 0, col: 0, rowOff: 0, colOff: 0 } }

		expect(
			refusalOf(
				await session.command({
					type: "drawing.insertImage",
					sheet,
					placement,
					bytes: new TextEncoder().encode("<html>"),
					name: "",
					alt: ""
				})
			)
		).toBe("unsupported")
		expect(refusalOf(await session.command({ type: "drawing.insertImage", sheet, placement, bytes: SVG, name: "", alt: "" }))).toBe(
			"unsupported"
		)
		session.close()
	})

	it("deletes a picture this session inserted together with its media", async () => {
		const { session, sheets } = await openDrawings()
		const sheet = sheets[0] ?? 0

		changeOf(
			await session.command({
				type: "drawing.insertImage",
				sheet,
				placement: { from: { row: 0, col: 8, rowOff: 0, colOff: 0 } },
				bytes: PNG,
				name: "",
				alt: ""
			})
		)

		const added = (await drawingsOf(session, sheet))?.objects.at(-1)

		if (added === undefined) throw new Error("not inserted")

		changeOf(await session.command({ type: "drawing.delete", sheet, objects: [{ key: added.key, stamp: added.stamp }] }))

		const files = entries(await savedBytes(session))

		expect(files["xl/media/image2.png"]).toBeUndefined()
		expect(files["xl/drawings/_rels/drawing1.xml.rels"]).not.toContain("image2.png")
		session.close()
	})

	it("moves anchors with inserted rows, and the listed objects follow", async () => {
		const { session, sheets } = await openDrawings()
		const sheet = sheets[0] ?? 0

		changeOf(await session.command({ type: "axis.insert", sheet, axis: "rows", at: 0, count: 2 }))

		const drawings = await drawingsOf(session, sheet)

		expect(objectOf(drawings, "Picture 1").anchor).toMatchObject({ from: { row: 3 }, to: { row: 7 } })
		// An absolute anchor stays.
		expect(objectOf(drawings, "Straight Connector 4").anchor).toEqual({ mode: "absolute", x: 0, y: 0, cx: 952500, cy: 0 })

		const picture = objectOf(drawings, "Picture 1")

		expect((await session.query("drawing.render", { sheet, key: picture.key, stamp: picture.stamp }))?.kind).toBe("drawing")
		session.close()
	})

	it("refuses object edits on a sheet that protects its objects", async () => {
		const bytes = withMedia(
			buildFixture({
				...DRAWINGS,
				sheets: [
					{
						name: "One",
						body: '<sheetData/><sheetProtection sheet="1" objects="1"/><drawing r:id="rId1"/>',
						rels: relationship("rId1", "drawing", "../drawings/drawing1.xml")
					}
				]
			}),
			{ "xl/media/image1.png": PNG }
		)
		const { session, sheets } = await openDrawings(bytes)
		const sheet = sheets[0] ?? 0
		const picture = objectOf(await drawingsOf(session, sheet), "Picture 1")

		expect(
			refusalOf(await session.command({ type: "drawing.move", sheet, key: picture.key, stamp: picture.stamp, anchor: MOVED }))
		).toBe("protectedSheet")
		session.close()
	})

	it("deletes a chart frame with its chart part, and an undo brings both back", async () => {
		const { session, sheets } = await openDrawings()
		const sheet = sheets[0] ?? 0
		const chart = objectOf(await drawingsOf(session, sheet), "Chart 3")

		changeOf(await session.command({ type: "drawing.delete", sheet, objects: [{ key: chart.key, stamp: chart.stamp }] }))

		const files = entries(await savedBytes(session))

		expect(files["xl/charts/chart1.xml"]).toBeUndefined()
		expect(files["xl/drawings/_rels/drawing1.xml.rels"]).not.toContain("chart1.xml")
		expect(files["[Content_Types].xml"]).not.toContain("/xl/charts/chart1.xml")

		changeOf(await session.undo())

		const restored = entries(await savedBytes(session))

		expect(restored["xl/charts/chart1.xml"]).toBe(entries(DRAWINGS_BYTES)["xl/charts/chart1.xml"])
		expect(restored["xl/drawings/drawing1.xml"]).toBe(entries(DRAWINGS_BYTES)["xl/drawings/drawing1.xml"])
		session.close()
	})

	it("draws a group's shapes in the group's frame", async () => {
		const group =
			"<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>" +
			"<xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>" +
			'<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="2" name="Group 1"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>' +
			'<xdr:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1219200" cy="381000"/><a:chOff x="0" y="0"/><a:chExt cx="1219200" cy="381000"/></a:xfrm></xdr:grpSpPr>' +
			'<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="3" name="A"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="609600" cy="381000"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></xdr:spPr></xdr:sp>' +
			'<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="4" name="B"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="609600" y="0"/><a:ext cx="609600" cy="381000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></xdr:spPr></xdr:sp>' +
			"</xdr:grpSp><xdr:clientData/></xdr:twoCellAnchor>"
		const bytes = buildFixture({
			...DRAWINGS,
			parts: { ...DRAWINGS.parts, "xl/drawings/drawing1.xml": { text: wsDr(group), type: TYPES.drawing } }
		})
		const { session, sheets } = await openDrawings(bytes)
		const sheet = sheets[0] ?? 0
		const object = objectOf(await drawingsOf(session, sheet), "Group 1")
		const rendered = await session.query("drawing.render", { sheet, key: object.key, stamp: object.stamp })

		expect(object.kind).toBe("group")
		expect(rendered?.kind === "drawing" && rendered.drawing.itemCount).toBe(2)
		session.close()
	})

	it("shows no objects on a sheet whose drawing part does not read, and refuses inserting there", async () => {
		const bytes = withMedia(
			buildFixture({
				...DRAWINGS,
				parts: { ...DRAWINGS.parts, "xl/drawings/drawing1.xml": { text: "<xdr:wsDr", type: TYPES.drawing } }
			}),
			{ "xl/media/image1.png": PNG }
		)
		const { session, sheets, events } = await openDrawings(bytes)
		const sheet = sheets[0] ?? 0

		expect(await drawingsOf(session, sheet)).toBeUndefined()
		expect(events.some(event => event.type === "featureUnavailable" && event.feature === "drawings")).toBe(true)
		expect(
			refusalOf(
				await session.command({
					type: "drawing.insertImage",
					sheet,
					placement: { from: { row: 0, col: 0, rowOff: 0, colOff: 0 } },
					bytes: PNG,
					name: "",
					alt: ""
				})
			)
		).toBe("drawingUnavailable")
		// Saved as it was.
		expect(entries(await savedBytes(session))["xl/drawings/drawing1.xml"]).toBe("<xdr:wsDr")
		session.close()
	})

	it("tells the client the sheet's last object is gone", async () => {
		const { session, sheets } = await openDrawings(
			withMedia(
				buildFixture({
					...DRAWINGS,
					parts: { ...DRAWINGS.parts, "xl/drawings/drawing1.xml": { text: wsDr(CONNECTOR), type: TYPES.drawing } }
				}),
				{
					"xl/media/image1.png": PNG
				}
			)
		)
		const sheet = sheets[0] ?? 0
		const connector = objectOf(await drawingsOf(session, sheet), "Straight Connector 4")
		const change = changeOf(
			await session.command({ type: "drawing.delete", sheet, objects: [{ key: connector.key, stamp: connector.stamp }] })
		)

		// A patch keeps what it leaves out: the extensions come whole, without the drawings.
		expect(change.layouts[0]?.layout.extensions).toEqual({})
		session.close()
	})

	it("reads, moves and inserts in a Strict workbook, writing the file's own namespaces", async () => {
		const { session, sheets } = await openDrawings(strictOf(DRAWINGS_BYTES))
		const one = sheets[0] ?? 0
		const two = sheets[1] ?? 0
		const drawings = await drawingsOf(session, one)

		expect(drawings?.objects.map(object => object.kind)).toEqual(["picture", "shape", "chart", "connector"])
		expect(objectOf(drawings, "Chart 3").chartPart).toBe("xl/charts/chart1.xml")

		await session.intent()

		const picture = objectOf(drawings, "Picture 1")

		changeOf(await session.command({ type: "drawing.move", sheet: one, key: picture.key, stamp: picture.stamp, anchor: MOVED }))
		changeOf(
			await session.command({
				type: "drawing.insertImage",
				sheet: two,
				placement: { from: { row: 0, col: 0, rowOff: 0, colOff: 0 } },
				bytes: PNG,
				name: "",
				alt: ""
			})
		)

		const files = entries(await savedBytes(session))

		expect(files["xl/drawings/drawing2.xml"]).toContain('xmlns:xdr="http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"')
		expect(files["xl/drawings/drawing2.xml"]).toContain('xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"')
		expect(files["xl/drawings/_rels/drawing2.xml.rels"]).toContain(
			'Type="http://purl.oclc.org/ooxml/officeDocument/relationships/image"'
		)
		expect(files["xl/worksheets/_rels/sheet2.xml.rels"]).toContain(
			'Type="http://purl.oclc.org/ooxml/officeDocument/relationships/drawing"'
		)
		expect(files["xl/worksheets/sheet2.xml"]).toContain('<drawing xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"')
		session.close()
	})

	it("moves both branches of an object in alternate content, and deletes the whole of it", async () => {
		const branch = (name: string): string => PICTURE.replace('name="Picture 1"', `name="${name}"`)
		const wrapped = `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><mc:Choice Requires="a14">${branch("Choice")}</mc:Choice><mc:Fallback>${branch("Fallback")}</mc:Fallback></mc:AlternateContent>`
		const bytes = withMedia(
			buildFixture({
				...DRAWINGS,
				parts: { ...DRAWINGS.parts, "xl/drawings/drawing1.xml": { text: wsDr(wrapped), type: TYPES.drawing } }
			}),
			{ "xl/media/image1.png": PNG }
		)
		const { session, sheets } = await openDrawings(bytes)
		const sheet = sheets[0] ?? 0
		const object = objectOf(await drawingsOf(session, sheet), "Choice")

		changeOf(await session.command({ type: "drawing.move", sheet, key: object.key, stamp: object.stamp, anchor: MOVED }))

		const drawing = entries(await savedBytes(session))["xl/drawings/drawing1.xml"] ?? ""

		expect(drawing.match(/<xdr:from><xdr:col>2<\/xdr:col><xdr:colOff>0<\/xdr:colOff><xdr:row>10<\/xdr:row>/g)).toHaveLength(2)

		const moved = objectOf(await drawingsOf(session, sheet), "Choice")

		changeOf(await session.command({ type: "drawing.delete", sheet, objects: [{ key: moved.key, stamp: moved.stamp }] }))
		expect(entries(await savedBytes(session))["xl/drawings/drawing1.xml"]).not.toContain("AlternateContent")
		session.close()
	})
})
