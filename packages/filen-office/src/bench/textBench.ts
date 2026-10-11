import { PerformanceObserver } from "node:perf_hooks"
import { Alignment, LayoutCursor, TextLayout } from "../core/text/index"
import { bidiClassTable, lineBreakTable } from "../core/text/unicode"
import { GEOMETRY, harness } from "../tests/textFixtures"

// Text layout throughput on 10k three-run paragraphs (~65k justified lines at 624 px), with a fake
// measurer: once with a small vocabulary (every width cached) and once with every word unique (every
// word measured once). Reports time per line and GC activity (scavenges approximate allocation).
//
//   node --expose-gc --import tsx src/bench/textBench.ts

const PARAGRAPHS = 10_000
const VOCABULARY = [
	"lorem",
	"ipsum",
	"dolor",
	"sit",
	"amet",
	"consectetur",
	"adipiscing",
	"elit",
	"sed",
	"do",
	"eiusmod",
	"tempor",
	"Größe",
	"naïve"
]

function corpus(unique: boolean): string[][] {
	const paragraphs: string[][] = []
	let serial = 0

	for (let p = 0; p < PARAGRAPHS; p++) {
		const runs: string[] = []

		for (let r = 0; r < 3; r++) {
			const words: string[] = []

			for (let w = 0; w < 20; w++) {
				const word = VOCABULARY[(p * 7 + r * 5 + w * 3) % VOCABULARY.length] ?? ""

				words.push(unique ? `${word}${(serial++).toString(36)}` : word)
			}

			runs.push(`${words.join(" ")} `)
		}

		paragraphs.push(runs)
	}

	return paragraphs
}

let scavenges = 0
let gcTime = 0

new PerformanceObserver(list => {
	for (const entry of list.getEntries()) {
		scavenges++
		gcTime += entry.duration
	}
}).observe({ entryTypes: ["gc"] })

async function run(name: string, paragraphs: readonly string[][], passes: number): Promise<void> {
	const h = harness()
	const bold = h.style({ letterSpacing: 0.5 })
	const geometry = { ...GEOMETRY, width: 624, alignment: Alignment.BOTH }
	const out = new TextLayout()
	const cursor = new LayoutCursor()

	for (let pass = 0; pass < passes; pass++) {
		let lines = 0
		let prepare = 0
		let layout = 0

		scavenges = 0
		gcTime = 0

		const started = performance.now()

		for (const runs of paragraphs) {
			h.builder.reset(h.base)

			for (let r = 0; r < runs.length; r++) h.builder.append(runs[r] ?? "", r === 1 ? bold : h.base)

			const a = performance.now()
			const prepared = h.layouter.prepare(h.builder.finish(), { rtl: false })
			const b = performance.now()

			out.reset()
			cursor.reset()
			h.layouter.layout(prepared, geometry, out, cursor, 0)
			prepare += b - a
			layout += performance.now() - b
			lines += out.lineCount
		}

		const elapsed = performance.now() - started

		// GC entries arrive asynchronously.
		await new Promise(resolve => setImmediate(resolve))

		console.log(
			`${name} pass ${pass}: ${lines} lines ${elapsed.toFixed(1)} ms (${((elapsed * 1000) / lines).toFixed(2)} µs/line; prepare ${prepare.toFixed(1)}, layout ${layout.toFixed(1)}), ${h.measurer.strings} strings measured, ${scavenges} GCs ${gcTime.toFixed(1)} ms`
		)
	}
}

const tables = performance.now()
const expanded = lineBreakTable().byteLength + bidiClassTable().byteLength

console.log(`unicode tables expanded in ${(performance.now() - tables).toFixed(2)} ms, ${(expanded / 1024).toFixed(1)} KiB`)

await run("cached vocabulary", corpus(false), 3)
await run("unique words", corpus(true), 2)
