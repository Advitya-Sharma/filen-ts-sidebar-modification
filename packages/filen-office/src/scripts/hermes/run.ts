import { execFileSync, spawnSync } from "node:child_process"
import { createHash } from "node:crypto"
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs"
import { createRequire } from "node:module"
import { tmpdir } from "node:os"
import { dirname, join, resolve } from "node:path"
import { fileURLToPath } from "node:url"
import { runInNewContext } from "node:vm"
import { decodeBase64, encodeBase64 } from "./base64"
import { INPUT_GLOBAL, RESULT_PREFIX } from "./protocol"

// Runs the core's open/layout/edit/save scenario (scenario.ts) under Hermes and under V8 from one bundle and
// compares the results byte for byte. The bundle is built the way the mobile release build makes its own:
// Metro with filen-mobile's Expo config and Babel preset (engine hermes, profile hermes-stable, not minified),
// then the release hermesc (-O -emit-binary). Hermes cannot read files, so the fixtures are prepended to the
// bundle as a script.
//
//   tsx src/scripts/hermes/run.ts [--hermes <hermes CLI>] [--root <office package>] [--mobile <app>] [--out <dir>] [fixtures...]
//
// Without a Hermes CLI (HERMES_CLI or --hermes) the bundle is only compiled with hermesc, which still checks
// that every module parses and compiles for Hermes.

const here = dirname(fileURLToPath(import.meta.url))
const DEFAULT_FIXTURES = [
	"01-plain.docx",
	"03-header-footer-footnotes.docx",
	"04-comments-tracked.docx",
	"05-fields-toc.docx",
	"06-textbox-shape.docx",
	"07-chart-customxml-unknown.docx",
	"08-large-300p.docx"
]

interface Options {
	readonly root: string
	readonly mobile: string
	readonly out: string
	readonly hermes: string | undefined
	readonly fixtures: readonly string[]
}

function options(argv: readonly string[]): Options {
	let root = resolve(here, "../../..")
	let mobile = resolve(here, "../../../../filen-mobile")
	let out = join(tmpdir(), "office-hermes")
	let hermes = process.env["HERMES_CLI"]
	const fixtures: string[] = []

	for (let i = 0; i < argv.length; i++) {
		const arg = argv[i] ?? ""
		const value = (): string => {
			const next = argv[++i]

			if (next === undefined) throw new Error(`${arg} needs a value`)

			return resolve(next)
		}

		if (arg === "--root") root = value()
		else if (arg === "--out") out = value()
		else if (arg === "--hermes") hermes = value()
		else if (arg === "--mobile") mobile = value()
		else fixtures.push(resolve(arg))
	}

	return {
		root,
		mobile,
		out,
		hermes,
		fixtures: fixtures.length > 0 ? fixtures : DEFAULT_FIXTURES.map(name => join(root, "src/tests/fixtures", name))
	}
}

interface MetroApi {
	loadConfig(argv: { config: string; cwd: string }): Promise<unknown>
	runBuild(config: unknown, options: Record<string, unknown>): Promise<unknown>
}

async function bundle(opts: Options): Promise<string> {
	const mobileRequire = createRequire(join(opts.mobile, "package.json"))
	const expoRequire = createRequire(mobileRequire.resolve("expo/package.json"))
	const metro = expoRequire("metro") as MetroApi
	const presetRequire = createRequire(mobileRequire.resolve("babel-preset-expo/package.json"))
	const babelRuntime = dirname(presetRequire.resolve("@babel/runtime/package.json"))
	const configPath = join(opts.out, "metro.config.cjs")
	const out = join(opts.out, "bundle.js")

	// The app's Expo defaults (resolver, Babel transformer, serializer and polyfills); the office package is
	// watched in case it is a copy outside the workspace; the cache stays in the output directory. Hermes has
	// no async generators, so the preset rewrites them (and for await) into @babel/runtime helper imports,
	// which neither package declares yet: they resolve to the preset's own copy here.
	writeFileSync(
		configPath,
		`const { getDefaultConfig } = require(${JSON.stringify(mobileRequire.resolve("expo/metro-config"))})
const { FileStore } = require(${JSON.stringify(expoRequire.resolve("metro-cache"))})
const config = getDefaultConfig(${JSON.stringify(opts.mobile)})
module.exports = {
	...config,
	watchFolders: [...(config.watchFolders ?? []), ${JSON.stringify(opts.root)}],
	resolver: {
		...config.resolver,
		extraNodeModules: { ...config.resolver?.extraNodeModules, "@babel/runtime": ${JSON.stringify(babelRuntime)} }
	},
	cacheStores: [new FileStore({ root: ${JSON.stringify(join(opts.out, "metro-cache"))} })],
	maxWorkers: 4
}
`
	)

	const config = await metro.loadConfig({ config: configPath, cwd: opts.mobile })

	await metro.runBuild(config, {
		entry: join(opts.root, "src/scripts/hermes/entry.ts"),
		out,
		platform: "ios",
		dev: false,
		minify: false,
		sourceMap: false,
		customTransformOptions: { engine: "hermes", bytecode: "1" },
		unstable_transformProfile: "hermes-stable"
	})

	return out
}

function hermesc(opts: Options): string {
	const mobileRequire = createRequire(join(opts.mobile, "package.json"))
	const rnRequire = createRequire(mobileRequire.resolve("react-native/package.json"))
	const compiler = dirname(rnRequire.resolve("hermes-compiler/package.json"))
	const binary = join(compiler, "hermesc", process.platform === "darwin" ? "osx-bin" : "linux64-bin", "hermesc")

	if (!existsSync(binary)) throw new Error(`No hermesc at ${binary}`)

	return binary
}

interface FixtureOut {
	readonly name: string
	readonly pages: number
	readonly structure: string
	readonly geometry: string
	readonly unedited: string
	readonly uneditedIdentical: boolean
	readonly steps: readonly string[]
	readonly edited: string
	readonly changedParts: readonly string[]
	readonly editedPages: number
	readonly editedStructure: string
	readonly editedGeometry: string
	readonly reopenedPages: number
	readonly reopenedMatches: number
	readonly ms: Readonly<Record<string, number>>
}

interface RunOut {
	readonly fixtures?: readonly FixtureOut[]
	readonly probes?: readonly { readonly name: string; readonly value: string }[]
	readonly errors?: readonly string[]
	readonly fatal?: string
}

function parse(output: string): RunOut {
	const line = output.split("\n").find(text => text.startsWith(RESULT_PREFIX))

	if (line === undefined) throw new Error(`No result in the output:\n${output.slice(0, 4000)}`)

	return JSON.parse(line.slice(RESULT_PREFIX.length)) as RunOut
}

// The bundle in a fresh V8 context: ECMAScript builtins and Intl, plus the CLI's print and timers. No
// TextDecoder, CompressionStream, MessageChannel or Node globals, like the Hermes CLI.
async function runNode(script: string): Promise<{ out: RunOut; ms: number }> {
	const started = performance.now()
	const lines: string[] = []
	const done = Promise.withResolvers<void>()
	const context = {
		print: (line: string) => {
			lines.push(line)

			if (line.startsWith(RESULT_PREFIX)) done.resolve()
		},
		setTimeout,
		clearTimeout,
		setInterval,
		clearInterval,
		TextEncoder
	}

	runInNewContext(script, context, { filename: "bundle.js" })
	await done.promise

	return { out: parse(lines.join("\n")), ms: performance.now() - started }
}

function runHermes(binary: string, bytecode: string): { out: RunOut; ms: number; maxRssMb: number | undefined } {
	const started = performance.now()
	// /usr/bin/time reports the peak resident set size (bytes on macOS, KiB on Linux).
	const timed = existsSync("/usr/bin/time")
	const run = timed
		? spawnSync("/usr/bin/time", [process.platform === "darwin" ? "-l" : "-v", binary, bytecode], {
				encoding: "utf8",
				maxBuffer: 1 << 28
			})
		: spawnSync(binary, [bytecode], { encoding: "utf8", maxBuffer: 1 << 28 })
	const ms = performance.now() - started

	if (run.status !== 0) throw new Error(`hermes exited ${run.status}:\n${run.stderr.slice(-4000)}\n${run.stdout.slice(-2000)}`)

	const rss = /(\d+)\s+maximum resident set size/.exec(run.stderr) ?? /Maximum resident set size \(kbytes\): (\d+)/.exec(run.stderr)
	const maxRssMb = rss?.[1] === undefined ? undefined : Number(rss[1]) / (process.platform === "darwin" ? 1024 * 1024 : 1024)

	return { out: parse(run.stdout), ms, maxRssMb }
}

function sha(base64: string): string {
	return createHash("sha256").update(decodeBase64(base64)).digest("hex").slice(0, 16)
}

function compare(node: RunOut, hermes: RunOut): string[] {
	const problems: string[] = []
	const byName = new Map((hermes.fixtures ?? []).map(fixture => [fixture.name, fixture]))

	for (const ours of node.fixtures ?? []) {
		const theirs = byName.get(ours.name)

		if (theirs === undefined) {
			problems.push(`${ours.name}: missing under Hermes`)

			continue
		}

		for (const key of [
			"pages",
			"structure",
			"geometry",
			"uneditedIdentical",
			"editedPages",
			"editedGeometry",
			"editedStructure",
			"reopenedPages",
			"reopenedMatches"
		] as const) {
			if (ours[key] !== theirs[key])
				problems.push(`${ours.name}: ${key} ${String(ours[key])} (V8) vs ${String(theirs[key])} (Hermes)`)
		}

		if (ours.steps.join() !== theirs.steps.join()) problems.push(`${ours.name}: steps ${ours.steps.join()} vs ${theirs.steps.join()}`)
		if (ours.changedParts.join() !== theirs.changedParts.join()) problems.push(`${ours.name}: changed parts differ`)
		if (ours.unedited !== theirs.unedited) problems.push(`${ours.name}: unedited save bytes differ`)
		if (ours.edited !== theirs.edited) problems.push(`${ours.name}: edited save bytes differ`)
	}

	const probes = new Map((node.probes ?? []).map(probe => [probe.name, probe.value]))

	for (const probe of hermes.probes ?? []) {
		if (probes.get(probe.name) !== probe.value)
			problems.push(`probe ${probe.name}: ${String(probes.get(probe.name))} (V8) vs ${probe.value} (Hermes)`)
	}

	for (const error of hermes.errors ?? []) problems.push(`Hermes error: ${error}`)
	for (const error of node.errors ?? []) problems.push(`V8 error: ${error}`)

	return problems
}

async function main(): Promise<void> {
	const opts = options(process.argv.slice(2))

	mkdirSync(opts.out, { recursive: true })

	let at = performance.now()
	const bundled = await bundle(opts)
	const bundleMs = performance.now() - at
	const input = {
		fixtures: opts.fixtures.map(path => ({ name: path.split("/").pop() ?? path, base64: encodeBase64(readFileSync(path)) }))
	}
	const script = `globalThis.${INPUT_GLOBAL} = ${JSON.stringify(input)};\n${readFileSync(bundled, "utf8")}`
	const scriptPath = join(opts.out, "run.js")
	const bytecode = join(opts.out, "run.hbc")

	writeFileSync(scriptPath, script)
	at = performance.now()
	execFileSync(hermesc(opts), ["-emit-binary", "-max-diagnostic-width=80", "-O", "-out", bytecode, scriptPath], { stdio: "inherit" })

	const compileMs = performance.now() - at
	const node = await runNode(script)
	const report: Record<string, unknown> = {
		bundleBytes: readFileSync(bundled).length,
		bytecodeBytes: readFileSync(bytecode).length,
		bundleMs: Math.round(bundleMs),
		compileMs: Math.round(compileMs),
		v8: { ms: Math.round(node.ms), probes: node.out.probes, errors: node.out.errors, fatal: node.out.fatal }
	}
	let problems: string[] = node.out.fatal === undefined ? [] : [`V8 fatal: ${node.out.fatal}`]

	if (opts.hermes === undefined) {
		problems.push("No Hermes CLI (set HERMES_CLI or pass --hermes): compiled with hermesc only")
	} else {
		const hermes = runHermes(opts.hermes, bytecode)

		report["hermes"] = {
			ms: Math.round(hermes.ms),
			maxRssMb: hermes.maxRssMb,
			probes: hermes.out.probes,
			errors: hermes.out.errors,
			fatal: hermes.out.fatal
		}

		if (hermes.out.fatal !== undefined) problems.push(`Hermes fatal: ${hermes.out.fatal}`)

		problems = [...problems, ...compare(node.out, hermes.out)]
		report["fixtures"] = (hermes.out.fixtures ?? []).map(fixture => {
			const v8 = node.out.fixtures?.find(other => other.name === fixture.name)

			return {
				name: fixture.name,
				pages: fixture.pages,
				uneditedIdentical: fixture.uneditedIdentical,
				editedSha: sha(fixture.edited),
				editedBytesEqual: v8?.edited === fixture.edited,
				steps: fixture.steps,
				changedParts: fixture.changedParts,
				reopenedMatches: fixture.reopenedMatches,
				hermesMs: fixture.ms,
				v8Ms: v8?.ms
			}
		})
	}

	report["problems"] = problems
	writeFileSync(join(opts.out, "report.json"), `${JSON.stringify(report, null, "\t")}\n`)
	process.stdout.write(`${JSON.stringify(report, null, "\t")}\n`)
	process.exitCode = problems.length === 0 ? 0 : 1
}

await main()
