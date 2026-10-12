import { fork } from "node:child_process"
import { rm } from "node:fs/promises"
import { join, resolve } from "node:path"
import { parseArgs } from "node:util"
import { MITIGATION_PRESETS } from "./agent"
import { LOOP_PRESETS } from "./loopHarness"
import { EVAL_CASES } from "./cases"
import { INJECTION_CASES } from "./cases/injection"
import { CATEGORIES, LANGUAGES } from "./cases/types"
import { ensureModel, loadManifest, PACKAGE_ROOT } from "./models"
import { writeReport } from "./report"
import type { EvalPreset, WorkerJob } from "./worker"

const USAGE = `Usage: pnpm eval <command> [options]

Commands:
  run        Run the evaluation (resumes unless --fresh), then write the report
  download   Download and verify the configured models
  report     Rebuild summary.json and summary.md from existing results
  list       Show the case counts

Options:
  --models <ids>       Comma-separated model ids (default: every enabled model)
  --suite <name>       main | injection | all (default: all)
  --presets <names>    Presets for the injection suite (default: shipped,loop)
  --main-presets <n>   Presets for the main suite (default: shipped)

Presets: shipped (the shipping loop with the tools the web app runs), loop (the shipping loop with every
tool), and the legacy harness's mitigation presets (${Object.keys(MITIGATION_PRESETS).join(", ")}).
  --categories <list>  Main-suite categories to run
  --languages <list>   Languages to run (both suites)
  --cases <ids>        Exact case ids to run
  --sample <n>         Keep every n-th case of the selection (quick subsets)
  --out <dir>          Results directory (default: results/default)
  --models-dir <dir>   Model cache (default: .models)
  --fresh              Delete the selected models' previous results first`

function list(value: string | undefined): string[] | undefined {
	return value === undefined
		? undefined
		: value
				.split(",")
				.map(item => item.trim())
				.filter(item => item.length > 0)
}

const ALL_PRESETS = [...Object.keys(LOOP_PRESETS), ...Object.keys(MITIGATION_PRESETS)] as EvalPreset[]

function presets(value: string | undefined, fallback: EvalPreset[]): EvalPreset[] {
	const names = list(value) ?? fallback

	for (const name of names) {
		if (!(ALL_PRESETS as string[]).includes(name)) {
			throw new Error(`Unknown preset "${name}"; known: ${ALL_PRESETS.join(", ")}`)
		}
	}

	return names as EvalPreset[]
}

function runWorker(job: WorkerJob): Promise<void> {
	return new Promise((resolvePromise, reject) => {
		// One process per model: memory is returned to the OS between models and peak RSS is per model.
		const child = fork(join(PACKAGE_ROOT, "src", "worker.ts"), [JSON.stringify(job)], {
			execArgv: process.execArgv,
			stdio: "inherit"
		})

		child.on("exit", code => (code === 0 ? resolvePromise() : reject(new Error(`Worker for ${job.modelId} exited with ${code}`))))
	})
}

async function main(): Promise<void> {
	const { positionals, values } = parseArgs({
		allowPositionals: true,
		options: {
			models: {
				type: "string"
			},
			suite: {
				type: "string"
			},
			presets: {
				type: "string"
			},
			"main-presets": {
				type: "string"
			},
			categories: {
				type: "string"
			},
			languages: {
				type: "string"
			},
			cases: {
				type: "string"
			},
			sample: {
				type: "string"
			},
			out: {
				type: "string"
			},
			"models-dir": {
				type: "string"
			},
			fresh: {
				type: "boolean"
			},
			help: {
				type: "boolean"
			}
		}
	})
	const command = positionals[0]

	if (command === undefined || values.help === true) {
		console.log(USAGE)

		return
	}

	const outDir = resolve(PACKAGE_ROOT, values.out ?? "results/default")
	const modelsDir = resolve(PACKAGE_ROOT, values["models-dir"] ?? ".models")
	const manifest = await loadManifest()
	const wantedModels = list(values.models)
	const models = manifest.filter(model => (wantedModels === undefined ? model.enabled : wantedModels.includes(model.id)))

	if (wantedModels !== undefined && models.length !== wantedModels.length) {
		throw new Error(`Unknown model in ${wantedModels.join(", ")}; known: ${manifest.map(model => model.id).join(", ")}`)
	}

	const languages = list(values.languages)
	const categories = list(values.categories)
	const caseIds = list(values.cases)
	const sample = Number(values.sample ?? "1")
	const keep = <T extends { id: string; lang: string }>(items: readonly T[], extra: (item: T) => boolean): T[] =>
		items
			.filter(
				item =>
					(languages === undefined || languages.includes(item.lang)) &&
					(caseIds === undefined || caseIds.includes(item.id)) &&
					extra(item)
			)
			.filter((_, index) => index % sample === 0)
	const suite = values.suite ?? "all"
	const mainCases = suite === "injection" ? [] : keep(EVAL_CASES, item => categories === undefined || categories.includes(item.category))
	const injectionCases = suite === "main" ? [] : keep(INJECTION_CASES, () => true)

	if (command === "list") {
		for (const category of CATEGORIES) {
			const items = EVAL_CASES.filter(item => item.category === category)

			console.log(
				`${category}: ${items.length} (${LANGUAGES.map(lang => `${lang} ${items.filter(item => item.lang === lang).length}`).join(", ")})`
			)
		}

		console.log(
			`main total: ${EVAL_CASES.length}; injection: ${INJECTION_CASES.length} in ${new Set(INJECTION_CASES.map(item => item.lang)).size} languages`
		)

		return
	}

	if (command === "report") {
		await writeReport({
			outDir,
			manifest
		})

		return
	}

	if (command !== "download" && command !== "run") {
		throw new Error(`Unknown command "${command}"\n\n${USAGE}`)
	}

	for (const model of models) {
		const modelPath = await ensureModel({
			model,
			dir: modelsDir
		})

		if (command === "download") {
			console.log(`${model.id}: ${modelPath}`)

			continue
		}

		if (values.fresh === true) {
			await rm(join(outDir, model.id), {
				recursive: true,
				force: true
			})
		}

		await runWorker({
			modelId: model.id,
			modelPath,
			outDir,
			main: {
				presets: presets(values["main-presets"], ["shipped"]),
				caseIds: mainCases.map(item => item.id)
			},
			injection: {
				// The two loops that ship; the legacy presets only when named, as research.
				presets: presets(values.presets, ["shipped", "loop"]),
				caseIds: injectionCases.map(item => item.id)
			}
		})
	}

	if (command === "run") {
		await writeReport({
			outDir,
			manifest
		})
	}
}

await main()
