import { describeModelShards } from "../vite/assistant-model"

// Prints the `shards` list of a model manifest (src/features/assistant/runtime/models.ts) for a GGUF, or for
// every piece of a llama-gguf-split model when given its first piece:
//   pnpm exec tsx scripts/assistant-model-manifest.ts /path/to/model-00001-of-00007.gguf
const path = process.argv[2]

if (path === undefined) {
	throw new Error("usage: assistant-model-manifest.ts <model.gguf | model-00001-of-0000N.gguf>")
}

console.log(JSON.stringify(await describeModelShards(path), null, "\t"))
