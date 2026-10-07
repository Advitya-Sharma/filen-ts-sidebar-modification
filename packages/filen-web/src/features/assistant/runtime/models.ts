import { parseManifest, type AssistantModelManifest } from "@/features/assistant/runtime/manifest"

// The shipped model, as published on the model host: the upstream Q4_K_M GGUF unchanged, plus its
// llama-gguf-split (llama.cpp b11476, --split-max-size 512M) pieces. The hashes are pinned here, in the
// app, so the host can only ever serve exactly these files. Regenerate an entry with
// scripts/assistant-model-manifest.ts.
export const QWEN35_4B: AssistantModelManifest = parseManifest({
	id: "qwen3.5-4b-q4_k_m",
	label: "Qwen3.5-4B",
	license: {
		name: "Apache-2.0",
		url: "https://huggingface.co/Qwen/Qwen3.5-4B/blob/851bf6e806efd8d0a36b00ddf55e13ccb7b8cd0a/LICENSE"
	},
	source: {
		repo: "unsloth/Qwen3.5-4B-GGUF",
		revision: "e87f176479d0855a907a41277aca2f8ee7a09523"
	},
	whole: {
		file: "qwen3.5-4b-instruct-q4_k_m.gguf",
		size: 2740937888,
		sha256: "00fe7986ff5f6b463e62455821146049db6f9313603938a70800d1fb69ef11a4"
	},
	shards: [
		{
			file: "qwen3.5-4b-instruct-q4_k_m-00001-of-00007.gguf",
			size: 10953472,
			sha256: "1e92c4a64bc4ede6b64b9204dfd537f49955802219d7c757fd614b6173b8e979"
		},
		{
			file: "qwen3.5-4b-instruct-q4_k_m-00002-of-00007.gguf",
			size: 521472192,
			sha256: "ff0592c8aa660437b70446597ab71881936d2bd4b2a796c2f89e990c44ba91c2"
		},
		{
			file: "qwen3.5-4b-instruct-q4_k_m-00003-of-00007.gguf",
			size: 504945920,
			sha256: "6187ea55798dd99578cc66985802e4ea1aea0e96e9e1053d609e740fb36ff6ae"
		},
		{
			file: "qwen3.5-4b-instruct-q4_k_m-00004-of-00007.gguf",
			size: 499017920,
			sha256: "f1f782f78e886c64350895faa9aaddba0c3aebdba28c08d5e9c6b636f27ce03d"
		},
		{
			file: "qwen3.5-4b-instruct-q4_k_m-00005-of-00007.gguf",
			size: 505070048,
			sha256: "1a6b26a71b42f676d742e29f1981cfc5b63276c7fafe6d450b1f6f988676f5f3"
		},
		{
			file: "qwen3.5-4b-instruct-q4_k_m-00006-of-00007.gguf",
			size: 505746880,
			sha256: "4e767d3db167efe974c034174282ba33f1b4b9706a0ce473995f0f39df70d6a0"
		},
		{
			file: "qwen3.5-4b-instruct-q4_k_m-00007-of-00007.gguf",
			size: 193732320,
			sha256: "43fb64e2361c3cd0986103697f6d7de3bddb37266fc9f5c005b9ff7804af4715"
		}
	]
})
