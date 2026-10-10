// The core type-checks against ESNext only, so it runs unchanged in a worker, a WebView and Hermes. These
// are the runtime globals all of them provide, declared with the minimal shape the modules use. Anything
// else that fails to resolve is a platform leak and must not enter the core.

declare class TextEncoder {
	encode(input?: string): Uint8Array
}
