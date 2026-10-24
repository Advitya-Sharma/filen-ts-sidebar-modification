import { defineConfig } from "vitest/config"
import { budget } from "./src/tests/timing"

export default defineConfig({
	test: {
		environment: "node",
		include: ["src/**/*.test.{ts,tsx}"],
		isolate: true,
		clearMocks: true,
		restoreMocks: true,
		unstubEnvs: true,
		unstubGlobals: true,
		// Vitest's 5 s default, scaled for a slower runner like every other bound on time.
		testTimeout: budget(5000),
		hookTimeout: budget(10_000)
	}
})
