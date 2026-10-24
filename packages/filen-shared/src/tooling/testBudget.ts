// Bounds on time in the office tests and the web e2e specs. Through Reflect: some callers' projects have no node
// typings.

// How many times a developer's machine the bounds on time are: FILEN_PERF_BUDGET_SCALE, set where the runner is
// slower (1 when unset). A complexity regression still exceeds the scaled bounds.
function budgetScale(): number {
	const proc: unknown = Reflect.get(globalThis, "process")
	const env: unknown = typeof proc === "object" && proc !== null ? Reflect.get(proc, "env") : undefined
	const raw: unknown = typeof env === "object" && env !== null ? Reflect.get(env, "FILEN_PERF_BUDGET_SCALE") : undefined
	const scale = typeof raw === "string" ? Number(raw) : Number.NaN

	return Number.isFinite(scale) && scale > 0 ? scale : 1
}

const BUDGET_SCALE = budgetScale()

// A bound on time, in ms, scaled for the machine running the tests.
export function budget(ms: number): number {
	return ms * BUDGET_SCALE
}
