// The file one run of the e2e suite records its shared fixture tree in (e2e/fixtures.ts), keyed by the run's
// id (E2E_RUN_ID, playwright.config.ts): concurrent local runs share the directory, and one run's teardown
// must never remove another's manifest. Only the harness imports this; the app never does.
export function fixtureManifestName(runId: string | undefined): string {
	if (runId === undefined || !/^[A-Za-z0-9_-]{1,64}$/.test(runId)) {
		throw new Error(
			`E2E_RUN_ID is ${runId === undefined || runId === "" ? "unset" : "not a file-name-safe id"}: run the suite through playwright.config.ts`
		)
	}

	return `tree-${runId}.json`
}
