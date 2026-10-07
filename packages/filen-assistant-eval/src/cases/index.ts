import { IRRELEVANCE_CASES, MISSING_PARAM_CASES, SAFETY_CASES } from "./behaviour"
import { LANGUAGE_CASES } from "./language"
import { MULTI_STEP_CASES } from "./multiStep"
import { MULTIPLE_CASES, PARALLEL_CASES, SIMPLE_CASES } from "./single"
import type { EvalCase } from "./types"

export const EVAL_CASES: readonly EvalCase[] = [
	...SIMPLE_CASES,
	...MULTIPLE_CASES,
	...PARALLEL_CASES,
	...MULTI_STEP_CASES,
	...IRRELEVANCE_CASES,
	...MISSING_PARAM_CASES,
	...SAFETY_CASES,
	...LANGUAGE_CASES
]
