// How the runner hands fixtures to the bundle and reads its result back.

// The global the runner's prepended script sets: { fixtures: { name, base64 }[] }.
export const INPUT_GLOBAL = "__OFFICE_HERMES_INPUT__"
// The bundle prints its result as one line starting with this.
export const RESULT_PREFIX = "OFFICE_HERMES_RESULT "
