export { CfbError, type CfbErrorCode } from "./errors"
export {
	CfbFile,
	DEFAULT_CFB_LIMITS,
	compareNames,
	isCompoundFile,
	type CfbEntry,
	type CfbEntryKind,
	type CfbLimits,
	type CfbOpenOptions
} from "./reader"
export { cfbBytes, planCfb, writeCfb, type CfbPlan, type CfbWriteEntry, type CfbWriteOptions } from "./writer"
