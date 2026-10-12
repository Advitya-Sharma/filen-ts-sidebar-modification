export { canonicalElement, canonicalFragment } from "./canonical"
export { NUMBERING_CONTENT_TYPE, abstractNumbering, numberingInstance, planNumbering, type NumberingWrite } from "./numbering"
export { writerProof } from "./proof"
export { PROPERTY_MARKS, elementsOf, sameFormatting, writeRunProperties } from "./runs"
export { DocxSaveError, isDocxSaveFailure, type DocxSaveFailure } from "./failure"
export {
	planDocxSave,
	prepareDocxSave,
	saveDocx,
	serializeDocx,
	type DocxChanges,
	type DocxPrepareOptions,
	type DocxPreparedSave,
	type DocxSavePlan,
	type StoryChange,
	type DocxSaveInput,
	type DocxSaveOptions,
	type DocxSaveResult,
	type DocxSaveTimings
} from "./save"
export {
	changedCrcs,
	nativeDeflate,
	verifyDocxZip,
	zipDocxSave,
	type DocxPartChanges,
	type DocxZipOptions,
	type DocxZipResult
} from "./zip"
export { STYLES_CONTENT_TYPE, hasStyleDefinition, styleDefinition, stylesWithDefinitions } from "./styles"
export {
	DocxWriteError,
	R_NAMESPACE,
	StoryWriter,
	W_NAMESPACE,
	type StoryDraft,
	type StoryOutput,
	type StoryWriteOptions,
	type WriteEnvironment,
	type WriteFailure,
	type WrittenContainer,
	type WrittenParagraph
} from "./story"
export { DocxVerifyError, paragraphSignature, textblocks, verifyDocxParts, type VerifyFailure, type VerifyPartsInput } from "./verify"
export { Names, Out } from "./markup"
export * from "./order.generated"
