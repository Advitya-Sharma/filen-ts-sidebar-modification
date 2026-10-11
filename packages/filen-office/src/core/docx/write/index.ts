export { canonicalElement, canonicalFragment } from "./canonical"
export { NUMBERING_CONTENT_TYPE, abstractNumbering, numberingInstance, planNumbering, type NumberingWrite } from "./numbering"
export { writerProof } from "./proof"
export { PROPERTY_MARKS, elementsOf, sameFormatting, writeRunProperties } from "./runs"
export {
	DocxSaveError,
	nativeDeflate,
	planDocxSave,
	saveDocx,
	serializeDocx,
	type DocxChanges,
	type DocxSavePlan,
	type StoryChange,
	type DocxSaveFailure,
	type DocxSaveInput,
	type DocxSaveOptions,
	type DocxSaveResult,
	type DocxSaveTimings
} from "./save"
export { STYLES_CONTENT_TYPE, hasStyleDefinition, styleDefinition, stylesWithDefinitions } from "./styles"
export {
	DocxWriteError,
	R_NAMESPACE,
	StoryWriter,
	W_NAMESPACE,
	type StoryOutput,
	type StoryWriteOptions,
	type WriteEnvironment,
	type WriteFailure,
	type WrittenContainer,
	type WrittenParagraph
} from "./story"
export {
	DocxVerifyError,
	paragraphSignature,
	textblocks,
	verifyDocxParts,
	verifyDocxZip,
	type VerifyFailure,
	type VerifyPartsInput
} from "./verify"
export { Names, Out } from "./markup"
export * from "./order.generated"
