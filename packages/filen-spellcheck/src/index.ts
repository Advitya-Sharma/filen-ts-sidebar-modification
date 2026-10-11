// Everything here is free of the engine: the engine (wasm), the worker and the page host are their own
// entry points (./engine, ./worker, ./host), so importing language data never pulls in a worker.
export { DICTIONARY_IDS, dictionaryFor, type DictionaryId } from "./languages"
export {
	WORD_SEPARATOR,
	type SpellcheckAsset,
	type SpellcheckAssets,
	type SpellcheckDictionaryAssets,
	type SpellcheckPort,
	type SpellcheckRequest,
	type SpellcheckResponse,
	type SpellcheckWorkerScope
} from "./protocol"
export {
	MAX_USER_WORDS,
	MAX_USER_WORD_LENGTH,
	USER_DICTIONARY_FILE_NAME,
	USER_DICTIONARY_VERSION,
	UserDictionaryVersionError,
	addUserWord,
	emptyUserDictionary,
	hasUserWord,
	mergeUserDictionaries,
	normalizeUserWord,
	parseUserDictionary,
	removeUserWord,
	sameUserDictionary,
	serializeUserDictionary,
	userWords,
	type UserDictionaryDoc
} from "./userDictionary"
