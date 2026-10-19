export {
	DRM_TRANSFORM_ID,
	STRONG_ENCRYPTION_TRANSFORM_ID,
	classifyDataSpaces,
	type DataSpacesSource,
	type DataSpacesVerdict
} from "./dataSpaces"
export { OfficeCryptoError, type OfficeCryptoErrorCode } from "./errors"
export {
	CERTIFICATE_ENCRYPTOR_URI,
	ENCRYPTION_NAMESPACE,
	MAX_ENCRYPTION_INFO_BYTES,
	MAX_SPIN_COUNT,
	PASSWORD_ENCRYPTOR_URI,
	parseEncryptionInfo,
	type AgileEncryptionInfo,
	type EncryptionInfo,
	type StandardEncryptionInfo
} from "./info"
export {
	AGILE_DEFAULT_SPIN_COUNT,
	DEFAULT_OFFICE_PASSWORD,
	exportOfficeKey,
	importOfficeKey,
	probeEncryptionInfo,
	releaseOfficeKey,
	unlockOfficePackage,
	type AbortSignalLike,
	type OfficeCryptoOptions,
	type OfficeKey,
	type OfficeKeyTransfer,
	type OfficeUnlockOptions,
	type OfficeUnlockResult
} from "./webCrypto"
export {
	encryptedPackageVerdict,
	openEncryptedPackage,
	type DecryptedPackage,
	type EncryptedPackageOpen,
	type EncryptionRefusal,
	type LockedPackage,
	type OpenEncryptedPackageOptions,
	type PackageUnlockOptions,
	type PackageUnlockResult
} from "./sealedPackage"
// Sealing (an editor's or a workbook worker's saves) is @filen/office/dom/encryption/seal.
export { isSealedOfficeFile } from "../../core/encryption/index"
