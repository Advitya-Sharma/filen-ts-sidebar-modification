#!/usr/bin/env bash
#
# Builds the assistant's inference engine, wllama's WebAssembly build of llama.cpp with its WebGPU backend,
# from pinned sources with glue.patch beside this script applied, and writes it with a manifest and the
# licenses of what it contains to src/vendor/wllama/. All of it is committed; the worker loads that
# wllama.wasm instead of the one in the npm package, whose JavaScript (patched through pnpm) it keeps.
#
# glue.patch changes wllama's glue and llama.cpp:
#
# - A load can set llama.cpp's host-memory prompt cache limit (cache_ram_mib). The engine's default, 8 GiB,
#   is twice the 4 GiB a wasm heap can grow to, so a long session ended in std::bad_alloc. The field is
#   appended to the load request, whose fields are read in order: the JavaScript side's prototype for it
#   (patches/@wllama__wllama@<version>.patch) must list the same fields in the same order.
# - Loading the weights passes no whole tensor (up to ~500 MiB) through one buffer: llama.cpp stages a tensor
#   bound for the GPU in 16 MiB pieces, since a wasm heap keeps its largest size until the page closes and
#   Chrome keeps the shared memory it copies a GPU buffer write of 4 MiB or more through, in the page's
#   process and in the GPU process; and reads of the file are requested in 16 MiB pieces, since each arrives
#   as a JavaScript buffer that stays in memory until a garbage collection.
#
# This took the page's process from ~2.5 to ~1.0 GiB and the GPU process from ~3.9 to ~3.3 GiB right after
# loading Qwen3.5-4B on WebGPU, with the same output and loads and turns as fast (the assistant's README has
# the measurements).
#
# The recipe is the one wllama publishes with (its scripts/build_wasm.sh, run in the emscripten/emsdk
# image), minus the compat build, which the app never loads; the prefix map gives the source tree the path
# it has there. Every run first builds the release unpatched and requires it to be byte for byte the npm
# package's wllama.wasm, which proves the toolchain and recipe; then it applies the patch, which recompiles
# only the files it changes, and relinks.
#
# Emscripten 4.0.20's LLVM orders the try blocks of a function in llama-kv-cache.cpp by memory address
# (https://github.com/llvm/llvm-project/issues/204883, fixed after that release), so about one compile of
# that file in four nests them differently from npm's build. The unpatched stage recompiles it, up to
# WLLAMA_ATTEMPTS times (default 10), until the module matches npm's; the patched build then differs from
# npm's by the patch alone and is the same on every run. A run that never matches stops without output.
#
#   scripts/wllama/build-wllama.sh           build, write src/vendor/wllama/
#   scripts/wllama/build-wllama.sh --check   build, compare with the committed wllama.wasm, write nothing
#
# Needs git, curl, shasum, unzip, python3, cmake and make. Emscripten 4.0.20 is installed with emsdk
# (about 2 GB) into the work directory, a temporary one unless WLLAMA_WORKDIR names a directory to keep
# between runs (emsdk, sources, downloads); every run still builds from a clean tree.
#
# To update wllama: set the version and commits below (LLAMA_CPP_COMMIT is the submodule commit of that
# wllama commit), the Emscripten release (emsdk's tag commit and the versions its tools print) and Dawn
# package from its scripts/build_wasm.sh and scripts/docker-compose.yml, and the npm package's
# wllama.wasm hash; check the try nesting still varies only where UNSTABLE_OBJECT says; rebase glue.patch
# and the pnpm patch, build, and run the notices generator.

set -euo pipefail

export LC_ALL=C

WLLAMA_VERSION="3.8.1"
WLLAMA_REPOSITORY="https://github.com/ngxson/wllama"
WLLAMA_COMMIT="7ed17361caf221a84aa1e80ca85a3d6324f3af85"
LLAMA_CPP_REPOSITORY="https://github.com/ggml-org/llama.cpp"
LLAMA_CPP_COMMIT="46ca246de9bb1c35269722a6240d37d9dfd79cad"
EMSCRIPTEN_VERSION="4.0.20"
EMSDK_REPOSITORY="https://github.com/emscripten-core/emsdk"
# The commit of emsdk's tag for that release.
EMSDK_COMMIT="e4fe26ef59168ff44f4c23c466e497bf60b3411e"
# Exactly as that release's emcc, clang and wasm-opt print their versions.
EMCC_VERSION="4.0.20 (6913738ec5371a88c4af5a80db0ab42bad3de681)"
CLANG_VERSION="clang version 22.0.0git (https:/github.com/llvm/llvm-project d163988dd2833f28fbca8c144265108d25ae7bd2)"
WASM_OPT_VERSION="wasm-opt version 124 (version_123-495-g6d5fed324)"
DAWN_TAG="v20260317.182325"
DAWN_PACKAGE_URL="https://github.com/google/dawn/releases/download/${DAWN_TAG}/emdawnwebgpu_pkg-${DAWN_TAG}.zip"
# The digest GitHub lists for the release asset.
DAWN_PACKAGE_SHA256="8dcae86c630d76b6794b271c6572becba36f4237d132bee121a56638c3d76575"
# The esm/wasm/wllama.wasm of the published package.
UPSTREAM_WASM_SHA256="551f13439366eff11e057b1b91fe46a7fdac29d56d67372d4c5f19a549dfe70c"

CMAKE_FLAGS=(-DGGML_WEBGPU=ON -DGGML_WEBGPU_JSPI=ON -DWLLAMA_TEST_BACKEND=OFF)
PREFIX_MAP="-ffile-prefix-map=<wllama source>=/source"
# Relative to the build directory: the object whose try nesting varies (see above).
UNSTABLE_OBJECT="llama.cpp/src/CMakeFiles/llama.dir/llama-kv-cache.cpp.o"

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PACKAGE="$(cd "${HERE}/../.." && pwd)"
OUT_DIR="${PACKAGE}/src/vendor/wllama"
PATCH="${HERE}/glue.patch"
ATTEMPTS="${WLLAMA_ATTEMPTS:-10}"

mode="build"

case "${1:-}" in
	"") ;;
	--check) mode="check" ;;
	*)
		echo "usage: $0 [--check]" >&2
		exit 2
		;;
esac

need() {
	command -v "$1" > /dev/null || {
		echo "missing $1" >&2
		exit 1
	}
}

for tool in git curl shasum unzip python3 cmake make; do
	need "${tool}"
done

hash_of() {
	shasum -a 256 "$1" | cut -d ' ' -f 1
}

if [[ -n "${WLLAMA_WORKDIR:-}" ]]; then
	work="${WLLAMA_WORKDIR}"
	mkdir -p "${work}"
else
	work="$(mktemp -d)"
	trap 'rm -rf "${work}"' EXIT
fi

work="$(cd "${work}" && pwd)"

# Fetched by commit (git verifies the content against it).
checkout() {
	local dir="$1" repository="$2" commit="$3"

	if [[ ! -d "${dir}/.git" ]]; then
		git init --quiet "${dir}"
	fi

	if [[ "$(git -C "${dir}" rev-parse --verify --quiet "${commit}^{commit}" || true)" != "${commit}" ]]; then
		git -C "${dir}" fetch --quiet --depth 1 "${repository}.git" "${commit}"
	fi

	git -C "${dir}" -c advice.detachedHead=false checkout --quiet --force "${commit}"
}

# The sources, reset to a clean tree.
fetch() {
	checkout "$@"
	git -C "$1" clean --quiet -ffdx
}

# Emscripten, from emsdk's release of that version. emsdk does not verify what it downloads; the versions
# its binaries print, and the unpatched build matching npm's byte for byte, do.
emsdk="${work}/emsdk"

checkout "${emsdk}" "${EMSDK_REPOSITORY}" "${EMSDK_COMMIT}"
"${emsdk}/emsdk" install "${EMSCRIPTEN_VERSION}" > /dev/null
"${emsdk}/emsdk" activate "${EMSCRIPTEN_VERSION}" > /dev/null
# shellcheck disable=SC1091
EMSDK_QUIET=1 source "${emsdk}/emsdk_env.sh" > /dev/null

toolchain="$(dirname "$(command -v emcc)")/../bin"

expect_version() {
	local tool="$1" found="$2" expected="$3"

	if [[ "${found}" != "${expected}" ]]; then
		echo "${tool}: ${found} found, ${expected} expected" >&2
		exit 1
	fi
}

emcc_version="$(emcc --version | sed -n '1s/^.*) \([0-9]\)/\1/p')"

expect_version emcc "${emcc_version}" "${EMCC_VERSION}"
expect_version clang "$("${toolchain}/clang" --version | head -n 1)" "${CLANG_VERSION}"
expect_version wasm-opt "$("${toolchain}/wasm-opt" --version)" "${WASM_OPT_VERSION}"

download() {
	local url="$1" file="$2"

	if [[ ! -f "${file}" ]]; then
		curl -fsSL --retry 3 --proto '=https' -o "${file}.part" "${url}"
		mv "${file}.part" "${file}"
	fi
}

src="${work}/wllama"

fetch "${src}" "${WLLAMA_REPOSITORY}" "${WLLAMA_COMMIT}"

submodule="$(git -C "${src}" ls-tree HEAD llama.cpp | awk '{ print $3 }')"

if [[ "${submodule}" != "${LLAMA_CPP_COMMIT}" ]]; then
	echo "wllama ${WLLAMA_COMMIT} pins llama.cpp ${submodule}, not ${LLAMA_CPP_COMMIT}" >&2
	exit 1
fi

fetch "${src}/llama.cpp" "${LLAMA_CPP_REPOSITORY}" "${LLAMA_CPP_COMMIT}"

dawn_zip="${work}/emdawnwebgpu_pkg-${DAWN_TAG}.zip"

download "${DAWN_PACKAGE_URL}" "${dawn_zip}"
echo "${DAWN_PACKAGE_SHA256}  ${dawn_zip}" | shasum -a 256 -c - > /dev/null

build="${src}/build"
dawn="${build}/emdawn/emdawnwebgpu_pkg"

mkdir -p "${build}/emdawn"
unzip -q "${dawn_zip}" -d "${build}/emdawn"

prefix_map="-ffile-prefix-map=${src}=/source"
jobs="$(getconf _NPROCESSORS_ONLN)"

make_engine() {
	(cd "${build}" && emmake make wllama -j "${jobs}" >> "${work}/make.log")
}

(
	cd "${build}"
	emcmake cmake .. "${CMAKE_FLAGS[@]}" "-DEMDAWNWEBGPU_DIR=${dawn}" \
		"-DCMAKE_C_FLAGS=${prefix_map}" "-DCMAKE_CXX_FLAGS=${prefix_map}" > "${work}/cmake.log"
)

: > "${work}/make.log"
make_engine

# The release as published, before the patch.
attempt=1

while [[ "$(hash_of "${build}/wllama.wasm")" != "${UPSTREAM_WASM_SHA256}" ]]; do
	if [[ ${attempt} -ge ${ATTEMPTS} ]]; then
		echo "unpatched build: differs from the npm package's wllama.wasm after ${attempt} attempts" >&2
		exit 1
	fi

	rm -f "${build}/${UNSTABLE_OBJECT}"
	make_engine
	attempt=$((attempt + 1))
done

echo "unpatched build: identical to the npm package's wllama.wasm (attempt ${attempt})"

git -C "${src}" apply --whitespace=nowarn "${PATCH}"
make_engine

output="$(hash_of "${build}/wllama.wasm")"

echo "patched build: $(wc -c < "${build}/wllama.wasm" | tr -d ' ') bytes, SHA-256 ${output}"

recorded=""

if [[ -f "${OUT_DIR}/manifest.json" ]]; then
	recorded="$(python3 -I -c 'import json, sys; print(json.load(open(sys.argv[1]))["output"]["wllama.wasm"]["sha256"])' "${OUT_DIR}/manifest.json")"
fi

if [[ "${mode}" == "check" ]]; then
	if [[ "$(hash_of "${OUT_DIR}/wllama.wasm")" != "${recorded}" ]]; then
		echo "the committed wllama.wasm is not the one its manifest records" >&2
		exit 1
	fi

	if [[ "${output}" != "${recorded}" ]]; then
		echo "against the committed wllama.wasm: different" >&2
		exit 1
	fi

	echo "against the committed wllama.wasm: identical"
	exit 0
fi

# The license files of what the module contains: llama.cpp and ggml, nlohmann/json (the one library
# llama.cpp vendors under terms that ask for attribution; stb_image, miniaudio and the sha256 code are in
# the public domain), Dawn's C++ WebGPU headers, and Emscripten (its runtime, and the C WebGPU
# implementation in Dawn's package) with musl.
emscripten_root="$(em-config EMSCRIPTEN_ROOT)"
json_version="$(awk 'match($0, /version [0-9][0-9.]*$/) { print substr($0, RSTART + 8); exit }' "${src}/llama.cpp/vendor/nlohmann/json.hpp")"
musl_version="$(cat "${emscripten_root}/system/lib/libc/musl/VERSION")"

mkdir -p "${OUT_DIR}"
rm -rf "${OUT_DIR}/licenses"
mkdir -p "${OUT_DIR}/licenses"
install -m 644 "${build}/wllama.wasm" "${OUT_DIR}/wllama.wasm"
cp "${src}/llama.cpp/LICENSE" "${OUT_DIR}/licenses/llama.cpp-LICENSE"
cp "${src}/llama.cpp/licenses/LICENSE-jsonhpp" "${OUT_DIR}/licenses/nlohmann-json-LICENSE"
cp "${dawn}/webgpu_cpp/LICENSE" "${OUT_DIR}/licenses/dawn-LICENSE"
cp "${emscripten_root}/LICENSE" "${OUT_DIR}/licenses/emscripten-LICENSE"
cp "${emscripten_root}/system/lib/libc/musl/COPYRIGHT" "${OUT_DIR}/licenses/musl-COPYRIGHT"

json_string() {
	local value="${1//\\/\\\\}"

	printf '"%s"' "${value//\"/\\\"}"
}

json_array() {
	local first=1

	printf '['
	for item in "$@"; do
		[[ ${first} -eq 1 ]] || printf ', '
		first=0
		json_string "${item}"
	done
	printf ']'
}

licenses_json() {
	local first=1

	for file in "${OUT_DIR}"/licenses/*; do
		[[ ${first} -eq 1 ]] || printf ',\n'
		first=0
		printf '\t\t"%s": "%s"' "$(basename "${file}")" "$(hash_of "${file}")"
	done
}

cat > "${OUT_DIR}/manifest.json" << EOF
{
	"description": "wllama's WebAssembly build of llama.cpp with the WebGPU backend, built by scripts/wllama/build-wllama.sh with scripts/wllama/glue.patch applied.",
	"wllama": {
		"version": "${WLLAMA_VERSION}",
		"repository": "${WLLAMA_REPOSITORY}",
		"commit": "${WLLAMA_COMMIT}"
	},
	"llama.cpp": {
		"repository": "${LLAMA_CPP_REPOSITORY}",
		"commit": "${LLAMA_CPP_COMMIT}"
	},
	"nlohmann/json": {
		"version": $(json_string "${json_version}")
	},
	"dawn": {
		"tag": "${DAWN_TAG}",
		"package": "${DAWN_PACKAGE_URL}",
		"sha256": "${DAWN_PACKAGE_SHA256}"
	},
	"emcc": $(json_string "${emcc_version}"),
	"emscripten": {
		"musl": $(json_string "${musl_version}")
	},
	"cmake": $(json_array "${CMAKE_FLAGS[@]}" "-DCMAKE_C_FLAGS=${PREFIX_MAP}" "-DCMAKE_CXX_FLAGS=${PREFIX_MAP}"),
	"sources": {
		"build-wllama.sh": "$(hash_of "${HERE}/build-wllama.sh")",
		"glue.patch": "$(hash_of "${PATCH}")"
	},
	"licenses": {
$(licenses_json)
	},
	"upstream": {
		"wllama.wasm": "${UPSTREAM_WASM_SHA256}"
	},
	"output": {
		"wllama.wasm": {
			"bytes": $(wc -c < "${OUT_DIR}/wllama.wasm" | tr -d ' '),
			"sha256": "${output}"
		}
	}
}
EOF

if [[ "${output}" == "${recorded}" ]]; then
	echo "wrote ${OUT_DIR}/wllama.wasm (the binary committed before), licenses/ and manifest.json"
else
	echo "wrote ${OUT_DIR}/wllama.wasm (a new binary), licenses/ and manifest.json"
fi
