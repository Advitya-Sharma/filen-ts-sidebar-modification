# @filen/office

Document engines and editors shared by the Filen apps. `pnpm run lint`, `pnpm run typecheck` and `pnpm test` are what CI runs.

## Real-world corpus

The corpus tests (`src/tests/*Corpus.test.ts`) run the engines over 510 licensed real-world documents. The files are never committed; `corpus/manifest.json` lists each one with its download URL, licence, size and SHA-256, and `corpus/README.md` lists the sources and licences.

```sh
pnpm run corpus:fetch   # downloads ~84 MB into .corpus/ (gitignored); re-runs only fetch what is missing
pnpm test               # the corpus tests now run instead of skipping
```

- Every file is verified against its SHA-256 before it is written; files no longer in the manifest are removed.
- `--strict` fails on any unavailable file (by default up to 10 may be), `--check-remote` also checks that the sources of cached files still exist, `--dir <directory>` fetches elsewhere.
- The tests look for the corpus in `FILEN_OFFICE_CORPUS` (must exist when set), then `.corpus/`, then the gitignored `packages/filen-web/docs/office/corpus` where it was first assembled.
- `FILEN_OFFICE_UNICODE_FULL=1` adds the full Unicode conformance files, downloaded by `pnpm exec tsx src/scripts/unicodeTables.ts`.
- `pnpm run corpus:manifest --from <MANIFEST.json>` regenerates `corpus/` when the corpus is rebuilt.

CI runs all of it in `.github/workflows/office-corpus.yml`, on changes to this package and nightly.
