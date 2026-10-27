# Contributing to cTable

Thank you for your interest in improving **cTable**, a pandas-inspired VBA tabular-data class library for Excel. Contributions are welcome for consideration by the maintainer, including bug reports, reproducible test cases, documentation, performance improvements, and focused code changes.

Contributions are welcome through GitHub issues and pull requests. For substantial changes, open an issue to discuss the proposed approach before beginning implementation. The maintainer determines which changes are accepted and when they are released.

## Before contributing

1. Read [README.md](README.md), [Installation](docs/INSTALLATION.md), and [Architecture](docs/ARCHITECTURE.md).
2. Search existing issues, if accessible, before opening a new one.
3. For substantial behavior or API changes, discuss the design with the maintainer before implementing it.
4. Never include confidential workbooks, organizational financial data, credentials, or real records in fixtures or diagnostic output.

## How to report a bug

Include the method and parameters involved, expected versus actual behavior, exact VBA error information (if any), a minimal **synthetic** reproduction, and the Excel version and platform. State whether lineage, type inference, or external references are enabled where relevant.

Avoid attaching sensitive production workbooks. A small invented worksheet with a few rows and expected results is preferable.

## Suggested change workflow

1. Create a feature or fix branch from the current default branch (normally `main`).
2. Make a focused change, preserving the public API unless a breaking change is explicitly accepted.
3. Add or update VBA regression tests and provide the expected outputs.
4. Compile the VBA project and execute relevant tests in an actual Excel/VBA host; indicate clearly if either step could not be performed.
5. Update affected documentation and the `[Unreleased]` section of [CHANGELOG.md](CHANGELOG.md).
6. Open a pull request with the motivation, behavior change, test results, platform tested, and any compatibility concerns.

Suggested commit descriptions should explain the change's purpose, not just which files were edited.

## Source and VBA conventions

- Treat `src/*.cls` and `src/*.bas` as exported VBA modules. **Preserve** their `VERSION`, `BEGIN`, `END`, and `Attribute` records when applicable.
- Import classes using the VBA editor's **File > Import File** workflow; do not paste exported files into blank classes if doing so drops attributes.
- In particular, **`cLineageRegistry.cls` requires `Attribute VB_PredeclaredId = True`**. Code uses its default instance; creating a separate registry with `New` is not supported.
- Keep `Option Explicit`; favor descriptive identifiers and documented parameter meanings.
- Public row/column positional arguments use **1-based indexes**, as defined by the library.
- Check whether a method mutates its receiver or returns a separate `cTable`. Use `Set` for VBA object assignments.
- Preserve case-insensitive comparison behavior where currently specified, and handle `Null` explicitly.
- Do not make broad line-ending or automated reformatting changes as part of a functional PR. In particular, avoid accidentally rewriting every exported VBA file as a side effect of Git normalization.

## Public API and compatibility

- Avoid changing existing names, parameter names, positions, return types, or defaults without documenting downstream impacts. VBA callers may use named arguments.
- For new parameters, consider optional defaults and backwards compatibility.
- For filters, test `Null` behavior in both inclusion and exclusion paths.
- For joins and append operations, test mismatched headers, duplicate keys, missing values, and column alignment.
- For sorted/grouped operations, test the effect of input row order and invalidation of cached group state.
- For snapshot history, test baseline behavior, additions/removals, changed attributes, repeated entity keys, and stale caches.

## Lineage and data provenance

Lineage is a functional contract, not merely optional metadata. When adding or modifying an operation, decide whether output values originate from source cells, computed expressions, manual writes, literals, or missing results.

- Preserve the correct source ID across filtering, reordering, projection, joins, append, and cloning.
- Use existing `cTable`/`cLineage` helpers rather than constructing independent lineage registries.
- If exact provenance cannot be determined, report **Unknown** instead of asserting an incorrect source.
- Test both lineage-enabled and lineage-disabled execution paths.
- Enable lineage **before** loading data in tests that expect workbook provenance.

See [Lineage](docs/LINEAGE.md) and [Architecture](docs/ARCHITECTURE.md).

## Dependencies and platform compatibility

The bundled baseline contains early-bound references to **Microsoft Scripting Runtime** and Excel object-model types. Do not claim that a change works without manual references or works on both Windows and macOS merely because it compiles on one platform.

When testing, record your actual environment (Excel version/build, Windows or macOS, and VBA references). Flag platform-specific Windows APIs, macOS `#If Mac` paths, file handling, and unavailable COM components. Dependency removal or late binding should be a separate, tested change.

## Testing expectations

At minimum, perform a **Debug > Compile VBAProject** and run focused checks covering the changed methods. Where relevant, include:

- Empty input and single/multiple-row tables.
- Named and numeric column references.
- `Null`, empty strings, mismatched data types, and conversion errors.
- Source files with headers, varying column order, and nonmatching schemas.
- Stable source attribution after filters, joins, appends, and sorts.
- Cross-platform behavior if a change touches external file/system APIs.

A reviewer must be able to reproduce the results. Never label static review or AI-generated expected output as an executed test.

## Documentation and release notes

Update method examples and defaults whenever the API changes. Add noteworthy user-facing changes to [CHANGELOG.md](CHANGELOG.md) under `[Unreleased]`. Update architecture and lineage documentation when component responsibilities, data flow, or provenance semantics change. Avoid recording every internal refactor in the changelog.

Formal releases and version tags are assigned by the maintainer.

## Licensing and third-party code

Original cTable contributions are intended for distribution under the repository's [MIT License](LICENSE). By submitting a contribution, you affirm that you have the rights needed to submit it on those terms. This policy is not a copyright assignment.

**Do not remove or replace third-party notices.** In particular, `src/JsonConverter.bas` is derived from VBA-JSON and contains third-party copyright and BSD licensing text. Preserve those notices, and consult [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md) and `licenses/VBA-JSON-LICENSE` before modifying or redistributing that module. Discuss any new dependency or copied third-party code with the maintainer and document its license.

## AI-assisted changes

AI tools may assist with drafts, refactoring, and tests, but the contributor remains responsible for examining generated code, confirming methods against the current source, validating licensing/provenance, and **running tests** in Excel. Do not claim an AI suggestion was compiled or executed when it was not.

## Review and acceptance

Maintainer: **Brian A. Crawford**. Acceptance, release timing, and compatibility decisions remain at the maintainer's discretion. A submitted pull request does not guarantee inclusion.
