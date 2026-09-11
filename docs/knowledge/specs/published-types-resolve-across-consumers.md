# Published types must resolve for TypeScript, ESM, and CommonJS consumers

Cluster: types

## Scenario

A developer installs the library and imports it from a TypeScript project, and separately from plain
ESM and CommonJS projects. They expect editor intellisense for the public API and expect `tsc` to
type-check consuming code against the shipped declarations. In some editor and toolchain combinations
the type information failed to resolve, so autocompletion and type checking silently degraded even
though the runtime import worked. The failure was environment-dependent, present in some editors and
not others, and tied to how the package manifest points at its entry and its declarations.

> Spec note, not a corpus case: this is a packaging and manifest guarantee verified by a build/CI
> consumer-smoke check and type-level tests, not by the runtime roundtrip adapter. The upstream
> thread was a PR against a deleted Grunt/Babel/browserify build; the diff is discarded and the
> requirement kept.

## Desired behavior

- The published package exposes its public API types so they resolve **correctly and identically**
  for three consumer shapes, TypeScript, ESM and CommonJS, regardless of `moduleResolution`
  (`node`, `node16`/`nodenext`, `bundler`). Editor intellisense and `tsc` both pick up the shipped
  declarations.
- An explicit `exports` map declares `types`, `import`, and `require` conditions so node16 and nodenext
  resolution succeeds, with `types` and `typesVersions` fallbacks covering older resolvers.
- The `.d.ts`, `.d.mts` and `.d.cts` variants match each condition, with no single hand-authored
  declaration that only resolves under one module system.

The same dual-build packaging must also **execute** cleanly, not just type-check. A user importing
the library into an SSR React framework and constructing the workbook class saw a production build
throw _"Cannot call a class as a function"_, a Babel `_classCallCheck`-style guard emitted by a
UMD/CommonJS transpile breaking under a specific framework bundler. The requirement: the package
bundles and instantiates cleanly inside modern framework build pipelines, including SSR React with
their own bundlers, imported statically, dynamically, or across a server/client boundary, with no
class-called-as-function or interop failure. For this fork the class is native ESM TypeScript with no
CommonJS-isms in source, so a `_classCallCheck` guard cannot be emitted by construction. The
remaining work is choosing the published module formats (ESM-only against dual ESM plus CJS) and
declaring `exports` conditions so framework bundlers resolve and run the right build. The CI
consumer-smoke check must therefore _instantiate_ the workbook in each consumer project, not merely
type-check it.

## Open questions

- Ship dual ESM plus CJS, or ESM-only with a compatibility shim? Whichever is chosen, the declaration
  variants must line up with the runtime conditions.
- CI guard: install the built tarball into throwaway TS, ESM, and CJS consumer projects and assert
  both runtime import and type resolution succeed, turning "intellisense works" into a
  machine-checked invariant rather than a manual editor screenshot.
- Prefer type-level tests (`expectTypeOf` or tsd-style) over per-file compile-spec suites.

Related: `public-types-node-stream-portability`, `async-iterable-types-compile-cleanly`,
`write-buffer-return-type-contract`, `browser-safe-io-boundary`.
