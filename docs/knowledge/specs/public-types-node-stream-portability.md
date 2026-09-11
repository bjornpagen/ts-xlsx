# Public type surface must be portable across Node and browser/bundler tsconfigs

Cluster: types

Reported as "error using exceljs".

## Problem

The legacy library shipped hand-written declarations that referenced Node's stream module via `import('stream').Stream` and the ambient `NodeJS` global namespace, such as `NodeJS.TypedArray`, directly in the public API. In consumer projects that do not surface `@types/node` in the compilation unit that type-checks the library, a very common setup for front-end, Angular and bundler-targeted apps, sometimes with a split `tsconfig.app.json`, type-checking fails hard with:

- `TS2307: Cannot find module 'stream'`
- `TS2503: Cannot find namespace 'NodeJS'`

These errors appear even when the consumer never touches the stream-based APIs, because the failing symbols sit in the same declaration file that describes the whole surface.

## Prior art / observed workarounds

- Adding `"types": ["node"]`, and installing `@types/node`, to the _app-level_ tsconfig such as `tsconfig.app.json` rather than the root `tsconfig.json`, resolves it. But this is a consumer-side band-aid and unintuitive to discover.
- Some users avoided the stream-typed methods entirely.

## Desired behavior for ts-xlsx

1. The public type surface must type-check cleanly in a browser or bundler-targeted project that does **not** have `@types/node` in scope, at least for the non-Node APIs. Importing the library must not, by itself, force every consuming compilation unit to resolve Node ambient globals.
2. Node-specific stream and buffer APIs should be typed against explicitly-imported types, such as `import type { Readable } from 'node:stream'`, that a bundler can tree-shake and a browser build can omit, rather than dynamic `import('stream')` string module references or raw `NodeJS.*` globals sprinkled through the surface. Prefer standard web types (`ReadableStream`, `ArrayBuffer`, `Uint8Array`, `DataView`) for cross-platform APIs, and confine Node types to clearly Node-only entry points.
3. Consider providing separate entry points, browser against node, so the browser surface never references Node types at all.
4. Because ts-xlsx is TypeScript-first with generated rather than hand-written declarations, this class of "declaration file references a type the consumer can't resolve" must be guarded by a type-level test that compiles the public surface under a minimal, Node-types-absent tsconfig.

## Open questions

- Which APIs are genuinely Node-only and which are dual-target, and can they be cleanly split by entry point?
- Do we support DOM `ReadableStream` and Node `Readable` at the same call site, or force explicit per-platform imports?
