# Importing the library must not mutate global built-ins

Cluster: browser

Reported as "Angular Material crashes".

## Desired behavior

Importing the library, in any environment and via any published entrypoint, MUST NOT mutate global built-ins (`Promise`, `Array.prototype`, `Object.assign`, `Symbol`, and the rest). Loading the package is a side-effect-free operation with respect to the global scope.

Rationale: a spreadsheet library is a leaf dependency. Host applications and frameworks legitimately replace or instrument global built-ins, and Angular with Zone.js swaps in an instrumented `Promise` to drive change detection while other frameworks patch scheduling primitives. A library that ships eagerly-applied core-js polyfills in its browser bundle silently overwrites those, producing action-at-a-distance breakage that is nearly impossible for consumers to attribute back to the library.

## Prior art (what went wrong upstream)

- The legacy browser bundle was produced by Babel plus core-js and executed `require('core-js/modules/es.promise')` at module-load time, unconditionally assigning `globalThis.Promise`.
- Under Angular this overwrote Zone.js's `ZoneAwarePromise`, so promise settlement stopped notifying the framework: double-click-to-act, stuck spinners, phantom route changes, and a console error `ZoneAwarePromise (window|global).Promise has been overwritten`.
- A `bare` and `bare.min` bundle without the polyfill existed but was not the default entrypoint, forcing consumers to hand-edit resolution and then re-supply the polyfills the Babel output assumed. This is exactly the packaging shape we are forking away from.
- The symptom eventually disappeared for some users across later builds purely as a side effect of build-config drift, not an intentional guarantee, meaning it could regress at any time without a durable rule.

## Requirements for this fork

1. **Ship modern ESM** targeting current runtimes; assume the consumer's toolchain, or a runtime baseline, provides `Promise`, async iterators and the rest. Do not down-compile to ES5.
2. **Zero global mutation at import.** No core-js entry modules, no polyfill side-effects, no assignment to `globalThis.*` on load. If a runtime lacks a needed built-in, that is the consumer's polyfill decision, made _before_ our code runs, not ours.
3. If any optional runtime-feature shim is ever offered, it must be an explicit, opt-in, separate import that the consumer chooses, never bundled into the default entrypoint.
4. A CI guard should assert the built package produces no global side-effects on import, for example by snapshotting key global identities before and after importing the entrypoint and asserting they are unchanged.

## Open questions

- What is the minimum runtime baseline we commit to (Node LTS plus evergreen browsers)? That baseline determines which built-ins we may freely assume without any shim.
- Do we ship a single ESM entrypoint for all environments, or distinct Node and browser conditional exports? Either is fine provided neither mutates globals.
