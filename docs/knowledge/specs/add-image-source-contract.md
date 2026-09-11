# The image source contract: buffer/path, with URLs handled deliberately

Cluster: images

## Scenario

A user wants to embed an image by pointing at a remote HTTP(S) URL, expecting the library to fetch
the bytes and embed them. Instead the image source is only interpreted as a local file path (Node
only) or an in-memory buffer, so passing a URL as the "filename" silently fails to produce an
embedded picture: no image, no error. The community answer is always the same, which is to fetch the
bytes yourself and pass a buffer.

> Spec note, not a corpus case: the durable question is what an "image source" _is_ and how a
> URL-shaped input should fail. That is a typing and API-contract decision, not a malformed-output
> bug. The Node-only-ness of the filepath source is already noted in `image-by-filename-is-node-only`.

## Desired behavior

The image-adding API has a clear, well-typed contract for what a source is. Today the accepted
sources are effectively a local file path (Node only) and a raw byte buffer; a remote URL is not
supported and passing one produces nothing. Decide deliberately between two positions:

1. **Keep network I/O out of scope** (recommended default). A spreadsheet library should not fetch
   on the caller's behalf. Precisely type the source as `buffer | filepath`, and **reject a
   URL-shaped path with a typed, actionable error** ("image source looks like a URL; fetch the bytes
   and pass a buffer") instead of silently embedding nothing.
2. **Opt-in URL support.** Offer an async helper that accepts a URL, but only via an explicitly
   injected fetcher so the library core stays network-free and the caller controls the request
   (timeouts, auth, SSRF posture).

Either way the failure mode is explicit and the source type is honest, so "pass a URL" is never a
silent no-op.

### Validate the image definition at the point of entry

Beyond the source _kind_, the image definition itself must be validated when it is added, not
discovered as a corrupt package at write time. The add-image entry point rejects, with a typed and
actionable error, a definition that:

- **carries no payload**, meaning neither a buffer nor a resolvable path/base64 source is present;
- **declares an unusable or missing extension/format**, since the extension drives the OOXML content
  type and the media part name, so an empty or unrecognized one silently corrupts the package (see
  `image-missing-extension-corrupts-package`);
- **contradicts itself**, meaning a declared format that disagrees with the actual payload's
  detectable type (magic bytes), where that mismatch would produce a file Excel rejects.

Validating at entry keeps the failure close to the caller's mistake and keeps the writer's invariant
simple: by the time an image reaches serialization, it is already known-well-formed.

## Open questions

- Is a URL ever accepted in core, or always via an injected fetcher in an optional helper?
- What exactly counts as "URL-shaped" for the reject path (a scheme allowlist, or any `://`)?
- Security: any fetch path is untrusted-input-facing (SSRF, redirect, size limits), which is a
  reason to keep it out of core and in a caller-controlled helper.

Related: `image-by-filename-is-node-only`, `image-missing-extension-corrupts-package`,
`load-accepts-arraybuffer-and-typed-arrays`.
