// Ways of driving an XML scan: the pull generators a reader writes a `for..of` over, the push
// adapter and the multi-pass driver that lets several readers share one scan of a part, the
// verbatim subtree capture, and the gatherer that reassembles one element's text.
//
// The scan itself, and the readings of a single attribute's text, are `xml-scan.ts`. The split is
// by what a helper does with the source rather than by who calls it: everything here walks a
// document, nothing there does. `elementSubtrees` is the reason the scanner exports its markup and
// tag primitives rather than keeping them private - it is a second scanner over the same text, and
// the two must agree to the character about where an element ends.

import {XmlParseError} from './errors.ts';
import {NamespaceScope} from './xml-namespaces.ts';
import {
  localName,
  markupAt,
  parseAttributes,
  type Tag,
  tagAt,
  type XmlAttributes,
  type XmlEvent,
  xmlEvents,
} from './xml-scan.ts';

export interface SaxHandlers {
  /**
   * An element start. `selfClosing` is true for `<x/>`; no matching {@link onClose} fires for it.
   *
   * `scope` carries the namespace bindings in force here, for the handful of readings whose identity
   * is a namespace rather than a prefix: an `r:id` under whatever prefix the file bound the
   * relationships namespace to, or telling an extension element from a main-namespace one in a file
   * that prefixes the main namespace. Most handlers match on {@link localName} and ignore it.
   */
  onOpen(name: string, attrs: XmlAttributes, selfClosing: boolean, scope: NamespaceScope): void;
  /** A run of character data (already entity-decoded; CDATA delivered verbatim). Omit to ignore text. */
  onText?(text: string): void;
  /** An element end (`</x>`); the synthetic end of a self-closing element is *not* reported here.
   * Omit to ignore closes. `scope` is the element's own bindings, still in force: it is popped after
   * this returns, so a close handler resolves the same namespaces its open handler did. */
  onClose?(name: string, scope: NamespaceScope): void;
}

/**
 * What {@link elementSubtrees} is to capture: for each container element's local name, the local name
 * of the children to take verbatim inside it (`'dxfs' -> 'dxf'`). Scoping the child to a container is
 * what keeps a `<color>` in `<mruColors>` from being confused with the many other `<color>` elements
 * a stylesheet carries.
 */
export type SubtreeSelection = ReadonlyMap<string, string>;

/** What {@link elementSubtrees} captured: the verbatim source of each selected child, keyed by its
 * container's local name, and each container's own attributes as written. */
export interface SubtreeCapture {
  readonly fragments: ReadonlyMap<string, readonly string[]>;
  readonly attributes: ReadonlyMap<string, XmlAttributes>;
}

/**
 * Every element tag in `source`, in document order, with the offsets an offset-based reader needs:
 * where the `<` sits, the tag as {@link tagAt} split it, and its namespace-stripped local name.
 * Markup that is not a tag -- a comment, a CDATA section, a processing instruction, a declaration --
 * is stepped over through {@link markupAt} and never yielded.
 *
 * The three functions below work on offsets rather than events, so none can go through
 * {@link xmlEvents}, which hands back decoded payloads and no positions. That is the whole of what
 * they share, and it is the part that must not drift: they are run over the same part
 * (`parseStyleTable` puts two over `xl/styles.xml`), so a form one walk skipped and another did not
 * would make them disagree about where an element ends. `markupAt` and `tagAt` are shared for that
 * reason; the loop that drives them is shared for the same one.
 *
 * Deliberately private: each of the three below builds a different state machine on top of this, and
 * publishing the walk would invite another offset scanner rather than another caller of these.
 */
function* rawTags(source: string): Generator<{lt: number; tag: Tag; local: string}> {
  const length = source.length;
  let i = 0;
  while (i < length) {
    const lt = source.indexOf('<', i);
    if (lt === -1) break;
    const markup = markupAt(source, lt);
    if (markup !== undefined) {
      i = markup.next;
      continue;
    }
    const tag = tagAt(source, lt);
    // Advanced before the yield, so a consumer that `continue`s still moves the scan on.
    i = tag.next;
    yield {lt, tag, local: localName(tag.name)};
  }
}

/**
 * Capture the verbatim source text of selected elements, in one scan.
 *
 * Some content is re-emitted byte for byte rather than modelled: a differential style, a custom
 * indexed palette, an author's recent-colour swatches, a table-style definition. Preserving the raw
 * text is what keeps a foreign `<dxf>`'s number format a real format code across a re-write instead
 * of a coerced `"[object Object]"`, so the reader needs a subtree's *source*, which an event stream
 * by definition cannot hand back.
 *
 * That gap is why four callers each grew a `<container>([\s\S]*?)</container>` scanner of their own,
 * which is a regular expression parsing XML, over untrusted input, in the same directory as the
 * reader written specifically to avoid that (ADR-0004). This is the same capability done properly:
 * one linear scan with comments, CDATA, processing instructions and declarations skipped as markup
 * rather than matched as text, and the nesting depth counted so a same-named descendant does not end
 * a capture early.
 *
 * Only a container's *first* occurrence is read, matching the single block these documents declare;
 * a second is ignored rather than merged. A captured element that never closes throws
 * {@link XmlParseError}, like the reader's other truncation cases: a partial subtree re-emitted
 * verbatim is broken markup handed on as though it were content.
 */
export function elementSubtrees(source: string, selection: SubtreeSelection): SubtreeCapture {
  const fragments = new Map<string, string[]>();
  const attributes = new Map<string, XmlAttributes>();
  // A container is read once: closed containers are recorded so a later one of the same name is
  // skipped rather than appended to.
  const finished = new Set<string>();
  // The container currently open, the child name to capture inside it, and how deep the same
  // container name is nested within itself.
  let container: {local: string; child: string; depth: number} | undefined;
  // The child element being captured: where its `<` sits, and how deep its own name is nested inside
  // it, so `</name>` for a descendant does not end the capture.
  let capture: {local: string; start: number; depth: number} | undefined;

  for (const {lt, tag, local} of rawTags(source)) {
    if (tag.close) {
      if (capture !== undefined && local === capture.local) {
        if (capture.depth > 0) capture.depth -= 1;
        else {
          fragments.get(container?.local ?? '')?.push(source.slice(capture.start, tag.next));
          capture = undefined;
        }
      } else if (capture === undefined && container !== undefined && local === container.local) {
        if (container.depth > 0) container.depth -= 1;
        else {
          finished.add(container.local);
          container = undefined;
        }
      }
      continue;
    }

    const {selfClosing} = tag;
    if (capture !== undefined) {
      if (!selfClosing && local === capture.local) capture.depth += 1;
    } else if (container !== undefined) {
      if (!selfClosing && local === container.local) container.depth += 1;
      else if (local === container.child) {
        if (selfClosing) fragments.get(container.local)?.push(source.slice(lt, tag.next));
        else capture = {local, start: lt, depth: 0};
      }
    } else {
      const child = selection.get(local);
      if (child !== undefined && !finished.has(local)) {
        attributes.set(local, parseAttributes(tag.attrSource));
        if (selfClosing) finished.add(local);
        else {
          fragments.set(local, fragments.get(local) ?? []);
          container = {local, child, depth: 0};
        }
      }
    }
  }

  if (capture !== undefined) {
    throw new XmlParseError(`unterminated <${capture.local}> element`);
  }
  return {fragments, attributes};
}

/**
 * Where an element sits in the source: the offsets an editor splices at, and the parsed attributes of
 * its opening tag.
 *
 * `contentStart`/`contentEnd` bound the element's children; for an empty element they are equal and
 * both sit just past the `/>`, so replacing that range turns `<latin/>` into `<latin>…</latin>` and
 * a caller that means to replace the *element* uses `start`/`end` instead.
 */
export interface ElementRange {
  /** Offset of the element's `<`. */
  readonly start: number;
  /** One past the element's final `>`. */
  readonly end: number;
  /** One past the opening tag's `>`. */
  readonly contentStart: number;
  /** Offset of the closing tag's `<`, or `end` for an empty element. */
  readonly contentEnd: number;
  /** The name as written, namespace prefix included. */
  readonly name: string;
  readonly attrs: XmlAttributes;
}

/**
 * Locate the first element reachable by a path of local names, as offsets into the source.
 *
 * The primitive for *editing* a part rather than reading one: everything outside the returned range
 * is spliced through byte for byte, so an edit changes what it names and nothing else -- the
 * whitespace, the comments, the attribute order, the prefix the source chose. That is a guarantee no
 * re-serialisation can make, and it is why this returns offsets instead of text.
 *
 * The alternative it replaces is a `<container>[\s\S]*?</container>` regular expression, which is a
 * regular expression parsing XML over untrusted input: it terminates on a `</container>` inside a
 * comment or a CDATA section, it cannot see an element written `<x></x>` where it expected `<x/>`,
 * and its container name is interpolated into a pattern where an NCName's legal `.` and `-` are
 * metacharacters. This scan classifies markup with {@link markupAt}, so a comment is skipped rather
 * than matched, and compares parsed names rather than raw text.
 *
 * Each step of `path` matches at any depth below the previous one, and the *first* match wins, which
 * is the single block these documents declare. An element that never closes throws
 * {@link XmlParseError}, on the same grounds as the other truncation cases here: an unterminated
 * range is not a range.
 */
export function elementRange(source: string, path: readonly string[]): ElementRange | undefined {
  // How much of `path` is currently satisfied, and where each satisfied step's element began. The
  // last entry is the element whose range is being measured once `matched === path.length`.
  const open: {local: string; depth: number}[] = [];
  let pending:
    | {start: number; contentStart: number; name: string; attrs: XmlAttributes}
    | undefined;

  for (const {lt, tag, local} of rawTags(source)) {
    if (tag.close) {
      const innermost = open.at(-1);
      if (innermost !== undefined && innermost.local === local) {
        if (innermost.depth > 0) innermost.depth -= 1;
        else {
          open.pop();
          if (pending !== undefined && open.length === path.length - 1) {
            return {...pending, end: tag.next, contentEnd: lt};
          }
        }
      }
      continue;
    }

    if (pending === undefined && local === path[open.length]) {
      const last = open.length === path.length - 1;
      if (last) {
        if (tag.selfClosing) {
          return {
            start: lt,
            end: tag.next,
            contentStart: tag.next,
            contentEnd: tag.next,
            name: tag.name,
            attrs: parseAttributes(tag.attrSource),
          };
        }
        pending = {
          start: lt,
          contentStart: tag.next,
          name: tag.name,
          attrs: parseAttributes(tag.attrSource),
        };
        open.push({local, depth: 0});
      } else if (!tag.selfClosing) {
        open.push({local, depth: 0});
      }
    } else {
      // Either inside the element being measured, or on a name no step of the path wants. Both are
      // the same question: does this tag deepen the innermost open element's nesting in itself?
      const innermost = open.at(-1);
      if (innermost !== undefined && !tag.selfClosing && innermost.local === local) {
        innermost.depth += 1;
      }
    }
  }

  if (pending !== undefined) {
    throw new XmlParseError(`unterminated <${pending.name}> element`);
  }
  return undefined;
}

/**
 * Locate every element with the local name `local`, at any depth, as offsets into the source, in
 * document order: {@link elementRange} for a part that repeats the element an edit is after, such as
 * the `<c:f>` behind each of a chart's series. A same-named element nested in a match is part of that
 * match's content rather than a range of its own. An element that never closes throws
 * {@link XmlParseError}.
 */
export function* elementRanges(source: string, local: string): Generator<ElementRange> {
  let pending:
    | {start: number; contentStart: number; name: string; attrs: XmlAttributes}
    | undefined;
  // How many same-named elements are open inside the pending one.
  let depth = 0;
  for (const {lt, tag, local: tagLocal} of rawTags(source)) {
    if (tagLocal !== local) continue;
    if (pending !== undefined) {
      if (tag.close) {
        if (depth > 0) depth -= 1;
        else {
          const open = pending;
          pending = undefined;
          yield {...open, end: tag.next, contentEnd: lt};
        }
      } else if (!tag.selfClosing) {
        depth += 1;
      }
      continue;
    }
    if (tag.close) continue;
    const attrs = parseAttributes(tag.attrSource);
    if (tag.selfClosing) {
      yield {
        start: lt,
        end: tag.next,
        contentStart: tag.next,
        contentEnd: tag.next,
        name: tag.name,
        attrs,
      };
    } else {
      pending = {start: lt, contentStart: tag.next, name: tag.name, attrs};
      depth = 0;
    }
  }
  if (pending !== undefined) {
    throw new XmlParseError(`unterminated <${pending.name}> element`);
  }
}

/** An element start surfaced by {@link openElements}: its qualified `name`, the namespace-stripped
 * `local` name the filter matched on, and its already-decoded `attrs`. */
export interface OpenElement {
  readonly name: string;
  readonly local: string;
  readonly attrs: XmlAttributes;
  /**
   * The namespace bindings in force at this element, for the attributes and elements whose identity
   * is a namespace rather than a prefix (`r:id`, the x14 extension elements). One shared, mutating
   * instance rather than a snapshot: it is valid while this element is the current one, which is the
   * whole of a `for..of` body, and copying it per element would cost every scan for the few that ask.
   */
  readonly scope: NamespaceScope;
}

/**
 * Yield each element start in `source` as an {@link OpenElement}, optionally restricted to the given
 * local names. This is the pull shape for the ubiquitous "scan opens, read attributes" pass: a caller
 * writes a plain `for..of` and reads `attrs` directly, instead of threading a mutable accumulator out
 * through a {@link parseXml} `onOpen` closure. With no names every start is yielded; with one or more,
 * only starts whose namespace-stripped name matches one. Text and close events are skipped.
 *
 * Throws {@link XmlParseError} on malformed markup.
 */
export function* openElements(source: string, ...localNames: string[]): Generator<OpenElement> {
  const filter = localNames.length > 0 ? new Set(localNames) : undefined;
  const scope = new NamespaceScope();
  for (const event of xmlEvents(source)) {
    if (event.kind === 'close') {
      scope.close();
      continue;
    }
    if (event.kind !== 'open') continue;
    scope.open(event.attrs);
    const local = localName(event.name);
    if (filter === undefined || filter.has(local)) {
      yield {name: event.name, local, attrs: event.attrs, scope};
    }
    if (event.selfClosing) scope.close();
  }
}

/**
 * Wrap an {@link XmlEvent} stream so a self-closing `<x/>` whose local name is in `names` is
 * presented as an open (with `selfClosing: false`) immediately followed by a close: the exact
 * event shape of `<x></x>`. This lets a consumer commit such an element from its close handling
 * alone, instead of hand-coding a parallel self-closing branch: {@link xmlEvents} fires no close
 * for `<x/>`, and forgetting that branch silently drops the empty element. Names not in the set
 * pass through untouched, so an element whose close would wrongly act on absent content (an empty
 * `<v/>`/`<f/>` that must not commit captured text) is left as a bare self-closing open.
 */
export function* closeEmptyElements(
  events: Iterable<XmlEvent>,
  names: ReadonlySet<string>,
): Generator<XmlEvent> {
  for (const event of events) {
    if (event.kind === 'open' && event.selfClosing && names.has(localName(event.name))) {
      yield {kind: 'open', name: event.name, attrs: event.attrs, selfClosing: false};
      yield {kind: 'close', name: event.name};
    } else {
      yield event;
    }
  }
}

/** Options for {@link parseXml}. */
export interface ParseXmlOptions {
  /**
   * Local names whose self-closing form should also fire {@link SaxHandlers.onClose}. For each name
   * listed, `<x/>` is delivered as an open followed by a close (via {@link closeEmptyElements}), so a
   * handler commits the element once in `onClose` rather than duplicating that logic in a self-closing
   * branch. Only name elements the handler is safe to run on close when empty.
   */
  readonly closeEmptyElements?: ReadonlySet<string>;
}

/**
 * One reader's share of a parse: the handlers it wants the events delivered to, and the self-closing
 * elements it needs expanded. Named separately from {@link ParseXmlOptions} because several readers
 * of the same part run over a single parse of it, and each has to bring its own requirements rather
 * than have the caller remember them; see {@link parseXmlPasses}.
 */
export interface SaxPass {
  readonly handlers: SaxHandlers;
  /** As {@link ParseXmlOptions.closeEmptyElements}, for this reader's elements. */
  readonly closeEmptyElements?: ReadonlySet<string>;
}

/** A {@link SaxPass} that gathers something during the parse rather than committing as it goes. */
export interface CollectingPass<T> extends SaxPass {
  /** What the pass collected. Meaningful only once the parse driving it has finished. */
  result(): T;
}

/**
 * Parse `source` once, delivering every event to each pass in turn. The alternative, a parse per
 * reader, costs a full scan of the document per reader and finds nothing in most of them: the
 * worksheet part is the largest in a package, and reading it five times over spent 45% of a large
 * file's read on four scans that matched no element.
 *
 * The expansions the passes ask for are unioned, so *every* pass sees `<x/>` as an open plus a close
 * for any name *any* of them named. That is the one way a pass can observe that it is sharing a
 * parse, and it is why the option is a set of element names rather than a flag: a pass sees an extra
 * close only for elements another pass had to name, and reaching a close for an element a reader
 * does not handle is already the ordinary case.
 */
export function parseXmlPasses(source: string, passes: readonly SaxPass[]): void {
  const expanded = new Set<string>();
  for (const pass of passes) {
    for (const name of pass.closeEmptyElements ?? []) expanded.add(name);
  }
  const handlers = passes.map((pass) => pass.handlers);
  parseXml(
    source,
    {
      onOpen(name, attrs, selfClosing, scope) {
        for (const handler of handlers) handler.onOpen(name, attrs, selfClosing, scope);
      },
      onText(text) {
        for (const handler of handlers) handler.onText?.(text);
      },
      onClose(name, scope) {
        for (const handler of handlers) handler.onClose?.(name, scope);
      },
    },
    expanded.size > 0 ? {closeEmptyElements: expanded} : undefined,
  );
}

/**
 * Parse an XML document, dispatching SAX events to `handlers`. A thin push adapter over
 * {@link xmlEvents}: one scanning core serves both the callback and the pull consumers.
 * Throws {@link XmlParseError} on malformed markup.
 */
export function parseXml(source: string, handlers: SaxHandlers, options?: ParseXmlOptions): void {
  const events = options?.closeEmptyElements
    ? closeEmptyElements(xmlEvents(source), options.closeEmptyElements)
    : xmlEvents(source);
  // One scope for the whole parse, shared by every pass over it: namespace bindings are a property of
  // the document, not of any one reader, and threading a scope per pass would have each of them
  // re-deriving the same thing from the same attributes.
  const scope = new NamespaceScope();
  for (const event of events) {
    switch (event.kind) {
      case 'open':
        scope.open(event.attrs);
        handlers.onOpen(event.name, event.attrs, event.selfClosing, scope);
        // A self-closing element fires no close, unless `closeEmptyElements` expanded it into a pair,
        // in which case the close below does the popping instead.
        if (event.selfClosing) scope.close();
        break;
      case 'text':
        handlers.onText?.(event.text);
        break;
      case 'close':
        // Closed after the handler, so a close handler still sees the bindings its element declared.
        handlers.onClose?.(event.name, scope);
        scope.close();
        break;
    }
  }
}

/**
 * Gathers one element's character data across the open/text/close events a SAX parse delivers it in.
 *
 * Nine parsers used to open-code this: latch a flag and clear a buffer on the open, append every
 * chunk while latched, consume the buffer and unlatch on the close. Six spellings of one idea, and
 * none of them honoured the one thing {@link SaxHandlers.onOpen} warns about. A self-closing `<x/>`
 * fires no matching close, so `<t/>`, `<text/>`, `<xm:f/>` and `<totalsRowFormula/>`, all legal and
 * all written by real files, latched a capture that nothing would ever close. What kept that from
 * corrupting anything was the order the next open happened to reset things in, which is an accident
 * rather than a property anyone chose, on a path that reads untrusted input. Taking `selfClosing`
 * here makes it structural, once.
 *
 * The other thing the open-coded versions disagreed on is what an unrelated element opening
 * mid-capture should do. Ending the capture is never what a caller wants: the text belongs to the
 * element that opened it, and a nested or sibling element is not that element. So an open that is
 * not for a captured name leaves a capture in progress alone, and {@link close} answers only for
 * the element that started it.
 *
 * Decoding stays outside. A `<t>` needs `decodeSpreadsheetText` over the whole element and never
 * over a chunk, an `<xm:f>` needs nothing, and a coordinate needs a number: the caller knows which.
 */
export class TextCapture {
  readonly #names: ReadonlySet<string>;
  #capturing: string | undefined;
  #text = '';

  /** @param names the element local name, or the set of names this instance may capture. */
  constructor(names: string | Iterable<string>) {
    this.#names = new Set(typeof names === 'string' ? [names] : names);
  }

  /** Whether a capture is currently open. */
  get capturing(): boolean {
    return this.#capturing !== undefined;
  }

  /** Begin capturing `local` if it is one of this instance's names and is not self-closing. */
  open(local: string, selfClosing: boolean): void {
    if (selfClosing || !this.#names.has(local)) return;
    this.#capturing = local;
    this.#text = '';
  }

  /** Feed a chunk of character data; ignored when no capture is open. */
  text(chunk: string): void {
    if (this.#capturing !== undefined) this.#text += chunk;
  }

  /** The gathered text when `local` closes the captured element, else `undefined`. Unlatches. */
  close(local: string): string | undefined {
    if (this.#capturing !== local) return undefined;
    this.#capturing = undefined;
    return this.#text;
  }

  /**
   * Abandon any capture in progress, for a caller whose own container has just ended or begun.
   *
   * {@link close} unlatches only for the element that opened the capture, which is the right answer
   * within one element but not across a boundary the caller owns: markup that opens a `<t>` and is
   * then truncated leaves the capture armed, and the next container's text lands in the abandoned
   * buffer instead of where that container wanted it. A caller that resets its own per-container
   * state resets this with it.
   */
  reset(): void {
    this.#capturing = undefined;
    this.#text = '';
  }
}

/**
 * Yield each named element's text as that element closes, as `{local, text}`.
 *
 * The third member of the pull-shaped family beside {@link openElements} ("scan opens, read
 * attributes") and {@link closeEmptyElements}: this one is "capture these elements' text, tell me
 * each as it closes". A parser whose whole job is reading a handful of text elements out of a part
 * writes a `for..of` over it instead of a {@link parseXml} handler triple whose open and text arms
 * are the same three lines every time.
 *
 * It is deliberately not for every {@link TextCapture} caller. A parser that interleaves capture
 * with per-element state of its own (a `<dataValidation>` gathering formulae, a `<tableColumn>`
 * attaching a totals formula to the column it is inside) needs the open and attribute events too,
 * and stays bespoke; forcing it through here would trade a handler triple for a second pass.
 *
 * A self-closing `<x/>` carries no text and fires no close, so it yields nothing, which is the
 * behaviour {@link TextCapture} exists to make structural rather than a branch each caller
 * remembers.
 *
 * Throws {@link XmlParseError} on malformed markup.
 */
export function* capturedText(
  source: string,
  names: string | Iterable<string>,
): Generator<{local: string; text: string}> {
  const capture = new TextCapture(names);
  for (const event of xmlEvents(source)) {
    if (event.kind === 'open') {
      capture.open(localName(event.name), event.selfClosing);
    } else if (event.kind === 'text') {
      capture.text(event.text);
    } else {
      const local = localName(event.name);
      const text = capture.close(local);
      if (text !== undefined) yield {local, text};
    }
  }
}
