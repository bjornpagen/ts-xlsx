// Turn a `customUI` ribbon-customisation part into a typed, read-only view of the ribbon tree.
//
// A `.xlsm` can carry two ribbon parts hung off the *package root* rels: `customUI/customUI.xml`
// (Office 2007 RibbonX, namespace `.../2006/01/customui`) and `customUI/customUI14.xml` (Office 2010+,
// namespace `.../2009/07/customui`, which adds backstage/QAT/commands). Both are preserved verbatim on
// round-trip (see `src/core/preserved.ts`); this module adds a reader on top, mirroring how
// `parseVbaProject` projects a read view over bytes the writer already round-trips opaquely.
//
// Scope (v1): the `<ribbon>` subtree only, tabs → groups → controls, plus each control's callback
// names (the `onAction` a click invokes is the whole reason a macro workbook ships a ribbon). A
// customUI document's `<commands>`, `<backstage>`, `<contextMenus>`, and the ribbon's `qat`/
// `contextualTabs` are NOT parsed; they still round-trip byte-for-byte, they are just not surfaced here.
//
// Security posture: the part is untrusted input. The underlying `xmlEvents` scanner never expands
// entities or DTDs (see `xml-scan.ts`); on top of that this parser caps nesting depth and fails closed
// with {@link CustomUiParseError} on any malformed or unrecognised structure rather than returning a
// half-built tree.

import {strFromU8} from 'fflate';

import {isRelType} from '../rel-type.ts';
import {tokenSet} from '../token-set.ts';
import {boolStrict, localName, type XmlAttributes, xmlEvents} from '../xml/xml-scan.ts';
import {CustomUiParseError} from './errors.ts';

/** The `customUI` schema a part is written against. The read model keys off this, not the (frequently
 * mis-copied) relationship type. `2007` is the original RibbonX (`customUI.xml`); `2010` is the later
 * schema (`customUI14.xml`) that also carries backstage/QAT/commands. */
export type RibbonDialect = '2007' | '2010';

/** The `customUI` root namespaces, one per {@link RibbonDialect}. */
export const CUSTOMUI_2006_NAMESPACE = 'http://schemas.microsoft.com/office/2006/01/customui';
export const CUSTOMUI_2009_NAMESPACE = 'http://schemas.microsoft.com/office/2009/07/customui';

// The OPC relationship Type URIs Office wires the two ribbon parts under, from the package-root rels.
// Both end `/ui/extensibility` (the 2010 one confusingly carries `2007` in its path); the reader
// matches on that suffix, in {@link isCustomUiRelType}, through the same `isRelType` the codecs use
// on every other preserved relationship type.
export const CUSTOMUI_2007_REL_TYPE =
  'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility';
export const CUSTOMUI_2010_REL_TYPE =
  'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility';

/** Whether a package-root relationship Type URI points at a `customUI` ribbon part. */
export function isCustomUiRelType(type: string): boolean {
  return isRelType(type, 'ui/extensibility');
}

/**
 * A control element inside a ribbon group. `kind` is the element's local name, narrowed to the closed
 * set of RibbonX control elements ({@link RibbonControlKind}); an element outside that set is surfaced
 * as `unknown` rather than dropped. The three identity attributes (`id` a document-defined control,
 * `idQ` a qualified id, `idMso` a built-in control) and the two most-consulted display/behaviour
 * attributes (`label`, `onAction`) are lifted out as typed conveniences; every attribute the element
 * actually carried, including the many `get*` dynamic callbacks and layout hints not modelled here,
 * is preserved verbatim in {@link attributes}, so nothing is lost. Container controls (a `menu`,
 * `splitButton`, `gallery`, `dropDown`, `box`, …) carry their nested controls/items in {@link children}.
 */
export interface RibbonControl {
  readonly kind: RibbonControlKind;
  /** A document-defined control id. */
  readonly id?: string;
  /** A namespace-qualified control id (`idQ`), used to reference a control across add-ins. */
  readonly idQ?: string;
  /** The id of a built-in (Microsoft-defined) control this element repurposes or places against. */
  readonly idMso?: string;
  /** The static label, when the element carries one (a dynamic label uses `getLabel`, in {@link attributes}). */
  readonly label?: string;
  /** The callback procedure name invoked on activation: the macro a click runs. */
  readonly onAction?: string;
  /** Every attribute on the element, verbatim and entity-decoded. The typed fields above are lifted from
   * here; this map is the complete record, including attributes this model does not lift out. */
  readonly attributes: Readonly<Record<string, string>>;
  /** Nested controls or items, for a container control; absent for a leaf control. */
  readonly children?: readonly RibbonControl[];
}

/** The RibbonX control elements this reader recognises. `item` is a `dropDown`/`gallery`/`comboBox`
 * entry; `unknown` is the fallback for any element outside this set (never silently dropped). */
export type RibbonControlKind =
  | 'button'
  | 'toggleButton'
  | 'checkBox'
  | 'editBox'
  | 'dropDown'
  | 'comboBox'
  | 'gallery'
  | 'menu'
  | 'dynamicMenu'
  | 'splitButton'
  | 'buttonGroup'
  | 'box'
  | 'labelControl'
  | 'separator'
  | 'menuSeparator'
  | 'dialogBoxLauncher'
  | 'control'
  | 'item'
  | 'unknown';

/** A `<group>` within a ribbon tab: its identity/label attributes and the controls it contains. */
export interface RibbonGroup {
  readonly id?: string;
  readonly idQ?: string;
  readonly idMso?: string;
  readonly label?: string;
  /** Every attribute on the `<group>`, verbatim. */
  readonly attributes: Readonly<Record<string, string>>;
  readonly controls: readonly RibbonControl[];
}

/** A `<tab>` within the ribbon: its identity/label attributes and the groups it contains. */
export interface RibbonTab {
  readonly id?: string;
  readonly idQ?: string;
  readonly idMso?: string;
  readonly label?: string;
  /** Every attribute on the `<tab>`, verbatim. */
  readonly attributes: Readonly<Record<string, string>>;
  readonly groups: readonly RibbonGroup[];
}

/** The parsed `<ribbon>` element: whether it starts from a blank ribbon, and its custom tabs. Only the
 * `<tabs>` subtree is modelled; `qat` and `contextualTabs` are not parsed in v1. */
export interface Ribbon {
  /** `startFromScratch="true"` reduces the built-in ribbon to a minimal set before custom tabs apply. */
  readonly startFromScratch: boolean;
  readonly tabs: readonly RibbonTab[];
}

/**
 * A parsed `customUI` part. `dialect` records which schema it was written against (derived from the
 * root namespace, the authoritative signal). `ribbon` is the parsed `<ribbon>` subtree, or `undefined`
 * when the document customises only backstage/QAT/commands (which v1 does not parse). Future work can
 * extend this with `backstage`/`qat` without changing the shape callers already depend on.
 */
export interface CustomUiDocument {
  readonly dialect: RibbonDialect;
  readonly ribbon?: Ribbon;
}

// A hostile part could nest elements thousands deep; the recursive control mapper would then overflow
// the call stack with an uncatchable RangeError. Capping tree depth keeps every failure a
// CustomUiParseError. Real ribbons nest only a handful deep (tab > group > menu > submenu), so this
// bound is unreachable by any legitimate document.
const MAX_DEPTH = 256;

// Keyed by the union minus `unknown`: that is this reader's word for an element it did not
// recognise, never one a part declares.
/** Narrow a raw element's local name to a {@link RibbonControlKind} this reader models. */
const isKnownControlKind = tokenSet<Exclude<RibbonControlKind, 'unknown'>>({
  button: true,
  toggleButton: true,
  checkBox: true,
  editBox: true,
  dropDown: true,
  comboBox: true,
  gallery: true,
  menu: true,
  dynamicMenu: true,
  splitButton: true,
  buttonGroup: true,
  box: true,
  labelControl: true,
  separator: true,
  menuSeparator: true,
  dialogBoxLauncher: true,
  control: true,
  item: true,
});

// A minimal element node built from the SAX event stream: enough to walk the small customUI tree
// without a general-purpose DOM dependency. `name` keeps the qualified form so namespace resolution can
// tell a default-namespaced `<customUI>` from a prefixed `<mso:customUI>`; `local` is the stripped name
// every structural match uses.
interface RawElement {
  readonly name: string;
  readonly local: string;
  readonly attrs: XmlAttributes;
  readonly children: RawElement[];
}

/**
 * Parse a `customUI` part (raw UTF-8 bytes or its decoded text) into a {@link CustomUiDocument}.
 *
 * @throws {CustomUiParseError} if the XML is malformed, the root is not a `<customUI>` element in
 *   a recognised namespace, or the tree nests beyond {@link MAX_DEPTH}.
 */
export function parseCustomUi(input: string | Uint8Array): CustomUiDocument {
  const xml = typeof input === 'string' ? input : strFromU8(input);

  let root: RawElement;
  try {
    root = buildTree(xml);
  } catch (error) {
    if (error instanceof CustomUiParseError) throw error;
    const detail = error instanceof Error ? error.message : String(error);
    throw new CustomUiParseError(`malformed customUI XML: ${detail}`, {cause: error});
  }

  const customUi = findChild(root, 'customUI');
  if (customUi === undefined) {
    throw new CustomUiParseError('part has no <customUI> root element');
  }

  const dialect = dialectOf(customUi);
  const ribbonEl = findChild(customUi, 'ribbon');
  return ribbonEl === undefined ? {dialect} : {dialect, ribbon: toRibbon(ribbonEl)};
}

// Resolve the namespace bound to the customUI element's own name and map it to a dialect. A default
// namespace (`xmlns=`) governs an unprefixed `<customUI>`; a prefixed `<p:customUI>` is governed by
// that prefix's `xmlns:p`. Neither of the two known namespaces present means this is not a customUI part
// this reader understands, so fail closed rather than guess.
function dialectOf(customUi: RawElement): RibbonDialect {
  const colon = customUi.name.indexOf(':');
  const nsAttr = colon === -1 ? 'xmlns' : `xmlns:${customUi.name.slice(0, colon)}`;
  const ns = customUi.attrs[nsAttr];
  if (ns === CUSTOMUI_2006_NAMESPACE) return '2007';
  if (ns === CUSTOMUI_2009_NAMESPACE) return '2010';
  throw new CustomUiParseError(`unrecognised customUI namespace: ${ns ?? '(none)'}`);
}

function toRibbon(ribbonEl: RawElement): Ribbon {
  const tabsEl = findChild(ribbonEl, 'tabs');
  const tabs = tabsEl === undefined ? [] : childrenNamed(tabsEl, 'tab').map(toTab);
  return {startFromScratch: boolStrict(ribbonEl.attrs.startFromScratch), tabs};
}

function toTab(tabEl: RawElement): RibbonTab {
  return {
    ...identity(tabEl.attrs),
    attributes: tabEl.attrs,
    groups: childrenNamed(tabEl, 'group').map(toGroup),
  };
}

function toGroup(groupEl: RawElement): RibbonGroup {
  return {
    ...identity(groupEl.attrs),
    attributes: groupEl.attrs,
    controls: groupEl.children.map((child) => toControl(child)),
  };
}

function toControl(el: RawElement): RibbonControl {
  const kind: RibbonControlKind = isKnownControlKind(el.local) ? el.local : 'unknown';
  return {
    kind,
    ...identity(el.attrs),
    ...(el.attrs.onAction !== undefined ? {onAction: el.attrs.onAction} : {}),
    attributes: el.attrs,
    ...(el.children.length > 0 ? {children: el.children.map((child) => toControl(child))} : {}),
  };
}

// The id/label attributes lifted onto every tab, group, and control. Emitted only when present, so an
// absent attribute stays absent (exactOptionalPropertyTypes) rather than becoming an explicit undefined.
function identity(attrs: XmlAttributes): {
  id?: string;
  idQ?: string;
  idMso?: string;
  label?: string;
} {
  return {
    ...(attrs.id !== undefined ? {id: attrs.id} : {}),
    ...(attrs.idQ !== undefined ? {idQ: attrs.idQ} : {}),
    ...(attrs.idMso !== undefined ? {idMso: attrs.idMso} : {}),
    ...(attrs.label !== undefined ? {label: attrs.label} : {}),
  };
}

// Build the element tree from the SAX event stream with an explicit stack (no recursion), capping depth
// so a hostile part cannot force the later recursive walk to overflow. Text and comments carry no ribbon
// meaning and are ignored. Throws CustomUiParseError on unbalanced markup; an XmlParseError from the
// scanner is caught and re-wrapped by the caller.
function buildTree(xml: string): RawElement {
  const root: RawElement = {name: '#root', local: '#root', attrs: {}, children: []};
  const stack: RawElement[] = [root];

  for (const event of xmlEvents(xml)) {
    if (event.kind === 'open') {
      const parent = stack[stack.length - 1];
      if (parent === undefined) throw new CustomUiParseError('unbalanced customUI markup');
      const el: RawElement = {
        name: event.name,
        local: localName(event.name),
        attrs: event.attrs,
        children: [],
      };
      parent.children.push(el);
      if (!event.selfClosing) {
        if (stack.length > MAX_DEPTH) throw new CustomUiParseError('customUI nesting too deep');
        stack.push(el);
      }
    } else if (event.kind === 'close') {
      if (stack.length <= 1) throw new CustomUiParseError('unbalanced customUI markup');
      stack.pop();
    }
  }

  if (stack.length !== 1) throw new CustomUiParseError('unbalanced customUI markup');
  return root;
}

function findChild(el: RawElement, local: string): RawElement | undefined {
  return el.children.find((child) => child.local === local);
}

function childrenNamed(el: RawElement, local: string): RawElement[] {
  return el.children.filter((child) => child.local === local);
}
