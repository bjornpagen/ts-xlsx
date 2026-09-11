// Reading `xl/workbook.xml` and the two `docProps` parts: the sheet declarations, the workbook's
// protection, its window view, its defined names, and the document properties.
//
// Beside its inverse. `workbook-xml.ts` writes exactly these elements, and this tree's stated
// preference is to keep both directions of one wire form in one place, where the two can be read
// against each other: `color-xml.ts`, `theme-xml.ts`, `tables.ts`, `rels.ts` and `font-xml.ts` all do.
// This part's reader had been living inside the orchestrator that happened to call it first, and the
// worksheet's layout blocks were doing the same until `sheet-properties.ts` took its half back.
//
// Every reader here is a {@link SaxPass} rather than a function taking the part's text, because
// `read.ts` drives all of them from one parse. Scanning the workbook part once per reader is what the
// worksheet part was already fixed for, and this part had got none of the treatment: six scans, of
// which four matched no element each.

import {parseDateText} from '../../core/date.ts';
import {definedNameKeys, unmangleFunctions} from '../../core/formula.ts';
import {
  WORKBOOK_PROTECTION_CREDENTIAL_ATTRS,
  type WorkbookProtection,
  type WorkbookProtectionCredentialAttr,
} from '../../core/workbook-protection.ts';
import type {DefinedName, Workbook, WorkbookView} from '../../core/workbook.ts';
import {isVisibility, type WorksheetState} from '../../core/worksheet.ts';
import {enumToken, numInteger} from '../../xml/xml-attrs.ts';
import {capturedText, type CollectingPass, type SaxPass, TextCapture} from '../../xml/xml-read.ts';
import {boolStrict, localName, type XmlAttributes} from '../../xml/xml-scan.ts';
import {relAttr} from '../opc/namespaces.ts';

// One `<sheet>` entry from `xl/workbook.xml`: its display name, the rel id linking to the sheet part,
// and its visibility state (absent for a normal, visible sheet).
export interface SheetEntry {
  readonly name: string;
  readonly relId: string;
  readonly state?: WorksheetState['state'];
}

export function workbookSheetsPass(): CollectingPass<SheetEntry[]> {
  const sheets: SheetEntry[] = [];
  return {
    handlers: {
      onOpen(name, attrs, _selfClosing, scope) {
        if (localName(name) !== 'sheet') return;
        const entry: {name: string; relId: string; state?: WorksheetState['state']} = {
          name: attrs.name ?? '',
          relId: relAttr(scope, attrs, 'id') ?? '',
        };
        // `visible` is the schema default and the model's, so it is dropped rather than stored:
        // keeping it would put a `state="visible"` attribute into a file Excel writes without one.
        const state = enumToken(attrs.state, isVisibility);
        if (state !== undefined && state !== 'visible') entry.state = state;
        sheets.push(entry);
      },
    },
    result: () => sheets,
  };
}

// Read the workbook's structure/window protection (`<workbookProtection>`). The three lock flags are
// decoded as booleans (an absent or "0" attribute stays unlocked), and only the whitelisted
// password/agile-hash attributes are preserved verbatim: a hostile or unknown attribute is dropped
// rather than echoed back on write. Returns undefined when the workbook declares no protection.
export function workbookProtectionPass(): CollectingPass<WorkbookProtection | undefined> {
  let found: WorkbookProtection | undefined;
  const handlers = {
    onOpen(name: string, attrs: XmlAttributes): void {
      if (localName(name) !== 'workbookProtection') return;
      const protection: {
        lockStructure?: boolean;
        lockWindows?: boolean;
        lockRevision?: boolean;
        credentials?: Partial<Record<WorkbookProtectionCredentialAttr, string>>;
      } = {};
      if (boolStrict(attrs.lockStructure)) protection.lockStructure = true;
      if (boolStrict(attrs.lockWindows)) protection.lockWindows = true;
      if (boolStrict(attrs.lockRevision)) protection.lockRevision = true;
      const credentials: Partial<Record<WorkbookProtectionCredentialAttr, string>> = {};
      for (const key of WORKBOOK_PROTECTION_CREDENTIAL_ATTRS) {
        const value = attrs[key];
        if (value !== undefined) credentials[key] = value;
      }
      if (Object.keys(credentials).length > 0) protection.credentials = credentials;
      found = protection;
    },
  };
  return {handlers, result: () => found};
}

/**
 * Apply `<workbookPr>`: the workbook's date system and its VBA code name.
 *
 * Both are read for the same reason and it is not that the model computes anything from them. The
 * date system is what every serial in every sheet counts from, so a workbook read without it is read
 * with every date four years and a day out; the code name is what the VBA project means by
 * `ThisWorkbook`, so a `.xlsm` written back without it has had its macros unbound from its document.
 * Neither has any signal in the file other than this element.
 */
export function workbookPropertiesPass(workbook: Workbook): SaxPass {
  let seen = false;
  return {
    handlers: {
      onOpen(name, attrs) {
        if (seen || localName(name) !== 'workbookPr') return;
        seen = true;
        if (boolStrict(attrs.date1904)) workbook.dateEpoch = 1904;
        if (attrs.codeName !== undefined) workbook.codeName = attrs.codeName;
      },
    },
  };
}

// Restore the workbook's saved window state from `<bookViews><workbookView/>` onto the model's view,
// so a round-trip hands back the geometry and active tab the author left rather than stamping the
// library's defaults over them. Only the first `<workbookView>` is read: the model carries one view,
// which is all Excel writes and all a single consuming window can restore.
//
// Each attribute is applied only when the source carried a usable value; an absent or non-numeric one
// leaves the default in place, so a truncated or hostile element degrades to a valid window rather
// than a NaN geometry that would serialise as garbage.
export function workbookViewPass(view: WorkbookView): SaxPass {
  let seen = false;
  return {
    handlers: {
      onOpen(name, attrs) {
        if (seen || localName(name) !== 'workbookView') return;
        seen = true;
        // The window may sit at a negative origin (a secondary monitor left of the primary), so only
        // the extents and the tab ordinal carry a floor.
        const x = numInteger(attrs.xWindow);
        if (x !== undefined) view.x = x;
        const y = numInteger(attrs.yWindow);
        if (y !== undefined) view.y = y;
        const width = numInteger(attrs.windowWidth, 0);
        if (width !== undefined) view.width = width;
        const height = numInteger(attrs.windowHeight, 0);
        if (height !== undefined) view.height = height;
        const activeTab = numInteger(attrs.activeTab, 0);
        if (activeTab !== undefined) view.activeTab = activeTab;
        const visibility = enumToken(attrs.visibility, isVisibility);
        if (visibility !== undefined && visibility !== 'visible') view.visibility = visibility;
        if (boolStrict(attrs.minimized)) view.minimized = true;
      },
    },
  };
}

// Reconstruct the workbook's defined names. Each `<definedName>` carries its name (and optional
// comment/hidden flag) as attributes and its refersTo formula as text content; a `localSheetId`
// maps back through the sheet order to the scope sheet's name. A name whose localSheetId is out of
// range (a foreign file referencing a sheet we did not load) is left global rather than dropped.
export interface DefinedNamesPass extends SaxPass {
  /**
   * The names, scoped against the sheet order.
   *
   * The order is supplied here rather than to the pass, and that is what lets the names ride the same
   * scan as everything else this part carries. A scoped name's `localSheetId` indexes the sheets in
   * declaration order, and the model's answer for index n is known only after the sheet loop has run
   * and repaired the names it had to; resolving during the scan would have forced a second one.
   */
  result(sheetOrder: readonly string[]): DefinedName[];

  /**
   * Every name the part defines, whatever its scope, as `definedNameKeys` spells them. Known once the
   * scan ends, before any sheet is read, and needed then: whether a function a formula passes as a value
   * sheds its `_xleta.` depends on them.
   */
  spellings(): ReadonlySet<string>;
}

export function definedNamesPass(): DefinedNamesPass {
  // `localSheetId` is kept as read and resolved in `result`; everything else is final on close.
  const drafts: {
    name: string;
    localSheetId: number;
    comment?: string;
    hidden?: boolean;
    refersTo: string;
  }[] = [];
  const refersTo = new TextCapture('definedName');
  let pending: {name: string; localSheetId: number; comment?: string; hidden?: boolean} | undefined;
  return {
    handlers: {
      onOpen(name, attrs, selfClosing) {
        if (localName(name) !== 'definedName' || attrs.name === undefined) return;
        // `_xlnm._FilterDatabase` is the built-in Excel derives from a sheet's autofilter, not a
        // user-defined name: it is reconstructed from the sheet's `<autoFilter>` element, so skip it
        // here to keep it off `Workbook.definedNames` and out of a duplicating round-trip.
        if (attrs.name === '_xlnm._FilterDatabase') return;
        refersTo.open(localName(name), selfClosing);
        pending = {name: attrs.name, localSheetId: numInteger(attrs.localSheetId, 0) ?? -1};
        if (attrs.comment !== undefined) pending.comment = attrs.comment;
        if (boolStrict(attrs.hidden)) pending.hidden = true;
      },
      onText(chunk) {
        refersTo.text(chunk);
      },
      onClose(name) {
        const text = refersTo.close(localName(name));
        if (text === undefined || pending === undefined) return;
        // Kept as stored until `result`: a name's formula can pass a function as a value, and whether
        // that sheds its prefix depends on names the part may declare after this one.
        drafts.push({...pending, refersTo: text});
        pending = undefined;
      },
    },
    // A `<definedName name="X"/>` with no formula is legal, fires no close, and used to be dropped
    // entirely while leaving the hand-rolled capture latched on whatever text came next. Expanded
    // into an open plus a close, it commits through the same path every other name does.
    closeEmptyElements: new Set(['definedName']),
    result(sheetOrder) {
      const names = definedNameKeys(drafts);
      return drafts.map(({localSheetId, refersTo: stored, ...draft}) => {
        // Stripped back to the readable form, the same normalisation a cell formula gets, so the model
        // never holds the on-disk mangling.
        const named = {...draft, refersTo: unmangleFunctions(stored, names)};
        // A name whose localSheetId is out of range (a foreign file referencing a sheet we did not
        // load) is left global rather than dropped.
        const scope = sheetOrder[localSheetId];
        return scope === undefined ? named : {...named, scope};
      });
    },
    spellings() {
      return definedNameKeys(drafts);
    },
  };
}

// Core document properties live in docProps/core.xml under mixed namespaces
// (dc:creator, cp:lastModifiedBy, dcterms:created/modified); local names disambiguate.
const CORE_PROPERTY_LOCAL_NAMES = new Set([
  'title',
  'creator',
  'lastModifiedBy',
  'created',
  'modified',
]);

export function applyCoreProperties(workbook: Workbook, xml: string): void {
  for (const {local, text} of capturedText(xml, CORE_PROPERTY_LOCAL_NAMES)) {
    if (local === 'title') workbook.properties.title = text;
    else if (local === 'creator') workbook.properties.creator = text;
    else if (local === 'lastModifiedBy') workbook.properties.lastModifiedBy = text;
    else {
      const date = parseDateText(text);
      if (date !== null) {
        if (local === 'created') workbook.properties.created = date;
        else workbook.properties.modified = date;
      }
    }
  }
}

// `Company` is the one document property OOXML keeps in the extended part rather than the core
// one. Everything else in app.xml is either derived (`TitlesOfParts`) or this library's own
// (`Application`), so nothing here reads more than the single element.
export function applyAppProperties(workbook: Workbook, xml: string): void {
  for (const {text} of capturedText(xml, 'Company')) workbook.properties.company = text;
}
