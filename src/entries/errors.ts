// Entry barrel for `@shbernal/ts-xlsx/errors`: the whole failure taxonomy, and nothing else.
//
// Every error class the library throws is exported here and from no other entry. That rule is
// what keeps the entry barrels disjoint (a container-level failure belongs to no single codec:
// `readXlsx` and `readXlsb` both raise `UnsupportedFormatError`), and it makes the answer to
// "what can this throw at me?" one import rather than a hunt across codecs.
//
// It is also the cheapest thing in the package by an order of magnitude: the classes reach
// nothing but each other, so a service that only needs to classify a failure (log it, map it to
// an HTTP status, decide whether to retry) pays for the taxonomy and not for a parser.
//
// `XlsxError` is the one-line answer to "was that us?"; `XlsxErrorCode` is the kind of failure,
// shared across classes on purpose (see `src/errors.ts`).

export {CustomUiParseError} from '../customui/errors.ts';
export {AuthoringError, InternalError, XlsxError, type XlsxErrorCode} from '../errors.ts';
export {CsvParseError} from '../io/csv/errors.ts';
export {
  PackageReadError,
  type UnsupportedFormat,
  UnsupportedFormatError,
} from '../io/opc/errors.ts';
export {XlsbParseError} from '../io/xlsb/errors.ts';
export {XlsxParseError} from '../io/xlsx/errors.ts';
export {VbaAuthorError, VbaParseError} from '../vba/errors.ts';
export {XmlParseError} from '../xml/errors.ts';
