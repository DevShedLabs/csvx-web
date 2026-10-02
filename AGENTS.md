# Architecture rules for this repo

The binding rules for this project live in `../csvx-spec/AGENTS.md`. Read it before making any
change here. The short version for this repo specifically:

- This app is a **consumer of a real CSVX engine** — `csvx-ts`, imported via its `/browser`
  subpath (`import ... from 'csvx-ts/browser'`) so filesystem-only code (`node:fs`-based wrappers)
  never gets bundled. It is UI/interaction code, nothing else.
- **No CSVX parsing, serialization, cell-type inference, formula evaluation, number formatting, or
  style application logic belongs in this repo.** This app used to contain a hand-rolled ZIP
  reader, CSV parser, `inferCellType`, `formatCellValue`, and `cellStyle` — all engine logic
  reimplemented here instead of calling a real engine. That was replaced (2026-10-01) with calls
  into `csvx-ts`: `loadWorkbookFromZip`, `writeWorkbookToZip`, `validateBuffer`. Don't add logic
  like the old implementation back in, even incrementally.
- Two kinds of logic that look like candidates for "quick fixes" here are deliberately *not*
  implemented, because no engine implements them yet: cell-type inference from raw text, and
  computed number formatting (turning `numberFormat: "$#,##0.00"` into a displayed `"$1,234.56"`).
  If a feature needs either, that's an engine task (`csvx-ts`), not something to patch in here —
  see `csvx-spec/AGENTS.md` rule 3.5 ("spec first, schema second, tests third, implementation
  last").
- Row/column **insert-in-the-middle and delete are not implemented**, on purpose: both require
  rewriting formula references that shift, and no engine does formula parsing yet. Append-only
  add (`+ Row`, the column `+` header button) is safe without that and is implemented. Don't add
  insert/delete without first confirming an engine can keep formula references correct across it.
- Column width (`Column.width` in `sheet-metadata.schema.json`) has no defined unit in the spec,
  and real imported data stores it in XLSX character-width units, not CSS pixels (confirmed by
  loading `example.csvx` and seeing columns render far too narrow when treated as pixels). Column
  resize in this app is therefore view-only React state, never written into the workbook — see the
  comment on it in `src/model.js`. Don't start writing a pixel number into that field without
  resolving the unit question in `csvx-spec` first.
- Style formatting (bold/italic/underline/colors/alignment) *is* implemented, because it's a
  literal field-to-CSS mapping of data the engine already gives you — not an interpretation. The
  line between "render declared data" (fine) and "decide what the data means" (not fine) is the
  one that matters throughout this repo.
- UI-only state (selection, scroll position, active editor, column-width overrides, toolbar
  layout) is fine to keep here. A second opinion about what a value, type, or style *means* is not.

See `../csvx-spec/AGENTS.md` for full detail and the reasoning behind these rules.
- Print is implemented as a Google-Sheets-style Print view (`src/PrintView.jsx`, `src/print.js`).
  Pagination is layout — `spec/03-sheets.md` ("Print settings") defines the settings and says
  pagination is the renderer's job — so it lives here, free of React so it can move into an engine.
  Every print setting is read from and written to the sheet's `print` object so it round-trips; never
  keep a print setting as UI-only state. The used-range rule (`usedRange` in `src/model.js`) is also
  from the spec.
- Row 1 is the CSV header, everywhere (`../csvx-spec/spec/03-sheets.md`). The convention — header is
  `HEADER_ROW` (-1), record i is row i + 2 — is owned by `csvx-ts` (`coordinateFor`,
  `indicesForCoordinate`, `rowNumberFor`, `rowIndexFor`, `rawCellText`), as is workbook recalculation
  (`recalculateWorkbook`). Never write a literal `+ 1`/`+ 2` row offset here; call those helpers.
  The grid renders the header as its first row (editing it renames the column), and
  `public/example.csvx` must be regenerated with `csvx import` whenever the importer's mapping changes.
