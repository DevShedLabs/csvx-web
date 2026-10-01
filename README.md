# CSVX Web

CSVX Web is the React/Vite demo workbench for opening, inspecting, editing, and exporting CSVX
files. It is a consumer of a real engine (`csvx-ts`), not a second definition of the format — see
`AGENTS.md` and `handoff.md` for the rules and the 2026-10-01 rewrite that made that true.

## Current scope

- Real `.csvx` open and export, through `csvx-ts/browser` (`loadWorkbookFromZip`,
  `writeWorkbookToZip`, `validateBuffer`) — not a hand-rolled parser
- Workbook and sheet navigation, with add/rename/delete
- CSV-backed spreadsheet data display, with cell selection by A1 coordinate
- Cell editing (Enter, Space, F2 to start; Enter to commit; Escape to cancel)
- Append-only row/column addition (`+ Row`, the column header `+` button)
- A formatting toolbar (bold/italic/underline/text/fill/border color/alignment) that writes literal
  style fields the engine already understands
- Formula and cached-value transparency (shown as-is, never recomputed)
- Validation diagnostics from the engine on load and before export, not invented in React
- Responsive layout and keyboard-visible focus states

Not implemented, because no engine implements the underlying behavior yet: cell-type inference,
computed number formatting, and row/column insert-in-the-middle or delete (these need
formula-reference rewriting). See `AGENTS.md` for why these are engine tasks, not something to add
here, and `handoff.md` for what this cost in practice.

`public/example.csvx` is loaded on startup as the demo workbook — a real fixture, not fabricated
data.

## Commands

```bash
npm install
npm run dev       # dev server on port 5174
npm run build
npm run preview
```

## Regenerating the example fixture

```bash
csvx convert example.xlsx public/example.csvx
```

## Direction

Keep the UI separate from CSVX semantics: this app renders and edits data the engine already
understands, and never holds a second opinion about what a value, type, or style means. See
`handoff.md`'s "Immediate next steps" for what's left.
