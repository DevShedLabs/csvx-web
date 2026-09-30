# Architecture rules for this repo

The binding rules for this project live in `../csvx-spec/AGENTS.md`. Read it before making any
change here. The short version for this repo specifically:

- This app is a **consumer of a real CSVX engine** (`csvx-go`, `csvx-ts`, or whichever is wired
  up) — it is UI/interaction code, nothing else.
- **No CSVX parsing, serialization, cell-type inference, formula evaluation, number formatting, or
  style application logic belongs in this repo.** `src/App.jsx` currently contains a hand-rolled
  ZIP reader, CSV parser, `inferCellType`, `formatCellValue`, and `cellStyle` — all of that is
  engine logic that was reimplemented here instead of calling a real engine. It made every past
  "fix" in this app a fix to a fake backend, not to CSVX support. This is being corrected; don't
  add more logic like it in the meantime.
- Until a real engine is wired in (WASM, local API, or bindings — a decision made explicitly, not
  defaulted into), do not treat this app's current read/write behavior as a reference for what
  CSVX does. It isn't one.
- UI-only state (selection, scroll position, active editor, toolbar layout) is fine to keep here.
  A second opinion about what a value, type, or style *means* is not.

See `../csvx-spec/AGENTS.md` for full detail and the reasoning behind these rules.
