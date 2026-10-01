// Workbook-shape helpers used by App.jsx. These only rearrange data the engine (csvx-ts) already
// produced or will accept back — none of them decide what a value, type, or style *means*. That
// distinction is the whole point of this rewrite: see AGENTS.md and ../csvx-spec/AGENTS.md rule 5.
//
// Specifically NOT here (on purpose, because no engine implements it yet): cell-type inference
// from raw text, and computed number formatting (parsing "$#,##0.00" and producing "$1,234.56").
// Row/column insert and delete (below) DO rewrite formula references on shift — that's lexical
// reference-adjustment (finding "B2"-shaped tokens and renumbering them), not formula evaluation,
// so it doesn't cross the line rule 5.1 draws; nothing here parses a formula's actual semantics or
// computes a result.
//
// Also notably absent: anything that writes to Column.width. schemas/sheet-metadata.schema.json
// never defines a unit for it, and real imported data (example.csvx, via csvx-go's XLSX importer)
// stores it in XLSX character-width units (e.g. 26.25), not CSS pixels — confirmed by actually
// running this app against that fixture, which is exactly the kind of gap "building it and using
// it for real" surfaces that spec-reading doesn't. App.jsx's pixelWidthForColumn() reads this
// field and approximates it to pixels purely for initial display; a user's manual resize is still
// view-only state, local to the browser session, and never writes back into this field — writing
// a pixel number into it would silently corrupt the original unit on export.

import { columnId } from 'csvx-ts/browser'

export function coordinateFor(columnIndex, rowIndex) {
  return `${columnId(columnIndex)}${rowIndex + 1}`
}

/** Parses "AB12" into zero-based {column, row} indices — the inverse of coordinateFor. */
export function indicesForCoordinate(coordinate) {
  const match = /^([A-Z]+)(\d+)$/.exec(coordinate)
  if (!match) return { column: 0, row: 0 }
  let column = 0
  for (const char of match[1]) column = column * 26 + (char.charCodeAt(0) - 64)
  return { column: column - 1, row: Number(match[2]) - 1 }
}

function columnIndexFromLabel(label) {
  let column = 0
  for (const char of label) column = column * 26 + (char.charCodeAt(0) - 64)
  return column - 1
}

export function cellMetadata(sheet, coordinate) {
  return sheet?.cells?.[coordinate]
}

export function styleForId(styles, id) {
  return (styles || []).find((style) => style.id === id) || {}
}

/** Literal field-to-CSS mapping, not interpretation. No number-format math, no type-based
 * heuristics — just rendering declared font/fill/border/alignment as CSS. */
export function cellCSS(style) {
  const font = style.font || {}
  const fill = style.fill || {}
  const alignment = style.alignment || {}
  const border = style.border || {}
  return {
    color: font.color || undefined,
    backgroundColor: fill.color || undefined,
    fontWeight: font.bold ? 700 : undefined,
    fontStyle: font.italic ? 'italic' : undefined,
    textDecoration: font.underline ? 'underline' : undefined,
    textAlign: alignment.horizontal || undefined,
    whiteSpace: alignment.wrapText ? 'normal' : 'nowrap',
    boxShadow: border.color ? `inset 0 0 0 0.0625rem ${border.color}` : undefined,
  }
}

export function findSheet(workbook, sheetId) {
  return workbook.sheets.find((sheet) => sheet.id === sheetId) || workbook.sheets[0]
}

/** Sets a cell's raw CSV value. Clears any existing `formula`/`cached` metadata on that cell,
 * since a literal value the user just typed and a stale formula result can't both be true at
 * once — this is invariant maintenance, not inventing what the new value means. Style and
 * validation metadata, which describe the cell rather than its content, are preserved. */
export function setCellValue(workbook, sheetId, rowIndex, columnIndex, value) {
  return {
    ...workbook,
    sheets: workbook.sheets.map((sheet) => {
      if (sheet.id !== sheetId) return sheet
      const records = sheet.records.map((row, index) => (index === rowIndex ? [...row] : row))
      records[rowIndex][columnIndex] = value
      const coordinate = coordinateFor(columnIndex, rowIndex)
      const cells = { ...sheet.cells }
      const existing = cells[coordinate]
      if (existing) {
        const { formula, cached, ...rest } = existing
        if (Object.keys(rest).length > 0) cells[coordinate] = rest
        else delete cells[coordinate]
      }
      return { ...sheet, records, cells }
    }),
  }
}

/** Appends a blank row at the end of a sheet. Safe without formula-reference rewriting because
 * nothing above it shifts. */
export function appendRow(workbook, sheetId) {
  return {
    ...workbook,
    sheets: workbook.sheets.map((sheet) =>
      sheet.id !== sheetId ? sheet : { ...sheet, records: [...sheet.records, Array(sheet.columns.length).fill('')] },
    ),
  }
}

/** Appends a blank column at the end of a sheet. Same reasoning as appendRow.
 *
 * The new column's name is its own letter, not an invented label like "Column 8" — column.name is
 * written verbatim as the CSV header cell on export (see writeWorkbookToZip) and csvx-ts's CSV
 * parser rejects an empty header cell outright, so it can't just be left blank either. Naming it
 * after its own id, combined with syncSyntheticColumnNames (below) keeping that name in step with
 * the id on every later shift, is what makes a freshly added column behave like Excel/Sheets/
 * Numbers: its header is just its (correctly shifted) letter, for as long as nobody renames it. */
export function appendColumn(workbook, sheetId) {
  return {
    ...workbook,
    sheets: workbook.sheets.map((sheet) => {
      if (sheet.id !== sheetId) return sheet
      const id = columnId(sheet.columns.length)
      const columns = [...sheet.columns, { id, name: id }]
      const records = sheet.records.map((row) => [...row, ''])
      return { ...sheet, columns, records }
    }),
  }
}

/** Keeps a column's name in sync with its id when a shift (insert/delete) changes that id — but
 * only while name still equals the *old* id, i.e. nobody has renamed it away from its letter yet.
 * The moment a column gets a real header name, it naturally stops matching its id and this stops
 * touching it. Without this, a freshly inserted column's name (itself just its letter — see
 * appendColumn) goes stale the next time something shifts it, reproducing the exact "header text
 * doesn't track position" bug this exists to prevent. */
function syncSyntheticColumnNames(columns) {
  return columns.map((column, index) => {
    const id = columnId(index)
    if (column.id === id) return column
    return { ...column, id, name: column.name === column.id ? id : column.name }
  })
}

// styles.schema.json requires an id matching ^[A-Za-z][A-Za-z0-9_-]{0,63}$ — a bare number (what
// this returned before) is schema-invalid. Caught by actually exporting and running the real
// fixture through the canonical validator, not by reasoning about the schema in the abstract.
// "s<N>" matches the convention csvx-go's XLSX importer already uses for real style ids.
function nextStyleIdCounter(styles) {
  const numericSuffixes = (styles || [])
    .map((style) => /^s(\d+)$/.exec(style.id)?.[1])
    .filter((value) => value !== undefined)
    .map(Number)
  return numericSuffixes.length ? Math.max(...numericSuffixes) + 1 : 0
}

/** Applies a literal font/fill/border/alignment patch to every given cell, each getting its own
 * new style record. This does not dedupe identical styles across cells — a real editor would want
 * to, but that's a quality improvement, not a correctness one; tracked as a known simplification
 * rather than solved here.
 *
 * The id counter is computed once up front and incremented locally rather than rescanning `styles`
 * per cell (as a naive `nextStyleId(styles)` call per iteration would) — for a large selection
 * (e.g. a whole column), rescanning plus `[...styles, style]` per cell made this O(n²). */
export function applyCellsFormat(workbook, sheetId, coordinates, patch) {
  const styles = workbook.styles ? [...workbook.styles] : []
  const sheet = findSheet(workbook, sheetId)
  const cells = { ...sheet.cells }
  let nextId = nextStyleIdCounter(styles)
  coordinates.forEach((coordinate) => {
    const existingMeta = cells[coordinate] || {}
    const existingStyle = styleForId(styles, existingMeta.style)
    const style = {
      id: `s${nextId}`,
      ...(existingStyle.numberFormat ? { numberFormat: existingStyle.numberFormat } : {}),
      font: { ...existingStyle.font, ...patch.font },
      fill: { ...existingStyle.fill, ...patch.fill },
      border: { ...existingStyle.border, ...patch.border },
      alignment: { ...existingStyle.alignment, ...patch.alignment },
    }
    nextId += 1
    styles.push(style)
    cells[coordinate] = { ...existingMeta, style: style.id }
  })
  return { ...workbook, styles, sheets: workbook.sheets.map((item) => (item.id === sheetId ? { ...item, cells } : item)) }
}

export function clearCellsFormat(workbook, sheetId, coordinates) {
  return {
    ...workbook,
    sheets: workbook.sheets.map((sheet) => {
      if (sheet.id !== sheetId) return sheet
      const cells = { ...sheet.cells }
      coordinates.forEach((coordinate) => {
        const { style, ...rest } = cells[coordinate] || {}
        if (Object.keys(rest).length > 0) cells[coordinate] = rest
        else delete cells[coordinate]
      })
      return { ...sheet, cells }
    }),
  }
}

export function addSheet(workbook) {
  const usedIds = new Set(workbook.sheets.map((sheet) => sheet.id))
  let index = workbook.sheets.length + 1
  while (usedIds.has(`sheet-${index}`)) index += 1
  const id = `sheet-${index}`
  const sheet = { id, name: `Sheet ${index}`, path: `sheets/${id}.csv`, columns: [{ id: 'A', name: 'Value' }], records: [['']], cells: {} }
  return { ...workbook, sheets: [...workbook.sheets, sheet] }
}

export function renameSheet(workbook, sheetId, name) {
  const trimmed = name.trim()
  if (!trimmed) return workbook
  return { ...workbook, sheets: workbook.sheets.map((sheet) => (sheet.id === sheetId ? { ...sheet, name: trimmed } : sheet)) }
}

export function deleteSheet(workbook, sheetId) {
  return { ...workbook, sheets: workbook.sheets.filter((sheet) => sheet.id !== sheetId) }
}

// Matches a (possibly sheet-qualified, possibly $-anchored) cell reference: [Sheet!]$A$1. Formulas
// aren't parsed into an AST anywhere in this stack (no engine implements that yet — see the file
// header comment), but shifting references on row/column insert or delete only requires finding
// these tokens and renumbering them, which is plain regex work, not formula evaluation.
const CELL_REFERENCE_PATTERN = /((?:'[^']+'|[A-Za-z0-9_]+)!)?(\$?)([A-Z]+)(\$?)(\d+)/g

function referencedSheetName(sheetPrefix, targetSheetName) {
  if (!sheetPrefix) return targetSheetName
  return sheetPrefix.slice(0, -1).replace(/^'|'$/g, '').replace(/''/g, "'")
}

function rewriteFormulaRows(formula, targetSheetName, insertionRow, delta, deleting) {
  if (!formula) return formula
  return formula.replace(CELL_REFERENCE_PATTERN, (match, sheetPrefix, colAnchor, column, rowAnchor, rowText) => {
    if (referencedSheetName(sheetPrefix, targetSheetName) !== targetSheetName) return match
    const row = Number(rowText) - 1
    if (deleting && row === insertionRow) return '#REF!'
    if (row < insertionRow) return match
    return `${sheetPrefix || ''}${colAnchor}${column}${rowAnchor}${row + delta + 1}`
  })
}

function rewriteFormulaColumns(formula, targetSheetName, insertionColumn, delta, deleting) {
  if (!formula) return formula
  return formula.replace(CELL_REFERENCE_PATTERN, (match, sheetPrefix, colAnchor, column, rowAnchor, rowText) => {
    if (referencedSheetName(sheetPrefix, targetSheetName) !== targetSheetName) return match
    const columnIndex = columnIndexFromLabel(column)
    if (deleting && columnIndex === insertionColumn) return '#REF!'
    if (columnIndex < insertionColumn) return match
    return `${sheetPrefix || ''}${colAnchor}${columnId(columnIndex + delta)}${rowAnchor}${rowText}`
  })
}

/** Shared by insertRow/deleteRow: rewrites every formula in the workbook (any sheet can reference
 * the target sheet by name) and, for the target sheet only, both shifts cell metadata to follow
 * its row and splices `records`. A formula whose own row moves still gets its references rewritten
 * — unlike a hand-rolled version of this that existed before the csvx-ts rewrite, which skipped
 * rewriting a formula sitting on a row that itself shifted, silently leaving stale references. */
function shiftSheetRows(workbook, sheetId, insertionRow, delta, deleting) {
  const targetSheet = findSheet(workbook, sheetId)
  return {
    ...workbook,
    sheets: workbook.sheets.map((sheetItem) => {
      const nextCells = {}
      Object.entries(sheetItem.cells || {}).forEach(([coordinate, metadata]) => {
        const rewritten = metadata.formula
          ? { ...metadata, formula: rewriteFormulaRows(metadata.formula, targetSheet.name, insertionRow, delta, deleting) }
          : metadata
        if (sheetItem.id !== sheetId) {
          nextCells[coordinate] = rewritten
          return
        }
        const { row, column } = indicesForCoordinate(coordinate)
        if (row < insertionRow) {
          nextCells[coordinate] = rewritten
          return
        }
        if (deleting && row === insertionRow) return
        nextCells[coordinateFor(column, row + delta)] = rewritten
      })
      if (sheetItem.id !== sheetId) return { ...sheetItem, cells: nextCells }
      const records = [...sheetItem.records]
      if (deleting) records.splice(insertionRow, 1)
      else records.splice(insertionRow, 0, Array(sheetItem.columns.length).fill(''))
      return { ...sheetItem, records, cells: nextCells }
    }),
  }
}

/** Analogous to shiftSheetRows, for columns: also keeps each remaining Column.id in sync with its
 * new index, matching the convention columnId(index) already establishes elsewhere (buildSheetFromCSV
 * in csvx-ts, appendColumn above) — without this, a column's id would silently point at the wrong
 * letter after a mid-sheet insert/delete. */
function shiftSheetColumns(workbook, sheetId, insertionColumn, delta, deleting) {
  const targetSheet = findSheet(workbook, sheetId)
  return {
    ...workbook,
    sheets: workbook.sheets.map((sheetItem) => {
      const nextCells = {}
      Object.entries(sheetItem.cells || {}).forEach(([coordinate, metadata]) => {
        const rewritten = metadata.formula
          ? { ...metadata, formula: rewriteFormulaColumns(metadata.formula, targetSheet.name, insertionColumn, delta, deleting) }
          : metadata
        if (sheetItem.id !== sheetId) {
          nextCells[coordinate] = rewritten
          return
        }
        const { row, column } = indicesForCoordinate(coordinate)
        if (column < insertionColumn) {
          nextCells[coordinate] = rewritten
          return
        }
        if (deleting && column === insertionColumn) return
        nextCells[coordinateFor(column + delta, row)] = rewritten
      })
      if (sheetItem.id !== sheetId) return { ...sheetItem, cells: nextCells }
      let columns = sheetItem.columns
      let records = sheetItem.records
      if (deleting) {
        columns = columns.filter((_, index) => index !== insertionColumn)
        records = records.map((row) => row.filter((_, index) => index !== insertionColumn))
      } else {
        columns = [...columns]
        const id = columnId(insertionColumn)
        columns.splice(insertionColumn, 0, { id, name: id })
        records = records.map((row) => {
          const next = [...row]
          next.splice(insertionColumn, 0, '')
          return next
        })
      }
      columns = syncSyntheticColumnNames(columns)
      return { ...sheetItem, columns, records, cells: nextCells }
    }),
  }
}

export function insertRow(workbook, sheetId, rowIndex) {
  return shiftSheetRows(workbook, sheetId, rowIndex, 1, false)
}

export function deleteRow(workbook, sheetId, rowIndex) {
  return shiftSheetRows(workbook, sheetId, rowIndex, -1, true)
}

export function insertColumn(workbook, sheetId, columnIndex) {
  return shiftSheetColumns(workbook, sheetId, columnIndex, 1, false)
}

export function deleteColumn(workbook, sheetId, columnIndex) {
  return shiftSheetColumns(workbook, sheetId, columnIndex, -1, true)
}
