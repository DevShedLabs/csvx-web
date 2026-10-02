// Workbook-shape helpers used by App.jsx. These only rearrange data the engine (csvx-ts) already
// produced or will accept back — none of them decide what a value, type, or style *means*. That
// distinction is the whole point of this rewrite: see AGENTS.md and ../csvx-spec/AGENTS.md rule 5.
//
// Specifically NOT here: cell-type inference from raw text (csvx-ts's resolveCellValue), computed
// number formatting (csvx-ts's formatValue), formula evaluation (csvx-ts's recalculateCells), and
// deciding which cell metadata survives an edit (csvx-ts's nextCellMetadata) — all genuine CSVX
// semantics, so all of it is a call into the engine, never a second opinion reimplemented here.
// Row/column insert and delete (below) DO rewrite formula references on shift — that's lexical
// reference-adjustment (finding "B2"-shaped tokens and renumbering them), not formula evaluation,
// so it doesn't cross the line rule 5.1 draws; nothing here parses a formula's actual semantics or
// computes a result.
//
// Column.width is in XLSX character-width units (e.g. 26.25), per spec/03-sheets.md — this was
// discovered by actually running this app against example.csvx's real imported data, exactly the
// kind of gap "building it and using it for real" surfaces that spec-reading doesn't, and the spec
// has since been updated to make the unit explicit rather than leaving every engine to guess. A
// resize converts through csvx-ts's pixelsToColumnWidth (see setColumnWidth below) rather than
// writing a raw pixel number into this field, which would corrupt it for any other reader.

import { HEADER_ROW, canonicalCellText, columnId, columnIndexFromId, coordinateFor, indicesForCoordinate, nextCellMetadata, pixelsToColumnWidth, rawCellText, recalculateWorkbook, resolveCellValue, rowIndexFor, rowNumberFor } from 'csvx-ts/browser'

// The row convention (header = row 1, record i = row i + 2) and workbook recalculation are engine
// logic, owned by csvx-ts; they are re-exported here only so the rest of the app has one import site.
export { HEADER_ROW, coordinateFor, indicesForCoordinate, rawCellText, recalculateWorkbook }

/** The "used range" (like Excel's default print area): row and column counts up to the last cell
 * that prints something — a value, a formula, or a visible fill or border. Formatting that draws
 * nothing (number format, font only), common on imported XLSX files that format thousands of empty
 * rows, doesn't count. Always at least 1×1. */
export function usedRange(sheet, styles) {
  const byId = new Map((styles || []).map((style) => [style.id, style]))
  let rows = 1
  let columns = 1
  for (let rowIndex = HEADER_ROW; rowIndex < (sheet?.records?.length || 0); rowIndex += 1) {
    ;(sheet.columns || []).forEach((_, columnIndex) => {
      const metadata = sheet.cells?.[coordinateFor(columnIndex, rowIndex)]
      const style = metadata?.style ? byId.get(metadata.style) : undefined
      const drawsSomething = metadata?.formula || style?.fill?.color || ['top', 'right', 'bottom', 'left'].some((edge) => declaredEdge(style, edge))
      if (rawCellText(sheet, rowIndex, columnIndex) || drawsSomething) {
        rows = Math.max(rows, rowNumberFor(rowIndex))
        columns = Math.max(columns, columnIndex + 1)
      }
    })
  }
  return { rows, columns }
}

/** Merges `patch` into a sheet's `print` settings (spec/03-sheets.md). A key set to undefined or
 * null is removed, so it falls back to the spec default; an empty result removes `print` entirely.
 * Properties this app doesn't know about are left alone, as the spec requires. */
export function setPrintSettings(workbook, sheetId, patch) {
  return {
    ...workbook,
    sheets: workbook.sheets.map((sheet) => {
      if (sheet.id !== sheetId) return sheet
      const next = { ...sheet.print }
      for (const [key, value] of Object.entries(patch)) {
        if (value === undefined || value === null) delete next[key]
        else next[key] = value
      }
      const { print, ...rest } = sheet
      return Object.keys(next).length > 0 ? { ...rest, print: next } : rest
    }),
  }
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
  return {
    color: font.color || undefined,
    backgroundColor: fill.color || undefined,
    fontWeight: font.bold ? 700 : undefined,
    fontStyle: font.italic ? 'italic' : undefined,
    textDecoration: font.underline ? 'underline' : undefined,
    textAlign: alignment.horizontal || undefined,
    whiteSpace: alignment.wrapText ? 'normal' : 'nowrap',
  }
}

const BORDER_WIDTHS = { thin: '1px', medium: '2px', thick: '3px', dashed: '1px', dotted: '1px', double: '3px' }
const BORDER_LINES = { dashed: 'dashed', dotted: 'dotted', double: 'double' }

/** The color to show in the border tool for a style: the shorthand color, else the first declared
 * edge color (top, right, bottom, left). Undefined when the cell declares none. */
export function borderColorOf(style) {
  const border = style?.border || {}
  return border.color || ['top', 'right', 'bottom', 'left'].map((edge) => border[edge]?.color).find(Boolean)
}

/** One edge of a style's border as a CSS shorthand, or undefined when nothing is declared.
 * Top-level `style`/`color` are the all-edges shorthand; a per-edge object overrides them. */
function declaredEdge(style, edge) {
  const border = style?.border || {}
  const declared = { style: border.style, color: border.color, ...border[edge] }
  if ((!declared.style && !declared.color) || declared.style === 'none') return undefined
  const lineStyle = declared.style || 'thin'
  return `${BORDER_WIDTHS[lineStyle] || '1px'} ${BORDER_LINES[lineStyle] || 'solid'} ${declared.color || '#000000'}`
}

/** Literal mapping of spec/08-styles.md `border` onto a cell's own edges (applied to the `<td>`,
 * never as an inset outline). A border belongs to the edge shared by two cells, so each edge is
 * resolved with the neighbor: the later cell in reading order (right/bottom) wins, per the spec.
 * Doing this here instead of leaving it to `border-collapse` matters: on a tie the collapse rules
 * favor the left/top cell's default gridline, which would hide a declared left/top edge.
 * Edges with no declared border return nothing so the default gridlines still show. */
export function cellBorderCSS(style, { above, below, left, right } = {}) {
  const resolved = {
    borderTop: declaredEdge(style, 'top') ?? declaredEdge(above, 'bottom'),
    borderRight: declaredEdge(right, 'left') ?? declaredEdge(style, 'right'),
    borderBottom: declaredEdge(below, 'top') ?? declaredEdge(style, 'bottom'),
    borderLeft: declaredEdge(style, 'left') ?? declaredEdge(left, 'right'),
  }
  return Object.fromEntries(Object.entries(resolved).filter(([, value]) => value))
}

export function findSheet(workbook, sheetId) {
  return workbook.sheets.find((sheet) => sheet.id === sheetId) || workbook.sheets[0]
}

/** Sets a cell's raw CSV value. A value starting with "=" is stored as a formula (and the sheet is
 * recalculated via csvx-ts — see recalculateWorkbook — so `cached` and the visible CSV text update
 * immediately); anything else is a literal edit. Which metadata fields survive either kind of edit
 * (style/validation: yes; type/cached, and formula for a literal edit: no) is decided by csvx-ts's
 * nextCellMetadata, not here — see spec/05-cell-values.md and AGENTS.md rule 5.2: this app doesn't
 * get a second opinion about what a stale `type` or `formula` means once the content it described
 * is gone. */
export function setCellValue(workbook, sheetId, rowIndex, columnIndex, value) {
  if (rowIndex < 0) return setHeaderName(workbook, sheetId, columnIndex, value)
  const isFormula = typeof value === 'string' && value.startsWith('=')
  const next = {
    ...workbook,
    sheets: workbook.sheets.map((sheet) => {
      if (sheet.id !== sheetId) return sheet
      const records = sheet.records.map((row, index) => (index === rowIndex ? [...row] : row))
      const coordinate = coordinateFor(columnIndex, rowIndex)
      const cells = { ...sheet.cells }
      const nextMetadata = nextCellMetadata(cells[coordinate], isFormula ? value : undefined)
      if (nextMetadata) cells[coordinate] = nextMetadata
      else delete cells[coordinate]
      if (isFormula) {
        records[rowIndex][columnIndex] = value
      } else {
        // A literal typed against a cell whose style declares a numberFormat (e.g. "$7.00" into a
        // currency-styled cell) is canonicalized to plain numeric text via resolveCellValue's
        // numberFormat parsing — see spec/08-styles.md — rather than stored verbatim with the
        // currency symbol baked into the CSV text forever. Anything that doesn't resolve to a
        // number (a genuine string, a declared non-numeric column type, unparseable text) is stored
        // exactly as typed; this never applies to a formula (handled above) or an already-cleared
        // per-cell type override (nextCellMetadata just dropped it, so only the column's declared
        // type, if any, can still apply here).
        const declaredType = sheet.columns[columnIndex]?.type
        const numberFormat = styleForId(workbook.styles, nextMetadata?.style).numberFormat
        const resolved = resolveCellValue(value, declaredType, numberFormat)
        records[rowIndex][columnIndex] = resolved.type === 'integer' || resolved.type === 'decimal' ? canonicalCellText(resolved) : value
      }
      return { ...sheet, records, cells }
    }),
  }
  return recalculateWorkbook(next)
}

/** Editing row 1 renames the column: the header cell's text is Column.name (spec/03-sheets.md). An
 * empty value restores the column's own letter, since a CSV header cell can't be empty. Style
 * metadata on the header cell is untouched; a stale formula or cache on it is dropped by the engine
 * (nextCellMetadata) like any literal edit. */
function setHeaderName(workbook, sheetId, columnIndex, value) {
  const next = {
    ...workbook,
    sheets: workbook.sheets.map((sheet) => {
      if (sheet.id !== sheetId) return sheet
      const columns = sheet.columns.map((column, index) => (index === columnIndex ? { ...column, name: value === '' ? columnId(index) : value } : column))
      const coordinate = coordinateFor(columnIndex, HEADER_ROW)
      const cells = { ...sheet.cells }
      const nextMetadata = nextCellMetadata(cells[coordinate], undefined)
      if (nextMetadata) cells[coordinate] = nextMetadata
      else delete cells[coordinate]
      return { ...sheet, columns, cells }
    }),
  }
  return recalculateWorkbook(next)
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

/** Persists a user's drag-resize back into Column.width, converting the live pixel width to the
 * field's real unit (XLSX character-width units, per spec/03-sheets.md) via csvx-ts's
 * pixelsToColumnWidth — never writing a raw pixel number into this field, which would corrupt it
 * for any other tool reading the file (including a round-trip through this same app after a real
 * XLSX import, where the field already holds that unit). */
export function setColumnWidth(workbook, sheetId, columnIndex, pixelWidth) {
  const width = pixelsToColumnWidth(pixelWidth)
  return {
    ...workbook,
    sheets: workbook.sheets.map((sheet) => {
      if (sheet.id !== sheetId) return sheet
      const columns = sheet.columns.map((column, index) => (index === columnIndex ? { ...column, width } : column))
      return { ...sheet, columns }
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

/** Applies a literal font/fill/border/alignment/numberFormat patch to every given cell, each
 * getting its own new style record. This does not dedupe identical styles across cells — a real
 * editor would want to, but that's a quality improvement, not a correctness one; tracked as a known
 * simplification rather than solved here.
 *
 * `patch.numberFormat` has three states, since "leave it alone" and "remove it" are both real
 * cases a caller needs (font/fill/border/alignment don't need this because their per-property merge
 * already does the right thing without a sentinel): omit the key entirely to preserve whatever
 * numberFormat the cell already had (what every existing caller — bold/italic/underline/alignment —
 * does); pass a format string to set it; pass `null` to remove it.
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
    const numberFormat = 'numberFormat' in patch ? patch.numberFormat : existingStyle.numberFormat
    const style = {
      id: `s${nextId}`,
      ...(numberFormat ? { numberFormat } : {}),
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
    const row = rowIndexFor(Number(rowText))
    if (deleting && row === insertionRow) return '#REF!'
    if (row < insertionRow) return match
    return `${sheetPrefix || ''}${colAnchor}${column}${rowAnchor}${rowNumberFor(row + delta)}`
  })
}

function rewriteFormulaColumns(formula, targetSheetName, insertionColumn, delta, deleting) {
  if (!formula) return formula
  return formula.replace(CELL_REFERENCE_PATTERN, (match, sheetPrefix, colAnchor, column, rowAnchor, rowText) => {
    if (referencedSheetName(sheetPrefix, targetSheetName) !== targetSheetName) return match
    const columnIndex = columnIndexFromId(column)
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

// Reference text is already rewritten by shiftSheetRows/shiftSheetColumns above; recalculateWorkbook
// (re-run here) is what turns those rewritten formulas into updated cached values and visible text.

export function insertRow(workbook, sheetId, rowIndex) {
  return recalculateWorkbook(shiftSheetRows(workbook, sheetId, rowIndex, 1, false))
}

export function deleteRow(workbook, sheetId, rowIndex) {
  return recalculateWorkbook(shiftSheetRows(workbook, sheetId, rowIndex, -1, true))
}

export function insertColumn(workbook, sheetId, columnIndex) {
  return recalculateWorkbook(shiftSheetColumns(workbook, sheetId, columnIndex, 1, false))
}

export function deleteColumn(workbook, sheetId, columnIndex) {
  return recalculateWorkbook(shiftSheetColumns(workbook, sheetId, columnIndex, -1, true))
}
