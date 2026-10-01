// Workbook-shape helpers used by App.jsx. These only rearrange data the engine (csvx-ts) already
// produced or will accept back — none of them decide what a value, type, or style *means*. That
// distinction is the whole point of this rewrite: see AGENTS.md and ../csvx-spec/AGENTS.md rule 5.
//
// Specifically NOT here (on purpose, because no engine implements it yet): cell-type inference
// from raw text, computed number formatting (parsing "$#,##0.00" and producing "$1,234.56"), and
// formula-reference rewriting on row/column insert or delete. Row/column *append* is safe without
// that machinery (nothing shifts), so it's supported; insert-in-the-middle and delete are not.
//
// Also notably absent: anything that writes to Column.width. schemas/sheet-metadata.schema.json
// never defines a unit for it, and real imported data (example.csvx, via csvx-go's XLSX importer)
// stores it in XLSX character-width units (e.g. 26.25), not CSS pixels — confirmed by actually
// running this app against that fixture, which is exactly the kind of gap "building it and using
// it for real" surfaces that spec-reading doesn't. Column resize in App.jsx is therefore
// view-only state, local to the browser session, and never touches this field — writing a pixel
// number into it would silently corrupt the original unit on export.

import { columnId } from 'csvx-ts/browser'

export function coordinateFor(columnIndex, rowIndex) {
  return `${columnId(columnIndex)}${rowIndex + 1}`
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

/** Appends a blank column at the end of a sheet. Same reasoning as appendRow. */
export function appendColumn(workbook, sheetId) {
  return {
    ...workbook,
    sheets: workbook.sheets.map((sheet) => {
      if (sheet.id !== sheetId) return sheet
      const columns = [...sheet.columns, { id: columnId(sheet.columns.length), name: `Column ${sheet.columns.length + 1}` }]
      const records = sheet.records.map((row) => [...row, ''])
      return { ...sheet, columns, records }
    }),
  }
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
