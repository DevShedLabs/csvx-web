import { useEffect, useMemo, useRef, useState } from 'react'

const demoWorkbook = {
  name: 'example.csvx', version: '1.0',
  sheets: [{ id: 'sheet-1', name: 'Sheet 1', columns: ['Value'], rows: [['1'], ['2']], cells: {} }],
  styles: {}, source: null,
}
const STORED_PACKAGE_KEY = 'csvx-web.current-package'

function readUInt16(bytes, offset) { return bytes[offset] | (bytes[offset + 1] << 8) }
function readUInt32(bytes, offset) { return bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24) }
function decode(bytes) { return new TextDecoder().decode(bytes) }

async function readZipEntries(buffer) {
  const bytes = new Uint8Array(buffer); let end = -1
  for (let index = bytes.length - 22; index >= 0; index -= 1) if (readUInt32(bytes, index) === 0x06054b50) { end = index; break }
  if (end < 0) throw new Error('Not a CSVX ZIP package')
  const entries = new Map(); let offset = readUInt32(bytes, end + 16); const count = readUInt16(bytes, end + 10)
  for (let index = 0; index < count; index += 1) {
    if (readUInt32(bytes, offset) !== 0x02014b50) throw new Error('Invalid CSVX ZIP directory')
    const method = readUInt16(bytes, offset + 10); const compressedSize = readUInt32(bytes, offset + 20); const nameLength = readUInt16(bytes, offset + 28); const extraLength = readUInt16(bytes, offset + 30); const commentLength = readUInt16(bytes, offset + 32)
    const name = decode(bytes.subarray(offset + 46, offset + 46 + nameLength)); const localOffset = readUInt32(bytes, offset + 42); const localNameLength = readUInt16(bytes, localOffset + 26); const localExtraLength = readUInt16(bytes, localOffset + 28); const start = localOffset + 30 + localNameLength + localExtraLength
    const compressed = bytes.slice(start, start + compressedSize); let content = compressed
    if (method === 8) content = new Uint8Array(await new Response(new Blob([compressed]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer())
    if (method !== 0 && method !== 8) throw new Error(`Unsupported ZIP compression for ${name}`)
    entries.set(name, content); offset += 46 + nameLength + extraLength + commentLength
  }
  return entries
}

function parseCSV(text) {
  const rows = []; let row = []; let value = ''; let quoted = false; let endedLine = false
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index]; endedLine = false
    if (character === '"') { if (quoted && text[index + 1] === '"') { value += '"'; index += 1 } else quoted = !quoted }
    else if (character === ',' && !quoted) { row.push(value); value = '' }
    else if ((character === '\n' || character === '\r') && !quoted) { if (character === '\r' && text[index + 1] === '\n') index += 1; row.push(value); value = ''; rows.push(row); row = []; endedLine = true }
    else value += character
  }
  if (value || row.length > 0 || !endedLine) { row.push(value); rows.push(row) }
  while (rows.length > 1 && rows[rows.length - 1].every((cell) => cell === '')) rows.pop()
  if (!rows.length) throw new Error('CSV sheet has no header row')
  return { columns: rows[0], rows: rows.slice(1) }
}

async function parseCSVX(file) {
  const entries = await readZipEntries(await file.arrayBuffer()); const readJSON = (name) => { const body = entries.get(name); if (!body) throw new Error(`Missing ${name}`); return JSON.parse(decode(body)) }
  const manifest = readJSON('manifest.json'); const workbook = readJSON(manifest.workbook); const styleResource = workbook.styles ? readJSON(workbook.styles) : { styles: {} }
  const styles = Array.isArray(styleResource.styles) ? Object.fromEntries(styleResource.styles.map((style) => [style.id, style])) : styleResource.styles || {}
  const sheets = workbook.sheets.map((entry) => { const csv = parseCSV(decode(entries.get(entry.path))); const metadata = entry.metadata ? readJSON(entry.metadata) : { cells: {} }; return { id: entry.id, name: entry.name, ...csv, cells: metadata.cells || {}, rowHeights: metadata.rowHeights || {}, metadata } })
  return { name: file.name, version: workbook.version, sheets, styles, source: workbook.source || null, manifest }
}

function formatCellValue(value, metadata, styles) {
  if (value === '') return ''; const format = metadata?.style ? styles?.[metadata.style]?.numberFormat : ''; if (!format || Number.isNaN(Number(value))) return value
  const numericValue = Number(value); const decimalPart = format.split('.')[1] || ''; const formatDecimals = (decimalPart.match(/0/g) || []).length; const rawDecimals = value.includes('.') ? value.split('.')[1].length : 0; const wouldLosePrecision = Number(numericValue.toFixed(formatDecimals)) !== numericValue; const decimals = (!format.includes('%') && wouldLosePrecision) ? Math.max(formatDecimals, rawDecimals) : formatDecimals; const outputValue = format.includes('%') ? numericValue * 100 : numericValue; const rounded = outputValue.toFixed(decimals); const [whole, fraction] = rounded.split('.'); const grouped = format.includes(',') ? whole.replace(/\B(?=(\d{3})+(?!\d))/g, ',') : whole; const number = fraction ? `${grouped}.${fraction}` : grouped
  if (format.includes('%')) return `${number}%`; if (format.includes('$')) return `$${number}`; if (format.includes('€')) return `€${number}`; if (format.includes('£')) return `£${number}`; return number
}
function describeCellType(metadata, styles) {
  const format = metadata?.style ? styles?.[metadata.style]?.numberFormat : ''
  if (format) {
    if (/[$€£¥]/.test(format)) return 'currency'
    if (format.includes('%')) return 'percentage'
  }
  return metadata?.type || ''
}
function cellStyle(value, metadata, styles) { const style = metadata?.style ? styles?.[metadata.style] : null; const font = style?.font || {}; const fill = style?.fill || {}; const alignment = style?.alignment || {}; const border = style?.border || {}; const numericType = ['integer', 'decimal'].includes(metadata?.type); const numericValue = value !== '' && !Number.isNaN(Number(value)); return { color: font.color || undefined, backgroundColor: fill.color || undefined, fontWeight: font.bold ? 700 : undefined, fontStyle: font.italic ? 'italic' : undefined, textDecoration: font.underline ? 'underline' : undefined, textAlign: alignment.horizontal || ((numericType || numericValue) ? 'right' : undefined), whiteSpace: alignment.wrapText ? 'normal' : 'nowrap', boxShadow: border.color ? `inset 0 0 0 0.0625rem ${border.color}` : undefined } }
function styleFor(metadata, styles) { return metadata?.style ? styles?.[metadata.style] || {} : {} }
function nextStyleId(styles) { const numericIds = Object.keys(styles || {}).map(Number).filter((value) => !Number.isNaN(value)); return String((numericIds.length ? Math.max(...numericIds) : -1) + 1) }
function applyCellFormat(workbook, targetSheet, coordinate, patch) {
  const existingMetadata = targetSheet.cells?.[coordinate] || {}
  const existingStyle = styleFor(existingMetadata, workbook.styles)
  const nextStyle = { ...existingStyle, font: { ...existingStyle.font, ...patch.font }, fill: { ...existingStyle.fill, ...patch.fill }, border: { ...existingStyle.border, ...patch.border }, alignment: { ...existingStyle.alignment, ...patch.alignment } }
  const styleId = nextStyleId(workbook.styles)
  return { ...workbook, styles: { ...workbook.styles, [styleId]: nextStyle }, sheets: workbook.sheets.map((item) => item.id !== targetSheet.id ? item : { ...item, cells: { ...item.cells, [coordinate]: { ...existingMetadata, style: styleId } } }) }
}
function applyCellsFormat(workbook, targetSheet, coordinates, patch) {
  let next = workbook
  coordinates.forEach((coordinate) => { next = applyCellFormat(next, targetSheet, coordinate, patch) })
  return next
}
function columnLabel(index) { let label = ''; let value = index + 1; while (value > 0) { const remainder = (value - 1) % 26; label = String.fromCharCode(65 + remainder) + label; value = Math.floor((value - 1) / 26) } return label }
function columnIndexFromLabel(label) { return label.split('').reduce((total, character) => total * 26 + character.charCodeAt(0) - 64, 0) - 1 }
function cellCoordinate(column, row) { return `${columnLabel(column)}${row + 1}` }
function cellPosition(coordinate) { const match = coordinate.match(/^([A-Z]+)(\d+)$/); if (!match) return { column: 0, row: 0 }; return { column: columnIndexFromLabel(match[1]), row: Number(match[2]) - 1 } }
const DEFAULT_COLUMN_WIDTH = 128
const MIN_COLUMN_WIDTH = 48
const ARROW_DIRECTIONS = { ArrowLeft: { column: -1, row: 0 }, ArrowRight: { column: 1, row: 0 }, ArrowUp: { column: 0, row: -1 }, ArrowDown: { column: 0, row: 1 } }
function cloneSheet(sheet) { return { ...sheet, rows: sheet.rows.map((row) => [...row]), cells: { ...sheet.cells } } }
function getCellText(sheet, coordinate) { const target = cellPosition(coordinate); return sheet.rows[target.row]?.[target.column] ?? '' }
function inferCellType(value) {
  if (value === '') return 'blank'
  if (value === 'true' || value === 'false') return 'boolean'
  if (/^-?\d+$/.test(value)) return 'integer'
  if (/^-?\d+\.\d+$/.test(value)) return 'decimal'
  return 'string'
}
function setCellText(workbook, sheetId, coordinate, value) { const target = cellPosition(coordinate); return { ...workbook, sheets: workbook.sheets.map((item) => { if (item.id !== sheetId) return item; const nextSheet = cloneSheet(item); while (nextSheet.rows.length <= target.row) nextSheet.rows.push(Array(nextSheet.columns.length).fill('')); while (nextSheet.rows[target.row].length < nextSheet.columns.length) nextSheet.rows[target.row].push(''); nextSheet.rows[target.row][target.column] = value; const { formula, ...preservedMetadata } = nextSheet.cells[coordinate] || {}; if (value === '' && Object.keys(preservedMetadata).length === 0) { delete nextSheet.cells[coordinate] } else { nextSheet.cells[coordinate] = { ...preservedMetadata, type: inferCellType(value) } } return nextSheet }) } }
function rewriteFormulaRows(formula, targetSheetName, insertionRow, delta, deleting) {
  if (!formula) return formula
  const referencePattern = /((?:'[^']+'|[A-Za-z0-9_ ]+)!)?([A-Z]+)(\d+)/g
  return formula.replace(referencePattern, (match, sheetPrefix = '', column, rowText) => {
    const referencedSheet = sheetPrefix ? sheetPrefix.slice(0, -1).replace(/^'|'$/g, "").replace(/''/g, "'") : targetSheetName
    if (referencedSheet !== targetSheetName) return match
    const row = Number(rowText)
    if (deleting && row === insertionRow) return '#REF!'
    if (row < insertionRow) return match
    return `${sheetPrefix}${column}${row + delta}`
  })
}
function remapRows(workbook, targetSheet, insertionIndex, delta, deleting) {
  const insertionRow = insertionIndex + 1
  return { ...workbook, sheets: workbook.sheets.map((item) => {
    const nextSheet = cloneSheet(item)
    const nextCells = {}
    Object.entries(nextSheet.cells || {}).forEach(([coordinate, metadata]) => {
      const position = cellPosition(coordinate)
      if (item.id === targetSheet.id && position.row + 1 >= insertionRow) {
        if (!(deleting && position.row + 1 === insertionRow)) nextCells[cellCoordinate(position.column, position.row + delta)] = metadata
        return
      }
      nextCells[coordinate] = { ...metadata, formula: rewriteFormulaRows(metadata.formula, targetSheet.name, insertionRow, delta, deleting) }
    })
    nextSheet.cells = nextCells
    if (deleting) nextSheet.rows.splice(insertionIndex, 1)
    else nextSheet.rows.splice(insertionIndex, 0, Array(nextSheet.columns.length).fill(''))
    return nextSheet
  }) }
}
function rewriteFormulaColumns(formula, targetSheetName, insertionIndex, delta, deleting) {
  if (!formula) return formula
  const referencePattern = /((?:'[^']+'|[A-Za-z0-9_ ]+)!)?([A-Z]+)(\d+)/g
  return formula.replace(referencePattern, (match, sheetPrefix = '', column, rowText) => {
    const referencedSheet = sheetPrefix ? sheetPrefix.slice(0, -1).replace(/^'|'$/g, "").replace(/''/g, "'") : targetSheetName
    if (referencedSheet !== targetSheetName) return match
    const columnIndex = columnIndexFromLabel(column)
    if (deleting && columnIndex === insertionIndex) return '#REF!'
    if (columnIndex < insertionIndex) return match
    return `${sheetPrefix}${columnLabel(columnIndex + delta)}${rowText}`
  })
}
function remapColumns(workbook, targetSheet, insertionIndex, delta, deleting) {
  return { ...workbook, sheets: workbook.sheets.map((item) => {
    const nextSheet = cloneSheet(item)
    const nextCells = {}
    Object.entries(nextSheet.cells || {}).forEach(([coordinate, metadata]) => {
      const position = cellPosition(coordinate)
      if (item.id === targetSheet.id && position.column >= insertionIndex) {
        if (!(deleting && position.column === insertionIndex)) nextCells[cellCoordinate(position.column + delta, position.row)] = metadata
        return
      }
      nextCells[coordinate] = { ...metadata, formula: rewriteFormulaColumns(metadata.formula, targetSheet.name, insertionIndex, delta, deleting) }
    })
    nextSheet.cells = nextCells
    if (item.id === targetSheet.id) {
      const nextColumnWidths = {}
      Object.entries(nextSheet.columnWidths || {}).forEach(([key, width]) => {
        const index = Number(key)
        if (deleting && index === insertionIndex) return
        nextColumnWidths[index >= insertionIndex ? index + delta : index] = width
      })
      nextSheet.columnWidths = nextColumnWidths
      if (deleting) {
        nextSheet.columns = nextSheet.columns.filter((_, index) => index !== insertionIndex)
        nextSheet.rows = nextSheet.rows.map((row) => row.filter((_, index) => index !== insertionIndex))
      } else {
        nextSheet.columns = [...nextSheet.columns]; nextSheet.columns.splice(insertionIndex, 0, '')
        nextSheet.rows = nextSheet.rows.map((row) => { const nextRow = [...row]; nextRow.splice(insertionIndex, 0, ''); return nextRow })
      }
    }
    return nextSheet
  }) }
}
function setColumnWidth(workbook, sheetId, columnIndex, width) {
  return { ...workbook, sheets: workbook.sheets.map((item) => item.id !== sheetId ? item : { ...item, columnWidths: { ...item.columnWidths, [columnIndex]: width } }) }
}
const NEW_SHEET_COLUMN_COUNT = 26
const NEW_SHEET_ROW_COUNT = 100
const NEW_SHEET_COLUMN_WIDTH = 100
function createSheet(workbook) {
  const usedIds = new Set(workbook.sheets.map((item) => item.id))
  const usedNames = new Set(workbook.sheets.map((item) => item.name))
  let index = workbook.sheets.length + 1
  while (usedIds.has(`sheet-${index}`) || usedNames.has(`Sheet ${index}`)) index += 1
  const columnWidths = {}
  for (let column = 0; column < NEW_SHEET_COLUMN_COUNT; column += 1) columnWidths[column] = NEW_SHEET_COLUMN_WIDTH
  return {
    id: `sheet-${index}`,
    name: `Sheet ${index}`,
    columns: Array(NEW_SHEET_COLUMN_COUNT).fill(''),
    rows: Array.from({ length: NEW_SHEET_ROW_COUNT }, () => Array(NEW_SHEET_COLUMN_COUNT).fill('')),
    cells: {},
    rowHeights: {},
    columnWidths,
  }
}
function renameSheet(workbook, sheetId, name) {
  const trimmed = name.trim()
  if (!trimmed) return workbook
  return { ...workbook, sheets: workbook.sheets.map((item) => item.id === sheetId ? { ...item, name: trimmed } : item) }
}
function deleteSheet(workbook, sheetId) {
  return { ...workbook, sheets: workbook.sheets.filter((item) => item.id !== sheetId) }
}

function App() {
  const [workbook, setWorkbook] = useState(demoWorkbook); const [activeSheet, setActiveSheet] = useState('sheet-1'); const [selectedCell, setSelectedCell] = useState('A1'); const [selectedCells, setSelectedCells] = useState(() => new Set(['A1'])); const [selectedColumns, setSelectedColumns] = useState(() => new Set()); const [selectedRows, setSelectedRows] = useState(() => new Set()); const [editingCell, setEditingCell] = useState(null); const [draftValue, setDraftValue] = useState(''); const [contextMenu, setContextMenu] = useState(null); const [error, setError] = useState(''); const [renamingSheetId, setRenamingSheetId] = useState(null); const [sheetNameDraft, setSheetNameDraft] = useState(''); const inputRef = useRef(null); const editorRef = useRef(null); const cellRefs = useRef(new Map()); const colRefs = useRef(new Map()); const contextMenuRef = useRef(null); const tableScrollRef = useRef(null); const sheetNameInputRef = useRef(null)
  const sheet = useMemo(() => workbook.sheets.find((item) => item.id === activeSheet) || workbook.sheets[0], [activeSheet, workbook]); const selectedMetadata = sheet?.cells?.[selectedCell]; const position = cellPosition(selectedCell); const selectedValue = sheet?.rows?.[position.row]?.[position.column] || ''; const selectedDisplayValue = selectedMetadata?.formula || formatCellValue(selectedValue, selectedMetadata, workbook.styles) || 'Blank cell'; const selectedType = selectedDisplayValue !== 'Blank cell' ? describeCellType(selectedMetadata, workbook.styles) : ''
  const selectedStyle = styleFor(selectedMetadata, workbook.styles); const selectedAlignment = selectedStyle.alignment?.horizontal || ''
  const selectionStyles = useMemo(() => [...selectedCells].map((coordinate) => styleFor(sheet?.cells?.[coordinate], workbook.styles)), [selectedCells, sheet, workbook.styles])
  const allSelectedHaveFont = (property) => selectionStyles.length > 0 && selectionStyles.every((style) => style.font?.[property])
  const allSelectedHaveAlignment = (value) => selectionStyles.length > 0 && selectionStyles.every((style) => (style.alignment?.horizontal || '') === value)
  useEffect(() => { const stored = localStorage.getItem(STORED_PACKAGE_KEY); if (!stored) return; try { const restored = JSON.parse(stored); if (restored?.name && Array.isArray(restored.sheets)) { setWorkbook(restored); setActiveSheet(restored.sheets[0]?.id) } } catch { localStorage.removeItem(STORED_PACKAGE_KEY) } }, [])
  useEffect(() => { if (editingCell) { const input = editorRef.current; input?.focus(); input?.setSelectionRange(input.value.length, input.value.length) } }, [editingCell])
  useEffect(() => { if (renamingSheetId) { sheetNameInputRef.current?.focus(); sheetNameInputRef.current?.select() } }, [renamingSheetId])
  useEffect(() => { if (!editingCell) cellRefs.current.get(selectedCell)?.focus() }, [selectedCell, editingCell, activeSheet])
  useEffect(() => { const closeMenu = () => setContextMenu(null); document.addEventListener('click', closeMenu); return () => document.removeEventListener('click', closeMenu) }, [])
  useEffect(() => {
    const el = tableScrollRef.current
    if (!el) return
    let axis = null; let resetTimer = null
    function handleWheel(event) {
      if (resetTimer) clearTimeout(resetTimer)
      resetTimer = setTimeout(() => { axis = null }, 200)
      if (!axis) axis = Math.abs(event.deltaX) > Math.abs(event.deltaY) ? 'x' : 'y'
      event.preventDefault()
      if (axis === 'x') el.scrollLeft += event.deltaX
      else el.scrollTop += event.deltaY
    }
    el.addEventListener('wheel', handleWheel, { passive: false })
    return () => { el.removeEventListener('wheel', handleWheel); if (resetTimer) clearTimeout(resetTimer) }
  }, [])
  useEffect(() => {
    function handleShortcut(event) {
      if (!(event.metaKey || event.ctrlKey)) return
      const key = event.key.toLowerCase()
      if (key === 'b') { event.preventDefault(); toggleFont('bold') }
      else if (key === 'i') { event.preventDefault(); toggleFont('italic') }
      else if (key === 'u') { event.preventDefault(); toggleFont('underline') }
    }
    document.addEventListener('keydown', handleShortcut)
    return () => document.removeEventListener('keydown', handleShortcut)
  }, [workbook, sheet, selectedCell])
  useEffect(() => {
    if (!contextMenu || !contextMenuRef.current) return
    const rect = contextMenuRef.current.getBoundingClientRect(); const margin = 8
    const maxLeft = window.innerWidth - rect.width - margin; const maxTop = window.innerHeight - rect.height - margin
    const clampedLeft = Math.min(contextMenu.x, Math.max(margin, maxLeft)); const clampedTop = Math.min(contextMenu.y, Math.max(margin, maxTop))
    if (clampedLeft !== contextMenu.x || clampedTop !== contextMenu.y) setContextMenu((current) => current ? { ...current, x: clampedLeft, y: clampedTop } : current)
  }, [contextMenu])
  function updateWorkbook(nextWorkbook) { setWorkbook(nextWorkbook); localStorage.setItem(STORED_PACKAGE_KEY, JSON.stringify(nextWorkbook)) }
  function selectCell(coordinate) { setSelectedCell(coordinate); setSelectedCells(new Set([coordinate])); setSelectedColumns(new Set()); setSelectedRows(new Set()); setContextMenu(null) }
  function switchSheet(sheetId) { setActiveSheet(sheetId); setSelectedCell('A1'); setSelectedCells(new Set(['A1'])); setSelectedColumns(new Set()); setSelectedRows(new Set()); setContextMenu(null) }
  function addSheet() { const newSheet = createSheet(workbook); updateWorkbook({ ...workbook, sheets: [...workbook.sheets, newSheet] }); switchSheet(newSheet.id) }
  function beginRenameSheet(sheetId, currentName) { setRenamingSheetId(sheetId); setSheetNameDraft(currentName); setContextMenu(null) }
  function commitRenameSheet() { if (!renamingSheetId) return; updateWorkbook(renameSheet(workbook, renamingSheetId, sheetNameInputRef.current?.value ?? sheetNameDraft)); setRenamingSheetId(null) }
  function cancelRenameSheet() { setRenamingSheetId(null) }
  function removeSheet(sheetId) {
    if (workbook.sheets.length <= 1) { setError('A workbook must contain at least one sheet.'); setContextMenu(null); return }
    const index = workbook.sheets.findIndex((item) => item.id === sheetId)
    const next = deleteSheet(workbook, sheetId)
    updateWorkbook(next)
    if (activeSheet === sheetId) { const fallback = next.sheets[Math.max(0, index - 1)] || next.sheets[0]; switchSheet(fallback.id) }
    setContextMenu(null)
  }
  function openSheetContextMenu(event, sheetId) { event.preventDefault(); setContextMenu({ x: event.clientX, y: event.clientY, kind: 'sheet', index: sheetId }) }
  function toggleCellSelection(coordinate) {
    setSelectedCell(coordinate)
    setSelectedCells((current) => { const next = new Set(current); if (next.has(coordinate)) next.delete(coordinate); else next.add(coordinate); return next.size ? next : new Set([coordinate]) })
    setContextMenu(null)
  }
  function selectColumn(index, additive) {
    const coordinates = sheet.rows.map((_, rowIndex) => cellCoordinate(index, rowIndex))
    setSelectedCell(cellCoordinate(index, 0))
    if (additive) { setSelectedCells((current) => new Set([...current, ...coordinates])); setSelectedColumns((current) => new Set(current).add(index)) }
    else { setSelectedCells(new Set(coordinates)); setSelectedColumns(new Set([index])); setSelectedRows(new Set()) }
    setContextMenu(null)
  }
  function selectRow(index, additive) {
    const coordinates = sheet.columns.map((_, columnIndex) => cellCoordinate(columnIndex, index))
    setSelectedCell(cellCoordinate(0, index))
    if (additive) { setSelectedCells((current) => new Set([...current, ...coordinates])); setSelectedRows((current) => new Set(current).add(index)) }
    else { setSelectedCells(new Set(coordinates)); setSelectedRows(new Set([index])); setSelectedColumns(new Set()) }
    setContextMenu(null)
  }
  function beginEdit(coordinate, initialValue) { setSelectedCell(coordinate); setDraftValue(initialValue !== undefined ? initialValue : getCellText(sheet, coordinate)); setEditingCell(coordinate); setContextMenu(null) }
  function commitEdit() { if (!editingCell) return; const value = editorRef.current?.value ?? draftValue; updateWorkbook(setCellText(workbook, sheet.id, editingCell, value)); setEditingCell(null) }
  function cancelEdit() { setEditingCell(null) }
  function moveSelection(direction) { const nextColumn = Math.max(0, Math.min(sheet.columns.length - 1, position.column + direction.column)); const nextRow = Math.max(0, Math.min(sheet.rows.length - 1, position.row + direction.row)); selectCell(cellCoordinate(nextColumn, nextRow)) }
  function commitEditAndMove(direction) { if (!editingCell) return; const value = editorRef.current?.value ?? draftValue; updateWorkbook(setCellText(workbook, sheet.id, editingCell, value)); setEditingCell(null); moveSelection(direction) }
  function toggleFont(property) { updateWorkbook(applyCellsFormat(workbook, sheet, selectedCells, { font: { [property]: !allSelectedHaveFont(property) } })); cellRefs.current.get(selectedCell)?.focus() }
  function setAlignment(value) { updateWorkbook(applyCellsFormat(workbook, sheet, selectedCells, { alignment: { horizontal: allSelectedHaveAlignment(value) ? '' : value } })); cellRefs.current.get(selectedCell)?.focus() }
  function setTextColor(color) { updateWorkbook(applyCellsFormat(workbook, sheet, selectedCells, { font: { color } })); cellRefs.current.get(selectedCell)?.focus() }
  function setFillColor(color) { updateWorkbook(applyCellsFormat(workbook, sheet, selectedCells, { fill: { color, pattern: 'solid' } })); cellRefs.current.get(selectedCell)?.focus() }
  function setBorderColor(color) { updateWorkbook(applyCellsFormat(workbook, sheet, selectedCells, { border: { color } })); cellRefs.current.get(selectedCell)?.focus() }
  function clearFormatting() {
    updateWorkbook({ ...workbook, sheets: workbook.sheets.map((item) => { if (item.id !== sheet.id) return item; const nextCells = { ...item.cells }; selectedCells.forEach((coordinate) => { const { style, ...rest } = nextCells[coordinate] || {}; if (Object.keys(rest).length) nextCells[coordinate] = rest; else delete nextCells[coordinate] }); return { ...item, cells: nextCells } }) })
    cellRefs.current.get(selectedCell)?.focus()
  }
  function addRow(index) { const next = remapRows(workbook, sheet, index, 1, false); updateWorkbook(next); setSelectedCell(cellCoordinate(0, index)); setContextMenu(null) }
  function deleteRow(index) { if (sheet.rows.length <= 1) { setError('A sheet must contain at least one data row.'); setContextMenu(null); return } const next = remapRows(workbook, sheet, index, -1, true); updateWorkbook(next); setSelectedCell(cellCoordinate(0, Math.max(0, index - 1))); setContextMenu(null) }
  function addColumn(index) { const next = remapColumns(workbook, sheet, index, 1, false); updateWorkbook(next); setSelectedCell(cellCoordinate(index, 0)); setContextMenu(null) }
  function deleteColumn(index) { if (sheet.columns.length <= 1) { setError('A sheet must contain at least one column.'); setContextMenu(null); return } const next = remapColumns(workbook, sheet, index, -1, true); updateWorkbook(next); setSelectedCell(cellCoordinate(Math.max(0, index - 1), 0)); setContextMenu(null) }
  function openContextMenu(event, kind, index) { event.preventDefault(); setSelectedCell(kind === 'row' ? cellCoordinate(0, index) : cellCoordinate(index, 0)); setContextMenu({ x: event.clientX, y: event.clientY, kind, index }) }
  function startColumnResize(event, columnIndex) {
    event.preventDefault(); event.stopPropagation()
    const startX = event.clientX; const startWidth = sheet.columnWidths?.[columnIndex] || DEFAULT_COLUMN_WIDTH; const colElement = colRefs.current.get(columnIndex)
    document.body.style.cursor = 'col-resize'; document.body.style.userSelect = 'none'
    function widthAt(moveEvent) { return Math.max(MIN_COLUMN_WIDTH, startWidth + (moveEvent.clientX - startX)) }
    function handleMove(moveEvent) { if (colElement) colElement.style.width = `${widthAt(moveEvent)}px` }
    function handleUp(moveEvent) {
      document.removeEventListener('mousemove', handleMove); document.removeEventListener('mouseup', handleUp)
      document.body.style.cursor = ''; document.body.style.userSelect = ''
      updateWorkbook(setColumnWidth(workbook, sheet.id, columnIndex, widthAt(moveEvent)))
    }
    document.addEventListener('mousemove', handleMove); document.addEventListener('mouseup', handleUp)
  }
  function handleCellKeyDown(event, coordinate) {
    if (editingCell === coordinate) { if (event.key === 'Enter') { event.preventDefault(); commitEdit() } if (event.key === 'Escape') { event.preventDefault(); cancelEdit() } return }
    if (event.key === 'Enter' || event.key === 'F2') { event.preventDefault(); beginEdit(coordinate); return }
    if (event.key === 'Backspace' || event.key === 'Delete') { event.preventDefault(); let next = workbook; selectedCells.forEach((target) => { next = setCellText(next, sheet.id, target, '') }); updateWorkbook(next); return }
    if (ARROW_DIRECTIONS[event.key]) { event.preventDefault(); moveSelection(ARROW_DIRECTIONS[event.key]); return }
    if (event.key === ' ') { event.preventDefault(); beginEdit(coordinate); return }
    if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) { event.preventDefault(); beginEdit(coordinate, event.key) }
  }
  async function handleOpen(event) { const file = event.target.files?.[0]; if (!file) return; try { const loaded = await parseCSVX(file); setWorkbook(loaded); setActiveSheet(loaded.sheets[0]?.id); setSelectedCell('A1'); setError(''); localStorage.setItem(STORED_PACKAGE_KEY, JSON.stringify(loaded)) } catch (loadError) { setError(loadError.message); setWorkbook(demoWorkbook) } event.target.value = '' }
  return <div className="app-shell"><header className="topbar"><div className="brand-lockup"><span className="brand-mark" aria-hidden="true">X</span><div><p className="eyebrow">CSVX demo</p><h1>Workbench</h1></div></div><div className="file-status" role="status"><span className="status-dot" aria-hidden="true" /><span>{workbook.name}</span><span className="muted">Core {workbook.version}</span></div><nav className="top-actions" aria-label="File actions"><input ref={inputRef} type="file" accept=".csvx,application/zip" onChange={handleOpen} className="sr-only" aria-label="Open CSVX package" /><button type="button" className="button button-quiet" onClick={() => inputRef.current?.click()}>Open</button><button type="button" className="button button-primary" disabled aria-disabled="true">Export</button></nav></header><main id="main-content" className="main-content" tabIndex="-1">{error && <p role="alert">{error}</p>}<div className="content-header"><div><p className="eyebrow">Sheet / {sheet.name}</p><h2>{sheet.name}</h2></div><span className="read-only-badge">Local edits</span></div><section className="formula-panel" aria-label="Cell inspector"><div className="name-box">{selectedCell}</div><div className="formula-symbol" aria-hidden="true">fx</div><div className="formula-value">{selectedDisplayValue}{selectedType ? <span className="value-type">{selectedType}</span> : null}</div><div className="format-toolbar" role="toolbar" aria-label="Cell formatting"><button type="button" className={`format-button ${allSelectedHaveFont('bold') ? 'is-active' : ''}`} aria-pressed={allSelectedHaveFont('bold')} onMouseDown={(event) => event.preventDefault()} onClick={() => toggleFont('bold')} aria-label="Bold" title="Bold (Cmd+B)"><strong>B</strong></button><button type="button" className={`format-button ${allSelectedHaveFont('italic') ? 'is-active' : ''}`} aria-pressed={allSelectedHaveFont('italic')} onMouseDown={(event) => event.preventDefault()} onClick={() => toggleFont('italic')} aria-label="Italic" title="Italic (Cmd+I)"><em>I</em></button><button type="button" className={`format-button ${allSelectedHaveFont('underline') ? 'is-active' : ''}`} aria-pressed={allSelectedHaveFont('underline')} onMouseDown={(event) => event.preventDefault()} onClick={() => toggleFont('underline')} aria-label="Underline" title="Underline (Cmd+U)"><span className="format-underline-glyph">U</span></button><span className="format-divider" aria-hidden="true" /><label className="format-button color-swatch" title="Text color"><span aria-hidden="true" className="color-swatch-text" style={{ borderBottomColor: selectedStyle.font?.color || '#202124' }}>A</span><input type="color" value={selectedStyle.font?.color || '#000000'} onChange={(event) => setTextColor(event.target.value)} aria-label="Text color" /></label><label className="format-button color-swatch" title="Fill color"><span aria-hidden="true" className="color-swatch-fill" style={{ backgroundColor: selectedStyle.fill?.color || 'transparent' }} /><input type="color" value={selectedStyle.fill?.color || '#ffffff'} onChange={(event) => setFillColor(event.target.value)} aria-label="Fill color" /></label><label className="format-button color-swatch" title="Border color"><span aria-hidden="true" className="color-swatch-border" style={{ borderColor: selectedStyle.border?.color || 'var(--border-strong)' }} /><input type="color" value={selectedStyle.border?.color || '#000000'} onChange={(event) => setBorderColor(event.target.value)} aria-label="Border color" /></label><button type="button" className="format-button format-clear" onMouseDown={(event) => event.preventDefault()} onClick={clearFormatting} aria-label="Clear formatting" title="Clear formatting">Clear</button><span className="format-divider" aria-hidden="true" /><button type="button" className={`format-button ${allSelectedHaveAlignment('left') ? 'is-active' : ''}`} aria-pressed={allSelectedHaveAlignment('left')} onMouseDown={(event) => event.preventDefault()} onClick={() => setAlignment('left')} aria-label="Align left" title="Align left"><svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" fill="none"><line x1="1" y1="3" x2="15" y2="3" /><line x1="1" y1="8" x2="10" y2="8" /><line x1="1" y1="13" x2="13" y2="13" /></svg></button><button type="button" className={`format-button ${allSelectedHaveAlignment('center') ? 'is-active' : ''}`} aria-pressed={allSelectedHaveAlignment('center')} onMouseDown={(event) => event.preventDefault()} onClick={() => setAlignment('center')} aria-label="Align center" title="Align center"><svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" fill="none"><line x1="1" y1="3" x2="15" y2="3" /><line x1="3" y1="8" x2="13" y2="8" /><line x1="2" y1="13" x2="14" y2="13" /></svg></button><button type="button" className={`format-button ${allSelectedHaveAlignment('right') ? 'is-active' : ''}`} aria-pressed={allSelectedHaveAlignment('right')} onMouseDown={(event) => event.preventDefault()} onClick={() => setAlignment('right')} aria-label="Align right" title="Align right"><svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" fill="none"><line x1="1" y1="3" x2="15" y2="3" /><line x1="6" y1="8" x2="15" y2="8" /><line x1="3" y1="13" x2="15" y2="13" /></svg></button></div></section><section className="grid-card" aria-labelledby="grid-title"><h3 id="grid-title" className="sr-only">{sheet.name} spreadsheet data</h3><div className="table-scroll" ref={tableScrollRef}><table className="spreadsheet"><caption className="sr-only">CSV-backed data in {sheet.name}. Press Enter, Space, or F2 to edit a cell. Right-click a row or column header for actions, or drag a column's edge to resize it.</caption><colgroup><col style={{ width: '3rem' }} />{sheet.columns.map((_, index) => <col key={index} ref={(element) => { if (element) colRefs.current.set(index, element); else colRefs.current.delete(index) }} style={{ width: `${sheet.columnWidths?.[index] || DEFAULT_COLUMN_WIDTH}px` }} />)}</colgroup><thead><tr><th scope="col" className="corner-cell" aria-label="Spreadsheet corner" />{sheet.columns.map((column, index) => <th scope="col" key={`${column}-${index}`} onContextMenu={(event) => openContextMenu(event, 'column', index)}><button type="button" className={`column-header-button ${selectedColumns.has(index) ? 'is-active' : ''}`} onClick={(event) => selectColumn(index, event.metaKey || event.ctrlKey)} aria-label={`Select column ${columnLabel(index)}`}>{columnLabel(index)}</button><span className="column-resize-handle" onMouseDown={(event) => startColumnResize(event, index)} onClick={(event) => event.stopPropagation()} aria-hidden="true" /></th>)}</tr></thead><tbody>{sheet.rows.map((row, rowIndex) => <tr key={`${sheet.id}-${rowIndex}`}><th scope="row" onContextMenu={(event) => openContextMenu(event, 'row', rowIndex)}><button type="button" className={`row-header-button ${selectedRows.has(rowIndex) ? 'is-active' : ''}`} onClick={(event) => selectRow(rowIndex, event.metaKey || event.ctrlKey)} aria-label={`Select row ${rowIndex + 1}`}>{rowIndex + 1}</button></th>{sheet.columns.map((_, columnIndex) => { const coordinate = cellCoordinate(columnIndex, rowIndex); const value = row[columnIndex] || ''; const metadata = sheet.cells?.[coordinate]; const isEditing = editingCell === coordinate; return <td key={coordinate}>{isEditing ? <input ref={editorRef} autoFocus type="text" className="cell-editor" style={cellStyle(value, metadata, workbook.styles)} value={draftValue} onChange={(event) => setDraftValue(event.currentTarget.value)} onInput={(event) => setDraftValue(event.currentTarget.value)} onBlur={commitEdit} onKeyDown={(event) => { if (event.metaKey || event.ctrlKey) { const key = event.key.toLowerCase(); if (key === 'b' || key === 'i' || key === 'u') { event.preventDefault(); event.stopPropagation(); toggleFont(key === 'b' ? 'bold' : key === 'i' ? 'italic' : 'underline'); return } } event.stopPropagation(); if (event.key === 'Enter') { event.preventDefault(); commitEdit(); return } if (event.key === 'Escape') { event.preventDefault(); cancelEdit(); return } if (event.key === 'ArrowUp' || event.key === 'ArrowDown') { event.preventDefault(); commitEditAndMove(ARROW_DIRECTIONS[event.key]); return } const input = event.currentTarget; const atStart = input.selectionStart === 0 && input.selectionEnd === 0; const atEnd = input.selectionStart === input.value.length && input.selectionEnd === input.value.length; if (event.key === 'ArrowLeft' && atStart) { event.preventDefault(); commitEditAndMove(ARROW_DIRECTIONS.ArrowLeft); return } if (event.key === 'ArrowRight' && atEnd) { event.preventDefault(); commitEditAndMove(ARROW_DIRECTIONS.ArrowRight) } }} onMouseDown={(event) => event.stopPropagation()} onClick={(event) => event.stopPropagation()} aria-label={`Edit ${coordinate}`} /> : <button type="button" ref={(element) => { if (element) cellRefs.current.set(coordinate, element); else cellRefs.current.delete(coordinate) }} className={`cell-button ${selectedCell === coordinate ? 'is-selected' : ''} ${selectedCell !== coordinate && selectedCells.has(coordinate) ? 'is-in-selection' : ''}`} style={cellStyle(value, metadata, workbook.styles)} onClick={(event) => { if (event.metaKey || event.ctrlKey) toggleCellSelection(coordinate); else selectCell(coordinate) }} onKeyDown={(event) => handleCellKeyDown(event, coordinate)} aria-label={`${coordinate}, value ${value || 'blank'}`}>{formatCellValue(value, metadata, workbook.styles)}{metadata?.formula ? <span className="formula-indicator" aria-label="Formula"> ƒ</span> : null}</button>}</td> })}</tr>)}</tbody></table></div></section></main><footer className="app-footer"><nav className="sheet-tabs" aria-label="Workbook sheets"><span className="sheet-tabs-label">Sheets<span className="count">{workbook.sheets.length}</span></span><div className="sheet-tabs-list">{workbook.sheets.map((item) => renamingSheetId === item.id ? <input key={item.id} ref={sheetNameInputRef} type="text" className="sheet-tab-input" value={sheetNameDraft} onChange={(event) => setSheetNameDraft(event.target.value)} onBlur={commitRenameSheet} onClick={(event) => event.stopPropagation()} onKeyDown={(event) => { event.stopPropagation(); if (event.key === 'Enter') { event.preventDefault(); commitRenameSheet() } if (event.key === 'Escape') { event.preventDefault(); cancelRenameSheet() } }} aria-label={`Rename sheet ${item.name}`} /> : <button key={item.id} type="button" className={`sheet-tab ${item.id === activeSheet ? 'is-active' : ''}`} onClick={() => switchSheet(item.id)} onDoubleClick={() => beginRenameSheet(item.id, item.name)} onContextMenu={(event) => openSheetContextMenu(event, item.id)} aria-current={item.id === activeSheet ? 'true' : undefined}>{item.name}</button>)}<button type="button" className="sheet-tab-add" onClick={addSheet} aria-label="Add sheet" title="Add sheet">+</button></div></nav><div className="status-bar" role="status"><strong>{sheet.rows.length}</strong>&nbsp;rows · <strong>{sheet.columns.length}</strong>&nbsp;columns</div><span className="footer-hint">Enter/F2 edit · Right-click row header for actions</span></footer>{contextMenu &&<div ref={contextMenuRef} className="row-context-menu" role="menu" style={{ left: contextMenu.x, top: contextMenu.y }} onClick={(event) => event.stopPropagation()}>{contextMenu.kind === 'row' ? <><p className="context-menu-label">Row {contextMenu.index + 1}</p><button type="button" role="menuitem" onClick={() => addRow(contextMenu.index)}>Insert row before</button><button type="button" role="menuitem" onClick={() => addRow(contextMenu.index + 1)}>Insert row after</button><button type="button" role="menuitem" className="danger-action" onClick={() => deleteRow(contextMenu.index)}>Delete row</button></> : contextMenu.kind === 'column' ? <><p className="context-menu-label">Column {columnLabel(contextMenu.index)}</p><button type="button" role="menuitem" onClick={() => addColumn(contextMenu.index)}>Insert column before</button><button type="button" role="menuitem" onClick={() => addColumn(contextMenu.index + 1)}>Insert column after</button><button type="button" role="menuitem" className="danger-action" onClick={() => deleteColumn(contextMenu.index)}>Delete column</button></> : <><p className="context-menu-label">{workbook.sheets.find((item) => item.id === contextMenu.index)?.name}</p><button type="button" role="menuitem" onClick={() => beginRenameSheet(contextMenu.index, workbook.sheets.find((item) => item.id === contextMenu.index)?.name || '')}>Rename sheet</button><button type="button" role="menuitem" className="danger-action" onClick={() => removeSheet(contextMenu.index)}>Delete sheet</button></>}</div>}</div>
}
export default App
