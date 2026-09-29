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
  const numericValue = Number(value); const decimalPart = format.split('.')[1] || ''; const decimals = (decimalPart.match(/0/g) || []).length; const outputValue = format.includes('%') ? numericValue * 100 : numericValue; const rounded = outputValue.toFixed(decimals); const [whole, fraction] = rounded.split('.'); const grouped = format.includes(',') ? whole.replace(/\B(?=(\d{3})+(?!\d))/g, ',') : whole; const number = fraction ? `${grouped}.${fraction}` : grouped
  if (format.includes('%')) return `${number}%`; if (format.includes('$')) return `$${number}`; if (format.includes('€')) return `€${number}`; if (format.includes('£')) return `£${number}`; return number
}
function cellStyle(value, metadata, styles) { const style = metadata?.style ? styles?.[metadata.style] : null; const font = style?.font || {}; const fill = style?.fill || {}; const alignment = style?.alignment || {}; const numericType = ['integer', 'decimal'].includes(metadata?.type); const numericValue = value !== '' && !Number.isNaN(Number(value)); return { color: font.color || undefined, backgroundColor: fill.color || undefined, fontWeight: font.bold ? 700 : undefined, fontStyle: font.italic ? 'italic' : undefined, textAlign: alignment.horizontal || ((numericType || numericValue) ? 'right' : undefined), whiteSpace: alignment.wrapText ? 'normal' : 'nowrap' } }
function columnLabel(index) { let label = ''; let value = index + 1; while (value > 0) { const remainder = (value - 1) % 26; label = String.fromCharCode(65 + remainder) + label; value = Math.floor((value - 1) / 26) } return label }
function cellCoordinate(column, row) { return `${columnLabel(column)}${row + 1}` }
function cellPosition(coordinate) { const match = coordinate.match(/^([A-Z]+)(\d+)$/); if (!match) return { column: 0, row: 0 }; const column = match[1].split('').reduce((total, character) => total * 26 + character.charCodeAt(0) - 64, 0) - 1; return { column, row: Number(match[2]) - 1 } }
function cloneSheet(sheet) { return { ...sheet, rows: sheet.rows.map((row) => [...row]), cells: { ...sheet.cells } } }
function getCellText(sheet, coordinate) { const target = cellPosition(coordinate); return sheet.rows[target.row]?.[target.column] ?? '' }
function setCellText(workbook, sheetId, coordinate, value) { const target = cellPosition(coordinate); return { ...workbook, sheets: workbook.sheets.map((item) => { if (item.id !== sheetId) return item; const nextSheet = cloneSheet(item); while (nextSheet.rows.length <= target.row) nextSheet.rows.push(Array(nextSheet.columns.length).fill('')); while (nextSheet.rows[target.row].length < nextSheet.columns.length) nextSheet.rows[target.row].push(''); nextSheet.rows[target.row][target.column] = value; return nextSheet }) } }
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

function App() {
  const [workbook, setWorkbook] = useState(demoWorkbook); const [activeSheet, setActiveSheet] = useState('sheet-1'); const [selectedCell, setSelectedCell] = useState('A1'); const [editingCell, setEditingCell] = useState(null); const [draftValue, setDraftValue] = useState(''); const [contextMenu, setContextMenu] = useState(null); const [error, setError] = useState(''); const inputRef = useRef(null); const editorRef = useRef(null)
  const sheet = useMemo(() => workbook.sheets.find((item) => item.id === activeSheet) || workbook.sheets[0], [activeSheet, workbook]); const selectedMetadata = sheet?.cells?.[selectedCell]; const position = cellPosition(selectedCell); const selectedValue = sheet?.rows?.[position.row]?.[position.column] || ''; const selectedDisplayValue = selectedMetadata?.formula || formatCellValue(selectedValue, selectedMetadata, workbook.styles) || 'Blank cell'; const selectedType = selectedMetadata?.type && selectedDisplayValue !== 'Blank cell' ? selectedMetadata.type : ''
  useEffect(() => { const stored = localStorage.getItem(STORED_PACKAGE_KEY); if (!stored) return; try { const restored = JSON.parse(stored); if (restored?.name && Array.isArray(restored.sheets)) { setWorkbook(restored); setActiveSheet(restored.sheets[0]?.id) } } catch { localStorage.removeItem(STORED_PACKAGE_KEY) } }, [])
  useEffect(() => { if (editingCell) editorRef.current?.focus() }, [editingCell])
  useEffect(() => { const closeMenu = () => setContextMenu(null); document.addEventListener('click', closeMenu); return () => document.removeEventListener('click', closeMenu) }, [])
  function updateWorkbook(nextWorkbook) { setWorkbook(nextWorkbook); localStorage.setItem(STORED_PACKAGE_KEY, JSON.stringify(nextWorkbook)) }
  function selectCell(coordinate) { setSelectedCell(coordinate); setContextMenu(null) }
  function beginEdit(coordinate) { setSelectedCell(coordinate); setDraftValue(getCellText(sheet, coordinate)); setEditingCell(coordinate); setContextMenu(null) }
  function commitEdit() { if (!editingCell) return; const value = editorRef.current?.value ?? draftValue; updateWorkbook(setCellText(workbook, sheet.id, editingCell, value)); setEditingCell(null) }
  function cancelEdit() { setEditingCell(null) }
  function addRow(index) { const next = remapRows(workbook, sheet, index, 1, false); updateWorkbook(next); setSelectedCell(cellCoordinate(0, index)); setContextMenu(null) }
  function deleteRow(index) { if (sheet.rows.length <= 1) { setError('A sheet must contain at least one data row.'); setContextMenu(null); return } const next = remapRows(workbook, sheet, index, -1, true); updateWorkbook(next); setSelectedCell(cellCoordinate(0, Math.max(0, index - 1))); setContextMenu(null) }
  function openRowMenu(event, rowIndex) { event.preventDefault(); setSelectedCell(cellCoordinate(0, rowIndex)); setContextMenu({ x: event.clientX, y: event.clientY, rowIndex }) }
  function handleCellKeyDown(event, coordinate) { if (editingCell === coordinate) { if (event.key === 'Enter') { event.preventDefault(); commitEdit() } if (event.key === 'Escape') { event.preventDefault(); cancelEdit() } return } if (event.key === 'Enter' || event.key === 'F2' || event.key === ' ') { event.preventDefault(); beginEdit(coordinate); return } const directions = { ArrowLeft: { column: -1, row: 0 }, ArrowRight: { column: 1, row: 0 }, ArrowUp: { column: 0, row: -1 }, ArrowDown: { column: 0, row: 1 } }; if (!directions[event.key]) return; event.preventDefault(); const nextColumn = Math.max(0, Math.min(sheet.columns.length - 1, position.column + directions[event.key].column)); const nextRow = Math.max(0, Math.min(sheet.rows.length - 1, position.row + directions[event.key].row)); selectCell(cellCoordinate(nextColumn, nextRow)) }
  async function handleOpen(event) { const file = event.target.files?.[0]; if (!file) return; try { const loaded = await parseCSVX(file); setWorkbook(loaded); setActiveSheet(loaded.sheets[0]?.id); setSelectedCell('A1'); setError(''); localStorage.setItem(STORED_PACKAGE_KEY, JSON.stringify(loaded)) } catch (loadError) { setError(loadError.message); setWorkbook(demoWorkbook) } event.target.value = '' }
  return <div className="app-shell"><header className="topbar"><div className="brand-lockup"><span className="brand-mark" aria-hidden="true">X</span><div><p className="eyebrow">CSVX demo</p><h1>Workbench</h1></div></div><div className="file-status" role="status"><span className="status-dot" aria-hidden="true" /><span>{workbook.name}</span><span className="muted">Core {workbook.version}</span></div><nav className="top-actions" aria-label="File actions"><input ref={inputRef} type="file" accept=".csvx,application/zip" onChange={handleOpen} className="sr-only" aria-label="Open CSVX package" /><button type="button" className="button button-quiet" onClick={() => inputRef.current?.click()}>Open</button><button type="button" className="button button-primary" disabled aria-disabled="true">Export</button></nav></header><div className="workspace"><aside className="sidebar" aria-label="Workbook navigation"><section className="sidebar-section"><div className="section-heading"><h2>Workbook</h2><span className="count">{workbook.sheets.length}</span></div><ul className="sheet-list">{workbook.sheets.map((item) => <li key={item.id}><button type="button" className={`sheet-item ${item.id === activeSheet ? 'is-active' : ''}`} onClick={() => { setActiveSheet(item.id); setSelectedCell('A1') }}><span className="sheet-icon" aria-hidden="true">▦</span>{item.name}</button></li>)}</ul></section></aside><main id="main-content" className="main-content" tabIndex="-1">{error && <p role="alert">{error}</p>}<div className="content-header"><div><p className="eyebrow">Sheet / {sheet.name}</p><h2>{sheet.name}</h2></div><span className="read-only-badge">Local edits</span></div><section className="formula-panel" aria-label="Cell inspector"><div className="name-box">{selectedCell}</div><div className="formula-symbol" aria-hidden="true">fx</div><div className="formula-value">{selectedDisplayValue}{selectedType ? <span className="value-type">{selectedType}</span> : null}</div></section><section className="grid-card" aria-labelledby="grid-title"><h3 id="grid-title" className="sr-only">{sheet.name} spreadsheet data</h3><div className="table-scroll"><table className="spreadsheet"><caption className="sr-only">CSV-backed data in {sheet.name}. Press Enter, Space, or F2 to edit a cell. Right-click a row header for row actions.</caption><thead><tr><th scope="col" className="corner-cell" aria-label="Spreadsheet corner" />{sheet.columns.map((column, index) => <th scope="col" key={`${column}-${index}`}>{columnLabel(index)}</th>)}</tr></thead><tbody>{sheet.rows.map((row, rowIndex) => <tr key={`${sheet.id}-${rowIndex}`}><th scope="row" onContextMenu={(event) => openRowMenu(event, rowIndex)}><button type="button" className="row-header-button" onClick={() => selectCell(cellCoordinate(0, rowIndex))} aria-label={`Select row ${rowIndex + 1}`}>{rowIndex + 1}</button></th>{sheet.columns.map((_, columnIndex) => { const coordinate = cellCoordinate(columnIndex, rowIndex); const value = row[columnIndex] || ''; const metadata = sheet.cells?.[coordinate]; const isEditing = editingCell === coordinate; return <td key={coordinate} style={cellStyle(value, metadata, workbook.styles)}>{isEditing ? <input ref={editorRef} autoFocus type="text" className="cell-editor" value={draftValue} onChange={(event) => setDraftValue(event.currentTarget.value)} onInput={(event) => setDraftValue(event.currentTarget.value)} onBlur={commitEdit} onKeyDown={(event) => { event.stopPropagation(); if (event.key === 'Enter') { event.preventDefault(); commitEdit() } if (event.key === 'Escape') { event.preventDefault(); cancelEdit() } }} onMouseDown={(event) => event.stopPropagation()} onClick={(event) => event.stopPropagation()} aria-label={`Edit ${coordinate}`} /> : <button type="button" className={`cell-button ${selectedCell === coordinate ? 'is-selected' : ''}`} onClick={() => selectCell(coordinate)} onKeyDown={(event) => handleCellKeyDown(event, coordinate)} aria-label={`${coordinate}, value ${value || 'blank'}`}>{formatCellValue(value, metadata, workbook.styles)}{metadata?.formula ? <span className="formula-indicator" aria-label="Formula"> ƒ</span> : null}</button>}</td> })}</tr>)}</tbody></table></div></section><div className="status-bar" role="status"><span><strong>{sheet.rows.length}</strong> data rows</span><span><strong>{sheet.columns.length}</strong> columns</span><span className="status-spacer" /><span>Enter/F2 edit · Right-click row header for actions</span></div></main></div>{contextMenu && <div className="row-context-menu" role="menu" style={{ left: contextMenu.x, top: contextMenu.y }} onClick={(event) => event.stopPropagation()}><p className="context-menu-label">Row {contextMenu.rowIndex + 1}</p><button type="button" role="menuitem" onClick={() => addRow(contextMenu.rowIndex)}>Insert row before</button><button type="button" role="menuitem" onClick={() => addRow(contextMenu.rowIndex + 1)}>Insert row after</button><button type="button" role="menuitem" className="danger-action" onClick={() => deleteRow(contextMenu.rowIndex)}>Delete row</button></div>}</div>
}
export default App
