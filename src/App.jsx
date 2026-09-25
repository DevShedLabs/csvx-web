import { useEffect, useMemo, useRef, useState } from 'react'

const demoWorkbook = { name: 'example.csvx', version: '1.0', sheets: [{ id: 'sheet-1', name: 'Sheet 1', columns: ['Value'], rows: [['1'], ['2']], cells: {} }], styles: {}, source: null }
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
    const name = decode(bytes.subarray(offset + 46, offset + 46 + nameLength)); const localOffset = readUInt32(bytes, offset + 42)
    const localNameLength = readUInt16(bytes, localOffset + 26); const localExtraLength = readUInt16(bytes, localOffset + 28); const start = localOffset + 30 + localNameLength + localExtraLength
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

function cellStyle(metadata, styles) { const style = metadata?.style ? styles?.[metadata.style] : null; const font = style?.font || {}; const fill = style?.fill || {}; const alignment = style?.alignment || {}; const numberFormat = style?.numberFormat || ''; const numericFormat = /[$€£%]|0[.,#]/.test(numberFormat); return { color: font.color || undefined, backgroundColor: fill.color || undefined, fontFamily: font.name ? `'${font.name}', ui-monospace, monospace` : undefined, fontSize: font.size ? `${font.size}pt` : undefined, fontWeight: font.bold ? 700 : undefined, fontStyle: font.italic ? 'italic' : undefined, textAlign: alignment.horizontal || (numericFormat ? 'right' : undefined), verticalAlign: alignment.vertical || undefined, whiteSpace: alignment.wrapText ? 'normal' : 'nowrap', transform: alignment.textRotation && alignment.textRotation !== '0' ? `rotate(${alignment.textRotation}deg)` : undefined, paddingLeft: alignment.indent ? `${Number(alignment.indent) * 0.5}rem` : undefined } }
function columnLabel(index) { let label = ''; let value = index + 1; while (value > 0) { const remainder = (value - 1) % 26; label = String.fromCharCode(65 + remainder) + label; value = Math.floor((value - 1) / 26) } return label }
function cellCoordinate(column, row) { return `${columnLabel(column)}${row + 1}` }
function cellPosition(coordinate) { const match = coordinate.match(/^([A-Z]+)(\d+)$/); if (!match) return { column: 0, row: 0 }; const column = match[1].split('').reduce((total, character) => total * 26 + character.charCodeAt(0) - 64, 0) - 1; return { column, row: Number(match[2]) - 1 } }

function App() {
  const [workbook, setWorkbook] = useState(demoWorkbook); const [activeSheet, setActiveSheet] = useState('sheet-1'); const [selectedCell, setSelectedCell] = useState('A1'); const [error, setError] = useState(''); const inputRef = useRef(null)
  const sheet = useMemo(() => workbook.sheets.find((item) => item.id === activeSheet) || workbook.sheets[0], [activeSheet, workbook]); const selectedMetadata = sheet?.cells?.[selectedCell]; const position = cellPosition(selectedCell); const selectedValue = sheet?.rows?.[position.row]?.[position.column] || ''; const selectedDisplayValue = selectedMetadata?.formula || formatCellValue(selectedValue, selectedMetadata, workbook.styles) || 'Blank cell'; const selectedType = selectedMetadata?.type && selectedDisplayValue !== 'Blank cell' ? selectedMetadata.type : ''
  useEffect(() => { const stored = localStorage.getItem(STORED_PACKAGE_KEY); if (!stored) return; try { const restored = JSON.parse(stored); if (restored?.name && Array.isArray(restored.sheets)) { setWorkbook(restored); setActiveSheet(restored.sheets[0]?.id) } } catch { localStorage.removeItem(STORED_PACKAGE_KEY) } }, [])
  function moveSelection(direction) { const column = Math.max(0, Math.min(sheet.columns.length - 1, position.column + direction.column)); const row = Math.max(0, Math.min(sheet.rows.length - 1, position.row + direction.row)); setSelectedCell(cellCoordinate(column, row)) }
  function handleCellKeyDown(event) { const directions = { ArrowLeft: { column: -1, row: 0 }, ArrowRight: { column: 1, row: 0 }, ArrowUp: { column: 0, row: -1 }, ArrowDown: { column: 0, row: 1 } }; if (!directions[event.key]) return; event.preventDefault(); moveSelection(directions[event.key]) }
  async function handleOpen(event) { const file = event.target.files?.[0]; if (!file) return; try { const loaded = await parseCSVX(file); setWorkbook(loaded); setActiveSheet(loaded.sheets[0]?.id); setSelectedCell('A1'); setError(''); localStorage.setItem(STORED_PACKAGE_KEY, JSON.stringify(loaded)) } catch (loadError) { setError(loadError.message); setWorkbook(demoWorkbook) } event.target.value = '' }
  return <div className="app-shell"><header className="topbar"><div className="brand-lockup"><span className="brand-mark" aria-hidden="true">X</span><div><p className="eyebrow">CSVX demo</p><h1>Workbench</h1></div></div><div className="file-status" role="status"><span className="status-dot" aria-hidden="true" /><span>{workbook.name}</span><span className="muted">Core {workbook.version}</span></div><nav className="top-actions" aria-label="File actions"><input ref={inputRef} type="file" accept=".csvx,application/zip" onChange={handleOpen} className="sr-only" aria-label="Open CSVX package" /><button type="button" className="button button-quiet" onClick={() => inputRef.current?.click()}>Open</button><button type="button" className="button button-primary" disabled aria-disabled="true">Export</button></nav></header><div className="workspace"><aside className="sidebar" aria-label="Workbook navigation"><section className="sidebar-section"><div className="section-heading"><h2>Workbook</h2><span className="count">{workbook.sheets.length}</span></div><ul className="sheet-list">{workbook.sheets.map((item) => <li key={item.id}><button type="button" className={`sheet-item ${item.id === activeSheet ? 'is-active' : ''}`} onClick={() => { setActiveSheet(item.id); setSelectedCell('A1') }}><span className="sheet-icon" aria-hidden="true">▦</span>{item.name}</button></li>)}</ul></section></aside><main id="main-content" className="main-content" tabIndex="-1">{error && <p role="alert">Unable to open package: {error}</p>}<div className="content-header"><div><p className="eyebrow">Sheet / {sheet.name}</p><h2>{sheet.name}</h2></div><span className="read-only-badge">Read-only demo</span></div><section className="formula-panel" aria-label="Cell inspector"><div className="name-box">{selectedCell}</div><div className="formula-symbol" aria-hidden="true">fx</div><div className="formula-value">{selectedDisplayValue}{selectedType ? <span className="value-type">{selectedType}</span> : null}</div></section><section className="grid-card" aria-labelledby="grid-title"><h3 id="grid-title" className="sr-only">{sheet.name} spreadsheet data</h3><div className="table-scroll"><table className="spreadsheet"><caption className="sr-only">CSV-backed data in {sheet.name}</caption><thead><tr><th scope="col" className="corner-cell" aria-label="Spreadsheet corner" />{sheet.columns.map((column, index) => <th scope="col" key={`${column}-${index}`} style={{ width: sheet.columns[index]?.width ? `${sheet.columns[index].width}ch` : undefined }}>{columnLabel(index)}</th>)}</tr></thead><tbody>{sheet.rows.map((row, rowIndex) => <tr key={`${sheet.id}-${rowIndex}`} style={{ height: sheet.rowHeights?.[rowIndex + 1] ? `${sheet.rowHeights[rowIndex + 1]}pt` : undefined }}><th scope="row">{rowIndex + 1}</th>{sheet.columns.map((_, columnIndex) => { const coordinate = cellCoordinate(columnIndex, rowIndex); const value = row[columnIndex] || ''; const metadata = sheet.cells?.[coordinate]; return <td key={coordinate}><button type="button" style={cellStyle(metadata, workbook.styles)} className={`cell-button ${selectedCell === coordinate ? 'is-selected' : ''}`} onClick={() => setSelectedCell(coordinate)} onKeyDown={handleCellKeyDown} aria-label={`${coordinate}, value ${value || 'blank'}`}>{formatCellValue(value, metadata, workbook.styles)}{metadata?.formula ? <span className="formula-indicator" aria-label="Formula"> ƒ</span> : null}</button></td> })}</tr>)}</tbody></table></div></section></main></div></div>
}
export default App
