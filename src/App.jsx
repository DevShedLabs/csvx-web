import { useEffect, useMemo, useRef, useState } from 'react'
import { loadWorkbookFromZip, validateBuffer, writeWorkbookToZip } from 'csvx-ts/browser'
import { downloadBytes } from './download.js'
import {
  addSheet,
  appendColumn,
  appendRow,
  applyCellsFormat,
  cellCSS,
  cellMetadata,
  clearCellsFormat,
  coordinateFor,
  deleteSheet,
  findSheet,
  renameSheet,
  setCellValue,
  styleForId,
} from './model.js'

const DEMO_URL = '/example.csvx'
const DEFAULT_COLUMN_WIDTH = 128
const MIN_COLUMN_WIDTH = 48
const ARROW_DIRECTIONS = { ArrowLeft: { column: -1, row: 0 }, ArrowRight: { column: 1, row: 0 }, ArrowUp: { column: 0, row: -1 }, ArrowDown: { column: 0, row: 1 } }

async function loadWorkbookFromBuffer(buffer) {
  const diagnostics = await validateBuffer(buffer)
  if (!diagnostics.valid) {
    const message = diagnostics.errors[0]?.message || 'Package failed to load'
    throw new Error(message)
  }
  return loadWorkbookFromZip(buffer)
}

function App() {
  const [workbook, setWorkbook] = useState(null)
  const [fileName, setFileName] = useState('')
  const [activeSheetId, setActiveSheetId] = useState(null)
  const [selectedCell, setSelectedCell] = useState({ row: 0, column: 0 })
  const [selectedCells, setSelectedCells] = useState(() => new Set(['A1']))
  const [selectedColumns, setSelectedColumns] = useState(() => new Set())
  const [selectedRows, setSelectedRows] = useState(() => new Set())
  const [editingCell, setEditingCell] = useState(null)
  const [draftValue, setDraftValue] = useState('')
  const [contextMenu, setContextMenu] = useState(null)
  const [error, setError] = useState('')
  const [renamingSheetId, setRenamingSheetId] = useState(null)
  const [sheetNameDraft, setSheetNameDraft] = useState('')
  // View-only, never written back into the workbook — see the comment on Column.width in
  // model.js for why: the schema's declared width is in an undefined (and, for real imported
  // data, non-pixel) unit, so resizing here must not overwrite it.
  const [columnWidthOverrides, setColumnWidthOverrides] = useState({})
  const inputRef = useRef(null)
  const editorRef = useRef(null)
  const cellRefs = useRef(new Map())
  const colRefs = useRef(new Map())
  const contextMenuRef = useRef(null)
  const tableScrollRef = useRef(null)
  const sheetNameInputRef = useRef(null)

  const sheet = useMemo(() => (workbook ? findSheet(workbook, activeSheetId) : null), [workbook, activeSheetId])
  const selectedCoordinate = coordinateFor(selectedCell.column, selectedCell.row)
  const selectedMetadata = cellMetadata(sheet, selectedCoordinate)
  const selectedRawValue = sheet?.records?.[selectedCell.row]?.[selectedCell.column] ?? ''
  const selectedStyle = styleForId(workbook?.styles, selectedMetadata?.style)
  const selectedAlignment = selectedStyle.alignment?.horizontal || ''
  const selectedType = selectedMetadata?.type || sheet?.columns?.[selectedCell.column]?.type || ''
  const selectionStyles = useMemo(
    () => [...selectedCells].map((coordinate) => styleForId(workbook?.styles, cellMetadata(sheet, coordinate)?.style)),
    [selectedCells, sheet, workbook],
  )
  const allSelectedHaveFont = (property) => selectionStyles.length > 0 && selectionStyles.every((style) => style.font?.[property])
  const allSelectedHaveAlignment = (value) => selectionStyles.length > 0 && selectionStyles.every((style) => (style.alignment?.horizontal || '') === value)

  // Load the real example fixture on first mount, through the real engine, instead of shipping
  // fabricated demo data — see AGENTS.md: this app must never hold a second opinion about what a
  // CSVX file contains.
  useEffect(() => {
    let cancelled = false
    fetch(DEMO_URL)
      .then((response) => response.arrayBuffer())
      .then(async (buffer) => {
        const loaded = await loadWorkbookFromBuffer(buffer)
        if (cancelled) return
        setWorkbook(loaded)
        setFileName('example.csvx')
        setActiveSheetId(loaded.sheets[0]?.id)
      })
      .catch((loadError) => {
        if (!cancelled) setError(`Could not load demo package: ${loadError.message}`)
      })
    return () => {
      cancelled = true
    }
  }, [])

  useEffect(() => {
    if (editingCell) {
      const input = editorRef.current
      input?.focus()
      input?.setSelectionRange(input.value.length, input.value.length)
    }
  }, [editingCell])
  useEffect(() => {
    if (renamingSheetId) {
      sheetNameInputRef.current?.focus()
      sheetNameInputRef.current?.select()
    }
  }, [renamingSheetId])
  useEffect(() => {
    if (!editingCell) cellRefs.current.get(selectedCoordinate)?.focus()
  }, [selectedCoordinate, editingCell, activeSheetId])
  useEffect(() => {
    const closeMenu = () => setContextMenu(null)
    document.addEventListener('click', closeMenu)
    return () => document.removeEventListener('click', closeMenu)
  }, [])
  useEffect(() => {
    const el = tableScrollRef.current
    if (!el) return
    let axis = null
    let resetTimer = null
    function handleWheel(event) {
      if (resetTimer) clearTimeout(resetTimer)
      resetTimer = setTimeout(() => {
        axis = null
      }, 200)
      if (!axis) axis = Math.abs(event.deltaX) > Math.abs(event.deltaY) ? 'x' : 'y'
      event.preventDefault()
      if (axis === 'x') el.scrollLeft += event.deltaX
      else el.scrollTop += event.deltaY
    }
    el.addEventListener('wheel', handleWheel, { passive: false })
    return () => {
      el.removeEventListener('wheel', handleWheel)
      if (resetTimer) clearTimeout(resetTimer)
    }
  }, [])
  useEffect(() => {
    function handleShortcut(event) {
      if (!(event.metaKey || event.ctrlKey)) return
      const key = event.key.toLowerCase()
      if (key === 'b') {
        event.preventDefault()
        toggleFont('bold')
      } else if (key === 'i') {
        event.preventDefault()
        toggleFont('italic')
      } else if (key === 'u') {
        event.preventDefault()
        toggleFont('underline')
      }
    }
    document.addEventListener('keydown', handleShortcut)
    return () => document.removeEventListener('keydown', handleShortcut)
  }, [workbook, sheet, selectedCoordinate])
  useEffect(() => {
    if (!contextMenu || !contextMenuRef.current) return
    const rect = contextMenuRef.current.getBoundingClientRect()
    const margin = 8
    const maxLeft = window.innerWidth - rect.width - margin
    const maxTop = window.innerHeight - rect.height - margin
    const clampedLeft = Math.min(contextMenu.x, Math.max(margin, maxLeft))
    const clampedTop = Math.min(contextMenu.y, Math.max(margin, maxTop))
    if (clampedLeft !== contextMenu.x || clampedTop !== contextMenu.y) {
      setContextMenu((current) => (current ? { ...current, x: clampedLeft, y: clampedTop } : current))
    }
  }, [contextMenu])

  function selectCell(row, column) {
    setSelectedCell({ row, column })
    setSelectedCells(new Set([coordinateFor(column, row)]))
    setSelectedColumns(new Set())
    setSelectedRows(new Set())
    setContextMenu(null)
  }
  function toggleCellSelection(row, column) {
    const coordinate = coordinateFor(column, row)
    setSelectedCell({ row, column })
    setSelectedCells((current) => {
      const next = new Set(current)
      if (next.has(coordinate)) next.delete(coordinate)
      else next.add(coordinate)
      return next.size ? next : new Set([coordinate])
    })
    setContextMenu(null)
  }
  function selectColumn(columnIndex, additive) {
    const coordinates = sheet.records.map((_, rowIndex) => coordinateFor(columnIndex, rowIndex))
    setSelectedCell({ row: 0, column: columnIndex })
    if (additive) {
      setSelectedCells((current) => new Set([...current, ...coordinates]))
      setSelectedColumns((current) => new Set(current).add(columnIndex))
    } else {
      setSelectedCells(new Set(coordinates))
      setSelectedColumns(new Set([columnIndex]))
      setSelectedRows(new Set())
    }
    setContextMenu(null)
  }
  function selectRow(rowIndex, additive) {
    const coordinates = sheet.columns.map((_, columnIndex) => coordinateFor(columnIndex, rowIndex))
    setSelectedCell({ row: rowIndex, column: 0 })
    if (additive) {
      setSelectedCells((current) => new Set([...current, ...coordinates]))
      setSelectedRows((current) => new Set(current).add(rowIndex))
    } else {
      setSelectedCells(new Set(coordinates))
      setSelectedRows(new Set([rowIndex]))
      setSelectedColumns(new Set())
    }
    setContextMenu(null)
  }
  function switchSheet(sheetId) {
    setActiveSheetId(sheetId)
    selectCell(0, 0)
  }

  function beginEdit(row, column, initialValue) {
    const coordinate = coordinateFor(column, row)
    setSelectedCell({ row, column })
    setDraftValue(initialValue !== undefined ? initialValue : sheet.records[row]?.[column] ?? '')
    setEditingCell(coordinate)
    setContextMenu(null)
  }
  function commitEdit() {
    if (!editingCell) return
    const value = editorRef.current?.value ?? draftValue
    setWorkbook(setCellValue(workbook, sheet.id, selectedCell.row, selectedCell.column, value))
    setEditingCell(null)
  }
  function cancelEdit() {
    setEditingCell(null)
  }
  function moveSelection(direction) {
    const nextColumn = Math.max(0, Math.min(sheet.columns.length - 1, selectedCell.column + direction.column))
    const nextRow = Math.max(0, Math.min(sheet.records.length - 1, selectedCell.row + direction.row))
    selectCell(nextRow, nextColumn)
  }
  function commitEditAndMove(direction) {
    if (!editingCell) return
    const value = editorRef.current?.value ?? draftValue
    setWorkbook(setCellValue(workbook, sheet.id, selectedCell.row, selectedCell.column, value))
    setEditingCell(null)
    moveSelection(direction)
  }
  function handleCellKeyDown(event, row, column) {
    const coordinate = coordinateFor(column, row)
    if (editingCell === coordinate) {
      if (event.key === 'Enter') {
        event.preventDefault()
        commitEdit()
      }
      if (event.key === 'Escape') {
        event.preventDefault()
        cancelEdit()
      }
      return
    }
    if (event.key === 'Enter' || event.key === 'F2') {
      event.preventDefault()
      beginEdit(row, column)
      return
    }
    if (event.key === 'Backspace' || event.key === 'Delete') {
      event.preventDefault()
      let next = workbook
      selectedCells.forEach((target) => {
        const [targetRow, targetColumn] = coordinateToIndices(target)
        next = setCellValue(next, sheet.id, targetRow, targetColumn, '')
      })
      setWorkbook(next)
      return
    }
    if (ARROW_DIRECTIONS[event.key]) {
      event.preventDefault()
      moveSelection(ARROW_DIRECTIONS[event.key])
      return
    }
    if (event.key === ' ') {
      event.preventDefault()
      beginEdit(row, column)
      return
    }
    if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
      event.preventDefault()
      beginEdit(row, column, event.key)
    }
  }
  function coordinateToIndices(coordinate) {
    const match = coordinate.match(/^([A-Z]+)(\d+)$/)
    if (!match) return [0, 0]
    let column = 0
    for (const char of match[1]) column = column * 26 + (char.charCodeAt(0) - 64)
    return [Number(match[2]) - 1, column - 1]
  }

  function toggleFont(property) {
    setWorkbook(applyCellsFormat(workbook, sheet.id, selectedCells, { font: { [property]: !allSelectedHaveFont(property) } }))
    cellRefs.current.get(selectedCoordinate)?.focus()
  }
  function setAlignment(value) {
    setWorkbook(applyCellsFormat(workbook, sheet.id, selectedCells, { alignment: { horizontal: allSelectedHaveAlignment(value) ? '' : value } }))
    cellRefs.current.get(selectedCoordinate)?.focus()
  }
  function setTextColor(color) {
    setWorkbook(applyCellsFormat(workbook, sheet.id, selectedCells, { font: { color } }))
    cellRefs.current.get(selectedCoordinate)?.focus()
  }
  function setFillColor(color) {
    setWorkbook(applyCellsFormat(workbook, sheet.id, selectedCells, { fill: { color } }))
    cellRefs.current.get(selectedCoordinate)?.focus()
  }
  function setBorderColor(color) {
    setWorkbook(applyCellsFormat(workbook, sheet.id, selectedCells, { border: { color } }))
    cellRefs.current.get(selectedCoordinate)?.focus()
  }
  function clearFormatting() {
    setWorkbook(clearCellsFormat(workbook, sheet.id, selectedCells))
    cellRefs.current.get(selectedCoordinate)?.focus()
  }

  function startColumnResize(event, columnIndex) {
    event.preventDefault()
    event.stopPropagation()
    const key = `${sheet.id}:${columnIndex}`
    const startX = event.clientX
    const startWidth = columnWidthOverrides[key] || DEFAULT_COLUMN_WIDTH
    const colElement = colRefs.current.get(columnIndex)
    document.body.style.cursor = 'col-resize'
    document.body.style.userSelect = 'none'
    function widthAt(moveEvent) {
      return Math.max(MIN_COLUMN_WIDTH, startWidth + (moveEvent.clientX - startX))
    }
    function handleMove(moveEvent) {
      if (colElement) colElement.style.width = `${widthAt(moveEvent)}px`
    }
    function handleUp(moveEvent) {
      document.removeEventListener('mousemove', handleMove)
      document.removeEventListener('mouseup', handleUp)
      document.body.style.cursor = ''
      document.body.style.userSelect = ''
      setColumnWidthOverrides((current) => ({ ...current, [key]: widthAt(moveEvent) }))
    }
    document.addEventListener('mousemove', handleMove)
    document.addEventListener('mouseup', handleUp)
  }

  function handleAddSheet() {
    const next = addSheet(workbook)
    setWorkbook(next)
    switchSheet(next.sheets[next.sheets.length - 1].id)
  }
  function beginRenameSheet(sheetId, currentName) {
    setRenamingSheetId(sheetId)
    setSheetNameDraft(currentName)
    setContextMenu(null)
  }
  function commitRenameSheet() {
    if (!renamingSheetId) return
    setWorkbook(renameSheet(workbook, renamingSheetId, sheetNameInputRef.current?.value ?? sheetNameDraft))
    setRenamingSheetId(null)
  }
  function cancelRenameSheet() {
    setRenamingSheetId(null)
  }
  function removeSheet(sheetId) {
    if (workbook.sheets.length <= 1) {
      setError('A workbook must contain at least one sheet.')
      setContextMenu(null)
      return
    }
    const index = workbook.sheets.findIndex((item) => item.id === sheetId)
    const next = deleteSheet(workbook, sheetId)
    setWorkbook(next)
    if (activeSheetId === sheetId) {
      const fallback = next.sheets[Math.max(0, index - 1)] || next.sheets[0]
      switchSheet(fallback.id)
    }
    setContextMenu(null)
  }
  function openSheetContextMenu(event, sheetId) {
    event.preventDefault()
    setContextMenu({ x: event.clientX, y: event.clientY, kind: 'sheet', index: sheetId })
  }
  function handleAddRow() {
    setWorkbook(appendRow(workbook, sheet.id))
  }
  function handleAddColumn() {
    setWorkbook(appendColumn(workbook, sheet.id))
  }

  async function handleOpen(event) {
    const file = event.target.files?.[0]
    if (!file) return
    try {
      const buffer = await file.arrayBuffer()
      const loaded = await loadWorkbookFromBuffer(buffer)
      setWorkbook(loaded)
      setFileName(file.name)
      setActiveSheetId(loaded.sheets[0]?.id)
      selectCell(0, 0)
      setError('')
    } catch (loadError) {
      setError(loadError.message)
    }
    event.target.value = ''
  }
  async function handleExport() {
    if (!workbook) return
    try {
      const bytes = await writeWorkbookToZip(workbook)
      const diagnostics = await validateBuffer(bytes)
      if (!diagnostics.valid) {
        setError(`Export produced an invalid package: ${diagnostics.errors[0]?.message}`)
        return
      }
      downloadBytes(bytes, fileName || 'workbook.csvx')
      setError('')
    } catch (exportError) {
      setError(exportError.message)
    }
  }

  if (!workbook || !sheet) {
    return (
      <div className="app-shell">
        <main id="main-content" className="main-content" tabIndex="-1">
          {error ? <p role="alert">{error}</p> : <p role="status">Loading…</p>}
        </main>
      </div>
    )
  }

  return (
    <div className="app-shell">
      <header className="topbar">
        <div className="brand-lockup">
          <span className="brand-mark" aria-hidden="true">X</span>
          <div>
            <p className="eyebrow">CSVX demo</p>
            <h1>Workbench</h1>
          </div>
        </div>
        <div className="file-status" role="status">
          <span className="status-dot" aria-hidden="true" />
          <span>{fileName}</span>
          <span className="status-sep" aria-hidden="true">/</span>
          <strong>{sheet.name}</strong>
          <span className="muted">Core {workbook.version}</span>
        </div>
        <nav className="top-actions" aria-label="File actions">
          <input ref={inputRef} type="file" accept=".csvx,application/zip" onChange={handleOpen} className="sr-only" aria-label="Open CSVX package" />
          <button type="button" className="button button-quiet" onClick={() => inputRef.current?.click()}>Open</button>
          <button type="button" className="button button-primary" onClick={handleExport}>Export</button>
        </nav>
      </header>
      <main id="main-content" className="main-content" tabIndex="-1">
        {error && <p role="alert">{error}</p>}
        <section className="formula-panel" aria-label="Cell inspector">
          <div className="name-box">{selectedCoordinate}</div>
          <div className="formula-symbol" aria-hidden="true">fx</div>
          <div className="formula-value">
            {selectedMetadata?.formula || selectedMetadata?.cached?.value || selectedRawValue || 'Blank cell'}
            {selectedType ? <span className="value-type">{selectedType}</span> : null}
          </div>
          <div className="format-toolbar" role="toolbar" aria-label="Cell formatting">
            <button type="button" className={`format-button ${allSelectedHaveFont('bold') ? 'is-active' : ''}`} aria-pressed={allSelectedHaveFont('bold')} onMouseDown={(e) => e.preventDefault()} onClick={() => toggleFont('bold')} aria-label="Bold" title="Bold (Cmd+B)"><strong>B</strong></button>
            <button type="button" className={`format-button ${allSelectedHaveFont('italic') ? 'is-active' : ''}`} aria-pressed={allSelectedHaveFont('italic')} onMouseDown={(e) => e.preventDefault()} onClick={() => toggleFont('italic')} aria-label="Italic" title="Italic (Cmd+I)"><em>I</em></button>
            <button type="button" className={`format-button ${allSelectedHaveFont('underline') ? 'is-active' : ''}`} aria-pressed={allSelectedHaveFont('underline')} onMouseDown={(e) => e.preventDefault()} onClick={() => toggleFont('underline')} aria-label="Underline" title="Underline (Cmd+U)"><span className="format-underline-glyph">U</span></button>
            <span className="format-divider" aria-hidden="true" />
            <label className="format-button color-swatch" title="Text color">
              <span aria-hidden="true" className="color-swatch-text" style={{ borderBottomColor: selectedStyle.font?.color || '#202124' }}>A</span>
              <input type="color" value={selectedStyle.font?.color || '#000000'} onChange={(e) => setTextColor(e.target.value)} aria-label="Text color" />
            </label>
            <label className="format-button color-swatch" title="Fill color">
              <span aria-hidden="true" className="color-swatch-fill" style={{ backgroundColor: selectedStyle.fill?.color || 'transparent' }} />
              <input type="color" value={selectedStyle.fill?.color || '#ffffff'} onChange={(e) => setFillColor(e.target.value)} aria-label="Fill color" />
            </label>
            <label className="format-button color-swatch" title="Border color">
              <span aria-hidden="true" className="color-swatch-border" style={{ borderColor: selectedStyle.border?.color || 'var(--border-strong)' }} />
              <input type="color" value={selectedStyle.border?.color || '#000000'} onChange={(e) => setBorderColor(e.target.value)} aria-label="Border color" />
            </label>
            <button type="button" className="format-button format-clear" onMouseDown={(e) => e.preventDefault()} onClick={clearFormatting} aria-label="Clear formatting" title="Clear formatting">Clear</button>
            <span className="format-divider" aria-hidden="true" />
            <button type="button" className={`format-button ${allSelectedHaveAlignment('left') ? 'is-active' : ''}`} aria-pressed={allSelectedHaveAlignment('left')} onMouseDown={(e) => e.preventDefault()} onClick={() => setAlignment('left')} aria-label="Align left" title="Align left">L</button>
            <button type="button" className={`format-button ${allSelectedHaveAlignment('center') ? 'is-active' : ''}`} aria-pressed={allSelectedHaveAlignment('center')} onMouseDown={(e) => e.preventDefault()} onClick={() => setAlignment('center')} aria-label="Align center" title="Align center">C</button>
            <button type="button" className={`format-button ${allSelectedHaveAlignment('right') ? 'is-active' : ''}`} aria-pressed={allSelectedHaveAlignment('right')} onMouseDown={(e) => e.preventDefault()} onClick={() => setAlignment('right')} aria-label="Align right" title="Align right">R</button>
          </div>
        </section>
        <section className="grid-card" aria-labelledby="grid-title">
          <h3 id="grid-title" className="sr-only">{sheet.name} spreadsheet data</h3>
          <div className="table-scroll" ref={tableScrollRef}>
            <table className="spreadsheet">
              <caption className="sr-only">CSV-backed data in {sheet.name}. Press Enter, Space, or F2 to edit a cell.</caption>
              <colgroup>
                <col style={{ width: '3rem' }} />
                {sheet.columns.map((column, index) => (
                  <col key={column.id || index} ref={(el) => { if (el) colRefs.current.set(index, el); else colRefs.current.delete(index) }} style={{ width: `${columnWidthOverrides[`${sheet.id}:${index}`] || DEFAULT_COLUMN_WIDTH}px` }} />
                ))}
              </colgroup>
              <thead>
                <tr>
                  <th scope="col" className="corner-cell" aria-label="Spreadsheet corner" />
                  {sheet.columns.map((column, index) => (
                    <th scope="col" key={column.id || index}>
                      <button
                        type="button"
                        className={`column-header-button ${selectedColumns.has(index) ? 'is-active' : ''}`}
                        onClick={(e) => selectColumn(index, e.metaKey || e.ctrlKey)}
                        aria-label={`Select column ${column.name || column.id}`}
                        title={column.width ? `Declared width: ${column.width} (unit not defined by the CSVX schema; not used for display)` : undefined}
                      >
                        {column.name || column.id}
                      </button>
                      <span className="column-resize-handle" onMouseDown={(e) => startColumnResize(e, index)} onClick={(e) => e.stopPropagation()} aria-hidden="true" />
                    </th>
                  ))}
                  <th scope="col">
                    <button type="button" className="column-header-button" onClick={handleAddColumn} aria-label="Add column">+</button>
                  </th>
                </tr>
              </thead>
              <tbody>
                {sheet.records.map((row, rowIndex) => (
                  <tr key={rowIndex}>
                    <th scope="row">
                      <button type="button" className={`row-header-button ${selectedRows.has(rowIndex) ? 'is-active' : ''}`} onClick={(e) => selectRow(rowIndex, e.metaKey || e.ctrlKey)} aria-label={`Select row ${rowIndex + 1}`}>
                        {rowIndex + 1}
                      </button>
                    </th>
                    {sheet.columns.map((_, columnIndex) => {
                      const coordinate = coordinateFor(columnIndex, rowIndex)
                      const value = row[columnIndex] || ''
                      const metadata = cellMetadata(sheet, coordinate)
                      const style = styleForId(workbook.styles, metadata?.style)
                      const isEditing = editingCell === coordinate
                      const displayValue = metadata?.cached?.value ?? value
                      return (
                        <td key={coordinate}>
                          {isEditing ? (
                            <input
                              ref={editorRef}
                              autoFocus
                              type="text"
                              className="cell-editor"
                              style={cellCSS(style)}
                              value={draftValue}
                              onChange={(e) => setDraftValue(e.currentTarget.value)}
                              onBlur={commitEdit}
                              onKeyDown={(e) => {
                                if (e.metaKey || e.ctrlKey) {
                                  const key = e.key.toLowerCase()
                                  if (key === 'b' || key === 'i' || key === 'u') {
                                    e.preventDefault()
                                    e.stopPropagation()
                                    toggleFont(key === 'b' ? 'bold' : key === 'i' ? 'italic' : 'underline')
                                    return
                                  }
                                }
                                e.stopPropagation()
                                if (e.key === 'Enter') { e.preventDefault(); commitEdit(); return }
                                if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); return }
                                if (e.key === 'ArrowUp' || e.key === 'ArrowDown') { e.preventDefault(); commitEditAndMove(ARROW_DIRECTIONS[e.key]); return }
                                const input = e.currentTarget
                                const atStart = input.selectionStart === 0 && input.selectionEnd === 0
                                const atEnd = input.selectionStart === input.value.length && input.selectionEnd === input.value.length
                                if (e.key === 'ArrowLeft' && atStart) { e.preventDefault(); commitEditAndMove(ARROW_DIRECTIONS.ArrowLeft); return }
                                if (e.key === 'ArrowRight' && atEnd) { e.preventDefault(); commitEditAndMove(ARROW_DIRECTIONS.ArrowRight) }
                              }}
                              onMouseDown={(e) => e.stopPropagation()}
                              onClick={(e) => e.stopPropagation()}
                              aria-label={`Edit ${coordinate}`}
                            />
                          ) : (
                            <button
                              type="button"
                              ref={(el) => { if (el) cellRefs.current.set(coordinate, el); else cellRefs.current.delete(coordinate) }}
                              className={`cell-button ${selectedCoordinate === coordinate ? 'is-selected' : ''} ${selectedCoordinate !== coordinate && selectedCells.has(coordinate) ? 'is-in-selection' : ''}`}
                              style={cellCSS(style)}
                              onClick={(e) => (e.metaKey || e.ctrlKey ? toggleCellSelection(rowIndex, columnIndex) : selectCell(rowIndex, columnIndex))}
                              onKeyDown={(e) => handleCellKeyDown(e, rowIndex, columnIndex)}
                              aria-label={`${coordinate}, value ${value || 'blank'}`}
                            >
                              {displayValue}
                              {metadata?.formula ? <span className="formula-indicator" aria-label="Formula"> ƒ</span> : null}
                            </button>
                          )}
                        </td>
                      )
                    })}
                    <td />
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <button type="button" className="button button-quiet" onClick={handleAddRow}>+ Row</button>
        </section>
      </main>
      <footer className="app-footer">
        <nav className="sheet-tabs" aria-label="Workbook sheets">
          <span className="sheet-tabs-label">Sheets<span className="count">{workbook.sheets.length}</span></span>
          <div className="sheet-tabs-list">
            {workbook.sheets.map((item) =>
              renamingSheetId === item.id ? (
                <input
                  key={item.id}
                  ref={sheetNameInputRef}
                  type="text"
                  className="sheet-tab-input"
                  value={sheetNameDraft}
                  onChange={(e) => setSheetNameDraft(e.target.value)}
                  onBlur={commitRenameSheet}
                  onClick={(e) => e.stopPropagation()}
                  onKeyDown={(e) => {
                    e.stopPropagation()
                    if (e.key === 'Enter') { e.preventDefault(); commitRenameSheet() }
                    if (e.key === 'Escape') { e.preventDefault(); cancelRenameSheet() }
                  }}
                  aria-label={`Rename sheet ${item.name}`}
                />
              ) : (
                <button
                  key={item.id}
                  type="button"
                  className={`sheet-tab ${item.id === activeSheetId ? 'is-active' : ''}`}
                  onClick={() => switchSheet(item.id)}
                  onDoubleClick={() => beginRenameSheet(item.id, item.name)}
                  onContextMenu={(e) => openSheetContextMenu(e, item.id)}
                  aria-current={item.id === activeSheetId ? 'true' : undefined}
                >
                  {item.name}
                </button>
              ),
            )}
            <button type="button" className="sheet-tab-add" onClick={handleAddSheet} aria-label="Add sheet" title="Add sheet">+</button>
          </div>
        </nav>
        <div className="status-bar" role="status">
          <strong>{sheet.records.length}</strong>&nbsp;rows · <strong>{sheet.columns.length}</strong>&nbsp;columns
        </div>
        <span className="footer-hint">Enter/F2 edit · Right-click a sheet tab for actions</span>
      </footer>
      {contextMenu && (
        <div ref={contextMenuRef} className="row-context-menu" role="menu" style={{ left: contextMenu.x, top: contextMenu.y }} onClick={(e) => e.stopPropagation()}>
          <p className="context-menu-label">{workbook.sheets.find((item) => item.id === contextMenu.index)?.name}</p>
          <button type="button" role="menuitem" onClick={() => beginRenameSheet(contextMenu.index, workbook.sheets.find((item) => item.id === contextMenu.index)?.name || '')}>Rename sheet</button>
          <button type="button" role="menuitem" className="danger-action" onClick={() => removeSheet(contextMenu.index)}>Delete sheet</button>
        </div>
      )}
    </div>
  )
}

export default App
