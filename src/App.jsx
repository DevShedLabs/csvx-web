import { useEffect, useMemo, useRef, useState } from 'react'
import { flushSync } from 'react-dom'
import PrintView from './PrintView.jsx'
import { columnId, columnWidthToPixels, importCSV, rowNumberFor, formatValue, loadWorkbookFromZip, resolveCellValue, validateBuffer, writeWorkbookToZip } from 'csvx-ts/browser'
// loadWorkbookFromBuffer below intentionally calls loadWorkbookFromZip once, not
// validateBuffer-then-loadWorkbookFromZip — validateBuffer (csvx-ts src/package.ts) just calls
// loadWorkbookFromZip internally and classifies the thrown error, so calling both re-parses the
// whole ZIP/CSV/JSON a second time for the same information.
import { downloadBytes } from './download.js'
import {
  addSheet,
  appendColumn,
  applyCellsFormat,
  applyTextCase,
  cellCSS,
  cellBorderCSS,
  setPrintSettings,
  borderColorOf,
  cellMetadata,
  rawCellText,
  HEADER_ROW,
  clearCellsFormat,
  coordinateFor,
  deleteColumns,
  deleteRows,
  deleteSheet,
  findSheet,
  indicesForCoordinate,
  insertColumns,
  insertRows,
  recalculateWorkbook,
  renameSheet,
  setCellValue,
  setColumnWidth,
} from './model.js'

const DEMO_URL = '/example.csvx'
const DEFAULT_COLUMN_WIDTH = 128
const MIN_COLUMN_WIDTH = 48
// USD only for now — a currency picker (and the locale-aware formatting that implies) is real,
// separate scope; see csvx-ts's format.ts comment on why general multi-currency/locale number
// parsing isn't attempted yet. This is just the one Excel-style format code csvx-ts already
// supports for display and round-trip-safe literal parsing (quoted-literal prefix, 2 decimals).
const USD_NUMBER_FORMAT = '"$"#,##0.00'
// File System Access API — Chromium only (not Firefox/Safari as of this writing). Where it's
// available, Open keeps a live handle so Save can write back in place; everywhere else, Open uses
// the classic <input type="file"> and Save always falls back to a download (Save As behavior).
const SUPPORTS_FILE_SYSTEM_ACCESS = typeof window !== 'undefined' && typeof window.showOpenFilePicker === 'function'
const CSVX_PICKER_TYPES = [{ description: 'CSVX workbook', accept: { 'application/zip': ['.csvx'] } }]
// Open also takes a plain CSV, which csvx-ts's importCSV (spec 11-import-export.md 11.1) converts.
// One combined entry so both extensions show by default (a second entry hides CSV behind the
// picker's file-type dropdown).
const OPEN_PICKER_TYPES = [{ description: 'CSVX workbook or CSV file', accept: { 'application/zip': ['.csvx'], 'text/csv': ['.csv'] } }]
const CSV_DELIMITERS = [
  { label: 'Comma (,)', value: ',' },
  { label: 'Semicolon (;)', value: ';' },
  { label: 'Tab', value: '\t' },
  { label: 'Pipe (|)', value: '|' },
]
const isCSVFile = (file) => /\.csv$/i.test(file.name) || file.type === 'text/csv'

// Column.width is in XLSX character-width units, per spec/03-sheets.md — columnWidthToPixels
// (csvx-ts) is the one canonical conversion every engine agrees on; a resize writes back through
// its inverse, pixelsToColumnWidth (see setColumnWidth in model.js), so this now round-trips
// instead of being display-only.
function pixelWidthForColumn(column) {
  if (typeof column?.width !== 'number') return DEFAULT_COLUMN_WIDTH
  return Math.max(MIN_COLUMN_WIDTH, columnWidthToPixels(column.width))
}
const ARROW_DIRECTIONS = { ArrowLeft: { column: -1, row: 0 }, ArrowRight: { column: 1, row: 0 }, ArrowUp: { column: 0, row: -1 }, ArrowDown: { column: 0, row: 1 } }
// Must match `.spreadsheet th, .spreadsheet td { height: 1.75rem }` in index.css (1.75rem * 16px).
// Real imported workbooks run to thousands of rows (example.csvx has ~1000); rendering every row
// as a live DOM <tr> of <button>s made every click/format action re-render tens of thousands of
// nodes. Only rows within ROW_OVERSCAN of the visible scroll window are actually mounted.
const ROW_HEIGHT = 28
const ROW_OVERSCAN = 10

const ALIGN_ICON_BARS = {
  left: [
    [1, 3, 14],
    [1, 7, 9],
    [1, 11, 12],
  ],
  center: [
    [1, 3, 14],
    [3.5, 7, 9],
    [2, 11, 12],
  ],
  right: [
    [1, 3, 14],
    [6, 7, 9],
    [3, 11, 12],
  ],
}

function AlignIcon({ variant }) {
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
      {ALIGN_ICON_BARS[variant].map(([x, y, width]) => (
        <rect key={y} x={x} y={y} width={width} height="1.5" rx="0.75" fill="currentColor" />
      ))}
    </svg>
  )
}

async function loadWorkbookFromBuffer(buffer) {
  try {
    const loaded = await loadWorkbookFromZip(buffer)
    // A formula's cached value is advisory (spec/05-cell-values.md) and may be stale or missing
    // (e.g. authored by a tool that can't calculate) — recalculate once on load rather than trust it.
    return recalculateWorkbook(loaded)
  } catch (loadError) {
    throw new Error(loadError.message || 'Package failed to load')
  }
}

function App() {
  const [workbook, setWorkbook] = useState(null)
  const [fileName, setFileName] = useState('')
  const [activeSheetId, setActiveSheetId] = useState(null)
  const [selectedCell, setSelectedCell] = useState({ row: HEADER_ROW, column: 0 })
  const [selectedCells, setSelectedCells] = useState(() => new Set(['A1']))
  const [selectedColumns, setSelectedColumns] = useState(() => new Set())
  const [selectedRows, setSelectedRows] = useState(() => new Set())
  const [editingCell, setEditingCell] = useState(null)
  // Which input is the live editor for editingCell — the grid's inline cell input, or the formula
  // bar. Only one can actually be focused at a time; this is what stops the grid cell's autoFocus
  // from stealing focus back from the formula bar while the user is typing into it there.
  const [editingSource, setEditingSource] = useState('grid')
  // Only the *initial* value of the in-progress edit — the <input> below is uncontrolled
  // (defaultValue, not value) so keystrokes don't setState on every character and re-render the
  // whole grid. Commit reads the live value from editorRef/formulaBarRef instead; draftValue is
  // just the fallback for the instant before that ref attaches.
  const [draftValue, setDraftValue] = useState('')
  // Bumped on cancel (and after a commit) to force the always-mounted formula bar input to remount
  // and re-read its defaultValue — unlike the grid's cell input, it doesn't unmount on its own when
  // editing ends, so a cancelled or committed edit would otherwise leave stale typed text showing.
  const [editRevision, setEditRevision] = useState(0)
  const [contextMenu, setContextMenu] = useState(null)
  const [error, setError] = useState('')
  // A CSV picked in Open, waiting for the user to confirm the import options (UI state only).
  const [pendingCSV, setPendingCSV] = useState(null)
  const [importNotice, setImportNotice] = useState('')
  const [renamingSheetId, setRenamingSheetId] = useState(null)
  const [sheetNameDraft, setSheetNameDraft] = useState('')
  const [scrollTop, setScrollTop] = useState(0)
  const [viewportHeight, setViewportHeight] = useState(0)
  // Guards the demo-fixture fetch below against clobbering a workbook the user opened themselves
  // while that fetch was still in flight — without this, a slow demo-fetch response arriving after
  // a manual Open would silently overwrite the user's file with the fixture.
  const userOpenedRef = useRef(false)
  // The live handle for the file the user opened via the File System Access API, if the browser
  // supports it and they didn't open via the plain <input type="file"> fallback or the demo fetch
  // — this is what lets Save write back to the same file in place instead of always downloading a
  // new copy. Not state: it never needs to trigger a re-render on its own.
  const fileHandleRef = useRef(null)
  const inputRef = useRef(null)
  const editorRef = useRef(null)
  const formulaBarRef = useRef(null)
  // The formula bar never unmounts the way the grid's inline editor does, so Enter/Escape calling
  // .blur() to leave it synchronously re-fires onBlur's commitEdit with the *previous* render's
  // editingCell (state updates from the keydown handler haven't flushed yet) — this suppresses that
  // redundant/incorrect re-commit for exactly one blur.
  const suppressFormulaBarBlurRef = useRef(false)
  const cellRefs = useRef(new Map())
  const colRefs = useRef(new Map())
  const contextMenuRef = useRef(null)
  const tableScrollRef = useRef(null)
  const dragRef = useRef(null)
  const suppressClickRef = useRef(false)
  const applyRangeRef = useRef(null)
  const [printOpen, setPrintOpen] = useState(false)
  const sheetNameInputRef = useRef(null)

  const sheet = useMemo(() => (workbook ? findSheet(workbook, activeSheetId) : null), [workbook, activeSheetId])
  // styleForId does a linear scan of workbook.styles. The whole table re-renders on every click or
  // selection change, and (without this map) every cell in the sheet would re-run that scan every
  // time — O(rows * columns * styles) per click, which is the actual source of the click-to-focus
  // lag on any workbook with a non-trivial number of styles. Built once per workbook.styles
  // reference instead.
  const stylesById = useMemo(() => new Map((workbook?.styles || []).map((style) => [style.id, style])), [workbook?.styles])
  const styleFor = (id) => stylesById.get(id) || {}
  // The header is row 1 of the grid, so the scroll window is over records + 1 rows; `g` below is a
  // grid row (0 = header) and the record index is g - 1.
  const totalRows = (sheet?.records?.length || 0) + 1
  const columnCount = sheet?.columns?.length || 0
  const startRow = Math.max(0, Math.floor(scrollTop / ROW_HEIGHT) - ROW_OVERSCAN)
  const visibleRowCount = Math.ceil((viewportHeight || 0) / ROW_HEIGHT) + ROW_OVERSCAN * 2
  const endRow = Math.min(totalRows, startRow + visibleRowCount)
  const topSpacerHeight = startRow * ROW_HEIGHT
  const bottomSpacerHeight = (totalRows - endRow) * ROW_HEIGHT
  const selectedCoordinate = coordinateFor(selectedCell.column, selectedCell.row)
  const selectedMetadata = cellMetadata(sheet, selectedCoordinate)
  const selectedRawValue = rawCellText(sheet, selectedCell.row, selectedCell.column)
  const selectedStyle = styleFor(selectedMetadata?.style)
  const selectedAlignment = selectedStyle.alignment?.horizontal || ''
  // The resolved type (via the engine, same as the grid's own display logic) — never a raw
  // metadata.type read directly, since an untyped cell's effective type (including one inferred
  // from its literal text, per resolveCellValue) is exactly what the badge should communicate, not
  // just whatever happens to be explicitly declared. Hidden for a genuinely blank cell — showing
  // "blank" on every empty cell would just be noise.
  const selectedDeclaredType = selectedMetadata?.type || (selectedCell.row < 0 ? 'string' : sheet?.columns?.[selectedCell.column]?.type)
  const selectedResolvedValue = selectedMetadata?.cached ?? resolveCellValue(selectedRawValue, selectedDeclaredType, selectedStyle.numberFormat)
  const selectedType = selectedResolvedValue.type === 'blank' ? '' : selectedResolvedValue.type
  const selectionStyles = useMemo(
    () => [...selectedCells].map((coordinate) => styleFor(cellMetadata(sheet, coordinate)?.style)),
    [selectedCells, sheet, stylesById],
  )
  const allSelectedHaveFont = (property) => selectionStyles.length > 0 && selectionStyles.every((style) => style.font?.[property])
  const allSelectedHaveAlignment = (value) => selectionStyles.length > 0 && selectionStyles.every((style) => (style.alignment?.horizontal || '') === value)
  const allSelectedAreCurrency = selectionStyles.length > 0 && selectionStyles.every((style) => style.numberFormat === USD_NUMBER_FORMAT)

  // Load the real example fixture on first mount, through the real engine, instead of shipping
  // fabricated demo data — see AGENTS.md: this app must never hold a second opinion about what a
  // CSVX file contains.
  useEffect(() => {
    let cancelled = false
    fetch(DEMO_URL)
      .then((response) => response.arrayBuffer())
      .then(async (buffer) => {
        const loaded = await loadWorkbookFromBuffer(buffer)
        if (cancelled || userOpenedRef.current) return
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
  // Any way of printing (Cmd/Ctrl+P, the browser menu) goes through the Print view, so the output
  // is always the paginated pages and never the live grid. flushSync renders the view before the
  // browser snapshots the page.
  useEffect(() => {
    const open = () => flushSync(() => setPrintOpen(true))
    window.addEventListener('beforeprint', open)
    return () => window.removeEventListener('beforeprint', open)
  }, [])
  // Tracks the scroller's height so the row-window size (see ROW_HEIGHT/ROW_OVERSCAN above) can
  // adapt to the actual viewport instead of a guessed row count.
  useEffect(() => {
    const el = tableScrollRef.current
    if (!el) return
    const updateHeight = () => setViewportHeight(el.clientHeight)
    updateHeight()
    const observer = new ResizeObserver(updateHeight)
    observer.observe(el)
    return () => observer.disconnect()
  }, [sheet?.id])
  // New sheet (different row count) starts scrolled to the top; otherwise a stale scrollTop from a
  // longer sheet could compute a row window past the end of a shorter one.
  useEffect(() => {
    const el = tableScrollRef.current
    if (el) el.scrollTop = 0
    setScrollTop(0)
  }, [sheet?.id])
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
  // Drag-select: while the button is down, find the cell under the pointer (clamped to the grid, so
  // dragging past an edge keeps extending) and auto-scroll near the edges — only the rows near the
  // viewport are mounted, so scrolling is what brings further cells into reach.
  useEffect(() => {
    const ROW_HEADER_WIDTH = 48
    const COLUMN_HEADER_HEIGHT = 36
    const EDGE = 28
    let pointer = null
    let timer = null
    const clamp = (value, low, high) => Math.min(Math.max(value, low), Math.max(low, high))
    function update() {
      const drag = dragRef.current
      const scroller = tableScrollRef.current
      if (!drag || !scroller || !pointer) return
      // Use the client area, not the outer rect: scrollbars sit inside the outer rect and aren't cells.
      const rect = scroller.getBoundingClientRect()
      const right = rect.left + scroller.clientWidth
      const bottom = rect.top + scroller.clientHeight
      const x = clamp(pointer.x, rect.left + ROW_HEADER_WIDTH + 2, right - 2)
      const y = clamp(pointer.y, rect.top + COLUMN_HEADER_HEIGHT + 2, bottom - 2)
      // The clamped point can land on a scrollbar or gutter rather than a cell, so step toward the
      // middle of the grid until a cell is found.
      const centerX = (rect.left + ROW_HEADER_WIDTH + right) / 2
      const centerY = (rect.top + COLUMN_HEADER_HEIGHT + bottom) / 2
      let cell = null
      for (let step = 0; step <= 8 && !cell; step += 1) {
        const px = x + Math.sign(centerX - x) * Math.min(Math.abs(centerX - x), step * 8)
        const py = y + Math.sign(centerY - y) * Math.min(Math.abs(centerY - y), step * 8)
        cell = document.elementFromPoint(px, py)?.closest('td[data-row]')
      }
      if (!cell) return
      const focus = { row: Number(cell.dataset.row), column: Number(cell.dataset.col) }
      const key = `${focus.row},${focus.column}`
      if (key === drag.last) return
      drag.last = key
      if (focus.row !== drag.anchor.row || focus.column !== drag.anchor.column) suppressClickRef.current = true
      applyRangeRef.current(drag.kind, drag.anchor, focus)
    }
    function scrollStep() {
      const scroller = tableScrollRef.current
      if (!dragRef.current || !scroller || !pointer) return
      const rect = scroller.getBoundingClientRect()
      const right = rect.left + scroller.clientWidth
      const bottom = rect.top + scroller.clientHeight
      const speed = (distance) => Math.round(Math.min(40, 4 + Math.abs(distance) / 2)) * Math.sign(distance)
      const top = rect.top + COLUMN_HEADER_HEIGHT
      const left = rect.left + ROW_HEADER_WIDTH
      let dy = 0
      let dx = 0
      if (pointer.y > bottom - EDGE) dy = speed(pointer.y - (bottom - EDGE))
      else if (pointer.y < top + EDGE) dy = speed(pointer.y - (top + EDGE))
      if (pointer.x > right - EDGE) dx = speed(pointer.x - (right - EDGE))
      else if (pointer.x < left + EDGE) dx = speed(pointer.x - (left + EDGE))
      if (dx || dy) {
        scroller.scrollBy(dx, dy)
        setScrollTop(scroller.scrollTop)
        // The newly mounted rows render on the next frame; re-locate the pointer then.
        requestAnimationFrame(update)
      }
    }
    function onMove(event) {
      if (!dragRef.current) return
      pointer = { x: event.clientX, y: event.clientY }
      if (!timer) timer = setInterval(scrollStep, 30)
      update()
    }
    function onUp() {
      dragRef.current = null
      pointer = null
      if (timer) clearInterval(timer)
      timer = null
    }
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
    return () => {
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
      if (timer) clearInterval(timer)
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
      } else if (key === 'a' && !/^(INPUT|TEXTAREA|SELECT)$/.test(event.target?.tagName)) {
        // Cmd/Ctrl+A would otherwise select the whole painted DOM, not the sheet's cells.
        event.preventDefault()
        selectAll()
      } else if (key === 'p') {
        event.preventDefault()
        setPrintOpen(true)
      } else if (key === 's') {
        // Saves in place instead of letting the browser try to "Save page as…".
        event.preventDefault()
        handleSave()
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

  // Brings a row into the mounted window (see ROW_HEIGHT/ROW_OVERSCAN) before anything tries to
  // focus it — needed for keyboard navigation (moveSelection), since a row outside the current
  // scroll window has no DOM node to focus yet. A no-op when the row is already visible.
  function ensureRowVisible(rowIndex) {
    const el = tableScrollRef.current
    if (!el) return
    const rowTop = (rowIndex + 1) * ROW_HEIGHT
    const rowBottom = rowTop + ROW_HEIGHT
    let nextScrollTop = el.scrollTop
    if (rowTop < el.scrollTop) nextScrollTop = rowTop
    else if (rowBottom > el.scrollTop + el.clientHeight) nextScrollTop = rowBottom - el.clientHeight
    if (nextScrollTop !== el.scrollTop) {
      el.scrollTop = nextScrollTop
      setScrollTop(nextScrollTop)
    }
  }
  function selectCell(row, column) {
    ensureRowVisible(row)
    setSelectedCell({ row, column })
    setSelectedCells(new Set([coordinateFor(column, row)]))
    setSelectedColumns(new Set())
    setSelectedRows(new Set())
    setContextMenu(null)
  }
  // Range selection (click-drag, Shift+click). `kind` says what is being dragged: 'cell' selects the
  // rectangle between anchor and focus, 'column' / 'row' select whole columns / rows between them
  // (a column includes the header row). The active cell stays at the anchor, as in Sheets and Excel.
  function applyRange(kind, anchor, focus) {
    const lastRow = sheet.records.length - 1
    const lastColumn = sheet.columns.length - 1
    let [r0, r1] = [Math.min(anchor.row, focus.row), Math.max(anchor.row, focus.row)]
    let [c0, c1] = [Math.min(anchor.column, focus.column), Math.max(anchor.column, focus.column)]
    if (kind === 'column') [r0, r1] = [HEADER_ROW, lastRow]
    if (kind === 'row') [c0, c1] = [0, lastColumn]
    const coordinates = new Set()
    for (let row = r0; row <= r1; row += 1) for (let column = c0; column <= c1; column += 1) coordinates.add(coordinateFor(column, row))
    const span = (from, to) => new Set(Array.from({ length: to - from + 1 }, (_, offset) => from + offset))
    setSelectedCell(anchor)
    setSelectedCells(coordinates)
    setSelectedColumns(kind === 'column' ? span(c0, c1) : new Set())
    setSelectedRows(kind === 'row' ? span(r0, r1) : new Set())
    setContextMenu(null)
  }
  applyRangeRef.current = applyRange
  // Mouse-down starts a drag (and a plain press selects immediately); Shift+press extends from the
  // active cell instead. Ctrl/Cmd is left to the existing click handlers (additive toggling).
  function startDrag(kind, target, event) {
    if (event.button !== 0 || event.metaKey || event.ctrlKey) return
    suppressClickRef.current = false
    let anchor = target
    if (event.shiftKey) {
      anchor = kind === 'column' ? { row: HEADER_ROW, column: selectedCell.column } : kind === 'row' ? { row: selectedCell.row, column: 0 } : selectedCell
      applyRange(kind, anchor, target)
      suppressClickRef.current = true
    } else if (kind === 'cell') selectCell(target.row, target.column)
    else if (kind === 'column') selectColumn(target.column, false)
    else selectRow(target.row, false)
    dragRef.current = { kind, anchor, last: `${target.row},${target.column}` }
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
    const coordinates = [HEADER_ROW, ...sheet.records.map((_, rowIndex) => rowIndex)].map((rowIndex) => coordinateFor(columnIndex, rowIndex))
    setSelectedCell({ row: HEADER_ROW, column: columnIndex })
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
  function styleAt(row, column) {
    if (row < HEADER_ROW || column < 0 || row >= sheet.records.length || column >= sheet.columns.length) return undefined
    return styleFor(cellMetadata(sheet, coordinateFor(column, row))?.style)
  }
  function neighborStyles(row, column) {
    return { above: styleAt(row - 1, column), below: styleAt(row + 1, column), left: styleAt(row, column - 1), right: styleAt(row, column + 1) }
  }
  function displayAt(row, column) {
    const metadata = cellMetadata(sheet, coordinateFor(column, row))
    const style = styleFor(metadata?.style)
    const declaredType = metadata?.type || (row < 0 ? 'string' : sheet.columns[column]?.type)
    const cellValue = metadata?.cached ?? resolveCellValue(rawCellText(sheet, row, column), declaredType, style.numberFormat)
    return formatValue(cellValue, style.numberFormat)
  }
  function selectAll() {
    const coordinates = []
    for (let rowIndex = HEADER_ROW; rowIndex < sheet.records.length; rowIndex += 1) sheet.columns.forEach((_, columnIndex) => coordinates.push(coordinateFor(columnIndex, rowIndex)))
    setSelectedCell({ row: HEADER_ROW, column: 0 })
    setSelectedCells(new Set(coordinates))
    setSelectedColumns(new Set())
    setSelectedRows(new Set())
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

  function beginEdit(row, column, initialValue, source = 'grid') {
    const coordinate = coordinateFor(column, row)
    setSelectedCell({ row, column })
    // Editing an existing formula cell edits the formula text, never its cached result — the CSV
    // cell itself holds the cache (spec/03-sheets.md), so sheet.records alone isn't the right
    // source once a formula is involved.
    const existingFormula = cellMetadata(sheet, coordinate)?.formula
    setDraftValue(initialValue !== undefined ? initialValue : existingFormula ?? rawCellText(sheet, row, column))
    setEditingCell(coordinate)
    setEditingSource(source)
    setContextMenu(null)
  }
  function currentEditorValue() {
    const ref = editingSource === 'formula-bar' ? formulaBarRef : editorRef
    return ref.current?.value ?? draftValue
  }
  function commitEdit() {
    if (!editingCell) return
    const value = currentEditorValue()
    setWorkbook(setCellValue(workbook, sheet.id, selectedCell.row, selectedCell.column, value))
    setEditingCell(null)
    setEditRevision((revision) => revision + 1)
  }
  function cancelEdit() {
    setEditingCell(null)
    setEditRevision((revision) => revision + 1)
  }
  function moveSelection(direction) {
    const nextColumn = Math.max(0, Math.min(sheet.columns.length - 1, selectedCell.column + direction.column))
    const nextRow = Math.max(HEADER_ROW, Math.min(sheet.records.length - 1, selectedCell.row + direction.row))
    selectCell(nextRow, nextColumn)
  }
  function commitEditAndMove(direction) {
    if (!editingCell) return
    const value = currentEditorValue()
    setWorkbook(setCellValue(workbook, sheet.id, selectedCell.row, selectedCell.column, value))
    setEditingCell(null)
    setEditRevision((revision) => revision + 1)
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
        const { row: targetRow, column: targetColumn } = indicesForCoordinate(target)
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
  function toggleFont(property) {
    setWorkbook(applyCellsFormat(workbook, sheet.id, selectedCells, { font: { [property]: !allSelectedHaveFont(property) } }))
    cellRefs.current.get(selectedCoordinate)?.focus()
  }
  function setAlignment(value) {
    setWorkbook(applyCellsFormat(workbook, sheet.id, selectedCells, { alignment: { horizontal: allSelectedHaveAlignment(value) ? '' : value } }))
    cellRefs.current.get(selectedCoordinate)?.focus()
  }
  function toggleCurrency() {
    setWorkbook(applyCellsFormat(workbook, sheet.id, selectedCells, { numberFormat: allSelectedAreCurrency ? null : USD_NUMBER_FORMAT }))
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
    setWorkbook(applyCellsFormat(workbook, sheet.id, selectedCells, { border: { top: { style: 'thin', color }, right: { style: 'thin', color }, bottom: { style: 'thin', color }, left: { style: 'thin', color } } }))
    cellRefs.current.get(selectedCoordinate)?.focus()
  }
  function changeTextCase(mode) {
    setWorkbook(applyTextCase(workbook, sheet.id, selectedCells, mode))
    cellRefs.current.get(selectedCoordinate)?.focus()
  }
  function clearFormatting() {
    setWorkbook(clearCellsFormat(workbook, sheet.id, selectedCells))
    cellRefs.current.get(selectedCoordinate)?.focus()
  }

  function startColumnResize(event, columnIndex) {
    event.preventDefault()
    event.stopPropagation()
    const startX = event.clientX
    const startWidth = pixelWidthForColumn(sheet.columns[columnIndex])
    const colElement = colRefs.current.get(columnIndex)
    document.body.style.cursor = 'col-resize'
    document.body.style.userSelect = 'none'
    function widthAt(moveEvent) {
      return Math.max(MIN_COLUMN_WIDTH, startWidth + (moveEvent.clientX - startX))
    }
    function handleMove(moveEvent) {
      // Direct DOM mutation during the drag, not React state — committing to the real model (and
      // triggering a full grid re-render) on every mousemove would make resizing feel sluggish.
      if (colElement) colElement.style.width = `${widthAt(moveEvent)}px`
    }
    function handleUp(moveEvent) {
      document.removeEventListener('mousemove', handleMove)
      document.removeEventListener('mouseup', handleUp)
      document.body.style.cursor = ''
      document.body.style.userSelect = ''
      setWorkbook((current) => setColumnWidth(current, sheet.id, columnIndex, widthAt(moveEvent)))
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
  function handleAddColumn() {
    setWorkbook(appendColumn(workbook, sheet.id))
  }
  // Right-clicking inside the current row/column selection keeps it and the menu acts on all of it;
  // right-clicking outside selects just that row/column first, as in Sheets and Excel.
  function openRowContextMenu(event, rowIndex) {
    event.preventDefault()
    // Row 1 is the header: it can't be inserted before or deleted, so it has no row menu.
    if (rowIndex < 0) return
    const keep = selectedRows.has(rowIndex)
    if (!keep) selectRow(rowIndex, false)
    const rows = keep ? [...selectedRows].filter((row) => row >= 0).sort((a, b) => a - b) : [rowIndex]
    setContextMenu({ x: event.clientX, y: event.clientY, kind: 'row', index: rowIndex, items: rows })
  }
  function openColumnContextMenu(event, columnIndex) {
    event.preventDefault()
    const keep = selectedColumns.has(columnIndex)
    if (!keep) selectColumn(columnIndex, false)
    const columns = keep ? [...selectedColumns].sort((a, b) => a - b) : [columnIndex]
    setContextMenu({ x: event.clientX, y: event.clientY, kind: 'column', index: columnIndex, items: columns })
  }
  // Batch operations: one pass and one recalculation however many rows or columns are involved.
  function handleInsertRows(atIndex, count) {
    setWorkbook(insertRows(workbook, sheet.id, atIndex, count))
    setContextMenu(null)
  }
  function handleDeleteRows(rows) {
    if (rows.length >= sheet.records.length) {
      setError('A sheet must contain at least one row.')
      setContextMenu(null)
      return
    }
    setWorkbook(deleteRows(workbook, sheet.id, rows))
    selectCell(Math.max(0, Math.min(...rows) - 1), selectedCell.column)
    setContextMenu(null)
  }
  function handleInsertColumns(atIndex, count) {
    setWorkbook(insertColumns(workbook, sheet.id, atIndex, count))
    setContextMenu(null)
  }
  function handleDeleteColumns(columns) {
    if (columns.length >= sheet.columns.length) {
      setError('A sheet must contain at least one column.')
      setContextMenu(null)
      return
    }
    setWorkbook(deleteColumns(workbook, sheet.id, columns))
    selectCell(selectedCell.row, Math.max(0, Math.min(...columns) - 1))
    setContextMenu(null)
  }

  async function openFile(file) {
    userOpenedRef.current = true
    if (isCSVFile(file)) {
      setPendingCSV({ file, header: true, infer: true, delimiter: ',' })
      return
    }
    setImportNotice('')
    const buffer = await file.arrayBuffer()
    const loaded = await loadWorkbookFromBuffer(buffer)
    setWorkbook(loaded)
    setFileName(file.name)
    setActiveSheetId(loaded.sheets[0]?.id)
    selectCell(0, 0)
    setError('')
  }
  // Converts the pending CSV with the engine; the app only supplies the options the user chose.
  async function confirmCSVImport() {
    const { file, header, infer, delimiter } = pendingCSV
    try {
      const bytes = new Uint8Array(await file.arrayBuffer())
      const stem = file.name.replace(/\.[^.]*$/, '')
      const { workbook: imported, warnings } = importCSV(bytes, { header, infer, delimiter, name: stem || undefined })
      fileHandleRef.current = null
      setWorkbook(imported)
      setFileName(`${stem || 'workbook'}.csvx`)
      setActiveSheetId(imported.sheets[0]?.id)
      selectCell(0, 0)
      setError('')
      setImportNotice(warnings.length ? `Imported with ${warnings.length} warning${warnings.length === 1 ? '' : 's'}: ${warnings.slice(0, 3).map((w) => `${w.location}: ${w.reason}`).join('; ')}${warnings.length > 3 ? '; …' : ''}` : '')
    } catch (importError) {
      setError(`Could not import CSV: ${importError.message}`)
    }
    setPendingCSV(null)
  }
  // Classic fallback path: browsers without the File System Access API (Firefox, Safari) or a user
  // who dismisses the native picker's capability in favor of a plain file input. No handle means
  // Save can't write back in place for this file — it'll fall back to Save As.
  async function handleOpenInputChange(event) {
    const file = event.target.files?.[0]
    if (!file) return
    fileHandleRef.current = null
    try {
      await openFile(file)
    } catch (loadError) {
      setError(loadError.message)
    }
    event.target.value = ''
  }
  async function handleOpenClick() {
    if (!SUPPORTS_FILE_SYSTEM_ACCESS) {
      inputRef.current?.click()
      return
    }
    try {
      const [handle] = await window.showOpenFilePicker({ types: OPEN_PICKER_TYPES })
      const file = await handle.getFile()
      // A CSV has no CSVX handle to save back to; Save falls back to Save As (.csvx).
      fileHandleRef.current = isCSVFile(file) ? null : handle
      await openFile(file)
    } catch (pickerError) {
      if (pickerError.name === 'AbortError') return
      setError(pickerError.message)
    }
  }
  async function buildExportBytes() {
    const bytes = await writeWorkbookToZip(workbook)
    const diagnostics = await validateBuffer(bytes)
    if (!diagnostics.valid) throw new Error(`Save produced an invalid package: ${diagnostics.errors[0]?.message}`)
    return bytes
  }
  // Writes back to the file the user opened, in place — the behavior any real editor's Cmd/Ctrl+S
  // gives you. Falls back to Save As when there's no live handle to write through (opened via the
  // classic input, opened the read-only demo fixture, or the browser doesn't support the API).
  async function handleSave() {
    if (!workbook) return
    try {
      const bytes = await buildExportBytes()
      if (fileHandleRef.current) {
        const writable = await fileHandleRef.current.createWritable()
        await writable.write(bytes)
        await writable.close()
        setError('')
        return
      }
      await saveAsWithBytes(bytes)
    } catch (saveError) {
      setError(saveError.message)
    }
  }
  async function saveAsWithBytes(bytes) {
    if (SUPPORTS_FILE_SYSTEM_ACCESS) {
      try {
        const handle = await window.showSaveFilePicker({ suggestedName: fileName || 'workbook.csvx', types: CSVX_PICKER_TYPES })
        const writable = await handle.createWritable()
        await writable.write(bytes)
        await writable.close()
        fileHandleRef.current = handle
        setFileName(handle.name)
        setError('')
        return
      } catch (pickerError) {
        if (pickerError.name === 'AbortError') return
        setError(pickerError.message)
        return
      }
    }
    downloadBytes(bytes, fileName || 'workbook.csvx')
    setError('')
  }
  async function handleSaveAs() {
    if (!workbook) return
    try {
      await saveAsWithBytes(await buildExportBytes())
    } catch (saveError) {
      setError(saveError.message)
    }
  }
  const importDialog = pendingCSV && (
  <div className="import-overlay" role="dialog" aria-modal="true" aria-labelledby="import-title">
            <form
              className="import-dialog"
              onSubmit={(event) => {
                event.preventDefault()
                confirmCSVImport()
              }}
            >
              <h2 id="import-title">Import CSV</h2>
              <p className="muted">{pendingCSV.file.name} will be converted to a CSVX workbook.</p>
              <label>
                <input type="checkbox" checked={pendingCSV.header} onChange={(event) => setPendingCSV({ ...pendingCSV, header: event.target.checked })} /> First row is a header
              </label>
              <label>
                <input type="checkbox" checked={pendingCSV.infer} onChange={(event) => setPendingCSV({ ...pendingCSV, infer: event.target.checked })} /> Detect column types
              </label>
              <label>
                Delimiter{' '}
                <select value={pendingCSV.delimiter} onChange={(event) => setPendingCSV({ ...pendingCSV, delimiter: event.target.value })}>
                  {CSV_DELIMITERS.map((option) => (
                    <option key={option.value} value={option.value}>{option.label}</option>
                  ))}
                </select>
              </label>
              <div className="import-actions">
                <button type="button" className="button button-quiet" onClick={() => setPendingCSV(null)}>Cancel</button>
                <button type="submit" className="button button-primary" autoFocus>Import</button>
              </div>
            </form>
          </div>
  )

  if (!workbook || !sheet) {
    return (
      <div className="app-shell">
        <main id="main-content" className="main-content" tabIndex="-1">
          {error ? <p role="alert">{error}</p> : <p role="status">Loading…</p>}
        </main>
        {importDialog}
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
          <input ref={inputRef} type="file" accept=".csvx,.csv,application/zip,text/csv" onChange={handleOpenInputChange} className="sr-only" aria-label="Open CSVX package or CSV file" />
          <button type="button" className="button button-quiet" onClick={handleOpenClick}>Open</button>
          <button type="button" className="button button-quiet" onClick={() => setPrintOpen(true)} title="Print (Cmd/Ctrl+P)">Print…</button>
          <button type="button" className="button button-quiet" onClick={handleSaveAs} title="Save a copy to a new file">Save As…</button>
          <button type="button" className="button button-primary" onClick={handleSave} title="Save (Cmd/Ctrl+S)">Save</button>
        </nav>
      </header>
      {importDialog}
      <main id="main-content" className="main-content" tabIndex="-1">
        {error && <p role="alert">{error}</p>}
        {importNotice && <p role="status" className="import-notice">{importNotice}</p>}
        <section className="formula-panel" aria-label="Cell inspector">
          <div className="name-box">{selectedCoordinate}</div>
          <div className="formula-symbol" aria-hidden="true">fx</div>
          <div className="formula-value">
            <input
              key={`${activeSheetId}:${selectedCoordinate}:${editRevision}`}
              ref={formulaBarRef}
              type="text"
              className="formula-input"
              defaultValue={selectedMetadata?.formula ?? selectedRawValue}
              placeholder="Blank cell"
              onFocus={() => {
                if (!(editingCell === selectedCoordinate && editingSource === 'formula-bar')) {
                  beginEdit(selectedCell.row, selectedCell.column, undefined, 'formula-bar')
                }
              }}
              onBlur={() => {
                if (suppressFormulaBarBlurRef.current) {
                  suppressFormulaBarBlurRef.current = false
                  return
                }
                commitEdit()
              }}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  e.preventDefault()
                  suppressFormulaBarBlurRef.current = true
                  commitEdit()
                  e.currentTarget.blur()
                } else if (e.key === 'Escape') {
                  e.preventDefault()
                  suppressFormulaBarBlurRef.current = true
                  cancelEdit()
                  e.currentTarget.blur()
                }
              }}
              aria-label={`Formula for ${selectedCoordinate}`}
            />
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
              <span aria-hidden="true" className="color-swatch-border" style={{ borderColor: borderColorOf(selectedStyle) || 'var(--border-strong)' }} />
              <input type="color" value={(borderColorOf(selectedStyle) || '#000000').toLowerCase()} onChange={(e) => setBorderColor(e.target.value)} aria-label="Border color" />
            </label>
            <button type="button" className={`format-button ${allSelectedAreCurrency ? 'is-active' : ''}`} aria-pressed={allSelectedAreCurrency} onMouseDown={(e) => e.preventDefault()} onClick={toggleCurrency} aria-label="Currency (USD)" title="Format as currency (USD)">$</button>
            <button type="button" className="format-button format-clear" onMouseDown={(e) => e.preventDefault()} onClick={clearFormatting} aria-label="Clear formatting" title="Clear formatting">Clear</button>
            <span className="format-divider" aria-hidden="true" />
            <button type="button" className="format-button format-case" onMouseDown={(e) => e.preventDefault()} onClick={() => changeTextCase('upper')} aria-label="Uppercase" title="Uppercase">AA</button>
            <button type="button" className="format-button format-case" onMouseDown={(e) => e.preventDefault()} onClick={() => changeTextCase('title')} aria-label="Capitalize each word" title="First letter of each word">Aa</button>
            <button type="button" className="format-button format-case" onMouseDown={(e) => e.preventDefault()} onClick={() => changeTextCase('lower')} aria-label="Lowercase" title="Lowercase">aa</button>
            <span className="format-divider" aria-hidden="true" />
            <button type="button" className={`format-button ${allSelectedHaveAlignment('left') ? 'is-active' : ''}`} aria-pressed={allSelectedHaveAlignment('left')} onMouseDown={(e) => e.preventDefault()} onClick={() => setAlignment('left')} aria-label="Align left" title="Align left"><AlignIcon variant="left" /></button>
            <button type="button" className={`format-button ${allSelectedHaveAlignment('center') ? 'is-active' : ''}`} aria-pressed={allSelectedHaveAlignment('center')} onMouseDown={(e) => e.preventDefault()} onClick={() => setAlignment('center')} aria-label="Align center" title="Align center"><AlignIcon variant="center" /></button>
            <button type="button" className={`format-button ${allSelectedHaveAlignment('right') ? 'is-active' : ''}`} aria-pressed={allSelectedHaveAlignment('right')} onMouseDown={(e) => e.preventDefault()} onClick={() => setAlignment('right')} aria-label="Align right" title="Align right"><AlignIcon variant="right" /></button>
          </div>
        </section>
        <section className="grid-card" aria-labelledby="grid-title">
          <h3 id="grid-title" className="sr-only">{sheet.name} spreadsheet data</h3>
          <div className="table-scroll" ref={tableScrollRef} onScroll={(e) => setScrollTop(e.currentTarget.scrollTop)}>
            <table className="spreadsheet">
              <caption className="sr-only">CSV-backed data in {sheet.name}. Press Enter, Space, or F2 to edit a cell.</caption>
              <colgroup>
                <col style={{ width: '3rem' }} />
                {sheet.columns.slice(0, columnCount).map((column, index) => (
                  <col key={column.id || index} ref={(el) => { if (el) colRefs.current.set(index, el); else colRefs.current.delete(index) }} style={{ width: `${pixelWidthForColumn(column)}px` }} />
                ))}
              </colgroup>
              <thead>
                <tr>
                  <th scope="col" className="corner-cell">
                    <button type="button" className="corner-button" onClick={selectAll} aria-label="Select all cells" title="Select all" />
                  </th>
                  {sheet.columns.map((column, index) => (
                    <th scope="col" key={column.id || index} data-col={index} style={cellBorderCSS(styleAt(HEADER_ROW, index), {}).borderTop ? { borderBottom: cellBorderCSS(styleAt(HEADER_ROW, index), {}).borderTop } : undefined} onContextMenu={(e) => openColumnContextMenu(e, index)}>
                      <button
                        type="button"
                        className={`column-header-button ${selectedColumns.has(index) ? 'is-active' : ''}`}
                        onMouseDown={(e) => startDrag('column', { row: HEADER_ROW, column: index }, e)}
                        onClick={(e) => { if (!suppressClickRef.current) selectColumn(index, e.metaKey || e.ctrlKey) }}
                        aria-label={`Select column ${column.name || column.id}`}
                        title={column.width ? `Width: ${column.width} (XLSX character-width units)` : undefined}
                      >
                        {columnId(index)}
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
                {topSpacerHeight > 0 && (
                  <tr aria-hidden="true" style={{ height: topSpacerHeight }}>
                    <td colSpan={sheet.columns.length + 2} style={{ padding: 0, border: 0 }} />
                  </tr>
                )}
                {Array.from({ length: endRow - startRow }, (_, offset) => startRow + offset - 1).map((rowIndex) => {
                  return (
                  <tr key={rowIndex}>
                    <th scope="row" data-row={rowIndex} style={cellBorderCSS(styleAt(rowIndex, 0), {}).borderLeft ? { borderRight: cellBorderCSS(styleAt(rowIndex, 0), {}).borderLeft } : undefined} onContextMenu={(e) => openRowContextMenu(e, rowIndex)}>
                      <button type="button" className={`row-header-button ${selectedRows.has(rowIndex) ? 'is-active' : ''}`} onMouseDown={(e) => startDrag('row', { row: rowIndex, column: 0 }, e)} onClick={(e) => { if (!suppressClickRef.current) selectRow(rowIndex, e.metaKey || e.ctrlKey) }} aria-label={`Select row ${rowNumberFor(rowIndex)}`}>
                        {rowNumberFor(rowIndex)}
                      </button>
                    </th>
                    {sheet.columns.slice(0, columnCount).map((_, columnIndex) => {
                      const coordinate = coordinateFor(columnIndex, rowIndex)
                      const value = rawCellText(sheet, rowIndex, columnIndex)
                      const metadata = cellMetadata(sheet, coordinate)
                      const style = styleFor(metadata?.style)
                      // editingSource gates this: while editing via the formula bar, the grid cell
                      // stays in display mode so its autoFocus input doesn't steal focus away.
                      const isEditing = editingCell === coordinate && editingSource === 'grid'
                      // formatValue is presentation-only (spec/08-styles.md) — it never changes what's
                      // stored, only how a formula's cached result or a plain typed value is shown.
                      const declaredType = metadata?.type || (rowIndex < 0 ? 'string' : sheet.columns[columnIndex]?.type)
                      const cellValue = metadata?.cached ?? resolveCellValue(value, declaredType, style.numberFormat)
                      const displayValue = formatValue(cellValue, style.numberFormat)
                      return (
                        <td key={coordinate} data-row={rowIndex} data-col={columnIndex} style={cellBorderCSS(style, neighborStyles(rowIndex, columnIndex))}>
                          {isEditing ? (
                            <input
                              ref={editorRef}
                              autoFocus
                              type="text"
                              className="cell-editor"
                              style={cellCSS(style)}
                              defaultValue={draftValue}
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
                              onMouseDown={(e) => startDrag('cell', { row: rowIndex, column: columnIndex }, e)}
                              onClick={(e) => {
                                if (e.metaKey || e.ctrlKey) toggleCellSelection(rowIndex, columnIndex)
                                else if (!suppressClickRef.current) selectCell(rowIndex, columnIndex)
                              }}
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
                  )
                })}
                {bottomSpacerHeight > 0 && (
                  <tr aria-hidden="true" style={{ height: bottomSpacerHeight }}>
                    <td colSpan={sheet.columns.length + 2} style={{ padding: 0, border: 0 }} />
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </section>
      </main>
      {printOpen && (
        <PrintView
          sheet={sheet}
          styles={workbook.styles}
          print={sheet.print}
          columnWidthPx={pixelWidthForColumn}
          styleAt={(row, column) => styleAt(row, column)}
          displayAt={displayAt}
          onChange={(patch) => setWorkbook(setPrintSettings(workbook, sheet.id, patch))}
          onClose={() => setPrintOpen(false)}
        />
      )}
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
          <strong>{sheet.records.length + 1}</strong>&nbsp;rows · <strong>{sheet.columns.length}</strong>&nbsp;columns
        </div>
        <span className="footer-hint">Enter/F2 edit · Right-click a sheet tab for actions</span>
      </footer>
      {contextMenu && contextMenu.kind === 'sheet' && (
        <div ref={contextMenuRef} className="row-context-menu" role="menu" style={{ left: contextMenu.x, top: contextMenu.y }} onClick={(e) => e.stopPropagation()}>
          <p className="context-menu-label">{workbook.sheets.find((item) => item.id === contextMenu.index)?.name}</p>
          <button type="button" role="menuitem" onClick={() => beginRenameSheet(contextMenu.index, workbook.sheets.find((item) => item.id === contextMenu.index)?.name || '')}>Rename sheet</button>
          <button type="button" role="menuitem" className="danger-action" onClick={() => removeSheet(contextMenu.index)}>Delete sheet</button>
        </div>
      )}
      {contextMenu && contextMenu.kind === 'row' && (() => {
        const rows = contextMenu.items
        const plural = rows.length > 1
        return (
          <div ref={contextMenuRef} className="row-context-menu" role="menu" style={{ left: contextMenu.x, top: contextMenu.y }} onClick={(e) => e.stopPropagation()}>
            <p className="context-menu-label">{plural ? `Rows ${rowNumberFor(rows[0])}–${rowNumberFor(rows[rows.length - 1])}` : `Row ${rowNumberFor(rows[0])}`}</p>
            <button type="button" role="menuitem" onClick={() => handleInsertRows(rows[0], rows.length)}>Insert {plural ? `${rows.length} rows` : 'row'} before</button>
            <button type="button" role="menuitem" onClick={() => handleInsertRows(rows[rows.length - 1] + 1, rows.length)}>Insert {plural ? `${rows.length} rows` : 'row'} after</button>
            <button type="button" role="menuitem" className="danger-action" onClick={() => handleDeleteRows(rows)}>Delete {plural ? `${rows.length} rows` : 'row'}</button>
          </div>
        )
      })()}
      {contextMenu && contextMenu.kind === 'column' && (() => {
        const columns = contextMenu.items
        const plural = columns.length > 1
        const label = (index) => columnId(index)
        return (
          <div ref={contextMenuRef} className="row-context-menu" role="menu" style={{ left: contextMenu.x, top: contextMenu.y }} onClick={(e) => e.stopPropagation()}>
            <p className="context-menu-label">{plural ? `Columns ${label(columns[0])}–${label(columns[columns.length - 1])}` : `Column ${label(columns[0])}`}</p>
            <button type="button" role="menuitem" onClick={() => handleInsertColumns(columns[0], columns.length)}>Insert {plural ? `${columns.length} columns` : 'column'} before</button>
            <button type="button" role="menuitem" onClick={() => handleInsertColumns(columns[columns.length - 1] + 1, columns.length)}>Insert {plural ? `${columns.length} columns` : 'column'} after</button>
            <button type="button" role="menuitem" className="danger-action" onClick={() => handleDeleteColumns(columns)}>Delete {plural ? `${columns.length} columns` : 'column'}</button>
          </div>
        )
      })()}
    </div>
  )
}

export default App
