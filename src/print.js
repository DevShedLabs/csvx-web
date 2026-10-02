// Pagination for the Print view. This is layout, not CSVX semantics: the spec (03-sheets.md, "Print
// settings") defines the settings and says pagination is the renderer's job, and this is that job.
// It is deliberately free of React and the DOM so it can move into an engine library later.
import { rowHeightToPixels } from 'csvx-ts/browser'
import { usedRange } from './model.js'

const DPI = 96
export const DEFAULT_ROW_PX = 28

/** Paper sizes in inches, portrait (width, height). */
export const PAPER_INCHES = { letter: [8.5, 11], legal: [8.5, 14], tabloid: [11, 17], a3: [11.69, 16.54], a4: [8.27, 11.69], a5: [5.83, 8.27] }

/** The defaults the spec declares for an absent property. */
export const PRINT_DEFAULTS = {
  orientation: 'portrait',
  paperSize: 'letter',
  margins: { top: 0.75, right: 0.7, bottom: 0.75, left: 0.7 },
  scale: 100,
  pageOrder: 'downThenOver',
  gridlines: false,
  centerHorizontally: false,
}

export function columnIndexFromLetters(letters) {
  let index = 0
  for (const char of letters) index = index * 26 + (char.charCodeAt(0) - 64)
  return index - 1
}

/** 'A1:H20' → zero-based inclusive bounds, or null when malformed. */
export function parseRange(text) {
  const match = /^([A-Z]+)([1-9][0-9]*):([A-Z]+)([1-9][0-9]*)$/.exec(text || '')
  if (!match) return null
  const [c0, c1] = [columnIndexFromLetters(match[1]), columnIndexFromLetters(match[3])]
  const [r0, r1] = [Number(match[2]) - 1, Number(match[4]) - 1]
  return { r0: Math.min(r0, r1), r1: Math.max(r0, r1), c0: Math.min(c0, c1), c1: Math.max(c0, c1) }
}

/** '1:2' → [0, 1]; 'A:B' → [0, 1]; null when malformed. */
export function parseRepeat(text, kind) {
  const match = (kind === 'rows' ? /^([1-9][0-9]*):([1-9][0-9]*)$/ : /^([A-Z]+):([A-Z]+)$/).exec(text || '')
  if (!match) return null
  const [a, b] = kind === 'rows' ? [Number(match[1]) - 1, Number(match[2]) - 1] : [columnIndexFromLetters(match[1]), columnIndexFromLetters(match[2])]
  return [Math.min(a, b), Math.max(a, b)]
}

export function effectiveSettings(print) {
  const merged = { ...PRINT_DEFAULTS, ...print }
  merged.margins = { ...PRINT_DEFAULTS.margins, ...print?.margins }
  return merged
}

/** Paper size with orientation applied, and the printable area inside the margins, in CSS pixels. */
export function pageGeometry(settings) {
  const [w, h] = PAPER_INCHES[settings.paperSize] || PAPER_INCHES.letter
  const [paperW, paperH] = settings.orientation === 'landscape' ? [h, w] : [w, h]
  const { top, right, bottom, left } = settings.margins
  return {
    width: paperW * DPI,
    height: paperH * DPI,
    margin: { top: top * DPI, right: right * DPI, bottom: bottom * DPI, left: left * DPI },
    printableWidth: Math.max(1, (paperW - left - right) * DPI),
    printableHeight: Math.max(1, (paperH - top - bottom) * DPI),
  }
}

/** Splits `indices` into groups that fit `limit` pixels, honoring forced breaks (1-based "break
 * after item n") and prepending the repeated items to every group that doesn't already include
 * them. Always puts at least one item in a group, so an oversized item can't loop forever. */
export function splitIntoGroups(indices, sizeOf, limit, breaksAfter, repeat) {
  const repeatSize = repeat ? repeat.indices.reduce((sum, index) => sum + sizeOf(index), 0) : 0
  const overheadFor = (first) => (repeat && first > repeat.end ? repeatSize : 0)
  const groups = []
  let current = []
  let used = 0
  for (const index of indices) {
    const size = sizeOf(index)
    if (current.length === 0) used = overheadFor(index)
    if (current.length > 0 && used + size > limit) {
      groups.push(current)
      current = []
      used = overheadFor(index)
    }
    current.push(index)
    used += size
    if (breaksAfter.has(index + 1)) {
      groups.push(current)
      current = []
    }
  }
  if (current.length > 0) groups.push(current)
  return groups
}

function span(from, to) {
  return Array.from({ length: Math.max(0, to - from + 1) }, (_, offset) => from + offset)
}

/** Lays the sheet's print area out as pages. `columnWidthPx(column)` and `rowHeightPx(row)` give
 * unscaled pixel sizes. Returns { settings, geometry, scale, area, pages } where each page lists
 * the zero-based `rows` and `columns` it shows (repeat rows/columns already included). */
export function paginate({ sheet, styles, print, columnWidthPx, rowHeightPx }) {
  const settings = effectiveSettings(print)
  const geometry = pageGeometry(settings)
  const used = usedRange(sheet, styles)
  const declared = parseRange(settings.area)
  const area = declared
    ? { r0: declared.r0, r1: Math.min(declared.r1, sheet.records.length), c0: declared.c0, c1: Math.min(declared.c1, sheet.columns.length - 1) }
    : { r0: 0, r1: used.rows - 1, c0: 0, c1: used.columns - 1 }
  const rowIndices = span(area.r0, area.r1)
  const columnIndices = span(area.c0, area.c1)

  const repeatRowBounds = parseRepeat(settings.repeatRows, 'rows')
  const repeatColumnBounds = parseRepeat(settings.repeatColumns, 'columns')
  const repeatRows = repeatRowBounds ? { start: repeatRowBounds[0], end: repeatRowBounds[1], indices: span(...repeatRowBounds) } : null
  const repeatColumns = repeatColumnBounds ? { start: repeatColumnBounds[0], end: repeatColumnBounds[1], indices: span(...repeatColumnBounds) } : null

  const totalWidth = columnIndices.reduce((sum, index) => sum + columnWidthPx(index), 0)
  const totalHeight = rowIndices.reduce((sum, index) => sum + rowHeightPx(index), 0)
  const fitW = settings.fitToWidth > 0 ? settings.fitToWidth : null
  const fitH = settings.fitToHeight > 0 ? settings.fitToHeight : null
  let scale = settings.scale / 100
  if (fitW || fitH) {
    // Fit only ever shrinks, as in XLSX. Repeated columns/rows cost space on every extra page.
    const repeatW = repeatColumns ? repeatColumns.indices.reduce((sum, index) => sum + columnWidthPx(index), 0) : 0
    const repeatH = repeatRows ? repeatRows.indices.reduce((sum, index) => sum + rowHeightPx(index), 0) : 0
    const candidates = [1]
    if (fitW) candidates.push((fitW * geometry.printableWidth) / (totalWidth + (fitW - 1) * repeatW))
    if (fitH) candidates.push((fitH * geometry.printableHeight) / (totalHeight + (fitH - 1) * repeatH))
    scale = Math.min(...candidates)
  }
  scale = Math.min(4, Math.max(0.1, scale))

  const breaks = (list) => new Set(list || [])
  const columnGroups = splitIntoGroups(columnIndices, columnWidthPx, geometry.printableWidth / scale, breaks(settings.columnBreaks), repeatColumns && { ...repeatColumns, end: repeatColumns.end })
  const rowGroups = splitIntoGroups(rowIndices, rowHeightPx, geometry.printableHeight / scale, breaks(settings.rowBreaks), repeatRows && { ...repeatRows, end: repeatRows.end })

  const withRepeat = (group, repeat) => (repeat && group[0] > repeat.end ? [...repeat.indices, ...group] : group)
  const pages = []
  const outer = settings.pageOrder === 'overThenDown' ? rowGroups : columnGroups
  const inner = settings.pageOrder === 'overThenDown' ? columnGroups : rowGroups
  for (const outerGroup of outer) {
    for (const innerGroup of inner) {
      const columnGroup = settings.pageOrder === 'overThenDown' ? innerGroup : outerGroup
      const rowGroup = settings.pageOrder === 'overThenDown' ? outerGroup : innerGroup
      pages.push({ number: pages.length + 1, columns: withRepeat(columnGroup, repeatColumns), rows: withRepeat(rowGroup, repeatRows) })
    }
  }
  return { settings, geometry, scale, area, pages }
}

/** Row height in unscaled pixels: the sheet's declared height (points) if any, else the grid's. */
export function rowHeightPxFor(sheet) {
  return (row) => {
    const points = sheet.rowHeights?.[row + 1]
    return typeof points === 'number' && points > 0 ? rowHeightToPixels(points) : DEFAULT_ROW_PX
  }
}
