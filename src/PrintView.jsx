import { useEffect, useMemo, useState } from 'react'
import { cellBorderCSS, cellCSS } from './model.js'
import { PAPER_INCHES, PRINT_DEFAULTS, effectiveSettings, paginate, parseRange, parseRepeat, rowHeightPxFor } from './print.js'

// Google-Sheets-style print view: a settings panel and a preview of the actual pages. What the
// preview shows is exactly what window.print() sends — see the @media print rules in styles.css.
// Every setting is a property of the sheet's `print` object (spec/03-sheets.md), so it is saved
// with the workbook and round-trips.

const MARGIN_PRESETS = {
  normal: { label: 'Normal', margins: undefined },
  narrow: { label: 'Narrow', margins: { top: 0.25, right: 0.25, bottom: 0.25, left: 0.25 } },
  wide: { label: 'Wide', margins: { top: 1, right: 1, bottom: 1, left: 1 } },
}
const PAPER_LABELS = { letter: 'Letter', legal: 'Legal', tabloid: 'Tabloid', a3: 'A3', a4: 'A4', a5: 'A5' }

function scaleMode(print) {
  const width = print?.fitToWidth > 0
  const height = print?.fitToHeight > 0
  if (width && height) return 'page'
  if (width) return 'width'
  if (height) return 'height'
  return 'custom'
}

function marginPreset(print) {
  const margins = print?.margins
  if (!margins) return 'normal'
  const same = (a, b) => ['top', 'right', 'bottom', 'left'].every((side) => (a[side] ?? PRINT_DEFAULTS.margins[side]) === (b[side] ?? PRINT_DEFAULTS.margins[side]))
  const match = Object.entries(MARGIN_PRESETS).find(([, preset]) => same(margins, preset.margins || {}))
  return match ? match[0] : 'custom'
}

/** A text setting that is only saved once it matches the spec's pattern; empty clears it. */
function PatternField({ label, value, placeholder, parse, onCommit }) {
  const [draft, setDraft] = useState(value || '')
  const [invalid, setInvalid] = useState(false)
  useEffect(() => { setDraft(value || ''); setInvalid(false) }, [value])
  function commit() {
    const text = draft.trim().toUpperCase()
    if (text === '') { setInvalid(false); onCommit(undefined); return }
    if (!parse(text)) { setInvalid(true); return }
    setInvalid(false)
    setDraft(text)
    onCommit(text)
  }
  return (
    <label className="print-field">
      <span>{label}</span>
      <input type="text" value={draft} placeholder={placeholder} aria-invalid={invalid} onChange={(e) => setDraft(e.target.value)} onBlur={commit} onKeyDown={(e) => e.key === 'Enter' && commit()} />
    </label>
  )
}

export default function PrintView({ sheet, styles, print, columnWidthPx, styleAt: styleAtRecord, displayAt: displayAtRecord, onChange, onClose }) {
  // Print works in sheet rows (0 = the header, spec row 1); the app addresses cells by record index.
  const styleAt = (row, column) => styleAtRecord(row - 1, column)
  const displayAt = (row, column) => displayAtRecord(row - 1, column)
  const layout = useMemo(
    () => paginate({ sheet, styles, print, columnWidthPx: (index) => columnWidthPx(sheet.columns[index]), rowHeightPx: rowHeightPxFor(sheet) }),
    [sheet, styles, print, columnWidthPx],
  )
  const { settings, geometry, scale, pages } = layout
  const rowHeightPx = rowHeightPxFor(sheet)

  useEffect(() => {
    const onKey = (event) => event.key === 'Escape' && onClose()
    document.addEventListener('keydown', onKey)
    return () => document.removeEventListener('keydown', onKey)
  }, [onClose])

  const mode = scaleMode(print)
  const fitNote = mode !== 'custom' ? `${Math.round(scale * 100)}%` : null

  function setScaleMode(next) {
    if (next === 'width') onChange({ fitToWidth: 1, fitToHeight: 0, scale: undefined })
    else if (next === 'height') onChange({ fitToWidth: 0, fitToHeight: 1, scale: undefined })
    else if (next === 'page') onChange({ fitToWidth: 1, fitToHeight: 1, scale: undefined })
    else onChange({ fitToWidth: undefined, fitToHeight: undefined, scale: print?.scale ?? 100 })
  }

  const [paperW, paperH] = [geometry.width / 96, geometry.height / 96]
  const pageNode = (page) => {
    const contentWidth = page.columns.reduce((sum, index) => sum + columnWidthPx(sheet.columns[index]), 0)
    const centered = settings.centerHorizontally ? Math.max(0, (geometry.printableWidth - contentWidth * scale) / 2) : 0
    return (
      <div className="print-page-wrap" key={page.number}>
        <section className="print-page" aria-label={`Page ${page.number} of ${pages.length}`} style={{ width: geometry.width, height: geometry.height, padding: `${geometry.margin.top}px ${geometry.margin.right}px ${geometry.margin.bottom}px ${geometry.margin.left}px` }}>
          <div className="print-page-body" style={{ width: geometry.printableWidth, height: geometry.printableHeight }}>
            <div style={{ transform: `scale(${scale})`, transformOrigin: 'top left', width: contentWidth, marginLeft: centered }}>
              <table className={`print-table ${settings.gridlines ? 'has-gridlines' : ''}`} style={{ width: contentWidth }}>
                <colgroup>{page.columns.map((index) => <col key={index} style={{ width: columnWidthPx(sheet.columns[index]) }} />)}</colgroup>
                <tbody>
                  {page.rows.map((row) => (
                    <tr key={row} style={{ height: rowHeightPx(row) }}>
                      {page.columns.map((column) => {
                        const style = styleAt(row, column)
                        const neighbors = { above: styleAt(row - 1, column), below: styleAt(row + 1, column), left: styleAt(row, column - 1), right: styleAt(row, column + 1) }
                        return <td key={column} style={{ ...cellCSS(style), ...cellBorderCSS(style, neighbors) }}>{displayAt(row, column)}</td>
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </section>
        <p className="print-page-label">Page {page.number} of {pages.length}</p>
      </div>
    )
  }

  return (
    <div className="print-overlay" role="dialog" aria-modal="true" aria-label="Print">
      <style>{`@page { size: ${paperW}in ${paperH}in; margin: 0 }`}</style>
      <aside className="print-panel">
        <h2>Print</h2>
        <p className="print-summary">{sheet.name} · {pages.length} {pages.length === 1 ? 'page' : 'pages'}{fitNote ? ` · ${fitNote}` : ''}</p>
        <div className="print-actions">
          <button type="button" className="button button-primary" onClick={() => window.print()}>Print</button>
          <button type="button" className="button button-quiet" onClick={onClose}>Close</button>
        </div>

        <label className="print-field"><span>Orientation</span>
          <select value={settings.orientation} onChange={(e) => onChange({ orientation: e.target.value })}>
            <option value="portrait">Portrait</option><option value="landscape">Landscape</option>
          </select>
        </label>
        <label className="print-field"><span>Paper size</span>
          <select value={settings.paperSize} onChange={(e) => onChange({ paperSize: e.target.value })}>
            {Object.keys(PAPER_INCHES).map((size) => <option key={size} value={size}>{PAPER_LABELS[size]}</option>)}
          </select>
        </label>
        <label className="print-field"><span>Scale</span>
          <select value={mode} onChange={(e) => setScaleMode(e.target.value)}>
            <option value="custom">Custom</option><option value="width">Fit to width</option><option value="height">Fit to height</option><option value="page">Fit to page</option>
          </select>
        </label>
        {mode === 'custom' && (
          <label className="print-field"><span>Scale %</span>
            <input type="number" min="10" max="400" value={print?.scale ?? 100} onChange={(e) => { const value = Number(e.target.value); if (value >= 10 && value <= 400) onChange({ scale: value === 100 ? undefined : value }) }} />
          </label>
        )}
        <label className="print-field"><span>Margins</span>
          <select value={marginPreset(print)} onChange={(e) => onChange({ margins: MARGIN_PRESETS[e.target.value]?.margins })}>
            {Object.entries(MARGIN_PRESETS).map(([key, preset]) => <option key={key} value={key}>{preset.label}</option>)}
            {marginPreset(print) === 'custom' && <option value="custom">Custom</option>}
          </select>
        </label>
        <label className="print-field"><span>Page order</span>
          <select value={settings.pageOrder} onChange={(e) => onChange({ pageOrder: e.target.value === 'downThenOver' ? undefined : e.target.value })}>
            <option value="downThenOver">Down, then over</option><option value="overThenDown">Over, then down</option>
          </select>
        </label>
        <label className="print-check"><input type="checkbox" checked={!!settings.gridlines} onChange={(e) => onChange({ gridlines: e.target.checked || undefined })} /> Show gridlines</label>
        <label className="print-check"><input type="checkbox" checked={!!settings.centerHorizontally} onChange={(e) => onChange({ centerHorizontally: e.target.checked || undefined })} /> Center horizontally</label>

        <PatternField label="Print area" value={print?.area} placeholder="Used range, e.g. A1:H20" parse={parseRange} onCommit={(value) => onChange({ area: value })} />
        <PatternField label="Repeat rows" value={print?.repeatRows} placeholder="e.g. 1:1" parse={(text) => parseRepeat(text, 'rows')} onCommit={(value) => onChange({ repeatRows: value })} />
        <PatternField label="Repeat columns" value={print?.repeatColumns} placeholder="e.g. A:A" parse={(text) => parseRepeat(text, 'columns')} onCommit={(value) => onChange({ repeatColumns: value })} />
      </aside>
      <div className="print-preview">{pages.map(pageNode)}</div>
    </div>
  )
}
