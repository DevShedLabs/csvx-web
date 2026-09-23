import { useMemo, useState } from 'react'

const demoWorkbook = {
  name: 'example.csvx',
  version: '1.0',
  sheets: [
    {
      id: 'sheet-1',
      name: 'Sheet 1',
      columns: ['Value'],
      rows: [['1'], ['2']],
    },
  ],
}

function App() {
  const [activeSheet, setActiveSheet] = useState(demoWorkbook.sheets[0].id)
  const [selectedCell, setSelectedCell] = useState('A1')

  const sheet = useMemo(
    () => demoWorkbook.sheets.find((item) => item.id === activeSheet),
    [activeSheet],
  )

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
          <span>{demoWorkbook.name}</span>
          <span className="muted">Core {demoWorkbook.version}</span>
        </div>
        <nav className="top-actions" aria-label="File actions">
          <button type="button" className="button button-quiet">Open</button>
          <button type="button" className="button button-primary">Export</button>
        </nav>
      </header>

      <div className="workspace">
        <aside className="sidebar" aria-label="Workbook navigation">
          <section className="sidebar-section">
            <div className="section-heading">
              <h2>Workbook</h2>
              <span className="count">{demoWorkbook.sheets.length}</span>
            </div>
            <ul className="sheet-list">
              {demoWorkbook.sheets.map((item) => (
                <li key={item.id}>
                  <button
                    type="button"
                    className={`sheet-item ${item.id === activeSheet ? 'is-active' : ''}`}
                    aria-current={item.id === activeSheet ? 'page' : undefined}
                    onClick={() => setActiveSheet(item.id)}
                  >
                    <span className="sheet-icon" aria-hidden="true">▦</span>
                    {item.name}
                  </button>
                </li>
              ))}
            </ul>
          </section>
          <section className="sidebar-section package-section">
            <div className="section-heading"><h2>Package</h2></div>
            <dl className="package-list">
              <div><dt>Format</dt><dd>CSVX</dd></div>
              <div><dt>Data layer</dt><dd>CSV</dd></div>
              <div><dt>Metadata</dt><dd>None</dd></div>
            </dl>
          </section>
        </aside>

        <main id="main-content" className="main-content" tabIndex="-1">
          <div className="content-header">
            <div>
              <p className="eyebrow">Sheet / {sheet.name}</p>
              <h2>{sheet.name}</h2>
            </div>
            <span className="read-only-badge">Read-only demo</span>
          </div>

          <section className="formula-panel" aria-label="Cell inspector">
            <div className="name-box" aria-label="Selected cell">{selectedCell}</div>
            <div className="formula-symbol" aria-hidden="true">fx</div>
            <div className="formula-value">Select a cell to inspect its CSV value</div>
          </section>

          <section className="grid-card" aria-labelledby="grid-title">
            <h3 id="grid-title" className="sr-only">{sheet.name} spreadsheet data</h3>
            <div className="table-scroll">
              <table className="spreadsheet">
                <caption className="sr-only">CSV-backed data in {sheet.name}</caption>
                <thead>
                  <tr>
                    <th scope="col" className="corner-cell" aria-label="Spreadsheet corner" />
                    {sheet.columns.map((column, index) => (
                      <th scope="col" key={column}>{String.fromCharCode(65 + index)}<span className="sr-only">: {column}</span></th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {sheet.rows.map((row, rowIndex) => (
                    <tr key={`${sheet.id}-${rowIndex}`}>
                      <th scope="row">{rowIndex + 1}</th>
                      {row.map((value, columnIndex) => {
                        const coordinate = `${String.fromCharCode(65 + columnIndex)}${rowIndex + 1}`
                        return (
                          <td key={coordinate}>
                            <button
                              type="button"
                              className={`cell-button ${selectedCell === coordinate ? 'is-selected' : ''}`}
                              onClick={() => setSelectedCell(coordinate)}
                              aria-label={`${coordinate}, value ${value}`}
                            >
                              {value}
                            </button>
                          </td>
                        )
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>

          <div className="status-bar" role="status">
            <span><strong>{sheet.rows.length}</strong> data rows</span>
            <span><strong>{sheet.columns.length}</strong> column</span>
            <span className="status-spacer" />
            <span>CSV data layer · No formulas</span>
          </div>
        </main>
      </div>
    </div>
  )
}

export default App
