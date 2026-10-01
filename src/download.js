/** Triggers a browser download of in-memory bytes. Pure UI mechanics, no CSVX semantics. */
export function downloadBytes(bytes, filename) {
  const blob = new Blob([bytes], { type: 'application/zip' })
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = filename
  document.body.appendChild(link)
  link.click()
  link.remove()
  // Revoking immediately after click() races the browser actually reading the blob for the
  // download — some browsers (notably ones that hand the URL to a separate download process)
  // can end up saving a truncated or empty file. Deferring the revoke lets the read complete first.
  setTimeout(() => URL.revokeObjectURL(url), 1000)
}
