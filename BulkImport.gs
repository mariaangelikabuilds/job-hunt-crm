/**
 * BulkImport.gs
 *
 * Modal dialog that takes pasted CSV and bulk-appends rows to the
 * Applications tab. Optionally bulk-scores them with rate limiting (one
 * call per ~3 seconds to stay well under the UrlFetchApp daily quota).
 *
 * CSV column order expected:
 *   Company, Role, JD Link, Date Applied, Status, Tags
 *
 * Headers row is detected and skipped if present.
 * Quoted fields are supported. Embedded commas inside quotes are handled.
 */

function showBulkImportDialog_() {
  const html = HtmlService.createTemplateFromFile('BulkImportDialog')
    .evaluate()
    .setWidth(560)
    .setHeight(540)
  SpreadsheetApp.getUi().showModalDialog(html, 'Import applications from CSV')
}

/**
 * Parses CSV text and appends rows. Returns counts and any failures.
 */
function bulkImportCsv(csvText, options) {
  const opts = options || {}
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_APPLICATIONS)
  const cols = getColumnMap_(sheet)

  const rows = parseCsv_(csvText)
  if (rows.length === 0) return { imported: 0, failed: 0, errors: [] }

  if (looksLikeHeaderRow_(rows[0])) rows.shift()

  const errors = []
  const importedRowIndices = []
  const lock = LockService.getDocumentLock()

  try {
    lock.waitLock(15000)
    for (let i = 0; i < rows.length; i++) {
      const fields = rows[i]
      const company = (fields[0] || '').trim()
      const role = (fields[1] || '').trim()
      const jdLink = (fields[2] || '').trim()
      const dateApplied = parseDateInput_(fields[3])
      const status = (fields[4] || 'Saved').trim()
      const tags = (fields[5] || '').trim()

      if (!company) {
        errors.push('Row ' + (i + 1) + ' skipped: no Company.')
        continue
      }

      const newRow = sheet.getLastRow() + 1
      sheet.getRange(newRow, cols['Company']).setValue(company)
      sheet.getRange(newRow, cols['Role']).setValue(role)
      sheet.getRange(newRow, cols['JD Link']).setValue(jdLink)
      if (dateApplied) sheet.getRange(newRow, cols['Date Applied']).setValue(dateApplied)
      sheet.getRange(newRow, cols['Status']).setValue(status)
      sheet.getRange(newRow, cols['Tags']).setValue(tags)
      sheet.getRange(newRow, cols['Last Touch']).setValue(new Date())
      sheet.getRange(newRow, cols['Notes']).setValue('Bulk-imported on ' + new Date().toISOString())

      importedRowIndices.push(newRow)
    }
  } finally {
    try { lock.releaseLock() } catch (_) {}
  }

  logActivity_({
    action: 'bulk_import',
    rowRef: 'count:' + importedRowIndices.length,
    notes: 'imported=' + importedRowIndices.length + ', failed=' + errors.length
  })

  let scored = 0
  if (opts.scoreAfter && importedRowIndices.length > 0) {
    for (let i = 0; i < importedRowIndices.length; i++) {
      const r = importedRowIndices[i]
      try {
        scoreJD_(r)
        scored++
      } catch (err) {
        errors.push('Row ' + r + ' score failed: ' + err.message)
      }
      Utilities.sleep(3000)
    }
  }

  return {
    imported: importedRowIndices.length,
    failed: errors.length,
    scored: scored,
    errors: errors.slice(0, 10)
  }
}

function looksLikeHeaderRow_(fields) {
  if (!fields || fields.length === 0) return false
  const first = String(fields[0] || '').toLowerCase().trim()
  return first === 'company'
}

function parseDateInput_(input) {
  if (!input) return null
  const s = String(input).trim()
  if (!s) return null
  const d = new Date(s)
  if (isNaN(d.getTime())) return null
  return d
}

/**
 * Minimal CSV parser. Handles quoted fields with embedded commas, escaped
 * quotes ("") inside quotes, and CRLF line endings. Not a full RFC 4180
 * implementation; sufficient for spreadsheet exports.
 */
function parseCsv_(text) {
  const rows = []
  let current = []
  let field = ''
  let inQuotes = false
  let i = 0
  const len = text.length

  while (i < len) {
    const c = text.charAt(i)

    if (inQuotes) {
      if (c === '"') {
        if (text.charAt(i + 1) === '"') {
          field += '"'
          i += 2
          continue
        }
        inQuotes = false
        i++
        continue
      }
      field += c
      i++
      continue
    }

    if (c === '"') {
      inQuotes = true
      i++
      continue
    }
    if (c === ',') {
      current.push(field)
      field = ''
      i++
      continue
    }
    if (c === '\n' || c === '\r') {
      current.push(field)
      field = ''
      if (current.length === 1 && current[0] === '') {
        current = []
      } else {
        rows.push(current)
        current = []
      }
      if (c === '\r' && text.charAt(i + 1) === '\n') i++
      i++
      continue
    }
    field += c
    i++
  }

  if (field !== '' || current.length > 0) {
    current.push(field)
    rows.push(current)
  }

  return rows.filter(r => r.some(f => f && f.trim()))
}
