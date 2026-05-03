/**
 * Activity.gs
 *
 * Append-only audit log. Every Claude call, status change, email send, or
 * destructive action writes a row to the Activity tab.
 *
 * Cost calculation uses claude-opus-4-7 list pricing as of 2026-04:
 *   $5  per 1M input tokens (standard)
 *   $25 per 1M output tokens
 *   ~10% of input cost for cache reads (~$0.50 per 1M)
 *   125% of input cost for cache creation (~$6.25 per 1M, 5min TTL default)
 *
 * If pricing changes, update PRICING below. The model field on each entry
 * lets a future migration recompute costs retroactively.
 */

const PRICING = {
  'claude-opus-4-7': {
    input: 5.0 / 1e6,
    output: 25.0 / 1e6,
    cache_read: 0.5 / 1e6,
    cache_creation: 6.25 / 1e6
  },
  'claude-sonnet-4-6': {
    input: 3.0 / 1e6,
    output: 15.0 / 1e6,
    cache_read: 0.3 / 1e6,
    cache_creation: 3.75 / 1e6
  },
  'claude-haiku-4-5': {
    input: 1.0 / 1e6,
    output: 5.0 / 1e6,
    cache_read: 0.1 / 1e6,
    cache_creation: 1.25 / 1e6
  }
}

const ACTIVITY_NOTES_MAX_LENGTH = 500

function logActivity_(entry) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName(SHEET_ACTIVITY)
  if (!sheet) {
    console.warn('Activity sheet missing. Run setup() to recreate it.')
    return
  }

  const usage = entry.usage || {}
  const model = entry.model || 'claude-opus-4-7'
  const cost = computeCost_(model, usage)

  const row = [
    new Date(),
    entry.action || 'unknown',
    entry.rowRef || '',
    usage.input_tokens || 0,
    usage.cache_read_input_tokens || 0,
    usage.cache_creation_input_tokens || 0,
    usage.output_tokens || 0,
    cost,
    sanitizeActivityNotes_(entry.notes)
  ]

  const lock = LockService.getDocumentLock()
  try {
    lock.waitLock(5000)
    sheet.appendRow(row)
  } catch (err) {
    console.warn('Could not write activity row: ' + err.message)
  } finally {
    try { lock.releaseLock() } catch (_) {}
  }
}

/**
 * Notes are written into the Activity audit log. Strip obvious PII patterns
 * (emails, phone numbers) and cap length so the audit tab doesn't accumulate
 * long-form JD content that could re-leak if the tab is unhidden or shared.
 */
function sanitizeActivityNotes_(raw) {
  let s = String(raw == null ? '' : raw)
  s = s.replace(/[\w.+-]+@[\w-]+\.[\w.-]+/g, '[email]')
  s = s.replace(/(?:\+?\d[\d\s().-]{7,}\d)/g, '[phone]')
  if (s.length > ACTIVITY_NOTES_MAX_LENGTH) {
    s = s.slice(0, ACTIVITY_NOTES_MAX_LENGTH - 3) + '...'
  }
  return s
}

function computeCost_(model, usage) {
  const p = PRICING[model] || PRICING['claude-opus-4-7']
  const input = (usage.input_tokens || 0) * p.input
  const cacheRead = (usage.cache_read_input_tokens || 0) * p.cache_read
  const cacheCreation = (usage.cache_creation_input_tokens || 0) * p.cache_creation
  const output = (usage.output_tokens || 0) * p.output
  return input + cacheRead + cacheCreation + output
}

function getMonthlySpendUsd_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName(SHEET_ACTIVITY)
  if (!sheet || sheet.getLastRow() < 2) return 0

  const activityRows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues()
  const now = new Date()
  const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1)

  let total = 0
  for (const row of activityRows) {
    const timestamp = row[0]
    const cost = row[7]
    if (timestamp instanceof Date && timestamp >= startOfMonth && typeof cost === 'number') {
      total += cost
    }
  }
  return total
}

function getMonthlySpendThresholdUsd_() {
  const raw = PropertiesService.getScriptProperties().getProperty('MONTHLY_SPEND_THRESHOLD_USD')
  const parsed = parseFloat(raw)
  return Number.isFinite(parsed) && parsed > 0 ? parsed : 20
}

function getCacheHitRate_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName(SHEET_ACTIVITY)
  if (!sheet || sheet.getLastRow() < 2) return null

  const tokenRows = sheet.getRange(2, 4, sheet.getLastRow() - 1, 2).getValues()
  let input = 0
  let cacheRead = 0
  for (const row of tokenRows) {
    input += row[0] || 0
    cacheRead += row[1] || 0
  }
  if (input + cacheRead === 0) return null
  return cacheRead / (input + cacheRead)
}
