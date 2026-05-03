/**
 * Embeddings.gs
 *
 * Vector similarity for "find similar past applications". When a JD is
 * scored, its text is embedded via OpenAI text-embedding-3-small (1536
 * dims, $0.02 per 1M tokens) and stored as base64-packed Float32 in the
 * hidden Embedding column.
 *
 * "Find similar" computes dot product against every other row's embedding
 * (OpenAI v3 vectors are unit-normalized, so dot product equals cosine
 * similarity). Returns top-k matches above a threshold.
 *
 * If OPENAI_API_KEY is not set, all embedding writes are no-ops and the
 * find-similar action returns an empty list with a clear message.
 *
 * Storage cost: 1536 floats * 4 bytes = 6144 bytes per row, base64 ~8200
 * chars. Sheet cell limit is 50000 chars, so this fits comfortably.
 */

const EMBEDDING_MODEL = 'text-embedding-3-small'
const EMBEDDING_DIMS = 1536
const EMBEDDING_API_URL = 'https://api.openai.com/v1/embeddings'
const SIMILARITY_DEFAULT_K = 3
const SIMILARITY_MIN_SCORE = 0.55

function generateEmbedding_(text) {
  if (!text || !text.trim()) return null

  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY')
  if (!apiKey) return null

  const response = UrlFetchApp.fetch(EMBEDDING_API_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + apiKey },
    payload: JSON.stringify({
      model: EMBEDDING_MODEL,
      input: text.slice(0, 8000)
    }),
    muteHttpExceptions: true
  })

  const code = response.getResponseCode()
  if (code < 200 || code >= 300) {
    console.warn('Embedding API failed: ' + code + ' ' + response.getContentText().slice(0, 200))
    return null
  }

  const parsed = JSON.parse(response.getContentText())
  const vector = parsed && parsed.data && parsed.data[0] && parsed.data[0].embedding
  if (!vector || vector.length !== EMBEDDING_DIMS) {
    console.warn('Unexpected embedding response shape')
    return null
  }

  return vector
}

/**
 * Float32 array → base64 string. Little-endian byte layout so the same
 * code on any reader produces consistent results.
 */
function packEmbedding_(floats) {
  const buffer = new ArrayBuffer(floats.length * 4)
  const view = new DataView(buffer)
  for (let i = 0; i < floats.length; i++) {
    view.setFloat32(i * 4, floats[i], true)
  }
  const u8 = new Uint8Array(buffer)
  const bytes = new Array(u8.length)
  for (let i = 0; i < u8.length; i++) bytes[i] = u8[i]
  return Utilities.base64Encode(bytes)
}

/**
 * base64 string → Float32 array. Apps Script's Utilities.base64Decode
 * returns a Java byte[] with signed values; we coerce to unsigned before
 * reading floats.
 */
function unpackEmbedding_(b64) {
  if (!b64) return null
  const bytes = Utilities.base64Decode(b64)
  if (!bytes || bytes.length % 4 !== 0) return null

  const buffer = new ArrayBuffer(bytes.length)
  const u8 = new Uint8Array(buffer)
  for (let i = 0; i < bytes.length; i++) u8[i] = bytes[i] & 0xff

  const view = new DataView(buffer)
  const floats = new Array(bytes.length / 4)
  for (let i = 0; i < floats.length; i++) floats[i] = view.getFloat32(i * 4, true)
  return floats
}

/**
 * Dot product. OpenAI v3 embeddings are unit-normalized so this equals
 * cosine similarity.
 */
function dot_(a, b) {
  if (!a || !b || a.length !== b.length) return 0
  let sum = 0
  for (let i = 0; i < a.length; i++) sum += a[i] * b[i]
  return sum
}

/**
 * Generates an embedding for the given row's JD Text and writes it to
 * the Embedding column. Called from scoreJD_() after a successful score.
 * No-op if the row has no JD text or OPENAI_API_KEY is not set.
 */
function storeRowEmbedding_(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_APPLICATIONS)
  if (!sheet) return

  const cols = getColumnMap_(sheet)
  if (!cols['Embedding']) return

  const jdText = sheet.getRange(row, cols['JD Text']).getValue()
  if (!jdText) return

  const company = sheet.getRange(row, cols['Company']).getValue() || ''
  const role = sheet.getRange(row, cols['Role']).getValue() || ''
  const enriched = company + ' ' + role + '\n\n' + jdText

  const vector = generateEmbedding_(enriched)
  if (!vector) return

  const packed = packEmbedding_(vector)
  const lock = LockService.getDocumentLock()
  try {
    lock.waitLock(5000)
    sheet.getRange(row, cols['Embedding']).setValue(packed)
  } finally {
    try { lock.releaseLock() } catch (_) {}
  }

  logActivity_({
    action: 'embedding_stored',
    rowRef: 'row:' + row,
    notes: company + ' / ' + role
  })
}

/**
 * For a target row, scans all other rows with embeddings and returns the
 * top-k by similarity score. Skips rows below SIMILARITY_MIN_SCORE.
 *
 * Returns an array of { row, company, role, status, fit_score, score }
 * sorted by score desc.
 */
function findSimilarRows_(targetRow, k) {
  const limit = k || SIMILARITY_DEFAULT_K
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_APPLICATIONS)
  if (!sheet || sheet.getLastRow() < 2) return []

  const cols = getColumnMap_(sheet)
  if (!cols['Embedding']) return []

  const targetEmbedding = unpackEmbedding_(sheet.getRange(targetRow, cols['Embedding']).getValue())
  if (!targetEmbedding) return []

  const candidateRows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues()
  const matches = []

  for (let i = 0; i < candidateRows.length; i++) {
    const rowIndex = i + 2
    if (rowIndex === targetRow) continue

    const packed = candidateRows[i][cols['Embedding'] - 1]
    if (!packed) continue

    const otherEmbedding = unpackEmbedding_(packed)
    if (!otherEmbedding) continue

    const score = dot_(targetEmbedding, otherEmbedding)
    if (score < SIMILARITY_MIN_SCORE) continue

    matches.push({
      row: rowIndex,
      company: candidateRows[i][cols['Company'] - 1],
      role: candidateRows[i][cols['Role'] - 1],
      status: candidateRows[i][cols['Status'] - 1],
      fit_score: candidateRows[i][cols['Fit Score'] - 1],
      score: score
    })
  }

  matches.sort((a, b) => b.score - a.score)
  return matches.slice(0, limit)
}

/**
 * Public handler called from the sidebar. Wraps findSimilarRows_ with a
 * "no embeddings yet" message when the target row has not been embedded.
 */
function findSimilarForRow(row) {
  if (!row || row < 2) throw new Error('Pick an application row first.')

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_APPLICATIONS)
  const cols = getColumnMap_(sheet)
  const packed = sheet.getRange(row, cols['Embedding']).getValue()

  if (!packed) {
    const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY')
    if (!apiKey) {
      return { matches: [], message: 'Add an OPENAI_API_KEY in Setup to enable similarity search.' }
    }
    storeRowEmbedding_(row)
  }

  const matches = findSimilarRows_(row, SIMILARITY_DEFAULT_K)
  if (matches.length === 0) {
    return { matches: [], message: 'No similar past applications above threshold.' }
  }

  return { matches: matches, message: null }
}
