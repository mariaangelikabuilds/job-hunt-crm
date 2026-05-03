/**
 * ResumeDiff.gs
 *
 * Line-level diff between the master resume (Drive Doc, cached) and a
 * tailored resume (Drive Doc per company). Surfaces add/delete/same lines
 * so Angel can audit a tailored resume for fabrication before sending.
 *
 * Why line diff and not a fuzzy semantic compare: fabrication shows up
 * as plain new tokens (skill names, metric numbers, project titles) not
 * present in the master. A line diff makes those visible at a glance;
 * any '+' line that is not a pure rephrase is the audit target.
 *
 * Algorithm: standard LCS on lines, O(n*m) memory. Resumes are short
 * (typically under 200 lines), so this is trivial in compute and well
 * under the Apps Script execution limit.
 */

function getResumeDiff(row) {
  if (!row || row < 2) throw new Error('Invalid row')

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_APPLICATIONS)
  const cols = getColumnMap_(sheet)
  const tailoredUrl = sheet.getRange(row, cols['Tailored Resume']).getValue()
  if (!tailoredUrl) throw new Error('No tailored resume on this row. Click Tailor resume first.')

  const tailoredId = extractDocId_(tailoredUrl)
  if (!tailoredId) throw new Error('Could not parse Doc ID from Tailored Resume URL.')

  const master = getResumeText_()
  const tailored = DocumentApp.openById(tailoredId).getBody().getText()

  const diff = computeResumeDiff_(master, tailored)
  return {
    diff: diff,
    summary: summarizeDiff_(diff)
  }
}

function computeResumeDiff_(masterText, tailoredText) {
  const a = String(masterText || '').split('\n')
  const b = String(tailoredText || '').split('\n')
  const dp = lcsLengths_(a, b)
  return backtrackDiff_(a, b, dp)
}

function lcsLengths_(a, b) {
  const m = a.length
  const n = b.length
  const dp = new Array(m + 1)
  for (let i = 0; i <= m; i++) dp[i] = new Array(n + 1).fill(0)
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      if (a[i - 1] === b[j - 1]) dp[i][j] = dp[i - 1][j - 1] + 1
      else dp[i][j] = dp[i - 1][j] >= dp[i][j - 1] ? dp[i - 1][j] : dp[i][j - 1]
    }
  }
  return dp
}

function backtrackDiff_(a, b, dp) {
  const out = []
  let i = a.length
  let j = b.length
  while (i > 0 && j > 0) {
    if (a[i - 1] === b[j - 1]) {
      out.push({ op: ' ', line: a[i - 1] })
      i--
      j--
    } else if (dp[i - 1][j] >= dp[i][j - 1]) {
      out.push({ op: '-', line: a[i - 1] })
      i--
    } else {
      out.push({ op: '+', line: b[j - 1] })
      j--
    }
  }
  while (i > 0) {
    out.push({ op: '-', line: a[i - 1] })
    i--
  }
  while (j > 0) {
    out.push({ op: '+', line: b[j - 1] })
    j--
  }
  return out.reverse()
}

function extractDocId_(url) {
  const match = String(url).match(/\/d\/([a-zA-Z0-9_-]+)/)
  return match ? match[1] : null
}

function summarizeDiff_(diff) {
  let added = 0
  let removed = 0
  let same = 0
  for (const entry of diff) {
    if (entry.op === '+') added++
    else if (entry.op === '-') removed++
    else same++
  }
  return { added: added, removed: removed, same: same }
}
