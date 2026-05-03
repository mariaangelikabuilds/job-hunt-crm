/**
 * Code.gs
 *
 * Menu, orchestration, and sheet helpers. The hot path is scoreJD_(row),
 * called from the menu and from the sidebar. Everything else is plumbing.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(APP_NAME)
    .addItem('Open sidebar', 'openSidebar_')
    .addItem('Score active row', 'scoreActiveRow_')
    .addItem('Generate follow-up for active row', 'generateFollowupForActiveRow_')
    .addItem('Find similar past applications', 'findSimilarForActiveRow_')
    .addSeparator()
    .addItem('Import from CSV', 'showBulkImportDialog_')
    .addItem('Generate resume from LinkedIn', 'showLinkedInImportDialog_')
    .addSeparator()
    .addItem('Refresh resume cache', 'refreshResumeCache_')
    .addItem('Refresh triggers', 'refreshTriggers_')
    .addSeparator()
    .addItem('Onboarding wizard', 'showOnboardingWizard_')
    .addItem('Setup (quick prompts)', 'setup')
    .addToUi()
}

function openSidebar_() {
  const html = HtmlService.createTemplateFromFile('Sidebar')
    .evaluate()
    .setTitle('Job CRM')
    .setWidth(360)
  SpreadsheetApp.getUi().showSidebar(html)
}

function scoreActiveRow_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getActiveSheet()
  if (sheet.getName() !== SHEET_APPLICATIONS) {
    SpreadsheetApp.getUi().alert('Switch to the Applications tab and select a row first.')
    return
  }
  const row = sheet.getActiveRange().getRow()
  if (row < 2) {
    SpreadsheetApp.getUi().alert('Select a data row, not the header.')
    return
  }
  scoreJD_(row)
  SpreadsheetApp.getUi().alert('Scored row ' + row + '. Open sidebar to see the breakdown.')
}

/**
 * Core scoring entry point. Reads the row, calls Claude, writes the result
 * back to the sheet, returns a structured object for the sidebar.
 */
function scoreJD_(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_APPLICATIONS)
  const cols = getColumnMap_(sheet)

  const rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0]
  const jdText = rowValues[cols['JD Text'] - 1]
  const jdLink = rowValues[cols['JD Link'] - 1]

  let effectiveJd = jdText
  if (!effectiveJd && jdLink) {
    const extracted = scrapeJD_(jdLink, 'scrape:row:' + row)
    if (extracted && extracted.jd_text) {
      effectiveJd = extracted.jd_text
      const scrapeLock = LockService.getDocumentLock()
      try {
        scrapeLock.waitLock(5000)
        sheet.getRange(row, cols['JD Text']).setValue(effectiveJd)
        if (extracted.company && !rowValues[cols['Company'] - 1]) {
          sheet.getRange(row, cols['Company']).setValue(extracted.company)
        }
        if (extracted.role && !rowValues[cols['Role'] - 1]) {
          sheet.getRange(row, cols['Role']).setValue(extracted.role)
        }
      } finally {
        try { scrapeLock.releaseLock() } catch (_) {}
      }
    }
  }

  if (!effectiveJd) {
    throw new Error('Row ' + row + ' has no JD Text and no scrapable JD Link. Some boards (LinkedIn, Indeed) block bots; paste the JD body directly into the JD Text column.')
  }

  const scored = scoreJDViaClaude_(effectiveJd, { rowRef: 'row:' + row })
  const parsed = scored.parsed

  const lock = LockService.getDocumentLock()
  try {
    lock.waitLock(5000)
    sheet.getRange(row, cols['Fit Score']).setValue(parsed.fit_score)
    sheet.getRange(row, cols['Cover Angles']).setValue(parsed.top_3_angles.join('\n'))
    sheet.getRange(row, cols['Red Flags']).setValue(parsed.red_flags.join('\n'))
    sheet.getRange(row, cols['Why Score']).setValue(parsed.why_score)
    sheet.getRange(row, cols['Last Touch']).setValue(new Date())
  } finally {
    try { lock.releaseLock() } catch (_) {}
  }

  try {
    storeRowEmbedding_(row)
  } catch (err) {
    console.warn('Embedding generation failed for row ' + row + ': ' + err.message)
  }

  return {
    row: row,
    fit_score: parsed.fit_score,
    top_3_angles: parsed.top_3_angles,
    red_flags: parsed.red_flags,
    why_score: parsed.why_score,
    cached: scored.cached,
    cache_read_tokens: scored.usage ? scored.usage.cache_read_input_tokens : 0,
    monthly_spend_usd: getMonthlySpendUsd_()
  }
}

/**
 * Reads the active row data for the sidebar. Called via google.script.run.
 */
function getActiveRowData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getActiveSheet()
  if (sheet.getName() !== SHEET_APPLICATIONS) {
    return { onApplicationsTab: false }
  }
  const row = sheet.getActiveRange().getRow()
  if (row < 2) return { onApplicationsTab: true, validRow: false }

  const cols = getColumnMap_(sheet)
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0]

  const get = name => values[cols[name] - 1]

  return {
    onApplicationsTab: true,
    validRow: true,
    row: row,
    company: get('Company'),
    role: get('Role'),
    jd_link: get('JD Link'),
    jd_text: get('JD Text'),
    status: get('Status'),
    tags: get('Tags'),
    fit_score: get('Fit Score'),
    cover_angles: get('Cover Angles'),
    red_flags: get('Red Flags'),
    why_score: get('Why Score'),
    last_touch: get('Last Touch'),
    next_action: get('Next Action'),
    drive_folder: get('Drive Folder'),
    tailored_resume: get('Tailored Resume'),
    cover_letter: get('Cover Letter'),
    picked_angle: get('Picked Angle'),
    monthly_spend_usd: getMonthlySpendUsd_(),
    monthly_spend_threshold_usd: getMonthlySpendThresholdUsd_(),
    cache_hit_rate: getCacheHitRate_()
  }
}

/**
 * Sidebar action: write the JD Text from the textarea, then score.
 */
function saveJdAndScore(row, jdText) {
  if (!row || row < 2) throw new Error('Invalid row')
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_APPLICATIONS)
  const cols = getColumnMap_(sheet)
  if (jdText) {
    const lock = LockService.getDocumentLock()
    try {
      lock.waitLock(5000)
      sheet.getRange(row, cols['JD Text']).setValue(jdText)
    } finally {
      try { lock.releaseLock() } catch (_) {}
    }
  }
  return scoreJD_(row)
}

/**
 * Sidebar action: open a Gmail draft with the picked cover angle as the hook.
 */
function draftCoverEmail(row, angleIndex) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_APPLICATIONS)
  const cols = getColumnMap_(sheet)
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0]
  const company = values[cols['Company'] - 1]
  const role = values[cols['Role'] - 1]
  const angles = String(values[cols['Cover Angles'] - 1] || '').split('\n').filter(Boolean)
  const angle = angles[angleIndex] || angles[0]

  if (!angle) throw new Error('No cover angles on this row. Score the JD first.')

  const subject = 'Application: ' + role + ' at ' + company
  const body = [
    angle,
    '',
    '[Add specific reference to a recent product, leadership decision, or public statement from ' + company + ' here. Anti-AI rule: must be a real verifiable detail, not generic praise.]',
    '',
    '[Two-three sentences pulling from the master resume that ladder up to this angle. Real metrics only.]',
    '',
    'Resume: [link to tailored resume Doc]',
    'Portfolio: angel.dev',
    '',
    'Angel Agutaya'
  ].join('\n')

  const draft = GmailApp.createDraft('', subject, body)
  const lock = LockService.getDocumentLock()
  try {
    lock.waitLock(5000)
    sheet.getRange(row, cols['Picked Angle']).setValue(angleIndex + 1)
    sheet.getRange(row, cols['Last Touch']).setValue(new Date())
  } finally {
    try { lock.releaseLock() } catch (_) {}
  }

  logActivity_({
    action: 'cover_draft',
    rowRef: 'row:' + row,
    notes: 'angle=' + (angleIndex + 1)
  })

  return draft.getId()
}

/**
 * Helper: returns a map of header name to 1-indexed column number.
 */
function getColumnMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
  const map = {}
  for (let i = 0; i < headers.length; i++) {
    if (headers[i]) map[headers[i]] = i + 1
  }
  return map
}

function generateFollowupForActiveRow_() {
  SpreadsheetApp.getUi().alert('Follow-up drafting is wired through the sidebar quick actions. Open the sidebar, pick the row, and click Draft follow-up.')
}

function findSimilarForActiveRow_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getActiveSheet()
  if (sheet.getName() !== SHEET_APPLICATIONS) {
    SpreadsheetApp.getUi().alert('Switch to the Applications tab and select a row first.')
    return
  }
  const row = sheet.getActiveRange().getRow()
  if (row < 2) {
    SpreadsheetApp.getUi().alert('Select a data row, not the header.')
    return
  }

  const result = findSimilarForRow(row)
  if (result.message && result.matches.length === 0) {
    SpreadsheetApp.getUi().alert(result.message)
    return
  }

  const lines = result.matches.map(m => {
    const score = Math.round(m.score * 100)
    return 'Row ' + m.row + ': ' + m.company + ' / ' + m.role + '  (' + score + '% match, ' + m.status + ')'
  })
  SpreadsheetApp.getUi().alert('Top ' + result.matches.length + ' similar past applications', lines.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK)
}
