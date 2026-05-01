/**
 * WebApp.gs
 *
 * Apps Script Web App that serves the mobile UI. Deploy via
 * Deploy → New deployment → Web app, executeAs: User deploying,
 * access: Only myself.
 *
 * The Web App is single-user. The deployment URL is a long opaque string
 * that only Angel can hit (gated by Google sign-in). Add the URL to her
 * iOS home screen as a PWA via Safari.
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile('MobileApp')
    .evaluate()
    .setTitle('Job CRM')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover')
    .addMetaTag('apple-mobile-web-app-capable', 'yes')
    .addMetaTag('apple-mobile-web-app-status-bar-style', 'default')
    .addMetaTag('apple-mobile-web-app-title', 'Job CRM')
    .addMetaTag('theme-color', '#2D5266')
}

/**
 * Web App POST handler. Currently dispatches Slack slash commands; future
 * inbound integrations (Slack interactions, other webhooks) get added here.
 *
 * Slack posts as application/x-www-form-urlencoded. Apps Script populates
 * e.parameter from form fields. We detect Slack by the presence of `command`
 * and `token` in the body.
 */
function doPost(e) {
  if (!e || !e.parameter) {
    return ContentService.createTextOutput('Bad request').setMimeType(ContentService.MimeType.TEXT)
  }

  if (e.parameter.command && e.parameter.token) {
    return handleSlackSlashCommand_(e)
  }

  return ContentService.createTextOutput('Unsupported payload').setMimeType(ContentService.MimeType.TEXT)
}

/**
 * Returns recent applications for the mobile home/recent view.
 * Sorted by Last Touch desc.
 */
function getRecentApplicationsForMobile(limit) {
  const max = limit || 50
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_APPLICATIONS)
  if (!sheet || sheet.getLastRow() < 2) return { rows: [], briefing: null }

  const cols = getColumnMap_(sheet)
  const rangeData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues()

  const rows = rangeData.map((r, idx) => ({
    row: idx + 2,
    company: r[cols['Company'] - 1],
    role: r[cols['Role'] - 1],
    status: r[cols['Status'] - 1],
    fit_score: r[cols['Fit Score'] - 1],
    last_touch: r[cols['Last Touch'] - 1] instanceof Date ? r[cols['Last Touch'] - 1].toISOString() : null,
    next_action: r[cols['Next Action'] - 1],
    drive_folder: r[cols['Drive Folder'] - 1]
  })).filter(r => r.company)

  rows.sort((a, b) => {
    const ta = a.last_touch ? new Date(a.last_touch).getTime() : 0
    const tb = b.last_touch ? new Date(b.last_touch).getTime() : 0
    return tb - ta
  })

  return {
    rows: rows.slice(0, max),
    briefing: buildDailyBriefing_(rows)
  }
}

function buildDailyBriefing_(rows) {
  const followupsDue = []
  const interviewsThisWeek = []
  const sevenDaysAgo = new Date()
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7)
  const sevenDaysFromNow = new Date()
  sevenDaysFromNow.setDate(sevenDaysFromNow.getDate() + 7)

  for (const r of rows) {
    if (r.status === 'Applied' && r.last_touch) {
      const lt = new Date(r.last_touch)
      if (lt < sevenDaysAgo) followupsDue.push(r)
    }
    if (r.status === 'Interview Scheduled' && r.next_action) {
      const interviewDate = new Date(r.next_action)
      if (!isNaN(interviewDate.getTime()) && interviewDate >= new Date() && interviewDate <= sevenDaysFromNow) {
        interviewsThisWeek.push(r)
      }
    }
  }

  return {
    follow_ups_due: followupsDue.length,
    interviews_this_week: interviewsThisWeek.length,
    follow_ups: followupsDue.slice(0, 5),
    interviews: interviewsThisWeek.slice(0, 5),
    monthly_spend_usd: getMonthlySpendUsd_()
  }
}

/**
 * Quick log: append a new row from the mobile form.
 * Returns the new row index for follow-up actions.
 */
function quickLogApplication(payload) {
  if (!payload || !payload.company) throw new Error('Company is required.')

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_APPLICATIONS)
  const cols = getColumnMap_(sheet)

  const row = sheet.getLastRow() + 1
  const lock = LockService.getDocumentLock()
  try {
    lock.waitLock(5000)
    sheet.getRange(row, cols['Company']).setValue(payload.company)
    sheet.getRange(row, cols['Role']).setValue(payload.role || '')
    sheet.getRange(row, cols['JD Link']).setValue(payload.jd_link || '')
    sheet.getRange(row, cols['Status']).setValue(payload.status || 'Saved')
    sheet.getRange(row, cols['Tags']).setValue(payload.tags || '')
    sheet.getRange(row, cols['Notes']).setValue(payload.notes || '')
    sheet.getRange(row, cols['Last Touch']).setValue(new Date())
    if ((payload.status || 'Saved') === 'Applied') {
      sheet.getRange(row, cols['Date Applied']).setValue(new Date())
    }
  } finally {
    try { lock.releaseLock() } catch (_) {}
  }

  if (payload.company) {
    try {
      const folderUrl = ensureCompanyDriveFolder_(payload.company)
      sheet.getRange(row, cols['Drive Folder']).setValue(folderUrl)
    } catch (err) {
      console.warn('Folder creation failed: ' + err.message)
    }
  }

  logActivity_({
    action: 'mobile_quick_log',
    rowRef: 'row:' + row,
    notes: payload.company + ' / ' + (payload.role || '?')
  })

  return { row: row }
}

/**
 * Mobile swipe action: update status.
 */
function mobileSetStatus(row, status) {
  if (!row || row < 2) throw new Error('Invalid row')
  const valid = ['Saved', 'Applied', 'Interview Scheduled', 'Interview Done', 'Offer', 'Rejected', 'Withdrawn']
  if (valid.indexOf(status) === -1) throw new Error('Invalid status: ' + status)

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_APPLICATIONS)
  const cols = getColumnMap_(sheet)
  return setStatusAndTouch_(row, sheet, cols, status)
}

/**
 * Mobile-friendly score call. Same as sidebar but stripped down for the
 * lightweight mobile view.
 */
function mobileScoreRow(row) {
  return scoreJD_(row)
}
