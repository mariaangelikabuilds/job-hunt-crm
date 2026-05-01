/**
 * Setup.gs
 *
 * Idempotent first-run setup. Run once after pasting all .gs files into the
 * Apps Script editor and binding the project to a fresh Sheet.
 *
 * - Prompts for ANTHROPIC_API_KEY and RESUME_DOC_ID (Script Properties)
 * - Creates the three sheet tabs (Applications, Analytics, Activity)
 * - Installs time-driven, onEdit, and onSelectionChange triggers
 * - Seeds the menu via onOpen
 *
 * Re-runnable. Existing triggers are not duplicated. Existing properties are
 * preserved unless the prompt is answered with a new value.
 */

const APP_NAME = 'Job CRM'
const SHEET_APPLICATIONS = 'Applications'
const SHEET_ANALYTICS = 'Analytics'
const SHEET_ACTIVITY = 'Activity'

const APPLICATIONS_HEADERS = [
  'Company',
  'Role',
  'JD Link',
  'JD Text',
  'Date Applied',
  'Status',
  'Tags',
  'Fit Score',
  'Cover Angles',
  'Red Flags',
  'Why Score',
  'Last Touch',
  'Next Action',
  'Drive Folder',
  'Tailored Resume',
  'Cover Letter',
  'Picked Angle',
  'Notes',
  'Embedding'
]

const STATUS_VALUES = [
  'Saved',
  'Applied',
  'Interview Scheduled',
  'Interview Done',
  'Offer',
  'Rejected',
  'Withdrawn'
]

const ACTIVITY_HEADERS = [
  'Timestamp',
  'Action',
  'Row Ref',
  'Input Tokens',
  'Cache Read Tokens',
  'Cache Creation Tokens',
  'Output Tokens',
  'Cost USD',
  'Notes'
]

function setup() {
  const ui = SpreadsheetApp.getUi()
  const props = PropertiesService.getScriptProperties()

  promptForProperty_(ui, props, 'ANTHROPIC_API_KEY', 'Paste your Anthropic API key. Get one at console.anthropic.com.')
  promptForProperty_(ui, props, 'RESUME_DOC_ID', 'Paste the Drive Doc ID for your master resume. Open the doc, copy the long string between /d/ and /edit in the URL.')
  promptForProperty_(ui, props, 'GMAIL_LABEL', 'Gmail label to watch for job application emails.', 'job-apply')
  promptForProperty_(ui, props, 'DRIVE_PARENT_FOLDER_ID', 'Drive folder ID where per-company sub-folders should live. Open the folder, copy the long string after /folders/ in the URL.')
  promptForProperty_(ui, props, 'MONTHLY_SPEND_THRESHOLD_USD', 'Monthly Anthropic API spend threshold for the sidebar amber indicator.', '20')

  promptForProperty_(ui, props, 'SLACK_WEBHOOK_URL', 'Optional. Slack incoming webhook URL for daily follow-up nudges. Leave blank to skip Slack.', '')
  promptForProperty_(ui, props, 'SLACK_NUDGE_TARGET', 'Where to send the daily nudge: "email", "slack", or "both". Defaults to email.', 'email')
  promptForProperty_(ui, props, 'SLACK_VERIFICATION_TOKEN', 'Optional. Slack slash-command verification token (from the Slack app config). Leave blank if not using slash commands.', '')
  promptForProperty_(ui, props, 'SLACK_TEAM_ID', 'Optional. Slack workspace team_id allowlist for slash commands (e.g., T01ABCDEF). Leave blank to allow any team that has the verification token.', '')

  promptForProperty_(ui, props, 'OPENAI_API_KEY', 'Optional. OpenAI API key for JD embeddings (find similar past applications). Leave blank to skip embeddings.', '')

  ensureSheetsExist_()
  ensureGmailLabelExists_(props.getProperty('GMAIL_LABEL'))
  installTriggers_()

  ui.alert('Setup complete', 'Open the Job CRM menu and click "Open sidebar". The panel binds to whichever row is selected on the Applications tab.', ui.ButtonSet.OK)
}

function promptForProperty_(ui, props, key, message, fallback) {
  const current = props.getProperty(key)
  const prefix = current ? 'Currently set. Press OK to keep, or paste a new value to replace.\n\n' : ''
  const result = ui.prompt(key, prefix + message, ui.ButtonSet.OK_CANCEL)
  if (result.getSelectedButton() !== ui.Button.OK) return
  const value = result.getResponseText().trim()
  if (value) {
    props.setProperty(key, value)
  } else if (!current && fallback) {
    props.setProperty(key, fallback)
  }
}

function ensureSheetsExist_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  ensureApplicationsSheet_(ss)
  ensureAnalyticsSheet_(ss)
  ensureActivitySheet_(ss)
}

function ensureApplicationsSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_APPLICATIONS)
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_APPLICATIONS, 0)
  }
  if (sheet.getLastColumn() === 0 || sheet.getRange(1, 1).getValue() !== APPLICATIONS_HEADERS[0]) {
    sheet.getRange(1, 1, 1, APPLICATIONS_HEADERS.length).setValues([APPLICATIONS_HEADERS])
    sheet.setFrozenRows(1)
    sheet.getRange(1, 1, 1, APPLICATIONS_HEADERS.length).setFontWeight('bold')
  } else {
    const existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    const existingSet = {}
    for (const h of existing) if (h) existingSet[h] = true
    let appendIndex = sheet.getLastColumn()
    for (const header of APPLICATIONS_HEADERS) {
      if (!existingSet[header]) {
        appendIndex += 1
        sheet.getRange(1, appendIndex).setValue(header).setFontWeight('bold')
      }
    }
  }
  const statusCol = APPLICATIONS_HEADERS.indexOf('Status') + 1
  const range = sheet.getRange(2, statusCol, sheet.getMaxRows() - 1, 1)
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_VALUES, true)
    .setAllowInvalid(false)
    .build()
  range.setDataValidation(rule)

  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
  const embeddingIdx = headerRow.indexOf('Embedding')
  if (embeddingIdx >= 0) sheet.hideColumns(embeddingIdx + 1)
}

function ensureAnalyticsSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_ANALYTICS)
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_ANALYTICS)
  }
  if (sheet.getRange(1, 1).getValue()) return

  const formulas = [
    ['Total applications', '=COUNTA(Applications!A2:A)'],
    ['Applied', '=COUNTIF(Applications!F2:F, "Applied")'],
    ['Interview Scheduled', '=COUNTIF(Applications!F2:F, "Interview Scheduled")'],
    ['Interview Done', '=COUNTIF(Applications!F2:F, "Interview Done")'],
    ['Offer', '=COUNTIF(Applications!F2:F, "Offer")'],
    ['Rejected', '=COUNTIF(Applications!F2:F, "Rejected")'],
    ['Average fit score', '=IFERROR(AVERAGE(Applications!H2:H), "no data")'],
    ['Response rate', '=IFERROR((COUNTIF(Applications!F2:F, "Interview Scheduled") + COUNTIF(Applications!F2:F, "Interview Done") + COUNTIF(Applications!F2:F, "Offer") + COUNTIF(Applications!F2:F, "Rejected")) / COUNTIF(Applications!F2:F, "Applied"), 0)'],
    ['', ''],
    ['Token usage', ''],
    ['Total input tokens', '=SUM(Activity!D2:D)'],
    ['Total cache read tokens', '=SUM(Activity!E2:E)'],
    ['Total output tokens', '=SUM(Activity!G2:G)'],
    ['Total spend USD', '=SUM(Activity!H2:H)'],
    ['Cache hit rate', '=IFERROR(SUM(Activity!E2:E) / (SUM(Activity!D2:D) + SUM(Activity!E2:E)), 0)']
  ]
  sheet.getRange(1, 1, formulas.length, 2).setValues(formulas)
  sheet.getRange(1, 1, formulas.length, 1).setFontWeight('bold')
  sheet.setColumnWidth(1, 220)
  sheet.setColumnWidth(2, 200)
  const rateRow = formulas.findIndex(r => r[0] === 'Response rate') + 1
  if (rateRow > 0) sheet.getRange(rateRow, 2).setNumberFormat('0%')
  const hitRow = formulas.findIndex(r => r[0] === 'Cache hit rate') + 1
  if (hitRow > 0) sheet.getRange(hitRow, 2).setNumberFormat('0%')
  const spendRow = formulas.findIndex(r => r[0] === 'Total spend USD') + 1
  if (spendRow > 0) sheet.getRange(spendRow, 2).setNumberFormat('$0.0000')
}

function ensureActivitySheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_ACTIVITY)
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_ACTIVITY)
  }
  if (sheet.getRange(1, 1).getValue() !== ACTIVITY_HEADERS[0]) {
    sheet.getRange(1, 1, 1, ACTIVITY_HEADERS.length).setValues([ACTIVITY_HEADERS])
    sheet.setFrozenRows(1)
    sheet.getRange(1, 1, 1, ACTIVITY_HEADERS.length).setFontWeight('bold')
  }
  sheet.hideSheet()
}

function ensureGmailLabelExists_(name) {
  if (!name) return
  const existing = GmailApp.getUserLabelByName(name)
  if (!existing) GmailApp.createLabel(name)
}

function installTriggers_() {
  const existing = ScriptApp.getProjectTriggers()
  const handlerNames = existing.map(t => t.getHandlerFunction())

  const ss = SpreadsheetApp.getActiveSpreadsheet()
  if (!handlerNames.includes('parseGmailLabel_')) {
    ScriptApp.newTrigger('parseGmailLabel_').timeBased().everyMinutes(5).create()
  }
  if (!handlerNames.includes('dailyNudge_')) {
    ScriptApp.newTrigger('dailyNudge_').timeBased().everyDays(1).atHour(9).create()
  }
  if (!handlerNames.includes('onStatusEdit_')) {
    ScriptApp.newTrigger('onStatusEdit_').forSpreadsheet(ss).onEdit().create()
  }
  if (!handlerNames.includes('onSelectionChange_')) {
    ScriptApp.newTrigger('onSelectionChange_').forSpreadsheet(ss).onSelectionChange().create()
  }
}

function refreshTriggers_() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t))
  installTriggers_()
  SpreadsheetApp.getUi().alert('Triggers reinstalled.')
}
