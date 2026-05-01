/**
 * Gmail.gs
 *
 * Two background flows:
 *
 * 1. parseGmailLabel_ runs every 5 minutes via time-driven trigger. Scans
 *    threads with the configured label (default "job-apply"), parses each
 *    message via Claude (cheap extraction call), appends a row to the
 *    Applications tab with Company / Role / JD Link / Date Applied filled
 *    in. Removes the label after processing so threads aren't reprocessed.
 *
 * 2. dailyNudge_ runs once a day at 9am Manila time. Finds rows where
 *    Status = Applied and Last Touch > 7 days. Sends a digest email to
 *    Angel with one-click draft follow-ups for each.
 *
 * Both flows use Claude minimally (one extraction call per thread, one
 * draft per overdue row). Token spend is bounded.
 */

function parseGmailLabel_() {
  const labelName = PropertiesService.getScriptProperties().getProperty('GMAIL_LABEL') || 'job-apply'
  const label = GmailApp.getUserLabelByName(labelName)
  if (!label) return

  const threads = label.getThreads(0, 25)
  if (threads.length === 0) return

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_APPLICATIONS)
  const cols = getColumnMap_(sheet)

  for (const thread of threads) {
    try {
      processThread_(thread, sheet, cols)
      thread.removeLabel(label)
    } catch (err) {
      console.warn('Could not process thread "' + thread.getFirstMessageSubject() + '": ' + err.message)
    }
  }
}

function processThread_(thread, sheet, cols) {
  const message = thread.getMessages()[0]
  const subject = message.getSubject() || ''
  const body = message.getPlainBody() || ''
  const from = message.getFrom() || ''

  const extracted = extractApplicationFromEmail_(subject, body, from, thread.getId())
  if (!extracted) return

  const lock = LockService.getDocumentLock()
  try {
    lock.waitLock(5000)
    const row = sheet.getLastRow() + 1
    sheet.getRange(row, cols['Company']).setValue(extracted.company || '')
    sheet.getRange(row, cols['Role']).setValue(extracted.role || '')
    sheet.getRange(row, cols['JD Link']).setValue(extracted.jd_link || '')
    sheet.getRange(row, cols['Date Applied']).setValue(message.getDate())
    sheet.getRange(row, cols['Status']).setValue(extracted.status || 'Applied')
    sheet.getRange(row, cols['Last Touch']).setValue(new Date())
    sheet.getRange(row, cols['Notes']).setValue('Auto-imported from Gmail thread ' + thread.getId())

    if (extracted.company) {
      try {
        const folderUrl = ensureCompanyDriveFolder_(extracted.company)
        sheet.getRange(row, cols['Drive Folder']).setValue(folderUrl)
      } catch (err) {
        console.warn('Folder creation failed: ' + err.message)
      }
    }
  } finally {
    try { lock.releaseLock() } catch (_) {}
  }

  logActivity_({
    action: 'email_imported',
    rowRef: 'thread:' + thread.getId(),
    notes: extracted.company + ' / ' + extracted.role
  })
}

const EMAIL_EXTRACTION_SCHEMA = {
  type: 'object',
  properties: {
    company: { type: 'string' },
    role: { type: 'string' },
    jd_link: { type: 'string' },
    status: {
      type: 'string',
      enum: ['Applied', 'Interview Scheduled', 'Rejected', 'Saved', '']
    }
  },
  required: ['company', 'role', 'jd_link', 'status'],
  additionalProperties: false
}

function extractApplicationFromEmail_(subject, body, from, threadRef) {
  const truncatedBody = body.length > 8000 ? body.slice(0, 8000) : body

  const userContent = [
    'From: ' + from,
    'Subject: ' + subject,
    '',
    truncatedBody
  ].join('\n')

  const systemBlocks = [
    {
      type: 'text',
      text: [
        'You parse job-application emails. Output JSON with company, role, jd_link, status.',
        '',
        'Status mapping:',
        '- "Applied" if this is a confirmation that her application was received',
        '- "Interview Scheduled" if a recruiter is offering interview times or scheduling one',
        '- "Rejected" if this is a rejection email',
        '- "Saved" if this is just a JD she forwarded to herself for later',
        '- "" empty string if unclear',
        '',
        'jd_link: extract the canonical JD URL if present (Greenhouse, Lever, Workable, Ashby, careers page links). Empty string if none.',
        '',
        'company / role: extract from the email content. Empty strings if not present.',
        '',
        'Do not invent. Empty fields are fine.'
      ].join('\n'),
      cache_control: { type: 'ephemeral' }
    }
  ]

  const payload = {
    model: DEFAULT_MODEL,
    max_tokens: 1024,
    system: systemBlocks,
    messages: [{ role: 'user', content: userContent }],
    thinking: { type: 'adaptive' },
    output_config: {
      effort: 'low',
      format: { type: 'json_schema', schema: EMAIL_EXTRACTION_SCHEMA }
    }
  }

  try {
    const response = callClaude_(payload)
    const parsed = parseJsonResponse_(response)

    logActivity_({
      action: 'email_extract',
      rowRef: 'thread:' + threadRef,
      model: DEFAULT_MODEL,
      usage: response.usage,
      notes: parsed.company + ' / ' + parsed.role
    })

    return parsed
  } catch (err) {
    console.warn('Email extraction failed: ' + err.message)
    return null
  }
}

function dailyNudge_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_APPLICATIONS)
  if (!sheet || sheet.getLastRow() < 2) return

  const cols = getColumnMap_(sheet)
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues()

  const sevenDaysAgo = new Date()
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7)

  const overdue = []
  for (let i = 0; i < data.length; i++) {
    const row = data[i]
    const status = row[cols['Status'] - 1]
    const lastTouch = row[cols['Last Touch'] - 1]
    if (status !== 'Applied') continue
    if (!(lastTouch instanceof Date) || lastTouch >= sevenDaysAgo) continue

    overdue.push({
      rowIndex: i + 2,
      company: row[cols['Company'] - 1],
      role: row[cols['Role'] - 1],
      lastTouch: lastTouch,
      coverAngles: row[cols['Cover Angles'] - 1]
    })
  }

  if (overdue.length === 0) return

  const target = (PropertiesService.getScriptProperties().getProperty('SLACK_NUDGE_TARGET') || 'email').toLowerCase()
  const wantsSlack = target === 'slack' || target === 'both'
  const wantsEmail = target === 'email' || target === 'both'

  if (wantsEmail) sendNudgeDigest_(overdue)
  if (wantsSlack) postNudgeToSlack_(overdue)

  logActivity_({
    action: 'daily_nudge',
    rowRef: 'count:' + overdue.length,
    notes: 'target=' + target + '; ' + overdue.map(o => o.company).join(', ')
  })
}

function sendNudgeDigest_(overdue) {
  const email = Session.getActiveUser().getEmail()
  if (!email) return

  const lines = [
    'Follow-ups due. ' + overdue.length + ' application' + (overdue.length === 1 ? '' : 's') + ' have not been touched in 7+ days.',
    '',
    'Quick tip: in the Job CRM sidebar, click any cover angle to draft a follow-up email pre-loaded with that angle.',
    ''
  ]

  for (const item of overdue) {
    const days = Math.floor((Date.now() - item.lastTouch.getTime()) / (1000 * 60 * 60 * 24))
    lines.push('Row ' + item.rowIndex + ': ' + item.company + ' / ' + item.role)
    lines.push('  Last touch: ' + days + ' days ago')
    if (item.coverAngles) {
      lines.push('  First angle: ' + String(item.coverAngles).split('\n')[0])
    }
    lines.push('')
  }

  GmailApp.sendEmail(email, 'Job CRM: ' + overdue.length + ' follow-up' + (overdue.length === 1 ? '' : 's') + ' due', lines.join('\n'))
}
