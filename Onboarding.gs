/**
 * Onboarding.gs
 *
 * Richer first-run wizard. The setup() function in Setup.gs uses plain UI
 * prompts for fast scripted use; this one is a modal dialog for the
 * smoother experience.
 *
 * Run via menu: Job CRM → Setup → Onboarding wizard, or invoke directly.
 */

function showOnboardingWizard_() {
  const html = HtmlService.createTemplateFromFile('OnboardingDialog')
    .evaluate()
    .setWidth(520)
    .setHeight(640)
  SpreadsheetApp.getUi().showModalDialog(html, 'Set up Job CRM')
}

/**
 * Reads current property values for the wizard to pre-fill.
 */
function getOnboardingState() {
  const props = PropertiesService.getScriptProperties()
  return {
    has_api_key: !!props.getProperty('ANTHROPIC_API_KEY'),
    resume_doc_id: props.getProperty('RESUME_DOC_ID') || '',
    drive_parent_folder_id: props.getProperty('DRIVE_PARENT_FOLDER_ID') || '',
    gmail_label: props.getProperty('GMAIL_LABEL') || 'job-apply',
    monthly_spend_threshold: props.getProperty('MONTHLY_SPEND_THRESHOLD_USD') || '20'
  }
}

/**
 * Validates a Drive URL or ID. Accepts full URLs or bare IDs. Returns the
 * extracted ID, or throws if the format is unrecognized.
 */
function parseDriveId_(input, kind) {
  if (!input) throw new Error(kind + ' is empty.')
  const trimmed = input.trim()

  let match = trimmed.match(/\/d\/([a-zA-Z0-9_-]{20,})/)
  if (match) return match[1]

  match = trimmed.match(/\/folders\/([a-zA-Z0-9_-]{20,})/)
  if (match) return match[1]

  match = trimmed.match(/[?&]id=([a-zA-Z0-9_-]{20,})/)
  if (match) return match[1]

  if (/^[a-zA-Z0-9_-]{20,}$/.test(trimmed)) return trimmed

  throw new Error(kind + ': could not parse a Drive ID from "' + trimmed.slice(0, 60) + '". Paste the full URL or just the ID.')
}

/**
 * Validate the API key by making a tiny Claude call.
 */
function testAnthropicKey(apiKey) {
  if (!apiKey || apiKey.length < 20) throw new Error('API key looks too short. Get one at console.anthropic.com.')

  const response = UrlFetchApp.fetch(ANTHROPIC_API_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': ANTHROPIC_VERSION
    },
    payload: JSON.stringify({
      model: 'claude-haiku-4-5',
      max_tokens: 32,
      messages: [{ role: 'user', content: 'Respond with the exact word: ok' }]
    }),
    muteHttpExceptions: true
  })

  const code = response.getResponseCode()
  const body = response.getContentText()

  if (code === 401 || code === 403) {
    throw new Error('API key rejected by Anthropic. Check the key at console.anthropic.com.')
  }
  if (code < 200 || code >= 300) {
    throw new Error('Test call failed with status ' + code + ': ' + body.slice(0, 300))
  }

  return { ok: true }
}

/**
 * Main wizard submit handler.
 */
function saveOnboarding(payload) {
  if (!payload) throw new Error('No payload')

  const props = PropertiesService.getScriptProperties()

  if (payload.api_key) {
    testAnthropicKey(payload.api_key)
    props.setProperty('ANTHROPIC_API_KEY', payload.api_key)
  } else if (!props.getProperty('ANTHROPIC_API_KEY')) {
    throw new Error('Anthropic API key is required.')
  }

  const resumeId = parseDriveId_(payload.resume_doc, 'Resume Doc')
  try {
    DocumentApp.openById(resumeId).getBody().getText()
  } catch (err) {
    throw new Error('Resume Doc is not readable. Confirm the URL is correct and you own the doc. Detail: ' + err.message)
  }
  props.setProperty('RESUME_DOC_ID', resumeId)

  const folderId = parseDriveId_(payload.drive_parent, 'Drive parent folder')
  try {
    DriveApp.getFolderById(folderId).getName()
  } catch (err) {
    throw new Error('Drive folder is not readable. Confirm the URL is correct. Detail: ' + err.message)
  }
  props.setProperty('DRIVE_PARENT_FOLDER_ID', folderId)

  const label = (payload.gmail_label || 'job-apply').trim()
  props.setProperty('GMAIL_LABEL', label)
  ensureGmailLabelExists_(label)

  const threshold = String(parseFloat(payload.monthly_spend_threshold) || 20)
  props.setProperty('MONTHLY_SPEND_THRESHOLD_USD', threshold)

  ensureSheetsExist_()
  installTriggers_()

  return { ok: true }
}
