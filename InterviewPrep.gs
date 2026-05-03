/**
 * InterviewPrep.gs
 *
 * Sidebar Interview Prep mode (active when Status = "Interview Scheduled").
 *
 * Generates company research and likely-question lists via Claude on demand.
 * Caches the result per row in PropertiesService so re-opening the sidebar
 * doesn't re-spend tokens. Tracks the prep checklist state per row.
 */

const PREP_CHECKLIST_ITEMS = [
  { key: 'resume_locked', label: 'Tailored resume version locked in' },
  { key: 'portfolio_tested', label: 'Portfolio link opens cleanly on a fresh browser' },
  { key: 'questions_drafted', label: 'Questions to ask the interviewer drafted' },
  { key: 'company_review', label: 'Company research reviewed (below)' },
  { key: 'likely_questions', label: 'Likely questions reviewed (below)' },
  { key: 'travel_route', label: 'Route or video link tested' }
]

const COMPANY_RESEARCH_SCHEMA = {
  type: 'object',
  properties: {
    snapshot: { type: 'string' },
    recent_signals: {
      type: 'array',
      items: { type: 'string' }
    },
    leadership_notes: {
      type: 'array',
      items: { type: 'string' }
    },
    public_values: {
      type: 'array',
      items: { type: 'string' }
    },
    vibe_signal: { type: 'string' }
  },
  required: ['snapshot', 'recent_signals', 'leadership_notes', 'public_values', 'vibe_signal'],
  additionalProperties: false
}

const LIKELY_QUESTIONS_SCHEMA = {
  type: 'object',
  properties: {
    questions: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          q: { type: 'string' },
          angle: { type: 'string' }
        },
        required: ['q', 'angle'],
        additionalProperties: false
      }
    }
  },
  required: ['questions'],
  additionalProperties: false
}

function getInterviewPrepData(row) {
  if (!row || row < 2) throw new Error('Invalid row')

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_APPLICATIONS)
  const cols = getColumnMap_(sheet)
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0]

  const company = values[cols['Company'] - 1]
  const role = values[cols['Role'] - 1]
  const jdText = values[cols['JD Text'] - 1]
  const coverAngles = values[cols['Cover Angles'] - 1]
  const nextAction = values[cols['Next Action'] - 1]

  const cachedResearch = readCachedJson_('research_row_' + row)
  const cachedQuestions = readCachedJson_('questions_row_' + row)
  const checklist = readChecklist_(row)
  const notes = String(values[cols['Notes'] - 1] || '')

  return {
    row: row,
    company: company,
    role: role,
    next_action: nextAction,
    cover_angles: coverAngles,
    research: cachedResearch,
    questions: cachedQuestions,
    checklist_items: PREP_CHECKLIST_ITEMS,
    checklist_state: checklist,
    notes: notes,
    has_jd: !!jdText
  }
}

function generateCompanyResearch(row) {
  if (!row || row < 2) throw new Error('Invalid row')

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_APPLICATIONS)
  const cols = getColumnMap_(sheet)
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0]

  const company = values[cols['Company'] - 1]
  const role = values[cols['Role'] - 1]
  const jdText = values[cols['JD Text'] - 1]

  if (!company) throw new Error('Company is empty.')
  if (!jdText) throw new Error('JD Text is empty.')

  const systemBlocks = [
    {
      type: 'text',
      text: [
        'You research a company for an interview prep brief. The candidate is interviewing for the role described in the JD below.',
        '',
        'Output JSON with five fields:',
        '- snapshot: 2-3 sentences. What the company actually does, who pays them, what stage they are at. No marketing language.',
        '- recent_signals: 3-6 bullet points of recent public signals (funding rounds, leadership changes, product launches, strategic shifts) that the candidate should know walking in. Cite the year if you have it. Mark anything you are uncertain about with "(unverified)".',
        '- leadership_notes: notes on the people the candidate is likely to meet (CEO, head of design, hiring manager if named). Public information only. Empty list if you have nothing concrete.',
        '- public_values: 2-4 stated values, mission elements, or cultural signals from the company\'s public-facing material. Empty list if generic.',
        '- vibe_signal: one sentence. What the candidate should brace for in tone (formal vs casual, technical depth, pace expected).',
        '',
        'Hard rule: do not invent. If you do not know something, say so. Mark uncertain claims "(unverified)". Banned phrases: "innovative", "industry-leading", "cutting-edge", "results-driven", "best-in-class".'
      ].join('\n'),
      cache_control: { type: 'ephemeral' }
    }
  ]

  const userMessage = [
    'Company: ' + company,
    'Role: ' + role,
    '',
    'Job description:',
    '',
    jdText
  ].join('\n')

  const payload = {
    model: DEFAULT_MODEL,
    max_tokens: 4096,
    system: systemBlocks,
    messages: [{ role: 'user', content: userMessage }],
    thinking: { type: 'adaptive' },
    output_config: {
      effort: 'medium',
      format: { type: 'json_schema', schema: COMPANY_RESEARCH_SCHEMA }
    }
  }

  const response = callClaude_(payload)
  const parsed = parseJsonResponse_(response)

  writeCachedJson_('research_row_' + row, parsed)

  logActivity_({
    action: 'company_research',
    rowRef: 'row:' + row,
    model: DEFAULT_MODEL,
    usage: response.usage,
    notes: company
  })

  return parsed
}

function generateLikelyQuestions(row) {
  if (!row || row < 2) throw new Error('Invalid row')

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_APPLICATIONS)
  const cols = getColumnMap_(sheet)
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0]

  const company = values[cols['Company'] - 1]
  const role = values[cols['Role'] - 1]
  const jdText = values[cols['JD Text'] - 1]
  const coverAngles = values[cols['Cover Angles'] - 1]
  const resume = getResumeText_()

  if (!company || !jdText) throw new Error('Company and JD Text required.')

  const systemBlocks = [
    {
      type: 'text',
      text: [
        'CANDIDATE RESUME (for context, do not quote verbatim):',
        '',
        resume,
        '',
        '---',
        '',
        'You predict 5-7 specific interview questions this candidate is likely to face for this role at this company. Each question must come with a one-line angle for how the candidate should respond, grounded in actual experience from the resume above.',
        '',
        'Mix:',
        '- 1-2 questions about technical fit',
        '- 1-2 questions about the specific role responsibilities in the JD',
        '- 1-2 questions about the candidate\'s past work or transitions',
        '- 1 question about the company-candidate fit',
        '',
        'Banned questions: "tell me about yourself" (too generic), "why our company" (every JD asks this; not insight). Pick questions a real interviewer at this company would actually ask.',
        '',
        'Output JSON: { questions: [{q, angle}, ...] }. Five to seven items. No preamble.'
      ].join('\n'),
      cache_control: { type: 'ephemeral' }
    }
  ]

  const userMessage = [
    'Company: ' + company,
    'Role: ' + role,
    '',
    'Cover angles already chosen (lean into these):',
    coverAngles || '(none)',
    '',
    '---',
    '',
    'Job description:',
    '',
    jdText
  ].join('\n')

  const payload = {
    model: DEFAULT_MODEL,
    max_tokens: 3072,
    system: systemBlocks,
    messages: [{ role: 'user', content: userMessage }],
    thinking: { type: 'adaptive' },
    output_config: {
      effort: 'medium',
      format: { type: 'json_schema', schema: LIKELY_QUESTIONS_SCHEMA }
    }
  }

  const response = callClaude_(payload)
  const parsed = parseJsonResponse_(response)

  writeCachedJson_('questions_row_' + row, parsed)

  logActivity_({
    action: 'likely_questions',
    rowRef: 'row:' + row,
    model: DEFAULT_MODEL,
    usage: response.usage,
    notes: company
  })

  return parsed
}

function setChecklistItem(row, key, checked) {
  const all = readChecklistMap_()
  if (!all[row]) all[row] = {}
  all[row][key] = !!checked
  PropertiesService.getDocumentProperties().setProperty('prep_checklist', JSON.stringify(all))
  return { ok: true }
}

function saveInterviewNotes(row, notes) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_APPLICATIONS)
  const cols = getColumnMap_(sheet)
  const lock = LockService.getDocumentLock()
  try {
    lock.waitLock(5000)
    sheet.getRange(row, cols['Notes']).setValue(notes || '')
    sheet.getRange(row, cols['Last Touch']).setValue(new Date())
  } finally {
    try { lock.releaseLock() } catch (_) {}
  }
  return { ok: true }
}

function readChecklist_(row) {
  const all = readChecklistMap_()
  return all[row] || {}
}

function readChecklistMap_() {
  const raw = PropertiesService.getDocumentProperties().getProperty('prep_checklist')
  if (!raw) return {}
  try {
    return JSON.parse(raw)
  } catch (e) {
    return {}
  }
}

function readCachedJson_(key) {
  const raw = PropertiesService.getDocumentProperties().getProperty(key)
  if (!raw) return null
  try {
    return JSON.parse(raw)
  } catch (e) {
    return null
  }
}

function writeCachedJson_(key, value) {
  PropertiesService.getDocumentProperties().setProperty(key, JSON.stringify(value))
}
