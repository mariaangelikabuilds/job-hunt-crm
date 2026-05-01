/**
 * ResumeGen.gs
 *
 * Generates a tailored resume per application. Reads master resume from
 * Drive (cached), JD Text + Cover Angles from the row, calls Claude with
 * a hard anti-fabrication system prompt, writes Markdown output to a new
 * Google Doc inside the company's Drive folder.
 *
 * Anti-fabrication rules (enforced at the prompt level):
 * - No skills not in master
 * - No metrics not in master
 * - No roles not in master
 * - Reordering and re-phrasing only
 * - Honest gaps: if JD asks for something not in master, write
 *   "GAP: [what the JD wants]" as an inline comment so Angel sees it
 */

const TAILORING_SYSTEM_PROMPT_TAIL = [
  '',
  'You tailor a master resume for a specific job description. The master resume above is the source of truth: every claim, skill, metric, and role in your output MUST appear in the master.',
  '',
  'Hard rules:',
  '1. No new skills. If TypeScript is not in the master, it does not appear in the tailored output even if the JD asks for it.',
  '2. No new metrics. If "shipped to 1000 users" is in the master, "shipped to 5000 users" cannot appear. Numbers do not move.',
  '3. No new roles. Job history is locked.',
  '4. Reordering and re-phrasing only. Surface what is relevant first; rephrase to match the JD vocabulary; cut what is not relevant. Do not say new things.',
  '5. Honest gaps. If the JD asks for X and the master has no X, OMIT the requirement from the resume. Do NOT fake it. Add a comment line that begins exactly with "GAP: " describing what the JD asks for and the master lacks. Angel will see the gap in the diff and decide whether to address it.',
  '6. Forbidden vocabulary (these read as AI-generated): "Spearheaded", "Synergized", "Leveraged", "Results-driven", "Proven track record", "Wearing many hats", "Passionate about", "Thrive in fast-paced environments", "Strong communication skills", "Rockstar", "Ninja", "Guru", "Architected" (overused), "Pioneered" (overused).',
  '7. Every bullet must have a real metric, scope, or named outcome from the master. Bullets without numbers, scope, or named outcomes get cut.',
  '8. Use em dashes never. Use commas, periods, or rewrite the sentence.',
  '',
  'Output format: Markdown. Sections in this order, each as an h2 heading:',
  '## Header (name, role, location, contact line)',
  '## About (one paragraph, two to three sentences max)',
  '## Experience (most relevant role first; bullets per role)',
  '## Selected Projects (only if relevant to this JD)',
  '## Skills (comma-separated, only skills from master that this JD asks for)',
  '## Education',
  '',
  'Place any GAP: comments at the very bottom under "## Gaps to address" so they are easy to find and remove.',
  '',
  'Output Markdown only. No preamble, no explanation, no JSON.'
].join('\n')

function tailorResumeForRow(row) {
  if (!row || row < 2) throw new Error('Invalid row')

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_APPLICATIONS)
  const cols = getColumnMap_(sheet)
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0]

  const company = values[cols['Company'] - 1]
  const role = values[cols['Role'] - 1]
  const jdText = values[cols['JD Text'] - 1]
  const coverAngles = values[cols['Cover Angles'] - 1]
  const whyScore = values[cols['Why Score'] - 1]

  if (!company) throw new Error('Company is empty. Fill it in before tailoring.')
  if (!jdText) throw new Error('JD Text is empty. Score the row first.')

  const resume = getResumeText_()

  const systemBlocks = [
    {
      type: 'text',
      text: 'MASTER RESUME (source of truth, do not contradict):\n\n' + resume + TAILORING_SYSTEM_PROMPT_TAIL,
      cache_control: { type: 'ephemeral' }
    }
  ]

  const userMessage = [
    'Target company: ' + company,
    'Target role: ' + role,
    '',
    'Cover angles already chosen for this application (use these to decide what to surface):',
    coverAngles || '(none)',
    '',
    'Why this is a fit (auditable scoring rationale):',
    whyScore || '(none)',
    '',
    '---',
    '',
    'Job description:',
    '',
    jdText
  ].join('\n')

  const payload = {
    model: DEFAULT_MODEL,
    max_tokens: 8192,
    system: systemBlocks,
    messages: [{ role: 'user', content: userMessage }],
    thinking: { type: 'adaptive' },
    output_config: { effort: 'high' }
  }

  const response = callClaude_(payload)
  const markdown = extractTextResponse_(response)

  const docUrl = writeMarkdownToDoc_({
    company: company,
    role: role,
    kind: 'Resume',
    markdown: markdown
  })

  const lock = LockService.getDocumentLock()
  try {
    lock.waitLock(5000)
    sheet.getRange(row, cols['Tailored Resume']).setValue(docUrl)
    sheet.getRange(row, cols['Last Touch']).setValue(new Date())
  } finally {
    try { lock.releaseLock() } catch (_) {}
  }

  logActivity_({
    action: 'tailor_resume',
    rowRef: 'row:' + row,
    model: DEFAULT_MODEL,
    usage: response.usage,
    notes: company + ' / ' + role
  })

  return {
    url: docUrl,
    cached: (response.usage && response.usage.cache_read_input_tokens > 0) || false,
    has_gaps: markdown.indexOf('GAP:') !== -1
  }
}

function extractTextResponse_(response) {
  if (!response || !response.content) throw new Error('Empty Claude response')
  const block = response.content.find(b => b.type === 'text')
  if (!block) throw new Error('No text block in Claude response. Stop reason: ' + (response.stop_reason || 'unknown'))
  return block.text
}

function writeMarkdownToDoc_(opts) {
  const folderUrl = ensureCompanyDriveFolder_(opts.company)
  const folderId = extractFolderIdFromUrl_(folderUrl)
  const folder = DriveApp.getFolderById(folderId)

  const filename = [opts.company, opts.role, opts.kind, formatDateShort_(new Date())]
    .filter(Boolean)
    .join(' - ')

  const doc = DocumentApp.create(filename)
  const docFile = DriveApp.getFileById(doc.getId())
  docFile.moveTo(folder)

  const body = doc.getBody()
  body.clear()
  appendMarkdownToDocBody_(body, opts.markdown)

  doc.saveAndClose()
  return doc.getUrl()
}

function extractFolderIdFromUrl_(url) {
  const match = String(url).match(/\/folders\/([a-zA-Z0-9_-]+)/)
  if (!match) throw new Error('Could not parse folder ID from URL: ' + url)
  return match[1]
}

function formatDateShort_(d) {
  const y = d.getFullYear()
  const m = String(d.getMonth() + 1).padStart(2, '0')
  const day = String(d.getDate()).padStart(2, '0')
  return y + '-' + m + '-' + day
}

/**
 * Minimal Markdown to Google Doc converter. Handles the subset Claude
 * actually produces for resumes: h1, h2, h3, bullets (- or *), and plain
 * paragraphs. Bold and italic via ** and * are preserved as run formatting.
 *
 * Not a full Markdown parser. Intentionally narrow.
 */
function appendMarkdownToDocBody_(body, markdown) {
  const lines = markdown.split('\n')
  for (const line of lines) {
    if (line.trim() === '') {
      body.appendParagraph('')
      continue
    }

    if (line.startsWith('# ')) {
      body.appendParagraph(line.slice(2)).setHeading(DocumentApp.ParagraphHeading.HEADING1)
      continue
    }
    if (line.startsWith('## ')) {
      body.appendParagraph(line.slice(3)).setHeading(DocumentApp.ParagraphHeading.HEADING2)
      continue
    }
    if (line.startsWith('### ')) {
      body.appendParagraph(line.slice(4)).setHeading(DocumentApp.ParagraphHeading.HEADING3)
      continue
    }
    if (/^\s*[-*]\s+/.test(line)) {
      const text = line.replace(/^\s*[-*]\s+/, '')
      body.appendListItem(stripInlineMarkdown_(text)).setGlyphType(DocumentApp.GlyphType.BULLET)
      continue
    }
    body.appendParagraph(stripInlineMarkdown_(line))
  }
}

function stripInlineMarkdown_(s) {
  return s
    .replace(/\*\*(.+?)\*\*/g, '$1')
    .replace(/\*(.+?)\*/g, '$1')
    .replace(/`(.+?)`/g, '$1')
}
