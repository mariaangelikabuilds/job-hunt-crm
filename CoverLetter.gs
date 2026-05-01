/**
 * CoverLetter.gs
 *
 * Drafts a cover letter for a row using the picked cover angle as the spine.
 * Anti-AI rules at the prompt level: must reference one concrete verifiable
 * thing about the company (real product, recent launch, named team member,
 * public statement). If Claude cannot ground in a specific, it must say so
 * inline rather than invent.
 *
 * Output written to a new Google Doc inside the company's Drive folder.
 */

const COVER_LETTER_SYSTEM_PROMPT_TAIL = [
  '',
  'You draft a cover letter for a specific application. The candidate resume above is the source of truth: every claim, skill, metric, and project you reference MUST appear in the master.',
  '',
  'Structure: three short paragraphs, 250 to 400 words total. Not five paragraphs. Not a wall of text.',
  '',
  '- Paragraph 1 (the hook): lead with the picked cover angle. Make it specific to this role at this company. Reference one real, verifiable thing about the company (a real product, a recent launch you can name, a leadership decision, a public statement they made). If you cannot ground in a real specific from the JD or your training data, say so explicitly with the line "[VERIFY: cite a specific public ${company} thing here, e.g. recent launch / leadership move / public statement]" so Angel knows to fill it in. Do NOT fabricate company specifics.',
  '- Paragraph 2 (the bridge): two to three sentences pulling concrete experience from the master resume that ladders up to the picked angle. Real metrics only. No vague claims.',
  '- Paragraph 3 (the close): direct invitation to talk. One sentence. No "thank you for your consideration", no "I look forward to hearing from you".',
  '',
  'Banned openers: "I am writing to apply for", "I am excited about this opportunity because", "I believe I am a strong fit because", "My passion for X aligns with your mission", "Excited to apply".',
  '',
  'Banned closers: "Thank you for your consideration", "I look forward to hearing from you", "Please find my resume attached", "I would welcome the opportunity to discuss".',
  '',
  'Banned vocabulary: "Spearheaded", "Synergized", "Leveraged", "Results-driven", "Proven track record", "Wearing many hats", "Passionate about", "Thrive in fast-paced environments", "Strong communication skills", "Rockstar", "Ninja", "Guru", "Driven by", "Empower", "Foster".',
  '',
  'Banned punctuation: em dashes. Use commas, periods, or rewrite the sentence.',
  '',
  'Tone: direct, specific, confident. Like a real person wrote it. Match the candidate voice from the master resume above. Do not flatter the company or the reader.',
  '',
  'Output the letter body only. No subject line, no addressing block, no signature, no postscript. Plain prose. No JSON, no preamble, no explanation.'
].join('\n')

function draftCoverLetterForRow(row, angleIndex) {
  if (!row || row < 2) throw new Error('Invalid row')
  if (typeof angleIndex !== 'number' || angleIndex < 0) angleIndex = 0

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_APPLICATIONS)
  const cols = getColumnMap_(sheet)
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0]

  const company = values[cols['Company'] - 1]
  const role = values[cols['Role'] - 1]
  const jdText = values[cols['JD Text'] - 1]
  const coverAnglesRaw = values[cols['Cover Angles'] - 1]
  const whyScore = values[cols['Why Score'] - 1]

  if (!company) throw new Error('Company is empty.')
  if (!jdText) throw new Error('JD Text is empty. Score the row first.')

  const angles = String(coverAnglesRaw || '').split('\n').filter(Boolean)
  if (angles.length === 0) throw new Error('No cover angles. Score the row first.')

  const pickedAngle = angles[Math.min(angleIndex, angles.length - 1)]

  const resume = getResumeText_()

  const systemText = 'MASTER RESUME (source of truth):\n\n' + resume +
    COVER_LETTER_SYSTEM_PROMPT_TAIL.replace(/\$\{company\}/g, company)

  const systemBlocks = [
    {
      type: 'text',
      text: systemText,
      cache_control: { type: 'ephemeral' }
    }
  ]

  const userMessage = [
    'Target company: ' + company,
    'Target role: ' + role,
    '',
    'Picked cover angle (the spine of the letter):',
    pickedAngle,
    '',
    'Scoring rationale (for context, do not paraphrase verbatim):',
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
    max_tokens: 2048,
    system: systemBlocks,
    messages: [{ role: 'user', content: userMessage }],
    thinking: { type: 'adaptive' },
    output_config: { effort: 'medium' }
  }

  const response = callClaude_(payload)
  const letterBody = extractTextResponse_(response)

  const docUrl = writeCoverLetterToDoc_({
    company: company,
    role: role,
    body: letterBody,
    pickedAngle: pickedAngle
  })

  const lock = LockService.getDocumentLock()
  try {
    lock.waitLock(5000)
    sheet.getRange(row, cols['Cover Letter']).setValue(docUrl)
    sheet.getRange(row, cols['Picked Angle']).setValue(angleIndex + 1)
    sheet.getRange(row, cols['Last Touch']).setValue(new Date())
  } finally {
    try { lock.releaseLock() } catch (_) {}
  }

  logActivity_({
    action: 'draft_cover_letter',
    rowRef: 'row:' + row,
    model: DEFAULT_MODEL,
    usage: response.usage,
    notes: company + ' angle=' + (angleIndex + 1)
  })

  return {
    url: docUrl,
    cached: (response.usage && response.usage.cache_read_input_tokens > 0) || false,
    needs_verify: letterBody.indexOf('[VERIFY:') !== -1
  }
}

function writeCoverLetterToDoc_(opts) {
  const folderUrl = ensureCompanyDriveFolder_(opts.company)
  const folderId = extractFolderIdFromUrl_(folderUrl)
  const folder = DriveApp.getFolderById(folderId)

  const filename = [opts.company, opts.role, 'Cover Letter', formatDateShort_(new Date())]
    .filter(Boolean)
    .join(' - ')

  const doc = DocumentApp.create(filename)
  const docFile = DriveApp.getFileById(doc.getId())
  docFile.moveTo(folder)

  const body = doc.getBody()
  body.clear()

  body.appendParagraph(opts.company + ' / ' + opts.role)
    .setForegroundColor('#5C5A56')
    .setFontSize(10)

  body.appendParagraph('Cover angle used: ' + opts.pickedAngle)
    .setForegroundColor('#5C5A56')
    .setFontSize(10)
    .setItalic(true)

  body.appendParagraph('')

  const paragraphs = opts.body.split(/\n\s*\n/).map(p => p.trim()).filter(Boolean)
  for (const p of paragraphs) {
    body.appendParagraph(stripInlineMarkdown_(p))
  }

  body.appendParagraph('')
  body.appendParagraph('Edit before sending. Replace any [VERIFY: ...] placeholders with real specifics.')
    .setForegroundColor('#A23B26')
    .setFontSize(10)
    .setItalic(true)

  doc.saveAndClose()
  return doc.getUrl()
}
