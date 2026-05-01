/**
 * LinkedInImport.gs
 *
 * "Generate a master resume from LinkedIn" feature. LinkedIn blocks bot
 * fetches with HTTP 999, so this never tries to scrape. Instead it accepts
 * three input paths, all of which work around the anti-bot wall:
 *
 * 1. Paste text - user copies the visible text from their own LinkedIn
 *    profile (logged in) and pastes into a textarea.
 * 2. Upload PDF - user uses LinkedIn's "Save to PDF" button on their own
 *    profile, uploads to Drive, gives us the file ID.
 * 3. Upload screenshot - user takes one or more screenshots of their own
 *    LinkedIn view, uploads to Drive, gives us the file IDs. Claude vision
 *    extracts the content.
 *
 * Output: a new Google Doc in the Drive parent folder, structured as a
 * master resume in Markdown. Angel reviews, edits, then sets the Doc ID
 * as RESUME_DOC_ID via the onboarding wizard. From that point on, every
 * scoring + tailoring + cover letter call uses this resume as the cache
 * anchor and source of truth.
 */

const RESUME_GENERATION_SYSTEM_PROMPT = [
  'You convert LinkedIn profile content into a master resume in Markdown.',
  '',
  'Output structure (follow exactly):',
  '## Header',
  '(Name on first line. Headline / target role on second line. Location on third. Email and link line on fourth.)',
  '',
  '## About',
  '(One paragraph, two to three sentences. The candidate\'s positioning statement. Pull from the LinkedIn About section. Compress, do not pad.)',
  '',
  '## Experience',
  '(Most recent role first. For each role:',
  '### Role title at Company',
  '*Dates · Location · Employment type*',
  '- Bullets. Each bullet must have a real metric, scope, named system, or named outcome from the LinkedIn content. No metric-free bullets.',
  ')',
  '',
  '## Selected projects',
  '(Optional. Only if LinkedIn has a Featured / Projects section with concrete artifacts.)',
  '',
  '## Skills',
  '(Comma-separated. Pull from the LinkedIn Skills, Top Skills, and Stack lines. Group as needed. No more than 25 skills total. No "Microsoft Office" or "Communication".)',
  '',
  '## Education',
  '(Per entry: degree, institution, dates.)',
  '',
  '## Certifications',
  '(Per entry: name, issuer, date.)',
  '',
  'Hard rules:',
  '1. NEVER invent. Every claim must trace back to the LinkedIn input. If a field is missing, write "TODO: [what is missing]" instead of inventing.',
  '2. Banned vocabulary: "Spearheaded", "Synergized", "Leveraged", "Results-driven", "Proven track record", "Wearing many hats", "Passionate about", "Thrive in fast-paced environments", "Strong communication skills", "Rockstar", "Ninja", "Guru", "Architected" (overused), "Pioneered" (overused).',
  '3. Banned punctuation: em dashes. Use commas, periods, or rewrite the sentence.',
  '4. Every Experience bullet has a metric, named system, or scope. Bullets without those get cut.',
  '5. Compress LinkedIn descriptions. Do not copy multi-paragraph bullet bodies verbatim. One tight sentence per bullet.',
  '6. Use the candidate\'s own headline as the role positioning at the top of Header (target role), not their current job title, if those differ. Career-pivot framing.',
  '',
  'Output Markdown only. No JSON, no preamble, no commentary.'
].join('\n')

/**
 * Menu entry: opens the LinkedIn import dialog.
 */
function showLinkedInImportDialog_() {
  const html = HtmlService.createTemplateFromFile('LinkedInDialog')
    .evaluate()
    .setWidth(560)
    .setHeight(620)
  SpreadsheetApp.getUi().showModalDialog(html, 'Generate resume from LinkedIn')
}

/**
 * Path 1: paste text.
 */
function generateResumeFromLinkedInText(pastedText) {
  if (!pastedText || pastedText.trim().length < 200) {
    throw new Error('Pasted content is too short. Paste the full visible text of your LinkedIn profile.')
  }

  const userContent = [
    { type: 'text', text: 'LinkedIn profile content (pasted text):\n\n' + pastedText }
  ]

  return runResumeGeneration_(userContent, 'paste')
}

/**
 * Path 2: PDF from Drive (LinkedIn "Save to PDF" output).
 */
function generateResumeFromLinkedInPdf(fileIdOrUrl) {
  const fileId = parseDriveId_(fileIdOrUrl, 'PDF file')
  let blob
  try {
    blob = DriveApp.getFileById(fileId).getBlob()
  } catch (err) {
    throw new Error('Could not read the PDF from Drive. Confirm the file ID is correct and you own the file. Detail: ' + err.message)
  }

  const contentType = blob.getContentType()
  if (contentType !== 'application/pdf') {
    throw new Error('That file is not a PDF. Got Content-Type: ' + contentType + '. Use LinkedIn\'s "Save to PDF" button on your profile.')
  }

  const base64 = Utilities.base64Encode(blob.getBytes())

  const userContent = [
    {
      type: 'document',
      source: {
        type: 'base64',
        media_type: 'application/pdf',
        data: base64
      }
    },
    { type: 'text', text: 'Convert the attached LinkedIn PDF export into a master resume Markdown.' }
  ]

  return runResumeGeneration_(userContent, 'pdf:' + fileId)
}

/**
 * Path 3: one or more screenshots from Drive. Accepts a comma-separated
 * list of file IDs or URLs. Up to 6 images per call (Claude content block
 * cap is generous; more than 6 means the user is screenshotting too much
 * detail, suggest using PDF instead).
 */
function generateResumeFromLinkedInScreenshots(fileIdsOrUrls) {
  if (!fileIdsOrUrls) throw new Error('No screenshot IDs provided.')
  const items = String(fileIdsOrUrls).split(/[,\n]/).map(s => s.trim()).filter(Boolean)
  if (items.length === 0) throw new Error('No screenshot IDs provided.')
  if (items.length > 6) throw new Error('Too many screenshots (' + items.length + '). Use 6 or fewer, or use the PDF path instead.')

  const userContent = []
  for (let i = 0; i < items.length; i++) {
    const fileId = parseDriveId_(items[i], 'Screenshot ' + (i + 1))
    let blob
    try {
      blob = DriveApp.getFileById(fileId).getBlob()
    } catch (err) {
      throw new Error('Could not read screenshot ' + (i + 1) + ' from Drive: ' + err.message)
    }

    const contentType = blob.getContentType()
    if (contentType.indexOf('image/') !== 0) {
      throw new Error('Screenshot ' + (i + 1) + ' is not an image. Got: ' + contentType)
    }

    const base64 = Utilities.base64Encode(blob.getBytes())
    userContent.push({
      type: 'image',
      source: {
        type: 'base64',
        media_type: contentType,
        data: base64
      }
    })
  }

  userContent.push({
    type: 'text',
    text: 'Convert the attached LinkedIn screenshot' + (items.length === 1 ? '' : 's') + ' into a master resume Markdown. The screenshots show the candidate\'s own LinkedIn profile from a logged-in view.'
  })

  return runResumeGeneration_(userContent, 'screenshot:count:' + items.length)
}

/**
 * Common Claude call + Doc creation path for all three input modes.
 */
function runResumeGeneration_(userContent, sourceTag) {
  const systemBlocks = [
    {
      type: 'text',
      text: RESUME_GENERATION_SYSTEM_PROMPT,
      cache_control: { type: 'ephemeral' }
    }
  ]

  const payload = {
    model: DEFAULT_MODEL,
    max_tokens: 8192,
    system: systemBlocks,
    messages: [{ role: 'user', content: userContent }],
    thinking: { type: 'adaptive' },
    output_config: { effort: 'high' }
  }

  const response = callClaude_(payload)
  const markdown = extractTextResponse_(response)

  const doc = createMasterResumeDoc_(markdown)

  logActivity_({
    action: 'linkedin_import',
    rowRef: 'source:' + sourceTag,
    model: DEFAULT_MODEL,
    usage: response.usage,
    notes: 'doc=' + doc.id
  })

  PropertiesService.getScriptProperties().setProperty('LAST_GENERATED_RESUME_DOC_ID', doc.id)

  return {
    url: doc.url,
    id: doc.id,
    has_todos: markdown.indexOf('TODO:') !== -1,
    cached: (response.usage && response.usage.cache_read_input_tokens > 0) || false
  }
}

/**
 * Writes the Markdown to a new Doc inside the Drive parent folder.
 * Filename includes the date so multiple regeneration attempts don't
 * collide.
 */
function createMasterResumeDoc_(markdown) {
  const parentId = PropertiesService.getScriptProperties().getProperty('DRIVE_PARENT_FOLDER_ID')
  if (!parentId) {
    throw new Error('DRIVE_PARENT_FOLDER_ID not set. Run the onboarding wizard first.')
  }

  const filename = 'Master Resume - ' + formatDateShort_(new Date())
  const doc = DocumentApp.create(filename)
  const docFile = DriveApp.getFileById(doc.getId())
  docFile.moveTo(DriveApp.getFolderById(parentId))

  const body = doc.getBody()
  body.clear()

  body.appendParagraph('Master resume')
    .setHeading(DocumentApp.ParagraphHeading.TITLE)

  body.appendParagraph('Generated from LinkedIn on ' + formatDateShort_(new Date()) + '. Review carefully before pointing RESUME_DOC_ID at this Doc. Replace any TODO: lines with real content.')
    .setForegroundColor('#5C5A56')
    .setFontSize(10)
    .setItalic(true)

  body.appendParagraph('')

  appendMarkdownToDocBody_(body, markdown)

  doc.saveAndClose()
  return { id: doc.getId(), url: doc.getUrl() }
}
