/**
 * Claude.gs
 *
 * Raw HTTP client for the Anthropic Messages API. Apps Script V8 has no SDK
 * available, so we build payloads directly and POST via UrlFetchApp.
 *
 * Caching design: the system prompt is constructed deterministically from
 * the resume body (Resume.gs) and a frozen rubric. Same bytes every call =
 * prompt-cache hits on every score after the first. Verify by reading
 * `usage.cache_read_input_tokens` from the response.
 *
 * Model defaults: claude-opus-4-7 with adaptive thinking and effort=medium.
 * Override via the second arg to scoreJD_() or extractFromHtml_() if needed.
 */

const ANTHROPIC_API_URL = 'https://api.anthropic.com/v1/messages'
const ANTHROPIC_VERSION = '2023-06-01'
const DEFAULT_MODEL = 'claude-opus-4-7'

const SCORING_RUBRIC = [
  'You score job descriptions for fit against the candidate resume above.',
  '',
  'Rubric (sum to fit_score, 0-100):',
  '- skills_match (0-30): how many skills the candidate has that the JD asks for',
  '- role_seniority_fit (0-15): is this her level (Product Design Engineer-track), under, or over?',
  '- comp_signal (0-15): infer comp range from JD; flag misalignment with mid-senior PDE comp',
  '- location_fit (0-10): remote-friendly / Manila-friendly / timezone-friendly?',
  '- growth_signal (0-15): would this advance her toward Product Design Engineer at Linear/Vercel/Stripe-tier company?',
  '- red_flag_count (0-15): subtract for vague responsibilities, no comp listed, "rockstar/ninja/guru" language, role-stuffing',
  '',
  'Cover angles: write 3 specific, opinionated angles she could lead a cover letter or outreach with. Each angle must reference something concrete from her experience that maps to a specific JD requirement. Not generic.',
  '',
  'Red flags: list anything in the JD that should make her pause (unpaid trials, unrealistic scope, equity-only comp, ageist language, etc.). Empty list if none.',
  '',
  'Why score: one paragraph explaining the number. Auditable. No filler.',
  '',
  'Banned vocabulary in cover angles: "Spearheaded", "Synergized", "Leveraged", "Results-driven", "Proven track record", "Wearing many hats", "Passionate about", "Thrive in fast-paced environments", "Strong communication skills", "Rockstar", "Ninja", "Guru", "Excited to apply", "I am writing to". Every angle must read like a real person wrote it.',
  '',
  'Output JSON only. No prose, no preamble.'
].join('\n')

// Anthropic's structured output validator rejects JSON Schema constraint
// keywords (minimum/maximum/minItems/maxItems). Constraints like the
// 0-100 fit score range and exactly-3 angles are enforced in the system
// prompt rubric instead.
const SCORING_SCHEMA = {
  type: 'object',
  properties: {
    fit_score: { type: 'integer' },
    top_3_angles: {
      type: 'array',
      items: { type: 'string' }
    },
    red_flags: {
      type: 'array',
      items: { type: 'string' }
    },
    why_score: { type: 'string' }
  },
  required: ['fit_score', 'top_3_angles', 'red_flags', 'why_score'],
  additionalProperties: false
}

const EXTRACTION_SCHEMA = {
  type: 'object',
  properties: {
    company: { type: 'string' },
    role: { type: 'string' },
    jd_text: { type: 'string' }
  },
  required: ['company', 'role', 'jd_text'],
  additionalProperties: false
}

function scoreJDViaClaude_(jdText, opts) {
  const options = opts || {}
  const model = options.model || DEFAULT_MODEL
  const resume = getResumeText_()

  const systemBlocks = [
    {
      type: 'text',
      text: 'CANDIDATE RESUME (master, source of truth):\n\n' + resume + '\n\n---\n\n' + SCORING_RUBRIC,
      cache_control: { type: 'ephemeral' }
    }
  ]

  const payload = {
    model: model,
    max_tokens: 4096,
    system: systemBlocks,
    messages: [{ role: 'user', content: jdText }],
    thinking: { type: 'adaptive' },
    output_config: {
      effort: options.effort || 'medium',
      format: {
        type: 'json_schema',
        schema: SCORING_SCHEMA
      }
    }
  }

  const response = callClaude_(payload)
  const parsed = parseJsonResponse_(response)

  logActivity_({
    action: 'score',
    rowRef: options.rowRef || '',
    model: model,
    usage: response.usage,
    notes: 'fit_score=' + (parsed.fit_score || '?')
  })

  return {
    parsed: parsed,
    usage: response.usage,
    cached: (response.usage && response.usage.cache_read_input_tokens > 0) || false
  }
}

function extractJDFromHtml_(html, opts) {
  const options = opts || {}
  const model = options.model || DEFAULT_MODEL
  const truncated = html.length > 50000 ? html.slice(0, 50000) : html

  const systemBlocks = [
    {
      type: 'text',
      text: [
        'You extract structured job posting data from raw HTML.',
        '',
        'Return JSON with three fields:',
        '- company: company name as it appears on the page',
        '- role: the role title',
        '- jd_text: the full job description body, plain text, with section headings preserved as plain prose. Strip HTML tags. Preserve bullet points as "- " prefixed lines.',
        '',
        'If a field cannot be found, return an empty string for that field. Do not invent.',
        '',
        'Output JSON only.'
      ].join('\n'),
      cache_control: { type: 'ephemeral' }
    }
  ]

  const payload = {
    model: model,
    max_tokens: 8192,
    system: systemBlocks,
    messages: [{ role: 'user', content: truncated }],
    thinking: { type: 'adaptive' },
    output_config: {
      effort: 'low',
      format: {
        type: 'json_schema',
        schema: EXTRACTION_SCHEMA
      }
    }
  }

  const response = callClaude_(payload)
  const parsed = parseJsonResponse_(response)

  logActivity_({
    action: 'extract_jd',
    rowRef: options.rowRef || '',
    model: model,
    usage: response.usage,
    notes: 'company=' + (parsed.company || '?')
  })

  return parsed
}

function callClaude_(payload) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY')
  if (!apiKey) {
    throw new Error('ANTHROPIC_API_KEY not set. Run Setup from the Job CRM menu.')
  }

  const fetchOptions = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': ANTHROPIC_VERSION
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  }

  let attempt = 0
  let lastError = null
  while (attempt < 3) {
    const response = UrlFetchApp.fetch(ANTHROPIC_API_URL, fetchOptions)
    const code = response.getResponseCode()
    const body = response.getContentText()

    if (code >= 200 && code < 300) {
      return JSON.parse(body)
    }

    if (code === 429 || code === 529 || code >= 500) {
      lastError = code + ' ' + body.slice(0, 500)
      const backoffMs = Math.pow(2, attempt) * 1500
      Utilities.sleep(backoffMs)
      attempt++
      continue
    }

    throw new Error('Anthropic API error ' + code + ': ' + body.slice(0, 1000))
  }

  throw new Error('Anthropic API retried ' + attempt + ' times. Last error: ' + lastError)
}

function parseJsonResponse_(response) {
  if (!response || !response.content || response.content.length === 0) {
    throw new Error('Empty response from Claude')
  }
  const textBlock = response.content.find(b => b.type === 'text')
  if (!textBlock) {
    throw new Error('No text block in Claude response. Stop reason: ' + (response.stop_reason || 'unknown'))
  }
  try {
    return JSON.parse(textBlock.text)
  } catch (err) {
    throw new Error('Could not parse JSON from Claude response: ' + textBlock.text.slice(0, 300))
  }
}
