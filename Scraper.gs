/**
 * Scraper.gs
 *
 * Fetches a JD URL and extracts plain JD text via Claude. Used when a row
 * has JD Link but no JD Text.
 *
 * Known limits:
 * - LinkedIn, Indeed serve Cloudflare-gated pages to non-browser fetches.
 *   These return 403 or HTML that's just a "you need JS" splash.
 * - Greenhouse, Lever, Workable, Ashby, custom careers pages usually
 *   return real HTML.
 *
 * Trims to 50K characters before passing to Claude (most JD pages have
 * heavy nav/footer chrome and a single page over that is usually a board
 * listing, not a single JD).
 */

const SCRAPE_USER_AGENT = 'Mozilla/5.0 (compatible; JobHuntCRM/1.0; +https://github.com/your-handle/job-hunt-crm)'
const SCRAPE_MAX_HTML_CHARS = 50000

function scrapeJD_(url, rowRef) {
  if (!url) return null

  let response
  try {
    response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        'User-Agent': SCRAPE_USER_AGENT,
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
      }
    })
  } catch (err) {
    console.warn('Scrape network failure for ' + url + ': ' + err.message)
    return null
  }

  const code = response.getResponseCode()
  if (code !== 200) {
    console.warn('Scrape ' + url + ' returned ' + code)
    return null
  }

  const contentType = response.getHeaders()['Content-Type'] || response.getHeaders()['content-type'] || ''
  if (contentType && contentType.indexOf('text/html') === -1 && contentType.indexOf('application/xhtml') === -1) {
    console.warn('Scrape ' + url + ' returned non-HTML: ' + contentType)
    return null
  }

  const html = response.getContentText()
  const cleaned = preCleanHtml_(html)

  try {
    const extracted = extractJDFromHtml_(cleaned, { rowRef: rowRef })
    return extracted
  } catch (err) {
    console.warn('Claude extraction failed for ' + url + ': ' + err.message)
    return null
  }
}

/**
 * Strip script, style, and noscript blocks before sending to Claude.
 * Saves tokens; the model doesn't need page JS or stylesheets.
 */
function preCleanHtml_(html) {
  let cleaned = html
    .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
    .replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, '')
    .replace(/<noscript\b[^<]*(?:(?!<\/noscript>)<[^<]*)*<\/noscript>/gi, '')
    .replace(/<!--[\s\S]*?-->/g, '')
  if (cleaned.length > SCRAPE_MAX_HTML_CHARS) {
    cleaned = cleaned.slice(0, SCRAPE_MAX_HTML_CHARS)
  }
  return cleaned
}
