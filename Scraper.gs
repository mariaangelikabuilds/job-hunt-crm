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
  if (!isPublicHttpUrl_(url)) {
    console.warn('Scrape rejected non-public URL: ' + url)
    return null
  }

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
 * Reject URLs that target localhost, private IP ranges, or cloud metadata
 * endpoints. UrlFetchApp runs in Google's sandbox so the practical SSRF
 * blast radius is small, but rejecting bad input is still the right shape.
 * Hostnames that resolve to private ranges via DNS slip past this string
 * check; that's the residual risk and is documented in the security audit.
 */
function isPublicHttpUrl_(url) {
  const s = String(url).trim()
  if (!/^https?:\/\//i.test(s)) return false
  let host
  try {
    const match = s.match(/^https?:\/\/([^\/?#]+)/i)
    if (!match) return false
    host = match[1].split('@').pop().split(':')[0].toLowerCase()
  } catch (_) {
    return false
  }
  if (!host) return false
  if (host === 'localhost' || host.endsWith('.localhost')) return false
  if (host === 'metadata.google.internal' || host === '169.254.169.254') return false
  if (/^127\./.test(host)) return false
  if (/^10\./.test(host)) return false
  if (/^192\.168\./.test(host)) return false
  if (/^169\.254\./.test(host)) return false
  if (/^172\.(1[6-9]|2\d|3[0-1])\./.test(host)) return false
  if (host === '0.0.0.0') return false
  if (/^::1$|^\[::1\]$|^\[fc/i.test(host)) return false
  return true
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
