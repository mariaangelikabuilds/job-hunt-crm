/**
 * Resume.gs
 *
 * The master resume lives as a Google Doc Angel owns. Its body is pulled
 * each call via DocumentApp, then snapshotted in CacheService for 6 hours
 * so prompt-cache hits work across consecutive scoring calls.
 *
 * When she edits the doc, she runs "Refresh resume cache" from the menu,
 * which clears the snapshot. Next call pays a one-time cache miss; every
 * call after that hits the cache normally.
 * Note: Resume caching uses Apps Script CacheService, which has a per-item size limit. 
 * Very large resume documents may fail to cache and should be shortened or handled with chunked caching.
 */

const RESUME_CACHE_KEY = 'resume_body_v1'
const RESUME_HASH_KEY = 'resume_hash_v1'
const RESUME_TTL_SECONDS = 21600

function getResumeText_() {
  const cache = CacheService.getScriptCache()
  const cached = cache.get(RESUME_CACHE_KEY)
  if (cached) return cached

  const docId = PropertiesService.getScriptProperties().getProperty('RESUME_DOC_ID')
  if (!docId) {
    throw new Error('RESUME_DOC_ID not set. Run Setup from the Job CRM menu.')
  }

  const body = DocumentApp.openById(docId).getBody().getText()
  if (!body || body.trim().length < 100) {
    throw new Error('Resume doc is empty or too short. Open the doc and confirm the body has real content.')
  }

  cache.put(RESUME_CACHE_KEY, body, RESUME_TTL_SECONDS)
  cache.put(RESUME_HASH_KEY, sha256Hex_(body), RESUME_TTL_SECONDS)
  return body
}

function refreshResumeCache_() {
  const cache = CacheService.getScriptCache()
  cache.removeAll([RESUME_CACHE_KEY, RESUME_HASH_KEY])
  const body = getResumeText_()
  SpreadsheetApp.getUi().alert(
    'Resume cache refreshed',
    'Pulled ' + body.length + ' characters from the resume doc. Next score call will repopulate the prompt cache.',
    SpreadsheetApp.getUi().ButtonSet.OK
  )
}

function getResumeHash_() {
  const cache = CacheService.getScriptCache()
  const hash = cache.get(RESUME_HASH_KEY)
  if (hash) return hash
  getResumeText_()
  return cache.get(RESUME_HASH_KEY)
}

function sha256Hex_(s) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
  return bytes.map(b => {
    const v = (b < 0 ? b + 256 : b).toString(16)
    return v.length === 1 ? '0' + v : v
  }).join('')
}
