/**
 * Slack.gs
 *
 * Two-way Slack integration for the Job Hunt CRM.
 *
 * Outbound (daily nudge): if SLACK_WEBHOOK_URL is set and SLACK_NUDGE_TARGET
 * includes "slack", dailyNudge_() forks here. Posts a digest to the
 * configured channel via incoming webhook. Block Kit, with a top-level
 * text fallback for screen readers and notifications.
 *
 * Inbound (slash command): WebApp.gs doPost routes Slack form posts here.
 * Supported subcommands:
 *
 *   /jobcrm log Company | Role | https://jd.url
 *
 * The slash command never scores. Scoring is a 5-10 second Claude call;
 * Slack expects a response within 3 seconds. Logging is fast (sheet append),
 * so it fits inside the budget. Scoring is done later from the sidebar.
 *
 * Why no HMAC signature verification: Apps Script doPost does not expose
 * request headers, so X-Slack-Signature cannot be checked. We fall back to
 * Slack's verification token (in the request body) plus a team_id allowlist.
 * This is acceptable for a single-user tool; documented in README.
 */

const SLACK_MAX_NUDGE_ROWS = 10

/**
 * Posts the daily nudge to Slack. Called from Gmail.gs when the user has
 * opted to receive Slack nudges (SLACK_NUDGE_TARGET in {slack, both}).
 *
 * No-op when there is nothing to send (we never post empty digests).
 */
function postNudgeToSlack_(overdue) {
  if (!overdue || overdue.length === 0) return

  const props = PropertiesService.getScriptProperties()
  const webhook = props.getProperty('SLACK_WEBHOOK_URL')
  if (!webhook) return

  const payload = buildNudgeBlocks_(overdue)

  const response = UrlFetchApp.fetch(webhook, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  })

  const code = response.getResponseCode()
  if (code < 200 || code >= 300) {
    console.warn('Slack webhook failed: ' + code + ' ' + response.getContentText().slice(0, 200))
    return
  }

  logActivity_({
    action: 'slack_nudge_sent',
    rowRef: 'count:' + overdue.length,
    notes: overdue.slice(0, 5).map(o => o.company).join(', ')
  })
}

function buildNudgeBlocks_(overdue) {
  const top = overdue.slice(0, SLACK_MAX_NUDGE_ROWS)
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName(SHEET_APPLICATIONS)
  const gid = sheet ? sheet.getSheetId() : 0
  const sheetUrl = ss.getUrl()

  const headerLine = top.length === 1
    ? '1 follow-up due.'
    : top.length + ' follow-ups due.'

  const blocks = [
    {
      type: 'header',
      text: { type: 'plain_text', text: 'Follow-ups due', emoji: false }
    },
    {
      type: 'section',
      text: {
        type: 'mrkdwn',
        text: headerLine + ' These applications have not been touched in 7+ days.'
      }
    },
    { type: 'divider' }
  ]

  for (const item of top) {
    const days = Math.floor((Date.now() - item.lastTouch.getTime()) / (1000 * 60 * 60 * 24))
    const rowUrl = sheetUrl + '#gid=' + gid + '&range=A' + item.rowIndex
    const lines = [
      '*' + escapeMrkdwn_(item.company) + '*  ' + escapeMrkdwn_(item.role || '?'),
      '_Last touch ' + days + ' days ago_'
    ]
    blocks.push({
      type: 'section',
      text: { type: 'mrkdwn', text: lines.join('\n') },
      accessory: {
        type: 'button',
        text: { type: 'plain_text', text: 'Open the row', emoji: false },
        url: rowUrl,
        action_id: 'open_row_' + item.rowIndex
      }
    })
  }

  if (overdue.length > top.length) {
    blocks.push({
      type: 'context',
      elements: [{
        type: 'mrkdwn',
        text: '+' + (overdue.length - top.length) + ' more not shown. Open the sheet for the full list.'
      }]
    })
  }

  const fallbackText = headerLine + ' Open the sheet to see them all.'

  return { text: fallbackText, blocks: blocks }
}

/**
 * Slash command entry point. Called by WebApp.gs doPost when the request
 * looks like a Slack form post (e.parameter.command and e.parameter.token
 * are present).
 *
 * Returns a ContentService TextOutput. Slack expects JSON with a
 * `response_type` of "ephemeral" (only the user sees it) or "in_channel".
 * We default to ephemeral for confirmations; errors are also ephemeral.
 */
function handleSlackSlashCommand_(e) {
  const props = PropertiesService.getScriptProperties()
  const expectedToken = props.getProperty('SLACK_VERIFICATION_TOKEN')
  const expectedTeam = props.getProperty('SLACK_TEAM_ID')

  if (expectedToken && e.parameter.token !== expectedToken) {
    return slackEphemeral_('Verification token did not match. The Slack app and the script are out of sync.')
  }
  if (expectedTeam && e.parameter.team_id !== expectedTeam) {
    return slackEphemeral_('This Slack workspace is not allow-listed.')
  }

  const text = (e.parameter.text || '').trim()
  if (!text) {
    return slackEphemeral_('Usage: `/jobcrm log Company | Role | https://jd.url`')
  }

  const space = text.indexOf(' ')
  const subcommand = space === -1 ? text.toLowerCase() : text.slice(0, space).toLowerCase()
  const rest = space === -1 ? '' : text.slice(space + 1).trim()

  if (subcommand === 'log') {
    return handleSlackLog_(rest, e.parameter.user_name || 'unknown')
  }

  return slackEphemeral_('Unknown subcommand `' + subcommand + '`. Try `/jobcrm log Company | Role | https://jd.url`.')
}

function handleSlackLog_(argString, userName) {
  if (!argString) {
    return slackEphemeral_('Usage: `/jobcrm log Company | Role | https://jd.url`')
  }

  const parts = argString.split('|').map(s => s.trim())
  const company = parts[0] || ''
  const role = parts[1] || ''
  const jdLink = parts[2] || ''

  if (!company) {
    return slackEphemeral_('Company is required. Usage: `/jobcrm log Company | Role | https://jd.url`')
  }

  let result
  try {
    result = quickLogApplication({
      company: company,
      role: role,
      jd_link: jdLink,
      status: 'Saved',
      notes: 'Logged from Slack by ' + userName
    })
  } catch (err) {
    return slackEphemeral_('Could not log: ' + err.message)
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const targetSheet = ss.getSheetByName(SHEET_APPLICATIONS)
  const gid = targetSheet ? targetSheet.getSheetId() : 0
  const rowUrl = ss.getUrl() + '#gid=' + gid + '&range=A' + result.row

  logActivity_({
    action: 'slack_log',
    rowRef: 'row:' + result.row,
    notes: company + ' / ' + role + ' (by ' + userName + ')'
  })

  const summary = role
    ? 'Logged ' + company + ' ' + role + ' to row ' + result.row + '.'
    : 'Logged ' + company + ' to row ' + result.row + '.'

  return slackEphemeralWithLink_(summary, 'Open the row', rowUrl)
}

function slackEphemeral_(message) {
  const body = {
    response_type: 'ephemeral',
    text: message
  }
  return ContentService
    .createTextOutput(JSON.stringify(body))
    .setMimeType(ContentService.MimeType.JSON)
}

function slackEphemeralWithLink_(message, linkText, linkUrl) {
  const body = {
    response_type: 'ephemeral',
    text: message,
    blocks: [
      {
        type: 'section',
        text: { type: 'mrkdwn', text: message }
      },
      {
        type: 'actions',
        elements: [{
          type: 'button',
          text: { type: 'plain_text', text: linkText, emoji: false },
          url: linkUrl
        }]
      }
    ]
  }
  return ContentService
    .createTextOutput(JSON.stringify(body))
    .setMimeType(ContentService.MimeType.JSON)
}

/**
 * Slack mrkdwn escaping. Anchored characters that have meaning in mrkdwn
 * are escaped so company/role text with `&`, `<`, `>` does not break
 * rendering.
 */
function escapeMrkdwn_(s) {
  if (!s) return ''
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
}
