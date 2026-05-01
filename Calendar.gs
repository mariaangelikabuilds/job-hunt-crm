/**
 * Calendar.gs
 *
 * Auto-creates a Calendar event placeholder when a row's Status flips to
 * "Interview Scheduled". Uses the Next Action column as the date hint.
 *
 * If Next Action doesn't parse as a date, the trigger writes a note instead
 * of failing silently. Angel can edit the placeholder event later.
 */

function onStatusEdit_(e) {
  if (!e || !e.range) return

  const sheet = e.range.getSheet()
  if (sheet.getName() !== SHEET_APPLICATIONS) return

  const cols = getColumnMap_(sheet)
  const editedCol = e.range.getColumn()
  if (editedCol !== cols['Status']) return

  const row = e.range.getRow()
  if (row < 2) return

  const status = e.range.getValue()
  if (status !== 'Interview Scheduled') return

  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0]
  const company = values[cols['Company'] - 1]
  const role = values[cols['Role'] - 1]
  const nextAction = values[cols['Next Action'] - 1]

  const eventDate = parseInterviewDate_(nextAction)
  if (!eventDate) {
    sheet.getRange(row, cols['Notes']).setValue('Set the interview date in Next Action (e.g. "2026-05-12 14:00") so a Calendar event can be created.')
    return
  }

  createInterviewEvent_(eventDate, company, role, row)
}

function parseInterviewDate_(input) {
  if (!input) return null
  if (input instanceof Date) return input
  const parsed = new Date(input)
  if (isNaN(parsed.getTime())) return null
  return parsed
}

function createInterviewEvent_(startDate, company, role, row) {
  const cal = CalendarApp.getDefaultCalendar()
  const endDate = new Date(startDate.getTime() + 60 * 60 * 1000)

  const title = 'Interview: ' + role + ' at ' + company
  const description = [
    'Auto-created by Job CRM (sheet row ' + row + ').',
    '',
    'Open Job CRM sidebar before this event for prep mode: company research, likely questions, prep checklist.'
  ].join('\n')

  cal.createEvent(title, startDate, endDate, {
    description: description
  })

  logActivity_({
    action: 'calendar_created',
    rowRef: 'row:' + row,
    notes: company + ' / ' + role + ' / ' + startDate.toISOString()
  })
}
