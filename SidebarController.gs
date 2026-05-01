/**
 * SidebarController.gs
 *
 * Backend handlers exposed to the sidebar via google.script.run. Keep these
 * focused: each one does one thing, returns plain data, throws on error.
 *
 * The sidebar calls getActiveRowData() (in Code.gs) to read state, and
 * delegates mutations through the handlers below.
 */

function runQuickAction(row, action) {
  if (!row || row < 2) throw new Error('Invalid row')
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_APPLICATIONS)
  const cols = getColumnMap_(sheet)

  switch (action) {
    case 'open-folder':
      return openDriveFolderForRow_(row, sheet, cols)
    case 'create-folder':
      return createDriveFolderForRow_(row, sheet, cols)
    case 'mark-applied':
      return setStatusAndTouch_(row, sheet, cols, 'Applied')
    case 'mark-interview':
      return setStatusAndTouch_(row, sheet, cols, 'Interview Scheduled')
    case 'mark-rejected':
      return setStatusAndTouch_(row, sheet, cols, 'Rejected')
    case 'mark-interview-done':
      return setStatusAndTouch_(row, sheet, cols, 'Interview Done')
    default:
      throw new Error('Unknown action: ' + action)
  }
}

function openDriveFolderForRow_(row, sheet, cols) {
  const url = sheet.getRange(row, cols['Drive Folder']).getValue()
  if (!url) throw new Error('No Drive folder set for this row.')
  return { ok: true, url: url }
}

function createDriveFolderForRow_(row, sheet, cols) {
  const company = sheet.getRange(row, cols['Company']).getValue()
  if (!company) throw new Error('Company is empty. Fill in Company before creating a folder.')

  const folderUrl = ensureCompanyDriveFolder_(company)

  const lock = LockService.getDocumentLock()
  try {
    lock.waitLock(5000)
    sheet.getRange(row, cols['Drive Folder']).setValue(folderUrl)
    sheet.getRange(row, cols['Last Touch']).setValue(new Date())
  } finally {
    try { lock.releaseLock() } catch (_) {}
  }

  logActivity_({
    action: 'folder_created',
    rowRef: 'row:' + row,
    notes: company
  })

  return { ok: true, url: folderUrl }
}

function setStatusAndTouch_(row, sheet, cols, status) {
  const lock = LockService.getDocumentLock()
  try {
    lock.waitLock(5000)
    sheet.getRange(row, cols['Status']).setValue(status)
    sheet.getRange(row, cols['Last Touch']).setValue(new Date())
    if (status === 'Applied' && !sheet.getRange(row, cols['Date Applied']).getValue()) {
      sheet.getRange(row, cols['Date Applied']).setValue(new Date())
    }
  } finally {
    try { lock.releaseLock() } catch (_) {}
  }

  logActivity_({
    action: 'status_change',
    rowRef: 'row:' + row,
    notes: status
  })

  return { ok: true, status: status }
}

/**
 * Sidebar lifecycle hook. Currently a no-op; reserved for future use
 * (e.g. registering the active sidebar instance for live updates).
 */
function registerSidebarReady() {
  return { ok: true }
}

/**
 * Selection-change trigger. Pushes a refresh signal that the sidebar's
 * window focus listener will pick up. Apps Script can't push to an open
 * sidebar directly, so we rely on the focus event to re-pull state.
 */
function onSelectionChange_(e) {
  // No-op for now. The sidebar's window.focus listener handles refresh.
  // Reserved hook for future server-side state syncing.
}
