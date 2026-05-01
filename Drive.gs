/**
 * Drive.gs
 *
 * Per-company folders live under DRIVE_PARENT_FOLDER_ID (set during Setup).
 * One folder per company, naming convention "[Company] - Job Hunt 2026".
 * Resume variants, cover letters, screenshots, and any other artifacts
 * land in there.
 */

function ensureCompanyDriveFolder_(companyName) {
  const parentId = PropertiesService.getScriptProperties().getProperty('DRIVE_PARENT_FOLDER_ID')
  if (!parentId) {
    throw new Error('DRIVE_PARENT_FOLDER_ID not set. Run Setup from the Job CRM menu.')
  }

  const parent = DriveApp.getFolderById(parentId)
  const folderName = formatCompanyFolderName_(companyName)

  const existing = parent.getFoldersByName(folderName)
  if (existing.hasNext()) {
    return existing.next().getUrl()
  }

  const folder = parent.createFolder(folderName)
  return folder.getUrl()
}

function formatCompanyFolderName_(company) {
  const year = new Date().getFullYear()
  return company.toString().trim() + ' - Job Hunt ' + year
}

function getOrCreateFolderById_(parentFolder, name) {
  const existing = parentFolder.getFoldersByName(name)
  if (existing.hasNext()) return existing.next()
  return parentFolder.createFolder(name)
}
