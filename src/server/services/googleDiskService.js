
const { google } = require('googleapis')
const fs = require('fs')
const path = require('path')
require('dotenv').config()
const { getSecretKey } = require('../guards/getCredentials')

const CREDENTIALS_PATH = process.env.GOOGLE_APPLICATION_CREDENTIALS
const FOLDER_ID = process.env.GOOGLE_DRIVE_FOLDER_ID

function getDriveService() {
  // Use getSecretKey to ensure credentials are read and cached consistently
  const serviceAccount = JSON.parse(fs.readFileSync(CREDENTIALS_PATH, 'utf8'))
  const auth = new google.auth.JWT({
    email: serviceAccount.client_email,
    key: serviceAccount.private_key,
    scopes: ['https://www.googleapis.com/auth/drive.file'],
  })
  return google.drive({ version: 'v3', auth })
}

/**
 * Uploads a file to Google Drive and returns the public link
 * @param {string} filePath - Absolute path to the file
 * @param {string} fileName - Name for the file on Drive
 * @returns {Promise<string>} - Public URL to the file
 */
async function uploadFileToDrive(filePath, fileName) {
  const drive = getDriveService()
  const fileMetadata = {
    name: fileName,
    parents: FOLDER_ID ? [FOLDER_ID] : [],
  }
  const media = {
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    body: fs.createReadStream(filePath),
  }
  const res = await drive.files.create({
    resource: fileMetadata,
    media,
    fields: 'id',
  })
  const fileId = res.data.id
  // Make file public
  await drive.permissions.create({
    fileId,
    requestBody: {
      role: 'reader',
      type: 'anyone',
    },
  })
  // Return public link
  return `https://drive.google.com/file/d/${fileId}/view?usp=sharing`
}

module.exports = {
  uploadFileToDrive,
}
