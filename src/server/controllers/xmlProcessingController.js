const HttpError = require('http-errors')
const xmlProcessingService = require('../services/xmlProcessingService')

module.exports.processXmlMultipart = async function (request, _reply) {
  const parts = request.parts()
  let fileType = null
  let fileBuffer = null
  let originalFileName = null

  for await (const part of parts) {
    if (part.type === 'field' && (part.fieldname === 'fileType' || part.fieldname === 'type')) {
      fileType = String(part.value || '').trim()
      continue
    }

    if (part.fieldname === 'file' && part.file) {
      originalFileName = part.filename || null
      const chunks = []
      for await (const chunk of part.file) {
        chunks.push(chunk)
      }
      fileBuffer = Buffer.concat(chunks)
    }
  }

  if (!fileType) {
    throw new HttpError[400]('Missing required field: fileType')
  }

  if (!fileBuffer || fileBuffer.length === 0) {
    throw new HttpError[400]('Missing uploaded XML file')
  }

  if (!['contract', 'appendix2'].includes(fileType)) {
    throw new HttpError[400]('Unsupported fileType. Allowed values: contract, appendix2')
  }

  const processed = xmlProcessingService.processXmlBuffer(fileBuffer, fileType, {
    sourceFileName: originalFileName
  })

  return {
    data: processed
  }
}
