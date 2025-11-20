module.exports = {
  description: 'Upload and process XLSX file',
  tags: ['xlsx'],
  summary: 'Process XLSX file and return result',
  consumes: ['multipart/form-data'],
  response: {
    200: {
      description: 'Processed XLSX file or Google Drive link',
      type: 'object',
      properties: {
        url: { type: 'string' }
      }
    },
    400: {
      description: 'Bad request',
      type: 'object',
      properties: {
        error: { type: 'string' }
      }
    }
  }
}
