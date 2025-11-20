module.exports = {
  description: 'Upload and process XLSX file',
  tags: ['xlsx'],
  summary: 'Process XLSX file and return result',
  consumes: ['multipart/form-data'],
  body: {
    type: 'object',
    properties: {
      file: { type: 'string', format: 'binary' },
      variant: { type: 'string', enum: ['1', '2'], default: '1' }
    },
    required: ['file']
  },
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
