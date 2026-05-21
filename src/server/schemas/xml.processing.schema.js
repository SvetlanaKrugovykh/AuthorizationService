module.exports = {
  description: 'Convert XML contract/appendix to structured JSON',
  tags: ['xml'],
  summary: 'Convert XML file into normalized structured data',
  headers: {
    type: 'object',
    properties: {
      Authorization: { type: 'string' }
    },
    required: ['Authorization']
  },
  consumes: ['multipart/form-data'],
  response: {
    200: {
      description: 'Successful response',
      type: 'object',
      properties: {
        data: { type: 'object' }
      }
    },
    400: {
      description: 'Bad request',
      type: 'object',
      properties: {
        statusCode: { type: 'integer' },
        error: { type: 'string' },
        message: { type: 'string' }
      }
    },
    500: {
      description: 'Internal server error',
      type: 'object',
      properties: {
        statusCode: { type: 'integer' },
        error: { type: 'string' },
        message: { type: 'string' }
      }
    }
  }
}
