const xmlProcessingController = require('../controllers/xmlProcessingController')
const isAuthorizedGuard = require('../guards/is-authorized.guard')
const xmlProcessingSchema = require('../schemas/xml.processing.schema')

module.exports = (fastify, _opts, done) => {
  fastify.route({
    method: 'POST',
    url: '/convert',
    handler: xmlProcessingController.processXmlMultipart,
    preHandler: [isAuthorizedGuard],
    schema: xmlProcessingSchema
  })

  done()
}
