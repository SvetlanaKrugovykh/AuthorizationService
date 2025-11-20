
const linkController = require('../controllers/linkController')
const isAuthorizedGuard = require('../guards/is-authorized.guard')
const linkSchema = require('../schemas/link.schema')
const linkSchemaMf = require('../schemas/link.schemaMf')
const linkSchemaXlsx = require('../schemas/link.schemaXlsx')

module.exports = (fastify, _opts, done) => {
  fastify.route({
    method: 'POST',
    url: '/through/',
    handler: linkController.linkThroughJson,
    preHandler: [isAuthorizedGuard],
    schema: linkSchema
  })

  fastify.route({
    method: 'POST',
    url: '/through/mf',
    handler: linkController.linkThroughMultipart,
    preHandler: [isAuthorizedGuard],
    schema: linkSchemaMf
  })

  fastify.route({
    method: 'POST',
    url: '/through/xlsx',
    handler: linkController.linkThroughXlsx,
    preHandler: [isAuthorizedGuard],
    schema: linkSchemaXlsx
  })

  done()
}