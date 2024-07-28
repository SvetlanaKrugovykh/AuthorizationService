const linkController = require('../controllers/linkController')
const isAuthorizedGuard = require('../guards/is-authorized.guard')
const linkSchema = require('../schemas/link.schema')

module.exports = (fastify, _opts, done) => {
  fastify.route({
    method: 'POST',
    url: '/link/through/',
    handler: linkController.linkThrough,
    preHandler: [
      isAuthorizedGuard
    ],
    schema: linkSchema
  })

  done()
}